// api/cloud_sync.js
// Vercel Serverless Function: アンケート解析システム クラウド自動記憶・複数PC間リアルタイム同期API
// 保存先: 非公開(Private)リポジトリ内の data/yachiyo_survey_{key}.json
// ※ 認証トークンは Vercel の暗号化環境変数から取得するためソースコードには一切含まれません。

module.exports = async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
  res.setHeader('Pragma', 'no-cache');
  res.setHeader('Expires', '0');
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  const token = process.env.GITHUB_STORAGE_TOKEN;
  const owner = process.env.GITHUB_STORAGE_OWNER || 'noriyukihoshina';
  const repo = process.env.GITHUB_STORAGE_REPO || 'shift-schedule-cloud-data';

  if (!token) {
    return res.status(503).json({
      ok: false,
      error: 'CLOUD_NOT_CONFIGURED'
    });
  }

  const rawKey = (req.method === 'GET' ? req.query?.key : req.body?.key) || 'main';
  const safeKey = String(rawKey).replace(/[^a-zA-Z0-9_-]/g, '') || 'main';
  const filePath = `data/yachiyo_survey_${safeKey}.json`;
  const ghApiUrl = `https://api.github.com/repos/${owner}/${repo}/contents/${filePath}`;

  const ghHeaders = {
    'Authorization': `token ${token}`,
    'Accept': 'application/vnd.github.v3+json',
    'User-Agent': 'YachiyoSurveyApp-CloudSync',
    'Content-Type': 'application/json'
  };

  async function fetchCurrentCloudFile() {
    const resp = await fetch(`${ghApiUrl}?t=${Date.now()}`, {
      method: 'GET',
      headers: ghHeaders,
      cache: 'no-store'
    });
    if (resp.status === 404) {
      return { exists: false, sha: null, data: null };
    }
    if (!resp.ok) {
      const errTxt = await resp.text();
      throw new Error(`GitHub GET error (${resp.status}): ${errTxt}`);
    }
    const json = await resp.json();
    // Large file support via git blob API if content is empty
    let contentStr = '';
    if (json.content) {
      contentStr = Buffer.from(json.content, 'base64').toString('utf8');
    } else if (json.sha) {
      const blobUrl = `https://api.github.com/repos/${owner}/${repo}/git/blobs/${json.sha}`;
      const blobResp = await fetch(blobUrl, { method: 'GET', headers: ghHeaders, cache: 'no-store' });
      if (blobResp.ok) {
        const blobJson = await blobResp.json();
        contentStr = Buffer.from(blobJson.content || '', 'base64').toString('utf8');
      }
    }
    let parsed = null;
    try {
      parsed = JSON.parse(contentStr);
    } catch (e) {
      parsed = null;
    }
    return { exists: true, sha: json.sha, data: parsed };
  }

  try {
    if (req.method === 'GET') {
      const current = await fetchCurrentCloudFile();
      return res.status(200).json({
        ok: true,
        key: safeKey,
        exists: current.exists,
        cloudData: current.data
      });
    }

    if (req.method === 'POST') {
      const body = typeof req.body === 'string' ? JSON.parse(req.body) : (req.body || {});
      const incomingData = body.data || {};

      for (let attempt = 0; attempt < 3; attempt++) {
        const current = await fetchCurrentCloudFile();
        let toSave = incomingData;

        // 画像チャンクの場合は既存のページ画像マップとマージして消えないように保護
        if (safeKey.startsWith('img_chunk_') && current.data && typeof current.data === 'object' && typeof incomingData === 'object') {
          toSave = {
            ...current.data,
            ...incomingData,
            updatedAt: new Date().toISOString()
          };
        } else {
          toSave = {
            ...incomingData,
            updatedAt: new Date().toISOString()
          };
        }

        const payloadStr = JSON.stringify(toSave);
        const base64Content = Buffer.from(payloadStr, 'utf8').toString('base64');

        const putBody = {
          message: `Sync survey ${safeKey} at ${toSave.updatedAt}`,
          content: base64Content
        };
        if (current.sha) {
          putBody.sha = current.sha;
        }

        const putResp = await fetch(ghApiUrl, {
          method: 'PUT',
          headers: ghHeaders,
          body: JSON.stringify(putBody)
        });

        if (putResp.ok) {
          return res.status(200).json({
            ok: true,
            key: safeKey,
            updatedAt: toSave.updatedAt
          });
        }

        if (putResp.status === 409 && attempt < 2) {
          await new Promise(r => setTimeout(r, 250 * (attempt + 1)));
          continue;
        }

        const errText = await putResp.text();
        throw new Error(`GitHub PUT error (${putResp.status}): ${errText}`);
      }
    }

    return res.status(405).json({ ok: false, error: 'Method Not Allowed' });
  } catch (err) {
    console.error('Survey Cloud Sync API Error:', err);
    return res.status(500).json({
      ok: false,
      error: err.message || 'Internal Server Error'
    });
  }
};
