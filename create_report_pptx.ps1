$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$outPath = "c:\Users\coffe\OneDrive\デスクトップ\Antigravity\アンケート解析\八千代リハビリテーション病院_どーんと祭出展結果報告_四役プレゼン.pptx"
$webPath = "c:\Users\coffe\OneDrive\デスクトップ\Antigravity\アンケート解析\report_presentation.pptx"
$artifactDir = "C:\Users\coffe\.gemini\antigravity\brain\57f12e92-5bc8-492c-8c7d-8ab9b420cf7a"

$pp = New-Object -ComObject PowerPoint.Application
foreach ($openP in @($pp.Presentations)) {
    if ($openP.Path -eq "" -or $openP.Name -like "*四役プレゼン*" -or $openP.Name -like "*report_presentation*") {
        $openP.Saved = -1
        $openP.Close()
    }
}
$pres = $pp.Presentations.Add($false)
$pres.PageSetup.SlideWidth = 960
$pres.PageSetup.SlideHeight = 540

function RGB([int]$r, [int]$g, [int]$b) {
    return [int]($r + ($g * 256) + ($b * 65536))
}

# ==============================================================================
# カラーパレット定義（メインカラーをブルーに完全統一＋柔らかいサブカラー＆意味のあるグラフ配色）
# ==============================================================================
$colBg        = RGB 244 247 252   # 優しいペールブルーグレー背景
$colWhite     = RGB 255 255 255
$colText      = RGB 30 41 59      # 柔らかいチャコールネイビー
$colSubText   = RGB 100 116 139   # ミディアムスレート
$colBorder    = RGB 216 226 238   # 柔らかいカード枠線
$colBarBg     = RGB 233 238 245   # 棒グラフ背景トラック

# メインカラー（全スライド共通で統一感を出す）
$colMain      = RGB 37 99 235     # メイン：ロイヤルブルー
$colMainSoft  = RGB 239 246 255   # メインパステル：ソフトアイスブルー

# 意味を持たせるグラフ＆サブアクセントカラー（柔らかみのあるトーン）
$colBlue      = RGB 37 99 235     # 多い所・肯定・主力（ブルー）
$colCoral     = RGB 225 75 75     # 少ない所・弱点・問題あり（ソフトコーラルレッド）
$colSlate     = RGB 148 163 184   # デフォルト・中間・標準（ウォームグレー）
$colLightGray = RGB 203 213 225   # 未記入・その他（ライトグレー）
$colTeal      = RGB 16 163 127    # 補助アクセント：ソフトティール（好評価・達成）
$colAmber     = RGB 234 138 25    # 補助アクセント：ソフトアンバー（注目ポイント）

# 柔らかさを演出するパステル背景色
$colPastelB   = RGB 239 246 255   # パステルブルー
$colPastelT   = RGB 236 253 245   # パステルティール
$colPastelA   = RGB 254 249 235   # パステルアンバー
$colPastelC   = RGB 254 242 242   # パステルコーラル
$colPastelG   = RGB 248 250 252   # パステルグレー

$totalSlides = 22

function Add-Rect($slide, [single]$left, [single]$top, [single]$width, [single]$height, [int]$fillRGB, [int]$lineRGB = -1, [bool]$round = $true, [single]$roundness = 0.08) {
    $type = if ($round) { 5 } else { 1 }
    $sh = $slide.Shapes.AddShape([int]$type, [single]$left, [single]$top, [single]$width, [single]$height)
    $sh.Fill.Visible = -1
    $sh.Fill.ForeColor.RGB = [int]$fillRGB
    if ($lineRGB -ge 0) {
        $sh.Line.Visible = -1
        $sh.Line.ForeColor.RGB = [int]$lineRGB
        $sh.Line.Weight = [single]1.5
    } else {
        $sh.Line.Visible = 0
    }
    if ($round -and $sh.Adjustments.Count -ge 1) {
        $sh.Adjustments.Item(1) = [single]$roundness
    }
    return $sh
}

function Add-Oval($slide, [single]$left, [single]$top, [single]$width, [single]$height, [int]$fillRGB, [int]$lineRGB = -1) {
    $sh = $slide.Shapes.AddShape(9, [single]$left, [single]$top, [single]$width, [single]$height)
    $sh.Fill.Visible = -1
    $sh.Fill.ForeColor.RGB = [int]$fillRGB
    if ($lineRGB -ge 0) {
        $sh.Line.Visible = -1
        $sh.Line.ForeColor.RGB = [int]$lineRGB
        $sh.Line.Weight = [single]1.5
    } else {
        $sh.Line.Visible = 0
    }
    return $sh
}

function Add-Text($slide, [single]$left, [single]$top, [single]$width, [single]$height, [string]$text, [single]$fontSize = 16, [int]$colorRGB = $colText, [bool]$bold = $false, [int]$align = 1) {
    $tb = $slide.Shapes.AddTextbox(1, [single]$left, [single]$top, [single]$width, [single]$height)
    $tf = $tb.TextFrame
    $tf.MarginLeft = [single]6
    $tf.MarginRight = [single]6
    $tf.MarginTop = [single]4
    $tf.MarginBottom = [single]4
    $tf.WordWrap = -1
    $tr = $tf.TextRange
    $tr.Text = $text
    $tr.Font.Name = "Meiryo UI"
    $tr.Font.NameFarEast = "Meiryo UI"
    $tr.Font.Size = [single]$fontSize
    $tr.Font.Color.RGB = [int]$colorRGB
    $tr.Font.Bold = if ($bold) { -1 } else { 0 }
    $tr.ParagraphFormat.Alignment = [int]$align
    return $tb
}

# 統一メインカラー（ロイヤルブルー）によるスッキリ1段のソフトヘッダー（サブリード重複なし）
function Add-SoftHeader($slide, [string]$badgeText, [string]$titleText, [int]$pageNo) {
    Add-Rect $slide 0 0 960 540 $colBg -1 $false | Out-Null
    # 上部ソフトヘッダーカード（1段でスッキリ・丸角）
    Add-Rect $slide 20 14 920 62 $colWhite $colBorder $true 0.18 | Out-Null
    # ピル型バッジ（全スライド統一のメインブルー）
    Add-Rect $slide 34 27 186 36 $colMain -1 $true 0.5 | Out-Null
    Add-Text $slide 34 30 186 30 $badgeText 13 $colWhite $true 2 | Out-Null
    # タイトル（大きく見やすく）
    Add-Text $slide 232 24 615 42 $titleText 21 $colText $true 1 | Out-Null
    # ページ番号ピル（メインパステル＋メインブルー文字で統一）
    Add-Rect $slide 856 29 68 32 $colMainSoft -1 $true 0.5 | Out-Null
    Add-Text $slide 856 31 68 28 ("$pageNo / $totalSlides") 13 $colMain $true 2 | Out-Null
}

# 丸みのある柔らかい白カード（見出しバーもメインカラーで統一、背景色で柔らかさを演出可能）
function Add-SoftCard($slide, [single]$left, [single]$top, [single]$width, [single]$height, [string]$title = "", [int]$accentColor = $colMain, [int]$bgRGB = $colWhite) {
    Add-Rect $slide $left $top $width $height $bgRGB $colBorder $true 0.06 | Out-Null
    if ($title -ne "") {
        Add-Rect $slide ($left + 18) ($top + 18) 8 28 $accentColor -1 $true 0.5 | Out-Null
        Add-Text $slide ($left + 34) ($top + 15) ($width - 50) 34 $title 18 $colText $true 1 | Out-Null
    }
}

# ドーナツ円グラフ描画関数
function Add-DonutChart($slide, [single]$cx, [single]$cy, [single]$diameter, [array]$slices, [string]$centerTopText, [string]$centerSubText) {
    $left = $cx - ($diameter / 2.0)
    $top  = $cy - ($diameter / 2.0)
    $startAngle = -90.0
    foreach ($s in $slices) {
        $pct = [double]$s.pct
        if ($pct -le 0.1) { continue }
        $sweep = ($pct / 100.0) * 360.0
        if ($pct -ge 99.9) {
            Add-Oval $slide $left $top $diameter $diameter $s.color $colWhite | Out-Null
        } else {
            $endAngle = $startAngle + $sweep
            $pie = $slide.Shapes.AddShape(142, [single]$left, [single]$top, [single]$diameter, [single]$diameter)
            $pie.Fill.Visible = -1
            $pie.Fill.ForeColor.RGB = [int]$s.color
            $pie.Line.Visible = -1
            $pie.Line.ForeColor.RGB = [int]$colWhite
            $pie.Line.Weight = [single]2.5
            $pie.Adjustments.Item(1) = [single]$startAngle
            $pie.Adjustments.Item(2) = [single]$endAngle
            $startAngle = $endAngle
        }
    }
    $holeD = $diameter * 0.58
    $holeL = $cx - ($holeD / 2.0)
    $holeT = $cy - ($holeD / 2.0)
    Add-Oval $slide $holeL $holeT $holeD $holeD $colWhite -1 | Out-Null
    if ($centerTopText -ne "") {
        Add-Text $slide ($cx - 75) ($cy - 22) 150 26 $centerTopText 14 $colText $true 2 | Out-Null
    }
    if ($centerSubText -ne "") {
        Add-Text $slide ($cx - 75) ($cy + 2) 150 24 $centerSubText 13 $colSubText $true 2 | Out-Null
    }
}

# 円グラフ用の凡例アイテム（柔らかいパステルハイライト対応・文字重なり完全防止）
function Add-LegendItem($slide, [single]$left, [single]$top, [single]$width, [int]$color, [string]$label, [string]$valText, [bool]$highlight = $false, [int]$hlBg = $colPastelB) {
    if ($highlight) {
        Add-Rect $slide ($left - 8) ($top - 4) ($width + 16) 38 $hlBg -1 $true 0.3 | Out-Null
    }
    $fSize = if ($width -le 230) { 14.0 } else { 15.0 }
    Add-Oval $slide $left ($top + 6) 15 15 $color -1 | Out-Null
    Add-Text $slide ($left + 20) $top ($width - 110) 28 $label $fSize $colText $true 1 | Out-Null
    $valColor = if ($color -eq $colLightGray) { $colSubText } else { $color }
    Add-Text $slide ($left + $width - 135) $top 135 28 $valText $fSize $valColor $true 3 | Out-Null
}

# 意味のある色付けを行う丸角横棒グラフ（デフォルト＝グレー、多い所＝青/赤、少ない所＝逆の色）
function Add-SoftBar($slide, [single]$left, [single]$top, [single]$labelW, [single]$maxBarW, [string]$label, [string]$valText, [single]$pctVal, [int]$barColor, [single]$fontSize = 15.5) {
    $isEmph = ($barColor -eq $colBlue -or $barColor -eq $colCoral)
    $lblColor = if ($isEmph) { $colText } else { $colSubText }
    $valColor = if ($isEmph) { $barColor } else { $colText }
    Add-Text $slide $left $top $labelW 30 $label $fontSize $lblColor $isEmph 1 | Out-Null
    $barX = $left + $labelW + 8
    Add-Rect $slide $barX ($top + 5) $maxBarW 22 $colBarBg -1 $true 0.5 | Out-Null
    $fillW = [math]::Max(10, [math]::Round($maxBarW * ($pctVal / 100.0)))
    Add-Rect $slide $barX ($top + 5) $fillW 22 $barColor -1 $true 0.5 | Out-Null
    Add-Text $slide ($barX + $maxBarW + 10) $top 180 30 $valText $fontSize $valColor $true 1 | Out-Null
}

# 棒グラフの色分けルール凡例ピル
function Add-BarColorRule($slide, [single]$left, [single]$top, [string]$highLabel, [int]$highCol, [string]$lowLabel, [int]$lowCol) {
    Add-Rect $slide $left $top 530 34 $colPastelG $colBorder $true 0.4 | Out-Null
    Add-Oval $slide ($left + 12) ($top + 10) 13 13 $highCol -1 | Out-Null
    Add-Text $slide ($left + 28) ($top + 4) 182 26 $highLabel 12.5 $highCol $true 1 | Out-Null
    Add-Oval $slide ($left + 212) ($top + 10) 13 13 $lowCol -1 | Out-Null
    Add-Text $slide ($left + 228) ($top + 4) 182 26 $lowLabel 12.5 $lowCol $true 1 | Out-Null
    Add-Oval $slide ($left + 412) ($top + 10) 13 13 $colSlate -1 | Out-Null
    Add-Text $slide ($left + 428) ($top + 4) 100 26 "灰：中間・標準" 12.5 $colSubText $true 1 | Out-Null
}

function Style-SoftTable($tbl, [int]$rows, [int]$cols, [single]$fontSize = 15) {
    for ($r = 1; $r -le $rows; $r++) {
        for ($c = 1; $c -le $cols; $c++) {
            $cell = $tbl.Cell($r, $c)
            $tr = $cell.Shape.TextFrame.TextRange
            $tr.Font.Name = "Meiryo UI"
            $tr.Font.NameFarEast = "Meiryo UI"
            $tr.Font.Size = [single]$fontSize
            if ($r -eq 1) {
                $cell.Shape.Fill.ForeColor.RGB = [int]$colMain
                $tr.Font.Color.RGB = [int]$colWhite
                $tr.Font.Bold = -1
            } else {
                $cell.Shape.Fill.ForeColor.RGB = if ($r % 2 -eq 0) { [int]$colWhite } else { [int]$colMainSoft }
                $tr.Font.Color.RGB = [int]$colText
            }
        }
    }
}

# ==============================================================================
# SLIDE 1: 表紙 ＆ 起案の3つの目的（メインカラー統一＋柔らかいパステルアクセント）
# ==============================================================================
$s1 = $pres.Slides.Add(1, 12)
Add-Rect $s1 0 0 960 540 $colBg -1 $false | Out-Null

# ヒーローバナー（丸角・メインブルー）
Add-Rect $s1 24 20 912 168 $colMain -1 $true 0.08 | Out-Null
# 表紙の柔らかいアイコン（丸＋医療クロス）
Add-Oval $s1 52 48 68 68 $colWhite -1 | Out-Null
Add-Rect $s1 78 62 16 40 $colTeal -1 $true 0.2 | Out-Null
Add-Rect $s1 66 74 40 16 $colTeal -1 $true 0.2 | Out-Null

Add-Text $s1 140 38 760 28 "八千代リハビリテーション病院 ｜ 四役報告資料（2026年10月17日・18日 開催）" 15 (RGB 219 234 254) $true 1 | Out-Null
Add-Text $s1 140 68 760 52 "八千代市産業祭「どーんと祭」出展結果報告書" 29 $colWhite $true 1 | Out-Null
Add-Text $s1 140 128 760 44 "ブース来場者：268名（1日目120名・2日目148名） ／ アンケート有効回答：236名`n出展規模：3区画・6部署協働（参加スタッフ14名） ／ 報告者：リハビリテーション部 保科憲幸" 14 (RGB 239 246 255) $false 1 | Out-Null

Add-Text $s1 32 204 896 30 "■ 起案書における3つの出展目的" 18 $colText $true 1 | Out-Null

# 3つの目的カード（パステル背景で柔らかさを演出）
Add-Rect $s1 24 242 292 272 $colWhite $colBorder $true 0.08 | Out-Null
Add-Rect $s1 40 258 260 40 $colPastelB -1 $true 0.4 | Out-Null
Add-Text $s1 40 264 260 30 "目的① 地域貢献" 17 $colMain $true 2 | Out-Null
Add-Text $s1 40 314 260 180 "多職種による健康測定・相談の機会を提供し、地域医療への貢献と住民の健康寿命延伸に寄与する。" 16 $colText $false 1 | Out-Null

Add-Rect $s1 334 242 292 272 $colWhite $colBorder $true 0.08 | Out-Null
Add-Rect $s1 350 258 260 40 $colPastelA -1 $true 0.4 | Out-Null
Add-Text $s1 350 264 260 30 "目的② 認知・集患" 17 $colAmber $true 2 | Out-Null
Add-Text $s1 350 314 260 180 "「顔が見える医療機関」として認知を確立し、将来入院を要する際に第一選択として思い浮かぶ下地を作る。" 16 $colText $false 1 | Out-Null

Add-Rect $s1 644 242 292 272 $colWhite $colBorder $true 0.08 | Out-Null
Add-Rect $s1 660 258 260 40 $colPastelT -1 $true 0.4 | Out-Null
Add-Text $s1 660 264 260 30 "目的③ 多職種連携の強化" 17 $colTeal $true 2 | Out-Null
Add-Text $s1 660 314 260 180 "6部署が協働してイベントを企画・運営することで院内の多職種連携を強化し、日常業務の連携基盤を強固にする。" 16 $colText $false 1 | Out-Null


# ==============================================================================
# SLIDE 2: 【第1部】来場者アンケートの質問事項と見えること
# ==============================================================================
$s2 = $pres.Slides.Add(2, 12)
Add-SoftHeader $s2 "第1部：来場者調査" "来場者アンケートの質問事項とそこから見えること" 2

$t2 = $s2.Shapes.AddTable(7, 3, 24, 92, 912, 424).Table
$t2.Columns.Item(1).Width = 240
$t2.Columns.Item(2).Width = 352
$t2.Columns.Item(3).Width = 320
$t2Data = @(
    @("質問項目", "選択肢", "この質問から何が見えるか"),
    @("基本情報", "年齢（10〜90代）／ 性別 ／ お住まい（市内・市外）", "来場者の属性と1日目・2日目の違い"),
    @("Q1. 参考になった内容", "筋肉量 ／ 血管年齢 ／ お薬 ／ 栄養 ／ 人生会議", "今回の出展内容に対する評価"),
    @("Q2. 当院をご存知でしたか", "よく知っている ／ 名前・場所は知る ／ 今日初めて知った", "地域における当院の現在の認知度"),
    @("Q3. 当院を利用したいですか", "ぜひ利用したい ／ 候補として検討 ／ 他院・不明", "アンケート回答時の当院の評価"),
    @("Q4. 健康で気になること", "転倒 ／ 認知機能 ／ 栄養 ／ 薬 ／ 介護終活 ／ その他", "地域住民の抱える課題や関心のチェック"),
    @("Q5. 来年に期待する内容", "自由記述（39件）", "来年度の出展内容のヒント")
)
for ($r = 0; $r -lt 7; $r++) {
    for ($c = 0; $c -lt 3; $c++) {
        $t2.Cell($r+1, $c+1).Shape.TextFrame.TextRange.Text = $t2Data[$r][$c]
    }
}
Style-SoftTable $t2 7 3 15.5
for ($r = 2; $r -le 7; $r++) {
    $t2.Cell($r, 1).Shape.TextFrame.TextRange.Font.Bold = -1
    $t2.Cell($r, 3).Shape.TextFrame.TextRange.Font.Bold = -1
    $t2.Cell($r, 3).Shape.TextFrame.TextRange.Font.Color.RGB = $colMain
}


# ==============================================================================
# SLIDE 3: 【第1部】基本情報①：お住まい ＆ 性別（2つの大きな円グラフ）
# ==============================================================================
$s3 = $pres.Slides.Add(3, 12)
Add-SoftHeader $s3 "第1部：来場者結果①" "基本情報①：お住まい と 性別（N=236名）" 3

# 左カード：お住まい円グラフ
Add-SoftCard $s3 20 90 450 316 "お住まいの内訳" $colMain
Add-DonutChart $s3 130 262 184 @(
    @{ pct = 72.5; color = $colBlue },
    @{ pct = 26.3; color = $colTeal },
    @{ pct = 1.3;  color = $colLightGray }
) "市内 72.5%" "171名 / 236名"

Add-LegendItem $s3 236 190 218 $colBlue      "八千代市内" "171名 (72.5%)" $true $colPastelB
Add-LegendItem $s3 236 244 218 $colTeal      "八千代市外" "62名 (26.3%)" $false
Add-LegendItem $s3 236 296 218 $colLightGray "未記入" "3名 (1.3%)" $false

# 右カード：性別円グラフ
Add-SoftCard $s3 490 90 450 316 "性別の内訳" $colMain
Add-DonutChart $s3 600 262 184 @(
    @{ pct = 64.8; color = $colAmber },
    @{ pct = 34.3; color = $colBlue },
    @{ pct = 0.8;  color = $colLightGray }
) "女性 64.8%" "153名 / 236名"

Add-LegendItem $s3 706 190 218 $colAmber     "女性" "153名 (64.8%)" $true $colPastelA
Add-LegendItem $s3 706 244 218 $colBlue      "男性" "81名 (34.3%)" $false
Add-LegendItem $s3 706 296 218 $colLightGray "未記入" "2名 (0.8%)" $false

# 下部：端的な解説ボックス
Add-Rect $s3 20 418 920 100 $colPastelB $colBorder $true 0.12 | Out-Null
Add-Text $s3 40 432 880 76 "・お住まいは「八千代市内」が171名（72.5%）で7割強を占め、「八千代市外」が62名（26.3%）であった。`n・性別は「女性」が153名（64.8%）、「男性」が81名（34.3%）であった。" 16.5 $colText $true 1 | Out-Null


# ==============================================================================
# SLIDE 4: 【第1部】基本情報②：年齢別の内訳（意味のある棒グラフ配色：最多＝青、最少＝赤、標準＝灰）
# ==============================================================================
$s4 = $pres.Slides.Add(4, 12)
Add-SoftHeader $s4 "第1部：来場者結果②" "基本情報②：年齢別の内訳（N=236名）" 4

Add-SoftCard $s4 20 90 566 428 "年代別の回答数・割合" $colMain
$ages4 = @(
    @("10〜20歳代", "18名 (7.6%)", 7.6, $colSlate),
    @("30歳代", "12名 (5.1%)", 5.1, $colCoral),
    @("40歳代", "33名 (14.0%)", 14.0, $colSlate),
    @("50歳代", "62名 (26.3%)", 26.3, $colBlue),
    @("60歳代", "39名 (16.5%)", 16.5, $colSlate),
    @("70歳代", "42名 (17.8%)", 17.8, $colSlate),
    @("80〜90歳代", "27名 (11.4%)", 11.4, $colSlate),
    @("未記入", "3名 (1.3%)", 1.3, $colLightGray)
)
for ($i = 0; $i -lt $ages4.Count; $i++) {
    $r = $ages4[$i]
    Add-SoftBar $s4 40 (150 + $i * 39) 115 240 $r[0] $r[1] ($r[2] * 3.2) $r[3] 15
}
Add-BarColorRule $s4 38 470 "青：最多層 (50代)" $colBlue "赤：最少層 (30代)" $colCoral

Add-SoftCard $s4 602 90 338 428 "結果の説明" $colMain
$s4Text = @"
・50歳代が62名（26.3%）で最も多い（青色）。

・次いで70歳代が42名（17.8%）、60歳代が39名（16.5%）、40歳代が33名（14.0%）となった。

・未記入を除くと、30歳代が12名（5.1%）で最も少なかった（赤色）。
"@
Add-Text $s4 622 156 298 330 $s4Text 16.5 $colText $false 1 | Out-Null


# ==============================================================================
# SLIDE 5: 【第1部】Q1. 本日参考になった内容（多い所＝青、少ない所＝赤、デフォルト＝灰）
# ==============================================================================
$s5 = $pres.Slides.Add(5, 12)
Add-SoftHeader $s5 "第1部：来場者結果③" "Q1. 本日参考になった内容（今回の出展の評価）" 5

Add-SoftCard $s5 20 90 566 428 "Q1. 参考になった内容（複数回答・N=236）" $colMain
Add-SoftBar $s5 40 165 185 210 "血管年齢測定" "172名 (72.9%)" 72.9 $colBlue 15.5
Add-SoftBar $s5 40 225 185 210 "筋肉量測定(InBody)" "157名 (66.5%)" 66.5 $colBlue 15.5
Add-SoftBar $s5 40 285 185 210 "栄養レッスン" "57名 (24.2%)" 24.2 $colSlate 15.5
Add-SoftBar $s5 40 345 185 210 "人生会議(終活)の案内" "41名 (17.4%)" 17.4 $colSlate 15.5
Add-SoftBar $s5 40 405 185 210 "お薬相談" "32名 (13.6%)" 13.6 $colCoral 15.5
Add-BarColorRule $s5 38 470 "青：上位（6割超）" $colBlue "赤：最少項目" $colCoral

Add-SoftCard $s5 602 90 338 428 "結果の説明" $colMain
$s5Text = @"
・「血管年齢測定」が172名（72.9%）、「筋肉量測定」が157名（66.5%）と、上位2項目が約7割を占めた（青色）。

・中間は「栄養レッスン」57名（24.2%）、「人生会議の案内」41名（17.4%）。

・最も少なかったのは「お薬相談」32名（13.6%）であった（赤色）。
"@
Add-Text $s5 622 156 298 330 $s5Text 16.5 $colText $false 1 | Out-Null


# ==============================================================================
# SLIDE 6: 【第1部】Q2. これまで当院をご存知でしたか？（大きな円グラフ）
# ==============================================================================
$s6 = $pres.Slides.Add(6, 12)
Add-SoftHeader $s6 "第1部：来場者結果④" "Q2. これまで当院をご存知でしたか？（現在の認知度）" 6

Add-SoftCard $s6 20 90 920 428 "Q2. 当院の認知度（単一回答・N=236）" $colMain

Add-DonutChart $s6 210 315 270 @(
    @{ pct = 53.4; color = $colBlue },
    @{ pct = 28.4; color = $colCoral },
    @{ pct = 16.5; color = $colTeal },
    @{ pct = 1.7;  color = $colLightGray }
) "初認知 28.4%" "名前のみ 53.4%"

Add-LegendItem $s6 410 165 490 $colBlue      "名前や場所は知っている" "126名 (53.4%)" $false
Add-LegendItem $s6 410 220 490 $colCoral     "今日初めて知った" "67名 (28.4%)" $true $colPastelC
Add-LegendItem $s6 410 275 490 $colTeal      "よく知っている（利用経験有）" "39名 (16.5%)" $false
Add-LegendItem $s6 410 328 490 $colLightGray "未記入" "4名 (1.7%)" $false

Add-Rect $s6 404 382 506 114 $colPastelB $colBorder $true 0.12 | Out-Null
Add-Text $s6 422 394 472 94 "・半数以上の126名（53.4%）が「名前や場所は知っている」と回答した。`n・67名（28.4%）が「今日初めて知った」、39名（16.5%）が「よく知っている」と回答した。" 16 $colText $true 1 | Out-Null


# ==============================================================================
# SLIDE 7: 【第1部】Q3. 万が一の際、当院を利用したいですか？（大きな円グラフ）
# ==============================================================================
$s7 = $pres.Slides.Add(7, 12)
Add-SoftHeader $s7 "第1部：来場者結果⑤" "Q3. 万が一の際、当院を利用したいですか？（当院の評価）" 7

Add-SoftCard $s7 20 90 920 428 "Q3. 万が一の際の利用意向（単一回答・N=236）" $colMain

Add-DonutChart $s7 210 315 270 @(
    @{ pct = 25.4; color = $colTeal },
    @{ pct = 64.8; color = $colBlue },
    @{ pct = 7.2;  color = $colSlate },
    @{ pct = 2.5;  color = $colLightGray }
) "好意的評価" "90.3% (213名)"

Add-LegendItem $s7 410 165 490 $colTeal      "ぜひ利用したい" "60名 (25.4%)" $true $colPastelT
Add-LegendItem $s7 410 220 490 $colBlue      "候補の一つとして検討したい" "153名 (64.8%)" $true $colPastelB
Add-LegendItem $s7 410 275 490 $colSlate     "他の病院を選ぶ・わからない" "17名 (7.2%)" $false
Add-LegendItem $s7 410 328 490 $colLightGray "未記入" "6名 (2.5%)" $false

Add-Rect $s7 404 382 506 114 $colPastelT $colBorder $true 0.12 | Out-Null
Add-Text $s7 422 394 472 94 "・「候補として検討したい」153名（64.8%）と「ぜひ利用したい」60名（25.4%）を合わせ、計213名（90.3%）が好意的な利用意向を示した。" 16 $colText $true 1 | Out-Null


# ==============================================================================
# SLIDE 8: 【第1部】Q4. 現在健康で気になること（多い所＝青、少ない所＝赤、標準＝灰）
# ==============================================================================
$s8 = $pres.Slides.Add(8, 12)
Add-SoftHeader $s8 "第1部：来場者結果⑥" "Q4. 現在健康について気になること（課題・関心のチェック）" 8

Add-SoftCard $s8 20 90 566 428 "Q4. 健康について気になること（複数回答・N=236）" $colMain
Add-SoftBar $s8 40 155 175 215 "足腰の衰え・転倒" "93名 (39.4%)" 39.4 $colBlue 15
Add-SoftBar $s8 40 207 175 215 "食事・栄養バランス" "79名 (33.5%)" 33.5 $colBlue 15
Add-SoftBar $s8 40 259 175 215 "物忘れ・認知機能" "78名 (33.1%)" 33.1 $colBlue 15
Add-SoftBar $s8 40 311 175 215 "将来の介護や終活" "71名 (30.1%)" 30.1 $colBlue 15
Add-SoftBar $s8 40 363 175 215 "とくにない" "34名 (14.4%)" 14.4 $colSlate 15
Add-SoftBar $s8 40 415 175 215 "薬の飲み方 ／ その他" "各11名 (4.7%)" 4.7 $colCoral 15
Add-BarColorRule $s8 38 470 "青：上位課題(30%超)" $colBlue "赤：最少項目(4.7%)" $colCoral

Add-SoftCard $s8 602 90 338 428 "結果の説明" $colMain
$s8Text = @"
・上位4項目（転倒39.4%、栄養33.5%、認知機能33.1%、介護・終活30.1%）がいずれも30%を超えた（青色）。

・最も少なかったのは「薬の飲み方」および「その他」の各11名（4.7%）であった（赤色）。

・「その他」には高血圧、糖尿病、膝の痛み、五十肩などが記載された。
"@
Add-Text $s8 622 156 298 330 $s8Text 16 $colText $false 1 | Out-Null


# ==============================================================================
# SLIDE 9: 【第1部】Q5. 来年に期待する内容（自由記述39件の内訳）
# ==============================================================================
$s9 = $pres.Slides.Add(9, 12)
Add-SoftHeader $s9 "第1部：来場者結果⑦" "Q5. 来年に期待する内容（来年度出展内容のヒント：39件）" 9

Add-SoftCard $s9 20 90 450 204 "① 運動・転倒・骨密度のチェック（12件）" $colMain $colWhite
Add-Rect $s9 38 144 414 134 $colPastelB -1 $true 0.1 | Out-Null
Add-Text $s9 48 154 394 114 "・「椅子に座ってできるストレッチや運動を知りたい」`n・「骨粗しょう症・骨密度や転倒危険度のチェック」" 16 $colText $false 1 | Out-Null

Add-SoftCard $s9 490 90 450 204 "② 今年度と同様の測定継続・好評の声（11件）" $colMain $colWhite
Add-Rect $s9 508 144 414 134 $colPastelT -1 $true 0.1 | Out-Null
Add-Text $s9 518 154 394 114 "・「筋肉量や血管年齢など今年と同じ測定を続けてほしい」`n・「マンションの入院経験者5名からリハが良かったと聞いた」" 16 $colText $false 1 | Out-Null

Add-SoftCard $s9 20 310 296 208 "③ 認知機能・フレイル（6件）" $colMain $colWhite
Add-Rect $s9 36 364 264 138 $colPastelA -1 $true 0.1 | Out-Null
Add-Text $s9 46 374 244 118 "・「認知機能の衰えチェックや体力測定、フレイル予防を知りたい」" 15.5 $colText $false 1 | Out-Null

Add-SoftCard $s9 332 310 296 208 "④ 介護・薬・受診相談（4件）" $colMain $colWhite
Add-Rect $s9 348 364 264 138 $colPastelG -1 $true 0.1 | Out-Null
Add-Text $s9 358 374 244 118 "・「介護生活の相談、家族が入院した際の薬、症状別の受診科」" 15.5 $colText $false 1 | Out-Null

Add-SoftCard $s9 644 310 296 208 "⑤ 備品・案内順路の要望（3件）" $colMain $colWhite
Add-Rect $s9 660 364 264 138 $colPastelC -1 $true 0.1 | Out-Null
Add-Text $s9 670 374 244 118 "・「記入用バインダー、InBodyの靴入れ・靴ベラ、順路表示」" 15.5 $colText $false 1 | Out-Null


# ==============================================================================
# SLIDE 10: 【第2部】スタッフアンケートの質問事項と見えること
# ==============================================================================
$s10 = $pres.Slides.Add(10, 12)
Add-SoftHeader $s10 "第2部：スタッフ調査" "スタッフアンケートの質問事項とそこから見えること（対象14名・回答13名）" 10

$t10 = $s10.Shapes.AddTable(7, 3, 24, 92, 912, 424).Table
$t10.Columns.Item(1).Width = 290
$t10.Columns.Item(2).Width = 200
$t10.Columns.Item(3).Width = 422
$t10Data = @(
    @("質問項目", "回答形式", "この質問から何が見えるか"),
    @("① 参加して良かったか（＋②理由）", "はい／いいえ＋記述", "参加満足度と多職種連携・地域PRの実感"),
    @("③ 地域住民に貢献できたか（＋④理由）", "はい／いいえ＋記述", "自部署の役割による地域貢献の達成度"),
    @("⑤ 3区画の広さはどうだったか（＋⑥問題点）", "問題なし／あり＋記述", "3区画の広さの妥当性と現場での問題点"),
    @("⑦ 人員配置は適切だったか（＋⑧問題点）", "問題なし／あり＋記述", "参加人数・休憩確保・誘導体制の妥当性"),
    @("⑨ 来年も参加する必要はあるか（＋⑩理由）", "はい／いいえ＋記述", "来年度の継続参加に対する意向と理由"),
    @("⑪ 来年に向けて改善・強化したいこと", "自由記述", "来年度の運営・設備・広報の改善提案")
)
for ($r = 0; $r -lt 7; $r++) {
    for ($c = 0; $c -lt 3; $c++) {
        $t10.Cell($r+1, $c+1).Shape.TextFrame.TextRange.Text = $t10Data[$r][$c]
    }
}
Style-SoftTable $t10 7 3 15.5
for ($r = 2; $r -le 7; $r++) {
    $t10.Cell($r, 1).Shape.TextFrame.TextRange.Font.Bold = -1
    $t10.Cell($r, 3).Shape.TextFrame.TextRange.Font.Bold = -1
    $t10.Cell($r, 3).Shape.TextFrame.TextRange.Font.Color.RGB = $colMain
}


# ==============================================================================
# SLIDE 11: 【第2部】スタッフ結果①：①参加満足度 ＆ ③地域貢献（2つの大きな円グラフ）
# ==============================================================================
$s11 = $pres.Slides.Add(11, 12)
Add-SoftHeader $s11 "第2部：スタッフ結果①" "スタッフ集計①：参加満足度（①）と 地域貢献度（③）" 11

Add-SoftCard $s11 20 90 450 316 "① 参加して良かったか（N=13）" $colMain
Add-DonutChart $s11 130 262 184 @(
    @{ pct = 100.0; color = $colBlue }
) "はい 100%" "13名 / 13名"
Add-LegendItem $s11 236 210 218 $colBlue      "はい" "13名 (100%)" $true $colPastelB
Add-LegendItem $s11 236 266 218 $colLightGray "いいえ" "0名 (0.0%)" $false

Add-SoftCard $s11 490 90 450 316 "③ 地域住民に貢献できたか（N=13）" $colMain
Add-DonutChart $s11 600 262 184 @(
    @{ pct = 92.3; color = $colBlue },
    @{ pct = 7.7;  color = $colSlate }
) "はい 92.3%" "12名 / 13名"
Add-LegendItem $s11 706 210 218 $colBlue  "はい" "12名 (92.3%)" $true $colPastelB
Add-LegendItem $s11 706 266 218 $colSlate "いいえ" "1名 (7.7%)" $false

Add-Rect $s11 20 418 920 100 $colPastelB $colBorder $true 0.12 | Out-Null
Add-Text $s11 40 430 880 78 "・設問①は13名全員（100%）が「はい」と回答し、地域住民との直接交流や他部門との一致団結・結束力が理由に挙げられた。`n・設問③は12名（92.3%）が「はい」と回答し、測定結果に基づく運動・栄養・生活指導やお薬カレンダー配布が評価された。" 15.5 $colText $true 1 | Out-Null


# ==============================================================================
# SLIDE 12: 【第2部】スタッフ結果②：⑤3区画の広さ ＆ ⑦人員配置（2つの大きな円グラフ）
# ==============================================================================
$s12 = $pres.Slides.Add(12, 12)
Add-SoftHeader $s12 "第2部：スタッフ結果②" "スタッフ集計②：3区画の広さ（⑤）と 人員配置（⑦）" 12

Add-SoftCard $s12 20 90 450 316 "⑤ 3区画の広さは適切だったか" $colMain
Add-DonutChart $s12 130 262 184 @(
    @{ pct = 61.5; color = $colBlue },
    @{ pct = 38.5; color = $colCoral }
) "問題あり" "38.5% (5名)"
Add-LegendItem $s12 236 210 218 $colBlue  "問題なし" "8名 (61.5%)" $false
Add-LegendItem $s12 236 266 218 $colCoral "問題あり" "5名 (38.5%)" $true $colPastelC

Add-SoftCard $s12 490 90 450 316 "⑦ 人員配置（参加人数）は適切だったか" $colMain
Add-DonutChart $s12 600 262 184 @(
    @{ pct = 76.9; color = $colCoral },
    @{ pct = 23.1; color = $colBlue }
) "問題あり" "76.9% (10名)"
Add-LegendItem $s12 706 210 218 $colCoral "問題あり" "10名 (76.9%)" $true $colPastelC
Add-LegendItem $s12 706 266 218 $colBlue  "問題なし" "3名 (23.1%)" $false

Add-Rect $s12 20 418 920 100 $colPastelC $colBorder $true 0.12 | Out-Null
Add-Text $s12 40 430 880 78 "・設問⑤（広さ）は5名（38.5%）が「問題あり（常時混雑・待合椅子不足で外出案内・InBodyが狭い）」と回答した。`n・設問⑦（人員配置）は10名（76.9%）が「問題あり（休憩が取れない・誘導係や測定補助が必要）」と回答した（⑨来年参加必要性は84.6%が「はい」）。" 15 $colText $true 1 | Out-Null


# ==============================================================================
# SLIDE 13: 【第2部】スタッフ結果③：自由記述の要点（良かった点 vs 問題点）
# ==============================================================================
$s13 = $pres.Slides.Add(13, 12)
Add-SoftHeader $s13 "第2部：スタッフ結果③" "スタッフ自由記述の要点（良かった点 と 現場の問題点）" 13

Add-SoftCard $s13 20 90 450 428 "良かった点（多職種連携・地域貢献）" $colMain
Add-Rect $s13 38 148 414 350 $colPastelB -1 $true 0.06 | Out-Null
$s13Good = @"
・他部門と同じ目標に向かって一致団結し、普段できないコミュニケーションが取れた。

・多職種と楽しく関われ、当院の結束力を感じた。

・InBody（リハ）の結果をもとに栄養相談や看護指導を行うなど、専門職間の連携がしやすかった。

・当院の認知度の低さ（医療センターと誤認される等）を現場で肌で知る機会になった。
"@
Add-Text $s13 50 162 390 326 $s13Good 16 $colText $false 1 | Out-Null

Add-SoftCard $s13 490 90 450 428 "問題点（広さ・人員・機器・PR）" $colMain
Add-Rect $s13 508 148 414 350 $colPastelC -1 $true 0.06 | Out-Null
$s13Bad = @"
・3区画では狭く満席が続き、座れない方に外出を勧めた。規模を広げるなら4区画が必要。

・InBodyの待ち時間が30分以上となり、最後は並ぶのを中止（お断り）した。2台体制が必要。

・休憩時間がほとんど取れなかった（10〜15分×2回程度）。誘導係・統括・測定の増員が必要。

・広報誌や病院案内パネルを持参せず、リハ病院の役割を十分に伝えきれなかった。
"@
Add-Text $s13 520 162 390 326 $s13Bad 16 $colText $false 1 | Out-Null


# ==============================================================================
# SLIDE 14: 【第3部】起案書のKPI・3つの目的の達成報告
# ==============================================================================
$s14 = $pres.Slides.Add(14, 12)
Add-SoftHeader $s14 "第3部：達成報告" "起案書のKPI（目標数値）と 3つの目的の達成報告" 14

# 上段：2つの大きなKPI達成カード
Add-Rect $s14 20 90 450 166 $colWhite $colBorder $true 0.08 | Out-Null
Add-Rect $s14 38 108 175 34 $colPastelB -1 $true 0.4 | Out-Null
Add-Text $s14 38 112 175 28 "KPI① ブース来場者数" 14.5 $colMain $true 2 | Out-Null
Add-Text $s14 38 150 414 52 "268 名 （達成率 134%）" 28 $colMain $true 1 | Out-Null
Add-Text $s14 38 208 414 32 "起案目標 200名 ／ 内訳：1日目 120名・2日目 148名" 15 $colSubText $true 1 | Out-Null

Add-Rect $s14 490 90 450 166 $colWhite $colBorder $true 0.08 | Out-Null
Add-Rect $s14 508 108 190 34 $colPastelT -1 $true 0.4 | Out-Null
Add-Text $s14 508 112 190 28 "KPI② アンケート回答数" 14.5 $colTeal $true 2 | Out-Null
Add-Text $s14 508 150 414 52 "236 名 （達成率 472%）" 28 $colTeal $true 1 | Out-Null
Add-Text $s14 508 208 414 32 "起案目標 50名の約4.7倍 ／ 内訳：1日目 106名・2日目 130名" 15 $colSubText $true 1 | Out-Null

# 下段：3つの目的の達成カード
Add-SoftCard $s14 20 272 920 246 "起案書における3つの目的の達成状況" $colMain
$s14Purp = @"
① 地域貢献（達成）：268名に測定・健康相談を提供し、スタッフの92.3%が地域貢献を実感した。

② 認知・集患（達成）：初認知67名・名前のみ126名と接点を持ち、90.3%（213名）から「利用・検討したい」との評価を獲得した。

③ 多職種連携の強化（達成）：6部署14名が協働して満足度100%を達成し、測定から栄養・看護・相談へつなぐ多職種連携と院内の結束力を強化できた。
"@
Add-Text $s14 40 330 880 174 $s14Purp 16 $colText $true 1 | Out-Null


# ==============================================================================
# SLIDE 15: 【第3部】今回の出展で得られたこと（6つの短いファクト）
# ==============================================================================
$s15 = $pres.Slides.Add(15, 12)
Add-SoftHeader $s15 "第3部：得られたこと" "今回の出展で得られたこと（事実の整理）" 15

$facts15 = @(
    @{ title = "① 目標を大きく上回る集客"; text = "2日間で268名が来場し、目標50名の4.7倍となる236名の回答を取得できた。"; col = $colMain; bg = $colPastelB },
    @{ title = "② 土日で異なる世代との接点"; text = "土曜は60代以上（52.8%）、日曜は40代以下が倍増（32.3%）し幅広い層と接した。"; col = $colTeal; bg = $colPastelT },
    @{ title = "③ 市内＋市外へのアプローチ"; text = "八千代市内（72.5%）に加え、市外（26.3%・62名）の住民とも接点を持てた。"; col = $colAmber; bg = $colPastelA },
    @{ title = "④ 人気コンテンツの明確化"; text = "血管年齢（72.9%）と筋肉量測定（66.5%）が強い集客力を持つと確認できた。"; col = $colMain; bg = $colPastelB },
    @{ title = "⑤ 9割超の好意的な利用意向"; text = "体験者の90.3%（213名）が、万が一の際に当院を「利用・検討したい」と回答した。"; col = $colTeal; bg = $colPastelT },
    @{ title = "⑥ 健康教室の4大テーマ把握"; text = "今後の知名度向上に向けた健康教室の4大関心テーマ（転倒・栄養・認知・終活）が判明した。"; col = $colAmber; bg = $colPastelA }
)
for ($i = 0; $i -lt 6; $i++) {
    $colIdx = $i % 3
    $rowIdx = [math]::Floor($i / 3)
    $x = 20 + ($colIdx * 310)
    $y = 92 + ($rowIdx * 214)
    $f = $facts15[$i]
    Add-Rect $s15 $x $y 296 198 $colWhite $colBorder $true 0.08 | Out-Null
    Add-Rect $s15 ($x + 16) ($y + 18) 264 38 $f.bg -1 $true 0.4 | Out-Null
    Add-Text $s15 ($x + 16) ($y + 22) 264 30 $f.title 15.5 $f.col $true 2 | Out-Null
    Add-Text $s15 ($x + 20) ($y + 70) 256 114 $f.text 15.5 $colText $false 1 | Out-Null
}


# ==============================================================================
# SLIDE 16: 【第3部：考察①】1日目 vs 2日目の来場者層の違い（2つの大きな円グラフ）
# ==============================================================================
$s16 = $pres.Slides.Add(16, 12)
Add-SoftHeader $s16 "第3部：考察① 日別比較" "1日目（106名）と 2日目（130名）の来場者層の違い" 16

# 左：1日目の年代円グラフ
Add-SoftCard $s16 20 90 450 316 "1日目（土曜・106名）：シニア中心" $colMain
Add-DonutChart $s16 130 262 184 @(
    @{ pct = 19.8; color = $colTeal },
    @{ pct = 26.4; color = $colSlate },
    @{ pct = 52.8; color = $colBlue },
    @{ pct = 0.9;  color = $colLightGray }
) "60代以上" "52.8% (56名)"
Add-LegendItem $s16 236 190 218 $colTeal  "10〜40代" "21名 (19.8%)" $false
Add-LegendItem $s16 236 242 218 $colSlate "50歳代" "28名 (26.4%)" $false
Add-LegendItem $s16 236 294 218 $colBlue  "60〜90代" "56名 (52.8%)" $true $colPastelB

# 右：2日目の年代円グラフ
Add-SoftCard $s16 490 90 450 316 "2日目（日曜・130名）：40代以下が倍増" $colMain
Add-DonutChart $s16 600 262 184 @(
    @{ pct = 32.3; color = $colTeal },
    @{ pct = 26.2; color = $colSlate },
    @{ pct = 40.0; color = $colBlue },
    @{ pct = 1.5;  color = $colLightGray }
) "40代以下" "32.3% (42名)"
Add-LegendItem $s16 706 190 218 $colTeal  "10〜40代" "42名 (32.3%)" $true $colPastelT
Add-LegendItem $s16 706 242 218 $colSlate "50歳代" "34名 (26.2%)" $false
Add-LegendItem $s16 706 294 218 $colBlue  "60〜90代" "52名 (40.0%)" $false

Add-Rect $s16 20 418 920 100 $colPastelB $colBorder $true 0.12 | Out-Null
Add-Text $s16 40 430 880 78 "・1日目は60歳代以上のシニア層が52.8%（56名）と過半数を占めた。`n・2日目は40歳代以下が21名（19.8%）→42名（32.3%）へ倍増し、八千代市内在住者も70.8%→73.8%へ増加した。" 15.5 $colText $true 1 | Out-Null


# ==============================================================================
# SLIDE 17: 【第3部：考察② 当院の弱点1】八千代市内（171名）の認知度（大きな円グラフ）
# ==============================================================================
$s17 = $pres.Slides.Add(17, 12)
Add-SoftHeader $s17 "第3部：考察② 当院の弱点" "アンケートからわかる当院の弱点①：八千代市民の認知度" 17

Add-SoftCard $s17 20 90 920 428 "八千代市内在住者（171名）における当院の認知度（Q2）" $colMain

Add-DonutChart $s17 210 315 270 @(
    @{ pct = 21.1; color = $colCoral },
    @{ pct = 59.6; color = $colAmber },
    @{ pct = 18.1; color = $colBlue },
    @{ pct = 1.2;  color = $colLightGray }
) "詳細不明・未認知" "80.7% (138名)"

Add-LegendItem $s17 410 160 490 $colCoral     "今日初めて知った（未認知）" "36名 (21.1%)" $true $colPastelC
Add-LegendItem $s17 410 215 490 $colAmber     "名前や場所は知っている" "102名 (59.6%)" $true $colPastelA
Add-LegendItem $s17 410 270 490 $colBlue      "よく知っている（利用経験有）" "31名 (18.1%)" $false
Add-LegendItem $s17 410 322 490 $colLightGray "未記入" "2名 (1.2%)" $false

Add-Rect $s17 404 374 506 126 $colPastelC $colBorder $true 0.12 | Out-Null
Add-Text $s17 422 384 472 108 "・当院は入院専門のリハ病院（外来は退院フォローのみ、介護は訪問・通所）のため、日常で住民が接する機会が少ない。`n・地元市民でも5人に1人（21.1%）が「今日初めて知った」と答え、「名前のみ（59.6%）」と合わせ80.7%が詳細を知らない点が弱点である。" 15.5 $colText $true 1 | Out-Null


# ==============================================================================
# SLIDE 18: 【第3部：考察③ 当院の弱点2】八千代市内の「年代別」認知度（最高＝赤、最低＝青、中間＝灰）
# ==============================================================================
$s18 = $pres.Slides.Add(18, 12)
Add-SoftHeader $s18 "第3部：考察③ 世代別の弱点" "アンケートからわかる当院の弱点②：八千代市内の年代別認知度" 18

Add-SoftCard $s18 20 90 566 428 "八千代市内の年代別「初認知＋名前のみ（詳細不明）」割合" $colMain
Add-SoftBar $s18 40 160 175 190 "市内 40歳代 (21名)" "85.7% (18名)" 85.7 $colSlate 15
Add-SoftBar $s18 40 218 175 190 "市内 50歳代 (47名)" "91.5% (43名)" 91.5 $colCoral 15
Add-SoftBar $s18 40 276 175 190 "市内 60歳代 (30名)" "86.7% (26名)" 86.7 $colSlate 15
Add-SoftBar $s18 40 334 175 190 "市内 70歳代 (29名)" "65.5% (19名)" 65.5 $colBlue 15
Add-SoftBar $s18 40 392 175 190 "市内 80歳代 (19名)" "84.2% (16名)" 84.2 $colSlate 15
Add-BarColorRule $s18 38 468 "赤：詳細不明が最多(50代)" $colCoral "青：詳細不明が最少(70代)" $colBlue

Add-SoftCard $s18 602 90 338 428 "世代別データの要点" $colMain
$s18Text = @"
・最多来場層の「市内50歳代」で、詳細を知らない割合が91.5%と最も高い（赤色）。

・70歳代は「よく知っている」が31.0%おり、詳細不明率は65.5%で最も低い（青色）。

・40代・60代・80代でも約85%が当院の中身を知らない状態にある（灰色）。
"@
Add-Text $s18 622 156 298 330 $s18Text 16 $colText $false 1 | Out-Null


# ==============================================================================
# SLIDE 19: 【第3部：考察④ クロス分析】認知度（Q2）× 当院の評価（Q3）（3つの円グラフ）
# ==============================================================================
$s19 = $pres.Slides.Add(19, 12)
Add-SoftHeader $s19 "第3部：考察④ クロス分析" "認知度（Q2）× アンケート時の当院の評価（Q3）" 19

# 1. 今日初めて知った (67名)
Add-SoftCard $s19 20 90 296 316 "今日初めて知った（67名）" $colMain
Add-DonutChart $s19 168 238 164 @(
    @{ pct = 14.9; color = $colTeal },
    @{ pct = 73.1; color = $colBlue },
    @{ pct = 11.9; color = $colSlate }
) "好意的評価" "88.1%"
Add-Rect $s19 36 334 264 56 $colPastelB -1 $true 0.2 | Out-Null
Add-Text $s19 36 338 264 50 "ぜひ利用：14.9%（10名）`n検討したい：73.1%（49名）" 14.5 $colText $true 2 | Out-Null

# 2. 名前や場所は知っている (126名)
Add-SoftCard $s19 332 90 296 316 "名前・場所は知る（126名）" $colMain
Add-DonutChart $s19 480 238 164 @(
    @{ pct = 19.0; color = $colTeal },
    @{ pct = 74.6; color = $colBlue },
    @{ pct = 6.3;  color = $colSlate }
) "好意的評価" "93.7%"
Add-Rect $s19 348 334 264 56 $colPastelB -1 $true 0.2 | Out-Null
Add-Text $s19 348 338 264 50 "ぜひ利用：19.0%（24名）`n検討したい：74.6%（94名）" 14.5 $colText $true 2 | Out-Null

# 3. よく知っている (39名)
Add-SoftCard $s19 644 90 296 316 "よく知っている（39名）" $colMain
Add-DonutChart $s19 792 238 164 @(
    @{ pct = 66.7; color = $colTeal },
    @{ pct = 25.6; color = $colBlue },
    @{ pct = 7.7;  color = $colSlate }
) "ぜひ利用" "66.7%"
Add-Rect $s19 660 334 264 56 $colPastelT -1 $true 0.2 | Out-Null
Add-Text $s19 660 338 264 50 "ぜひ利用：66.7%（26名）`n検討したい：25.6%（10名）" 14.5 $colTeal $true 2 | Out-Null

Add-Rect $s19 20 418 920 100 $colPastelB $colBorder $true 0.12 | Out-Null
Add-Text $s19 40 430 880 78 "・「今日初めて知った」層でも体験後は88.1%、「名前・場所のみ」層では93.7%が利用・検討したいと回答した。`n・さらに「よく知っている」層では66.7%が「ぜひ利用したい」と答えており、当院の中身を知ってもらう機会を増やせば高い評価につながる。" 15.5 $colText $true 1 | Out-Null


# ==============================================================================
# SLIDE 20: 【第4部：来年度提案①】来年度の出展に向けた改善点（3つの柱）
# ==============================================================================
$s20 = $pres.Slides.Add(20, 12)
Add-SoftHeader $s20 "第4部：来年度提案①" "来年度の出展に向けた3つの改善提案" 20

Add-Rect $s20 20 90 296 428 $colWhite $colBorder $true 0.08 | Out-Null
Add-Rect $s20 36 110 264 44 $colPastelC -1 $true 0.4 | Out-Null
Add-Text $s20 36 116 264 32 "① 3区画 ➔ 4区画へ拡大" 17 $colCoral $true 2 | Out-Null
$p20_1 = @"
・3区画では満席が続き、座れない方に外出を案内する機会損失が生じた。

・来年度は「4区画」を確保し、待合椅子と相談スペースを広げる。
"@
Add-Text $s20 38 174 260 310 $p20_1 16 $colText $false 1 | Out-Null

Add-Rect $s20 332 90 296 428 $colWhite $colBorder $true 0.08 | Out-Null
Add-Rect $s20 348 110 264 44 $colPastelA -1 $true 0.4 | Out-Null
Add-Text $s20 348 116 264 32 "② 増員と休憩1時間の確保" 17 $colAmber $true 2 | Out-Null
$p20_2 = @"
・スタッフの76.9%が人員不足（休憩10〜15分のみ）を課題に挙げた。

・各部署の参加人数を増やし「休憩1時間」を確保するとともに、「全体統括」と「誘導係（整理券方式）」を置く。
"@
Add-Text $s20 350 174 260 310 $p20_2 16 $colText $false 1 | Out-Null

Add-Rect $s20 644 90 296 428 $colWhite $colBorder $true 0.08 | Out-Null
Add-Rect $s20 660 110 264 44 $colPastelB -1 $true 0.4 | Out-Null
Add-Text $s20 660 116 264 32 "③ InBody2台・PRパネル" 17 $colMain $true 2 | Out-Null
$p20_3 = @"
・30分待ちで受付中止となったInBodyを「2台体制」にする。

・リハ病院の役割を伝える「広報誌・病院紹介パネル」を持参し、靴ベラ・バインダー・ノベルティを拡充する。
"@
Add-Text $s20 662 174 260 310 $p20_3 16 $colText $false 1 | Out-Null


# ==============================================================================
# SLIDE 21: 【第4部：来年度提案②】今後の知名度向上に向けた「健康教室」アクション
# ==============================================================================
$s21 = $pres.Slides.Add(21, 12)
Add-SoftHeader $s21 "第4部：来年度提案②" "今後の知名度向上に向けたアクション（定期的な健康教室）" 21

$themes21 = @(
    @{ title = "テーマ① 転倒予防・運動 (39.4%)"; desc = "自宅でできるストレッチや筋力維持・バランスチェック（PT/OT）"; col = $colMain; bg = $colPastelB },
    @{ title = "テーマ② 食事・栄養バランス (33.5%)"; desc = "筋肉量を維持する食事や生活習慣病予防の栄養講座（管理栄養士）"; col = $colTeal; bg = $colPastelT },
    @{ title = "テーマ③ 物忘れ・認知機能 (33.1%)"; desc = "認知症予防・フレイル予防のチェックと生活指導（リハ・看護）"; col = $colAmber; bg = $colPastelA },
    @{ title = "テーマ④ 将来の介護や終活 (30.1%)"; desc = "回復期リハ入院や訪問・通所リハの仕組み、介護相談（連携室・薬剤）"; col = $colMain; bg = $colPastelB }
)
for ($i = 0; $i -lt 4; $i++) {
    $colIdx = $i % 2
    $rowIdx = [math]::Floor($i / 2)
    $x = 20 + ($colIdx * 468)
    $y = 92 + ($rowIdx * 150)
    $th = $themes21[$i]
    Add-Rect $s21 $x $y 452 136 $colWhite $colBorder $true 0.1 | Out-Null
    Add-Rect $s21 ($x + 16) ($y + 16) 420 36 $th.bg -1 $true 0.4 | Out-Null
    Add-Text $s21 ($x + 24) ($y + 19) 404 30 $th.title 16.5 $th.col $true 1 | Out-Null
    Add-Text $s21 ($x + 24) ($y + 62) 404 62 $th.desc 15.5 $colText $false 1 | Out-Null
}

Add-Rect $s21 20 398 920 120 $colPastelB $colBorder $true 0.12 | Out-Null
Add-Text $s21 40 412 880 96 "・八千代市民の80.7%が当院を詳しく知らない弱点を補うには、年1回の産業祭だけでなく継続的な接点づくりが必要である。`n・今回Q4で把握した上記4テーマ（いずれも関心度30%超）で定期的に「健康教室」を開催し、当院の役割（入院リハ・訪問・通所）をコツコツと浸透させる。" 15.5 $colText $true 1 | Out-Null


# ==============================================================================
# SLIDE 22: 【第4部：総括まとめ】弱点・得られたこと・来年度の課題（来場者×スタッフ）
# ==============================================================================
$s22 = $pres.Slides.Add(22, 12)
Add-SoftHeader $s22 "第4部：総括まとめ" "まとめ：当院の弱点・出展で得られたこと・来年度の課題" 22

# 1. 当院の弱点
Add-Rect $s22 20 90 920 134 $colWhite $colBorder $true 0.1 | Out-Null
Add-Rect $s22 36 106 190 102 $colPastelC -1 $true 0.2 | Out-Null
Add-Text $s22 42 132 178 56 "1. 当院の弱点`n（知名度の低さ）" 16.5 $colCoral $true 2 | Out-Null
Add-Text $s22 242 106 680 104 "・八千代市民（171名）の80.7%が「今日初めて知った（21.1%）」または「名前のみ（59.6%）」と回答。`n・スタッフも「認知度の低さや他院との誤認」を痛感しており、一般外来がない当院にとって地域での知名度向上が最重要課題である。" 15.5 $colText $true 1 | Out-Null

# 2. 出展で得られたこと
Add-Rect $s22 20 236 920 134 $colWhite $colBorder $true 0.1 | Out-Null
Add-Rect $s22 36 252 190 102 $colPastelT -1 $true 0.2 | Out-Null
Add-Text $s22 42 278 178 56 "2. 出展で`n得られたこと" 16.5 $colTeal $true 2 | Out-Null
Add-Text $s22 242 252 680 104 "・6部署協働で来場268名・回答236名（目標比472%）を達成し、多職種連携の強化（満足度100%）を実現した。`n・初めて知った層でも体験後は88.1%（全体90.3%）が「利用・検討したい」と評価し、対面活動の有効性と健康教室の4大テーマが判明した。" 15.5 $colText $true 1 | Out-Null

# 3. 来年度の課題
Add-Rect $s22 20 382 920 136 $colWhite $colBorder $true 0.1 | Out-Null
Add-Rect $s22 36 398 190 104 $colPastelB -1 $true 0.2 | Out-Null
Add-Text $s22 42 424 178 56 "3. 来年度の課題`n（環境と体制整備）" 16.5 $colMain $true 2 | Out-Null
Add-Text $s22 242 398 680 106 "・来場者の「測定への人気集中・順路や備品の要望」と、スタッフの「3区画の狭さ・InBody受付中止・人員と休憩不足（76.9%）」が一致。`n・来年度は「① 4区画への拡大」「② InBody 2台体制・病院紹介パネル持参」「③ 誘導係・各部署の増員（休憩1時間確保）」を実施する。" 15.5 $colText $true 1 | Out-Null

for ($i = 1; $i -le $pres.Slides.Count; $i++) {
    $imgPath = Join-Path $artifactDir ("slide_v4_" + $i + ".png")
    $pres.Slides.Item($i).Export($imgPath, "PNG", 1920, 1080)
}

$pres.SaveAs($webPath)
$pres.Close()
if ($pp.Presentations.Count -eq 0) {
    $pp.Quit()
}
Start-Sleep -Milliseconds 500
try {
    Copy-Item -Path $webPath -Destination $outPath -Force -ErrorAction Stop
    Write-Host "Updated $outPath directly."
} catch {
    $altPath = "c:\Users\coffe\OneDrive\デスクトップ\Antigravity\アンケート解析\八千代リハビリテーション病院_どーんと祭出展結果報告_四役プレゼン_修正版.pptx"
    Copy-Item -Path $webPath -Destination $altPath -Force
    Write-Host "Original file is open in PowerPoint; saved to $altPath and $webPath"
}
Write-Host "Successfully generated 22-slide unified & meaningful-color PowerPoint!"
