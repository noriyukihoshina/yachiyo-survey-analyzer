$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$outPath = "c:\Users\coffe\OneDrive\デスクトップ\Antigravity\アンケート解析\八千代リハビリテーション病院_どーんと祭出展結果報告_四役プレゼン.pptx"
$webPath = "c:\Users\coffe\OneDrive\デスクトップ\Antigravity\アンケート解析\report_presentation.pptx"
$artifactDir = "C:\Users\coffe\.gemini\antigravity\brain\57f12e92-5bc8-492c-8c7d-8ab9b420cf7a"

$pp = New-Object -ComObject PowerPoint.Application
$pres = $pp.Presentations.Add($false)
$pres.PageSetup.SlideWidth = 960
$pres.PageSetup.SlideHeight = 540

function RGB([int]$r, [int]$g, [int]$b) {
    return [int]($r + ($g * 256) + ($b * 65536))
}

# 柔らかく親しみやすい、かつ視認性の高いカラーパレット
$colBg        = RGB 244 247 251   # 明るく優しいペールブルーグレー背景
$colWhite     = RGB 255 255 255
$colText      = RGB 30 41 59      # 柔らかいチャコールネイビー（真っ黒を避ける）
$colSubText   = RGB 100 116 139   # ミディアムスレート
$colBorder    = RGB 219 228 240   # 優しい枠線色
$colBarBg     = RGB 226 232 240

# 彩り豊かなアクセントカラー（5色＋補助色）
$colBlue      = RGB 37 99 235     # 1. フレッシュブルー（メイン）
$colTeal      = RGB 16 185 129    # 2. エメラルドティール（好評価・達成）
$colAmber     = RGB 245 158 11    # 3. ウォームアンバー（注目・新規）
$colCoral     = RGB 239 68 68     # 4. コーラルレッド（弱点・課題強調）
$colPurple    = RGB 139 92 246    # 5. ソフトバイオレット（第4要素・多職種）
$colSky       = RGB 14 165 233    # 補助スカイブルー
$colSlate     = RGB 148 163 184   # 補助グレー（未記入・その他用）

# パステル背景色（丸角カード用）
$colPastelB   = RGB 239 246 255
$colPastelT   = RGB 236 253 245
$colPastelA   = RGB 255 251 235
$colPastelC   = RGB 254 242 242
$colPastelP   = RGB 245 243 255

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

# 柔らかく親しみやすいヘッダー（白ベース＋丸みのあるカラーピルバッジ）
function Add-SoftHeader($slide, [string]$badgeText, [string]$titleText, [string]$subLead, [int]$pageNo, [int]$themeColor = $colBlue, [int]$pastelColor = $colPastelB) {
    Add-Rect $slide 0 0 960 540 $colBg -1 $false | Out-Null
    # 上部ソフトヘッダーカード
    Add-Rect $slide 20 14 920 96 $colWhite $colBorder $true 0.12 | Out-Null
    # ピル型バッジ
    Add-Rect $slide 36 26 184 32 $themeColor -1 $true 0.5 | Out-Null
    Add-Text $slide 36 28 184 28 $badgeText 12.5 $colWhite $true 2 | Out-Null
    # タイトル（大きく見やすく）
    Add-Text $slide 230 22 620 38 $titleText 20 $colText $true 1 | Out-Null
    # ページ番号ピル
    Add-Rect $slide 856 26 68 30 $pastelColor -1 $true 0.5 | Out-Null
    Add-Text $slide 856 28 68 26 ("$pageNo / $totalSlides") 12.5 $themeColor $true 2 | Out-Null

    # サブリード（端的な1行メッセージ）
    Add-Rect $slide 36 66 888 34 $pastelColor -1 $true 0.3 | Out-Null
    Add-Text $slide 48 69 868 28 $subLead 15 $themeColor $true 1 | Out-Null
}

# 丸みのある柔らかい白カード
function Add-SoftCard($slide, [single]$left, [single]$top, [single]$width, [single]$height, [string]$title = "", [int]$accentColor = $colBlue, [int]$bgRGB = $colWhite) {
    Add-Rect $slide $left $top $width $height $bgRGB $colBorder $true 0.06 | Out-Null
    if ($title -ne "") {
        # 丸みのあるアクセントバー（左横）
        Add-Rect $slide ($left + 18) ($top + 18) 8 26 $accentColor -1 $true 0.5 | Out-Null
        Add-Text $slide ($left + 34) ($top + 15) ($width - 50) 32 $title 17.5 $colText $true 1 | Out-Null
    }
}

# ドーナツ円グラフ描画関数（パッと見て直感的にわかる大きな円グラフ）
function Add-DonutChart($slide, [single]$cx, [single]$cy, [single]$diameter, [array]$slices, [string]$centerTopText, [string]$centerSubText) {
    $left = $cx - ($diameter / 2.0)
    $top  = $cy - ($diameter / 2.0)
    # ベース円（100%単一スライス時の対策も兼ねる）
    $startAngle = -90.0
    foreach ($s in $slices) {
        $pct = [double]$s.pct
        if ($pct -le 0.1) { continue }
        $sweep = ($pct / 100.0) * 360.0
        if ($pct -ge 99.9) {
            $ov = Add-Oval $slide $left $top $diameter $diameter $s.color $colWhite
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
    # 中央の白穴（ドーナツホール）
    $holeD = $diameter * 0.58
    $holeL = $cx - ($holeD / 2.0)
    $holeT = $cy - ($holeD / 2.0)
    Add-Oval $slide $holeL $holeT $holeD $holeD $colWhite -1 | Out-Null
    if ($centerTopText -ne "") {
        Add-Text $slide ($cx - 75) ($cy - 22) 150 26 $centerTopText 13.5 $colText $true 2 | Out-Null
    }
    if ($centerSubText -ne "") {
        Add-Text $slide ($cx - 75) ($cy + 2) 150 24 $centerSubText 12.5 $colSubText $true 2 | Out-Null
    }
}

# 円グラフ用の見やすい凡例アイテム（カラー丸ドット＋項目名＋大きな数値）
function Add-LegendItem($slide, [single]$left, [single]$top, [single]$width, [int]$color, [string]$label, [string]$valText, [bool]$highlight = $false) {
    if ($highlight) {
        Add-Rect $slide ($left - 6) ($top - 4) ($width + 12) 38 $colPastelA $colAmber $true 0.25 | Out-Null
    }
    Add-Oval $slide $left ($top + 5) 16 16 $color -1 | Out-Null
    Add-Text $slide ($left + 22) $top 220 28 $label 14 $colText $true 1 | Out-Null
    Add-Text $slide ($left + $width - 155) $top 155 28 $valText 14.5 $color $true 3 | Out-Null
}

# 丸みのある横棒グラフ（複数選択や年代別用）
function Add-SoftBar($slide, [single]$left, [single]$top, [single]$labelW, [single]$maxBarW, [string]$label, [string]$valText, [single]$pctVal, [int]$barColor, [single]$fontSize = 15) {
    Add-Text $slide $left $top $labelW 28 $label $fontSize $colText $true 1 | Out-Null
    $barX = $left + $labelW + 8
    Add-Rect $slide $barX ($top + 5) $maxBarW 20 $colBarBg -1 $true 0.5 | Out-Null
    $fillW = [math]::Max(10, [math]::Round($maxBarW * ($pctVal / 100.0)))
    Add-Rect $slide $barX ($top + 5) $fillW 20 $barColor -1 $true 0.5 | Out-Null
    Add-Text $slide ($barX + $maxBarW + 10) $top 145 28 $valText $fontSize $barColor $true 1 | Out-Null
}

function Style-SoftTable($tbl, [int]$rows, [int]$cols, [int]$headerBg = $colBlue, [single]$fontSize = 14.5) {
    for ($r = 1; $r -le $rows; $r++) {
        for ($c = 1; $c -le $cols; $c++) {
            $cell = $tbl.Cell($r, $c)
            $tr = $cell.Shape.TextFrame.TextRange
            $tr.Font.Name = "Meiryo UI"
            $tr.Font.NameFarEast = "Meiryo UI"
            $tr.Font.Size = [single]$fontSize
            if ($r -eq 1) {
                $cell.Shape.Fill.ForeColor.RGB = [int]$headerBg
                $tr.Font.Color.RGB = [int]$colWhite
                $tr.Font.Bold = -1
            } else {
                $cell.Shape.Fill.ForeColor.RGB = if ($r % 2 -eq 0) { [int]$colWhite } else { [int]$colPastelB }
                $tr.Font.Color.RGB = [int]$colText
            }
        }
    }
}

# ==============================================================================
# SLIDE 1: 表紙 ＆ 起案の3つの目的（丸みのある柔らかいデザイン）
# ==============================================================================
$s1 = $pres.Slides.Add(1, 12)
Add-Rect $s1 0 0 960 540 $colBg -1 $false | Out-Null

# ヒーローバナー（丸角）
Add-Rect $s1 24 20 912 168 $colBlue -1 $true 0.08 | Out-Null
# 表紙の柔らかいアイコン（丸＋医療クロス）
Add-Oval $s1 52 48 68 68 $colWhite -1 | Out-Null
Add-Rect $s1 78 62 16 40 $colTeal -1 $true 0.2 | Out-Null
Add-Rect $s1 66 74 40 16 $colTeal -1 $true 0.2 | Out-Null

Add-Text $s1 140 38 760 28 "八千代リハビリテーション病院 ｜ 四役報告資料（2026年10月17日・18日 開催）" 15 (RGB 219 234 254) $true 1 | Out-Null
Add-Text $s1 140 68 760 52 "八千代市産業祭「どーんと祭」出展結果報告書" 29 $colWhite $true 1 | Out-Null
Add-Text $s1 140 128 760 44 "ブース来場者：268名（1日目120名・2日目148名） ／ アンケート有効回答：236名`n出展規模：3区画・6部署協働（参加スタッフ14名） ／ 報告者：リハビリテーション部 保科憲幸" 14 (RGB 239 246 255) $false 1 | Out-Null

Add-Text $s1 32 204 896 30 "■ 起案書における3つの出展目的" 17.5 $colText $true 1 | Out-Null

# 3つの目的カード（文字数を絞って大きく）
Add-Rect $s1 24 242 292 272 $colWhite $colBorder $true 0.08 | Out-Null
Add-Rect $s1 40 258 260 38 $colPastelB -1 $true 0.4 | Out-Null
Add-Text $s1 40 263 260 30 "目的① 地域貢献" 17 $colBlue $true 2 | Out-Null
Add-Text $s1 40 312 260 180 "多職種による健康測定・相談の機会を提供し、地域医療への貢献と住民の健康寿命延伸に寄与する。" 16 $colText $false 1 | Out-Null

Add-Rect $s1 334 242 292 272 $colWhite $colBorder $true 0.08 | Out-Null
Add-Rect $s1 350 258 260 38 $colPastelA -1 $true 0.4 | Out-Null
Add-Text $s1 350 263 260 30 "目的② 認知・集患" 17 $colAmber $true 2 | Out-Null
Add-Text $s1 350 312 260 180 "「顔が見える医療機関」として認知を確立し、将来入院を要する際に第一選択として思い浮かぶ下地を作る。" 16 $colText $false 1 | Out-Null

Add-Rect $s1 644 242 292 272 $colWhite $colBorder $true 0.08 | Out-Null
Add-Rect $s1 660 258 260 38 $colPastelT -1 $true 0.4 | Out-Null
Add-Text $s1 660 263 260 30 "目的③ 多職種連携の強化" 17 $colTeal $true 2 | Out-Null
Add-Text $s1 660 312 260 180 "6部署が協働してイベントを企画・運営することで院内の多職種連携を強化し、日常業務の連携基盤を強固にする。" 16 $colText $false 1 | Out-Null


# ==============================================================================
# SLIDE 2: 【第1部】来場者アンケートの質問事項と見えること（短く端的に）
# ==============================================================================
$s2 = $pres.Slides.Add(2, 12)
Add-SoftHeader $s2 "第1部：来場者調査" "来場者アンケートの質問事項とそこから見えること" "基本情報とQ1〜Q5の6項目から、今回の出展評価・現在の認知度・当院の評価・健康課題を把握" 2 $colBlue $colPastelB

$t2 = $s2.Shapes.AddTable(7, 3, 24, 124, 912, 392).Table
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
Style-SoftTable $t2 7 3 $colBlue 15
for ($r = 2; $r -le 7; $r++) {
    $t2.Cell($r, 1).Shape.TextFrame.TextRange.Font.Bold = -1
    $t2.Cell($r, 3).Shape.TextFrame.TextRange.Font.Bold = -1
    $t2.Cell($r, 3).Shape.TextFrame.TextRange.Font.Color.RGB = $colBlue
}


# ==============================================================================
# SLIDE 3: 【第1部】基本情報①：お住まい ＆ 性別（2つの大きな円グラフ）
# ==============================================================================
$s3 = $pres.Slides.Add(3, 12)
Add-SoftHeader $s3 "第1部：来場者結果①" "基本情報①：お住まい と 性別（N=236名）" "お住まいは八千代市内が72.5%（市外26.3%）。性別は女性が64.8%を占めた" 3 $colBlue $colPastelB

# 左カード：お住まい円グラフ
Add-SoftCard $s3 20 122 450 288 "お住まいの内訳" $colBlue
Add-DonutChart $s3 145 280 190 @(
    @{ pct = 72.5; color = $colBlue },
    @{ pct = 26.3; color = $colTeal },
    @{ pct = 1.3;  color = $colSlate }
) "市内 72.5%" "171名 / 236名"

Add-LegendItem $s3 255 210 195 $colBlue  "八千代市内" "171名 (72.5%)" $false
Add-LegendItem $s3 255 258 195 $colTeal  "八千代市外" "62名 (26.3%)" $false
Add-LegendItem $s3 255 306 195 $colSlate "未記入" "3名 (1.3%)" $false

# 右カード：性別円グラフ
Add-SoftCard $s3 490 122 450 288 "性別の内訳" $colAmber
Add-DonutChart $s3 615 280 190 @(
    @{ pct = 64.8; color = $colAmber },
    @{ pct = 34.3; color = $colBlue },
    @{ pct = 0.8;  color = $colSlate }
) "女性 64.8%" "153名 / 236名"

Add-LegendItem $s3 725 210 195 $colAmber "女性" "153名 (64.8%)" $false
Add-LegendItem $s3 725 258 195 $colBlue  "男性" "81名 (34.3%)" $false
Add-LegendItem $s3 725 306 195 $colSlate "未記入" "2名 (0.8%)" $false

# 下部：端的な解説ボックス（2行のみ）
Add-Rect $s3 20 422 920 96 $colWhite $colBorder $true 0.12 | Out-Null
Add-Text $s3 40 434 880 76 "・お住まいは「八千代市内」が171名（72.5%）、「八千代市外」が62名（26.3%）であった。`n・性別は「女性」が153名（64.8%）、「男性」が81名（34.3%）であった。" 16.5 $colText $true 1 | Out-Null


# ==============================================================================
# SLIDE 4: 【第1部】基本情報②：年齢別の内訳（大きなグラフ＋短い解説）
# ==============================================================================
$s4 = $pres.Slides.Add(4, 12)
Add-SoftHeader $s4 "第1部：来場者結果②" "基本情報②：年齢別の内訳（N=236名）" "50歳代が26.3%（62名）で最多。次いで70歳代（17.8%）、60歳代（16.5%）、40歳代（14.0%）" 4 $colBlue $colPastelB

Add-SoftCard $s4 20 122 560 396 "年代別の回答数・割合" $colBlue
$ages4 = @(
    @("10〜20歳代", "18名 (7.6%)", 7.6, $colSky),
    @("30歳代", "12名 (5.1%)", 5.1, $colSky),
    @("40歳代", "33名 (14.0%)", 14.0, $colTeal),
    @("50歳代", "62名 (26.3%)", 26.3, $colBlue),
    @("60歳代", "39名 (16.5%)", 16.5, $colBlue),
    @("70歳代", "42名 (17.8%)", 17.8, $colAmber),
    @("80〜90歳代", "27名 (11.4%)", 11.4, $colPurple),
    @("未記入", "3名 (1.3%)", 1.3, $colSlate)
)
for ($i = 0; $i -lt $ages4.Count; $i++) {
    $r = $ages4[$i]
    Add-SoftBar $s4 40 (178 + $i * 40) 115 245 $r[0] $r[1] ($r[2] * 3.2) $r[3] 15
}

Add-SoftCard $s4 596 122 344 396 "結果のポイント" $colTeal
$s4Text = @"
・50歳代が62名（26.3%）で最も多い。

・次いで70歳代が42名（17.8%）、60歳代が39名（16.5%）、40歳代が33名（14.0%）となった。

・50〜70歳代で全体の60.6%（143名）を占める一方、40歳代以下も26.7%（63名）が来場した。
"@
Add-Text $s4 616 184 304 310 $s4Text 16.5 $colText $false 1 | Out-Null


# ==============================================================================
# SLIDE 5: 【第1部】Q1. 本日参考になった内容（今回の出展の評価）
# ==============================================================================
$s5 = $pres.Slides.Add(5, 12)
Add-SoftHeader $s5 "第1部：来場者結果③" "Q1. 本日参考になった内容（今回の出展の評価）" "「血管年齢測定（72.9%）」と「筋肉量測定（66.5%）」に回答者の約7割の評価が集まった" 5 $colBlue $colPastelB

Add-SoftCard $s5 20 122 570 396 "Q1. 参考になった内容（複数回答・N=236）" $colBlue
Add-SoftBar $s5 40 190 185 215 "血管年齢測定" "172名 (72.9%)" 72.9 $colBlue 15.5
Add-SoftBar $s5 40 250 185 215 "筋肉量測定(InBody)" "157名 (66.5%)" 66.5 $colTeal 15.5
Add-SoftBar $s5 40 310 185 215 "栄養レッスン" "57名 (24.2%)" 24.2 $colAmber 15.5
Add-SoftBar $s5 40 370 185 215 "人生会議(終活)の案内" "41名 (17.4%)" 17.4 $colPurple 15.5
Add-SoftBar $s5 40 430 185 215 "お薬相談" "32名 (13.6%)" 13.6 $colSky 15.5

Add-SoftCard $s5 606 122 334 396 "結果の説明" $colTeal
$s5Text = @"
・1位は「血管年齢測定」で172名（72.9%）。

・2位は「筋肉量測定」で157名（66.5%）。

・以下、「栄養レッスン」57名（24.2%）、「人生会議の案内」41名（17.4%）、「お薬相談」32名（13.6%）であった。
"@
Add-Text $s5 626 184 294 310 $s5Text 16.5 $colText $false 1 | Out-Null


# ==============================================================================
# SLIDE 6: 【第1部】Q2. これまで当院をご存知でしたか？（大きな円グラフ）
# ==============================================================================
$s6 = $pres.Slides.Add(6, 12)
Add-SoftHeader $s6 "第1部：来場者結果④" "Q2. これまで当院をご存知でしたか？（現在の認知度）" "「名前や場所は知っている」が53.4%、「今日初めて知った」が28.4%、「よく知っている」は16.5%" 6 $colAmber $colPastelA

Add-SoftCard $s6 20 122 920 396 "Q2. 当院の認知度（単一回答・N=236）" $colAmber

Add-DonutChart $s6 210 335 260 @(
    @{ pct = 53.4; color = $colBlue },
    @{ pct = 28.4; color = $colAmber },
    @{ pct = 16.5; color = $colTeal },
    @{ pct = 1.7;  color = $colSlate }
) "初認知 28.4%" "名前のみ 53.4%"

Add-LegendItem $s6 410 190 490 $colBlue  "名前や場所は知っている" "126名 (53.4%)" $false
Add-LegendItem $s6 410 242 490 $colAmber "今日初めて知った" "67名 (28.4%)" $true
Add-LegendItem $s6 410 294 490 $colTeal  "よく知っている（利用経験有）" "39名 (16.5%)" $false
Add-LegendItem $s6 410 342 490 $colSlate "未記入" "4名 (1.7%)" $false

Add-Rect $s6 404 394 506 102 $colPastelB -1 $true 0.12 | Out-Null
Add-Text $s6 420 404 476 86 "・半数以上の126名（53.4%）が「名前や場所は知っている」と回答した。`n・67名（28.4%）が「今日初めて知った」、39名（16.5%）が「よく知っている」と回答した。" 15.5 $colText $true 1 | Out-Null


# ==============================================================================
# SLIDE 7: 【第1部】Q3. 万が一の際、当院を利用したいですか？（大きな円グラフ）
# ==============================================================================
$s7 = $pres.Slides.Add(7, 12)
Add-SoftHeader $s7 "第1部：来場者結果⑤" "Q3. 万が一の際、当院を利用したいですか？（当院の評価）" "「ぜひ利用したい（25.4%）」と「候補として検討したい（64.8%）」を合わせ、90.3%が好意的評価" 7 $colTeal $colPastelT

Add-SoftCard $s7 20 122 920 396 "Q3. 万が一の際の利用意向（単一回答・N=236）" $colTeal

Add-DonutChart $s7 210 335 260 @(
    @{ pct = 25.4; color = $colTeal },
    @{ pct = 64.8; color = $colBlue },
    @{ pct = 7.2;  color = $colSlate },
    @{ pct = 2.5;  color = $colBarBg }
) "好意的評価" "90.3% (213名)"

Add-LegendItem $s7 410 190 490 $colTeal  "ぜひ利用したい" "60名 (25.4%)" $true
Add-LegendItem $s7 410 242 490 $colBlue  "候補の一つとして検討したい" "153名 (64.8%)" $false
Add-LegendItem $s7 410 294 490 $colSlate "他の病院を選ぶ・わからない" "17名 (7.2%)" $false
Add-LegendItem $s7 410 342 490 $colMuted "未記入" "6名 (2.5%)" $false

Add-Rect $s7 404 394 506 102 $colPastelT -1 $true 0.12 | Out-Null
Add-Text $s7 420 404 476 86 "・「候補として検討したい」153名（64.8%）と「ぜひ利用したい」60名（25.4%）を合わせ、計213名（90.3%）が前向きな利用意向を示した。" 16 $colText $true 1 | Out-Null


# ==============================================================================
# SLIDE 8: 【第1部】Q4. 現在健康で気になること（健康課題のチェック）
# ==============================================================================
$s8 = $pres.Slides.Add(8, 12)
Add-SoftHeader $s8 "第1部：来場者結果⑥" "Q4. 現在健康について気になること（課題・関心のチェック）" "「転倒（39.4%）」「栄養（33.5%）」「認知機能（33.1%）」「介護・終活（30.1%）」の4項目が3割超" 8 $colBlue $colPastelB

Add-SoftCard $s8 20 122 570 396 "Q4. 健康について気になること（複数回答・N=236）" $colBlue
Add-SoftBar $s8 40 180 175 220 "足腰の衰え・転倒" "93名 (39.4%)" 39.4 $colBlue 15
Add-SoftBar $s8 40 230 175 220 "食事・栄養バランス" "79名 (33.5%)" 33.5 $colTeal 15
Add-SoftBar $s8 40 280 175 220 "物忘れ・認知機能" "78名 (33.1%)" 33.1 $colAmber 15
Add-SoftBar $s8 40 330 175 220 "将来の介護や終活" "71名 (30.1%)" 30.1 $colPurple 15
Add-SoftBar $s8 40 380 175 220 "とくにない" "34名 (14.4%)" 14.4 $colSlate 15
Add-SoftBar $s8 40 430 175 220 "薬の飲み方 ／ その他" "各11名 (4.7%)" 4.7 $colMuted 15

Add-SoftCard $s8 606 122 334 396 "結果の説明" $colBlue
$s8Text = @"
・上位4項目（足腰の衰え・転倒、食事・栄養、物忘れ・認知機能、将来の介護や終活）がいずれも30%を超えた。

・「その他（11名）」には、高血圧、糖尿病、ひざの痛み、五十肩などが記載された。
"@
Add-Text $s8 626 184 294 310 $s8Text 16.5 $colText $false 1 | Out-Null


# ==============================================================================
# SLIDE 9: 【第1部】Q5. 来年に期待する内容（自由記述39件の内訳）
# ==============================================================================
$s9 = $pres.Slides.Add(9, 12)
Add-SoftHeader $s9 "第1部：来場者結果⑦" "Q5. 来年に期待する内容（来年度出展内容のヒント）" "「運動・転倒・骨密度チェック（12件）」「今年度と同様の測定継続（11件）」「認知機能チェック（6件）」が中心" 9 $colBlue $colPastelB

# 5つのカラフルな丸角カードで視覚的に整理
Add-SoftCard $s9 20 124 450 186 "① 運動・転倒・骨密度のチェック（12件）" $colBlue $colWhite
Add-Text $s9 38 176 414 120 "・「椅子に座ってできるストレッチや運動を知りたい」`n・「骨粗しょう症・骨密度や転倒危険度のチェック」" 15.5 $colText $false 1 | Out-Null

Add-SoftCard $s9 490 124 450 186 "② 今年度と同様の測定継続・好評の声（11件）" $colTeal $colWhite
Add-Text $s9 508 176 414 120 "・「筋肉量や血管年齢など今年と同じ測定を続けてほしい」`n・「マンションの入院経験者5名からリハが良かったと聞いた」" 15.5 $colText $false 1 | Out-Null

Add-SoftCard $s9 20 326 296 190 "③ 認知機能・フレイル（6件）" $colAmber $colWhite
Add-Text $s9 36 378 264 124 "・「認知機能の衰えチェックや体力測定、フレイル予防を知りたい」" 15 $colText $false 1 | Out-Null

Add-SoftCard $s9 332 326 296 190 "④ 介護・薬・受診相談（4件）" $colPurple $colWhite
Add-Text $s9 348 378 264 124 "・「介護生活の相談、家族が入院した際の薬、症状別の受診科」" 15 $colText $false 1 | Out-Null

Add-SoftCard $s9 644 326 296 190 "⑤ 備品・案内順路の要望（3件）" $colCoral $colWhite
Add-Text $s9 660 378 264 124 "・「記入用バインダー、InBodyの靴入れ・靴ベラ、順路表示」" 15 $colText $false 1 | Out-Null


# ==============================================================================
# SLIDE 10: 【第2部】スタッフアンケートの質問事項と見えること
# ==============================================================================
$s10 = $pres.Slides.Add(10, 12)
Add-SoftHeader $s10 "第2部：スタッフ調査" "スタッフアンケートの質問事項とそこから見えること" "参加した6部署・全14名（現在13名回答済）を対象に、出展意義・広さ・人員配置・改善点を調査" 10 $colTeal $colPastelT

$t10 = $s10.Shapes.AddTable(7, 3, 24, 124, 912, 392).Table
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
Style-SoftTable $t10 7 3 $colTeal 15
for ($r = 2; $r -le 7; $r++) {
    $t10.Cell($r, 1).Shape.TextFrame.TextRange.Font.Bold = -1
    $t10.Cell($r, 3).Shape.TextFrame.TextRange.Font.Bold = -1
    $t10.Cell($r, 3).Shape.TextFrame.TextRange.Font.Color.RGB = $colTeal
}


# ==============================================================================
# SLIDE 11: 【第2部】スタッフ結果①：①参加満足度 ＆ ③地域貢献（2つの大きな円グラフ）
# ==============================================================================
$s11 = $pres.Slides.Add(11, 12)
Add-SoftHeader $s11 "第2部：スタッフ結果①" "スタッフ集計①：参加満足度（①）と 地域貢献度（③）" "回答13名全員（100%）が「参加して良かった」、92.3%（12名）が「地域貢献できた」と回答" 11 $colTeal $colPastelT

Add-SoftCard $s11 20 122 450 288 "① 参加して良かったか（N=13）" $colTeal
Add-DonutChart $s11 145 280 190 @(
    @{ pct = 100.0; color = $colTeal }
) "はい 100%" "13名 / 13名"
Add-LegendItem $s11 255 230 195 $colTeal  "はい" "13名 (100%)" $true
Add-LegendItem $s11 255 282 195 $colSlate "いいえ" "0名 (0.0%)" $false

Add-SoftCard $s11 490 122 450 288 "③ 地域住民に貢献できたか（N=13）" $colBlue
Add-DonutChart $s11 615 280 190 @(
    @{ pct = 92.3; color = $colBlue },
    @{ pct = 7.7;  color = $colSlate }
) "はい 92.3%" "12名 / 13名"
Add-LegendItem $s11 725 230 195 $colBlue  "はい" "12名 (92.3%)" $true
Add-LegendItem $s11 725 282 195 $colSlate "いいえ" "1名 (7.7%)" $false

Add-Rect $s11 20 422 920 96 $colWhite $colBorder $true 0.12 | Out-Null
Add-Text $s11 40 432 880 78 "・設問①は13名全員（100%）が「はい」と回答し、地域住民との直接交流や他部門との一致団結・結束力が理由に挙げられた。`n・設問③は12名（92.3%）が「はい」と回答し、測定結果に基づく運動・栄養・生活指導やお薬カレンダー配布が評価された。" 15.5 $colText $true 1 | Out-Null


# ==============================================================================
# SLIDE 12: 【第2部】スタッフ結果②：⑤3区画の広さ ＆ ⑦人員配置（2つの大きな円グラフ）
# ==============================================================================
$s12 = $pres.Slides.Add(12, 12)
Add-SoftHeader $s12 "第2部：スタッフ結果②" "スタッフ集計②：3区画の広さ（⑤）と 人員配置（⑦）" "3区画の広さは38.5%（5名）、人員配置は76.9%（10名）が「問題あり」と回答" 12 $colCoral $colPastelC

Add-SoftCard $s12 20 122 450 288 "⑤ 3区画の広さは適切だったか" $colAmber
Add-DonutChart $s12 145 280 190 @(
    @{ pct = 61.5; color = $colBlue },
    @{ pct = 38.5; color = $colCoral }
) "問題あり" "38.5% (5名)"
Add-LegendItem $s12 255 230 195 $colBlue  "問題なし" "8名 (61.5%)" $false
Add-LegendItem $s12 255 282 195 $colCoral "問題あり" "5名 (38.5%)" $true

Add-SoftCard $s12 490 122 450 288 "⑦ 人員配置（参加人数）は適切だったか" $colCoral
Add-DonutChart $s12 615 280 190 @(
    @{ pct = 76.9; color = $colCoral },
    @{ pct = 23.1; color = $colBlue }
) "問題あり" "76.9% (10名)"
Add-LegendItem $s12 725 230 195 $colCoral "問題あり" "10名 (76.9%)" $true
Add-LegendItem $s12 725 282 195 $colBlue  "問題なし" "3名 (23.1%)" $false

Add-Rect $s12 20 422 920 96 $colWhite $colBorder $true 0.12 | Out-Null
Add-Text $s12 40 432 880 78 "・設問⑤（広さ）は5名（38.5%）が「問題あり（常時混雑・待合椅子不足で外出案内・InBodyが狭い）」と回答した。`n・設問⑦（人員配置）は10名（76.9%）が「問題あり（休憩が取れない・誘導係や測定補助が必要）」と回答した（⑨来年参加必要性は84.6%が「はい」）。" 15 $colText $true 1 | Out-Null


# ==============================================================================
# SLIDE 13: 【第2部】スタッフ結果③：自由記述の要点（良かった点 vs 問題点）
# ==============================================================================
$s13 = $pres.Slides.Add(13, 12)
Add-SoftHeader $s13 "第2部：スタッフ結果③" "スタッフ自由記述の要点（良かった点 と 現場の問題点）" "「多職種の一致団結・連携」が高く評価された一方、「狭さ・InBody待ち・休憩不足」が課題に" 13 $colTeal $colPastelT

Add-SoftCard $s13 20 124 450 392 "良かった点（多職種連携・地域貢献）" $colTeal
$s13Good = @"
・他部門と同じ目標に向かって一致団結し、普段できないコミュニケーションが取れた。

・多職種と楽しく関われ、当院の結束力を感じた。

・InBody（リハ）の結果をもとに栄養相談や看護指導を行うなど、専門職間の連携がしやすかった。

・当院の認知度の低さ（医療センターと誤認される等）を現場で肌で知る機会になった。
"@
Add-Text $s13 40 180 410 320 $s13Good 15.5 $colText $false 1 | Out-Null

Add-SoftCard $s13 490 124 450 392 "問題点（広さ・人員・機器・PR）" $colCoral
$s13Bad = @"
・3区画では狭く満席が続き、座れない方に外出を勧めた。規模を広げるなら4区画が必要。

・InBodyの待ち時間が30分以上となり、最後は並ぶのを中止（お断り）した。2台体制が必要。

・休憩時間がほとんど取れなかった（10〜15分×2回程度）。誘導係・統括・測定の増員が必要。

・広報誌や病院案内パネルを持参せず、リハ病院の役割を十分に伝えきれなかった。
"@
Add-Text $s13 510 180 410 320 $s13Bad 15.5 $colText $false 1 | Out-Null


# ==============================================================================
# SLIDE 14: 【第3部】起案書のKPI・3つの目的の達成報告
# ==============================================================================
$s14 = $pres.Slides.Add(14, 12)
Add-SoftHeader $s14 "第3部：達成報告" "起案書のKPI（目標数値）と 3つの目的の達成報告" "ブース来場268名（目標比134%）・アンケート236名（目標比472%）を達成し、3目的すべてを達成" 14 $colBlue $colPastelB

# 上段：2つの大きなKPI達成カード
Add-Rect $s14 20 124 450 156 $colWhite $colBorder $true 0.08 | Out-Null
Add-Rect $s14 36 140 160 32 $colPastelB -1 $true 0.4 | Out-Null
Add-Text $s14 36 143 160 26 "KPI① ブース来場者数" 14 $colBlue $true 2 | Out-Null
Add-Text $s14 36 180 420 50 "268 名 （達成率 134%）" 28 $colBlue $true 1 | Out-Null
Add-Text $s14 36 234 420 32 "起案目標 200名 ／ 内訳：1日目 120名・2日目 148名" 14.5 $colSubText $true 1 | Out-Null

Add-Rect $s14 490 124 450 156 $colWhite $colBorder $true 0.08 | Out-Null
Add-Rect $s14 506 140 180 32 $colPastelT -1 $true 0.4 | Out-Null
Add-Text $s14 506 143 180 26 "KPI② アンケート回答数" 14 $colTeal $true 2 | Out-Null
Add-Text $s14 506 180 420 50 "236 名 （達成率 472%）" 28 $colTeal $true 1 | Out-Null
Add-Text $s14 506 234 420 32 "起案目標 50名の約4.7倍 ／ 内訳：1日目 106名・2日目 130名" 14.5 $colSubText $true 1 | Out-Null

# 下段：3つの目的の達成カード
Add-SoftCard $s14 20 294 920 222 "起案書における3つの目的の達成状況" $colBlue
$s14Purp = @"
① 地域貢献（達成）：268名に測定・健康相談を提供し、スタッフの92.3%が地域貢献を実感した。

② 認知・集患（達成）：初認知67名・名前のみ126名と接点を持ち、90.3%（213名）から「利用・検討したい」との評価を獲得した。

③ 多職種連携の強化（達成）：6部署14名が協働して満足度100%を達成し、測定から栄養・看護・相談へつなぐ多職種連携と院内の結束力を強化できた。
"@
Add-Text $s14 40 348 880 156 $s14Purp 15.5 $colText $true 1 | Out-Null


# ==============================================================================
# SLIDE 15: 【第3部】今回の出展で得られたこと（6つの短いファクト）
# ==============================================================================
$s15 = $pres.Slides.Add(15, 12)
Add-SoftHeader $s15 "第3部：得られたこと" "今回の出展で得られたこと（事実の整理）" "2日間の出展を通じて得られた6つの客観的ファクト" 15 $colTeal $colPastelT

$facts15 = @(
    @{ title = "① 目標を大きく上回る集客"; text = "2日間で268名が来場し、目標50名の4.7倍となる236名の回答を取得できた。"; col = $colBlue; bg = $colPastelB },
    @{ title = "② 土日で異なる世代との接点"; text = "土曜は60代以上（52.8%）、日曜は40代以下が倍増（32.3%）し幅広い層と接した。"; col = $colTeal; bg = $colPastelT },
    @{ title = "③ 市内＋市外へのアプローチ"; text = "八千代市内（72.5%）に加え、市外（26.3%・62名）の住民とも接点を持てた。"; col = $colAmber; bg = $colPastelA },
    @{ title = "④ 人気コンテンツの明確化"; text = "血管年齢（72.9%）と筋肉量測定（66.5%）が強い集客力を持つと確認できた。"; col = $colPurple; bg = $colPastelP },
    @{ title = "⑤ 9割超の好意的な利用意向"; text = "体験者の90.3%（213名）が、万が一の際に当院を「利用・検討したい」と回答した。"; col = $colTeal; bg = $colPastelT },
    @{ title = "⑥ 健康教室の4大テーマ把握"; text = "今後の知名度向上に向けた健康教室の4大関心テーマ（転倒・栄養・認知・終活）が判明した。"; col = $colBlue; bg = $colPastelB }
)
for ($i = 0; $i -lt 6; $i++) {
    $colIdx = $i % 3
    $rowIdx = [math]::Floor($i / 3)
    $x = 20 + ($colIdx * 310)
    $y = 126 + ($rowIdx * 196)
    $f = $facts15[$i]
    Add-Rect $s15 $x $y 296 180 $colWhite $colBorder $true 0.08 | Out-Null
    Add-Rect $s15 ($x + 16) ($y + 16) 264 36 $f.bg -1 $true 0.4 | Out-Null
    Add-Text $s15 ($x + 16) ($y + 20) 264 28 $f.title 15.5 $f.col $true 2 | Out-Null
    Add-Text $s15 ($x + 18) ($y + 64) 260 104 $f.text 15 $colText $false 1 | Out-Null
}


# ==============================================================================
# SLIDE 16: 【第3部：考察①】1日目 vs 2日目の来場者層の違い（2つの大きな円グラフ）
# ==============================================================================
$s16 = $pres.Slides.Add(16, 12)
Add-SoftHeader $s16 "第3部：考察① 日別比較" "1日目（106名）と 2日目（130名）の来場者層の違い" "1日目は60代以上のシニア層が過半数（52.8%）。2日目は40代以下の若年・子育て層が倍増（32.3%）" 16 $colAmber $colPastelA

# 左：1日目の年代円グラフ
Add-SoftCard $s16 20 122 450 288 "1日目（土曜・106名）：シニア中心" $colBlue
Add-DonutChart $s16 140 280 186 @(
    @{ pct = 19.8; color = $colTeal },
    @{ pct = 26.4; color = $colAmber },
    @{ pct = 52.8; color = $colBlue },
    @{ pct = 0.9;  color = $colSlate }
) "60代以上" "52.8% (56名)"
Add-LegendItem $s16 245 205 205 $colTeal  "10〜40歳代" "21名 (19.8%)" $false
Add-LegendItem $s16 245 253 205 $colAmber "50歳代" "28名 (26.4%)" $false
Add-LegendItem $s16 245 301 205 $colBlue  "60〜90歳代" "56名 (52.8%)" $true

# 右：2日目の年代円グラフ
Add-SoftCard $s16 490 122 450 288 "2日目（日曜・130名）：40代以下が倍増" $colTeal
Add-DonutChart $s16 610 280 186 @(
    @{ pct = 32.3; color = $colTeal },
    @{ pct = 26.2; color = $colAmber },
    @{ pct = 40.0; color = $colBlue },
    @{ pct = 1.5;  color = $colSlate }
) "40代以下" "32.3% (42名)"
Add-LegendItem $s16 715 205 205 $colTeal  "10〜40歳代" "42名 (32.3%)" $true
Add-LegendItem $s16 715 253 205 $colAmber "50歳代" "34名 (26.2%)" $false
Add-LegendItem $s16 715 301 205 $colBlue  "60〜90歳代" "52名 (40.0%)" $false

Add-Rect $s16 20 422 920 96 $colWhite $colBorder $true 0.12 | Out-Null
Add-Text $s16 40 432 880 78 "・2日目は40歳代以下が21名（19.8%）→42名（32.3%）へ倍増した（特に40歳代は9名→24名へ2.7倍）。`n・お住まいも2日目は八千代市内が70.8%（75名）→73.8%（96名）へ増加し、地元の家族連れや若年層にリーチできた。" 15.5 $colText $true 1 | Out-Null


# ==============================================================================
# SLIDE 17: 【第3部：考察② 当院の弱点1】八千代市内（171名）の認知度（大きな円グラフ）
# ==============================================================================
$s17 = $pres.Slides.Add(17, 12)
Add-SoftHeader $s17 "第3部：考察② 当院の弱点" "アンケートからわかる当院の弱点①：八千代市民の認知度" "地元・八千代市民（171名）でも「よく知っている」は18.1%のみ。80.7%が当院の中身を知らない" 17 $colCoral $colPastelC

Add-SoftCard $s17 20 122 920 396 "八千代市内在住者（171名）における当院の認知度（Q2）" $colCoral

Add-DonutChart $s17 210 335 260 @(
    @{ pct = 21.1; color = $colCoral },
    @{ pct = 59.6; color = $colAmber },
    @{ pct = 18.1; color = $colTeal },
    @{ pct = 1.2;  color = $colSlate }
) "詳細不明・未認知" "80.7% (138名)"

Add-LegendItem $s17 410 186 490 $colCoral "今日初めて知った（未認知）" "36名 (21.1%)" $true
Add-LegendItem $s17 410 238 490 $colAmber "名前や場所は知っている" "102名 (59.6%)" $true
Add-LegendItem $s17 410 290 490 $colTeal  "よく知っている（利用経験有）" "31名 (18.1%)" $false
Add-LegendItem $s17 410 338 490 $colSlate "未記入" "2名 (1.2%)" $false

Add-Rect $s17 404 384 506 116 $colPastelC $colCoral $true 0.12 | Out-Null
Add-Text $s17 420 392 476 100 "・当院は入院専門のリハ病院（外来は退院フォローのみ、介護は訪問・通所）のため、日常で住民が接する機会が少ない。`n・地元市民でも5人に1人（21.1%）が「今日初めて知った」と答え、「名前のみ（59.6%）」と合わせ80.7%が詳細を知らない点が弱点である。" 15 $colText $true 1 | Out-Null


# ==============================================================================
# SLIDE 18: 【第3部：考察③ 当院の弱点2】八千代市内の「年代別」認知度
# ==============================================================================
$s18 = $pres.Slides.Add(18, 12)
Add-SoftHeader $s18 "第3部：考察③ 世代別の弱点" "アンケートからわかる当院の弱点②：八千代市内の年代別認知度" "家族の入院先を選ぶ市内40〜60代や、当事者の市内80代でも「初認知・名前のみ」が8〜9割を占める" 18 $colCoral $colPastelC

Add-SoftCard $s18 20 122 570 396 "八千代市内の主要年代における「初認知＋名前のみ」割合" $colCoral
Add-SoftBar $s18 40 190 175 210 "市内 40歳代 (21名)" "85.7% (初認知33%)" 85.7 $colCoral 15
Add-SoftBar $s18 40 255 175 210 "市内 50歳代 (47名)" "91.5% (名前のみ79%)" 91.5 $colCoral 15
Add-SoftBar $s18 40 320 175 210 "市内 60歳代 (30名)" "86.7% (初認知20%)" 86.7 $colAmber 15
Add-SoftBar $s18 40 385 175 210 "市内 80歳代 (19名)" "84.2% (初認知32%)" 84.2 $colCoral 15
Add-Text $s18 40 448 530 48 "※数値は「今日初めて知った」＋「名前や場所のみ知っている」の合計割合（＝当院を詳しく知らない層の割合）" 13.5 $colSubText $true 1 | Out-Null

Add-SoftCard $s18 606 122 334 396 "世代別データの要点" $colCoral
$s18Text = @"
・親の介護・入院先を決める市内40歳代の33.3%（3人に1人）が「今日初めて知った」と回答。

・最多層の市内50歳代でも「よく知っている」は8.5%のみで、91.5%が中身を知らない。

・当事者世代の市内80歳代でも31.6%が「今日初めて知った」と回答している。
"@
Add-Text $s18 626 184 294 310 $s18Text 16 $colText $false 1 | Out-Null


# ==============================================================================
# SLIDE 19: 【第3部：考察④ クロス分析】認知度（Q2）× 当院の評価（Q3）（3つの円グラフ）
# ==============================================================================
$s19 = $pres.Slides.Add(19, 12)
Add-SoftHeader $s19 "第3部：考察④ クロス分析" "認知度（Q2）× アンケート時の当院の評価（Q3）" "初めて知った層でも体験後は88.1%が好意的評価。「よく知る」層は66.7%が「ぜひ利用したい」と回答" 19 $colTeal $colPastelT

# 3つの円グラフカードを横並び
# 1. 今日初めて知った (67名)
Add-SoftCard $s19 20 122 296 288 "今日初めて知った（67名）" $colAmber
Add-DonutChart $s19 168 258 156 @(
    @{ pct = 14.9; color = $colTeal },
    @{ pct = 73.1; color = $colBlue },
    @{ pct = 11.9; color = $colSlate }
) "好意的評価" "88.1%"
Add-Text $s19 32 346 272 54 "ぜひ利用：14.9%（10名）`n検討したい：73.1%（49名）" 14.5 $colText $true 2 | Out-Null

# 2. 名前や場所は知っている (126名)
Add-SoftCard $s19 332 122 296 288 "名前・場所は知る（126名）" $colBlue
Add-DonutChart $s19 480 258 156 @(
    @{ pct = 19.0; color = $colTeal },
    @{ pct = 74.6; color = $colBlue },
    @{ pct = 6.3;  color = $colSlate }
) "好意的評価" "93.7%"
Add-Text $s19 344 346 272 54 "ぜひ利用：19.0%（24名）`n検討したい：74.6%（94名）" 14.5 $colText $true 2 | Out-Null

# 3. よく知っている (39名)
Add-SoftCard $s19 644 122 296 288 "よく知っている（39名）" $colTeal
Add-DonutChart $s19 792 258 156 @(
    @{ pct = 66.7; color = $colTeal },
    @{ pct = 25.6; color = $colBlue },
    @{ pct = 7.7;  color = $colSlate }
) "ぜひ利用" "66.7%"
Add-Text $s19 656 346 272 54 "ぜひ利用：66.7%（26名）`n検討したい：25.6%（10名）" 14.5 $colTeal $true 2 | Out-Null

Add-Rect $s19 20 422 920 96 $colWhite $colBorder $true 0.12 | Out-Null
Add-Text $s19 40 432 880 78 "・「今日初めて知った」層でも体験後は88.1%、「名前・場所のみ」層では93.7%が利用・検討したいと回答した。`n・さらに「よく知っている」層では66.7%が「ぜひ利用したい」と答えており、当院の中身を知ってもらう機会を増やせば高い評価につながる。" 15.5 $colText $true 1 | Out-Null


# ==============================================================================
# SLIDE 20: 【第4部：来年度提案①】来年度の出展に向けた改善点（3つの柱）
# ==============================================================================
$s20 = $pres.Slides.Add(20, 12)
Add-SoftHeader $s20 "第4部：来年度提案①" "来年度の出展に向けた3つの改善提案" "今年度の混雑・InBody受付中止・スタッフ休憩不足を解消するため、4区画化と増員を提案" 20 $colBlue $colPastelB

Add-Rect $s20 20 124 296 392 $colWhite $colBorder $true 0.08 | Out-Null
Add-Rect $s20 36 142 264 42 $colPastelC -1 $true 0.4 | Out-Null
Add-Text $s20 36 148 264 32 "① 3区画 ➔ 4区画へ拡大" 17 $colCoral $true 2 | Out-Null
$p20_1 = @"
・3区画では満席が続き、座れない方に外出を案内する機会損失が生じた。

・来年度は「4区画」を確保し、待合椅子と相談スペースを広げる。
"@
Add-Text $s20 38 202 260 290 $p20_1 16 $colText $false 1 | Out-Null

Add-Rect $s20 332 124 296 392 $colWhite $colBorder $true 0.08 | Out-Null
Add-Rect $s20 348 142 264 42 $colPastelA -1 $true 0.4 | Out-Null
Add-Text $s20 348 148 264 32 "② 増員と休憩1時間の確保" 17 $colAmber $true 2 | Out-Null
$p20_2 = @"
・スタッフの76.9%が人員不足（休憩10〜15分のみ）を課題に挙げた。

・各部署の参加人数を増やし「休憩1時間」を確保するとともに、「全体統括」と「誘導係（整理券方式）」を置く。
"@
Add-Text $s20 350 202 260 290 $p20_2 16 $colText $false 1 | Out-Null

Add-Rect $s20 644 124 296 392 $colWhite $colBorder $true 0.08 | Out-Null
Add-Rect $s20 660 142 264 42 $colPastelB -1 $true 0.4 | Out-Null
Add-Text $s20 660 148 264 32 "③ InBody2台・PRパネル" 17 $colBlue $true 2 | Out-Null
$p20_3 = @"
・30分待ちで受付中止となったInBodyを「2台体制」にする。

・リハ病院の役割を伝える「広報誌・病院紹介パネル」を持参し、靴ベラ・バインダー・ノベルティを拡充する。
"@
Add-Text $s20 662 202 260 290 $p20_3 16 $colText $false 1 | Out-Null


# ==============================================================================
# SLIDE 21: 【第4部：来年度提案②】今後の知名度向上に向けた「健康教室」アクション
# ==============================================================================
$s21 = $pres.Slides.Add(21, 12)
Add-SoftHeader $s21 "第4部：来年度提案②" "今後の知名度向上に向けたアクション（定期的な健康教室）" "Q4で判明した地域住民の4大関心テーマ（各30%超）を活かし、健康教室でコツコツ認知を広げる" 21 $colTeal $colPastelT

# 4つのテーマカード
$themes21 = @(
    @{ title = "テーマ① 転倒予防・運動 (39.4%)"; desc = "自宅でできるストレッチや筋力維持・バランスチェック（PT/OT）"; col = $colBlue; bg = $colPastelB },
    @{ title = "テーマ② 食事・栄養バランス (33.5%)"; desc = "筋肉量を維持する食事や生活習慣病予防の栄養講座（管理栄養士）"; col = $colTeal; bg = $colPastelT },
    @{ title = "テーマ③ 物忘れ・認知機能 (33.1%)"; desc = "認知症予防・フレイル予防のチェックと生活指導（リハ・看護）"; col = $colAmber; bg = $colPastelA },
    @{ title = "テーマ④ 将来の介護や終活 (30.1%)"; desc = "回復期リハ入院や訪問・通所リハの仕組み、介護相談（連携室・薬剤）"; col = $colPurple; bg = $colPastelP }
)
for ($i = 0; $i -lt 4; $i++) {
    $colIdx = $i % 2
    $rowIdx = [math]::Floor($i / 2)
    $x = 20 + ($colIdx * 468)
    $y = 124 + ($rowIdx * 138)
    $th = $themes21[$i]
    Add-Rect $s21 $x $y 452 124 $colWhite $colBorder $true 0.1 | Out-Null
    Add-Rect $s21 ($x + 16) ($y + 14) 420 34 $th.bg -1 $true 0.4 | Out-Null
    Add-Text $s21 ($x + 24) ($y + 17) 404 28 $th.title 16 $th.col $true 1 | Out-Null
    Add-Text $s21 ($x + 24) ($y + 56) 404 58 $th.desc 15 $colText $false 1 | Out-Null
}

Add-Rect $s21 20 406 920 110 $colWhite $colTeal $true 0.12 | Out-Null
Add-Text $s21 40 418 880 88 "・八千代市民の80.7%が当院を詳しく知らない弱点を補うには、年1回の産業祭だけでなく継続的な接点づくりが必要である。`n・今回Q4で把握した上記4テーマで定期的に「健康教室」を開催し、当院の役割（入院リハ・訪問・通所）をコツコツと浸透させる。" 15.5 $colText $true 1 | Out-Null


# ==============================================================================
# SLIDE 22: 【第4部：総括まとめ】弱点・得られたこと・来年度の課題（来場者×スタッフ）
# ==============================================================================
$s22 = $pres.Slides.Add(22, 12)
Add-SoftHeader $s22 "第4部：総括まとめ" "まとめ：当院の弱点・出展で得られたこと・来年度の課題" "来場者アンケート（236名）とスタッフアンケート（13名）を掛け合わせた最終結論" 22 $colBlue $colPastelB

# 1. 当院の弱点
Add-Rect $s22 20 124 920 122 $colWhite $colBorder $true 0.1 | Out-Null
Add-Rect $s22 34 138 190 94 $colPastelC -1 $true 0.2 | Out-Null
Add-Text $s22 40 160 178 56 "1. 当院の弱点`n（知名度の低さ）" 16.5 $colCoral $true 2 | Out-Null
Add-Text $s22 240 136 684 100 "・八千代市民（171名）の80.7%が「今日初めて知った（21.1%）」または「名前のみ（59.6%）」と回答。`n・スタッフも「認知度の低さや他院との誤認」を痛感しており、一般外来がない当院にとって地域での知名度向上が最重要課題である。" 15 $colText $true 1 | Out-Null

# 2. 出展で得られたこと
Add-Rect $s22 20 258 920 122 $colWhite $colBorder $true 0.1 | Out-Null
Add-Rect $s22 34 272 190 94 $colPastelT -1 $true 0.2 | Out-Null
Add-Text $s22 40 294 178 56 "2. 出展で`n得られたこと" 16.5 $colTeal $true 2 | Out-Null
Add-Text $s22 240 270 684 100 "・6部署協働で来場268名・回答236名（目標比472%）を達成し、多職種連携の強化（満足度100%）を実現した。`n・初めて知った層でも体験後は88.1%（全体90.3%）が「利用・検討したい」と評価し、対面活動の有効性と健康教室の4大テーマが判明した。" 15 $colText $true 1 | Out-Null

# 3. 来年度の課題
Add-Rect $s22 20 392 920 124 $colWhite $colBorder $true 0.1 | Out-Null
Add-Rect $s22 34 406 190 96 $colPastelB -1 $true 0.2 | Out-Null
Add-Text $s22 40 428 178 56 "3. 来年度の課題`n（環境と体制整備）" 16.5 $colBlue $true 2 | Out-Null
Add-Text $s22 240 404 684 102 "・来場者の「測定への人気集中・順路や備品の要望」と、スタッフの「3区画の狭さ・InBody受付中止・人員と休憩不足（76.9%）」が一致。`n・来年度は「① 4区画への拡大」「② InBody 2台体制・病院紹介パネル持参」「③ 誘導係・各部署の増員（休憩1時間確保）」を実施する。" 15 $colText $true 1 | Out-Null

$pres.SaveAs($outPath)
$pres.SaveAs($webPath)

for ($i = 1; $i -le $pres.Slides.Count; $i++) {
    $imgPath = Join-Path $artifactDir ("slide_v3_" + $i + ".png")
    $pres.Slides.Item($i).Export($imgPath, "PNG", 1920, 1080)
}

$pres.Close()
$pp.Quit()
Write-Host "Successfully generated 22-slide soft & colorful PowerPoint: $outPath"
