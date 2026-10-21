$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$outPath = "c:\Users\coffe\OneDrive\デスクトップ\Antigravity\アンケート解析\アンケート結果報告.pptx"
$webPath = "c:\Users\coffe\OneDrive\デスクトップ\Antigravity\アンケート解析\survey_result_handout.pptx"
$artifactDir = "C:\Users\coffe\.gemini\antigravity\brain\57f12e92-5bc8-492c-8c7d-8ab9b420cf7a"

$pp = New-Object -ComObject PowerPoint.Application
foreach ($openP in @($pp.Presentations)) {
    if ($openP.Path -eq "" -or $openP.Name -like "*アンケート結果報告*" -or $openP.Name -like "*survey_result_handout*") {
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
# カラーパレット定義
# ・第1部（来場者アンケート結果 P.2〜P.6）：メインカラー＝ロイヤルブルー
# ・第2部（スタッフアンケート結果 P.7〜P.12）：メインカラー＝エメラルドティール
# ・棒グラフ配色（4役用と統一）：デフォルト＝グレー、多い所＝青、少ない所＝赤
# ==============================================================================
$colBg          = RGB 244 247 252
$colWhite       = RGB 255 255 255
$colText        = RGB 30 41 59
$colSubText     = RGB 100 116 139
$colBorder      = RGB 216 226 238
$colBarBg       = RGB 233 238 245

# 2種類のメインカラー（来場者＝ブルー、スタッフ＝ティールグリーン）
$colVisitorMain = RGB 37 99 235     # 来場者結果メインカラー（ロイヤルブルー）
$colVisitorSoft = RGB 239 246 255   # 来場者パステルブルー
$colStaffMain   = RGB 16 163 127    # スタッフ結果メインカラー（エメラルドティール）
$colStaffSoft   = RGB 236 253 245   # スタッフパステルティール

# グラフ＆サブカラー（柔らかみのあるトーン）
$colBlue        = RGB 37 99 235     # 多い所・肯定（ブルー）
$colTeal        = RGB 16 163 127    # スタッフ肯定・好評価（ティール）
$colCoral       = RGB 225 75 75     # 少ない所・問題あり（ソフトコーラルレッド）
$colAmber       = RGB 234 138 25    # 補助アクセント（ソフトアンバー）
$colSlate       = RGB 148 163 184   # デフォルト・中間・標準（ウォームグレー）
$colLightGray   = RGB 203 213 225   # 未記入（ライトグレー）

# 柔らかさを演出するパステル背景色
$colPastelB     = RGB 239 246 255
$colPastelT     = RGB 236 253 245
$colPastelA     = RGB 254 249 235
$colPastelC     = RGB 254 242 242
$colPastelG     = RGB 248 250 252

$totalSlides = 12

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

function Add-Text($slide, [single]$left, [single]$top, [single]$width, [single]$height, [string]$text, [single]$fontSize = 15, [int]$colorRGB = $colText, [bool]$bold = $false, [int]$align = 1) {
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

# パート別メインカラーで統一された1段シンプルヘッダー（来場者＝ブルー、スタッフ＝ティール）
function Add-SimpleHeader($slide, [string]$badgeText, [string]$titleText, [int]$pageNo, [int]$themeColor, [int]$pastelColor) {
    Add-Rect $slide 0 0 960 540 $colBg -1 $false | Out-Null
    Add-Rect $slide 20 14 920 62 $colWhite $colBorder $true 0.18 | Out-Null
    Add-Rect $slide 34 27 182 36 $themeColor -1 $true 0.5 | Out-Null
    Add-Text $slide 34 30 182 30 $badgeText 13 $colWhite $true 2 | Out-Null
    $tSize = if ($titleText.Length -ge 31) { 18.0 } else { 20.0 }
    Add-Text $slide 226 25 622 40 $titleText $tSize $colText $true 1 | Out-Null
    Add-Rect $slide 854 29 72 32 $pastelColor -1 $true 0.5 | Out-Null
    Add-Text $slide 854 31 72 28 ("$pageNo / $totalSlides") 12.5 $themeColor $true 2 | Out-Null
}

function Add-SoftCard($slide, [single]$left, [single]$top, [single]$width, [single]$height, [string]$title = "", [int]$accentColor = $colVisitorMain, [int]$bgRGB = $colWhite) {
    Add-Rect $slide $left $top $width $height $bgRGB $colBorder $true 0.05 | Out-Null
    if ($title -ne "") {
        Add-Rect $slide ($left + 16) ($top + 16) 8 26 $accentColor -1 $true 0.5 | Out-Null
        Add-Text $slide ($left + 32) ($top + 13) ($width - 46) 32 $title 16.5 $colText $true 1 | Out-Null
        Add-Rect $slide ($left + 16) ($top + 48) ($width - 32) 1 $colBarBg -1 $false | Out-Null
    }
}

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
        Add-Text $slide ($cx - 75) ($cy + 2) 150 24 $centerSubText 12.5 $colSubText $true 2 | Out-Null
    }
}

function Add-LegendRow($slide, [single]$left, [single]$top, [single]$width, [int]$color, [string]$label, [string]$valText, [bool]$highlight = $false, [int]$hlBg = $colPastelB) {
    if ($highlight) {
        Add-Rect $slide ($left - 8) ($top - 3) ($width + 16) 34 $hlBg -1 $true 0.3 | Out-Null
    }
    Add-Oval $slide $left ($top + 6) 15 15 $color -1 | Out-Null
    Add-Text $slide ($left + 22) $top ($width - 145) 28 $label 14.5 $colText $true 1 | Out-Null
    $valColor = if ($color -eq $colLightGray) { $colSubText } else { $color }
    Add-Text $slide ($left + $width - 145) $top 145 28 $valText 14.5 $valColor $true 3 | Out-Null
}

# 意味のある色付けを行う丸角横棒グラフ（デフォルト＝グレー、多い所＝青、少ない所＝赤）
function Add-SoftBar($slide, [single]$left, [single]$top, [single]$labelW, [single]$maxBarW, [string]$label, [string]$valText, [single]$pctVal, [int]$barColor, [single]$fontSize = 14) {
    $isEmph = ($barColor -eq $colBlue -or $barColor -eq $colCoral)
    $lblColor = if ($isEmph) { $colText } else { $colSubText }
    $valColor = if ($isEmph) { $barColor } else { $colText }
    Add-Text $slide $left $top $labelW 26 $label $fontSize $lblColor $isEmph 1 | Out-Null
    $barX = $left + $labelW + 6
    Add-Rect $slide $barX ($top + 5) $maxBarW 18 $colBarBg -1 $true 0.5 | Out-Null
    $fillW = [math]::Max(8, [math]::Round($maxBarW * ($pctVal / 100.0)))
    Add-Rect $barX ($top + 5) $fillW 18 $barColor -1 $true 0.5 | Out-Null
}

# Add-SoftBarの安全定義
function Add-MeaningBar($slide, [single]$left, [single]$top, [single]$labelW, [single]$maxBarW, [string]$label, [string]$valText, [single]$pctVal, [int]$barColor, [single]$fontSize = 14) {
    $isEmph = ($barColor -eq $colBlue -or $barColor -eq $colCoral)
    $lblColor = if ($isEmph) { $colText } else { $colSubText }
    $valColor = if ($isEmph) { $barColor } else { $colText }
    Add-Text $slide $left $top $labelW 26 $label $fontSize $lblColor $isEmph 1 | Out-Null
    $barX = $left + $labelW + 6
    Add-Rect $slide $barX ($top + 5) $maxBarW 18 $colBarBg -1 $true 0.5 | Out-Null
    $fillW = [math]::Max(8, [math]::Round($maxBarW * ($pctVal / 100.0)))
    Add-Rect $slide $barX ($top + 5) $fillW 18 $barColor -1 $true 0.5 | Out-Null
    Add-Text $slide ($barX + $maxBarW + 8) $top 135 26 $valText $fontSize $valColor $true 1 | Out-Null
}

function Add-MiniBarRule($slide, [single]$left, [single]$top, [single]$width, [string]$highText, [string]$lowText) {
    Add-Rect $slide $left $top $width 30 $colPastelG $colBorder $true 0.4 | Out-Null
    Add-Oval $slide ($left + 10) ($top + 9) 12 12 $colBlue -1 | Out-Null
    Add-Text $slide ($left + 24) ($top + 3) 145 24 $highText 11.5 $colBlue $true 1 | Out-Null
    Add-Oval $slide ($left + 168) ($top + 9) 12 12 $colCoral -1 | Out-Null
    Add-Text $slide ($left + 182) ($top + 3) 135 24 $lowText 11.5 $colCoral $true 1 | Out-Null
    Add-Oval $slide ($left + 316) ($top + 9) 12 12 $colSlate -1 | Out-Null
    Add-Text $slide ($left + 330) ($top + 3) 95 24 "灰：中間・標準" 11.5 $colSubText $true 1 | Out-Null
}

# ==============================================================================
# SLIDE 1: 表紙「アンケート結果報告」（来場者＝ブルー、スタッフ＝ティールの2色構成を明示）
# ==============================================================================
$s1 = $pres.Slides.Add(1, 12)
Add-Rect $s1 0 0 960 540 $colBg -1 $false | Out-Null
Add-Rect $s1 28 28 904 196 $colVisitorMain -1 $true 0.08 | Out-Null

Add-Oval $s1 60 64 72 72 $colWhite -1 | Out-Null
Add-Rect $s1 88 78 16 44 $colStaffMain -1 $true 0.2 | Out-Null
Add-Rect $s1 74 92 44 16 $colStaffMain -1 $true 0.2 | Out-Null

Add-Text $s1 154 50 740 32 "八千代リハビリテーション病院 ｜ 2026年10月17日（土）・18日（日）開催" 16 (RGB 219 234 254) $true 1 | Out-Null
Add-Text $s1 154 88 740 64 "八千代市産業祭「どーんと祭」`nアンケート結果報告" 30 $colWhite $true 1 | Out-Null

# 左：第1部 来場者アンケート（ロイヤルブルー）
Add-Rect $s1 28 248 440 262 $colWhite $colBorder $true 0.06 | Out-Null
Add-Rect $s1 46 266 404 44 $colVisitorSoft -1 $true 0.4 | Out-Null
Add-Text $s1 46 273 404 32 "■ 1. 来場者アンケート集計（青テーマ）" 17 $colVisitorMain $true 2 | Out-Null
$s1Left = @"
・ブース来場者数：計 268 名
  （1日目：120名 ／ 2日目：148名）

・アンケート有効回答数：計 236 名
  （1日目：106名 ／ 2日目：130名）

・掲載ページ：P.2 〜 P.6
"@
Add-Text $s1 52 326 392 170 $s1Left 16 $colText $true 1 | Out-Null

# 右：第2部 スタッフアンケート（エメラルドティール）
Add-Rect $s1 492 248 440 262 $colWhite $colBorder $true 0.06 | Out-Null
Add-Rect $s1 510 266 404 44 $colStaffSoft -1 $true 0.4 | Out-Null
Add-Text $s1 510 273 404 32 "■ 2. スタッフアンケート集計（緑テーマ）" 17 $colStaffMain $true 2 | Out-Null
$s1Right = @"
・対象スタッフ：計 14 名（6部署）

・回答完了数：計 13 名

・掲載ページ：P.7 〜 P.12
"@
Add-Text $s1 516 326 392 170 $s1Right 16 $colText $true 1 | Out-Null


# ==============================================================================
# SLIDE 2: 【来場者＝青】来場者① 基本情報：お住まい（円グラフ）＆ 性別（円グラフ）
# ==============================================================================
$s2 = $pres.Slides.Add(2, 12)
Add-SimpleHeader $s2 "来場者アンケート①" "基本情報：お住まい ／ 性別（N=236）" 2 $colVisitorMain $colVisitorSoft

# 左：お住まい円グラフ
Add-SoftCard $s2 20 88 450 432 "お住まい（単一回答・N=236）" $colVisitorMain
Add-DonutChart $s2 245 248 204 @(
    @{ pct = 72.5; color = $colBlue },
    @{ pct = 26.3; color = $colTeal },
    @{ pct = 1.3;  color = $colLightGray }
) "八千代市内" "72.5% (171名)"
Add-LegendRow $s2 55 378 380 $colBlue      "八千代市内" "171名 (72.5%)" $true $colPastelB
Add-LegendRow $s2 55 420 380 $colTeal      "八千代市外" "62名 (26.3%)" $false
Add-LegendRow $s2 55 460 380 $colLightGray "未記入" "3名 (1.3%)" $false

# 右：性別円グラフ
Add-SoftCard $s2 490 88 450 432 "性別（単一回答・N=236）" $colVisitorMain
Add-DonutChart $s2 715 248 204 @(
    @{ pct = 64.8; color = $colAmber },
    @{ pct = 34.3; color = $colBlue },
    @{ pct = 0.8;  color = $colLightGray }
) "女性" "64.8% (153名)"
Add-LegendRow $s2 525 378 380 $colAmber     "女性" "153名 (64.8%)" $true $colPastelA
Add-LegendRow $s2 525 420 380 $colBlue      "男性" "81名 (34.3%)" $false
Add-LegendRow $s2 525 460 380 $colLightGray "未記入" "2名 (0.8%)" $false


# ==============================================================================
# SLIDE 3: 【来場者＝青】来場者② 基本情報：年齢（棒グラフ）＆ Q1. 参考になった内容（棒グラフ）
# ==============================================================================
$s3 = $pres.Slides.Add(3, 12)
Add-SimpleHeader $s3 "来場者アンケート②" "基本情報：年齢 ／ Q1. 本日参考になった内容（N=236）" 3 $colVisitorMain $colVisitorSoft

# 左：年齢棒グラフ（最多50代＝青、最少90代＝赤、その他＝グレー）
Add-SoftCard $s3 20 88 450 432 "年齢（単一回答・N=236）" $colVisitorMain
$ageRows = @(
    @("10歳代", "9名 (3.8%)", 3.8, $colSlate),
    @("20歳代", "9名 (3.8%)", 3.8, $colSlate),
    @("30歳代", "12名 (5.1%)", 5.1, $colSlate),
    @("40歳代", "33名 (14.0%)", 14.0, $colSlate),
    @("50歳代", "62名 (26.3%)", 26.3, $colBlue),
    @("60歳代", "39名 (16.5%)", 16.5, $colSlate),
    @("70歳代", "42名 (17.8%)", 17.8, $colSlate),
    @("80歳代", "25名 (10.6%)", 10.6, $colSlate),
    @("90歳代", "2名 (0.8%)", 0.8, $colCoral),
    @("未記入", "3名 (1.3%)", 1.3, $colLightGray)
)
for ($i = 0; $i -lt $ageRows.Count; $i++) {
    $r = $ageRows[$i]
    Add-MeaningBar $s3 36 (144 + $i * 32) 76 190 $r[0] $r[1] ($r[2] * 3.2) $r[3] 13.5
}
Add-MiniBarRule $s3 32 474 426 "青：最多 (50歳代)" "赤：最少 (90歳代)"

# 右：Q1 棒グラフ（上位2項目＝青、最少＝赤、中間＝グレー）
Add-SoftCard $s3 490 88 450 432 "Q1. 本日参考になった内容（複数回答・N=236）" $colVisitorMain
$q1Rows = @(
    @("血管年齢測定", "172名 (72.9%)", 72.9, $colBlue),
    @("筋肉量測定", "157名 (66.5%)", 66.5, $colBlue),
    @("栄養レッスン", "57名 (24.2%)", 24.2, $colSlate),
    @("人生会議(終活)の案内", "41名 (17.4%)", 17.4, $colSlate),
    @("お薬相談", "32名 (13.6%)", 13.6, $colCoral)
)
for ($i = 0; $i -lt $q1Rows.Count; $i++) {
    $r = $q1Rows[$i]
    Add-MeaningBar $s3 506 (164 + $i * 56) 155 130 $r[0] $r[1] $r[2] $r[3] 14
}
Add-MiniBarRule $s3 502 474 426 "青：上位 (6割超)" "赤：最少項目"


# ==============================================================================
# SLIDE 4: 【来場者＝青】来場者③ Q2. 当院の認知度（円グラフ）＆ Q3. 利用意向（円グラフ）
# ==============================================================================
$s4 = $pres.Slides.Add(4, 12)
Add-SimpleHeader $s4 "来場者アンケート③" "Q2. これまで当院をご存知でしたか？ ／ Q3. 万が一の際、利用したいですか？" 4 $colVisitorMain $colVisitorSoft

# 左：Q2 円グラフ
Add-SoftCard $s4 20 88 450 432 "Q2. これまで当院をご存知でしたか？（N=236）" $colVisitorMain
Add-DonutChart $s4 245 240 196 @(
    @{ pct = 53.4; color = $colBlue },
    @{ pct = 28.4; color = $colCoral },
    @{ pct = 16.5; color = $colTeal },
    @{ pct = 1.7;  color = $colLightGray }
) "名前・場所は知る" "53.4% (126名)"
Add-LegendRow $s4 42 356 406 $colBlue      "名前や場所は知っている" "126名 (53.4%)" $true $colPastelB
Add-LegendRow $s4 42 396 406 $colCoral     "今日初めて知った" "67名 (28.4%)" $true $colPastelC
Add-LegendRow $s4 42 436 406 $colTeal      "よく知っている（利用経験有）" "39名 (16.5%)" $false
Add-LegendRow $s4 42 472 406 $colLightGray "未記入" "4名 (1.7%)" $false

# 右：Q3 円グラフ
Add-SoftCard $s4 490 88 450 432 "Q3. 万が一の際、当院を利用したいですか？（N=236）" $colVisitorMain
Add-DonutChart $s4 715 240 196 @(
    @{ pct = 25.4; color = $colTeal },
    @{ pct = 64.8; color = $colBlue },
    @{ pct = 7.2;  color = $colSlate },
    @{ pct = 2.5;  color = $colLightGray }
) "利用・検討したい" "90.3% (213名)"
Add-LegendRow $s4 512 356 406 $colTeal      "ぜひ利用したい" "60名 (25.4%)" $true $colPastelT
Add-LegendRow $s4 512 396 406 $colBlue      "候補の一つとして検討したい" "153名 (64.8%)" $true $colPastelB
Add-LegendRow $s4 512 436 406 $colSlate     "他の病院を選ぶ・わからない" "17名 (7.2%)" $false
Add-LegendRow $s4 512 472 406 $colLightGray "未記入" "6名 (2.5%)" $false


# ==============================================================================
# SLIDE 5: 【来場者＝青】来場者④ Q4. 健康で気になること（左：意味のある棒グラフ ／ 右：その他記述）
# ==============================================================================
$s5 = $pres.Slides.Add(5, 12)
Add-SimpleHeader $s5 "来場者アンケート④" "Q4. 現在健康について気になることはありますか？（グラフ＋「その他」自由記載）" 5 $colVisitorMain $colVisitorSoft

# 左：Q4 棒グラフ（3割超の上位4項目＝青、最少＝赤、中間＝グレー）
Add-SoftCard $s5 20 88 490 432 "Q4. 健康について気になること（複数回答・N=236）" $colVisitorMain
$q4Rows = @(
    @("足腰の衰え・転倒", "93名 (39.4%)", 39.4, $colBlue),
    @("食事・栄養バランス", "79名 (33.5%)", 33.5, $colBlue),
    @("物忘れ・認知機能", "78名 (33.1%)", 33.1, $colBlue),
    @("将来の介護や終活", "71名 (30.1%)", 30.1, $colBlue),
    @("とくにない", "34名 (14.4%)", 14.4, $colSlate),
    @("薬の飲み方", "11名 (4.7%)", 4.7, $colCoral),
    @("その他", "11名 (4.7%)", 4.7, $colCoral)
)
for ($i = 0; $i -lt $q4Rows.Count; $i++) {
    $r = $q4Rows[$i]
    Add-MeaningBar $s5 36 (148 + $i * 44) 155 170 $r[0] $r[1] ($r[2] * 2.0) $r[3] 14
}
Add-MiniBarRule $s5 52 474 426 "青：上位 (30%超)" "赤：最少 (4.7%)"

# 右：Q4「その他」の自由記載一覧
Add-SoftCard $s5 526 88 414 432 "Q4「その他」の自由記載内容（計10件）" $colVisitorMain
Add-Rect $s5 542 148 382 354 $colPastelB -1 $true 0.05 | Out-Null
$q4OtherText = @"
・血圧 ／ 血圧が高い（2件）

・糖尿病（1件）

・ひざの痛み（1件）

・足がつる、むくみ（1件）

・胃痛（1件）

・痔（1件）

・早く目が覚める、夜間トイレに行く（1件）

・薬に頼らず努力で（1件）

・病気について（1件）
"@
Add-Text $s5 554 156 358 340 $q4OtherText 14.2 $colText $false 1 | Out-Null


# ==============================================================================
# SLIDE 6: 【来場者＝青】来場者⑤ Q5. 来年に期待する内容は何ですか？（自由記載の要約・羅列）
# ==============================================================================
$s6 = $pres.Slides.Add(6, 12)
Add-SimpleHeader $s6 "来場者アンケート⑤" "Q5. 来年に期待する内容は何ですか？（自由記載39件の要約）" 6 $colVisitorMain $colVisitorSoft

Add-SoftCard $s6 20 88 450 432 "測定・運動・健康チェックに関する記載" $colVisitorMain
Add-Rect $s6 36 148 418 354 $colPastelB -1 $true 0.05 | Out-Null
$q5Left = @"
・今年と同じ内容・同様の測定（筋肉量測定・血管年齢測定など）を来年も継続してほしい（11件）

・自宅の椅子に座っている時のストレッチ、日々できる簡単な運動体験、筋肉をつける姿勢を知りたい（4件）

・認知機能の衰えチェック、認知症・フレイル予防、体力測定をしてほしい（7件）

・転倒予防、バランスが悪いか・転倒危険度のチェックをしてほしい（3件）

・骨粗しょう症・骨密度チェックと対策をしてほしい（2件）

・今回測定した項目以外にも、さらに色々な測定・気軽な健康チェックをしてほしい（2件）
"@
Add-Text $s6 48 156 394 340 $q5Left 14 $colText $false 1 | Out-Null

Add-SoftCard $s6 490 88 450 432 "相談・感想・運営面に関する記載" $colVisitorMain
Add-Rect $s6 506 148 418 354 $colPastelG -1 $true 0.05 | Out-Null
$q5Right = @"
・介護生活の相談、食事、家族が病院に入院した際の飲む薬について知りたい（1件）

・薬について（ジェネリックが効かない）（1件）

・こんな症状には何科に行ったらいいか、五十肩、リハビリへの取り組み方を知りたい（3件）

・マンションで5人ほど入院した方がいて、とてもリハビリが良かったとのこと（1件）

・リハビリセンサーや測定が参考になりとても良かった、来年も元気でいられるアドバイスをお願いしたい（4件）

・記入用のバインダーがあると助かる（1件）

・InBodyで靴を脱ぐ際、靴入れと靴ベラがあると良かった（1件）

・測定の案内（スタンプラリー）の順番をわかりやすくしてほしい（1件）
"@
Add-Text $s6 518 156 394 340 $q5Right 13.8 $colText $false 1 | Out-Null


# ==============================================================================
# SLIDE 7: 【スタッフ＝緑】スタッフ① 設問① 参加して良かったか（左：グラフ ／ 右：②理由）
# ==============================================================================
$s7 = $pres.Slides.Add(7, 12)
Add-SimpleHeader $s7 "スタッフアンケート①" "① 今回のどーんと祭に参加して良かったですか？ ／ ② ①の理由" 7 $colStaffMain $colStaffSoft

Add-SoftCard $s7 20 88 420 432 "① 参加して良かったですか？（N=13）" $colStaffMain
Add-DonutChart $s7 230 250 202 @(
    @{ pct = 100.0; color = $colTeal }
) "はい" "100% (13名)"
Add-LegendRow $s7 46 388 368 $colTeal      "はい" "13名 (100.0%)" $true $colPastelT
Add-LegendRow $s7 46 432 368 $colLightGray "いいえ" "0名 (0.0%)" $false

Add-SoftCard $s7 456 88 484 432 "② ①の理由を教えてください（自由記載の要約）" $colStaffMain
Add-Rect $s7 472 148 452 354 $colPastelT -1 $true 0.05 | Out-Null
$st2Text = @"
・地域住民の方と直接お話しし、当院や介護予防教室をPRできたため。

・一般の方々の生活習慣や悩みをリアルに伺え、ご家族が悩まれていることに寄り添える病院だと伝えられたため。

・他部門メンバーと同じ目標に向かって一致団結し、普段できないようなコミュニケーションを取れたため。

・多職種連携ができた、多職種と楽しく地域社会に関われた、当院の結束力が感じられたため。

・たくさんの方が来場されやりがいを感じた、予防医療の大切さや病院アピールの改善点が分かったため。

・普段は病院内で働いているため、外で地域の人と関わる活動が新鮮で良い経験になったため。
"@
Add-Text $s7 484 156 428 340 $st2Text 14 $colText $false 1 | Out-Null


# ==============================================================================
# SLIDE 8: 【スタッフ＝緑】スタッフ② 設問③ 地域住民に貢献できたか（左：グラフ ／ 右：④理由）
# ==============================================================================
$s8 = $pres.Slides.Add(8, 12)
Add-SimpleHeader $s8 "スタッフアンケート②" "③ 自部署の出展内容や役割は地域住民に貢献できましたか？ ／ ④ ③の理由" 8 $colStaffMain $colStaffSoft

Add-SoftCard $s8 20 88 420 432 "③ 地域住民に貢献できましたか？（N=13）" $colStaffMain
Add-DonutChart $s8 230 250 202 @(
    @{ pct = 92.3; color = $colTeal },
    @{ pct = 7.7;  color = $colSlate }
) "はい" "92.3% (12名)"
Add-LegendRow $s8 46 388 368 $colTeal  "はい" "12名 (92.3%)" $true $colPastelT
Add-LegendRow $s8 46 432 368 $colSlate "いいえ" "1名 (7.7%)" $false

Add-SoftCard $s8 456 88 484 432 "④ ③の理由を教えてください（自由記載の要約）" $colStaffMain
Add-Rect $s8 472 148 452 354 $colPastelT -1 $true 0.05 | Out-Null
$st4Text = @"
【「はい」と回答した理由（12名）】
・筋肉量測定（InBody）に興味を持たれた方が多く、測定結果をもとに運動指導やアドバイスができたため。
・自分の身体の状態を知る機会となり、専門職から様々なアドバイスを行えたため。
・限られた時間の中でできる限りの栄養アドバイスを行えたため。
・お薬カレンダーの無料配布を喜んでいた方がいたため。
・高齢者以外の層にも興味を持っていただける内容であり、近隣の医師会ブースと比較しても多くの住民が来場してくれたため。
・受付として待ち時間がかかることを事前に案内できたため。

【「いいえ」と回答した理由（1名）】
・担当テーマが「人生会議」であり、来場者が求めていた内容だったか判断が難しかったため。
"@
Add-Text $s8 484 156 428 340 $st4Text 13.6 $colText $false 1 | Out-Null


# ==============================================================================
# SLIDE 9: 【スタッフ＝緑】スタッフ③ 設問⑤ 3区画の広さはいかがだったか（左：グラフ ／ 右：⑥問題点）
# ==============================================================================
$s9 = $pres.Slides.Add(9, 12)
Add-SimpleHeader $s9 "スタッフアンケート③" "⑤ 今回3区画を借りて出展しましたが広さはいかがでしたか？ ／ ⑥ 問題点" 9 $colStaffMain $colStaffSoft

Add-SoftCard $s9 20 88 420 432 "⑤ 3区画の広さはいかがでしたか？（N=13）" $colStaffMain
Add-DonutChart $s9 230 250 202 @(
    @{ pct = 61.5; color = $colTeal },
    @{ pct = 38.5; color = $colCoral }
) "問題あり" "38.5% (5名)"
Add-LegendRow $s9 46 388 368 $colTeal  "問題なし" "8名 (61.5%)" $false
Add-LegendRow $s9 46 432 368 $colCoral "問題あり" "5名 (38.5%)" $true $colPastelC

Add-SoftCard $s9 456 88 484 432 "⑥「問題あり」と答えた方の具体的な問題点（要約）" $colStaffMain
Add-Rect $s9 472 148 452 354 $colPastelC -1 $true 0.05 | Out-Null
$st6Text = @"
・常に混雑していたため狭かった。

・時間帯にもよるが、お待ちの方々が座れず、外出をお勧めして対応した。

・満席の状態が続くことがあり、移動時に狭そうな様子が見受けられた。

・InBodyの測定スペースが狭かった。

・同規模で行うならそのままでもよいが、InBodyを増やすなど規模を広げるのであれば、もう半テント（4区画）必要と感じた。
"@
Add-Text $s9 484 160 428 330 $st6Text 14.8 $colText $false 1 | Out-Null


# ==============================================================================
# SLIDE 10: 【スタッフ＝緑】スタッフ④ 設問⑦ 人員配置は適切だったか（左：グラフ ／ 右：⑧問題点）
# ==============================================================================
$s10 = $pres.Slides.Add(10, 12)
Add-SimpleHeader $s10 "スタッフアンケート④" "⑦ 今回の人員配置（部署参加人数）は適切でしたか？ ／ ⑧ 問題点" 10 $colStaffMain $colStaffSoft

Add-SoftCard $s10 20 88 420 432 "⑦ 人員配置は適切でしたか？（N=13）" $colStaffMain
Add-DonutChart $s10 230 250 202 @(
    @{ pct = 76.9; color = $colCoral },
    @{ pct = 23.1; color = $colTeal }
) "問題あり" "76.9% (10名)"
Add-LegendRow $s10 46 388 368 $colCoral "問題あり" "10名 (76.9%)" $true $colPastelC
Add-LegendRow $s10 46 432 368 $colTeal  "問題なし" "3名 (23.1%)" $false

Add-SoftCard $s10 456 88 484 432 "⑧「問題あり」と答えた方の具体的な問題点（要約）" $colStaffMain
Add-Rect $s10 472 148 452 354 $colPastelC -1 $true 0.05 | Out-Null
$st8Text = @"
・来場者が多く、休憩時間がほとんど取れなかった（長くても10〜15分×2回程度）。

・1人が休憩に入ると残りの人数で回すのが大変だったため、リハ・測定スタッフや看護スタッフをあと1〜3名多く配置してほしい。

・並んでいる方や待っている方を公平に案内する「誘導係（2名程度）」が必要だった。

・全体を把握・調整する「全体統括」の役割が必要だった。

・事務スタッフがもう1名いると良いと感じた。

・2日間連続の出勤は負担が大きかった。
"@
Add-Text $s10 484 156 428 340 $st8Text 14.2 $colText $false 1 | Out-Null


# ==============================================================================
# SLIDE 11: 【スタッフ＝緑】スタッフ⑤ 設問⑨ 来年も自部署が参加する必要性はあるか（左：グラフ ／ 右：⑩理由）
# ==============================================================================
$s11 = $pres.Slides.Add(11, 12)
Add-SimpleHeader $s11 "スタッフアンケート⑤" "⑨ 来年も自部署が参加する必要性はあると考えますか？ ／ ⑩ ⑨の理由" 11 $colStaffMain $colStaffSoft

Add-SoftCard $s11 20 88 420 432 "⑨ 来年も参加する必要性はありますか？（N=13）" $colStaffMain
Add-DonutChart $s11 230 250 202 @(
    @{ pct = 84.6; color = $colTeal },
    @{ pct = 15.4; color = $colSlate }
) "はい" "84.6% (11名)"
Add-LegendRow $s11 46 388 368 $colTeal  "はい" "11名 (84.6%)" $true $colPastelT
Add-LegendRow $s11 46 432 368 $colSlate "いいえ" "2名 (15.4%)" $false

Add-SoftCard $s11 456 88 484 432 "⑩ ⑨の理由を教えてください（自由記載の要約）" $colStaffMain
Add-Rect $s11 472 148 452 354 $colPastelT -1 $true 0.05 | Out-Null
$st10Text = @"
【「はい」と回答した理由（11名）】
・まだまだ当院の認知度が低いことが分かり、病院を宣伝する良い機会になるため。
・InBody測定が大変好評であり、健康に対する地域の皆様の関心が高いため。
・栄養相談の需要が想像以上に高いと分かったため。
・血管年齢測定に加えてその他の健康相談をされる方も複数おり、看護師として生活アドバイスを行う必要があるため。
・多職種で関わることで一致団結感が生まれるため。
・病院として出展するなら全部署が参加した方が良いと感じるため（少人数部署は2日間のうち1日参加や相談がある時のみの対応でもよい）。

【「いいえ」と回答した理由（2名）】
・2日連続の参加はきついため。 ／ ・自部署に限らず人員としては必要だと思うため。
"@
Add-Text $s11 484 154 428 344 $st10Text 13.4 $colText $false 1 | Out-Null


# ==============================================================================
# SLIDE 12: 【スタッフ＝緑】スタッフ⑥ 設問⑪ 来年に向けて改善・強化・取り入れた方がよいこと
# ==============================================================================
$s12 = $pres.Slides.Add(12, 12)
Add-SimpleHeader $s12 "スタッフアンケート⑥" "⑪ 来年に向けて改善・強化・取り入れた方がよいこと（自由記載の要約）" 12 $colStaffMain $colStaffSoft

Add-SoftCard $s12 20 88 450 432 "機器・動線・シフト体制に関する記載" $colStaffMain
Add-Rect $s12 36 148 418 354 $colPastelT -1 $true 0.05 | Out-Null
$st11Left = @"
・InBodyのスペース拡張と「2台体制」への増設（希望者が多く30分以上待ちとなり最後は並ぶのを中止したため／1台では150人が限度／または先着何名までと決める）。

・待ち時間防止のため整理番号を配って時間で来てもらう方式、またはブースを自由に回るのではなく「栄養→薬剤→MSW→看護」と順路を決めて誘導する方式の導入。

・InBodyの待ち時間に健康情報を提供するなど有効活用する工夫、栄養相談はInBodyの結果ありきで話す連携、栄養士セミナー形式の検討。

・参加職員の休憩時間を1時間確保してほしい（各日5〜6人での運用、2日目の現地集合組の設定、スタッフ用おにぎり・サンドイッチ等の準備）。

・医師の参加の検討。
"@
Add-Text $s12 48 156 394 340 $st11Left 13.6 $colText $false 1 | Out-Null

Add-SoftCard $s12 490 88 450 432 "広報・持参品・備品に関する記載" $colStaffMain
Add-Rect $s12 506 148 418 354 $colPastelG -1 $true 0.05 | Out-Null
$st11Right = @"
・病院広報誌、どんな病院か（リハ病院の役割）が分かる紹介パネル、Instagram等の映像、求人広告・ポスターの持参（※リハ病院の役割を知らない方や医療センターと誤認している方がいたため）。

・「あなたの健康カルテ」に①受付番号、②実年齢・西暦の生年月日（血管年齢の入力に必要なため）の記入欄を追加。

・ノベルティや配布物の拡充（アンケート回答者全員へのプレゼント、病院名入りボールペン、防災グッズを200個程度へ増量、お薬カレンダーを120個程度へ増量、エコバッグの増量、食べ物の無料配布など）。

・区画スペース拡大時の折りたたみ椅子増設、入口の右向き矢印表示、InBody用のペースメーカーNG表示の準備。
"@
Add-Text $s12 518 156 394 340 $st11Right 13.6 $colText $false 1 | Out-Null

for ($i = 1; $i -le $pres.Slides.Count; $i++) {
    $imgPath = Join-Path $artifactDir ("handout_" + $i + ".png")
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
    $altPath = "c:\Users\coffe\OneDrive\デスクトップ\Antigravity\アンケート解析\アンケート結果報告_修正版.pptx"
    Copy-Item -Path $webPath -Destination $altPath -Force
    Write-Host "Saved to $altPath and $webPath"
}
Write-Host "Successfully generated 12-slide Staff Handout PowerPoint!"
