$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$outPath = "c:\Users\coffe\OneDrive\デスクトップ\Antigravity\アンケート解析\八千代リハビリテーション病院_どーんと祭出展結果報告_四役プレゼン.pptx"

$pp = New-Object -ComObject PowerPoint.Application
# PowerPoint 16: 16:9 Widescreen (960 x 540 pt)
$pres = $pp.Presentations.Add($false)
$pres.PageSetup.SlideWidth = 960
$pres.PageSetup.SlideHeight = 540

function RGB($r, $g, $b) {
    return [int]($r + ($g * 256) + ($b * 65536))
}

$colNavy      = RGB 15 35 66
$colBlue      = RGB 30 64 175
$colLightBlue = RGB 239 246 255
$colBorderB   = RGB 191 219 254
$colTeal      = RGB 13 148 136
$colLightTeal = RGB 240 253 250
$colAmber     = RGB 217 119 6
$colLightAmb  = RGB 255 251 235
$colRed       = RGB 185 28 28
$colLightRed  = RGB 254 242 242
$colWhite     = RGB 255 255 255
$colBgGray    = RGB 248 250 252
$colTextDark  = RGB 15 23 42
$colTextMuted = RGB 71 85 105
$colBarBg     = RGB 226 232 240

function Add-Rect($slide, $left, $top, $width, $height, $fillRGB, $lineRGB = -1, $round = $false) {
    $type = if ($round) { 5 } else { 1 } # 5=RoundedRect, 1=Rect
    $sh = $slide.Shapes.AddShape([int]$type, [single]$left, [single]$top, [single]$width, [single]$height)
    $sh.Fill.Visible = -1
    $sh.Fill.ForeColor.RGB = [int]$fillRGB
    if ($lineRGB -ge 0) {
        $sh.Line.Visible = -1
        $sh.Line.ForeColor.RGB = [int]$lineRGB
        $sh.Line.Weight = [single]1.2
    } else {
        $sh.Line.Visible = 0
    }
    if ($round -and $sh.Adjustments.Count -ge 1) {
        $sh.Adjustments.Item(1) = [single]0.06
    }
    return $sh
}

function Add-Text($slide, $left, $top, $width, $height, $text, $fontSize = 14, $colorRGB = $colTextDark, $bold = $false, $align = 1) {
    $tb = $slide.Shapes.AddTextbox(1, [single]$left, [single]$top, [single]$width, [single]$height)
    $tf = $tb.TextFrame
    $tf.MarginLeft = 4
    $tf.MarginRight = 4
    $tf.MarginTop = 2
    $tf.MarginBottom = 2
    $tf.WordWrap = -1
    $tr = $tf.TextRange
    $tr.Text = $text
    $tr.Font.Name = "Meiryo UI"
    $tr.Font.NameFarEast = "Meiryo UI"
    $tr.Font.Size = [single]$fontSize
    $tr.Font.Color.RGB = [int]$colorRGB
    $tr.Font.Bold = if ($bold) { -1 } else { 0 }
    $tr.ParagraphFormat.Alignment = $align # 1=Left, 2=Center, 3=Right
    return $tb
}

function Add-Header($slide, $stepBadge, $titleText, $keyMessage, $pageNo) {
    Add-Rect $slide 0 0 960 540 $colBgGray | Out-Null
    Add-Rect $slide 0 0 960 56 $colNavy | Out-Null
    $b = Add-Rect $slide 24 12 110 32 $colTeal -1 $true
    Add-Text $slide 24 15 110 28 $stepBadge 12 $colWhite $true 2 | Out-Null
    Add-Text $slide 146 11 680 34 $titleText 20 $colWhite $true 1 | Out-Null
    Add-Text $slide 830 16 106 24 ("P. " + $pageNo + " / 5") 12 (RGB 203 213 225) $true 3 | Out-Null

    # Key message banner
    Add-Rect $slide 24 68 912 46 $colLightBlue $colBorderB $true | Out-Null
    Add-Rect $slide 24 68 8 46 $colBlue | Out-Null
    Add-Text $slide 42 74 884 36 $keyMessage 13.5 $colNavy $true 1 | Out-Null
}

function Add-HBar($slide, $left, $top, $labelW, $maxBarW, $label, $valText, $pctVal, $barColor) {
    Add-Text $slide $left $top $labelW 24 $label 12.5 $colTextDark $true 1 | Out-Null
    $barX = $left + $labelW + 6
    Add-Rect $slide $barX ($top + 3) $maxBarW 18 $colBarBg | Out-Null
    $fillW = [math]::Max(6, [math]::Round($maxBarW * ($pctVal / 100.0)))
    Add-Rect $slide $barX ($top + 3) $fillW 18 $barColor | Out-Null
    Add-Text $slide ($barX + $maxBarW + 8) $top 115 24 $valText 12.5 $colNavy $true 1 | Out-Null
}

# ==============================================================================
# SLIDE 1: 表紙 ＆ エグゼクティブサマリー（400字報告書＋3大KPI）
# ==============================================================================
$s1 = $pres.Slides.Add(1, 12) # 12 = ppLayoutBlank
Add-Rect $s1 0 0 960 540 $colBgGray | Out-Null
Add-Rect $s1 0 0 960 118 $colNavy | Out-Null
Add-Rect $s1 0 114 960 5 $colTeal | Out-Null

Add-Text $s1 36 18 880 28 "八千代リハビリテーション病院｜四役報告資料（2026年10月17日・18日 実施）" 14 (RGB 147 197 253) $true 1 | Out-Null
Add-Text $s1 36 48 880 54 "八千代市産業祭（どーんと祭）出展結果 報告書" 28 $colWhite $true 1 | Out-Null

# 3 KPI Cards
# Card 1
Add-Rect $s1 32 134 288 132 $colWhite $colBorderB $true | Out-Null
Add-Rect $s1 32 134 288 8 $colBlue | Out-Null
Add-Text $s1 46 148 260 24 "① アンケート有効回答数（目標50名）" 12.5 $colTextMuted $true 1 | Out-Null
Add-Text $s1 46 172 260 48 "236 名（4.7倍）" 28 $colBlue $true 1 | Out-Null
Add-Text $s1 46 224 260 36 "達成率 472.0%（1日目:106名 / 2日目:130名）`n※白紙除外後の確定有効回答数" 11 $colTextDark $false 1 | Out-Null

# Card 2
Add-Rect $s1 336 134 288 132 $colWhite (RGB 153 246 228) $true | Out-Null
Add-Rect $s1 336 134 288 8 $colTeal | Out-Null
Add-Text $s1 350 148 260 24 "② 当院の今後の利用・検討意向" 12.5 $colTextMuted $true 1 | Out-Null
Add-Text $s1 350 172 260 48 "90.3 %（213名）" 28 $colTeal $true 1 | Out-Null
Add-Text $s1 350 224 260 36 "「ぜひ利用したい」25.4%（60名）`n＋「候補として検討」64.8%（153名）" 11 $colTextDark $false 1 | Out-Null

# Card 3
Add-Rect $s1 640 134 288 132 $colWhite (RGB 253 230 138) $true | Out-Null
Add-Rect $s1 640 134 288 8 $colAmber | Out-Null
Add-Text $s1 654 148 260 24 "③ 新規認知獲得 ＆ 広域集客" 12.5 $colTextMuted $true 1 | Out-Null
Add-Text $s1 654 174 268 44 "28.4 %（新規67名）" 24 $colAmber $true 1 | Out-Null
Add-Text $s1 654 224 260 36 "「今日初めて当院を知った」が約3割`n八千代市内 72.5% ／ 市外 26.3%（62名）" 11 $colTextDark $false 1 | Out-Null

# Executive Summary Box (400字報告書全文・構造化)
Add-Rect $s1 32 282 896 232 $colWhite $colBorderB $true | Out-Null
Add-Rect $s1 32 282 896 36 $colNavy | Out-Null
Add-Text $s1 48 287 860 26 "📋 出展結果 総括サマリー（簡易報告書）" 14 $colWhite $true 1 | Out-Null

$summaryBody = @"
■ 目的と全体成果：10月17日・18日の2日間、病院認知度向上と地域貢献を目的に産業祭へ初出展した。アンケート有効回答数は 236名（1日目106名・2日目130名）に達し、起案目標50名の約4.7倍（達成率472%）と想定を大きく上回る成果を収めた。

■ 来場者属性と反応：来場者は八千代市内72.5%、市外26.3%と広域に及び、1日目は60代以上のシニア層中心（52.8%）、2日目は40代以下の若年・子育て世代が倍増（19.8%→32.3%）した。体験では「血管年齢測定（72.9%）」「筋肉量測定（66.5%）」が特に好評で、「当院を今日初めて知った」新規層が28.4%（67名）、万が一の際に当院を「利用・検討したい」層は 90.3%（213名）に上った。また地域の健康課題は「転倒予防（39.4%）」「栄養（33.5%）」「認知機能（33.1%）」「介護・終活（30.1%）」が各3割超と高く、今後の健康教室の有力テーマが明確となった。

■ 来期に向けた課題：一方、3区画では常時混雑し測定制限も生じたため、参加スタッフ全員（100%）が「参加して良かった」と回答しつつも75%が人員・スペース不足を挙げた。来期は「4区画への拡大」と「誘導・測定スタッフの増員（休憩確保）」を課題としたい。
"@
Add-Text $s1 48 326 864 180 $summaryBody 12.2 $colTextDark $false 1 | Out-Null

# ==============================================================================
# SLIDE 2: 来場者属性分析（1日目 vs 2日目の年齢変化・居住地・性別）
# ==============================================================================
$s2 = $pres.Slides.Add(2, 12)
Add-Header $s2 "基本属性分析" "来場者の年齢層（1日目 vs 2日目）・お住まい・性別" "1日目は60代以上のシニア層中心（52.8%）、2日目は40代以下の若年・子育て世代が倍増（19.8%→32.3%）！" 2

# Left Card: 1日目 vs 2日目 年齢構成比較
Add-Rect $s2 24 126 530 390 $colWhite $colBorderB $true | Out-Null
Add-Text $s2 40 136 500 28 "📊 年代別の来場変化（1日目：106名 vs 2日目：130名）" 15 $colNavy $true 1 | Out-Null
Add-Text $s2 40 164 500 22 "※2日目は40歳代が 8.5% → 18.5%（9名→24名）と2.7倍に急増し、家族連れが多数来場" 11 $colTextMuted $false 1 | Out-Null

# Age Table inside Left Card
$tblAge = $s2.Shapes.AddTable(7, 4, 40, 192, 498, 230).Table
$headersAge = @("年代区分", "全体 (236名)", "1日目 (106名)", "2日目 (130名)")
for ($c=1; $c -le 4; $c++) {
    $cell = $tblAge.Cell(1, $c)
    $cell.Shape.Fill.ForeColor.RGB = $colNavy
    $tr = $cell.Shape.TextFrame.TextRange
    $tr.Text = $headersAge[$c-1]
    $tr.Font.NameFarEast = "Meiryo UI"
    $tr.Font.Size = [single]11.5
    $tr.Font.Bold = -1
    $tr.Font.Color.RGB = $colWhite
}
$rowsAge = @(
    @("10〜30歳代（若年層）", "30名 (12.7%)", "12名 (11.3%)", "18名 (13.8%) ▲"),
    @("40歳代（子育て世代）", "33名 (14.0%)", "9名 (8.5%)", "24名 (18.5%) ▲2.7倍"),
    @("50歳代（最多層）", "62名 (26.3%)", "28名 (26.4%)", "34名 (26.2%)"),
    @("60〜70歳代（シニア層）", "81名 (34.3%)", "40名 (37.7%)", "41名 (31.5%)"),
    @("80〜90歳代（高齢層）", "27名 (11.4%)", "16名 (15.1%)", "11名 (8.5%) ▼"),
    @("未記入", "3名 (1.3%)", "1名 (0.9%)", "2名 (1.5%)")
)
for ($r=0; $r -lt $rowsAge.Count; $r++) {
    for ($c=0; $c -lt 4; $c++) {
        $cell = $tblAge.Cell($r+2, $c+1)
        $cell.Shape.Fill.ForeColor.RGB = if ($r -eq 1) { $colLightAmb } elseif ($r % 2 -eq 0) { $colWhite } else { $colBgGray }
        $tr = $cell.Shape.TextFrame.TextRange
        $tr.Text = $rowsAge[$r][$c]
        $tr.Font.NameFarEast = "Meiryo UI"
        $tr.Font.Size = 11
        $tr.Font.Bold = if ($r -eq 1 -or $r -eq 2) { -1 } else { 0 }
        $tr.Font.Color.RGB = if ($r -eq 1 -and $c -eq 3) { $colRed } else { $colTextDark }
    }
}

Add-Rect $s2 40 442 498 58 $colLightTeal (RGB 153 246 228) $true | Out-Null
Add-Text $s2 52 448 476 48 "💡 アナリスト考察：1日目は健康意識の高い60代以上が過半数（52.8%）。2日目は40代以下が32.3%（42名）へと倍増し、親世代の介護・入院を支えるキーパーソン層へ直接アプローチできた。" 11.2 $colTextDark $true 1 | Out-Null

# Right Top Card: お住まい（八千代市内 vs 市外）
Add-Rect $s2 570 126 366 202 $colWhite $colBorderB $true | Out-Null
Add-Text $s2 586 136 334 26 "🏠 お住まい内訳（広域集客効果）" 15 $colNavy $true 1 | Out-Null
Add-HBar $s2 586 172 96 130 "八千代市内" "171名 (72.5%)" 72.5 $colBlue
Add-HBar $s2 586 204 96 130 "八千代市外" "62名 (26.3%)" 26.3 $colTeal
Add-HBar $s2 586 236 96 130 "未記入" "3名 (1.3%)" 1.3 (RGB 148 163 184)
Add-Text $s2 586 270 338 48 "・1日目：市内 70.8%(75名) ／ 市外 27.4%(29名)`n・2日目：市内 73.8%(96名) ／ 市外 25.4%(33名)`n👉 全体の約4人に1人（62名）が市外からの広域来場！" 11 $colTextDark $false 1 | Out-Null

# Right Bottom Card: 性別内訳
Add-Rect $s2 570 342 366 174 $colWhite $colBorderB $true | Out-Null
Add-Text $s2 586 352 334 26 "🚻 性別内訳（全236名）" 15 $colNavy $true 1 | Out-Null
Add-HBar $s2 586 388 96 130 "女性" "153名 (64.8%)" 64.8 (RGB 219 39 119)
Add-HBar $s2 586 420 96 130 "男性" "81名 (34.3%)" 34.3 $colBlue
Add-Text $s2 586 456 338 48 "・女性が約65%（1日目62.3% → 2日目66.9%）を占め、家族の健康管理や栄養・測定への関心が非常に高い。" 11 $colTextDark $false 1 | Out-Null

# ==============================================================================
# SLIDE 3: Q1 ブース体験評価 ＆ Q2 認知度・Q3 利用意向
# ==============================================================================
$s3 = $pres.Slides.Add(3, 12)
Add-Header $s3 "Q1・Q2・Q3 分析" "ブース体験評価（Q1）と 当院認知度（Q2）・利用意向（Q3）" "「血管年齢測定(72.9%)」「筋肉量測定(66.5%)」が集客を牽引し、新規認知28.4%・利用意向90.3%を獲得！" 3

# Left Card: Q1
Add-Rect $s3 24 126 450 390 $colWhite $colBorderB $true | Out-Null
Add-Text $s3 40 136 420 28 "🩺 Q1. 本日参考になった内容（複数回答）" 15 $colNavy $true 1 | Out-Null
Add-HBar $s3 40 180 148 160 "① 血管年齢測定" "172名 (72.9%)" 72.9 $colBlue
Add-HBar $s3 40 218 148 160 "② 筋肉量測定(InBody)" "157名 (66.5%)" 66.5 $colTeal
Add-HBar $s3 40 256 148 160 "③ 栄養レッスン" "57名 (24.2%)" 24.2 $colAmber
Add-HBar $s3 40 294 148 160 "④ 人生会議(終活)案内" "41名 (17.4%)" 17.4 (RGB 109 40 217)
Add-HBar $s3 40 332 148 160 "⑤ お薬相談" "32名 (13.6%)" 13.6 (RGB 71 85 105)

Add-Rect $s3 40 376 418 122 $colLightAmb (RGB 253 230 138) $true | Out-Null
Add-Text $s3 52 384 396 108 "📌 重要ファクト（機会損失の発生）：`n・血管年齢と筋肉量（InBody）が圧倒的な集客フックとして機能。`n・ただし2日目はInBody希望者が殺到し、終盤に並ぶのを中止（入場制限）せざるを得なかったため、筋肉量測定の回答率が1日目(76.4%)→2日目(58.5%)へ低下。来期の「InBody 2台体制・4区画化」でさらなる伸びしろあり。" 11 $colTextDark $false 1 | Out-Null

# Right Top Card: Q2 認知度
Add-Rect $s3 490 126 446 188 $colWhite $colBorderB $true | Out-Null
Add-Text $s3 506 136 414 26 "🏥 Q2. これまで当院をご存知でしたか？" 15 $colNavy $true 1 | Out-Null
Add-HBar $s3 506 172 175 130 "名前や場所は知っている" "126名 (53.4%)" 53.4 $colBlue
Add-HBar $s3 506 204 175 130 "今日初めて知った【新規】" "67名 (28.4%)" 28.4 $colAmber
Add-HBar $s3 506 236 175 130 "よく知っている(利用経験有)" "39名 (16.5%)" 16.5 $colTeal
Add-Text $s3 506 268 416 38 "※「今日初めて知った」67名（28.4%）への直接認知獲得に成功！`n※2日目は「利用経験あり」が10.4%→21.5%（28名）へ倍増。" 11 $colTextDark $true 1 | Out-Null

# Right Bottom Card: Q3 利用意向
Add-Rect $s3 490 328 446 188 $colWhite $colBorderB $true | Out-Null
Add-Text $s3 506 338 414 26 "🤝 Q3. 万が一の際、当院を利用したいですか？" 15 $colNavy $true 1 | Out-Null
Add-HBar $s3 506 374 175 130 "候補の一つとして検討" "153名 (64.8%)" 64.8 $colBlue
Add-HBar $s3 506 406 175 130 "ぜひ利用したい" "60名 (25.4%)" 25.4 $colTeal
Add-HBar $s3 506 438 175 130 "他の病院・わからない" "17名 (7.2%)" 7.2 (RGB 148 163 184)
Add-Rect $s3 506 468 414 36 $colLightTeal (RGB 153 246 228) $true | Out-Null
Add-Text $s3 516 474 396 26 "✅ 「ぜひ利用＋検討したい」合計 213名（90.3%）が好意的意向！" 12 $colTeal $true 1 | Out-Null

# ==============================================================================
# SLIDE 4: Q4 地域の健康課題ニーズ ＆ Q5 自由記述（今後の健康教室テーマ）
# ==============================================================================
$s4 = $pres.Slides.Add(4, 12)
Add-Header $s4 "Q4・Q5 ニーズ分析" "地域の健康課題（Q4）と 来年への期待（Q5自由記述）" "「転倒予防(39.4%)」「栄養(33.5%)」「認知機能(33.1%)」「介護・終活(30.1%)」の4大ニーズが明確化！" 4

# Left Card: Q4
Add-Rect $s4 24 126 450 390 $colWhite $colBorderB $true | Out-Null
Add-Text $s4 40 136 420 28 "🏃 Q4. 現在健康で気になること（複数回答）" 15 $colNavy $true 1 | Out-Null
Add-HBar $s4 40 176 152 155 "① 足腰の衰え・転倒" "93名 (39.4%)" 39.4 $colBlue
Add-HBar $s4 40 210 152 155 "② 食事・栄養バランス" "79名 (33.5%)" 33.5 $colTeal
Add-HBar $s4 40 244 152 155 "③ 物忘れ・認知機能" "78名 (33.1%)" 33.1 $colAmber
Add-HBar $s4 40 278 152 155 "④ 将来の介護や終活" "71名 (30.1%)" 30.1 (RGB 109 40 217)
Add-HBar $s4 40 312 152 155 "⑤ とくにない" "34名 (14.4%)" 14.4 (RGB 148 163 184)
Add-HBar $s4 40 346 152 155 "⑥ 薬の飲み方 / その他" "各11名 (4.7%)" 4.7 (RGB 100 116 139)

Add-Rect $s4 40 384 418 116 $colLightBlue $colBorderB $true | Out-Null
Add-Text $s4 52 392 396 100 "💡 健康教室のテーマ設定への示唆：`n・起案目的の一つである「地域住民が抱える健康課題の情報収集」において、上位4項目がいずれも30%超（70〜93名）の高水準で並んだ。`n・PT/OTによる「転倒予防・自宅ストレッチ」、管理栄養士の「栄養講座」、認知症予防、ACP（人生会議）を柱とした健康教室が極めて有望。" 11 $colTextDark $false 1 | Out-Null

# Right Card: Q5 Free Comments
Add-Rect $s4 490 126 446 390 $colWhite $colBorderB $true | Out-Null
Add-Text $s4 506 136 414 28 "💬 Q5. 来年に期待する内容・来場者の声（抜粋）" 15 $colNavy $true 1 | Out-Null

$q5Text = @"
【① 運動・転倒予防・骨密度チェックの要望（多数）】
・「自宅の椅子に座っている時のストレッチや筋肉をつける姿勢を知りたい」
・「日々できるカンタンな運動体験」「バランス・転倒危険度チェック」
・「骨粗しょう症・骨密度チェックと対策をしてほしい」

【② 認知機能・フレイルチェックの要望】
・「認知機能の衰えチェック、体力測定をしてほしい」（複数名）
・「認知症・フレイル予防について知りたい」

【③ 介護・薬・医療相談の要望】
・「介護生活の相談、食事、家族が入院した際の薬について」
・「こんな症状には何科に行ったらいいか知りたい」

【④ 当院への好意的な評価・運営改善ヒント】
・「マンションで5人ほど入院した方がいて、とてもリハビリが良かったとの事」
・「リハビリセンサーや測定が参考になりとても良かった。来年も期待」
・「InBodyで靴を脱ぐ際、靴入れと靴ベラ、記入用バインダーがあると助かる」
"@
Add-Text $s4 506 170 416 334 $q5Text 11.2 $colTextDark $false 1 | Out-Null

# ==============================================================================
# SLIDE 5: スタッフアンケート結果（12名）と 来期に向けた課題・改善提案
# ==============================================================================
$s5 = $pres.Slides.Add(5, 12)
Add-Header $s5 "来期課題・提案" "スタッフアンケート結果（12名）と 次年度に向けた改善提案" "スタッフ満足度100%の一方、3区画では常時混雑し機会損失が発生。来期は「4区画化＋増員」を提案" 5

# Left Card: スタッフアンケート結果
Add-Rect $s5 24 126 446 390 $colWhite $colBorderB $true | Out-Null
Add-Text $s5 40 136 414 28 "🧑‍⚕️ 参加スタッフアンケート結果（回答12名）" 15 $colNavy $true 1 | Out-Null

Add-HBar $s5 40 176 172 115 "参加して良かった（はい）" "12名 (100%)" 100.0 $colTeal
Add-HBar $s5 40 210 172 115 "地域貢献できた（はい）" "11名 (91.7%)" 91.7 $colBlue
Add-HBar $s5 40 244 172 115 "人員・配置に課題あり" "9名 (75.0%)" 75.0 $colRed

Add-Rect $s5 40 282 414 218 $colLightRed (RGB 254 202 202) $true | Out-Null
$staffVoice = @"
⚠️ 現場スタッフから挙がった具体的な課題（3区画・人員不足の影響）：
1. 【スペース狭小による機会損失】
   ・「3区画では常に混雑し狭かった。お待ち頂く方が座れず外出を勧めた」
   ・「InBody測定希望者が多すぎて、最後は並ぶのを中止（お断り）した」
2. 【休憩時間の未確保と誘導負担】
   ・「休憩時間が全く取れなかった。シフト制で休憩を回せる人員が必要」
   ・「呼び込み・誘導係、測定補助、事務スタッフの増員が必要」
3. 【機器・備品の改善点】
   ・「InBodyは2台体制が望ましい」「Wi-Fi環境の事前確認、靴ベラ・バインダー準備」
"@
Add-Text $s5 50 290 396 202 $staffVoice 11 $colTextDark $false 1 | Out-Null

# Right Card: 来期（次年度）への3大提案
Add-Rect $s5 486 126 450 390 $colWhite $colBorderB $true | Out-Null
Add-Text $s5 502 136 418 28 "🚀 来期（次年度）に向けた3つの改善提案" 15 $colNavy $true 1 | Out-Null

# Proposal 1
Add-Rect $s5 502 174 418 98 $colLightBlue $colBorderB $true | Out-Null
Add-Text $s5 516 180 392 86 "提案①：出展ブースを「3区画 → 4区画」へ拡大`n・今年度は四役相談のもと3区画（実績予算21.4万円）で実施したが、目標の4.7倍（236名）が殺到し待合椅子不足・入場制限が発生。`n・当初起案の「4区画」を確保し、測定動線・待合スペース・相談ブースを分離することで取りこぼし（機会損失）をゼロにする。" 11.2 $colTextDark $false 1 | Out-Null

# Proposal 2
Add-Rect $s5 502 284 418 98 $colLightTeal (RGB 153 246 228) $true | Out-Null
Add-Text $s5 516 290 392 86 "提案②：スタッフ増員による「休憩時間確保」と「誘導強化」`n・呼び込み/スタンプラリー誘導係・測定スタッフを増員し、交代制で確実に休憩時間を確保できるシフト体制を構築。`n・人気集中のInBodyは「2台体制」＋靴入れ・靴ベラ・バインダー完備で回転率を向上。" 11.2 $colTextDark $false 1 | Out-Null

# Proposal 3
Add-Rect $s5 502 394 418 106 $colLightAmb (RGB 253 230 138) $true | Out-Null
Add-Text $s5 516 400 392 94 "提案③：アンケートで判明した「4大健康課題」を健康教室へ展開`n・今回3割超のニーズが確認された「①転倒予防・骨密度」「②栄養」「③認知機能・フレイル」「④介護・終活」をテーマに院内・地域の健康教室を企画。`n・単発イベントで終わらせず、当院への継続的なファン化・受診/入院選択（利用意向90.3%）へつなげる。" 11.2 $colTextDark $false 1 | Out-Null

$pres.SaveAs($outPath)

$artifactDir = "C:\Users\coffe\.gemini\antigravity\brain\57f12e92-5bc8-492c-8c7d-8ab9b420cf7a"
for ($i = 1; $i -le $pres.Slides.Count; $i++) {
    $imgPath = Join-Path $artifactDir ("slide_" + $i + ".png")
    $pres.Slides.Item($i).Export($imgPath, "PNG", 1920, 1080)
}

$pres.Close()
$pp.Quit()
Write-Host "Successfully generated PowerPoint and exported 5 slide images: $outPath"
