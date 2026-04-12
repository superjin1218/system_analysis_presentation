[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = "Stop"

function OleColor([int]$r, [int]$g, [int]$b) {
  return ($r -bor ($g -shl 8) -bor ($b -shl 16))
}

function UniquePath([string]$Path) {
  if (-not (Test-Path -LiteralPath $Path)) {
    return $Path
  }

  $dir = Split-Path -LiteralPath $Path
  $base = [System.IO.Path]::GetFileNameWithoutExtension($Path)
  $ext = [System.IO.Path]::GetExtension($Path)
  $i = 2

  while ($true) {
    $candidate = Join-Path $dir "$base-$i$ext"
    if (-not (Test-Path -LiteralPath $candidate)) {
      return $candidate
    }
    $i++
  }
}

$script:msoTrue = -1
$script:msoFalse = 0
$script:ppLayoutBlank = 12
$script:msoShapeRoundedRectangle = 5
$script:msoShapeOval = 9
$script:msoTextOrientationHorizontal = 1
$script:ppAlignLeft = 1
$script:ppAlignCenter = 2
$script:ppAlignRight = 3
$script:FontName = "Pretendard"

$script:Colors = @{
  Background = OleColor 248 251 255
  White = OleColor 255 255 255
  Navy = OleColor 15 23 42
  Blue = OleColor 37 99 235
  Teal = OleColor 20 184 166
  Purple = OleColor 124 58 237
  Orange = OleColor 249 115 22
  Amber = OleColor 245 158 11
  Muted = OleColor 71 85 105
  Border = OleColor 226 232 240
  BorderStrong = OleColor 203 213 225
  SoftBlue = OleColor 219 234 254
  SoftMint = OleColor 204 251 241
  SoftPurple = OleColor 237 233 254
  SoftGray = OleColor 241 245 249
  DarkBlue = OleColor 30 41 59
}

function Set-ShapeText {
  param(
    $Shape,
    [string]$Text,
    [double]$FontSize,
    [int]$Color,
    [bool]$Bold = $false,
    [int]$Alignment = 1
  )

  $Shape.TextFrame2.TextRange.Text = $Text
  $Shape.TextFrame2.MarginLeft = 0
  $Shape.TextFrame2.MarginRight = 0
  $Shape.TextFrame2.MarginTop = 0
  $Shape.TextFrame2.MarginBottom = 0
  $Shape.TextFrame2.WordWrap = $script:msoTrue
  $Shape.TextFrame2.AutoSize = 0
  $Shape.TextFrame2.VerticalAnchor = 1
  $Shape.TextFrame2.TextRange.Font.Name = $script:FontName
  $Shape.TextFrame2.TextRange.Font.Size = $FontSize
  $Shape.TextFrame2.TextRange.Font.Bold = $(if ($Bold) { $script:msoTrue } else { $script:msoFalse })
  $Shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = $Color
  $Shape.TextFrame2.TextRange.ParagraphFormat.Alignment = $Alignment
  $Shape.TextFrame2.TextRange.ParagraphFormat.SpaceAfter = 0
  $Shape.TextFrame2.TextRange.ParagraphFormat.SpaceBefore = 0
}

function Add-TextBox {
  param(
    $Slide,
    [double]$Left,
    [double]$Top,
    [double]$Width,
    [double]$Height,
    [string]$Text,
    [double]$FontSize,
    [int]$Color,
    [bool]$Bold = $false,
    [int]$Alignment = 1
  )

  $safeWidth = [Math]::Max([double]$Width, 8)
  $safeHeight = [Math]::Max([double]$Height, 8)
  $shape = $Slide.Shapes.AddTextbox($script:msoTextOrientationHorizontal, $Left, $Top, $safeWidth, $safeHeight)
  $shape.Line.Visible = $script:msoFalse
  $shape.Fill.Visible = $script:msoFalse
  Set-ShapeText -Shape $shape -Text $Text -FontSize $FontSize -Color $Color -Bold $Bold -Alignment $Alignment
  return $shape
}

function Add-RoundedRect {
  param(
    $Slide,
    [double]$Left,
    [double]$Top,
    [double]$Width,
    [double]$Height,
    [int]$FillColor,
    [double]$FillTransparency = 0,
    [int]$LineColor = 0,
    [double]$LineTransparency = 0
  )

  $shape = $Slide.Shapes.AddShape($script:msoShapeRoundedRectangle, $Left, $Top, $Width, $Height)
  $shape.Fill.ForeColor.RGB = $FillColor
  $shape.Fill.Transparency = $FillTransparency
  $shape.Line.ForeColor.RGB = $LineColor
  $shape.Line.Transparency = $LineTransparency
  $shape.Line.Weight = 1
  return $shape
}

function Add-Oval {
  param(
    $Slide,
    [double]$Left,
    [double]$Top,
    [double]$Width,
    [double]$Height,
    [int]$FillColor,
    [double]$FillTransparency = 0,
    [int]$LineColor = 0,
    [double]$LineTransparency = 1,
    [double]$LineWeight = 1
  )

  $shape = $Slide.Shapes.AddShape($script:msoShapeOval, $Left, $Top, $Width, $Height)
  $shape.Fill.ForeColor.RGB = $FillColor
  $shape.Fill.Transparency = $FillTransparency
  $shape.Line.ForeColor.RGB = $LineColor
  $shape.Line.Transparency = $LineTransparency
  $shape.Line.Weight = $LineWeight
  return $shape
}

function Add-Line {
  param(
    $Slide,
    [double]$X1,
    [double]$Y1,
    [double]$X2,
    [double]$Y2,
    [int]$Color,
    [double]$Weight = 1.4,
    [double]$Transparency = 0.2
  )

  $line = $Slide.Shapes.AddLine($X1, $Y1, $X2, $Y2)
  $line.Line.ForeColor.RGB = $Color
  $line.Line.Weight = $Weight
  $line.Line.Transparency = $Transparency
  return $line
}

function Apply-Backdrop {
  param($Slide)

  $Slide.Background.Fill.ForeColor.RGB = $script:Colors.Background

  $blobA = Add-Oval -Slide $Slide -Left 710 -Top -40 -Width 230 -Height 230 -FillColor $script:Colors.SoftBlue -FillTransparency 0.45 -LineColor $script:Colors.SoftBlue -LineTransparency 1
  $blobB = Add-Oval -Slide $Slide -Left -70 -Top 370 -Width 220 -Height 220 -FillColor $script:Colors.SoftMint -FillTransparency 0.55 -LineColor $script:Colors.SoftMint -LineTransparency 1
  $blobC = Add-Oval -Slide $Slide -Left 625 -Top 365 -Width 160 -Height 160 -FillColor $script:Colors.SoftPurple -FillTransparency 0.6 -LineColor $script:Colors.SoftPurple -LineTransparency 1

  $blobA.ZOrder(2) | Out-Null
  $blobB.ZOrder(2) | Out-Null
  $blobC.ZOrder(2) | Out-Null
}

function Add-Header {
  param(
    $Slide,
    [string]$Label,
    [int]$Number,
    [string]$Title,
    [string]$Subtitle = ""
  )

  $pill = Add-RoundedRect -Slide $Slide -Left 42 -Top 22 -Width 120 -Height 24 -FillColor $script:Colors.White -LineColor $script:Colors.Border
  $pill.Fill.Transparency = 0.08
  Add-TextBox -Slide $Slide -Left 54 -Top 26 -Width 96 -Height 14 -Text $Label -FontSize 8.5 -Color $script:Colors.Muted -Bold $true | Out-Null
  Add-TextBox -Slide $Slide -Left 850 -Top 24 -Width 70 -Height 18 -Text ("{0:00} / 15" -f $Number) -FontSize 9 -Color $script:Colors.Muted -Bold $true -Alignment $script:ppAlignRight | Out-Null
  Add-Line -Slide $Slide -X1 42 -Y1 52 -X2 112 -Y2 52 -Color $script:Colors.Blue -Weight 3 -Transparency 0 | Out-Null
  Add-TextBox -Slide $Slide -Left 42 -Top 62 -Width 810 -Height 40 -Text $Title -FontSize 26 -Color $script:Colors.Navy -Bold $true | Out-Null

  if ($Subtitle) {
    Add-TextBox -Slide $Slide -Left 42 -Top 105 -Width 830 -Height 34 -Text $Subtitle -FontSize 12.5 -Color $script:Colors.Muted | Out-Null
  }
}

function Add-CardBlock {
  param(
    $Slide,
    [double]$Left,
    [double]$Top,
    [double]$Width,
    [double]$Height,
    [string]$Kicker = "",
    [string]$Title = "",
    [string]$Body = "",
    [bool]$Dark = $false,
    [double]$TitleSize = 18,
    [double]$BodySize = 11.5
  )

  $fill = if ($Dark) { $script:Colors.Navy } else { $script:Colors.White }
  $line = if ($Dark) { $script:Colors.DarkBlue } else { $script:Colors.BorderStrong }
  $card = Add-RoundedRect -Slide $Slide -Left $Left -Top $Top -Width $Width -Height $Height -FillColor $fill -LineColor $line
  $card.Fill.Transparency = if ($Dark) { 0.02 } else { 0.05 }

  $titleColor = if ($Dark) { $script:Colors.White } else { $script:Colors.Navy }
  $bodyColor = if ($Dark) { $script:Colors.SoftBlue } else { $script:Colors.Muted }
  $kickerColor = if ($Dark) { $script:Colors.SoftBlue } else { $script:Colors.Muted }

  $cursor = $Top + 16
  if ($Kicker) {
    Add-TextBox -Slide $Slide -Left ($Left + 16) -Top $cursor -Width ($Width - 32) -Height 14 -Text $Kicker -FontSize 8.5 -Color $kickerColor -Bold $true | Out-Null
    $cursor += 20
  }

  if ($Title) {
    Add-TextBox -Slide $Slide -Left ($Left + 16) -Top $cursor -Width ($Width - 32) -Height 42 -Text $Title -FontSize $TitleSize -Color $titleColor -Bold $true | Out-Null
    $cursor += 44
  }

  if ($Body) {
    Add-TextBox -Slide $Slide -Left ($Left + 16) -Top $cursor -Width ($Width - 32) -Height ($Height - ($cursor - $Top) - 18) -Text $Body -FontSize $BodySize -Color $bodyColor | Out-Null
  }

  return $card
}

function Add-Callout {
  param(
    $Slide,
    [double]$Left,
    [double]$Top,
    [double]$Width,
    [double]$Height,
    [string]$Text
  )

  $box = Add-RoundedRect -Slide $Slide -Left $Left -Top $Top -Width $Width -Height $Height -FillColor $script:Colors.SoftBlue -FillTransparency 0.1 -LineColor $script:Colors.Border
  Add-TextBox -Slide $Slide -Left ($Left + 18) -Top ($Top + 12) -Width ($Width - 36) -Height ($Height - 20) -Text $Text -FontSize 11.5 -Color $script:Colors.Navy -Bold $true | Out-Null
  return $box
}

function Add-FooterNote {
  param(
    $Slide,
    [string]$LeftText,
    [string]$RightText = ""
  )

  Add-TextBox -Slide $Slide -Left 42 -Top 505 -Width 360 -Height 18 -Text $LeftText -FontSize 8.5 -Color $script:Colors.Muted -Bold $true | Out-Null
  if ($RightText) {
    Add-TextBox -Slide $Slide -Left 550 -Top 505 -Width 330 -Height 18 -Text $RightText -FontSize 8.5 -Color $script:Colors.Muted -Alignment $script:ppAlignRight | Out-Null
  }
}

function Add-BarChart {
  param(
    $Slide,
    [double]$Left,
    [double]$Top,
    [double]$Width,
    [double]$Height,
    [array]$Bars,
    [double]$ScaleMax = 100
  )

  $panel = Add-RoundedRect -Slide $Slide -Left $Left -Top $Top -Width $Width -Height $Height -FillColor $script:Colors.White -LineColor $script:Colors.BorderStrong
  $plotLeft = $Left + 62
  $plotTop = $Top + 48
  $plotWidth = $Width - 92
  $plotHeight = $Height - 112
  $plotBottom = $plotTop + $plotHeight

  foreach ($step in 0..4) {
    $pct = $step * 25
    $scaleValue = [Math]::Round($ScaleMax - (($ScaleMax / 4.0) * $step), 1)
    $label = if ($scaleValue % 1 -eq 0) { "{0}%" -f [int]$scaleValue } else { "{0:N1}%" -f $scaleValue }
    $y = $plotBottom - ($plotHeight * ($pct / 100.0))
    Add-Line -Slide $Slide -X1 $plotLeft -Y1 $y -X2 ($plotLeft + $plotWidth) -Y2 $y -Color $script:Colors.Border -Weight 0.8 -Transparency 0.2 | Out-Null
    Add-TextBox -Slide $Slide -Left ($Left + 10) -Top ($y - 7) -Width 42 -Height 14 -Text $label -FontSize 8 -Color $script:Colors.Muted -Alignment $script:ppAlignRight | Out-Null
  }

  $count = $Bars.Count
  $slot = $plotWidth / $count
  $barWidth = [Math]::Min(68, $slot * 0.55)

  for ($i = 0; $i -lt $count; $i++) {
    $bar = $Bars[$i]
    $x = $plotLeft + ($slot * $i) + (($slot - $barWidth) / 2)
    $barHeight = $plotHeight * ($bar.Value / $ScaleMax)
    $y = $plotBottom - $barHeight

    $shape = $Slide.Shapes.AddShape($script:msoShapeRoundedRectangle, $x, $y, $barWidth, $barHeight)
    $shape.Fill.ForeColor.RGB = $bar.Color
    $shape.Line.Visible = $script:msoFalse

    $valueText = if ($bar.ContainsKey('Text')) { $bar.Text } elseif ($bar.Value % 1 -eq 0) { "{0}%" -f [int]$bar.Value } else { "{0:N1}%" -f $bar.Value }
    Add-TextBox -Slide $Slide -Left ($x - 10) -Top ($y - 22) -Width ($barWidth + 20) -Height 16 -Text $valueText -FontSize 10 -Color $script:Colors.Navy -Bold $true -Alignment $script:ppAlignCenter | Out-Null
    Add-TextBox -Slide $Slide -Left ($x - 25) -Top ($plotBottom + 8) -Width ($barWidth + 50) -Height 36 -Text $bar.Label -FontSize 9 -Color $script:Colors.Muted -Alignment $script:ppAlignCenter | Out-Null
  }
}

function Add-HBar {
  param(
    $Slide,
    [double]$Left,
    [double]$Top,
    [double]$Width,
    [string]$Label,
    [string]$ValueText,
    [double]$Ratio,
    [int]$Color
  )

  Add-TextBox -Slide $Slide -Left $Left -Top $Top -Width ($Width - 72) -Height 16 -Text $Label -FontSize 10.5 -Color $script:Colors.Navy -Bold $true | Out-Null
  Add-TextBox -Slide $Slide -Left ($Left + $Width - 70) -Top $Top -Width 70 -Height 16 -Text $ValueText -FontSize 10.5 -Color $script:Colors.Navy -Bold $true -Alignment $script:ppAlignRight | Out-Null
  Add-RoundedRect -Slide $Slide -Left $Left -Top ($Top + 22) -Width $Width -Height 14 -FillColor $script:Colors.SoftGray -LineColor $script:Colors.SoftGray | Out-Null
  $fill = Add-RoundedRect -Slide $Slide -Left $Left -Top ($Top + 22) -Width ($Width * $Ratio) -Height 14 -FillColor $Color -LineColor $Color
  $fill.Fill.Transparency = 0.05
}

function Add-ComparisonCard {
  param(
    $Slide,
    [double]$Left,
    [double]$Top,
    [double]$Width,
    [double]$Height,
    [string]$System,
    [string]$Feature,
    [string]$Strength,
    [string]$Difference,
    [bool]$Highlight = $false
  )

  $card = Add-RoundedRect -Slide $Slide -Left $Left -Top $Top -Width $Width -Height $Height -FillColor $(if ($Highlight) { $script:Colors.Navy } else { $script:Colors.White }) -LineColor $(if ($Highlight) { $script:Colors.Blue } else { $script:Colors.BorderStrong })
  $lineWeight = if ($Highlight) { [single]1.6 } else { [single]1.0 }
  $card.Line.Weight = $lineWeight
  $titleColor = if ($Highlight) { $script:Colors.White } else { $script:Colors.Navy }
  $bodyColor = if ($Highlight) { $script:Colors.SoftBlue } else { $script:Colors.Muted }

  Add-TextBox -Slide $Slide -Left ($Left + 16) -Top ($Top + 14) -Width ($Width - 32) -Height 18 -Text $System -FontSize 15 -Color $titleColor -Bold $true | Out-Null
  Add-TextBox -Slide $Slide -Left ($Left + 16) -Top ($Top + 42) -Width ($Width - 32) -Height 84 -Text ("주요 특징: {0}`n강점: {1}`n차이점: {2}" -f $Feature, $Strength, $Difference) -FontSize 10.4 -Color $bodyColor | Out-Null
}

function Add-ChatBubble {
  param(
    $Slide,
    [double]$Left,
    [double]$Top,
    [double]$Width,
    [double]$Height,
    [string]$Role,
    [string]$Text,
    [bool]$Right = $false
  )

  $fill = if ($Right) { $script:Colors.SoftBlue } else { $script:Colors.White }
  $line = if ($Right) { $script:Colors.Blue } else { $script:Colors.BorderStrong }
  Add-RoundedRect -Slide $Slide -Left $Left -Top $Top -Width $Width -Height $Height -FillColor $fill -FillTransparency 0.03 -LineColor $line | Out-Null
  Add-TextBox -Slide $Slide -Left ($Left + 14) -Top ($Top + 10) -Width ($Width - 28) -Height 14 -Text $Role -FontSize 8.8 -Color $(if ($Right) { $script:Colors.Blue } else { $script:Colors.Teal }) -Bold $true | Out-Null
  Add-TextBox -Slide $Slide -Left ($Left + 14) -Top ($Top + 28) -Width ($Width - 28) -Height ($Height - 34) -Text $Text -FontSize 10.2 -Color $script:Colors.Navy | Out-Null
}

function Add-DataTableRow {
  param(
    $Slide,
    [double]$Left,
    [double]$Top,
    [double]$Width,
    [string]$Chip,
    [string]$Label,
    [string]$ValueText,
    [double]$Ratio,
    [int]$Color
  )

  $row = Add-RoundedRect -Slide $Slide -Left $Left -Top $Top -Width $Width -Height 44 -FillColor $script:Colors.White -LineColor $script:Colors.BorderStrong
  $row.Fill.Transparency = 0.03

  $chip = Add-RoundedRect -Slide $Slide -Left ($Left + 12) -Top ($Top + 9) -Width 62 -Height 24 -FillColor $Color -LineColor $Color
  Add-TextBox -Slide $Slide -Left ($Left + 12) -Top ($Top + 13) -Width 62 -Height 12 -Text $Chip -FontSize 8.5 -Color $script:Colors.White -Bold $true -Alignment $script:ppAlignCenter | Out-Null
  Add-TextBox -Slide $Slide -Left ($Left + 86) -Top ($Top + 13) -Width 170 -Height 14 -Text $Label -FontSize 10.2 -Color $script:Colors.Navy -Bold $true | Out-Null
  Add-TextBox -Slide $Slide -Left ($Left + 260) -Top ($Top + 13) -Width 54 -Height 14 -Text $ValueText -FontSize 10.5 -Color $script:Colors.Navy -Bold $true -Alignment $script:ppAlignRight | Out-Null

  Add-RoundedRect -Slide $Slide -Left ($Left + 330) -Top ($Top + 17) -Width 110 -Height 10 -FillColor $script:Colors.SoftGray -LineColor $script:Colors.SoftGray | Out-Null
  $fill = Add-RoundedRect -Slide $Slide -Left ($Left + 330) -Top ($Top + 17) -Width (110 * $Ratio) -Height 10 -FillColor $Color -LineColor $Color
}

$output = UniquePath (Join-Path -Path $PSScriptRoot -ChildPath "이론과제#1-발표자료-리디자인.pptx")
$ppt = $null
$presentation = $null

try {
  $ppt = New-Object -ComObject PowerPoint.Application
  $ppt.Visible = $script:msoTrue
  $presentation = $ppt.Presentations.Add()

  while ($presentation.Slides.Count -gt 0) {
    $presentation.Slides.Item(1).Delete()
  }

  $presentation.PageSetup.SlideWidth = 960
  $presentation.PageSetup.SlideHeight = 540

  $slide = $presentation.Slides.Add(1, $script:ppLayoutBlank)
  Apply-Backdrop -Slide $slide
  Add-TextBox -Slide $slide -Left 42 -Top 28 -Width 190 -Height 18 -Text "소프트웨어 시스템 개요 발표" -FontSize 10 -Color $script:Colors.Muted -Bold $true | Out-Null
  Add-Line -Slide $slide -X1 42 -Y1 56 -X2 146 -Y2 56 -Color $script:Colors.Blue -Weight 3 -Transparency 0 | Out-Null
  Add-TextBox -Slide $slide -Left 42 -Top 76 -Width 480 -Height 120 -Text "AI 환각 방지`n모델 오케스트레이션 시스템" -FontSize 31 -Color $script:Colors.Navy -Bold $true | Out-Null
  Add-TextBox -Slide $slide -Left 42 -Top 184 -Width 470 -Height 82 -Text "여러 AI 모델이 하나의 질문을 놓고 서로의 답변을 검토하고 반박한 뒤, 판정 모델이 최종 결론과 남은 쟁점을 정리하는 검증형 질의응답 시스템이다." -FontSize 13.2 -Color $script:Colors.Muted | Out-Null
  Add-RoundedRect -Slide $slide -Left 42 -Top 286 -Width 150 -Height 28 -FillColor $script:Colors.White -LineColor $script:Colors.BorderStrong | Out-Null
  Add-TextBox -Slide $slide -Left 56 -Top 293 -Width 124 -Height 14 -Text "12234195 조진우" -FontSize 10 -Color $script:Colors.Navy -Bold $true | Out-Null
  Add-RoundedRect -Slide $slide -Left 202 -Top 286 -Width 112 -Height 28 -FillColor $script:Colors.White -LineColor $script:Colors.BorderStrong | Out-Null
  Add-TextBox -Slide $slide -Left 214 -Top 293 -Width 88 -Height 14 -Text "Pretendard" -FontSize 10 -Color $script:Colors.Navy -Bold $true -Alignment $script:ppAlignCenter | Out-Null
  foreach ($pill in @(
    @{ L = 42; T = 338; W = 120; Text = "모델 선택 가능" },
    @{ L = 172; T = 338; W = 120; Text = "대화 라운드 설정" },
    @{ L = 302; T = 338; W = 92; Text = "판정 모델" },
    @{ L = 404; T = 338; W = 118; Text = "불확실성 표시" }
  )) {
    Add-RoundedRect -Slide $slide -Left $pill.L -Top $pill.T -Width $pill.W -Height 28 -FillColor $script:Colors.White -LineColor $script:Colors.BorderStrong | Out-Null
    Add-TextBox -Slide $slide -Left ($pill.L + 10) -Top ($pill.T + 7) -Width ($pill.W - 20) -Height 14 -Text $pill.Text -FontSize 9.2 -Color $script:Colors.Navy -Bold $true -Alignment $script:ppAlignCenter | Out-Null
  }
  $heroPanel = Add-RoundedRect -Slide $slide -Left 566 -Top 70 -Width 324 -Height 360 -FillColor $script:Colors.White -LineColor $script:Colors.BorderStrong
  $heroPanel.Fill.Transparency = 0.08
  Add-Oval -Slide $slide -Left 674 -Top 158 -Width 108 -Height 108 -FillColor $script:Colors.SoftBlue -FillTransparency 0.12 -LineColor $script:Colors.Blue -LineTransparency 0.08 -LineWeight 3 | Out-Null
  Add-TextBox -Slide $slide -Left 696 -Top 192 -Width 64 -Height 22 -Text "판정" -FontSize 18 -Color $script:Colors.Blue -Bold $true -Alignment $script:ppAlignCenter | Out-Null
  foreach ($edge in @(
    @{ X1 = 728; Y1 = 212; X2 = 646; Y2 = 124; Color = $script:Colors.Teal },
    @{ X1 = 728; Y1 = 212; X2 = 828; Y2 = 134; Color = $script:Colors.Purple },
    @{ X1 = 728; Y1 = 212; X2 = 636; Y2 = 314; Color = $script:Colors.Blue },
    @{ X1 = 728; Y1 = 212; X2 = 824; Y2 = 318; Color = $script:Colors.Orange }
  )) {
    Add-Line -Slide $slide -X1 $edge.X1 -Y1 $edge.Y1 -X2 $edge.X2 -Y2 $edge.Y2 -Color $edge.Color -Weight 1.5 -Transparency 0.22 | Out-Null
  }
  foreach ($node in @(
    @{ L = 598; T = 100; W = 96; Text = "GPT"; Color = $script:Colors.Navy },
    @{ L = 790; T = 110; W = 96; Text = "Claude"; Color = $script:Colors.Orange },
    @{ L = 592; T = 318; W = 96; Text = "Grok"; Color = $script:Colors.DarkBlue },
    @{ L = 784; T = 324; W = 96; Text = "Gemini"; Color = $script:Colors.Purple }
  )) {
    Add-RoundedRect -Slide $slide -Left $node.L -Top $node.T -Width $node.W -Height 34 -FillColor $script:Colors.White -LineColor $script:Colors.BorderStrong | Out-Null
    Add-TextBox -Slide $slide -Left ($node.L + 12) -Top ($node.T + 9) -Width ($node.W - 24) -Height 14 -Text $node.Text -FontSize 10.5 -Color $node.Color -Bold $true -Alignment $script:ppAlignCenter | Out-Null
  }
  Add-CardBlock -Slide $slide -Left 590 -Top 392 -Width 276 -Height 74 -Kicker "핵심 메시지" -Title "단일 응답보다 검증된 응답" -Body "여러 모델의 검토를 거친 답변이 더 신뢰할 만하다." -TitleSize 17 -BodySize 10.2 | Out-Null
  Add-FooterNote -Slide $slide -LeftText "이론과제 1 발표자료" -RightText "01 / 15"

  $slide = $presentation.Slides.Add(2, $script:ppLayoutBlank)
  Apply-Backdrop -Slide $slide
  Add-Header -Slide $slide -Label "발표 개요" -Number 2 -Title "발표 흐름" -Subtitle "문제 배경과 근거를 먼저 제시한 뒤, 제안 시스템의 개념, 사용자, 핵심 기능, 비교, 결론 순서로 발표한다."
  foreach ($item in @(
    @{ L = 42; K = "01"; Tt = "문제와 근거"; B = "AI 환각이 왜 중요한지와 최신 수치를 먼저 확인한다." },
    @{ L = 266; K = "02"; Tt = "시스템 개념"; B = "무엇을 제안하는지와 작동 구조를 설명한다." },
    @{ L = 490; K = "03"; Tt = "사용자와 기능"; B = "누가 쓰는지와 핵심 기능 세 가지를 정리한다." },
    @{ L = 714; K = "04"; Tt = "비교와 결론"; B = "유사 시스템과의 차이와 최종 메시지를 정리한다." }
  )) {
    Add-CardBlock -Slide $slide -Left $item.L -Top 158 -Width 206 -Height 150 -Kicker $item.K -Title $item.Tt -Body $item.B -TitleSize 18 -BodySize 11.2 | Out-Null
  }
  Add-Callout -Slide $slide -Left 42 -Top 338 -Width 878 -Height 58 -Text "이번 발표의 중심은 '더 똑똑한 단일 모델'이 아니라, '검증 가능한 구조를 가진 AI 시스템'이라는 점이다." | Out-Null
  Add-FooterNote -Slide $slide -LeftText "Agenda" -RightText "02 / 15"

  $slide = $presentation.Slides.Add(3, $script:ppLayoutBlank)
  Apply-Backdrop -Slide $slide
  Add-Header -Slide $slide -Label "문제 배경" -Number 3 -Title "왜 AI 환각을 줄이는 구조가 필요한가"
  Add-CardBlock -Slide $slide -Left 42 -Top 148 -Width 330 -Height 220 -Kicker "문제 정의" -Title "유창한 답변이 항상 정확한 답변은 아니다" -Body "• 그럴듯한 문장은 오류를 더 늦게 발견하게 만든다.`n• 과제, 보고서, 조사처럼 정확성이 중요한 작업에서는 사용자가 매번 직접 검증해야 한다.`n• 결과만 보여주는 AI는 판단 과정이 보이지 않아 신뢰도를 따지기 어렵다." -Dark $true -TitleSize 18 -BodySize 11.1 | Out-Null
  Add-CardBlock -Slide $slide -Left 396 -Top 148 -Width 242 -Height 94 -Title "Fluent" -Body "틀린 내용도 자연스럽게 말하면 신뢰되기 쉽다." -TitleSize 18 -BodySize 11.2 | Out-Null
  Add-CardBlock -Slide $slide -Left 656 -Top 148 -Width 242 -Height 94 -Title "Costly" -Body "사용자가 다시 검색하고 검토하는 비용이 커진다." -TitleSize 18 -BodySize 11.2 | Out-Null
  Add-CardBlock -Slide $slide -Left 396 -Top 266 -Width 242 -Height 94 -Title "Opaque" -Body "판단 과정이 숨겨지면 답변의 질을 평가하기 어렵다." -TitleSize 18 -BodySize 11.2 | Out-Null
  Add-CardBlock -Slide $slide -Left 656 -Top 266 -Width 242 -Height 94 -Title "Risky" -Body "작은 오류 하나가 잘못된 문서와 의사결정으로 이어질 수 있다." -TitleSize 18 -BodySize 11.2 | Out-Null
  Add-Callout -Slide $slide -Left 42 -Top 392 -Width 420 -Height 56 -Text "따라서 필요한 것은 결과를 더 화려하게 만드는 기술이 아니라, 결과를 검증하는 구조이다." | Out-Null
  Add-Callout -Slide $slide -Left 500 -Top 392 -Width 398 -Height 56 -Text "Nature 2024는 hallucination을 사용자가 출력의 정확성을 신뢰하지 못하게 만드는 핵심 문제로 다룬다." | Out-Null
  Add-FooterNote -Slide $slide -LeftText "문제 배경" -RightText "03 / 15"

  $slide = $presentation.Slides.Add(4, $script:ppLayoutBlank)
  Apply-Backdrop -Slide $slide
  Add-Header -Slide $slide -Label "근거 1" -Number 4 -Title "최신 모델도 hallucination을 완전히 없애지 못했다" -Subtitle "Vectara Hallucination Leaderboard의 2026년 3월 20일 스냅샷은 최신 프런티어 모델도 여전히 hallucination을 낸다는 점을 보여준다."
  Add-BarChart -Slide $slide -Left 42 -Top 154 -Width 432 -Height 250 -Bars @(
    @{ Label = "GPT-5.4"; Value = 7.0; Text = "7.0%"; Color = $script:Colors.Navy },
    @{ Label = "Gemini 2.5 Pro"; Value = 7.0; Text = "7.0%"; Color = $script:Colors.Teal },
    @{ Label = "Claude Sonnet 4.6"; Value = 10.6; Text = "10.6%"; Color = $script:Colors.Purple },
    @{ Label = "Grok 4.1 Fast"; Value = 17.8; Text = "17.8%"; Color = $script:Colors.Amber }
  ) -ScaleMax 20
  Add-RoundedRect -Slide $slide -Left 502 -Top 154 -Width 418 -Height 250 -FillColor $script:Colors.White -LineColor $script:Colors.BorderStrong | Out-Null
  Add-TextBox -Slide $slide -Left 522 -Top 172 -Width 180 -Height 16 -Text "모델별 수치 표" -FontSize 9.2 -Color $script:Colors.Muted -Bold $true | Out-Null
  Add-TextBox -Slide $slide -Left 522 -Top 194 -Width 170 -Height 18 -Text "rate와 상대 위치" -FontSize 17 -Color $script:Colors.Navy -Bold $true | Out-Null
  Add-DataTableRow -Slide $slide -Left 522 -Top 226 -Width 378 -Chip "GPT" -Label "GPT-5.4" -ValueText "7.0%" -Ratio 0.35 -Color $script:Colors.Navy
  Add-DataTableRow -Slide $slide -Left 522 -Top 276 -Width 378 -Chip "Gemini" -Label "Gemini 2.5 Pro" -ValueText "7.0%" -Ratio 0.35 -Color $script:Colors.Teal
  Add-DataTableRow -Slide $slide -Left 522 -Top 326 -Width 378 -Chip "Claude" -Label "Claude Sonnet 4.6" -ValueText "10.6%" -Ratio 0.53 -Color $script:Colors.Purple
  Add-DataTableRow -Slide $slide -Left 522 -Top 376 -Width 378 -Chip "Grok" -Label "Grok 4.1 Fast" -ValueText "17.8%" -Ratio 0.89 -Color $script:Colors.Amber
  Add-Callout -Slide $slide -Left 42 -Top 428 -Width 878 -Height 46 -Text "같은 최신 세대 모델 사이에서도 rate 차이가 있고, 가장 낮은 수치 역시 0%와는 거리가 남아 있다. 따라서 최신 모델 하나만 믿는 구조는 충분하지 않다." | Out-Null
  Add-FooterNote -Slide $slide -LeftText "Vectara Hallucination Leaderboard" -RightText "04 / 15"

  $slide = $presentation.Slides.Add(5, $script:ppLayoutBlank)
  Apply-Backdrop -Slide $slide
  Add-Header -Slide $slide -Label "근거 2" -Number 5 -Title "낮은 빈도라도 영향이 크면 환각은 위험하다" -Subtitle "의료 요약 분야 2025년 연구는 hallucination과 omission이 존재하며, 그중 일부는 major 오류가 될 수 있음을 보여준다."
  Add-CardBlock -Slide $slide -Left 42 -Top 154 -Width 418 -Height 248 -Kicker "문장 단위 오류율" -Title "수치가 낮아 보여도 고위험 영역에서는 무시할 수 없다" -Body "환각 1.47%, 생략 3.45%라는 수치는 작아 보일 수 있다. 그러나 정확성 비용이 큰 환경에서는 작은 비율도 중요한 리스크가 된다." -TitleSize 17 -BodySize 11 | Out-Null
  Add-HBar -Slide $slide -Left 64 -Top 272 -Width 374 -Label "Hallucination rate" -ValueText "1.47%" -Ratio 0.4261 -Color $script:Colors.Blue
  Add-HBar -Slide $slide -Left 64 -Top 340 -Width 374 -Label "Omission rate" -ValueText "3.45%" -Ratio 1 -Color $script:Colors.Teal
  Add-CardBlock -Slide $slide -Left 496 -Top 154 -Width 424 -Height 248 -Kicker "Major risk" -Title "환각 문장 191개 중 84개가 major 오류였다" -Body "오류가 드물어 보여도, 그 영향이 크면 시스템 위험은 충분히 심각하다. 그래서 시스템은 정답 생성뿐 아니라 검토와 경고 구조를 함께 가져야 한다." -TitleSize 17 -BodySize 11.2 | Out-Null
  Add-Oval -Slide $slide -Left 718 -Top 228 -Width 132 -Height 132 -FillColor $script:Colors.White -FillTransparency 1 -LineColor $script:Colors.Blue -LineTransparency 0.08 -LineWeight 16 | Out-Null
  Add-TextBox -Slide $slide -Left 744 -Top 266 -Width 82 -Height 24 -Text "44%" -FontSize 30 -Color $script:Colors.Navy -Bold $true -Alignment $script:ppAlignCenter | Out-Null
  Add-TextBox -Slide $slide -Left 716 -Top 302 -Width 136 -Height 18 -Text "major 오류 비중" -FontSize 10.2 -Color $script:Colors.Muted -Bold $true -Alignment $script:ppAlignCenter | Out-Null
  Add-Callout -Slide $slide -Left 42 -Top 428 -Width 878 -Height 46 -Text "즉 환각은 '드물다'는 이유만으로 무시할 수 없다. 위험이 큰 영역일수록 답변을 상호검증하는 구조가 더 중요하다." | Out-Null
  Add-FooterNote -Slide $slide -LeftText "npj Digital Medicine (2025)" -RightText "05 / 15"

  $slide = $presentation.Slides.Add(6, $script:ppLayoutBlank)
  Apply-Backdrop -Slide $slide
  Add-Header -Slide $slide -Label "시스템명" -Number 6 -Title "시스템명과 핵심 정의"
  Add-CardBlock -Slide $slide -Left 42 -Top 160 -Width 314 -Height 214 -Kicker "시스템명" -Title "AI 환각 방지 모델 오케스트레이션 시스템" -Body "영문명은 AI Hallucination Mitigation Model Orchestration System이며, 시스템이 하는 일을 이름에 직접 반영했다." -Dark $true -TitleSize 21 -BodySize 11.3 | Out-Null
  Add-CardBlock -Slide $slide -Left 386 -Top 160 -Width 534 -Height 214 -Kicker "한 줄 정의" -Title "사용자가 모델 조합과 대화 횟수를 직접 정하는 검증형 AI 시스템" -Body "• 여러 모델이 하나의 질문을 놓고 검토, 반박, 보완을 반복한다.`n• 판정 모델이 최종 답변, 확신도, 남은 쟁점을 정리한다.`n• 핵심은 정답을 더 화려하게 만드는 것이 아니라, 답변이 검토되는 구조를 만드는 데 있다." -TitleSize 18 -BodySize 11.3 | Out-Null
  Add-Callout -Slide $slide -Left 42 -Top 404 -Width 878 -Height 46 -Text "즉 이 시스템은 'AI가 말한 답'보다 '검증을 거친 답'을 제공하려는 방향의 소프트웨어 시스템이다." | Out-Null
  Add-FooterNote -Slide $slide -LeftText "시스템명" -RightText "06 / 15"

  $slide = $presentation.Slides.Add(7, $script:ppLayoutBlank)
  Apply-Backdrop -Slide $slide
  Add-Header -Slide $slide -Label "시스템 개념" -Number 7 -Title "초안 · 토론 · 판정의 역할을 분리해 검증 구조를 만든다"
  Add-CardBlock -Slide $slide -Left 42 -Top 170 -Width 270 -Height 194 -Kicker "1단계" -Title "초안 모델" -Body "첫 답변을 생성한다. 문제 정의와 핵심 설명을 먼저 제시하는 역할을 맡는다." -TitleSize 19 -BodySize 11.4 | Out-Null
  Add-CardBlock -Slide $slide -Left 345 -Top 170 -Width 270 -Height 194 -Kicker "2단계" -Title "토론 모델" -Body "사실성, 누락 조건, 반례 가능성을 검토한다. 서로의 주장에 반박하며 답변을 보완한다." -TitleSize 19 -BodySize 11.4 | Out-Null
  Add-CardBlock -Slide $slide -Left 648 -Top 170 -Width 270 -Height 194 -Kicker "3단계" -Title "판정 모델" -Body "전체 대화를 읽고 최종 답변, 채택 이유, 불확실성, 남은 쟁점을 정리한다." -TitleSize 19 -BodySize 11.4 | Out-Null
  Add-Line -Slide $slide -X1 312 -Y1 266 -X2 344 -Y2 266 -Color $script:Colors.Blue -Weight 2 -Transparency 0.15 | Out-Null
  Add-Line -Slide $slide -X1 615 -Y1 266 -X2 647 -Y2 266 -Color $script:Colors.Blue -Weight 2 -Transparency 0.15 | Out-Null
  Add-Callout -Slide $slide -Left 42 -Top 400 -Width 878 -Height 52 -Text "이 구조의 장점은 모델이 틀릴 수 있다는 점을 시스템 설계에 전제로 반영한다는 것이다. 즉 하나의 답을 바로 믿지 않고, 검토 과정을 통해 신뢰도를 높인다." | Out-Null
  Add-FooterNote -Slide $slide -LeftText "시스템 설명" -RightText "07 / 15"

  $slide = $presentation.Slides.Add(8, $script:ppLayoutBlank)
  Apply-Backdrop -Slide $slide
  Add-Header -Slide $slide -Label "작동 흐름" -Number 8 -Title "실제 사용 흐름은 5단계로 설명할 수 있다"
  $flowX = 42
  foreach ($step in @(
    @{ N = "1"; T = "질문 입력"; B = "사용자가 검증이 필요한 질문을 입력한다." },
    @{ N = "2"; T = "모델 선택"; B = "참여 모델과 라운드 수를 직접 설정한다." },
    @{ N = "3"; T = "초안 생성"; B = "첫 번째 모델이 기본 답변을 작성한다." },
    @{ N = "4"; T = "상호검증"; B = "다른 모델이 반박과 보완을 반복한다." },
    @{ N = "5"; T = "최종 판정"; B = "판정 모델이 최종 답과 남은 쟁점을 정리한다." }
  )) {
    Add-RoundedRect -Slide $slide -Left $flowX -Top 188 -Width 168 -Height 174 -FillColor $script:Colors.White -LineColor $script:Colors.BorderStrong | Out-Null
    Add-Oval -Slide $slide -Left ($flowX + 16) -Top 204 -Width 28 -Height 28 -FillColor $script:Colors.SoftBlue -FillTransparency 0 -LineColor $script:Colors.Blue -LineTransparency 0.1 -LineWeight 1 | Out-Null
    Add-TextBox -Slide $slide -Left ($flowX + 16) -Top 209 -Width 28 -Height 14 -Text $step.N -FontSize 11 -Color $script:Colors.Blue -Bold $true -Alignment $script:ppAlignCenter | Out-Null
    Add-TextBox -Slide $slide -Left ($flowX + 16) -Top 244 -Width 136 -Height 20 -Text $step.T -FontSize 12.5 -Color $script:Colors.Navy -Bold $true | Out-Null
    Add-TextBox -Slide $slide -Left ($flowX + 16) -Top 272 -Width 136 -Height 56 -Text $step.B -FontSize 10.3 -Color $script:Colors.Muted | Out-Null
    $flowX += 178
  }
  Add-Callout -Slide $slide -Left 42 -Top 400 -Width 878 -Height 50 -Text "이 시스템은 사용자가 라운드 수를 정한다는 점이 중요하다. 즉 검토 강도와 시간 비용을 사용자가 스스로 조절할 수 있다." | Out-Null
  Add-FooterNote -Slide $slide -LeftText "시스템 설명" -RightText "08 / 15"

  $slide = $presentation.Slides.Add(9, $script:ppLayoutBlank)
  Apply-Backdrop -Slide $slide
  Add-Header -Slide $slide -Label "시스템 사용자 정의" -Number 9 -Title "누가 언제 어디서 어떻게 쓰는가"
  foreach ($p in @(
    @{ L = 42; K = "누가"; Tt = "대학생과 학습자"; B = "과제, 보고서, 발표 자료를 준비하는 사용자" },
    @{ L = 268; K = "또 누가"; Tt = "연구자와 실무 사용자"; B = "자료 조사와 검토 비용이 큰 사용자" },
    @{ L = 494; K = "언제"; Tt = "정확성이 중요한 질문"; B = "여러 관점을 비교하고 싶은 상황" },
    @{ L = 720; K = "어디서·어떻게"; Tt = "웹 기반 인터페이스"; B = "질문 입력 → 모델 선택 → 토론 확인 → 결과 검토" }
  )) {
    Add-CardBlock -Slide $slide -Left $p.L -Top 156 -Width 190 -Height 138 -Kicker $p.K -Title $p.Tt -Body $p.B -TitleSize 16 -BodySize 10.8 | Out-Null
  }
  Add-CardBlock -Slide $slide -Left 42 -Top 326 -Width 420 -Height 118 -Kicker "사용 맥락" -Title "결과보다 검토 과정이 중요한 사용자에게 적합" -Body "이 시스템은 빠른 답만 원하는 사람보다, 답을 믿어도 되는지 확인해야 하는 사람에게 더 적합하다." -TitleSize 17 -BodySize 11.2 | Out-Null
  Add-CardBlock -Slide $slide -Left 500 -Top 326 -Width 420 -Height 118 -Kicker "사용 방식" -Title "복잡한 로직 없이 선택만으로 사용 가능" -Body "사용자는 오케스트레이션 내부 구조를 몰라도 모델과 라운드 수만 선택하면 된다." -TitleSize 17 -BodySize 11.2 | Out-Null
  Add-FooterNote -Slide $slide -LeftText "시스템 사용자 정의" -RightText "09 / 15"

  $slide = $presentation.Slides.Add(10, $script:ppLayoutBlank)
  Apply-Backdrop -Slide $slide
  Add-Header -Slide $slide -Label "사용 장면" -Number 10 -Title "사용자는 이런 장면에서 시스템을 활용할 수 있다"
  Add-CardBlock -Slide $slide -Left 42 -Top 166 -Width 282 -Height 214 -Kicker "대학생 시나리오" -Title "과제 자료 조사 전" -Body "여러 모델의 토론을 통해 개념 정의가 일치하는지, 누락된 조건이 없는지 먼저 확인한 뒤 과제를 작성한다.`n`n정확도 확인 / 과제 작성 전 점검" -TitleSize 18 -BodySize 11.1 | Out-Null
  Add-CardBlock -Slide $slide -Left 339 -Top 166 -Width 282 -Height 214 -Kicker "연구·학습 시나리오" -Title "새로운 개념을 이해할 때" -Body "하나의 설명만 듣는 대신 여러 모델이 같은 질문을 어떻게 해석하는지 비교하며 개념을 입체적으로 이해한다.`n`n개념 비교 / 오해 방지" -TitleSize 18 -BodySize 11.1 | Out-Null
  Add-CardBlock -Slide $slide -Left 636 -Top 166 -Width 282 -Height 214 -Kicker "실무 시나리오" -Title "기획 문서 초안 작성 전" -Body "여러 모델이 지적한 쟁점과 보완 포인트를 먼저 본 뒤, 정리된 답변을 문서 초안 작성에 활용한다.`n`n문서 품질 향상 / 검토 시간 절감" -TitleSize 18 -BodySize 11.1 | Out-Null
  Add-FooterNote -Slide $slide -LeftText "사용 장면" -RightText "10 / 15"

  $slide = $presentation.Slides.Add(11, $script:ppLayoutBlank)
  Apply-Backdrop -Slide $slide
  Add-Header -Slide $slide -Label "핵심 기능 1" -Number 11 -Title "모델 선택과 라운드 제어"
  Add-CardBlock -Slide $slide -Left 42 -Top 160 -Width 308 -Height 238 -Kicker "사용자 제어" -Title "참여 모델과 검토 강도를 직접 정한다" -Body "사용자는 초안 모델, 토론 모델, 판정 모델의 조합을 선택할 수 있고, 대화 라운드 수를 정해 검토 강도와 시간 비용을 조절할 수 있다." -Dark $true -TitleSize 19 -BodySize 11.3 | Out-Null
  Add-RoundedRect -Slide $slide -Left 388 -Top 160 -Width 532 -Height 238 -FillColor $script:Colors.White -LineColor $script:Colors.BorderStrong | Out-Null
  Add-TextBox -Slide $slide -Left 412 -Top 182 -Width 210 -Height 16 -Text "설정 예시" -FontSize 9.2 -Color $script:Colors.Muted -Bold $true | Out-Null
  Add-TextBox -Slide $slide -Left 412 -Top 206 -Width 180 -Height 18 -Text "모델 조합" -FontSize 18 -Color $script:Colors.Navy -Bold $true | Out-Null
  foreach ($opt in @(
    @{ L = 412; T = 238; W = 118; Text = "초안: GPT"; Color = $script:Colors.Navy },
    @{ L = 540; T = 238; W = 138; Text = "토론: Claude"; Color = $script:Colors.Orange },
    @{ L = 688; T = 238; W = 138; Text = "판정: Gemini"; Color = $script:Colors.Purple }
  )) {
    Add-RoundedRect -Slide $slide -Left $opt.L -Top $opt.T -Width $opt.W -Height 34 -FillColor $script:Colors.White -LineColor $script:Colors.BorderStrong | Out-Null
    Add-TextBox -Slide $slide -Left ($opt.L + 12) -Top ($opt.T + 9) -Width ($opt.W - 24) -Height 14 -Text $opt.Text -FontSize 10 -Color $opt.Color -Bold $true -Alignment $script:ppAlignCenter | Out-Null
  }
  Add-TextBox -Slide $slide -Left 412 -Top 300 -Width 180 -Height 18 -Text "대화 라운드 수" -FontSize 18 -Color $script:Colors.Navy -Bold $true | Out-Null
  foreach ($round in @(
    @{ L = 412; Text = "1회" },
    @{ L = 482; Text = "2회" },
    @{ L = 552; Text = "3회" },
    @{ L = 622; Text = "4회" }
  )) {
    Add-RoundedRect -Slide $slide -Left $round.L -Top 334 -Width 58 -Height 30 -FillColor $(if ($round.Text -eq "3회") { $script:Colors.SoftBlue } else { $script:Colors.White }) -LineColor $(if ($round.Text -eq "3회") { $script:Colors.Blue } else { $script:Colors.BorderStrong }) | Out-Null
    Add-TextBox -Slide $slide -Left $round.L -Top 342 -Width 58 -Height 12 -Text $round.Text -FontSize 10 -Color $script:Colors.Navy -Bold $true -Alignment $script:ppAlignCenter | Out-Null
  }
  Add-FooterNote -Slide $slide -LeftText "핵심 기능" -RightText "11 / 15"

  $slide = $presentation.Slides.Add(12, $script:ppLayoutBlank)
  Apply-Backdrop -Slide $slide
  Add-Header -Slide $slide -Label "핵심 기능 2" -Number 12 -Title "검증 과정을 시각적으로 보여준다"
  Add-CardBlock -Slide $slide -Left 42 -Top 154 -Width 350 -Height 270 -Kicker "메신저형 UI" -Title "토론 과정을 말풍선 흐름으로 보여준다" -Body "최종 답변만 보여주는 대신, 각 모델이 어떤 이유로 다른 모델의 주장을 수정하거나 반박했는지를 실시간으로 보여준다.`n`n이 기능은 설명 가능성과 신뢰도를 높이는 핵심 장치이다." -TitleSize 18 -BodySize 11.2 | Out-Null
  Add-RoundedRect -Slide $slide -Left 422 -Top 154 -Width 498 -Height 270 -FillColor $script:Colors.White -LineColor $script:Colors.BorderStrong | Out-Null
  Add-ChatBubble -Slide $slide -Left 444 -Top 180 -Width 294 -Height 56 -Role "초안 모델" -Text "이 질문의 핵심은 A로 볼 수 있습니다. 다만 B 조건은 아직 충분히 반영되지 않았습니다."
  Add-ChatBubble -Slide $slide -Left 560 -Top 246 -Width 332 -Height 56 -Role "토론 모델 1" -Text "A 설명은 가능하지만 방금 답변은 B 조건과 C 예외를 빠뜨렸습니다." -Right $true
  Add-ChatBubble -Slide $slide -Left 444 -Top 312 -Width 322 -Height 56 -Role "토론 모델 2" -Text "C 예외뿐 아니라 D 사례도 함께 보완해야 합니다."
  Add-ChatBubble -Slide $slide -Left 560 -Top 378 -Width 332 -Height 56 -Role "판정 모델" -Text "최종적으로는 A를 중심으로 설명하되 B, C, D 조건을 함께 적는 것이 가장 안전합니다." -Right $true
  Add-FooterNote -Slide $slide -LeftText "핵심 기능" -RightText "12 / 15"

  $slide = $presentation.Slides.Add(13, $script:ppLayoutBlank)
  Apply-Backdrop -Slide $slide
  Add-Header -Slide $slide -Label "핵심 기능 3" -Number 13 -Title "판정과 불확실성 표시"
  Add-CardBlock -Slide $slide -Left 42 -Top 158 -Width 282 -Height 174 -Kicker "기능 1" -Title "판정 모델 최종 정리" -Body "전체 대화를 읽고 최종 답변과 채택 이유를 정리한다." -TitleSize 17 -BodySize 11.3 | Out-Null
  Add-CardBlock -Slide $slide -Left 339 -Top 158 -Width 282 -Height 174 -Kicker "기능 2" -Title "확신도와 남은 쟁점" -Body "모델 간 합의가 부족하면 확신도를 낮게 표시하고, 아직 해소되지 않은 쟁점을 따로 제시한다." -TitleSize 17 -BodySize 11.3 | Out-Null
  Add-CardBlock -Slide $slide -Left 636 -Top 158 -Width 282 -Height 174 -Kicker "기능 3" -Title "판단 보류" -Body "근거가 부족하거나 의견 충돌이 크면 무리하게 하나의 답을 내지 않고 보류 상태를 표시한다." -TitleSize 17 -BodySize 11.3 | Out-Null
  Add-CardBlock -Slide $slide -Left 42 -Top 358 -Width 420 -Height 100 -Kicker "설계 원칙" -Title "모른다는 사실을 숨기지 않는 시스템" -Body "좋은 시스템은 항상 답을 내는 시스템이 아니라, 답하기 어려운 경우를 정직하게 표현하는 시스템이다." -TitleSize 16.5 -BodySize 10.9 | Out-Null
  Add-CardBlock -Slide $slide -Left 500 -Top 358 -Width 420 -Height 100 -Kicker "기대 효과" -Title "결과 품질과 신뢰성을 함께 관리" -Body "답변 내용뿐 아니라 답변의 확실성까지 함께 보여주기 때문에 사용자가 결과를 더 안전하게 활용할 수 있다." -TitleSize 16.5 -BodySize 10.9 | Out-Null
  Add-FooterNote -Slide $slide -LeftText "핵심 기능" -RightText "13 / 15"

  $slide = $presentation.Slides.Add(14, $script:ppLayoutBlank)
  Apply-Backdrop -Slide $slide
  Add-Header -Slide $slide -Label "유사 시스템 비교" -Number 14 -Title "유사 관련 시스템과 비교했을 때 무엇이 다른가"
  Add-ComparisonCard -Slide $slide -Left 42 -Top 154 -Width 420 -Height 126 -System "Microsoft AutoGen" -Feature "여러 에이전트를 역할별로 연결하는 프레임워크" -Strength "다중 에이전트 상호작용 구조에 강하다" -Difference "일반 사용자가 모델과 라운드 수를 직접 선택하는 UI는 기본이 아니다"
  Add-ComparisonCard -Slide $slide -Left 500 -Top 154 -Width 420 -Height 126 -System "CrewAI" -Feature "AI 에이전트 간 역할 분담과 협업 실행" -Strength "역할 기반 협업 구조를 만들기 쉽다" -Difference "환각 억제와 불확실성 표시를 전면 목표로 두지는 않는다"
  Add-ComparisonCard -Slide $slide -Left 42 -Top 300 -Width 420 -Height 126 -System "Perplexity AI" -Feature "검색 기반 답변과 근거 제시 중심 서비스" -Strength "빠른 정보 탐색과 출처 기반 응답에 강하다" -Difference "모델 간 토론 과정을 사용자에게 보여주는 구조와는 다르다"
  Add-ComparisonCard -Slide $slide -Left 500 -Top 300 -Width 420 -Height 126 -System "제안 시스템" -Feature "사용자가 모델 조합과 라운드 수를 정하고, 토론과 판정 결과를 함께 보는 시스템" -Strength "사용자 제어, 토론 시각화, 판정 모델, 불확실성 표시를 결합" -Difference "다중 모델 협업을 일반 사용자가 직접 쓰는 검증형 인터페이스로 바꾼다" -Highlight $true
  Add-Callout -Slide $slide -Left 42 -Top 448 -Width 878 -Height 46 -Text "기존 시스템이 협업 프레임워크나 검색 중심 서비스에 강점이 있다면, 제안 시스템은 그것을 사용자가 직접 조절할 수 있는 환각 저감 인터페이스로 만든다는 점에서 차별화된다." | Out-Null
  Add-FooterNote -Slide $slide -LeftText "유사 관련 시스템" -RightText "14 / 15"

  $slide = $presentation.Slides.Add(15, $script:ppLayoutBlank)
  Apply-Backdrop -Slide $slide
  Add-Header -Slide $slide -Label "결론" -Number 15 -Title "결론: 단일 응답보다 검증된 응답이 중요하다"
  Add-CardBlock -Slide $slide -Left 42 -Top 154 -Width 398 -Height 280 -Kicker "최종 정리" -Title "이 발표의 핵심 결론" -Body "• AI 환각은 여전히 현실적인 문제다.`n• 최신 모델도 hallucination을 완전히 없애지 못했다.`n• 따라서 필요한 것은 답변을 더 길게 만드는 기술이 아니라, 답변을 검토하는 구조와 불확실성을 표시하는 구조이다.`n• 제안 시스템은 모델 선택, 라운드 설정, 토론 시각화, 판정 모델, 판단 보류 기능으로 이를 해결하려 한다." -Dark $true -TitleSize 18 -BodySize 10.9 | Out-Null
  Add-CardBlock -Slide $slide -Left 468 -Top 154 -Width 452 -Height 280 -Kicker "참고 자료" -Title "발표에 사용한 주요 자료" -Body "1. Vectara Hallucination Leaderboard, updated Mar 20, 2026.`n2. Farquhar et al., Nature, Detecting hallucinations in large language models using semantic entropy, 2024.`n3. Asgari et al., npj Digital Medicine, A framework to assess clinical safety and hallucination rates of LLMs for medical text summarisation, 2025.`n`n그래프와 설명은 위 자료의 공개 수치를 바탕으로 발표용으로 재구성했다." -TitleSize 18 -BodySize 10.7 | Out-Null
  Add-FooterNote -Slide $slide -LeftText "발표 마무리" -RightText "15 / 15"

  $presentation.SaveAs($output, 24)
  $presentation.Close()
  $ppt.Quit()
  "CREATED: $output"
}
finally {
  if ($presentation) {
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation)
  }
  if ($ppt) {
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
