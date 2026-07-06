param(
  [Parameter(Mandatory=$true)][string]$ManifestPath,
  [Parameter(Mandatory=$true)][string]$OutputDir,
  [Parameter(Mandatory=$true)][string]$TemplateDir
)

$ErrorActionPreference = 'Stop'
$manifest = Get-Content -LiteralPath $ManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

function Safe-Part([string]$text, [string]$fallback) {
  $value = if ([string]::IsNullOrWhiteSpace($text)) { $fallback } else { $text }
  return ($value -replace '[\/:*?"<>|]', '_' -replace '\s+', '')
}

function Unique-Path([string]$path) {
  if (!(Test-Path -LiteralPath $path)) { return $path }
  $dir = [IO.Path]::GetDirectoryName($path)
  $base = [IO.Path]::GetFileNameWithoutExtension($path)
  $ext = [IO.Path]::GetExtension($path)
  $n = 2
  do {
    $candidate = Join-Path $dir ($base + '_' + $n + $ext)
    $n++
  } while (Test-Path -LiteralPath $candidate)
  return $candidate
}

function Sheet-Like($workbook, [string]$pattern) {
  foreach ($ws in @($workbook.Worksheets)) {
    if ($ws.Name -match $pattern) { return $ws }
  }
  return $null
}

function T($sheet, [string]$addr) {
  try { return [string]$sheet.Range($addr).Text } catch { return '' }
}

function Put($sheet, [string]$addr, $value) {
  if ($null -eq $value) { $value = '' }
  $sheet.Range($addr).Value2 = $value
}

function PutDash($sheet, [string]$addr, $value) {
  $s = ([string]$value).Trim()
  if ([string]::IsNullOrWhiteSpace($s) -or $s -eq '/') { $s = '-' }
  $sheet.Range($addr).Value2 = $s
}

function Addr([bool]$condition, [string]$whenTrue, [string]$whenFalse) {
  if ($condition) { return $whenTrue }
  return $whenFalse
}

function First-NonBlank($values) {
  foreach ($v in $values) {
    if ($null -ne $v -and -not [string]::IsNullOrWhiteSpace([string]$v)) { return $v }
  }
  return ''
}

function First-Useful($values) {
  foreach ($v in $values) {
    $s = [string]$v
    if (-not [string]::IsNullOrWhiteSpace($s) -and $s -notmatch '^(\/|Service|Clean|\*)$') { return $v }
  }
  return ''
}

function Num($value) {
  $s = ([string]$value).Trim() -replace ',', ''
  $n = 0.0
  if ([double]::TryParse($s, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$n)) { return $n }
  return $null
}

function First-NumberText($values) {
  foreach ($v in $values) {
    $s = ([string]$v).Trim()
    if ([string]::IsNullOrWhiteSpace($s)) { continue }
    if ($null -ne (Num $s)) { return $s }
  }
  return ''
}

function Add-One($value) {
  $n = Num $value
  if ($null -eq $n) { return $value }
  return [int]($n + 1)
}

function Max-Number($values, [double]$default) {
  $max = $default
  foreach ($v in $values) {
    $n = Num $v
    if ($null -ne $n -and $n -gt $max) { $max = $n }
  }
  return $max
}

function Min-Number($values, [double]$default) {
  $min = $default
  foreach ($v in $values) {
    $n = Num $v
    if ($null -ne $n -and $n -lt $min) { $min = $n }
  }
  return $min
}

function Model-FromName([string]$name) {
  $m = [regex]::Match($name, '(ZS025|KC12|M\d{2,3}[A-Z]*|MX\d+[A-Z]*|MC\d+|MS\d+|MA\d+)', 'IgnoreCase')
  if ($m.Success) { return $m.Value.ToUpper() }
  return ''
}

function Plates-FromName([string]$name) {
  $m = [regex]::Match($name, '(\d+)\s*pl', 'IgnoreCase')
  if ($m.Success) { return $m.Groups[1].Value }
  return ''
}

function Clear-KnownSlashCells($sheet, [bool]$isGphe) {
  $cells = @('D12','H12')
  foreach ($addr in $cells) { Put $sheet $addr '' }
}

function Override-Value($item, [string]$key, $fallback = '') {
  if ($null -ne $item.overrides) {
    $prop = $item.overrides.PSObject.Properties[$key]
    if ($null -ne $prop -and -not [string]::IsNullOrWhiteSpace([string]$prop.Value)) { return $prop.Value }
  }
  return $fallback
}

function Unit-Text([string]$unit) {
  $u = ([string]$unit).Trim()
  if ([string]::IsNullOrWhiteSpace($u)) { return '' }
  if ($u.StartsWith('(') -and $u.EndsWith(')')) { return $u }
  return '(' + $u + ')'
}

function Split-Pair([string]$text) {
  $parts = ([string]$text) -split '\s*/\s*|\s*,\s*'
  if ($parts.Count -ge 2) { return @($parts[0].Trim(), $parts[1].Trim()) }
  return @(([string]$text).Trim(), '')
}

function Set-ModelPrefix([string]$model, [string]$material) {
  $m = ([string]$model).Trim()
  if ($m -notmatch '^(MC|MS)') { return $m }
  if ($material -match 'Stainless|SUS|^MS$') { return ($m -replace '^(MC|MS)', 'MS') }
  if ($material -match 'Copper|^MC$') { return ($m -replace '^(MC|MS)', 'MC') }
  return $m
}

function Display-Solder([string]$material, [string]$model) {
  if ($material -match 'Stainless|SUS|^MS$') { return 'Stainless Steel' }
  if ($material -match 'Copper|^MC$') { return 'Copper' }
  if ($model -match '^MS') { return 'Stainless Steel' }
  return 'Copper'
}

function Heat-Factor([string]$unit) {
  $u = (([string]$unit).Trim('(',')',' ') -replace '\s+', '').ToLowerInvariant()
  switch -Regex ($u) {
    '^kw$' { return 1000.0 }
    '^w$' { return 1.0 }
    '^(kcal/hr|kcalh|kcal/hour)$' { return 4186.8 / 3600.0 }
    '^(btu/hr|btuh|btu/hour)$' { return 0.29307107 }
    default { return $null }
  }
}

function Pressure-Factor([string]$unit) {
  $u = (([string]$unit).Trim('(',')',' ') -replace '\s+', '' -replace '_', '' -replace '-', '').ToLowerInvariant()
  if ($u -match '^(pa|kpa|mpa|bar|kgf/cm2|kgfcm2|atm|psi|torr)(g|a)$') { $u = $matches[1] }
  switch -Regex ($u) {
    '^pa$' { return 1.0 }
    '^kpa$' { return 1000.0 }
    '^mpa$' { return 1000000.0 }
    '^bar$' { return 100000.0 }
    '^kgf/cm2$|^kgfcm2$' { return 98066.5 }
    '^atm$' { return 101325.0 }
    '^psi$' { return 6894.757293168 }
    '^torr$' { return 133.322368421 }
    default { return $null }
  }
}

function Convert-UnitValue($value, [string]$fromUnit, [string]$toUnit, [string]$kind) {
  $n = Num $value
  if ($null -eq $n) { return $value }
  $from = if ($kind -eq 'heat') { Heat-Factor $fromUnit } else { Pressure-Factor $fromUnit }
  $to = if ($kind -eq 'heat') { Heat-Factor $toUnit } else { Pressure-Factor $toUnit }
  if ($null -eq $from -or $null -eq $to -or $to -eq 0) { return $value }
  return [Math]::Round(($n * $from / $to), 6)
}

function Convert-Cells($sheet, [string[]]$cells, [string]$unitCell, [string]$targetUnit, [string]$kind) {
  if ([string]::IsNullOrWhiteSpace($targetUnit)) { return }
  $fromUnit = T $sheet $unitCell
  foreach ($addr in $cells) {
    Put $sheet $addr (Convert-UnitValue (T $sheet $addr) $fromUnit $targetUnit $kind)
    try { $sheet.Range($addr).NumberFormat = '0.######' } catch {}
  }
  Put $sheet $unitCell (Unit-Text $targetUnit)
}

function Apply-HtriOverrides($ds, $item, [bool]$isGphe) {
  Put $ds 'B1' (Override-Value $item 'Customer' (T $ds 'B1'))
  Put $ds 'B2' (Override-Value $item 'Project Name' (T $ds 'B2'))
  Put $ds 'B3' (Override-Value $item 'Contact' (T $ds 'B3'))
  Put $ds 'J3' (Override-Value $item 'Item No.' (T $ds 'J3'))
  Put $ds 'B4' (Override-Value $item 'Service' (T $ds 'B4'))
  Put $ds 'C11' (Override-Value $item 'Media Hot' (T $ds 'C11'))
  Put $ds 'G11' (Override-Value $item 'Media Cold' (T $ds 'G11'))

  $heatValue = Override-Value $item 'Heat capacity' ''
  if (-not [string]::IsNullOrWhiteSpace([string]$heatValue)) { Put $ds 'E12' $heatValue }
  $heatUnit = Override-Value $item 'Heat capacity Unit' ''
  if (-not [string]::IsNullOrWhiteSpace([string]$heatUnit)) { Put $ds 'J12' (Unit-Text $heatUnit) }
  Convert-Cells $ds @('E12') 'J12' (Override-Value $item 'Heat capacity Target Unit' '') 'heat'

  $pdHot = Addr $isGphe 'C16' 'C17'
  $pdCold = Addr $isGphe 'G16' 'G17'
  $pdUnit = Addr $isGphe 'J16' 'J17'
  Put $ds $pdHot (Override-Value $item 'Pressure drop Hot' (T $ds $pdHot))
  Put $ds $pdCold (Override-Value $item 'Pressure drop Cold' (T $ds $pdCold))
  $pdUnitValue = Override-Value $item 'Pressure drop Unit' ''
  if (-not [string]::IsNullOrWhiteSpace([string]$pdUnitValue)) { Put $ds $pdUnit (Unit-Text $pdUnitValue) }
  Convert-Cells $ds @($pdHot, $pdCold) $pdUnit (Override-Value $item 'Pressure drop Target Unit' '') 'pressure'

  $opHot = Addr $isGphe 'C19' 'C18'
  $opCold = Addr $isGphe 'G19' 'G18'
  $opUnit = Addr $isGphe 'J19' 'J18'
  Put $ds $opHot (Override-Value $item 'Operating Pressure Hot' (T $ds $opHot))
  Put $ds $opCold (Override-Value $item 'Operating Pressure Cold' (T $ds $opCold))
  $opUnitValue = Override-Value $item 'Operating Pressure Unit' ''
  if (-not [string]::IsNullOrWhiteSpace([string]$opUnitValue)) { Put $ds $opUnit (Unit-Text $opUnitValue) }
  Convert-Cells $ds @($opHot, $opCold) $opUnit (Override-Value $item 'Operating Pressure Target Unit' '') 'pressure'

  $dt = Split-Pair (Override-Value $item 'Design Temperature' '')
  if (-not [string]::IsNullOrWhiteSpace($dt[0])) {
    Put $ds (Addr $isGphe 'D45' 'C40') $dt[0]
    if (-not [string]::IsNullOrWhiteSpace($dt[1])) { Put $ds (Addr $isGphe 'H45' 'G40') $dt[1] }
  }
  $dtUnit = Override-Value $item 'Design Temperature Unit' ''
  if (-not [string]::IsNullOrWhiteSpace([string]$dtUnit)) { Put $ds (Addr $isGphe 'J45' 'J40') (Unit-Text $dtUnit) }

  $dp = Split-Pair (Override-Value $item 'Design / Test Pressure' '')
  if (-not [string]::IsNullOrWhiteSpace($dp[0])) {
    Put $ds (Addr $isGphe 'E46' 'C41') $dp[0]
    if (-not [string]::IsNullOrWhiteSpace($dp[1])) { Put $ds (Addr $isGphe 'G46' 'G41') $dp[1] }
  }
  $designUnit = Override-Value $item 'Design / Test Pressure Unit' ''
  if (-not [string]::IsNullOrWhiteSpace([string]$designUnit)) { Put $ds (Addr $isGphe 'J46' 'J41') (Unit-Text $designUnit) }
  Convert-Cells $ds @((Addr $isGphe 'E46' 'C41'), (Addr $isGphe 'G46' 'G41')) (Addr $isGphe 'J46' 'J41') (Override-Value $item 'Design / Test Pressure Target Unit' '') 'pressure'
}

function Apply-HtriValues($target, $source, $item, [string]$kind) {
  $ds = Sheet-Like $target 'MC Datasheet|Datasheet'
  $api = Sheet-Like $source 'API\s*662'
  $out = Sheet-Like $source 'Output\s*Summary'
  $fin = Sheet-Like $source 'Final\s*Results'
  if ($null -eq $ds) { throw 'Template MC Datasheet sheet not found.' }
  if ($null -eq $api) { throw 'Source API 662 sheet not found.' }

  try { $ds.UsedRange.Value2 = $ds.UsedRange.Value2 } catch {}

  $base = [IO.Path]::GetFileNameWithoutExtension($item.name)
  $model = First-NonBlank @((Model-FromName $base), (T $api 'L10'), (T $fin 'L11'))
  $materialOverride = Override-Value $item 'Soldering Material' ''
  $model = Set-ModelPrefix $model $materialOverride
  $plates = First-NonBlank @((Plates-FromName $base), (T $api 'L38'))
  $isGphe = $kind -eq 'GPHE'

  Put $ds 'B1' (T $api 'D6')
  Put $ds 'B2' ''
  Put $ds 'B3' ''
  Put $ds 'J3' (T $api 'M9')
  Put $ds 'B4' (T $api 'D9')
  Put $ds 'J4' (Get-Date -Format 'yyyy-MM-dd')
  Put $ds 'C6' ("1 x $model - $plates Countercurrent Flow")
  Put $ds 'A8' '  Thermal data for 1 unit(s) in parallel and 1 unit(s) in series'

  Put $ds 'C11' (T $api 'L14')
  Put $ds 'G11' (T $api 'V14')
  Put $ds 'E12' (First-NumberText @((T $fin 'I30'), (T $out 'S26'), (T $out 'T26'), (T $api 'L40')))
  Put $ds 'D12' ''
  Put $ds 'H12' ''
  Put $ds 'J12' '(kcal/hr)'
  Put $ds 'C13' (T $api 'L16')
  Put $ds 'G13' (T $api 'V16')
  Put $ds 'J13' ('(' + (T $api 'I16') + ')')
  Put $ds 'C14' (First-NumberText @((T $out 'G14'), (T $out 'H14'), (T $fin 'I14'), (T $fin 'J14'), (T $api 'M22')))
  Put $ds 'C15' (First-NumberText @((T $out 'M14'), (T $out 'N14'), (T $fin 'M14'), (T $fin 'N14'), (T $api 'R22')))
  Put $ds 'G14' (First-NumberText @((T $out 'Q14'), (T $out 'R14'), (T $fin 'O14'), (T $fin 'P14'), (T $api 'W22')))
  Put $ds 'G15' (First-NumberText @((T $out 'S14'), (T $out 'T14'), (T $fin 'Q14'), (T $fin 'R14'), (T $api 'AB22')))
  Put $ds 'J14' ('(' + (T $api 'I22') + ')')
  Put $ds 'J15' ('(' + (T $api 'I22') + ')')

  $vapHotIn = First-NumberText @((T $out 'G15'), (T $out 'H15'), (T $fin 'I13'), (T $fin 'J13'))
  $vapHotOut = First-NumberText @((T $out 'M15'), (T $out 'N15'), (T $fin 'M13'), (T $fin 'N13'))
  $vapColdIn = First-NumberText @((T $out 'Q15'), (T $out 'R15'), (T $fin 'O13'), (T $fin 'P13'))
  $vapColdOut = First-NumberText @((T $out 'S15'), (T $out 'T15'), (T $fin 'Q13'), (T $fin 'R13'))
  Put $ds (Addr $isGphe 'C17' 'C16') $vapHotIn
  Put $ds (Addr $isGphe 'E17' 'E16') $vapHotOut
  Put $ds (Addr $isGphe 'G17' 'G16') $vapColdIn
  Put $ds (Addr $isGphe 'I17' 'I16') $vapColdOut

  if ($isGphe) {
    Put $ds 'C16' (First-NumberText @((T $out 'G18'), (T $out 'H18'), (T $fin 'I18'), (T $fin 'J18'), (T $api 'M34')))
    Put $ds 'G16' (First-NumberText @((T $out 'Q18'), (T $out 'R18'), (T $fin 'O18'), (T $fin 'P18'), (T $api 'W34')))
    Put $ds 'J16' ('(' + (T $api 'I34') + ')')
    Put $ds 'C19' (First-NumberText @((T $out 'G17'), (T $out 'H17'), (T $fin 'I17'), (T $fin 'J17'), (T $api 'M33')))
    Put $ds 'G19' (First-NumberText @((T $out 'Q17'), (T $out 'R17'), (T $fin 'O17'), (T $fin 'P17'), (T $api 'W33')))
    Put $ds 'J19' ('(' + (T $api 'I33') + ')')
  } else {
    Put $ds 'C17' (First-NumberText @((T $out 'G18'), (T $out 'H18'), (T $fin 'I18'), (T $fin 'J18'), (T $api 'R34')))
    Put $ds 'G17' (First-NumberText @((T $out 'Q18'), (T $out 'R18'), (T $fin 'O18'), (T $fin 'P18'), (T $api 'AB34')))
    Put $ds 'J17' ('(' + (T $api 'I34') + ')')
    Put $ds 'C18' (First-NumberText @((T $out 'G17'), (T $out 'H17'), (T $fin 'I17'), (T $fin 'J17'), (T $api 'M33')))
    Put $ds 'G18' (First-NumberText @((T $out 'Q17'), (T $out 'R17'), (T $fin 'O17'), (T $fin 'P17'), (T $api 'W33')))
    Put $ds 'J18' ('(' + (T $api 'I33') + ')')
  }

  $refHot = First-NumberText @((T $fin 'I15'), (T $fin 'J15'), (T $out 'G16'), (T $out 'H16'), (T $api 'M22'))
  $refCold = First-NumberText @((T $fin 'O15'), (T $fin 'P15'), (T $out 'Q16'), (T $out 'R16'), (T $api 'W22'))
  Put $ds (Addr $isGphe 'C23' 'C22') $refHot
  Put $ds (Addr $isGphe 'G23' 'G22') $refCold
  Put $ds (Addr $isGphe 'J23' 'J22') (T $api 'I22')

  PutDash $ds (Addr $isGphe 'C24' 'C23') (T $api 'M25')
  PutDash $ds (Addr $isGphe 'E24' 'E23') (T $api 'N25')
  PutDash $ds (Addr $isGphe 'G24' 'G23') (T $api 'V25')
  PutDash $ds (Addr $isGphe 'I24' 'I23') (T $api 'X25')
  Put $ds (Addr $isGphe 'J24' 'J23') (T $api 'I25')
  PutDash $ds (Addr $isGphe 'C25' 'C24') (T $api 'M26')
  PutDash $ds (Addr $isGphe 'E25' 'E24') (T $api 'N26')
  PutDash $ds (Addr $isGphe 'G25' 'G24') (T $api 'V26')
  PutDash $ds (Addr $isGphe 'I25' 'I24') (T $api 'X26')
  Put $ds (Addr $isGphe 'J25' 'J24') (T $api 'I26')
  PutDash $ds (Addr $isGphe 'C26' 'C25') (T $api 'M27')
  PutDash $ds (Addr $isGphe 'E26' 'E25') (T $api 'N27')
  PutDash $ds (Addr $isGphe 'G26' 'G25') (T $api 'V27')
  PutDash $ds (Addr $isGphe 'I26' 'I25') (T $api 'X27')
  Put $ds (Addr $isGphe 'J26' 'J25') (T $api 'I27')
  PutDash $ds (Addr $isGphe 'C27' 'C26') (T $api 'M28')
  PutDash $ds (Addr $isGphe 'E27' 'E26') (T $api 'N28')
  PutDash $ds (Addr $isGphe 'G27' 'G26') (T $api 'V28')
  PutDash $ds (Addr $isGphe 'I27' 'I26') (T $api 'X28')
  Put $ds (Addr $isGphe 'J27' 'J26') 'mN-s/m2'

  Put $ds (Addr $isGphe 'C33' 'C30') (First-NumberText @((T $api 'F11'), (T $out 'S27')))
  Put $ds (Addr $isGphe 'G33' 'G30') (First-NumberText @((T $api 'F11'), (T $out 'S27')))
  Put $ds (Addr $isGphe 'C34' 'C31') $plates
  Put $ds (Addr $isGphe 'G34' 'G31') $plates
  Put $ds (Addr $isGphe 'C36' 'C32') (First-NumberText @((T $fin 'I31'), (T $fin 'J31'), (T $out 'G28'), (T $out 'H28'), (T $api 'M41')))
  Put $ds (Addr $isGphe 'C37' 'C33') (First-NumberText @((T $out 'S24'), (T $out 'T24'), (T $api 'M42')))
  Put $ds (Addr $isGphe 'G37' 'G33') (First-NumberText @((T $out 'S25'), (T $out 'T25'), (T $api 'V42'), (T $fin 'I29'), (T $fin 'J29')))
  Put $ds (Addr $isGphe 'C38' 'C34') (First-NumberText @((T $out 'S28'), (T $out 'T28'), (T $api 'M43')))
  Put $ds (Addr $isGphe 'C39' 'C35') (First-NonBlank @((T $api 'L44'), 'AISI 316'))

  $parallel = First-Useful @((T $api 'N10'), '1')
  $series = First-Useful @((T $api 'Q10'), '1')
  $pNum = Num $parallel
  $sNum = Num $series
  $totalUnits = if ($null -ne $pNum -and $null -ne $sNum) { [int]($pNum * $sNum) } else { '1' }

  $hotPressure = First-NumberText @((T $out 'G17'), (T $out 'H17'), (T $fin 'I17'), (T $fin 'J17'), (T $api 'M33'))
  $coldPressure = First-NumberText @((T $out 'Q17'), (T $out 'R17'), (T $fin 'O17'), (T $fin 'P17'), (T $api 'W33'))
  $pressureBase = Max-Number @($hotPressure, $coldPressure) 10.0
  if ($pressureBase -le 10.0) {
    $designPressure = 10
  } else {
    $designPressure = [Math]::Ceiling([Math]::Max($pressureBase + 2.0, $pressureBase * 1.1))
  }
  $testPressure = [Math]::Round($designPressure * 1.3, 1)
  $minTemp = [Math]::Floor((Min-Number @((T $out 'G14'), (T $out 'H14'), (T $out 'M14'), (T $out 'N14'), (T $out 'Q14'), (T $out 'R14'), (T $out 'S14'), (T $out 'T14'), (T $fin 'I14'), (T $fin 'J14'), (T $fin 'M14'), (T $fin 'N14'), (T $fin 'O14'), (T $fin 'P14'), (T $fin 'Q14'), (T $fin 'R14'), (T $api 'M22'), (T $api 'R22'), (T $api 'W22'), (T $api 'AB22')) 0.0))
  $maxTemp = [Math]::Ceiling((Max-Number @((T $out 'G14'), (T $out 'H14'), (T $out 'M14'), (T $out 'N14'), (T $out 'Q14'), (T $out 'R14'), (T $out 'S14'), (T $out 'T14'), (T $fin 'I14'), (T $fin 'J14'), (T $fin 'M14'), (T $fin 'N14'), (T $fin 'O14'), (T $fin 'P14'), (T $fin 'Q14'), (T $fin 'R14'), (T $api 'M22'), (T $api 'R22'), (T $api 'W22'), (T $api 'AB22')) 100.0))

  if ($isGphe) {
    Put $ds 'C31' $parallel
    Put $ds 'F31' $series
    Put $ds 'G31' $totalUnits
    Put $ds 'C32' $model
    Put $ds 'F35' (T $api 'L39')
    Put $ds 'C41' ((T $api 'N37') + ' x ' + (T $api 'Q37'))
    Put $ds 'G41' ((T $api 'X37') + ' x ' + (Add-One (T $api 'AA37')))
    Put $ds 'C42' ((T $api 'I49') + ' ' + (T $api 'J49'))
    Put $ds 'G42' ((T $api 'S49') + ' ' + (T $api 'T49'))
    Put $ds 'D45' $minTemp
    Put $ds 'F45' '/'
    Put $ds 'H45' $maxTemp
    Put $ds 'J45' '(Deg C)'
    Put $ds 'E46' $designPressure
    Put $ds 'F46' '/'
    Put $ds 'G46' $testPressure
    Put $ds 'J46' '(barG)'
  } else {
    $solder = Display-Solder $materialOverride $model
    Put $ds 'C36' $solder
    Put $ds 'C37' (T $api 'N37')
    Put $ds 'E37' (T $api 'Q37')
    Put $ds 'G37' (T $api 'X37')
    Put $ds 'I37' (Add-One (T $api 'AA37'))
    Put $ds 'C38' $parallel
    Put $ds 'F38' $series
    Put $ds 'G38' $totalUnits
    Put $ds 'C39' 'Stainless steel'
    Put $ds 'C40' $minTemp
    Put $ds 'F40' '/'
    Put $ds 'G40' $maxTemp
    Put $ds 'J40' '(Deg C)'
    Put $ds 'C41' $designPressure
    Put $ds 'F41' '/'
    Put $ds 'G41' $testPressure
    Put $ds 'J41' '(barG)'
  }
  Apply-HtriOverrides $ds $item $isGphe
  Clear-KnownSlashCells $ds $isGphe
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  try { $excel.AutomationSecurity = 3 } catch {}
  $outputs = @()
  foreach ($item in $manifest.files) {
    $kind = if ($item.kind -eq 'GPHE') { 'GPHE' } else { 'BPHE' }
    $template = Join-Path $TemplateDir ($kind + '.xlsm')
    if (!(Test-Path -LiteralPath $template)) { throw "Template not found: $template" }
    if (!(Test-Path -LiteralPath $item.path)) { throw "HTRI file not found: $($item.path)" }

    $sourceBase = [IO.Path]::GetFileNameWithoutExtension($item.name)
    $modelForName = Safe-Part (Model-FromName $sourceBase) 'HTRI'
    $modelForName = Safe-Part (Set-ModelPrefix $modelForName (Override-Value $item 'Soldering Material' '')) $modelForName
    $platesForName = Plates-FromName $sourceBase
    if ([string]::IsNullOrWhiteSpace($platesForName)) {
      $fileBase = $modelForName + '_Datasheet'
    } else {
      $fileBase = $modelForName + '_' + $platesForName + 'pl_Datasheet'
    }
    $outPath = Unique-Path (Join-Path $OutputDir ($fileBase + '.xlsm'))
    Copy-Item -LiteralPath $template -Destination $outPath -Force

    $target = $excel.Workbooks.Open($outPath)
    $source = $excel.Workbooks.Open($item.path)
    try {
      Apply-HtriValues $target $source $item $kind
      $excel.CalculateFullRebuild()
      $target.SaveAs($outPath, 52)
    } finally {
      $source.Close($false)
      $target.Close($true)
    }
    $outputs += $outPath
  }

  if ($outputs.Count -eq 1) {
    [pscustomobject]@{ type='file'; path=$outputs[0] } | ConvertTo-Json -Compress
  } else {
    $zipPath = Join-Path $OutputDir 'HTRI_Datasheets.zip'
    if (Test-Path -LiteralPath $zipPath) { Remove-Item -LiteralPath $zipPath -Force }
    Compress-Archive -LiteralPath $outputs -DestinationPath $zipPath -Force
    [pscustomobject]@{ type='zip'; path=$zipPath } | ConvertTo-Json -Compress
  }
} finally {
  $excel.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
