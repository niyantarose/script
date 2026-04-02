param(
  [string]$InputCsv = 'C:\Users\Owner\Downloads\quantity_name202603161403.csv',
  [string]$OutputCsv = 'C:\Users\Owner\Downloads\magazine_master_candidates.csv'
)

Add-Type -AssemblyName Microsoft.VisualBasic
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Read-CsvShiftJis {
  param([string]$Path)

  $parser = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($Path, [System.Text.Encoding]::GetEncoding(932))
  $parser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
  $parser.SetDelimiters(',')
  $parser.HasFieldsEnclosedInQuotes = $true

  $headers = $parser.ReadFields()
  $rows = @()
  while (-not $parser.EndOfData) {
    $fields = $parser.ReadFields()
    $row = [ordered]@{}
    for ($i = 0; $i -lt $headers.Length; $i++) {
      $name = $headers[$i]
      $row[$name] = if ($i -lt $fields.Length) { $fields[$i] } else { '' }
    }
    $rows += [pscustomobject]$row
  }
  $parser.Close()
  return $rows
}

function Normalize-MagazineText {
  param([string]$Text)

  return ([string]$Text).Normalize('FormKC').ToUpper() `
    -replace '[（(].*?[)）]', ' ' `
    -replace '&', ' AND ' `
    -replace '\b(KOREA|TAIWAN|HONG\s*KONG|HONGKONG|CHINA|THAILAND|HK|TW|CN)\b', ' ' `
    -replace '(韓国版|韓國版|台湾版|臺灣版|中国版|中國版|香港版|タイ版)', ' ' `
    -replace '[^A-Z0-9]+', ' ' `
    -replace '\s+', ' '
}

function Get-LanguageCode {
  param([string]$Text)
  $upper = [string]$Text
  if ($upper -match '台湾|臺灣|TAIWAN') { return 'TW' }
  if ($upper -match '香港|HONG\s*KONG|HONGKONG') { return 'HK' }
  if ($upper -match '中国|中國|CHINA') { return 'CN' }
  if ($upper -match 'タイ|THAILAND|THAI|泰国|泰國') { return 'TH' }
  if ($upper -match 'KOREA|韓国|韓國') { return 'KR' }
  return ''
}

function Get-IssueInfo {
  param([string]$Text)
  $value = [string]$Text
  if ($value -match 'Vol\.?\s*\d+' -or $value -match '第\s*\d+\s*(號|号|期)') {
    return [pscustomobject]@{ KeyType = '号数型'; HasMarker = $true; Reason = 'issue_marker' }
  }
  if (
    $value -match '\d{1,2}(?:\s*[./\-・~～]\s*\d{1,2})?\s*月(?:號|号)?\s*/\s*20\d{2}' -or
    $value -match '20\d{2}\s*(年|/|\-|\.)\s*\d{1,2}(?:\s*[./\-・~～]\s*\d{1,2})?\s*月'
  ) {
    return [pscustomobject]@{ KeyType = '年月型'; HasMarker = $true; Reason = 'year_month_marker' }
  }
  return [pscustomobject]@{ KeyType = '年月型'; HasMarker = $false; Reason = '' }
}

function Get-Abbreviation {
  param([string]$Name)
  $cleaned = (Normalize-MagazineText $Name) -replace '\s+', ''
  if (-not $cleaned) { return '' }
  if ($cleaned.Length -le 8) { return $cleaned }
  return $cleaned.Substring(0, 8)
}

$script:KnownMagazineDefinitions = $null
function Get-KnownMagazineDefinitions {
  if ($script:KnownMagazineDefinitions) { return $script:KnownMagazineDefinitions }
  $script:KnownMagazineDefinitions = @(
    @{ Canonical = 'W KOREA'; Language = 'KR'; Aliases = @('W KOREA', 'W') },
    @{ Canonical = 'VOGUE'; Language = ''; Aliases = @('VOGUE') },
    @{ Canonical = 'VOGUE KOREA'; Language = 'KR'; Aliases = @('VOGUE KOREA') },
    @{ Canonical = 'VOGUE TAIWAN'; Language = 'TW'; Aliases = @('VOGUE TAIWAN') },
    @{ Canonical = 'ELLE'; Language = ''; Aliases = @('ELLE') },
    @{ Canonical = 'ELLE KOREA'; Language = 'KR'; Aliases = @('ELLE KOREA') },
    @{ Canonical = 'ELLE HONG KONG'; Language = 'HK'; Aliases = @('ELLE HONG KONG') },
    @{ Canonical = 'BAZAAR'; Language = ''; Aliases = @('BAZAAR', "HARPER'S BAZAAR") },
    @{ Canonical = 'BAZAAR KOREA'; Language = 'KR'; Aliases = @('BAZAAR KOREA', "HARPER'S BAZAAR KOREA") },
    @{ Canonical = 'BAZAAR TAIWAN'; Language = 'TW'; Aliases = @('BAZAAR TAIWAN', "HARPER'S BAZAAR TAIWAN") },
    @{ Canonical = 'GQ'; Language = ''; Aliases = @('GQ') },
    @{ Canonical = 'GQ KOREA'; Language = 'KR'; Aliases = @('GQ KOREA') },
    @{ Canonical = 'marie claire'; Language = ''; Aliases = @('MARIE CLAIRE', 'marie claire', 'Marie Claire') },
    @{ Canonical = 'marie claire KOREA'; Language = 'KR'; Aliases = @('MARIE CLAIRE KOREA', 'marie claire Korea') },
    @{ Canonical = 'COSMOPOLITAN'; Language = ''; Aliases = @('COSMOPOLITAN') },
    @{ Canonical = 'COSMOPOLITAN KOREA'; Language = 'KR'; Aliases = @('COSMOPOLITAN KOREA') },
    @{ Canonical = 'MAXIM'; Language = ''; Aliases = @('MAXIM') },
    @{ Canonical = 'MAXIM KOREA'; Language = 'KR'; Aliases = @('MAXIM KOREA') },
    @{ Canonical = 'Esquire'; Language = ''; Aliases = @('ESQUIRE', '時尚先生 ESQUIRE') },
    @{ Canonical = 'Esquire Korea'; Language = 'KR'; Aliases = @('ESQUIRE KOREA') },
    @{ Canonical = 'Esquire Hong Kong'; Language = 'HK'; Aliases = @('ESQUIRE HONG KONG', 'ESQUIRE HK') },
    @{ Canonical = 'DAZED'; Language = ''; Aliases = @('DAZED') },
    @{ Canonical = 'DAZED KOREA'; Language = 'KR'; Aliases = @('DAZED KOREA', 'DAZED&CONFUSED KOREA') },
    @{ Canonical = 'allure'; Language = ''; Aliases = @('ALLURE', 'allure') },
    @{ Canonical = 'allure KOREA'; Language = 'KR'; Aliases = @('ALLURE KOREA', 'allure Korea') },
    @{ Canonical = '1st LOOK'; Language = 'KR'; Aliases = @('1ST LOOK', '1STLOOK') },
    @{ Canonical = 'CINE21'; Language = 'KR'; Aliases = @('CINE21', 'CINE 21') },
    @{ Canonical = 'Singles'; Language = 'KR'; Aliases = @('SINGLES', 'Singles') },
    @{ Canonical = 'THE STAR'; Language = 'KR'; Aliases = @('THE STAR', 'STAR') },
    @{ Canonical = 'NYLON KOREA'; Language = 'KR'; Aliases = @('NYLON KOREA', 'NYLON') },
    @{ Canonical = "Men's Health"; Language = 'KR'; Aliases = @("MEN'S HEALTH", 'MENS HEALTH') },
    @{ Canonical = "L'OFFICIEL KOREA"; Language = 'KR'; Aliases = @("L'OFFICIEL", 'LOFFICIEL', 'L OFFICIEL') },
    @{ Canonical = 'ARENA HOMME+ KOREA'; Language = 'KR'; Aliases = @('ARENA HOMME+', 'ARENA HOMME', 'ARENA') },
    @{ Canonical = 'WWD Korea'; Language = 'KR'; Aliases = @('WWD KOREA', 'WWD') },
    @{ Canonical = 'GRAZIA KOREA'; Language = 'KR'; Aliases = @('GRAZIA KOREA', 'GRAZIA') },
    @{ Canonical = 'MAPS'; Language = 'KR'; Aliases = @('MAPS') },
    @{ Canonical = 'arte'; Language = 'KR'; Aliases = @('ARTE') },
    @{ Canonical = 'scenePLAYBILL'; Language = 'KR'; Aliases = @('SCENEPLAYBILL') },
    @{ Canonical = 'CAMPUS PLUS'; Language = 'KR'; Aliases = @('CAMPUS PLUS') },
    @{ Canonical = 'TROTZINE'; Language = 'KR'; Aliases = @('TROTZINE') },
    @{ Canonical = 'Praew'; Language = 'TH'; Aliases = @('PRAEW') },
    @{ Canonical = 'Sudsapda'; Language = 'TH'; Aliases = @('SUDSAPDA') },
    @{ Canonical = "MEN'S UNO HK"; Language = 'HK'; Aliases = @("MEN'S UNO HK", "MEN'S UNO", 'S UNO HK') },
    @{ Canonical = 'STAR FOCUS'; Language = 'KR'; Aliases = @('STAR FOCUS') },
    @{ Canonical = 'THE MUSICAL'; Language = 'KR'; Aliases = @('THE MUSICAL') },
    @{ Canonical = 'Perla China'; Language = 'CN'; Aliases = @('PERLA CHINA') },
    @{ Canonical = 'SENSE China'; Language = 'CN'; Aliases = @('SENSE CHINA') },
    @{ Canonical = 'DE DELING'; Language = 'CN'; Aliases = @('DE DELING') },
    @{ Canonical = 'SPOTLiGHT China'; Language = 'CN'; Aliases = @('SPOTLIGHT CHINA') },
    @{ Canonical = 'CITER'; Language = 'CN'; Aliases = @('CITER') }
  )
  return $script:KnownMagazineDefinitions
}

function Test-ShortAliasBoundary {
  param([string]$Text, [string]$Alias)
  $escaped = [regex]::Escape(([string]$Alias).ToUpper())
  $regex = "(^|[^A-Z0-9])$escaped([^A-Z0-9]|$)"
  return ([string]$Text).ToUpper() -match $regex
}

function Get-KnownMagazineMatch {
  param([string]$Text, [string]$LanguageCode)

  $raw = [string]$Text
  $rawNorm = Normalize-MagazineText $raw
  if (-not $rawNorm) { return $null }

  $best = $null
  $bestScore = -1
  foreach ($def in (Get-KnownMagazineDefinitions)) {
    $langScore = if ($LanguageCode -and $def.Language -and $LanguageCode -eq $def.Language) { 100 } elseif (-not $def.Language) { 30 } else { 0 }
    foreach ($alias in $def.Aliases) {
      $aliasNorm = Normalize-MagazineText $alias
      if (-not $aliasNorm) { continue }

      $score = -1
      if ($rawNorm -eq $aliasNorm) { $score = 260 }
      elseif ($rawNorm.StartsWith($aliasNorm)) { $score = 210 + $aliasNorm.Length }
      elseif ($rawNorm.Contains($aliasNorm)) { $score = 180 + $aliasNorm.Length }
      elseif ($aliasNorm.Length -le 2 -and (Test-ShortAliasBoundary -Text $raw -Alias $alias)) { $score = 190 }

      if ($score -lt 0) { continue }
      $score += $langScore

      if ($score -gt $bestScore) {
        $bestScore = $score
        $best = [pscustomobject]@{
          Canonical = $def.Canonical
          Language = if ($LanguageCode) { $LanguageCode } elseif ($def.Language) { $def.Language } else { '' }
          Score = $score
          Reason = "known:$($alias)"
        }
      }
    }
  }

  return $best
}

function Get-CandidateStem {
  param([string]$Text)

  $value = ([string]$Text) -replace '[　]', ' '
  $value = $value -replace '\s+', ' '
  $match = [regex]::Match($value, '(?=\s*(?:20\d{2}\s*(?:年|/|\-|\.)\s*\d{1,2}\s*月|\d{1,2}(?:\s*[./\-・~～]\s*\d{1,2})?\s*月(?:號|号)?\s*/\s*20\d{2}|Vol\.?\s*\d+|第\s*\d+\s*(?:號|号|期)))', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
  if ($match.Success -and $match.Index -gt 0) {
    $value = $value.Substring(0, $match.Index)
  }

  $prefixPatterns = @(
    '^[★☆◎◯○●\s\d%％OFFoff+\-]*',
    '^(韓国|韓國|中国|中國|台湾|臺灣|香港|タイ)(?:語)?\s*(芸能|女性|男性)?\s*雑誌\s*',
    '^(韓国|韓國|中国|中國|台湾|臺灣|香港|タイ)\s+',
    '^雑誌\s*',
    '^まんが[『「"]?'
  )
  foreach ($pattern in $prefixPatterns) {
    $value = $value -replace $pattern, ''
  }

  $value = $value -replace '[\[【（(].*?[\]】）)]', ' '
  $value = $value -replace '^[『「\"\'']+', ''
  $value = $value -replace '\s+', ' '
  $value = $value.Trim(' ', '-', '/', ':', '・', '／', '』', '」')
  return $value
}

function Test-ExcludedRow {
  param([string]$Text)

  $patterns = @(
    'コミック', '漫畫', '漫画', '小説', '小說', '畫集', '画集', '写真集', '寫真集',
    'フォトブック', 'PHOTOBOOK', 'ブルーレイ', 'BLU-RAY', 'DVD', 'CD', 'OST',
    'アクリル', '缶バッジ', 'キーホルダー', 'トレカ', 'ポスター', 'カレンダー',
    'クリアファイル', '周邊', 'グッズ', '透明立方', 'ぬい', 'タペストリー'
  )

  foreach ($pattern in $patterns) {
    if ([string]$Text -match $pattern) { return $true }
  }
  return $false
}

function Test-GenericCandidate {
  param([string]$Name)

  $norm = (Normalize-MagazineText $Name)
  $generic = @('STYLE', 'MAGAZINE', 'THE', 'VOL', 'ISSUE', 'SPECIAL EDITION', 'LIMITED EDITION')
  return $generic -contains $norm
}

function Build-CandidateRecord {
  param(
    [string]$ProductCode,
    [string]$ProductName
  )

  $language = Get-LanguageCode -Text $ProductName
  $issueInfo = Get-IssueInfo -Text $ProductName
  $known = Get-KnownMagazineMatch -Text $ProductName -LanguageCode $language
  $excluded = Test-ExcludedRow -Text $ProductName

  $candidateName = ''
  $score = 0
  $reason = @()

  if ($known) {
    $candidateName = $known.Canonical
    $language = if ($language) { $language } else { $known.Language }
    $score = $known.Score
    $reason += $known.Reason
  } else {
    if (-not $issueInfo.HasMarker) { return $null }

    $candidateName = Get-CandidateStem -Text $ProductName
    if (-not $candidateName) { return $null }
    if (Test-GenericCandidate -Name $candidateName) { return $null }

    $score = 40
    if ($issueInfo.HasMarker) { $score += 30 }
    if ($candidateName.Length -ge 3) { $score += 10 }
    $reason += 'stem'
    $reason += $issueInfo.Reason
  }

  if ($excluded -and -not $known) {
    $score -= 80
    $reason += 'excluded_keyword'
  }

  return [pscustomobject]@{
    対応言語 = $language
    雑誌名候補 = $candidateName
    略称コード候補 = Get-Abbreviation -Name $candidateName
    基本キー型候補 = $issueInfo.KeyType
    サンプル商品コード = [string]$ProductCode
    サンプル商品名 = [string]$ProductName
    原題タイトルサンプル = [string]$ProductName
    出現数 = 1
    信頼度 = $score
    備考 = ($reason -join ',')
  }
}

$candidates = @{}
$parser = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($InputCsv, [System.Text.Encoding]::GetEncoding(932))
$parser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
$parser.SetDelimiters(',')
$parser.HasFieldsEnclosedInQuotes = $true

$headers = $parser.ReadFields()
$codeIndex = [Array]::IndexOf($headers, 'code')
$nameIndex = [Array]::IndexOf($headers, 'name')

while (-not $parser.EndOfData) {
  $fields = $parser.ReadFields()
  $productCode = if ($codeIndex -ge 0 -and $codeIndex -lt $fields.Length) { [string]$fields[$codeIndex] } else { '' }
  $productName = if ($nameIndex -ge 0 -and $nameIndex -lt $fields.Length) { [string]$fields[$nameIndex] } else { '' }
  if (-not $productName) { continue }

  $candidate = Build-CandidateRecord -ProductCode $productCode -ProductName $productName
  if (-not $candidate) { continue }

  $key = (([string]$candidate.対応言語).ToUpper() + '|' + (Normalize-MagazineText $candidate.雑誌名候補))
  if (-not $candidates.ContainsKey($key)) {
    $candidate | Add-Member -NotePropertyName 登録日時 -NotePropertyValue (Get-Date)
    $candidate | Add-Member -NotePropertyName 博客來URL -NotePropertyValue ''
    $candidate | Add-Member -NotePropertyName 重複キー -NotePropertyValue $key
    $candidate | Add-Member -NotePropertyName 状態 -NotePropertyValue ''
    $candidates[$key] = $candidate
    continue
  }

  $current = $candidates[$key]
  $current.出現数 = [int]$current.出現数 + 1
  if ([int]$candidate.信頼度 -gt [int]$current.信頼度) {
    $current.信頼度 = $candidate.信頼度
    $current.サンプル商品コード = $candidate.サンプル商品コード
    $current.サンプル商品名 = $candidate.サンプル商品名
    $current.原題タイトルサンプル = $candidate.原題タイトルサンプル
    $current.備考 = $candidate.備考
  }
}
$parser.Close()

$result = foreach ($item in $candidates.Values) {
  $score = [int]$item.信頼度
  if ([int]$item.出現数 -ge 3) { $score += 20 }
  if ([int]$item.出現数 -ge 10) { $score += 20 }
  if ([int]$item.出現数 -ge 50) { $score += 20 }

  $status = '要確認'
  if ($item.備考 -match 'excluded_keyword') {
    $status = '除外候補'
  } elseif ($item.備考 -match 'known:' -or $score -ge 120) {
    $status = '自動候補'
  } elseif ($score -lt 70) {
    $status = '除外候補'
  }

  [pscustomobject]@{
    登録日時 = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    対応言語 = $item.対応言語
    雑誌名候補 = $item.雑誌名候補
    略称コード候補 = $item.略称コード候補
    基本キー型候補 = $item.基本キー型候補
    サンプル商品コード = $item.サンプル商品コード
    サンプル商品名 = $item.サンプル商品名
    博客來URL = ''
    原題タイトルサンプル = $item.原題タイトルサンプル
    出現数 = $item.出現数
    信頼度 = $score
    重複キー = $item.重複キー
    状態 = $status
    備考 = $item.備考
  }
}

$result |
  Sort-Object 状態, @{ Expression = { - [int]$_.出現数 } }, 対応言語, 雑誌名候補 |
  Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8

Write-Host "saved: $OutputCsv"
Write-Host "count: $($result.Count)"
