$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$templatePath = Join-Path $root 'Halal.html'
$docxPath = Join-Path $root 'Rack City Website content 15.05.2026.docx'
$utf8 = New-Object System.Text.UTF8Encoding($false)
$assetDir = Join-Path $root 'assets'
$cssDir = Join-Path $assetDir 'css'
$jsDir = Join-Path $assetDir 'js'
$siteCssPath = Join-Path $cssDir 'site.css'
$siteJsPath = Join-Path $jsDir 'site.js'
$cssBlocks = New-Object System.Collections.Generic.List[string]
$cssBlockSet = @{}
$scriptBlocks = New-Object System.Collections.Generic.List[string]
$scriptBlockSet = @{}
$styleClassMap = @{}
$utilityRules = New-Object System.Collections.Generic.List[string]

function HtmlText($value) {
  return [System.Security.SecurityElement]::Escape([string]$value)
}

function Replace-First($text, $pattern, $replacement) {
  $regex = New-Object System.Text.RegularExpressions.Regex($pattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)
  return $regex.Replace($text, $replacement, 1)
}

function Add-CssBlock($css) {
  $clean = ([string]$css).Trim()
  if ($clean.Length -eq 0) { return }
  if (-not $cssBlockSet.ContainsKey($clean)) {
    $cssBlockSet[$clean] = $true
    $cssBlocks.Add($clean) | Out-Null
  }
}

function Add-ScriptBlock($script) {
  $clean = ([string]$script).Trim()
  if ($clean.Length -eq 0) { return }
  if (-not $scriptBlockSet.ContainsKey($clean)) {
    $scriptBlockSet[$clean] = $true
    $scriptBlocks.Add($clean) | Out-Null
  }
}

function Get-StyleClass($style) {
  $clean = ([System.Net.WebUtility]::HtmlDecode([string]$style)).Trim()
  if ($clean.Length -eq 0) { return '' }
  if (-not $clean.EndsWith(';')) { $clean += ';' }
  if ($styleClassMap.ContainsKey($clean)) { return $styleClassMap[$clean] }
  $className = 'u-style-' + ($styleClassMap.Count + 1)
  $styleClassMap[$clean] = $className
  $utilityRules.Add(".$className { $clean }") | Out-Null
  return $className
}

function Add-ClassToTag($tag, $className) {
  if ([string]::IsNullOrWhiteSpace($className)) { return $tag }
  $classRegex = New-Object System.Text.RegularExpressions.Regex('class="([^"]*)"', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
  if ($classRegex.IsMatch($tag)) {
    return $classRegex.Replace($tag, { param($m) 'class="' + $m.Groups[1].Value + ' ' + $className + '"' }, 1)
  }
  return $tag -replace '^<([^\s>/]+)', ('<$1 class="' + $className + '"')
}

function Externalize-Assets($html) {
  $result = [string]$html

  $styleRegex = New-Object System.Text.RegularExpressions.Regex('<style\b[^>]*>([\s\S]*?)</style>', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
  $result = $styleRegex.Replace($result, {
    param($m)
    Add-CssBlock $m.Groups[1].Value
    ''
  })

  $scriptRegex = New-Object System.Text.RegularExpressions.Regex('<script(?![^>]*type="application/ld\+json")[^>]*>([\s\S]*?)</script>', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
  $result = $scriptRegex.Replace($result, {
    param($m)
    Add-ScriptBlock $m.Groups[1].Value
    ''
  })

  $styleAttrRegex = New-Object System.Text.RegularExpressions.Regex('\sstyle="([^"]*)"', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
  $tagRegex = New-Object System.Text.RegularExpressions.Regex('<[a-zA-Z][^>]*\sstyle="[^"]*"[^>]*>', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
  $result = $tagRegex.Replace($result, {
    param($m)
    $styleMatch = $styleAttrRegex.Match($m.Value)
    $className = Get-StyleClass $styleMatch.Groups[1].Value
    $tagWithoutStyle = $styleAttrRegex.Replace($m.Value, '', 1)
    Add-ClassToTag $tagWithoutStyle $className
  })

  $inlineEventRegex = New-Object System.Text.RegularExpressions.Regex('\son[a-zA-Z]+="[^"]*"', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
  $result = $inlineEventRegex.Replace($result, '')

  if ($result -notmatch '<link rel="stylesheet" href="/assets/css/site\.css">') {
    $result = $result -replace '</head>', "  <link rel=`"stylesheet`" href=`"/assets/css/site.css`">`n</head>"
  }
  if ($scriptBlocks.Count -gt 0 -and $result -notmatch '<script src="/assets/js/site\.js" defer></script>') {
    $result = $result -replace '</body>', "  <script src=`"/assets/js/site.js`" defer></script>`n</body>"
  }

  return $result
}

function Slice($text, $startMarker, $endMarker) {
  $start = $text.IndexOf($startMarker)
  if ($start -lt 0) { throw "Missing marker $startMarker" }
  $end = $text.IndexOf($endMarker, $start)
  if ($end -lt 0) { throw "Missing marker $endMarker" }
  return $text.Substring($start, $end - $start)
}

function SliceInclusiveFrom($text, $startMarker) {
  $start = $text.IndexOf($startMarker)
  if ($start -lt 0) { throw "Missing marker $startMarker" }
  return $text.Substring($start)
}

function Read-DocxParagraphs($path) {
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $resolvedPath = (Resolve-Path $path).Path
  $sourcePath = $resolvedPath
  $tempPath = $null
  try {
    $tempPath = Join-Path ([System.IO.Path]::GetTempPath()) ("rack-city-content-" + [guid]::NewGuid().ToString() + ".docx")
    Copy-Item -LiteralPath $resolvedPath -Destination $tempPath -Force
    $sourcePath = $tempPath
  } catch {
    $sourcePath = $resolvedPath
  }

  $zip = [System.IO.Compression.ZipFile]::OpenRead($sourcePath)
  try {
    $entry = $zip.GetEntry('word/document.xml')
    $reader = New-Object System.IO.StreamReader($entry.Open())
    try {
      [xml]$doc = $reader.ReadToEnd()
    } finally {
      $reader.Close()
    }
  } finally {
    $zip.Dispose()
    if ($tempPath -and (Test-Path $tempPath)) {
      Remove-Item -LiteralPath $tempPath -Force
    }
  }

  $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
  $ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
  $paras = $doc.SelectNodes('//w:p', $ns)
  $items = @()
  $i = 0
  foreach ($p in $paras) {
    $texts = $p.SelectNodes('.//w:t', $ns)
    $s = ($texts | ForEach-Object { $_.'#text' }) -join ''
    if ($s.Trim().Length -gt 0) {
      $items += [pscustomobject]@{ Index = $i; Text = $s.Trim() }
      $i++
    }
  }
  return $items
}

function Parse-Meta($line) {
  $pattern = 'Slug:\s*(?<slug>.*?)(?=Focus keyphrase:)Focus keyphrase:\s*(?<focus>.*?)(?=SEO title:)SEO title:\s*(?<title>.*?)(?=Meta description:)Meta description:\s*(?<description>.*?)(?=Canonical:)Canonical:\s*(?<canonical>\S+)'
  $m = [regex]::Match($line, $pattern)
  if (-not $m.Success) {
    throw "Could not parse page metadata: $line"
  }
  return [ordered]@{
    slug = $m.Groups['slug'].Value.Trim()
    focus = $m.Groups['focus'].Value.Trim()
    title = $m.Groups['title'].Value.Trim()
    description = $m.Groups['description'].Value.Trim()
    canonical = $m.Groups['canonical'].Value.Trim()
  }
}

function QuestionAnswer($line) {
  $qPos = $line.IndexOf('?')
  if ($qPos -lt 1) { return $null }
  $q = $line.Substring(0, $qPos + 1).Trim()
  $a = $line.Substring($qPos + 1).Trim()
  if ($q.Length -lt 4 -or $a.Length -lt 1) { return $null }
  return [ordered]@{ q = $q; a = $a }
}

function Parse-Faqs($copyLines, $schemaLines) {
  $faqs = @()
  foreach ($line in $copyLines) {
    $qa = QuestionAnswer $line
    if ($qa -ne $null) {
      $q = CleanPublicText $qa.q
      $a = CleanPublicText $qa.a
      if ($q.Length -gt 0 -and $a.Length -gt 0) {
        $faqs += [pscustomobject]@{ q = $q; a = $a }
      }
    }
  }

  if ($faqs.Count -gt 0) { return $faqs }

  $schemaText = $schemaLines -join "`n"
  $qMatches = [regex]::Matches($schemaText, '"name"\s*:\s*"([^"]+)"[\s\S]*?"text"\s*:\s*"([^"]+)"')
  foreach ($m in $qMatches) {
    $faqs += [ordered]@{
      q = $m.Groups[1].Value
      a = $m.Groups[2].Value
    }
  }
  return $faqs
}

function JsonScript($obj) {
  $json = $obj | ConvertTo-Json -Depth 12
  return "<script type=`"application/ld+json`">`n$json`n</script>"
}

function BreadcrumbScript($pageName, $canonical, $isHome) {
  $items = @(
    [ordered]@{
      '@type' = 'ListItem'
      position = 1
      name = 'Home'
      item = 'https://www.rackcitykitchen.com/'
    }
  )
  if (-not $isHome) {
    $items += [ordered]@{
      '@type' = 'ListItem'
      position = 2
      name = $pageName
      item = $canonical
    }
  }
  return JsonScript ([ordered]@{
    '@context' = 'https://schema.org'
    '@type' = 'BreadcrumbList'
    itemListElement = $items
  })
}

function FaqScript($faqs) {
  $entities = @()
  foreach ($faq in $faqs) {
    $entities += [ordered]@{
      '@type' = 'Question'
      name = $faq.q
      acceptedAnswer = [ordered]@{
        '@type' = 'Answer'
        text = $faq.a
      }
    }
  }
  return JsonScript ([ordered]@{
    '@context' = 'https://schema.org'
    '@type' = 'FAQPage'
    mainEntity = $entities
  })
}

function Extract-JsonScripts($lines) {
  $scripts = @()
  $collecting = $false
  $current = @()
  foreach ($line in $lines) {
    if ($line -eq '<script type="application/ld+json">') {
      $collecting = $true
      $current = @($line)
      continue
    }
    if ($collecting) {
      $current += $line
      if ($line -eq '</script>') {
        $scripts += ,($current -join "`n")
        $collecting = $false
        $current = @()
      }
    }
  }
  return $scripts
}

function ApplyHead($html, $page, $restaurantSchema, $breadcrumbSchema, $faqSchema) {
  $title = HtmlText $page.Title
  $description = HtmlText $page.Description
  $canonical = HtmlText $page.Canonical

  $html = Replace-First $html '<title>.*?</title>' "<title>$title</title>"
  $html = Replace-First $html '<meta name="description" content=".*?">' "<meta name=`"description`" content=`"$description`">"
  $html = Replace-First $html '<link rel="canonical" href=".*?">' "<link rel=`"canonical`" href=`"$canonical`">"
  $html = Replace-First $html '<meta property="og:title" content=".*?">' "<meta property=`"og:title`" content=`"$title`">"
  $html = Replace-First $html '<meta property="og:description" content=".*?">' "<meta property=`"og:description`" content=`"$description`">"
  $html = Replace-First $html '<meta property="og:url" content=".*?">' "<meta property=`"og:url`" content=`"$canonical`">"

  $schemaMatches = [System.Text.RegularExpressions.Regex]::Matches($html, '<script type="application/ld\+json">[\s\S]*?</script>')
  if ($schemaMatches.Count -lt 3) {
    throw 'Expected Restaurant, BreadcrumbList and FAQPage schema blocks in Halal.html.'
  }

  $html = $html.Remove($schemaMatches[2].Index, $schemaMatches[2].Length).Insert($schemaMatches[2].Index, $faqSchema)
  $html = $html.Remove($schemaMatches[1].Index, $schemaMatches[1].Length).Insert($schemaMatches[1].Index, $breadcrumbSchema)
  $html = $html.Remove($schemaMatches[0].Index, $schemaMatches[0].Length).Insert($schemaMatches[0].Index, $restaurantSchema)
  return $html
}

function ImageForLine($line, $position) {
  $text = ([string]$line).ToLowerInvariant()
  if ($text -match 'pizza|margherita|meatfeast|pepperoni|tikka|dough') {
    return 'https://images.unsplash.com/photo-1574071318508-1cdbab80d002?w=700&q=85'
  }
  if ($text -match 'chicken|chick|drip|tender|wing|peri|parmo|parmasan') {
    if ($text -match 'parmo|parmasan') {
      return 'https://images.unsplash.com/photo-1601924994987-69e26d50dc26?w=700&q=85'
    }
    return 'https://images.unsplash.com/photo-1606755962773-d324e0a13086?w=700&q=85'
  }
  if ($text -match 'fries|chips|spice bag|loaded') {
    return 'https://images.unsplash.com/photo-1576107232684-1279f390859f?w=700&q=85'
  }
  if ($text -match 'donner|pitta|wrap|burrito|box') {
    return 'https://images.unsplash.com/photo-1555939594-58d7cb561ad1?w=700&q=85'
  }
  if ($text -match 'delivery|collection|order|postcode|driver|address|contact|phone|email|map') {
    return 'https://images.unsplash.com/photo-1551782450-a2132b4ba21d?w=900&q=85'
  }
  if ($text -match 'burger|smash|angus|beef|cheese|bigboy|americano|rack city special') {
    return 'https://images.unsplash.com/photo-1568901346375-23c9450c58cd?w=1200&q=85'
  }
  if ($text -match 'milkshake|dessert|cake|brownie|cookie') {
    return 'https://images.unsplash.com/photo-1550547660-d9450f859349?w=900&q=85'
  }

  $fallback = @(
    'https://images.unsplash.com/photo-1568901346375-23c9450c58cd?w=1200&q=85',
    'https://images.unsplash.com/photo-1606755962773-d324e0a13086?w=700&q=85',
    'https://images.unsplash.com/photo-1555939594-58d7cb561ad1?w=700&q=85',
    'https://images.unsplash.com/photo-1601924994987-69e26d50dc26?w=700&q=85',
    'https://images.unsplash.com/photo-1576107232684-1279f390859f?w=700&q=85',
    'https://images.unsplash.com/photo-1574071318508-1cdbab80d002?w=700&q=85',
    'https://images.unsplash.com/photo-1551782450-a2132b4ba21d?w=900&q=85'
  )
  return $fallback[$position % $fallback.Count]
}

function ImageSizeForLine($line) {
  $len = ([string]$line).Length
  if ($len -le 12) {
    return [ordered]@{ width = 92; height = 62; card = '320px' }
  }
  if ($len -le 40) {
    return [ordered]@{ width = 118; height = 78; card = '460px' }
  }
  if ($len -le 110) {
    return [ordered]@{ width = 156; height = 104; card = '640px' }
  }
  return [ordered]@{ width = 210; height = 132; card = '820px' }
}

function DisplayText($line) {
  $text = [string]$line
  if ($text -match '^H1:\s*(.+)$') { return $Matches[1].Trim() }
  if ($text -match '^Section:\s*(.+)$') { return $Matches[1].Trim() }
  if ($text -match '^Category\s+\d+:\s*(.+)$') { return $Matches[1].Trim() }
  if ($text -match '^Global note:\s*(.+)$') { return $Matches[1].Trim() }
  if ($text -match '^Trust badges \(above fold\):\s*(.+)$') { return $Matches[1].Trim() }
  if ($text -match '^CTAs:\s*(.+)$') { return $Matches[1].Trim() }
  if ($text -match '^Local entity reinforcement block \(footer\):\s*(.+)$') { return $Matches[1].Trim() }
  if ($text -match '^Author trust signal \(footer\):\s*(.+)$') { return $Matches[1].Trim() }
  if ($text -match '^FAQPage schema.*$') { return 'FAQ schema' }
  if ($text -match '^BreadcrumbList schema.*$') { return 'Breadcrumb schema' }
  return $text
}

function CleanPublicText($value) {
  $text = ([string]$value).Trim()
  if ($text.Length -eq 0) { return $text }
  $text = $text -replace '\s*\([^)]*(icon|word title|line benefit|categories with internal links|internal links|steps)[^)]*\)', ''
  $text = $text -replace '\s*Halal status of [^.]+ pending client confirmation\.', ''
  $text = $text -replace '\s*Halal status pending client confirmation\.', ''
  $text = $text -replace '\s*Pricing pending client confirmation \([^)]+\)\.', ''
  $text = $text -replace '\s*Flavours pending client confirmation\.?', ''
  $text = $text -replace '\s*Mozzarella sticks availability pending client confirmation \([^)]+\)\.', ''
  $text = $text -replace '\s*\([^)]*pending confirmation[^)]*\)', ''
  $text = $text -replace '\s+', ' '
  $text = $text -replace '(May 2026)(Written by)', '$1. $2'
  $text = $text -replace '(Blackburn\.)(Content reviewed)', '$1 $2'
  return $text.Trim()
}

function IsInternalInstruction($line) {
  $text = ([string]$line).Trim()
  if ($text.Length -eq 0) { return $true }
  if ($text -match '^(Icon|Title|Dish|Benefit|Burger|Description|Price|Pizza|9" Price|12" Price|Item|Offer|Inclusions|Town|Postcodes|Notes|Smash Burger|Regular Burger)$') { return $true }
  if ($text -match '^(Trust badges \(above fold\):|CTAs:|Menu items \(|Menu items \(from master menu\):|Local entity reinforcement block|Author trust signal|Allergen information available on request|SCHEMA|BreadcrumbList schema|FAQPage schema|LocalBusiness / Restaurant schema|json$)') { return $true }
  if ($text -match '^\(LocalBusiness schema same|^\(Note:|^Hero image alt text|^Page purpose|^Conversion block|^Internal links|^Schema notes|^Complete page copy|^Client confirmations still required') { return $true }
  if ($text -match 'pending client confirmation|pending confirmation|Fabricated items|omitted for brevity|must be included in final|master menu') { return $true }
  if ($text -match '^\[Embed Google Map') { return $true }
  return $false
}

function IsSectionHeading($line) {
  return $line -match '^(Section:|Category \d+:|Frequently Asked Questions|SCHEMA|Customer Review|Customer Suggestion|How to Order|Delivery|Collection|Opening Hours|Address|Phone|Email|Social Media|Map|Our Values|Our Kitchen|Why Franchise With Us\?|What You Get As a Franchise Partner|Investment & Availability|Enquiry Form|Delivery Questions|Halal Questions|Menu Questions|Order and Collection Questions|Franchise Questions|Contact Questions|Delivery by Town|How to Check Delivery|Late Night Hours|Late Night Menu|Late Night Delivery|Comparison Table|What Makes a Smash Burger Good\?|Where to Get Smash Burgers in Blackburn|What Is In a Parmo\?|Where Did the Parmo Originate\?|Halal Parmo in Blackburn|Why Our Donner Is Different|Delivery Zones|How Delivery Works|Our Halal Menu|Why Our Smash Burgers Are Different|Results That Matter|Our Menu|Real Customer Reviews|Our Services|Why Blackburn Chooses Rack City Kitchen|100% Halal Certified|Final CTA Band)'
}

function IsOperationalLine($line) {
  $text = ([string](DisplayText $line)).ToLowerInvariant()
  if ($text -match '^(monday|tuesday|wednesday|thursday|friday|saturday|sunday):') { return $true }
  if ($text -match 'opening hours|late night hours|delivery|collection|address|phone|email|map|social media|instagram|tiktok|facebook|postcode|minimum order|last delivery|last collection|get directions|call for|order online|checkout|faq|question|answer') { return $true }
  if ($text -match '^£|^add £|^\d|closed$|bb1|bb2|bb3|bb5|bb6|01254|info@') { return $true }
  if (QuestionAnswer $line) { return $true }
  return $false
}

function ShouldUseImage($line) {
  if (IsOperationalLine $line) { return $false }
  $text = ([string](DisplayText $line)).ToLowerInvariant()
  return $text -match 'burger|smash|angus|beef|chicken|chick|drip|tender|wing|peri|parmo|parmasan|pizza|margherita|meatfeast|pepperoni|tikka|dough|fries|chips|spice bag|loaded|donner|pitta|wrap|burrito|box|milkshake|dessert|cake|brownie|cookie'
}

function IsReviewLine($line) {
  $text = ([string](DisplayText $line)).Trim()
  return $text -match '(Google Review|Google|Shizzy Shiz|Aaron Hargreaves|Naomi Shep)' -and $text -match '^[\p{Pi}"'']'
}

function IsCalloutLine($line) {
  $text = ([string]$line).Trim()
  if ($text -match '^Global note:|^Important:|^Add extras:|^Add ons available:|^Make it a meal upgrade|^Allergen information') { return $true }
  return $false
}

function RenderReviewCardHtml($line) {
  $encoded = HtmlText (DisplayText $line)
  return @"
    <article class="doc-review-card reveal">
      <div class="doc-review-stars" aria-hidden="true">5/5 REVIEW</div>
      <p>$encoded</p>
    </article>
"@
}

function RenderReviewsTrack($reviewLines) {
  if ($reviewLines.Count -eq 0) { return '' }
  $cards = @()
  foreach ($review in $reviewLines) { $cards += RenderReviewCardHtml $review }
  return ($cards -join "`n")
}

function IsFeatureLine($line) {
  $text = ([string](DisplayText $line)).Trim()
  if ($text -match '^[^?]{2,80}\s*[\u2010-\u2015-]\s*.{3,}$') { return $true }
  if ($text -match '^(\d+\+|Under \d+|Contact us|Full training|Distinctive|Standardised|Marketing|Launch support|Ongoing)' -or $text -match 'Google rating') { return $true }
  return $false
}

function CopyLineHtml($line, $position) {
  if (IsInternalInstruction $line) { return '' }
  $display = CleanPublicText (DisplayText $line)
  if ($display.Length -eq 0) { return '' }
  $encoded = HtmlText $display
  if ($position -eq 0 -and $line.StartsWith('H1:')) {
    return "    <h2 class=`"doc-page-heading reveal`" style=`"font-family:'Bebas Neue',sans-serif;font-size:clamp(44px,5vw,76px);letter-spacing:2px;line-height:0.92;margin:0 0 22px;color:var(--white);`">$encoded</h2>"
  }
  if ($line -match '^Local entity reinforcement block \(footer\):|^Author trust signal \(footer\):') {
    return "    <p class=`"doc-text reveal`" style=`"max-width:760px;font-size:13px;font-weight:500;line-height:1.7;color:rgba(255,255,255,0.42);margin:18px 0;text-transform:uppercase;letter-spacing:1px;`">$encoded</p>"
  }
  if (IsSectionHeading $line) {
    return "    <h2 class=`"doc-heading reveal`" style=`"font-family:'Bebas Neue',sans-serif;font-size:clamp(30px,3vw,52px);letter-spacing:2px;line-height:0.95;margin:58px 0 20px;`">$encoded</h2>"
  }
  $qa = QuestionAnswer $line
  if ($qa -ne $null) {
    $q = HtmlText (CleanPublicText $qa.q)
    $a = HtmlText (CleanPublicText $qa.a)
    return @"
    <article class="faq-item doc-faq-item reveal">
      <button class="faq-question" type="button" aria-expanded="false">
        <span>$q</span>
        <span class="faq-icon" aria-hidden="true">+</span>
      </button>
      <div class="faq-answer">
        <p>$a</p>
      </div>
    </article>
"@
  }
  if (IsReviewLine $line) {
    return RenderReviewsTrack @($line)
  }
  if (IsCalloutLine $line) {
    return @"
    <aside class="doc-callout reveal">
      <span class="doc-callout-kicker">Note</span>
      <p>$encoded</p>
    </aside>
"@
  }
  if (IsFeatureLine $line) {
    return "    <p class=`"doc-feature-line reveal`">$encoded</p>"
  }
  return "    <p class=`"doc-text reveal`" style=`"max-width:820px;font-size:16px;font-weight:300;line-height:1.85;color:rgba(255,255,255,0.72);margin:12px 0;`">$encoded</p>"
}

function IsPriceLine($line) {
  $text = ([string](DisplayText $line)).Trim()
  $pound = [string][char]0x00A3
  if ($text.StartsWith($pound) -or $text.Contains($pound)) { return $true }
  if ($text -match '^£') { return $true }
  if ($text -match '^Add £') { return $true }
  if ($text -match '^\d+x') { return $true }
  if ($text -match '^\d+\s') { return $true }
  if ($text -match '^(Free|Contact us|Estimated|Franchise fee|Option to change|Rotating|–|-)$') { return $true }
  return $false
}

function IsTableStop($line) {
  if (-not $line) { return $true }
  if ($line.StartsWith('H1:')) { return $true }
  if (IsSectionHeading $line) { return $true }
  if ($line -match '^Global note:|^Add extras:|^Add ons available:|^Make it a meal upgrade|^Allergen information|^Local entity reinforcement block|^Author trust signal|^SCHEMA|^FAQPage schema|^BreadcrumbList schema') { return $true }
  return $false
}

function IsColumnLabel($line) {
  return $line -in @('Flavours', 'Description', 'Inclusions', 'Notes')
}

function RenderItemGrid($items, $label) {
  if ($items.Count -eq 0) { return '' }
  $cards = ''
  foreach ($item in $items) {
    $nameText = CleanPublicText $item.name
    $descText = CleanPublicText $item.description
    $priceText = CleanPublicText $item.price
    if ($nameText.Length -eq 0 -or $priceText.Length -eq 0) { continue }
    $name = HtmlText $nameText
    $desc = HtmlText $descText
    $price = HtmlText $priceText
    $image = ImageForLine $item.name 0
    if (-not (ShouldUseImage $item.name)) {
      $image = ImageForLine $item.description 0
    }
    $cards += @"
      <article class="doc-item-card reveal">
        <img src="$image" alt="$name" loading="lazy" width="640" height="360">
        <div class="doc-item-body">
          <div class="doc-item-top">
            <h3>$name</h3>
            <span class="doc-item-price">$price</span>
          </div>
          <p>$desc</p>
        </div>
      </article>

"@
  }
  if ($cards.Length -eq 0) { return '' }
  return @"
    <div class="doc-item-grid" aria-label="$(HtmlText $label)">
$cards
    </div>
"@
}

function RenderInfoGrid($items, $label) {
  if ($items.Count -eq 0) { return '' }
  $cards = ''
  foreach ($item in $items) {
    $icon = HtmlText (CleanPublicText $item.icon)
    $titleText = CleanPublicText $item.title
    $benefitText = CleanPublicText $item.benefit
    if ($titleText.Length -eq 0 -or $benefitText.Length -eq 0) { continue }
    $title = HtmlText $titleText
    $benefit = HtmlText $benefitText
    $cards += @"
      <article class="doc-info-card reveal">
        <div class="doc-info-icon">$icon</div>
        <h3>$title</h3>
        <p>$benefit</p>
      </article>

"@
  }
  if ($cards.Length -eq 0) { return '' }
  return @"
    <div class="doc-info-grid" aria-label="$(HtmlText $label)">
$cards
    </div>
"@
}

function RenderTileGrid($items, $label) {
  if ($items.Count -eq 0) { return '' }
  $cards = ''
  foreach ($item in $items) {
    $titleText = CleanPublicText $item
    if ($titleText.Length -eq 0) { continue }
    $title = HtmlText $titleText
    $image = ImageForLine $titleText 0
    $cards += @"
      <article class="doc-tile-card reveal">
        <img src="$image" alt="$title" loading="lazy" width="640" height="360">
        <h3>$title</h3>
      </article>

"@
  }
  if ($cards.Length -eq 0) { return '' }
  return @"
    <div class="doc-tile-grid" aria-label="$(HtmlText $label)">
$cards
    </div>
"@
}

function BuildCopyHtml($lines) {
  $html = @()
  $i = 0
  $pos = 0
  while ($i -lt $lines.Count) {
    $line = $lines[$i]

    $display = DisplayText $line

    if ($line -eq 'Icon' -and ($i + 2) -lt $lines.Count -and $lines[$i + 1] -in @('Title', 'Dish') -and $lines[$i + 2] -eq 'Benefit') {
      $label = 'Highlights'
      $i += 3
      $items = @()
      while (($i + 2) -lt $lines.Count -and -not (IsTableStop $lines[$i]) -and -not (IsTableStop $lines[$i + 1]) -and -not (IsTableStop $lines[$i + 2])) {
        $items += [pscustomobject]@{
          icon = DisplayText $lines[$i]
          title = DisplayText $lines[$i + 1]
          benefit = DisplayText $lines[$i + 2]
        }
        $i += 3
      }
      $html += RenderInfoGrid $items $label
      continue
    }

    $sectionTitle = CleanPublicText (DisplayText $line)
    if ((IsSectionHeading $line) -and ($sectionTitle -eq 'Our Menu' -or $sectionTitle -like 'Late Night Menu*')) {
      $heading = CopyLineHtml $line $pos
      $i++
      $pos++
      $items = @()
      while ($i -lt $lines.Count -and -not (IsTableStop $lines[$i]) -and -not (QuestionAnswer $lines[$i])) {
        if (-not (IsInternalInstruction $lines[$i])) {
          $itemText = CleanPublicText (DisplayText $lines[$i])
          if ($itemText.Length -gt 0) { $items += $itemText }
        }
        $i++
        $pos++
      }
      $html += $heading
      $html += RenderTileGrid $items $sectionTitle
      continue
    }

    if ($line -in @('Burger', 'Item', 'Offer') -and ($i + 5) -lt $lines.Count -and $lines[$i + 1] -in @('Description', 'Inclusions') -and $lines[$i + 2] -eq 'Price' -and (IsPriceLine $lines[$i + 5])) {
      $label = $line
      $i += 3
      $items = @()
      while (($i + 2) -lt $lines.Count -and -not (IsTableStop $lines[$i]) -and -not (IsTableStop $lines[$i + 1]) -and -not (IsTableStop $lines[$i + 2]) -and (IsPriceLine $lines[$i + 2])) {
        $items += [pscustomobject]@{
          name = DisplayText $lines[$i]
          description = DisplayText $lines[$i + 1]
          price = DisplayText $lines[$i + 2]
        }
        $i += 3
      }
      $html += RenderItemGrid $items $label
      continue
    }

    if ($line -eq 'Item' -and ($i + 2) -lt $lines.Count -and $lines[$i + 1] -eq 'Price') {
      $label = 'Items'
      $i += 2
      $items = @()
      while ($i -lt $lines.Count -and -not (IsTableStop $lines[$i])) {
        $name = DisplayText $lines[$i]
        $description = ''
        $price = ''
        if (($i + 1) -lt $lines.Count -and -not (IsTableStop $lines[$i + 1]) -and (IsPriceLine $lines[$i + 1])) {
          $price = DisplayText $lines[$i + 1]
          $i += 2
        } elseif (($i + 2) -lt $lines.Count -and -not (IsTableStop $lines[$i + 1]) -and -not (IsTableStop $lines[$i + 2]) -and (IsPriceLine $lines[$i + 2])) {
          if (IsColumnLabel $lines[$i]) {
            $name = DisplayText $lines[$i + 1]
          } else {
            $name = ((DisplayText $lines[$i]) + ' ' + (DisplayText $lines[$i + 1])).Trim()
          }
          $price = DisplayText $lines[$i + 2]
          $i += 3
        } else {
          break
        }

        if ($i -lt $lines.Count -and -not (IsTableStop $lines[$i]) -and -not (IsPriceLine $lines[$i]) -and -not (($i + 1) -lt $lines.Count -and (IsPriceLine $lines[$i + 1]))) {
          $description = DisplayText $lines[$i]
          $i++
        }

        $items += [pscustomobject]@{
          name = $name
          description = $description
          price = $price
        }
      }
      if ($items.Count -gt 0) {
        $html += RenderItemGrid $items $label
        continue
      }
    }

    if ($line -eq 'Pizza' -and ($i + 7) -lt $lines.Count -and $lines[$i + 1] -eq '9" Price' -and $lines[$i + 2] -eq '12" Price' -and $lines[$i + 3] -eq 'Description' -and (IsPriceLine $lines[$i + 5]) -and (IsPriceLine $lines[$i + 6])) {
      $i += 4
      $items = @()
      while (($i + 3) -lt $lines.Count -and -not (IsTableStop $lines[$i]) -and -not (IsTableStop $lines[$i + 1]) -and -not (IsTableStop $lines[$i + 2]) -and -not (IsTableStop $lines[$i + 3]) -and (IsPriceLine $lines[$i + 1]) -and (IsPriceLine $lines[$i + 2])) {
        $items += [pscustomobject]@{
          name = DisplayText $lines[$i]
          description = DisplayText $lines[$i + 3]
          price = ('9" ' + (DisplayText $lines[$i + 1]) + ' / 12" ' + (DisplayText $lines[$i + 2]))
        }
        $i += 4
      }
      $html += RenderItemGrid $items 'Pizzas'
      continue
    }

    if (IsReviewLine $line) {
      $reviews = @()
      while ($i -lt $lines.Count -and (IsReviewLine $lines[$i])) {
        $reviews += $lines[$i]
        $i++
        $pos++
      }
      $html += RenderReviewsTrack $reviews
      continue
    }

    $html += CopyLineHtml $line $pos
    $i++
    $pos++
  }
  return ($html -join "`n")
}

function BuildReviewsSection() {
  return @'
<section class="reviews-section" aria-label="Real Customer Reviews">
  <div class="reviews-header reveal"><div><div class="section-tag">Real Customer Reviews</div><h2 class="section-title">REAL CUSTOMER<br><span class="stroke">REVIEWS</span></h2></div><p>4.6&#9733; Google rating (327 reviews)</p></div>
  <div class="reviews-track-wrap" aria-hidden="true"><div class="reviews-track fwd">
    <div class="review-card"><div class="review-stars"><span>&#9733;</span><span>&#9733;</span><span>&#9733;</span><span>&#9733;</span><span>&#9733;</span></div><p class="review-text">&ldquo;The OG Smash Burger was a classic done right. The beef patty was well-seasoned, smashed to perfection with crispy edges. One of the best smash burgers I&rsquo;ve had.&rdquo;</p><div class="review-author">Shizzy Shiz (Google Review)</div></div>
    <div class="review-card"><div class="review-stars"><span>&#9733;</span><span>&#9733;</span><span>&#9733;</span><span>&#9733;</span><span>&#9733;</span></div><p class="review-text">&ldquo;Arrived just as opened &ndash; manager was really friendly. Food wait time was not long. Food was class 10/10. Smashed it!&rdquo;</p><div class="review-author">Aaron Hargreaves (Google Review)</div></div>
    <div class="review-card"><div class="review-stars"><span>&#9733;</span><span>&#9733;</span><span>&#9733;</span><span>&#9733;</span><span>&#9733;</span></div><p class="review-text">&ldquo;The food is always fresh and full of flavour. Staff are so polite and friendly. Best spot to eat in Blackburn.&rdquo;</p><div class="review-author">Naomi Shep (Google Review)</div></div>
  </div></div>
</section>
'@
}

function BuildHomeMain($page) {
  return @"
<main>
<style>
  .home-services .cat-grid { grid-template-columns: repeat(3, minmax(0, 1fr)); align-items: stretch; }
  .home-services .cat-card,
  .home-services .cat-card:nth-child(1) { grid-column: auto; aspect-ratio: 1 / 1; min-height: 420px; height: 100%; }
  .home-services .cat-content { min-height: 142px; display: flex; flex-direction: column; justify-content: flex-end; }
  .home-services .cat-desc { max-width: 100%; }
  .home-final-cta { align-self: center; justify-self: center; width: min(520px, 100%); text-align: center; display: flex; flex-direction: column; align-items: center; }
  .home-final-cta .section-tag { justify-content: center; }
  .home-final-cta .cta-row { display: flex; justify-content: center; gap: 0; flex-wrap: wrap; width: 100%; margin-top: 0; }
  .home-final-cta .cta-row .btn-outline { min-width: 190px; justify-content: center; }
  @media (max-width: 980px) {
    .home-services .cat-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
    .home-services .cat-card,
    .home-services .cat-card:nth-child(1) { min-height: 360px; }
  }
  @media (max-width: 620px) {
    .home-services .cat-grid { grid-template-columns: 1fr; }
    .home-services .cat-card,
    .home-services .cat-card:nth-child(1) { aspect-ratio: 4 / 3; min-height: 300px; }
    .home-final-cta .cta-row .btn-outline { width: 100%; }
  }
</style>
<section class="hero" id="hero" aria-label="Homepage">
  <div class="hero-bg" id="heroBg" role="img" aria-label="Homepage" style="background-image:url('/assets/images/og-smash-burger-banner-landscape.jpg')"></div>
  <div class="hero-overlay"></div>
  <div class="hero-content">
    <div class="hero-left">
      <div class="hero-eyebrow">halal smash burgers Blackburn</div>
      <h1 class="hero-title"><span class="line"><span class="line-inner">Halal Smash Burgers Blackburn &ndash; Fresh, Fast, Delivered</span></span></h1>
      <p class="hero-sub">Halal smash burgers, loaded fries and parmasans in Blackburn. &pound;2 delivery. 4.6&#9733; Google rating (327 reviews). Order online.</p>
      <div class="hero-btns"><a href="/order" class="btn-primary">Order Online &rarr;</a><a href="/menu/" class="btn-outline">View Full Menu &darr;</a></div>
    </div>
    <div class="hero-right">
      <div class="trust-badge"><div class="trust-icon">&#11088;</div><div class="trust-text"><div class="t-label">4.6&#9733; Google rating (327 reviews)</div></div></div>
      <div class="trust-badge"><div class="trust-icon">&#10003;</div><div class="trust-text"><div class="t-label">100% Halal Certified</div></div></div>
      <div class="trust-badge"><div class="trust-icon">&#128666;</div><div class="trust-text"><div class="t-label">&pound;2 Delivery</div></div></div>
      <div class="trust-badge"><div class="trust-icon">&#128293;</div><div class="trust-text"><div class="t-label">Fresh Daily</div></div></div>
    </div>
  </div>
</section>

<section class="story-section" aria-labelledby="home-heading">
  <div class="story-img"><img src="https://images.unsplash.com/photo-1568901346375-23c9450c58cd?w=1200&q=85" alt="Halal smash burgers Blackburn" loading="lazy"><div class="story-img-overlay"></div><div class="story-img-label" aria-hidden="true">RCK</div></div>
  <div class="story-content">
    <h2 class="section-title reveal" id="home-heading">HALAL SMASH BURGERS<br><span class="stroke">BLACKBURN</span></h2>
    <p class="reveal">Rack City Kitchen makes fresh halal smash burgers in Blackburn. Based at 36 Copy Nook, we serve Blackburn with halal smash burgers, loaded fries, parmasans and late-night takeaway. We use 100% halal Angus beef. Every patty is smashed thin on a hot grill. Crispy edges. Juicy centre.</p>
  </div>
</section>

$(BuildReviewsSection)

<section class="categories-section home-services" aria-labelledby="services-heading">
  <div class="categories-header"><div><div class="section-tag reveal">Our Services</div><h2 class="section-title reveal" id="services-heading">OUR<br><span class="stroke">SERVICES</span></h2></div><p class="reveal">Smash Burgers. Chicken Burgers. Loaded Fries. Stir Fry Donner. Parmasans. Pizzas.</p></div>
  <div class="cat-grid">
    <a href="/smash-burgers-blackburn/" class="cat-card reveal"><img src="https://images.unsplash.com/photo-1568901346375-23c9450c58cd?w=1200&q=85" alt="Smash Burgers" loading="lazy"><div class="cat-overlay"></div><div class="cat-content"><span class="cat-name">SMASH BURGERS</span><p class="cat-desc">13 burgers, Angus beef, smashed thin.</p></div><div class="cat-arrow">&#8599;</div></a>
    <a href="/chicken-burgers-blackburn/" class="cat-card reveal"><img src="https://images.unsplash.com/photo-1606755962773-d324e0a13086?w=700&q=85" alt="Chicken Burgers" loading="lazy"><div class="cat-overlay"></div><div class="cat-content"><span class="cat-name">CHICKEN BURGERS</span><p class="cat-desc">Classic, Hot Chick, Drip, Parmo, Peri.</p></div><div class="cat-arrow">&#8599;</div></a>
    <a href="/loaded-fries-blackburn/" class="cat-card reveal"><img src="https://images.unsplash.com/photo-1576107232684-1279f390859f?w=700&q=85" alt="Loaded Fries" loading="lazy"><div class="cat-overlay"></div><div class="cat-content"><span class="cat-name">LOADED FRIES</span><p class="cat-desc">Donner on chips &amp; spice bag.</p></div><div class="cat-arrow">&#8599;</div></a>
    <a href="/stir-fry-donner-blackburn/" class="cat-card reveal"><img src="https://images.unsplash.com/photo-1555939594-58d7cb561ad1?w=700&q=85" alt="Stir Fry Donner" loading="lazy"><div class="cat-overlay"></div><div class="cat-content"><span class="cat-name">STIR FRY DONNER</span><p class="cat-desc">Donner on chips, in pitta, or portion.</p></div><div class="cat-arrow">&#8599;</div></a>
    <a href="/parmasans-blackburn/" class="cat-card reveal"><img src="https://images.unsplash.com/photo-1601924994987-69e26d50dc26?w=700&q=85" alt="Parmasans" loading="lazy"><div class="cat-overlay"></div><div class="cat-content"><span class="cat-name">PARMASANS</span><p class="cat-desc">Classic, Hot Shot, Donner, Pepperoni, Peri.</p></div><div class="cat-arrow">&#8599;</div></a>
    <a href="/pizza-blackburn/" class="cat-card reveal"><img src="https://images.unsplash.com/photo-1574071318508-1cdbab80d002?w=700&q=85" alt="Pizzas" loading="lazy"><div class="cat-overlay"></div><div class="cat-content"><span class="cat-name">PIZZAS</span><p class="cat-desc">11 options, 9&quot; &amp; 12&quot;, with a drink.</p></div><div class="cat-arrow">&#8599;</div></a>
  </div>
</section>

<div class="stats-bar" aria-label="Results That Matter">
  <div class="stat-item reveal"><span class="stat-num">500+</span><span class="stat-label">happy customers</span></div>
  <div class="stat-item reveal"><span class="stat-num">4.6&#9733;</span><span class="stat-label">Google rating (327 reviews)</span></div>
  <div class="stat-item reveal"><span class="stat-num">20+</span><span class="stat-label">signature items</span></div>
  <div class="stat-item reveal"><span class="stat-num">45</span><span class="stat-label">Under 45 min average delivery time</span></div>
</div>

<section class="menu-feature" aria-labelledby="menu-heading">
  <div class="menu-feature-header"><div><div class="section-tag reveal">Our Menu</div><h2 class="section-title reveal" id="menu-heading">OUR<br><span class="stroke">MENU</span></h2></div><p class="reveal">OG Smash Burger. Hot Chick Burger. Drip Burger. Classic Parmo. Stir Fry Donner Pizza. Spice Bag. Peri Peri Chicken.</p></div>
  <div class="menu-grid">
    <div class="menu-item-card reveal"><div class="mic-num">01</div><span class="mic-cat">Our Menu</span><img class="mic-img" src="https://images.unsplash.com/photo-1568901346375-23c9450c58cd?w=600&q=80" alt="OG Smash Burger" loading="lazy"><div class="mic-name">OG SMASH BURGER</div><p class="mic-desc">OG Smash Burger</p></div>
    <div class="menu-item-card reveal"><div class="mic-num">02</div><span class="mic-cat">Our Menu</span><img class="mic-img" src="https://images.unsplash.com/photo-1606755962773-d324e0a13086?w=600&q=80" alt="Hot Chick Burger" loading="lazy"><div class="mic-name">HOT CHICK BURGER</div><p class="mic-desc">Hot Chick Burger</p></div>
    <div class="menu-item-card reveal"><div class="mic-num">03</div><span class="mic-cat">Our Menu</span><img class="mic-img" src="https://images.unsplash.com/photo-1606755962773-d324e0a13086?w=600&q=80" alt="Drip Burger" loading="lazy"><div class="mic-name">DRIP BURGER</div><p class="mic-desc">Drip Burger</p></div>
    <div class="menu-item-card reveal"><div class="mic-num">04</div><span class="mic-cat">Our Menu</span><img class="mic-img" src="https://images.unsplash.com/photo-1601924994987-69e26d50dc26?w=600&q=80" alt="Classic Parmo" loading="lazy"><div class="mic-name">CLASSIC PARMO</div><p class="mic-desc">Classic Parmo</p></div>
    <div class="menu-item-card reveal"><div class="mic-num">05</div><span class="mic-cat">Our Menu</span><img class="mic-img" src="https://images.unsplash.com/photo-1574071318508-1cdbab80d002?w=600&q=80" alt="Stir Fry Donner Pizza" loading="lazy"><div class="mic-name">STIR FRY DONNER PIZZA</div><p class="mic-desc">Stir Fry Donner Pizza</p></div>
    <div class="menu-item-card reveal"><div class="mic-num">06</div><span class="mic-cat">Our Menu</span><img class="mic-img" src="https://images.unsplash.com/photo-1576107232684-1279f390859f?w=600&q=80" alt="Spice Bag" loading="lazy"><div class="mic-name">SPICE BAG</div><p class="mic-desc">Spice Bag</p></div>
    <div class="menu-item-card reveal"><div class="mic-num">07</div><span class="mic-cat">Our Menu</span><img class="mic-img" src="https://images.unsplash.com/photo-1606755962773-d324e0a13086?w=600&q=80" alt="Peri Peri Chicken" loading="lazy"><div class="mic-name">PERI PERI CHICKEN</div><p class="mic-desc">Peri Peri Chicken</p></div>
  </div>
  <div class="menu-cta"><a href="/order" class="btn-primary">Order Online &rarr;</a><a href="/menu/" class="btn-outline">View Full Menu &rarr;</a></div>
</section>

<section class="how-section" aria-labelledby="how-heading">
  <div class="section-tag reveal">How It Works</div><h2 class="section-title reveal" id="how-heading">HOW IT<br><span class="stroke">WORKS</span></h2>
  <div class="how-grid">
    <div class="how-step reveal"><div class="how-step-num" aria-hidden="true">01</div><div class="how-step-icon">&#128241;</div><h3>Order online</h3><p>Choose delivery or collection.</p></div>
    <div class="how-step reveal"><div class="how-step-num" aria-hidden="true">02</div><div class="how-step-icon">&#128293;</div><h3>We cook fresh</h3><p>Every item made to order.</p></div>
    <div class="how-step reveal"><div class="how-step-num" aria-hidden="true">03</div><div class="how-step-icon">&#9989;</div><h3>You enjoy</h3><p>Delivered hot or collected fast.</p></div>
  </div>
</section>

<div class="halal-strip" aria-label="Why Blackburn Chooses Rack City Kitchen">
  <div class="halal-point reveal"><span class="halal-point-icon" aria-hidden="true">&#10003;</span><h4>100% Halal</h4><p>Sourced from approved suppliers.</p></div>
  <div class="halal-point reveal"><span class="halal-point-icon" aria-hidden="true">&#128307;</span><h4>Late-Night</h4><p>Open until 23:40 (Thu-Mon).</p></div>
  <div class="halal-point reveal"><span class="halal-point-icon" aria-hidden="true">&#128293;</span><h4>Homemade Donner</h4><p>Wok-fired in-house.</p></div>
  <div class="halal-point reveal"><span class="halal-point-icon" aria-hidden="true">&#127828;</span><h4>Fresh Smashed Beef</h4><p>No frozen patties.</p></div>
</div>

<section class="order-split" aria-labelledby="delivery-heading">
  <div class="order-img-side"><img src="https://images.unsplash.com/photo-1551782450-a2132b4ba21d?w=900&q=85" alt="Delivery and collection" loading="lazy"><div class="order-img-overlay"></div></div>
  <div class="order-content-side"><div class="section-tag reveal">Delivery &amp; Collection</div><h2 class="section-title reveal" id="delivery-heading">DELIVERY &amp;<br><span class="stroke">COLLECTION</span></h2><p class="reveal">We deliver across Blackburn and selected surrounding areas. Enter your postcode at checkout to confirm availability. Standard delivery charge &pound;2. Minimum order &pound;15. You can also collect your order for free. Select collection at checkout. Last orders for delivery: 23:10. Last orders for collection: 23:30.</p><div class="doc-callout reveal"><span class="doc-callout-kicker">Important</span><p>Darwen (BB3) &ndash; selected postcodes only. Enter your full postcode at checkout to confirm delivery.</p></div></div>
</section>

<section class="delivery-section" aria-labelledby="halal-heading">
  <div class="delivery-left"><div class="section-tag reveal">100% Halal Certified</div><h2 class="section-title reveal" id="halal-heading">100% HALAL<br><span class="stroke">CERTIFIED</span></h2><p class="reveal">Rack City Kitchen is fully halal certified. All meat is sourced from approved halal suppliers. All ingredients are 100% halal</p></div>
  <div class="delivery-right home-final-cta"><div class="section-tag reveal">Final CTA Band</div><h3 class="reveal">HUNGRY?</h3><p class="reveal">Order now &ndash; fresh halal burgers delivered to your door.</p><div class="cta-row reveal"><a href="/order" class="btn-outline">Order Online &rarr;</a><a href="/menu/" class="btn-outline">View Full Menu &rarr;</a></div></div>
</section>

<section class="faq-section" aria-labelledby="faq-heading">
  <div class="section-tag reveal">Frequently Asked Questions</div><h2 class="section-title reveal" id="faq-heading">FREQUENTLY<br><span class="stroke">ASKED</span></h2>
  <div class="faq-wrap">
    <div class="faq-item reveal"><button class="faq-question" type="button" aria-expanded="false">What is a smash burger?<span class="faq-icon" aria-hidden="true">+</span></button><div class="faq-answer"><p>A smash burger is a beef patty pressed thin on a hot grill. Crispy edges, juicy centre.</p></div></div>
    <div class="faq-item reveal"><button class="faq-question" type="button" aria-expanded="false">Is Rack City Kitchen halal?<span class="faq-icon" aria-hidden="true">+</span></button><div class="faq-answer"><p>Yes. All meat is 100% halal certified. All ingredients are sourced from approved halal suppliers.</p></div></div>
    <div class="faq-item reveal"><button class="faq-question" type="button" aria-expanded="false">Do you deliver to Accrington?<span class="faq-icon" aria-hidden="true">+</span></button><div class="faq-answer"><p>We deliver across Blackburn and selected surrounding areas. Enter your postcode at checkout to confirm.</p></div></div>
    <div class="faq-item reveal"><button class="faq-question" type="button" aria-expanded="false">Can I collect my order?<span class="faq-icon" aria-hidden="true">+</span></button><div class="faq-answer"><p>Yes. Order online and collect from 36 Copy Nook, Blackburn.</p></div></div>
  </div>
</section>

<section class="cta-band" aria-label="Order call to action">
  <div class="cta-band-inner"><div class="cta-band-left reveal-left"><div class="section-tag">Final CTA Band</div><h2 class="cta-band-title">HUNGRY?<br>ORDER<br>NOW.</h2></div><div class="cta-band-right reveal-right"><p>Order now &ndash; fresh halal burgers delivered to your door.</p><div class="cta-band-btns"><a href="/order" class="btn-black">Order Online &rarr;</a><a href="/menu/" class="btn-black-outline">View Full Menu</a></div></div></div>
</section>
</main>
"@
}

function BuildMain($page) {
  if ($page.Slug -eq '/' -or $page.Name -eq 'Homepage') {
    $homeMain = BuildHomeMain $page
    if ($homeMain) { return $homeMain }
  }

  $h1Line = ($page.CopyLines | Where-Object { $_.StartsWith('H1:') } | Select-Object -First 1)
  if (-not $h1Line) { $h1Line = $page.Name }
  $heroTitle = HtmlText (DisplayText $h1Line)
  $heroIntro = HtmlText $page.Description
  $pageImage = ImageForLine $page.Name 0
  $contentClass = if ($page.Slug -eq '/menu/') { 'menu-feature inner-editorial full-menu-page' } else { 'menu-feature inner-editorial' }

  $copy = BuildCopyHtml $page.CopyLines

  return @"
<main>
<!-- MAIN CONTENT -->
<section class="hero" id="hero" aria-label="$(HtmlText $page.Name)">
  <div class="hero-bg" id="heroBg" role="img" aria-label="$(HtmlText $page.Name)" style="background-image:url('https://images.unsplash.com/photo-1568901346375-23c9450c58cd?w=1920&q=90')"></div>
  <div class="hero-overlay"></div>
  <div class="hero-content">
    <div class="hero-left">
      <div class="hero-eyebrow">$(HtmlText $page.Focus)</div>
      <h1 class="hero-title">
        <span class="line"><span class="line-inner">$heroTitle</span></span>
      </h1>
      <p class="hero-sub">$heroIntro</p>
      <div class="hero-btns">
        <a href="/order" class="btn-primary">Order Online &rarr;</a>
        <a href="/menu/" class="btn-outline">View Full Menu &darr;</a>
      </div>
    </div>
    <div class="hero-right">
      <div class="trust-badge"><div class="trust-icon">&#11088;</div><div class="trust-text"><div class="t-label">4.6 Star Rating</div><div class="t-detail">327 verified Google reviews</div></div></div>
      <div class="trust-badge"><div class="trust-icon">&#10003;</div><div class="trust-text"><div class="t-label">100% Halal Certified</div><div class="t-detail">All meat from approved suppliers</div></div></div>
      <div class="trust-badge"><div class="trust-icon">&#128666;</div><div class="trust-text"><div class="t-label">&pound;2 Delivery</div><div class="t-detail">Minimum order &pound;15</div></div></div>
      <div class="trust-badge"><div class="trust-icon">&#128293;</div><div class="trust-text"><div class="t-label">Fresh Daily</div><div class="t-detail">Made to order in Blackburn</div></div></div>
    </div>
  </div>
  <div class="hero-scroll"><div class="scroll-line"></div>Scroll to explore</div>
</section>

$ticker

<section class="$contentClass" aria-label="$(HtmlText $page.Name) content" style="padding-top:96px;">
  <div class="inner-editorial-shell" style="max-width:1180px;margin:0 auto;">
    <style>
      .menu-feature { position:relative; overflow:hidden; background:linear-gradient(180deg,var(--off-black) 0%,var(--black) 48%,var(--charcoal) 100%); border-top:1px solid rgba(255,255,255,0.06); }
      .inner-editorial-shell { position:relative; display:grid; grid-template-columns:minmax(420px,0.95fr) minmax(0,1.05fr); gap:72px; align-items:start; }
      .inner-editorial-shell::before { content:'RACK CITY'; position:absolute; top:-52px; right:0; font-family:'Bebas Neue',sans-serif; font-size:clamp(72px,13vw,180px); letter-spacing:4px; color:rgba(255,255,255,0.025); line-height:0.8; pointer-events:none; }
      .doc-media-panel { position:sticky; top:110px; min-height:680px; overflow:hidden; background:var(--black); border:1px solid rgba(255,255,255,0.08); border-radius:8px; box-shadow:0 30px 90px rgba(0,0,0,.38); }
      .doc-media-panel img { width:100%; height:100%; min-height:680px; object-fit:cover; filter:grayscale(1) brightness(.62) contrast(1.08); transform:scale(1.02); }
      .doc-media-panel::after { content:''; position:absolute; inset:0; background:linear-gradient(180deg,rgba(0,0,0,0) 10%,rgba(0,0,0,.82) 100%); }
      .doc-media-label { position:absolute; left:32px; bottom:32px; z-index:1; font-family:'Bebas Neue',sans-serif; font-size:clamp(56px,6vw,104px); line-height:.86; letter-spacing:2px; color:rgba(255,255,255,.94); max-width:360px; }
      .doc-content-flow { min-width:0; padding-top:10px; }
      .doc-page-heading { max-width:980px; margin-bottom:28px !important; }
      .doc-heading { position:relative; max-width:920px; padding:0; margin-top:76px; border:0; background:transparent; }
      .doc-heading::before { content:''; position:absolute; left:0; top:-18px; width:72px; height:1px; background:var(--white); }
      .doc-text { max-width:780px; padding:0; font-size:17px !important; line-height:1.9 !important; }
      .doc-content-flow > .doc-text:first-of-type { padding:0; border:0; background:transparent; box-shadow:none; font-size:18px !important; color:rgba(255,255,255,.78) !important; }
      .doc-feature-line { max-width:860px; margin:14px 0; padding:22px 24px 22px 58px; border:1px solid rgba(255,255,255,0.1); border-radius:8px; background:rgba(255,255,255,0.045); color:rgba(255,255,255,0.74); font-size:15px; font-weight:300; line-height:1.65; position:relative; box-shadow:0 18px 55px rgba(0,0,0,.2); }
      .doc-feature-line:nth-of-type(even) { background:rgba(255,255,255,0.065); transform:none; }
      .doc-feature-line::before { content:'+'; position:absolute; left:24px; top:22px; font-family:'Bebas Neue',sans-serif; color:var(--white); font-size:22px; line-height:1; }
      .doc-callout { max-width:860px; margin:22px 0 32px; padding:28px 30px; border-left:2px solid var(--white); border-radius:0 8px 8px 0; background:rgba(255,255,255,0.06); box-shadow:0 22px 60px rgba(0,0,0,.22); }
      .doc-callout-kicker { display:block; font-family:'Barlow Condensed',sans-serif; font-size:12px; font-weight:700; letter-spacing:2px; text-transform:uppercase; color:rgba(255,255,255,0.45); margin-bottom:8px; }
      .doc-callout p { margin:0; color:rgba(255,255,255,0.76); font-size:15px; font-weight:300; line-height:1.75; }
      .doc-review-card { max-width:860px; margin:18px 0; padding:32px; border:1px solid rgba(255,255,255,0.08); border-radius:8px; background:linear-gradient(180deg,var(--charcoal),var(--off-black)); box-shadow:0 24px 70px rgba(0,0,0,.28); }
      .doc-review-stars { color:var(--white); font-size:11px; letter-spacing:4px; margin-bottom:16px; font-family:'Barlow Condensed',sans-serif; text-transform:uppercase; }
      .doc-review-card p { margin:0; font-family:'DM Serif Display',serif; font-size:clamp(23px,2.2vw,34px); line-height:1.18; color:rgba(255,255,255,0.92); }
      .doc-item-grid { display:grid; grid-template-columns:repeat(auto-fill,minmax(260px,1fr)); grid-auto-rows:1fr; gap:18px; align-items:stretch; justify-content:start; margin:22px 0 44px; }
      .doc-item-card { height:100%; min-height:100%; display:grid; grid-template-rows:154px 1fr; border:1px solid rgba(255,255,255,0.08); background:rgba(255,255,255,0.03); transition:transform .35s var(--ease-out), background .35s var(--ease-out), border-color .35s var(--ease-out); }
      .doc-item-card:hover { transform:translateY(-4px); background:rgba(255,255,255,0.055); border-color:rgba(255,255,255,0.18); }
      .doc-item-card img { width:100%; height:154px; object-fit:cover; filter:grayscale(1); border-bottom:1px solid rgba(255,255,255,0.08); transition:filter .35s var(--ease-out), transform .35s var(--ease-out); }
      .doc-item-card:hover img { filter:grayscale(0); transform:scale(1.025); }
      .doc-item-body { display:flex; flex-direction:column; gap:14px; justify-content:space-between; padding:18px; height:100%; }
      .doc-item-top { display:flex; justify-content:space-between; gap:14px; align-items:flex-start; }
      .doc-item-top h3 { font-family:'Barlow Condensed',sans-serif; font-size:22px; letter-spacing:1px; text-transform:uppercase; margin:0; line-height:1; color:rgba(255,255,255,0.92); }
      .doc-item-price { flex-shrink:0; font-family:'Bebas Neue',sans-serif; font-size:24px; letter-spacing:1px; color:var(--white); line-height:1; text-align:right; }
      .doc-item-body p { font-size:14px; font-weight:300; line-height:1.65; color:rgba(255,255,255,0.64); margin:0; }
      .doc-info-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(220px,1fr)); gap:2px; margin:24px 0 44px; border:1px solid rgba(255,255,255,0.08); background:rgba(255,255,255,0.08); }
      .doc-info-card { background:var(--black); padding:26px; min-height:190px; display:flex; flex-direction:column; gap:14px; transition:background .35s var(--ease-out), transform .35s var(--ease-out); }
      .doc-info-card:hover { background:var(--charcoal); transform:translateY(-2px); }
      .doc-info-icon { font-size:34px; line-height:1; }
      .doc-info-card h3 { font-family:'Barlow Condensed',sans-serif; font-size:24px; letter-spacing:1px; text-transform:uppercase; margin:0; color:rgba(255,255,255,0.92); }
      .doc-info-card p { font-size:14px; font-weight:300; line-height:1.65; color:rgba(255,255,255,0.64); margin:0; }
      .doc-tile-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(210px,1fr)); gap:2px; margin:24px 0 48px; background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.08); }
      .doc-tile-card { position:relative; overflow:hidden; background:var(--off-black); display:grid; grid-template-rows:170px 1fr; min-height:250px; }
      .doc-tile-card img { width:100%; height:170px; object-fit:cover; filter:grayscale(1) brightness(.82); transition:filter .35s var(--ease-out), transform .35s var(--ease-out); }
      .doc-tile-card:hover img { filter:grayscale(0) brightness(.72); transform:scale(1.045); }
      .doc-tile-card h3 { font-family:'Barlow Condensed',sans-serif; font-size:22px; letter-spacing:1px; text-transform:uppercase; margin:0; color:rgba(255,255,255,0.92); line-height:1.05; padding:18px; align-self:center; }
      .doc-faq-item { max-width:760px; }
      .full-menu-page .inner-editorial-shell { display:block; max-width:1320px !important; }
      .full-menu-page .inner-editorial-shell::before { right:clamp(12px,4vw,48px); }
      .full-menu-page .doc-media-panel { display:none; }
      .full-menu-page .doc-content-flow { width:100%; max-width:1320px; margin:0 auto; padding-top:0; }
      .full-menu-page .doc-page-heading, .full-menu-page .doc-text { max-width:860px; }
      .full-menu-page .doc-item-grid { grid-template-columns:repeat(auto-fit,minmax(230px,1fr)); gap:22px; margin:28px 0 58px; }
      .full-menu-page .doc-item-card { grid-template-rows:160px 1fr; min-width:0; }
      .full-menu-page .doc-item-card img { height:160px; }
      .full-menu-page .doc-item-body { padding:18px; gap:12px; }
      .full-menu-page .doc-item-top { display:grid; grid-template-columns:minmax(0,1fr) auto; gap:12px; align-items:start; }
      .full-menu-page .doc-item-top h3 { min-width:0; font-size:20px; line-height:1.02; overflow-wrap:anywhere; }
      .full-menu-page .doc-item-price { font-size:22px; white-space:nowrap; }
      .full-menu-page .doc-item-body p { font-size:13px; line-height:1.55; }
      @media (max-width: 720px) {
        .inner-editorial-shell { grid-template-columns:1fr; gap:32px; }
        .doc-media-panel { position:relative; top:auto; min-height:340px; }
        .doc-media-panel img { min-height:340px; }
        .doc-media-label { font-size:46px; }
        .doc-feature-line:nth-of-type(even) { transform:none; }
        .doc-item-grid { grid-template-columns:1fr; }
        .doc-review-card { padding:20px; }
        .doc-feature-line { padding-right:14px; }
        .full-menu-page .doc-item-grid { grid-template-columns:1fr; gap:18px; }
      }
    </style>
    <aside class="doc-media-panel reveal" aria-hidden="true">
      <img src="$pageImage" alt="" loading="lazy">
      <div class="doc-media-label">$(HtmlText $page.Name)</div>
    </aside>
    <div class="doc-content-flow">
$copy
    </div>
  </div>
</section>

<section class="cta-band" aria-label="Order call to action">
  <div class="cta-band-inner">
    <div class="cta-band-left reveal-left">
      <div class="section-tag">Ready to Order?</div>
      <h2 class="cta-band-title">HUNGRY?<br>ORDER<br>NOW.</h2>
    </div>
    <div class="cta-band-right reveal-right">
      <p>Fresh halal burgers delivered to your door across Blackburn and surrounding areas. &pound;2 delivery. Minimum order &pound;15.</p>
      <div class="cta-band-btns">
        <a href="/order" class="btn-black">Order Online &rarr;</a>
        <a href="/menu/" class="btn-black-outline">View Full Menu</a>
      </div>
    </div>
  </div>
</section>
</main>

"@
}

$template = [System.IO.File]::ReadAllText($templatePath, [System.Text.Encoding]::UTF8)
$bodyToTicker = Slice $template '<body>' '<!-- HERO -->'
$ticker = Slice $template '<!-- TICKER -->' '<!-- REVIEWS MARQUEE -->'
$tail = SliceInclusiveFrom $template '<!-- FRANCHISE TEASER -->'
$tail = $tail -replace "Blackburn's premium halal takeaway\. Built for the bold\. Crafted for consistency\. Designed to expand\.", 'Serving Blackburn, Accrington, Darwen, Rishton, Great Harwood and surrounding Lancashire areas. Written by the Rack City Kitchen team in Blackburn.'
$paras = Read-DocxParagraphs $docxPath

$pageStarts = @()
foreach ($p in $paras) {
  if ($p.Text -match '^Page \d+\. ') { $pageStarts += $p.Index }
}
if ($pageStarts.Count -ne 18) {
  throw "Expected 18 pages in DOCX. Found $($pageStarts.Count)."
}

$completionIndex = ($paras | Where-Object { $_.Text -eq 'COMPLETION SUMMARY' } | Select-Object -First 1).Index
if (-not $completionIndex) {
  $completionIndex = ($paras | Select-Object -Last 1).Index + 1
}

$rawPages = @()
for ($i = 0; $i -lt $pageStarts.Count; $i++) {
  $start = $pageStarts[$i]
  $end = if ($i -lt ($pageStarts.Count - 1)) { $pageStarts[$i + 1] } else { $completionIndex }
  $lines = @($paras | Where-Object { $_.Index -ge $start -and $_.Index -lt $end } | Sort-Object Index | ForEach-Object { $_.Text })
  $meta = Parse-Meta $lines[1]
  $copyStart = -1
  $schemaStart = -1
  for ($j = 0; $j -lt $lines.Count; $j++) {
    if ($copyStart -lt 0 -and $lines[$j] -eq 'FULL PAGE COPY') { $copyStart = $j }
    if ($schemaStart -lt 0 -and $lines[$j].StartsWith('SCHEMA')) { $schemaStart = $j }
  }
  if ($copyStart -lt 0 -or $schemaStart -lt 0) {
    throw "Could not find FULL PAGE COPY or SCHEMA markers for $($lines[0])."
  }

  $copyLines = @($lines[($copyStart + 1)..($schemaStart - 1)])
  $schemaLines = @($lines[$schemaStart..($lines.Count - 1)])
  $scripts = Extract-JsonScripts $schemaLines

  $pageName = ($lines[0] -replace '^Page \d+\.\s*', '').Trim()
  $faqs = Parse-Faqs $copyLines $schemaLines

  $rawPages += [pscustomobject]@{
    Number = $i + 1
    Name = $pageName
    Slug = $meta.slug
    Focus = $meta.focus
    Title = $meta.title
    Description = $meta.description
    Canonical = $meta.canonical
    CopyLines = $copyLines
    SchemaLines = $schemaLines
    JsonScripts = $scripts
    Faqs = $faqs
  }
}

$restaurantSchema = ($rawPages[0].JsonScripts | Select-Object -First 1)
if (-not $restaurantSchema) {
  throw 'Could not extract Restaurant schema from Page 1.'
}

foreach ($page in $rawPages) {
  $isHome = ([string]$page.Slug).Trim() -eq '/' -or $page.Name -eq 'Homepage'
  $breadcrumbName = if ($isHome) { 'Home' } else { $page.Name }
  $breadcrumbSchema = BreadcrumbScript $breadcrumbName $page.Canonical $isHome
  $faqSchema = FaqScript $page.Faqs
  $headApplied = ApplyHead $template $page $restaurantSchema $breadcrumbSchema $faqSchema
  $head = $headApplied.Substring(0, $headApplied.IndexOf('<body>'))
  $main = BuildMain $page
  $html = $head + $bodyToTicker + $main + $tail
  $html = Externalize-Assets $html
  $html = $html -replace "Blackburn's premium halal takeaway\. Built for the bold\. Crafted for consistency\. Designed to expand\.", 'Serving Blackburn, Accrington, Darwen, Rishton, Great Harwood and surrounding Lancashire areas. Written by the Rack City Kitchen team in Blackburn.'
  $html = [regex]::Replace($html, 'Rack City Kitchen is fully halal certified\.\s*All meat is sourced from approved halal suppliers\.\s*All ingredients are 100% halal(?: throughout our kitchen\. We do not compromise on this\. Ever\.)?\.?', 'Rack City Kitchen is fully halal certified. All meat is sourced from approved halal suppliers. All ingredients are 100% halal.')

  if ($isHome) {
    $outPath = Join-Path $root 'index.html'
  } else {
    $slugPath = $page.Slug.Trim('/')
    $dir = Join-Path $root $slugPath
    if (-not (Test-Path $dir)) {
      New-Item -ItemType Directory -Path $dir | Out-Null
    }
    $outPath = Join-Path $dir 'index.html'
  }

  [System.IO.File]::WriteAllText($outPath, $html, $utf8)
}

if (-not (Test-Path $cssDir)) {
  New-Item -ItemType Directory -Path $cssDir -Force | Out-Null
}
if (-not (Test-Path $jsDir)) {
  New-Item -ItemType Directory -Path $jsDir -Force | Out-Null
}


$cssOutput = @()
$cssOutput += '/* Generated by build-rack-city-pages.ps1. Edit the generator, then rebuild. */'
$cssOutput += $cssBlocks
if ($utilityRules.Count -gt 0) {
  $cssOutput += '/* Extracted inline style utilities */'
  $cssOutput += $utilityRules
}
[System.IO.File]::WriteAllText($siteCssPath, (($cssOutput -join "`n`n") + "`n"), $utf8)

$jsOutput = @()
$jsOutput += '// Generated by build-rack-city-pages.ps1. Edit the generator or Halal.html source script, then rebuild.'
$jsOutput += $scriptBlocks
$jsOutput += @"
document.querySelectorAll('.nav-overlay a').forEach(link => {
  link.addEventListener('click', () => {
    if (typeof closeNav === 'function') closeNav();
  });
});
"@
[System.IO.File]::WriteAllText($siteJsPath, (($jsOutput -join "`n`n") + "`n"), $utf8)

Write-Host "Generated $($rawPages.Count) pages from Rack City Website content 15.05.2026.docx."
