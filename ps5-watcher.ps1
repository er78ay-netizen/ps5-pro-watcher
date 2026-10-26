<#
  PS5 Pro stock watcher - Denmark and Sweden
  -------------------------------------------
  Checks authorised retailers for PlayStation 5 Pro stock and pushes an alert
  to your phone (ntfy app) plus a Windows notification when one becomes available.

  One run = one check. Windows Task Scheduler runs it every 3 minutes
  (install.cmd sets that up).

  Manual use:
    powershell -ExecutionPolicy Bypass -File ps5-watcher.ps1 -ShowStatus   # run a check and print the table
    powershell -ExecutionPolicy Bypass -File ps5-watcher.ps1 -TestAlert    # send a test notification

  Files written next to this script:
    status.txt  - result of the latest check (open it any time)
    log.txt     - history of checks, alerts and errors
    state.json  - what was available last time (so you only get alerted on changes)
#>
[CmdletBinding()]
param(
    [switch]$TestAlert,
    [switch]$ShowStatus
)

$ErrorActionPreference = 'Stop'
$ScriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$ConfigPath = Join-Path $ScriptDir 'config.json'
$StateDir   = $ScriptDir
if ($env:PS5W_STATE_DIR) { $StateDir = $env:PS5W_STATE_DIR }   # cloud runs keep state in a cached folder
$StatePath  = Join-Path $StateDir 'state.json'
$OnWindows  = ($env:OS -eq 'Windows_NT')
$InCloud    = ($env:GITHUB_ACTIONS -eq 'true')
$LogPath    = Join-Path $ScriptDir 'log.txt'
$StatusPath = Join-Path $ScriptDir 'status.txt'

# ============================================================================
#  What to check
# ============================================================================

# Product pages read through the page's machine-readable stock flag
# (schema.org availability). If a page redirects away from the product,
# the retailer is not selling it right now.
$ProductPages = @(
    @{ Retailer = 'Elgiganten'; Country = 'DK'; Id = '994805';  Url = 'https://www.elgiganten.dk/product/gaming/spillekonsol-tilbehor/playstation/ps5-playstation-5-spillekonsol/playstation-5-pro-2tb-spillekonsol/994805' }
    @{ Retailer = 'Elgiganten'; Country = 'DK'; Id = '830990';  Url = 'https://www.elgiganten.dk/product/gaming/spillekonsol-tilbehor/playstation/ps5-playstation-5-spillekonsol/playstation-5-pro-2tb/830990' }
    @{ Retailer = 'Elgiganten'; Country = 'SE'; Id = '994805';  Url = 'https://www.elgiganten.se/product/gaming/spelkonsoler-tillbehor/playstation/ps5-playstation-5-konsol/playstation-5-pro-2tb-spelkonsol/994805' }
    @{ Retailer = 'Elgiganten'; Country = 'SE'; Id = '830990';  Url = 'https://www.elgiganten.se/product/gaming/spelkonsoler-tillbehor/playstation/ps5-playstation-5-konsol/playstation-5-pro-2tb/830990' }
    @{ Retailer = 'Komplett';   Country = 'DK'; Id = '1330018'; Url = 'https://www.komplett.dk/product/1330018/gaming/konsoller/playstation/playstation-5-pro-2tb' }
    @{ Retailer = 'Komplett';   Country = 'SE'; Id = '1330018'; Url = 'https://www.komplett.se/product/1330018/gaming/konsoler/playstation/playstation-5-pro-2tb' }
    @{ Retailer = 'Proshop';    Country = 'DK'; Id = '3301537'; Url = 'https://www.proshop.dk/Spillekonsol/Sony-PlayStation-5-Pro/3301537' }
    @{ Retailer = 'Proshop';    Country = 'DK'; Id = '3417363'; Url = 'https://www.proshop.dk/Spillekonsol/Sony-PlayStation-5-Pro/3417363' }
    @{ Retailer = 'Proshop';    Country = 'SE'; Id = '3301537'; Url = 'https://www.proshop.se/Spelkonsol/Sony-PlayStation-5-Pro/3301537' }
    @{ Retailer = 'Proshop';    Country = 'SE'; Id = '3417363'; Url = 'https://www.proshop.se/Spelkonsol/Sony-PlayStation-5-Pro/3417363' }
    @{ Retailer = 'Coolshop';   Country = 'DK'; Id = '23NV4C';  Url = 'https://www.coolshop.dk/produkt/sony-playstation-5-pro-console/23NV4C/' }
    @{ Retailer = 'Coolshop';   Country = 'SE'; Id = '23NV4C';  Url = 'https://www.coolshop.se/produkt/sony-playstation-5-pro-console/23NV4C/' }
    @{ Retailer = 'Inet';       Country = 'SE'; Id = '6612681'; Url = 'https://www.inet.se/produkt/6612681/sony-playstation-5-pro-2tb-2025' }
    @{ Retailer = 'NetOnNet';   Country = 'SE'; Id = '1060020'; Url = 'https://www.netonnet.se/art/gaming/spel-och-konsol/playstation/playstation-konsol/sony-playstation-5-pro-2025/1060020.14413/' }
    @{ Retailer = 'NetOnNet';   Country = 'SE'; Id = '1035751'; Url = 'https://www.netonnet.se/art/gaming/spel-och-konsol/playstation/sony-playstation-5-pro/1035751.14404/' }
    @{ Retailer = 'NetOnNet';   Country = 'SE'; Id = '1060018'; Url = 'https://www.netonnet.se/art/gaming/spel-och-konsol/playstation/playstation-konsol/sony-playstation-5-pro-ea-fc26-bundle/1060018.14413/' }
)

# Search / category pages scanned for NEW PS5 Pro listings (new bundles,
# re-listed returns). Anything found is checked like a product page above.
$ListingPages = @(
    @{ Retailer = 'Inet';     Country = 'SE'; Base = 'https://www.inet.se';     Url = 'https://www.inet.se/hitta?q=playstation%205%20pro';         Pattern = '(?i)(?<path>/produkt/(?<id>\d+)/(?<slug>[a-z0-9-]+))' }
    @{ Retailer = 'Komplett'; Country = 'DK'; Base = 'https://www.komplett.dk'; Url = 'https://www.komplett.dk/search?q=playstation%205%20pro';  Pattern = '(?i)(?<path>/product/(?<id>\d+)/(?:[a-z0-9-]+/){0,6}(?<slug>[a-z0-9-]+))' }
    @{ Retailer = 'Komplett'; Country = 'SE'; Base = 'https://www.komplett.se'; Url = 'https://www.komplett.se/search?q=playstation%205%20pro';  Pattern = '(?i)(?<path>/product/(?<id>\d+)/(?:[a-z0-9-]+/){0,6}(?<slug>[a-z0-9-]+))' }
    @{ Retailer = 'Proshop';  Country = 'DK'; Base = 'https://www.proshop.dk';  Url = 'https://www.proshop.dk/Spillekonsol';                      Pattern = '(?i)(?<path>/Spillekonsol/(?<slug>[a-z0-9-]+)/(?<id>\d{6,8}))' }
    @{ Retailer = 'Proshop';  Country = 'SE'; Base = 'https://www.proshop.se';  Url = 'https://www.proshop.se/Spelkonsol';                       Pattern = '(?i)(?<path>/Spelkonsol/(?<slug>[a-z0-9-]+)/(?<id>\d{6,8}))' }
)

# Power: their own stock API (online stock count + number of stores with stock).
$PowerSites = @(
    @{ Country = 'DK'; Base = 'https://www.power.dk'; Category = 7790 }
    @{ Country = 'SE'; Base = 'https://www.power.se'; Category = 7416 }
)

# Webhallen: their own search API (online + per-store stock, incl. "Fyndvara" returns).
$WebhallenSearch = 'https://www.webhallen.com/api/productdiscovery/search/playstation%205%20pro?page=1&touchpoint=DESKTOP&totalProductCountSet=false&origin=ORGANIC'

# ============================================================================
#  Helpers
# ============================================================================

function Write-Log([string]$Message) {
    $line = '{0}  {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    try {
        if ((Test-Path $LogPath) -and ((Get-Item $LogPath).Length -gt 2MB)) {
            Move-Item -Path $LogPath -Destination ($LogPath + '.old') -Force
        }
        Add-Content -Path $LogPath -Value $line -Encoding UTF8
    } catch { }
    if ($ShowStatus) { Write-Host $line }
}

$Config = $null
try {
    $Config = Get-Content -Path $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
} catch {
    Write-Log ("Could not read config.json: " + $_.Exception.Message)
    $Config = New-Object psobject
}

function Get-Cfg([string]$Name, $Default) {
    if ($Config.PSObject.Properties.Name -contains $Name -and $null -ne $Config.$Name) { return $Config.$Name }
    return $Default
}

# --- HTTP -------------------------------------------------------------------
try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { }
Add-Type -AssemblyName System.Net.Http
$Handler = New-Object System.Net.Http.HttpClientHandler
$Handler.AllowAutoRedirect = $true
$Handler.MaxAutomaticRedirections = 10
$Handler.AutomaticDecompression = [System.Net.DecompressionMethods]::GZip -bor [System.Net.DecompressionMethods]::Deflate
$Http = New-Object System.Net.Http.HttpClient($Handler)
$Http.Timeout = [TimeSpan]::FromSeconds(20)
$UserAgent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/141.0.0.0 Safari/537.36'
$script:DeadHosts = @{}   # hosts that blocked or timed out during this check - not retried until the next check

function Get-Url([string]$Url, [switch]$Json) {
    $hostName = ([Uri]$Url).Host
    if ($script:DeadHosts.ContainsKey($hostName)) { throw ('skipped - {0} did not answer earlier in this check' -f $hostName) }
    $req = New-Object System.Net.Http.HttpRequestMessage([System.Net.Http.HttpMethod]::Get, $Url)
    $accept = 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'
    if ($Json) { $accept = 'application/json, text/plain, */*' }
    [void]$req.Headers.TryAddWithoutValidation('User-Agent', $UserAgent)
    [void]$req.Headers.TryAddWithoutValidation('Accept', $accept)
    [void]$req.Headers.TryAddWithoutValidation('Accept-Language', 'da-DK,da;q=0.9,sv-SE;q=0.8,sv;q=0.7,en;q=0.6')
    try {
        $resp = $Http.SendAsync($req).GetAwaiter().GetResult()
        $body = $resp.Content.ReadAsStringAsync().GetAwaiter().GetResult()
    } catch {
        $script:DeadHosts[$hostName] = $true
        throw
    }
    $status = [int]$resp.StatusCode
    $final = $resp.RequestMessage.RequestUri.AbsoluteUri
    Start-Sleep -Milliseconds (Get-Random -Minimum 200 -Maximum 700)
    if ($status -ge 400) {
        if ($status -eq 403 -or $status -eq 429 -or $status -ge 500) { $script:DeadHosts[$hostName] = $true }
        throw ('HTTP {0} from {1}' -f $status, $Url)
    }
    return [pscustomobject]@{ Status = $status; FinalUrl = $final; Body = $body }
}

# --- Text helpers -------------------------------------------------------------
function ConvertTo-PlainName([string]$s) {
    # lower-case, strip accents and symbols so "PlayStation(R)5 Pro" and "ps5-pro" compare the same
    if (-not $s) { return '' }
    $t = $s.ToLowerInvariant()
    $t = $t.Replace([string][char]0x00F8, 'o').Replace([string][char]0x00E6, 'ae')
    $t = $t.Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $t.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [Globalization.UnicodeCategory]::NonSpacingMark) { [void]$sb.Append($ch) }
    }
    $t = $sb.ToString()
    $t = [regex]::Replace($t, '[^a-z0-9+]+', ' ')
    return $t.Trim()
}

$AccessoryWords = '\b(cover|covers|konsolcover|konsolcovers|holje|holjen|hoelje|skin|skins|sticker|stand|stander|stall|beslag|vaegbeslag|vaggfaste|mount|dock|docking|oplader|laddningsstation|laddstation|charging|skivenhet|diskdrev|drev|drive|kabel|cable|fodral|taske|vaska|case|kylare|koeler|cooling|ssd|faceplate|kontroll|kontrol|controller)\b'

function Test-IsProConsole([string]$Name) {
    # True for "PlayStation 5 Pro ..." / "PS5 Pro ..." / "Sony PlayStation 5 Pro ..." consoles and bundles,
    # false for accessories such as "PS5 Pro konsolcovers" or "Controller for PS5 Pro".
    $n = ConvertTo-PlainName $Name
    if ($n -notmatch '^(sony\s+)?(playstation|ps)\s*5\s*pro\b') { return $false }
    # "PlayStation 5 Pro + DualSense Edge controller" is a console bundle, so judge only the part before "+"
    $head = ($n -split '\+')[0]
    if ($head -match $AccessoryWords) { return $false }
    return $true
}

function ConvertFrom-JsonText([string]$s) {
    try { return [string](ConvertFrom-Json ('"' + $s + '"')) } catch { return $s }
}

$KnownAvailability = 'InStock|OutOfStock|SoldOut|Discontinued|PreOrder|PreSale|BackOrder|LimitedAvailability|OnlineOnly|InStoreOnly|MadeToOrder|Reserved'

function Get-SchemaAvailability([string]$Html) {
    $vals = New-Object System.Collections.Generic.List[string]
    # 1) schema.org JSON-LD Product block
    foreach ($b in [regex]::Matches($Html, '(?is)<script[^>]*ld\+json[^>]*>(.*?)</script>')) {
        $block = $b.Groups[1].Value
        if ($block -match '"@type"\s*:\s*\[?\s*"Product"') {
            foreach ($m in [regex]::Matches($block, '"availability"\s*:\s*"(?:https?://schema\.org/)?(' + $KnownAvailability + ')"')) {
                $vals.Add($m.Groups[1].Value)
            }
            if ($vals.Count -gt 0) { break }
        }
    }
    # 2) <meta property="product:availability" content="in stock">
    if ($vals.Count -eq 0) {
        $m = [regex]::Match($Html, '(?i)<meta[^>]+property="(?:product|og):availability"[^>]+content="([^"]+)"')
        if (-not $m.Success) { $m = [regex]::Match($Html, '(?i)<meta[^>]+content="([^"]+)"[^>]+property="(?:product|og):availability"') }
        if ($m.Success) {
            $v = $m.Groups[1].Value.ToLowerInvariant().Trim()
            switch -regex ($v) {
                '^(in ?stock|instock|available)$'     { $vals.Add('InStock'); break }
                'pre-?order'                          { $vals.Add('PreOrder'); break }
                'available for order|back-?order'     { $vals.Add('BackOrder'); break }
                'discontinued'                        { $vals.Add('Discontinued'); break }
                'out ?of ?stock|oos|sold ?out'        { $vals.Add('OutOfStock'); break }
            }
        }
    }
    # 3) any schema availability value in the page (some shops embed it outside a script tag)
    if ($vals.Count -eq 0) {
        $m = [regex]::Match($Html, '"availability"\s*:\s*"(?:https?://schema\.org/)?(' + $KnownAvailability + ')"')
        if ($m.Success) { $vals.Add($m.Groups[1].Value) }
    }
    $unique = @($vals.ToArray() | Select-Object -Unique)
    return ,$unique
}

function Get-StatusFromAvailability([string[]]$Values) {
    if (-not $Values -or $Values.Count -eq 0) { return 'unknown' }
    foreach ($v in $Values) { if ($v -in @('InStock', 'LimitedAvailability', 'OnlineOnly')) { return 'in_stock' } }
    if ($Values -contains 'InStoreOnly') { return 'in_store' }
    foreach ($v in $Values) { if ($v -in @('PreOrder', 'PreSale', 'BackOrder', 'MadeToOrder')) { return 'orderable' } }
    return 'out_of_stock'
}

function Get-ProductName([string]$Html) {
    $m = [regex]::Match($Html, '"@type"\s*:\s*"Product"[\s\S]{0,4000}?"name"\s*:\s*"((?:[^"\\]|\\.)+)"')
    if ($m.Success) {
        $n = [System.Net.WebUtility]::HtmlDecode((ConvertFrom-JsonText $m.Groups[1].Value))
        if ((ConvertTo-PlainName $n) -match '(playstation|ps)\s*5\s*pro') { return $n.Trim() }
    }
    $m = [regex]::Match($Html, '(?is)<title[^>]*>(.*?)</title>')
    if ($m.Success) { return ([System.Net.WebUtility]::HtmlDecode($m.Groups[1].Value) -replace '\s+', ' ').Trim() }
    return $null
}

function Get-ProductPrice([string]$Html) {
    $m = [regex]::Match($Html, '"@type"\s*:\s*"(?:Product|Offer)"[\s\S]{0,6000}?"price"\s*:\s*"?(\d+(?:[.,]\d+)?)')
    if ($m.Success) { return $m.Groups[1].Value }
    return $null
}

# --- Results ----------------------------------------------------------------
$Items = New-Object System.Collections.Generic.List[object]
$SourceOk = @{}       # source name -> $true / error message
$script:CurrentSource = ''

function Add-Item {
    param($Retailer, $Country, $Key, $Name, $Url, $Status, $Detail)
    $Items.Add([pscustomobject]@{
        Retailer = $Retailer; Country = $Country; Key = $Key.ToLowerInvariant(); Name = $Name
        Url = $Url; Status = $Status; Detail = $Detail; Source = $script:CurrentSource
    })
}

function Get-Currency($Country) { if ($Country -eq 'SE') { 'SEK' } else { 'DKK' } }

# ============================================================================
#  Checks
# ============================================================================

function Test-ProductPage($P) {
    $key = '{0}-{1}-{2}' -f $P.Retailer, $P.Country, $P.Id
    $r = Get-Url $P.Url
    if ($r.FinalUrl -notmatch [regex]::Escape($P.Id)) {
        Add-Item $P.Retailer $P.Country $key ('PlayStation 5 Pro (#{0})' -f $P.Id) $P.Url 'not_listed' 'page redirects - not for sale right now'
        return
    }
    $vals = Get-SchemaAvailability $r.Body
    $status = Get-StatusFromAvailability $vals
    $name = Get-ProductName $r.Body
    if (-not $name) { $name = 'PlayStation 5 Pro' }
    $price = Get-ProductPrice $r.Body
    $detail = 'flag: ' + $(if ($vals.Count) { $vals -join ',' } else { 'none found' })
    if ($price) { $detail = ('{0} {1}, {2}' -f $price, (Get-Currency $P.Country), $detail) }
    Add-Item $P.Retailer $P.Country $key $name $r.FinalUrl $status $detail
}

function Find-NewListings($L, [hashtable]$KnownIds) {
    $r = Get-Url $L.Url
    $found = @()
    $seen = @{}
    foreach ($m in [regex]::Matches($r.Body, $L.Pattern)) {
        $id = $m.Groups['id'].Value
        if ($seen.ContainsKey($id)) { continue }
        $seen[$id] = $true
        if ($KnownIds.ContainsKey(('{0}-{1}-{2}' -f $L.Retailer, $L.Country, $id).ToLowerInvariant())) { continue }
        if (-not (Test-IsProConsole $m.Groups['slug'].Value)) { continue }
        $found += @{ Retailer = $L.Retailer; Country = $L.Country; Id = $id; Url = ($L.Base + $m.Groups['path'].Value) }
    }
    return ,$found
}

function Test-Power($S) {
    $urls = @(
        ('{0}/api/v2/productlists?cat={1}&size=60&s=5&from=0&o=false&cd=false' -f $S.Base, $S.Category),
        ('{0}/api/v2/productlists?size=60&q=playstation%205%20pro&s=5&from=0&o=false&cd=false' -f $S.Base)
    )
    $products = @{}
    foreach ($u in $urls) {
        $j = (Get-Url $u -Json).Body | ConvertFrom-Json
        foreach ($p in $j.products) { $products[[string]$p.productId] = $p }
    }
    $count = 0
    foreach ($p in $products.Values) {
        if (-not (Test-IsProConsole $p.title)) { continue }
        if ([string]$p.categoryName -notmatch '(?i)konsol|bundle') { continue }
        $count++
        $web = [int]$p.stockCount; $stores = [int]$p.storesStockCount
        $status = 'out_of_stock'
        if ($web -gt 0) { $status = 'in_stock' } elseif ($stores -gt 0) { $status = 'in_store' }
        $detail = '{0} {1}, online: {2}, stores with stock: {3}' -f $p.price, (Get-Currency $S.Country), $web, $stores
        Add-Item 'Power' $S.Country ('power-{0}-{1}' -f $S.Country, $p.productId) $p.title ($S.Base + $p.url) $status $detail
    }
    if ($count -eq 0) {
        Add-Item 'Power' $S.Country ('power-{0}-none' -f $S.Country) 'PlayStation 5 Pro' ($S.Base + '/search/?q=playstation%205%20pro') 'not_listed' 'no PS5 Pro listed in the console range'
    }
}

function Test-Webhallen {
    $j = (Get-Url $WebhallenSearch -Json).Body | ConvertFrom-Json
    $count = 0
    foreach ($p in $j.products) {
        if (-not (Test-IsProConsole $p.name)) { continue }
        if (([string]$p.categoryTree + ' ' + [string]$p.subTitle) -notmatch '(?i)konsol') { continue }
        $count++
        $web = 0; $stores = 0
        if ($p.stock) {
            $web = [int]$p.stock.web
            foreach ($prop in $p.stock.PSObject.Properties) {
                if ($prop.Name -match '^\d+$' -and [int]$prop.Value -gt 0) { $stores++ }
            }
        }
        $status = 'out_of_stock'
        if ($web -gt 0) { $status = 'in_stock' } elseif ($stores -gt 0) { $status = 'in_store' }
        $name = $p.name
        if ($p.isFyndware) { $name = $name + ' [Fyndvara - returned/open box]' }
        $price = $null; if ($p.price) { $price = $p.price.price }
        $detail = '{0} SEK, online: {1}, stores with stock: {2}' -f $price, $web, $stores
        if ($p.stock -and $p.stock.incoming -and $p.stock.incoming.deliveryDate) {
            $detail += (', incoming {0} ({1})' -f $p.stock.incoming.deliveryDate, $p.stock.incoming.quantity)
        }
        Add-Item 'Webhallen' 'SE' ('webhallen-se-{0}' -f $p.id) $name ('https://www.webhallen.com/se/product/{0}' -f $p.id) $status $detail
    }
    if ($count -eq 0) {
        Add-Item 'Webhallen' 'SE' 'webhallen-se-none' 'PlayStation 5 Pro' 'https://www.webhallen.com/se/search?searchString=playstation%205%20pro' 'not_listed' 'no PS5 Pro found in search'
    }
}

# ============================================================================
#  Notifications
# ============================================================================

function Send-Push([string]$Title, [string]$Message, [string]$Url, [int]$Priority = 4) {
    $topic = [string](Get-Cfg 'ntfyTopic' '')
    if ($env:NTFY_TOPIC) { $topic = $env:NTFY_TOPIC }   # cloud: comes from a GitHub secret
    if (-not $topic) { Write-Log 'No ntfy topic set (config.json or NTFY_TOPIC secret) - skipping phone push'; return }
    $server = ([string](Get-Cfg 'ntfyServer' 'https://ntfy.sh')).TrimEnd('/')
    $payload = [ordered]@{ topic = $topic; title = $Title; message = $Message; priority = $Priority; tags = @('video_game') }
    if ($Url) {
        $payload.click = $Url
        $payload.actions = @(@{ action = 'view'; label = 'Open shop'; url = $Url })
    }
    $json = $payload | ConvertTo-Json -Depth 5 -Compress
    try {
        $req = New-Object System.Net.Http.HttpRequestMessage([System.Net.Http.HttpMethod]::Post, $server)
        $req.Content = New-Object System.Net.Http.StringContent($json, [System.Text.Encoding]::UTF8, 'application/json')
        $resp = $Http.SendAsync($req).GetAwaiter().GetResult()
        if ([int]$resp.StatusCode -ge 300) { Write-Log ('Phone push failed: HTTP {0}' -f [int]$resp.StatusCode) }
        else { Write-Log ('Phone push sent: ' + $Title) }
    } catch {
        Write-Log ('Phone push failed: ' + $_.Exception.Message)
    }
}

function Show-Toast([string]$Title, [string]$Text) {
    if (-not $OnWindows -or $InCloud) { return }
    try {
        $null = [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]
        $tpl = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastText02)
        $nodes = $tpl.GetElementsByTagName('text')
        [void]$nodes.Item(0).AppendChild($tpl.CreateTextNode($Title))
        [void]$nodes.Item(1).AppendChild($tpl.CreateTextNode($Text))
        $toast = [Windows.UI.Notifications.ToastNotification]::new($tpl)
        $appId = '{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe'
        [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($appId).Show($toast)
    } catch {
        Write-Log ('Windows notification not shown: ' + $_.Exception.Message)
    }
}

$StatusLabel = @{
    in_stock = 'IN STOCK'; in_store = 'IN STORES'; orderable = 'ORDERABLE'
    out_of_stock = 'out of stock'; not_listed = 'not listed'; unknown = 'unknown'; error = 'ERROR'
}
$Rank = @{ orderable = 1; in_store = 2; in_stock = 3 }

# ============================================================================
#  Main
# ============================================================================

if ($env:PS5W_TEST -eq '1') { return }   # lets the test harness load the functions only

if ($TestAlert) {
    Send-Push 'PS5 Pro watcher: test alert' 'If you can read this on your phone, alerts are working. You will get a message like this when a PS5 Pro comes in stock in DK or SE.' 'https://ntfy.sh' 3
    Show-Toast 'PS5 Pro watcher: test alert' 'Windows notifications are working.'
    Write-Host 'Test alert sent. Check your phone (ntfy app) and the Windows notification area.'
    exit 0
}

# Only one check at a time
$Mutex = $null
try {
    $Mutex = New-Object System.Threading.Mutex($false, 'Local\PS5ProStockWatcher')
    if (-not $Mutex.WaitOne(0)) { exit 0 }
} catch { $Mutex = $null }

# Load state
$State = @{ items = @{}; failures = @{}; runs = 0 }
if (Test-Path $StatePath) {
    try {
        $raw = Get-Content -Path $StatePath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($raw.runs) { $State.runs = [int]$raw.runs }
        if ($raw.items)    { foreach ($p in $raw.items.PSObject.Properties)    { $State.items[$p.Name]    = @{ status = [string]$p.Value.status; since = [string]$p.Value.since; source = [string]$p.Value.source } } }
        if ($raw.failures) { foreach ($p in $raw.failures.PSObject.Properties) { $State.failures[$p.Name] = @{ count = [int]$p.Value.count; warned = [bool]$p.Value.warned } } }
    } catch {
        Write-Log 'state.json unreadable - starting fresh'
    }
}

$State.runs = [int]$State.runs + 1
$Paused = @{}

function Invoke-Source([string]$Name, [scriptblock]$Block) {
    $script:CurrentSource = $Name
    # A shop that keeps failing is only retried every 10th check (~30 min) to avoid hammering it
    $f = $State.failures[$Name]
    if ($f -and [int]$f.count -ge 20 -and ($State.runs % 10) -ne 0) { $Paused[$Name] = [int]$f.count; return }
    try {
        & $Block
        $SourceOk[$Name] = $true
    } catch {
        $msg = $_.Exception.Message
        if ($_.Exception.InnerException) { $msg = $_.Exception.InnerException.Message }
        $SourceOk[$Name] = $msg
        Write-Log ('{0}: {1}' -f $Name, $msg)
    }
}

$knownIds = @{}
foreach ($p in $ProductPages) { $knownIds[('{0}-{1}-{2}' -f $p.Retailer, $p.Country, $p.Id).ToLowerInvariant()] = $true }

foreach ($p in $ProductPages) {
    $pp = $p
    Invoke-Source ('{0} {1} #{2}' -f $p.Retailer, $p.Country, $p.Id) { Test-ProductPage $pp }
}
foreach ($l in $ListingPages) {
    $ll = $l
    Invoke-Source ('{0} {1} listing' -f $l.Retailer, $l.Country) {
        foreach ($n in (Find-NewListings $ll $knownIds)) {
            $knownIds[('{0}-{1}-{2}' -f $n.Retailer, $n.Country, $n.Id).ToLowerInvariant()] = $true
            Test-ProductPage $n
        }
    }
}
foreach ($s in $PowerSites) {
    $ss = $s
    Invoke-Source ('Power {0}' -f $s.Country) { Test-Power $ss }
}
Invoke-Source 'Webhallen SE' { Test-Webhallen }

# Failed and paused sources show up as ERROR rows
foreach ($name in $SourceOk.Keys) {
    if ($SourceOk[$name] -ne $true) {
        Add-Item ($name -split ' ')[0] ($name -split ' ')[1] ('error-' + ($name -replace '\W+', '-')) $name '' 'error' ([string]$SourceOk[$name])
    }
}
foreach ($name in $Paused.Keys) {
    Add-Item ($name -split ' ')[0] ($name -split ' ')[1] ('error-' + ($name -replace '\W+', '-')) $name '' 'error' ('paused after {0} failed checks - retried every 10th check' -f $Paused[$name])
}

# ---------------------------------------------------------------- alerts ---
$alertStore     = [bool](Get-Cfg 'alertOnStoreOnlyStock' $true)
$alertOrderable = [bool](Get-Cfg 'alertOnOrderable' $true)
$openBrowser    = [bool](Get-Cfg 'openBrowserOnAlert' $true) -and $OnWindows -and -not $InCloud
$now = Get-Date -Format 'yyyy-MM-dd HH:mm'

foreach ($it in $Items) {
    if ($it.Status -in @('error', 'unknown')) { continue }   # keep the previous state on a failed read
    $prev = 'none'
    if ($State.items.ContainsKey($it.Key)) { $prev = $State.items[$it.Key].status }

    $alertable = $Rank.ContainsKey($it.Status)
    if ($it.Status -eq 'in_store'  -and -not $alertStore)     { $alertable = $false }
    if ($it.Status -eq 'orderable' -and -not $alertOrderable) { $alertable = $false }
    $wasAvailable = $Rank.ContainsKey($prev)
    $upgrade = $wasAvailable -and $Rank.ContainsKey($it.Status) -and ($Rank[$it.Status] -gt $Rank[$prev])

    if ($alertable -and ((-not $wasAvailable) -or $upgrade)) {
        $title = 'PS5 Pro {0}: {1} {2}' -f $StatusLabel[$it.Status], $it.Retailer, $it.Country
        $msg = '{0}{1}{2}' -f $it.Name, [Environment]::NewLine, $it.Detail
        $prio = 4; if ($it.Status -eq 'in_stock') { $prio = 5 }
        Write-Log ('ALERT ' + $title + ' - ' + $it.Name + ' - ' + $it.Url)
        Send-Push $title $msg $it.Url $prio
        Show-Toast $title $it.Name
        if ($openBrowser -and $it.Status -ne 'in_store' -and $it.Url) {
            try { Start-Process $it.Url } catch { }
        }
    }
    if ($prev -ne $it.Status) {
        if ($prev -ne 'none') { Write-Log ('{0} {1}: {2} -> {3} ({4})' -f $it.Retailer, $it.Country, $prev, $it.Status, $it.Name) }
        $State.items[$it.Key] = @{ status = $it.Status; since = $now; source = $it.Source }
    }
}

# A listing that disappeared from a shop that answered fine is no longer for sale,
# so a later comeback triggers a fresh alert.
$seenKeys = @{}
foreach ($it in $Items) { $seenKeys[$it.Key] = $true }
foreach ($k in @($State.items.Keys)) {
    $e = $State.items[$k]
    if ($seenKeys.ContainsKey($k) -or -not $e.source) { continue }
    if ($SourceOk.ContainsKey($e.source) -and $SourceOk[$e.source] -eq $true -and $e.status -ne 'not_listed') {
        Write-Log ('{0}: {1} -> not_listed (gone from the shop)' -f $k, $e.status)
        $State.items[$k] = @{ status = 'not_listed'; since = $now; source = $e.source }
    }
}

# ------------------------------------------------- repeated-failure warning ---
$newlyBroken = @()
foreach ($name in $SourceOk.Keys) {
    $f = $State.failures[$name]
    if (-not $f) { $f = @{ count = 0; warned = $false } }
    if ($SourceOk[$name] -eq $true) {
        if ($f.warned) { Write-Log ($name + ' is working again') }
        $f = @{ count = 0; warned = $false }
    } else {
        $f.count = [int]$f.count + 1
        if ($f.count -ge 20 -and -not $f.warned) {
            $newlyBroken += ('{0} ({1})' -f $name, $SourceOk[$name])
            $f.warned = $true
        }
    }
    $State.failures[$name] = $f
}
if ($newlyBroken.Count -gt 0) {
    Send-Push 'PS5 Pro watcher: some shops cannot be checked' ('These checks have failed 20 times in a row - the site may have changed or is blocking the checker. Everything else is still watched.' + [Environment]::NewLine + ($newlyBroken -join [Environment]::NewLine)) '' 3
}

try {
    ([pscustomobject]@{ runs = $State.runs; items = $State.items; failures = $State.failures }) | ConvertTo-Json -Depth 5 | Set-Content -Path $StatePath -Encoding UTF8
} catch { Write-Log ('Could not save state.json: ' + $_.Exception.Message) }

# ---------------------------------------------------------- status file ---
$available = @($Items | Where-Object { $Rank.ContainsKey($_.Status) })
$errors = @($Items | Where-Object { $_.Status -eq 'error' })
$lines = New-Object System.Collections.Generic.List[string]
$lines.Add(('PS5 Pro stock - checked {0}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')))
$lines.Add(('{0} listings checked, {1} available, {2} checks failed' -f $Items.Count, $available.Count, $errors.Count))
$lines.Add('')
$lines.Add(('{0,-11} {1,-3} {2,-13} {3,-58} {4}' -f 'Shop', 'Cty', 'Status', 'Product', 'Details'))
$lines.Add(('-' * 120))
foreach ($it in ($Items | Sort-Object Country, Retailer, Name)) {
    $nm = [string]$it.Name; if ($nm.Length -gt 58) { $nm = $nm.Substring(0, 55) + '...' }
    $lines.Add(('{0,-11} {1,-3} {2,-13} {3,-58} {4}' -f $it.Retailer, $it.Country, $StatusLabel[$it.Status], $nm, $it.Detail))
    if ($Rank.ContainsKey($it.Status)) { $lines.Add(('{0,-29}-> {1}' -f '', $it.Url)) }
}
$lines.Add('')
$lines.Add('Not covered (no readable stock data without a browser): Bilka, Foetex, BR, MediaMarkt SE.')
$lines | Set-Content -Path $StatusPath -Encoding UTF8
if ($ShowStatus) { $lines | ForEach-Object { Write-Host $_ } }

# On GitHub, show the table on the run's summary page
if ($env:GITHUB_STEP_SUMMARY) {
    $md = New-Object System.Collections.Generic.List[string]
    $md.Add(('### PS5 Pro stock - {0} UTC' -f (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm')))
    $md.Add(('{0} listings checked, **{1} available**, {2} checks failed' -f $Items.Count, $available.Count, $errors.Count))
    $md.Add('')
    $md.Add('| Shop | Country | Status | Product | Details |')
    $md.Add('|---|---|---|---|---|')
    foreach ($it in ($Items | Sort-Object Country, Retailer, Name)) {
        $st = $StatusLabel[$it.Status]; if ($Rank.ContainsKey($it.Status)) { $st = '**' + $st + '**' }
        $nm = ([string]$it.Name) -replace '\|', '/'
        if ($it.Url -and $it.Status -ne 'error') { $nm = '[{0}]({1})' -f $nm, $it.Url }
        $md.Add(('| {0} | {1} | {2} | {3} | {4} |' -f $it.Retailer, $it.Country, $st, $nm, (([string]$it.Detail) -replace '\|', '/')))
    }
    try { $md | Add-Content -Path $env:GITHUB_STEP_SUMMARY -Encoding UTF8 } catch { }
}

Write-Log ('Check done: {0} listings, {1} available, {2} failed' -f $Items.Count, $available.Count, $errors.Count)
if ($Mutex) { try { $Mutex.ReleaseMutex() } catch { } }
