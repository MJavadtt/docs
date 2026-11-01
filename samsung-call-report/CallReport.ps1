#Requires -Version 5.1
<#
  Samsung Call Report
  ---------------------------------------------------------------
  - Reads the call log of a connected Samsung/Android phone via ADB
  - Builds Excel reports (work / personal) for the last N days
  - E-mails them automatically
  - Registers a Windows scheduled task so it repeats on its own

  Run by double-clicking CallReport.bat
#>
param([switch]$Scheduled)

$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root
try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch {}
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$LogFile    = Join-Path $Root 'log.txt'
$StateFile  = Join-Path $Root 'state.json'
$SecretFile = Join-Path $Root 'secret.dat'
$ReportDir  = Join-Path $Root 'Reports'
$TaskName   = 'SamsungCallReport'

# ----------------------------------------------------------------- helpers
function Log([string]$msg, [string]$color = 'Gray') {
    $line = '[{0:yyyy-MM-dd HH:mm:ss}] {1}' -f (Get-Date), $msg
    try { Add-Content -Path $LogFile -Value $line -Encoding UTF8 } catch {}
    if (-not $Scheduled) { Write-Host $msg -ForegroundColor $color }
}

function Read-Config {
    $cfg = @{}
    $path = Join-Path $Root 'config.ini'
    if (Test-Path $path) {
        foreach ($raw in (Get-Content $path -Encoding UTF8)) {
            $l = $raw.Trim()
            if ($l -and -not $l.StartsWith('#') -and $l.Contains('=')) {
                $i = $l.IndexOf('=')
                $cfg[$l.Substring(0, $i).Trim()] = $l.Substring($i + 1).Trim()
            }
        }
    }
    $defaults = @{
        ToEmail = 'm.javadtt@yahoo.com'; SenderEmail = 'm.javadtt@yahoo.com'; ManagerEmail = ''
        Days = '90'; SendEveryDays = '30'; ScheduleTime = '12:00'
        WorkKeywords = ''; UnknownNumbersAreWork = 'no'; WirelessAdb = 'yes'
        SmtpHost = ''; SmtpPort = '587'
    }
    foreach ($k in $defaults.Keys) { if (-not $cfg.ContainsKey($k) -or $cfg[$k] -eq '') { $cfg[$k] = $defaults[$k] } }
    return $cfg
}

function Read-State {
    if (Test-Path $StateFile) {
        try { return (Get-Content $StateFile -Raw -Encoding UTF8 | ConvertFrom-Json) } catch {}
    }
    return [pscustomobject]@{ LastSent = ''; PhoneIp = '' }
}
function Save-State($state) { $state | ConvertTo-Json | Set-Content $StateFile -Encoding UTF8 }

function Is-Yes([string]$v) { return @('yes', 'y', 'true', '1', 'بله') -contains $v.Trim().ToLower() }

function Normalize-Number([string]$n) {
    if (-not $n) { return '' }
    $d = $n -replace '\D', ''
    if ($d.Length -gt 10) { $d = $d.Substring($d.Length - 10) }
    return $d
}

# ----------------------------------------------------------------- ADB
function Get-Adb {
    $local = Join-Path $Root 'platform-tools\adb.exe'
    if (Test-Path $local) { return $local }
    $cmd = Get-Command adb.exe -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    Log 'در حال دانلود ابزار ADB از سایت رسمی گوگل (فقط بار اول)...' 'Cyan'
    $zip = Join-Path $env:TEMP 'platform-tools-latest-windows.zip'
    Invoke-WebRequest -Uri 'https://dl.google.com/android/repository/platform-tools-latest-windows.zip' -OutFile $zip -UseBasicParsing
    Expand-Archive -Path $zip -DestinationPath $Root -Force
    Remove-Item $zip -Force -ErrorAction SilentlyContinue
    if (-not (Test-Path $local)) { throw 'دانلود ADB ناموفق بود.' }
    return $local
}

function Invoke-Adb([string[]]$AdbArgs) {
    $ErrorActionPreference = 'Continue'
    $out = & $script:Adb @AdbArgs 2>&1 | ForEach-Object { "$_" }
    return @($out)
}

function Get-DeviceList {
    $list = @()
    foreach ($l in (Invoke-Adb @('devices'))) {
        if ($l -match '^(\S+)\s+(device|unauthorized|offline)\s*$') {
            $list += [pscustomobject]@{ Serial = $Matches[1]; State = $Matches[2] }
        }
    }
    return $list
}

function Wait-Device([int]$timeoutSec) {
    $deadline = (Get-Date).AddSeconds($timeoutSec)
    $toldUnauth = $false; $toldNone = $false
    while ((Get-Date) -lt $deadline) {
        $devs = Get-DeviceList
        $ok = $devs | Where-Object { $_.State -eq 'device' } | Select-Object -First 1
        if ($ok) { return $ok.Serial }
        if (-not $Scheduled) {
            if (($devs | Where-Object { $_.State -eq 'unauthorized' }) -and -not $toldUnauth) {
                Log '>>> روی صفحه گوشی پیغام «Allow USB debugging?» آمده: تیک «Always allow» را بزنید و Allow را لمس کنید.' 'Yellow'
                $toldUnauth = $true
            } elseif (-not $devs -and -not $toldNone) {
                Log '>>> گوشی پیدا نشد. گوشی را با کابل به کامپیوتر وصل کنید و مطمئن شوید «USB debugging» روشن است (راهنما: README.md).' 'Yellow'
                Log '    منتظر اتصال گوشی هستم...' 'Yellow'
                $toldNone = $true
            }
        }
        Start-Sleep -Seconds 2
    }
    return $null
}

function Get-CallLog([string]$serial) {
    $raw = Invoke-Adb @('-s', $serial, 'shell', 'content', 'query', '--uri', 'content://call_log/calls',
                        '--projection', 'number:name:date:duration:type')
    $calls = New-Object System.Collections.Generic.List[object]
    foreach ($l in $raw) {
        if ($l -match '^Row:\s*\d+\s+number=(.*?), name=(.*?), date=(\d+), duration=(\d+), type=(\d+)\s*$') {
            $name = $Matches[2]; if ($name -eq 'NULL') { $name = '' }
            $num = $Matches[1]; if ($num -eq 'NULL') { $num = '' }
            $calls.Add([pscustomobject]@{
                Number = $num; Name = $name; DateMs = [int64]$Matches[3]
                Duration = [int]$Matches[4]; Type = [int]$Matches[5]
            })
        }
    }
    if ($calls.Count -eq 0) {
        Log ('خروجی گوشی: ' + (($raw | Select-Object -First 5) -join ' | ')) 'Red'
    }
    return $calls
}

function Get-ContactNames([string]$serial) {
    # best effort: fresher names than the cached ones in the call log
    $map = @{}
    $raw = Invoke-Adb @('-s', $serial, 'shell', 'content', 'query', '--uri', 'content://com.android.contacts/data/phones',
                        '--projection', 'display_name:data1')
    foreach ($l in $raw) {
        if ($l -match '^Row:\s*\d+\s+display_name=(.*?), data1=(.*?)\s*$') {
            $k = Normalize-Number $Matches[2]
            if ($k -and $Matches[1] -ne 'NULL' -and -not $map.ContainsKey($k)) { $map[$k] = $Matches[1] }
        }
    }
    return $map
}

function Enable-WirelessAdb([string]$serial, $state) {
    if ($serial -match ':\d+$') { return }   # already wireless
    $ipOut = Invoke-Adb @('-s', $serial, 'shell', 'ip', '-f', 'inet', 'addr', 'show', 'wlan0')
    $ip = $null
    foreach ($l in $ipOut) { if ($l -match 'inet\s+(\d+\.\d+\.\d+\.\d+)') { $ip = $Matches[1] } }
    if (-not $ip) { Log 'وای‌فای گوشی روشن نیست؛ اتصال بی‌سیم فعال نشد (اشکالی ندارد).' ; return }
    Invoke-Adb @('-s', $serial, 'tcpip', '5555') | Out-Null
    $state.PhoneIp = $ip
    Log "اتصال بی‌سیم فعال شد (IP گوشی: $ip). تا وقتی گوشی ری‌استارت نشده، ارسال خودکار بدون کابل هم کار می‌کند." 'Green'
}

# ----------------------------------------------------------------- XLSX writer (no Excel needed)
function Get-ColName([int]$i) {
    $s = ''; $n = $i + 1
    while ($n -gt 0) { $m = ($n - 1) % 26; $s = "$([char](65 + $m))$s"; $n = [int][math]::Floor(($n - 1) / 26) }
    return $s
}
function Esc([object]$v) {
    if ($null -eq $v) { return '' }
    $s = "$v" -replace '[\x00-\x08\x0B\x0C\x0E-\x1F]', ''
    return [Security.SecurityElement]::Escape($s)
}

function Build-SheetXml($sheet) {
    $sb = New-Object Text.StringBuilder
    $cols = $sheet.Headers.Count
    $rowCount = $sheet.Rows.Count + 1
    $lastCol = Get-ColName ($cols - 1)
    [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sb.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
    [void]$sb.Append("<dimension ref=`"A1:$lastCol$rowCount`"/>")
    [void]$sb.Append('<sheetViews><sheetView workbookViewId="0" rightToLeft="1"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')
    [void]$sb.Append('<cols>')
    for ($c = 0; $c -lt $cols; $c++) {
        $w = 16; if ($sheet.Widths -and $c -lt $sheet.Widths.Count) { $w = $sheet.Widths[$c] }
        [void]$sb.Append("<col min=`"$($c+1)`" max=`"$($c+1)`" width=`"$w`" customWidth=`"1`"/>")
    }
    [void]$sb.Append('</cols><sheetData>')
    $all = @(, $sheet.Headers) + @($sheet.Rows)
    for ($r = 0; $r -lt $all.Count; $r++) {
        $rn = $r + 1
        [void]$sb.Append("<row r=`"$rn`">")
        $row = $all[$r]
        for ($c = 0; $c -lt $row.Count; $c++) {
            $ref = (Get-ColName $c) + $rn
            $v = $row[$c]
            $style = ''; if ($r -eq 0) { $style = ' s="1"' }
            if ($r -gt 0 -and ($v -is [int] -or $v -is [long] -or $v -is [double] -or $v -is [decimal])) {
                $num = ([double]$v).ToString([Globalization.CultureInfo]::InvariantCulture)
                [void]$sb.Append("<c r=`"$ref`"$style><v>$num</v></c>")
            } else {
                [void]$sb.Append("<c r=`"$ref`" t=`"inlineStr`"$style><is><t xml:space=`"preserve`">$(Esc $v)</t></is></c>")
            }
        }
        [void]$sb.Append('</row>')
    }
    [void]$sb.Append('</sheetData>')
    if ($sheet.Rows.Count -gt 0 -and -not $sheet.NoFilter) { [void]$sb.Append("<autoFilter ref=`"A1:$lastCol$rowCount`"/>") }
    [void]$sb.Append('</worksheet>')
    return $sb.ToString()
}

function New-Xlsx([string]$path, [object[]]$sheets) {
    Add-Type -AssemblyName System.IO.Compression
    if (Test-Path $path) { Remove-Item $path -Force }
    $utf8 = New-Object Text.UTF8Encoding($false)
    $fs = [IO.File]::Open($path, [IO.FileMode]::CreateNew)
    $zip = New-Object IO.Compression.ZipArchive($fs, [IO.Compression.ZipArchiveMode]::Create)
    $add = {
        param($name, $text)
        $e = $zip.CreateEntry($name)
        $w = New-Object IO.StreamWriter($e.Open(), $utf8)
        $w.Write($text); $w.Dispose()
    }
    $ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    $wbSheets = ''; $wbRels = ''
    for ($i = 0; $i -lt $sheets.Count; $i++) {
        $n = $i + 1
        $ct += "<Override PartName=`"/xl/worksheets/sheet$n.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>"
        $wbSheets += "<sheet name=`"$(Esc $sheets[$i].Name)`" sheetId=`"$n`" r:id=`"rId$n`"/>"
        $wbRels += "<Relationship Id=`"rId$n`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`" Target=`"worksheets/sheet$n.xml`"/>"
    }
    $ct += '</Types>'
    $defNames = ''
    for ($i = 0; $i -lt $sheets.Count; $i++) {
        $sh = $sheets[$i]
        if ($sh.Rows.Count -gt 0 -and -not $sh.NoFilter) {
            $sn = $sh.Name -replace "'", "''"
            $ref = '$A$1:$' + (Get-ColName ($sh.Headers.Count - 1)) + '$' + ($sh.Rows.Count + 1)
            $defNames += "<definedName name=`"_xlnm._FilterDatabase`" localSheetId=`"$i`" hidden=`"1`">$(Esc ("'" + $sn + "'!" + $ref))</definedName>"
        }
    }
    if ($defNames) { $defNames = "<definedNames>$defNames</definedNames>" }
    $stylesId = $sheets.Count + 1
    $wbRels += "<Relationship Id=`"rId$stylesId`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles`" Target=`"styles.xml`"/>"

    & $add '[Content_Types].xml' $ct
    & $add '_rels/.rels' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    & $add 'xl/workbook.xml' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' + $wbSheets + '</sheets>' + $defNames + '</workbook>')
    & $add 'xl/_rels/workbook.xml.rels' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + $wbRels + '</Relationships>')
    & $add 'xl/styles.xml' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Tahoma"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Tahoma"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF2F5597"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'
    for ($i = 0; $i -lt $sheets.Count; $i++) {
        & $add "xl/worksheets/sheet$($i+1).xml" (Build-SheetXml $sheets[$i])
    }
    $zip.Dispose(); $fs.Dispose()
}

# ----------------------------------------------------------------- report building
$TypeNames = @{ 1 = 'ورودی'; 2 = 'خروجی'; 3 = 'بی‌پاسخ (از دست رفته)'; 4 = 'پیغام صوتی'; 5 = 'رد شده'; 6 = 'مسدود شده'; 7 = 'پاسخ در دستگاه دیگر' }
$WeekDays  = @('یکشنبه', 'دوشنبه', 'سه‌شنبه', 'چهارشنبه', 'پنجشنبه', 'جمعه', 'شنبه')
$Persian   = New-Object Globalization.PersianCalendar

function To-Jalali([datetime]$d) { return '{0:0000}/{1:00}/{2:00}' -f $Persian.GetYear($d), $Persian.GetMonth($d), $Persian.GetDayOfMonth($d) }
function Fmt-Duration([int]$s) { return '{0}:{1:00}:{2:00}' -f [math]::Floor($s / 3600), [math]::Floor(($s % 3600) / 60), ($s % 60) }

function Load-WorkList {
    $nums = @{}; $names = @()
    $p = Join-Path $Root 'work_contacts.txt'
    if (Test-Path $p) {
        foreach ($raw in (Get-Content $p -Encoding UTF8)) {
            $l = $raw.Trim()
            if (-not $l -or $l.StartsWith('#')) { continue }
            $n = Normalize-Number $l
            if ($n.Length -ge 7) { $nums[$n] = $true } else { $names += $l }
        }
    }
    return @{ Numbers = $nums; Names = $names }
}

function Classify($call, $workList, $keywords, [bool]$unknownIsWork) {
    $n = Normalize-Number $call.Number
    if ($n -and $workList.Numbers.ContainsKey($n)) { return 'کاری' }
    foreach ($k in ($workList.Names + $keywords)) {
        if ($k -and $call.Name -and $call.Name.IndexOf($k, [StringComparison]::OrdinalIgnoreCase) -ge 0) { return 'کاری' }
    }
    if (-not $call.Name -and $unknownIsWork) { return 'کاری' }
    return 'شخصی'
}

$CallHeaders = @('ردیف', 'تاریخ (شمسی)', 'تاریخ (میلادی)', 'روز هفته', 'ساعت', 'نام مخاطب', 'شماره', 'نوع تماس / تماس‌گیرنده', 'مدت (ساعت:دقیقه:ثانیه)', 'مدت (ثانیه)', 'دسته')
$CallWidths  = @(7, 14, 14, 11, 10, 26, 18, 24, 18, 12, 9)

function To-Rows($calls) {
    $rows = New-Object System.Collections.Generic.List[object]
    $i = 0
    foreach ($c in $calls) {
        $i++
        $who = $TypeNames[$c.Type]; if (-not $who) { $who = "نوع $($c.Type)" }
        if ($c.Type -eq 1) { $who = 'ورودی (مخاطب زنگ زده)' } elseif ($c.Type -eq 2) { $who = 'خروجی (شما زنگ زده‌اید)' }
        $name = $c.Name; if (-not $name) { $name = '(ناشناس)' }
        $rows.Add(@([int]$i, (To-Jalali $c.Date), $c.Date.ToString('yyyy-MM-dd'), $WeekDays[[int]$c.Date.DayOfWeek],
                    $c.Date.ToString('HH:mm'), $name, $c.Number, $who, (Fmt-Duration $c.Duration), [int]$c.Duration, $c.Category))
    }
    return , $rows
}

function Contact-Summary($calls) {
    $rows = New-Object System.Collections.Generic.List[object]
    $groups = $calls | Group-Object { $k = Normalize-Number $_.Number; if (-not $k) { $k = $_.Number }; $k } |
              Sort-Object { ($_.Group | Measure-Object Duration -Sum).Sum } -Descending
    foreach ($g in $groups) {
        $first = $g.Group[0]
        $name = ($g.Group | Where-Object { $_.Name } | Select-Object -First 1).Name; if (-not $name) { $name = '(ناشناس)' }
        $sum = [int](($g.Group | Measure-Object Duration -Sum).Sum)
        $last = ($g.Group | Sort-Object DateMs -Descending | Select-Object -First 1).Date
        $rows.Add(@($name, $first.Number, [int]$g.Count,
                    [int]@($g.Group | Where-Object { $_.Type -eq 1 }).Count,
                    [int]@($g.Group | Where-Object { $_.Type -eq 2 }).Count,
                    [int]@($g.Group | Where-Object { $_.Type -eq 3 -or $_.Type -eq 5 }).Count,
                    (Fmt-Duration $sum), [math]::Round($sum / 60.0, 1), ((To-Jalali $last) + ' ' + $last.ToString('HH:mm'))))
    }
    return , $rows
}

function Stats($calls) {
    $sum = [int](($calls | Measure-Object Duration -Sum).Sum)
    return [pscustomobject]@{
        Total    = [int]@($calls).Count
        In       = [int]@($calls | Where-Object { $_.Type -eq 1 }).Count
        Out      = [int]@($calls | Where-Object { $_.Type -eq 2 }).Count
        Missed   = [int]@($calls | Where-Object { $_.Type -eq 3 -or $_.Type -eq 5 }).Count
        Seconds  = $sum
        Contacts = [int]@($calls | ForEach-Object { Normalize-Number $_.Number } | Sort-Object -Unique).Count
    }
}

# ----------------------------------------------------------------- e-mail
function Get-Password {
    if (Test-Path $SecretFile) {
        try {
            $sec = (Get-Content $SecretFile -Raw).Trim() | ConvertTo-SecureString
            return (New-Object Net.NetworkCredential('', $sec)).Password
        } catch { Remove-Item $SecretFile -Force }
    }
    if ($Scheduled) { throw 'رمز ایمیل ذخیره نشده؛ یک‌بار CallReport.bat را دستی اجرا کنید.' }
    Write-Host ''
    Write-Host '=================================================================' -ForegroundColor Cyan
    Write-Host ' فقط همین یک بار: «App Password» یاهو را وارد کنید' -ForegroundColor Cyan
    Write-Host ' (رمز معمولی یاهو کار نمی‌کند. ساخت آن: login.yahoo.com > Account Security' -ForegroundColor Cyan
    Write-Host '  > Generate app password . راهنمای کامل در README.md)' -ForegroundColor Cyan
    Write-Host ' رمز به صورت رمزنگاری‌شده و فقط برای کاربر ویندوز شما ذخیره می‌شود.' -ForegroundColor Cyan
    Write-Host '=================================================================' -ForegroundColor Cyan
    $sec = Read-Host -AsSecureString 'App Password'
    $plain = ((New-Object Net.NetworkCredential('', $sec)).Password) -replace '\s', ''
    ConvertTo-SecureString $plain -AsPlainText -Force | ConvertFrom-SecureString | Set-Content $SecretFile
    return $plain
}

function Send-Mail($cfg, [string]$to, [string]$subject, [string]$html, [string[]]$files) {
    $smtpHost = $cfg.SmtpHost
    if (-not $smtpHost) {
        $domain = ($cfg.SenderEmail -split '@')[-1].ToLower()
        if ($domain -match 'yahoo|ymail|rocketmail') { $smtpHost = 'smtp.mail.yahoo.com' }
        elseif ($domain -match 'gmail|googlemail') { $smtpHost = 'smtp.gmail.com' }
        elseif ($domain -match 'outlook|hotmail|live') { $smtpHost = 'smtp-mail.outlook.com' }
        else { $smtpHost = "smtp.$domain" }
    }
    $msg = New-Object Net.Mail.MailMessage
    $msg.From = New-Object Net.Mail.MailAddress($cfg.SenderEmail)
    foreach ($t in ($to -split '[,;]')) { if ($t.Trim()) { $msg.To.Add($t.Trim()) } }
    $msg.Subject = $subject; $msg.SubjectEncoding = [Text.Encoding]::UTF8
    $msg.Body = $html; $msg.BodyEncoding = [Text.Encoding]::UTF8; $msg.IsBodyHtml = $true
    foreach ($f in $files) { $msg.Attachments.Add((New-Object Net.Mail.Attachment($f))) }
    $smtp = New-Object Net.Mail.SmtpClient($smtpHost, [int]$cfg.SmtpPort)
    $smtp.EnableSsl = $true
    $smtp.Credentials = New-Object Net.NetworkCredential($cfg.SenderEmail, $script:Password)
    try { $smtp.Send($msg) } finally { $msg.Dispose(); $smtp.Dispose() }
}

function Register-Task($cfg) {
    try {
        $action  = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$PSCommandPath`" -Scheduled"
        $trigger = New-ScheduledTaskTrigger -Daily -At $cfg.ScheduleTime
        $set     = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
        Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $set -Force | Out-Null
        Log "ارسال خودکار فعال شد: هر $($cfg.SendEveryDays) روز یک‌بار گزارش جدید ایمیل می‌شود." 'Green'
    } catch {
        Log "ثبت زمان‌بندی خودکار ویندوز ناموفق بود: $($_.Exception.Message)" 'Red'
    }
}

# ================================================================= MAIN
try {
    $cfg = Read-Config
    $state = Read-State
    Log '=== Samsung Call Report ===' 'Cyan'

    if ($Scheduled -and $state.LastSent) {
        $since = (Get-Date) - [datetime]::Parse($state.LastSent, [Globalization.CultureInfo]::InvariantCulture)
        if ($since.TotalDays -lt [double]$cfg.SendEveryDays) { Log 'هنوز زمان گزارش بعدی نرسیده.'; exit 0 }
    }

    $script:Adb = Get-Adb
    Invoke-Adb @('start-server') | Out-Null

    $serial = $null
    if ($Scheduled) {
        $serial = Wait-Device 5
        if (-not $serial -and $state.PhoneIp) {
            Invoke-Adb @('connect', "$($state.PhoneIp):5555") | Out-Null
            $serial = Wait-Device 15
        }
        if (-not $serial) { Log 'گوشی در دسترس نبود؛ فردا دوباره تلاش می‌شود.'; exit 0 }
    } else {
        $serial = Wait-Device 300
        if (-not $serial) { throw 'گوشی در ۵ دقیقه متصل نشد. راهنمای README.md را ببینید و دوباره اجرا کنید.' }
    }
    Log "گوشی متصل است ($serial). در حال خواندن تاریخچه تماس‌ها..." 'Green'

    $calls = Get-CallLog $serial
    if ($calls.Count -eq 0) { throw 'تاریخچه تماسی از گوشی خوانده نشد.' }
    $contacts = @{}
    try { $contacts = Get-ContactNames $serial } catch {}

    $days = [int]$cfg.Days
    $to   = Get-Date
    $from = $to.AddDays(-$days)
    $cutoff = ([DateTimeOffset]$from).ToUnixTimeMilliseconds()

    $workList = Load-WorkList
    $keywords = @($cfg.WorkKeywords -split '[,،]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    $unknownIsWork = Is-Yes $cfg.UnknownNumbersAreWork

    $period = @($calls | Where-Object { $_.DateMs -ge $cutoff } | Sort-Object DateMs -Descending)
    foreach ($c in $period) {
        $k = Normalize-Number $c.Number
        if ($k -and $contacts.ContainsKey($k)) { $c.Name = $contacts[$k] }
        $c | Add-Member -NotePropertyName Date -NotePropertyValue ([DateTimeOffset]::FromUnixTimeMilliseconds($c.DateMs).LocalDateTime) -Force
        $c | Add-Member -NotePropertyName Category -NotePropertyValue (Classify $c $workList $keywords $unknownIsWork) -Force
    }
    $work     = @($period | Where-Object { $_.Category -eq 'کاری' })
    $personal = @($period | Where-Object { $_.Category -eq 'شخصی' })
    Log "تعداد تماس‌ها در $days روز گذشته: $($period.Count)  (کاری: $($work.Count) | شخصی: $($personal.Count))" 'Green'

    # ---- Excel files
    if (-not (Test-Path $ReportDir)) { New-Item -ItemType Directory -Path $ReportDir | Out-Null }
    $stamp = $to.ToString('yyyy-MM-dd')
    $rangeFa = "از $(To-Jalali $from) تا $(To-Jalali $to)"
    $ws = Stats $work; $ps = Stats $personal

    $summaryRows = New-Object System.Collections.Generic.List[object]
    $summaryRows.Add(@('بازه گزارش', $rangeFa))
    $summaryRows.Add(@('تعداد کل تماس‌های کاری', $ws.Total))
    $summaryRows.Add(@('تماس‌های ورودی', $ws.In))
    $summaryRows.Add(@('تماس‌های خروجی', $ws.Out))
    $summaryRows.Add(@('بی‌پاسخ / رد شده', $ws.Missed))
    $summaryRows.Add(@('مجموع زمان مکالمه (ساعت:دقیقه:ثانیه)', (Fmt-Duration $ws.Seconds)))
    $summaryRows.Add(@('مجموع زمان مکالمه (دقیقه)', [math]::Round($ws.Seconds / 60.0, 1)))
    $summaryRows.Add(@('تعداد مخاطبین کاری', $ws.Contacts))
    $summaryRows.Add(@('تاریخ تهیه گزارش', ((To-Jalali $to) + ' ' + $to.ToString('HH:mm'))))

    $contactHeaders = @('نام مخاطب', 'شماره', 'تعداد تماس', 'ورودی', 'خروجی', 'بی‌پاسخ/رد شده', 'مجموع مدت', 'مجموع (دقیقه)', 'آخرین تماس')
    $contactWidths  = @(26, 18, 11, 9, 9, 14, 14, 13, 20)

    $workFile = Join-Path $ReportDir "CallReport_Work_$stamp.xlsx"
    New-Xlsx $workFile @(
        @{ Name = 'خلاصه گزارش'; Headers = @('شرح', 'مقدار'); Rows = $summaryRows; Widths = @(38, 30); NoFilter = $true },
        @{ Name = 'تماس‌های کاری'; Headers = $CallHeaders; Rows = (To-Rows $work); Widths = $CallWidths },
        @{ Name = 'مخاطبین کاری'; Headers = $contactHeaders; Rows = (Contact-Summary $work); Widths = $contactWidths }
    )
    $allFile = Join-Path $ReportDir "CallReport_All_$stamp.xlsx"
    New-Xlsx $allFile @(
        @{ Name = 'همه تماس‌ها'; Headers = $CallHeaders; Rows = (To-Rows $period); Widths = $CallWidths },
        @{ Name = 'تماس‌های شخصی'; Headers = $CallHeaders; Rows = (To-Rows $personal); Widths = $CallWidths },
        @{ Name = 'همه مخاطبین'; Headers = ($contactHeaders); Rows = (Contact-Summary $period); Widths = $contactWidths }
    )
    Log "فایل‌های اکسل ساخته شد: $ReportDir" 'Green'

    # ---- e-mail
    $script:Password = Get-Password
    $td = 'style="padding:6px 12px;border:1px solid #ccc"'
    $html = @"
<div dir="rtl" style="font-family:Tahoma,Arial;font-size:14px">
<h3 style="color:#2F5597">گزارش فعالیت تماس‌های کاری</h3>
<p>بازه: $rangeFa ($days روز)</p>
<table style="border-collapse:collapse">
<tr><td $td>تماس‌های کاری</td><td $td><b>$($ws.Total)</b></td><td $td>تماس‌های شخصی</td><td $td>$($ps.Total)</td></tr>
<tr><td $td>ورودی</td><td $td>$($ws.In)</td><td $td>خروجی</td><td $td>$($ws.Out)</td></tr>
<tr><td $td>بی‌پاسخ / رد شده</td><td $td>$($ws.Missed)</td><td $td>مخاطبین کاری</td><td $td>$($ws.Contacts)</td></tr>
<tr><td $td>مجموع مکالمه کاری</td><td $td colspan="3">$(Fmt-Duration $ws.Seconds)</td></tr>
</table>
<p>فایل <b>CallReport_Work</b> مخصوص ارائه به مدیر است (فقط تماس‌های کاری).<br/>
فایل <b>CallReport_All</b> شامل همه تماس‌ها (کاری و شخصی) فقط برای خودتان است.</p>
</div>
"@
    Send-Mail $cfg $cfg.ToEmail "گزارش تماس‌ها - $rangeFa" $html @($workFile, $allFile)
    Log "ایمیل گزارش به $($cfg.ToEmail) ارسال شد." 'Green'

    if ($cfg.ManagerEmail) {
        $mhtml = $html -replace '<p>فایل <b>CallReport_Work</b>[\s\S]*?</p>', '<p>فایل پیوست شامل جزئیات تماس‌های کاری است.</p>' `
                       -replace '<td [^>]*>تماس‌های شخصی</td><td [^>]*>\d+</td>', ''
        Send-Mail $cfg $cfg.ManagerEmail "گزارش فعالیت تماس‌های کاری - $rangeFa" $mhtml @($workFile)
        Log "گزارش کاری برای مدیر ($($cfg.ManagerEmail)) ارسال شد." 'Green'
    }

    $state.LastSent = (Get-Date).ToString('o', [Globalization.CultureInfo]::InvariantCulture)
    if (Is-Yes $cfg.WirelessAdb) { try { Enable-WirelessAdb $serial $state } catch {} }
    Save-State $state

    if (-not $Scheduled) {
        Register-Task $cfg
        Log '' ; Log 'تمام شد! ✔  گزارش ایمیل شد و پوشه گزارش‌ها باز می‌شود.' 'Green'
        Invoke-Item $ReportDir
    }
} catch {
    $m = $_.Exception.Message
    if ($m -match 'authentication|5\.7\.\d|535|Authentication') {
        if (Test-Path $SecretFile) { Remove-Item $SecretFile -Force }
        $m = "ورود به ایمیل ناموفق بود (App Password اشتباه است). دوباره اجرا کنید و App Password صحیح را وارد کنید. [$m]"
    }
    Log "خطا: $m" 'Red'
    exit 1
}
