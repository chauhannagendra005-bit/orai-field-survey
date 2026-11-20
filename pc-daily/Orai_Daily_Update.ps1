<#
  Orai Fieldwork App - daily Billed / Unbilled update            (Developed by Nagendra Singh)

  Runs every day at 9:30 AM (Windows Task Scheduler, retried every 30 minutes till 1:30 PM).
  1. Finds the newest complete pair in this folder:
        BILLED_DVVNL_DIV233422_DDMMYYYY.csv.gz  +  UNBILLED_DVVNL_DIV233422_DDMMYYYY.csv.gz
  2. Builds a small change file for the app's active master: per consumer the
     last bill date, MR source, bill basis, units, reading and payments
     (last pay date / amount from both reports). For BILLED consumers the
     Total Outstanding becomes the report's AMOUNT_PAYABLE.
  3. Uploads it to the Orai server. Phones load it on their next sync, and
     PROGRESS > Billing shows this month's billed / unbilled.
  Nothing is uploaded twice: the pair already sent for the current master is skipped.

  First time: double-click Setup_Orai_Daily.bat (saves the admin login, schedules the job).
  Log: orai_update.log in this folder.
#>
param(
  [switch]$Setup,
  [switch]$Force,
  [switch]$DryRun,
  [string]$Folder = $PSScriptRoot
)

$ErrorActionPreference = 'Stop'
$ServerUrl  = 'https://script.google.com/macros/s/AKfycbzM4cldusGnSdXWefQLi4R2c5qwag2n-wOL6O2Jevon8CaHJ8bze-twypQkOvkmMu92Gg/exec'
$ConfigFile = Join-Path $Folder 'orai_update_config.json'
$StateFile  = Join-Path $Folder 'orai_update_state.json'
$LogFile    = Join-Path $Folder 'orai_update.log'
$TaskName   = 'Orai Fieldwork Daily Update'

function Log([string]$msg) {
  $line = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '  ' + $msg
  Write-Host $line
  try { Add-Content -Path $LogFile -Value $line -Encoding UTF8 } catch {}
}

try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

function Call-Server([hashtable]$body, [int]$timeoutSec = 300) {
  $json  = $body | ConvertTo-Json -Compress -Depth 5
  $bytes = [Text.Encoding]::UTF8.GetBytes($json)
  $r = Invoke-RestMethod -Uri $ServerUrl -Method Post -Body $bytes -ContentType 'text/plain; charset=utf-8' `
         -TimeoutSec $timeoutSec -MaximumRedirection 5
  if ($r -is [string]) { $r = $r | ConvertFrom-Json }
  return $r
}

function Get-Login {
  if (-not (Test-Path $ConfigFile)) { throw 'No saved login. Double-click Setup_Orai_Daily.bat first.' }
  $cfg = Get-Content $ConfigFile -Raw | ConvertFrom-Json
  $sec = ConvertTo-SecureString $cfg.password
  $ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec)
  try { $pw = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr) }
  finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr) }
  $r = Call-Server @{ action = 'login'; user_id = $cfg.user_id; password = $pw } 120
  if (-not $r.success) { throw ('Login failed: ' + $r.message + '  (run Setup_Orai_Daily.bat again if the password changed)') }
  if (-not $r.is_admin) { throw 'The saved user is not an admin.' }
  return $r.token
}

# ------------------------------------------------------------------ setup
if ($Setup) {
  Write-Host ''
  Write-Host 'Orai Fieldwork App - daily Billed / Unbilled update setup' -ForegroundColor Cyan
  Write-Host 'Enter the admin login of the Orai Fieldwork App (it is saved encrypted for this Windows user only).'
  $uid = Read-Host 'Admin User ID [admin]'
  if (-not $uid) { $uid = 'admin' }
  $sec = Read-Host 'Admin Password' -AsSecureString
  $ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec)
  try { $plain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr) }
  finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr) }
  $r = Call-Server @{ action = 'login'; user_id = $uid; password = $plain } 120
  if (-not $r.success) { Write-Host ('Login failed: ' + $r.message) -ForegroundColor Red; exit 1 }
  if (-not $r.is_admin) { Write-Host 'This user is not an admin.' -ForegroundColor Red; exit 1 }
  @{ user_id = $uid; password = ($sec | ConvertFrom-SecureString) } | ConvertTo-Json | Set-Content $ConfigFile -Encoding UTF8
  Write-Host 'Login OK and saved.' -ForegroundColor Green

  $script = Join-Path $Folder 'Orai_Daily_Update.ps1'
  $action = New-ScheduledTaskAction -Execute 'powershell.exe' `
              -Argument ('-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + $script + '"') `
              -WorkingDirectory $Folder
  $trigger = New-ScheduledTaskTrigger -Daily -At '09:30'
  $rep = New-ScheduledTaskTrigger -Once -At '09:30' -RepetitionInterval (New-TimeSpan -Minutes 30) -RepetitionDuration (New-TimeSpan -Hours 4)
  $trigger.Repetition = $rep.Repetition
  $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
                -ExecutionTimeLimit (New-TimeSpan -Minutes 30)
  Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings `
    -Description 'Uploads the daily BILLED/UNBILLED update to the Orai Fieldwork App (9:30 AM, retries till 1:30 PM).' -Force | Out-Null
  Write-Host ('Scheduled: "' + $TaskName + '" every day at 9:30 AM (retries every 30 min till 1:30 PM).') -ForegroundColor Green
  Write-Host 'Running it once now...'
  Log 'Setup done; first run follows.'
  & $script -Folder $Folder
  exit 0
}

# ------------------------------------------------------------------ helpers
$MON = @{ JAN=1; FEB=2; MAR=3; APR=4; MAY=5; JUN=6; JUL=7; AUG=8; SEP=9; OCT=10; NOV=11; DEC=12 }
function To-Ymd([string]$v) {
  if (-not $v) { return '' }
  $m = [regex]::Match($v, '^\s*(\d{1,2})-([A-Za-z]{3})-(\d{2,4})')
  if (-not $m.Success) { return '' }
  $mo = $MON[$m.Groups[2].Value.ToUpper()]
  if (-not $mo) { return '' }
  $y = [int]$m.Groups[3].Value
  if ($y -lt 100) { $y += 2000 }
  return ('{0:D4}-{1:D2}-{2:D2}' -f $y, $mo, [int]$m.Groups[1].Value)
}
function J([string]$s) {
  if ($null -eq $s -or $s -eq '') { return '""' }
  if ($s -notmatch '["\\\x00-\x1f]') { return '"' + $s + '"' }        # fast path: nothing to escape
  $t = $s.Replace('\', '\\').Replace('"', '\"')
  $t = [regex]::Replace($t, '[\x00-\x1f]', { param($m) '\u{0:x4}' -f [int][char]$m.Value })
  return '"' + $t + '"'
}
function Num([string]$s) {
  $s = ($s + '').Trim()
  if ($s -eq '') { return '""' }
  $d = 0.0
  if ([double]::TryParse($s, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$d)) {
    return $d.ToString('R', [Globalization.CultureInfo]::InvariantCulture)
  }
  return (J $s)
}

Add-Type -AssemblyName Microsoft.VisualBasic
function Read-Gz([string]$path, [scriptblock]$onRow) {
  $fs = [IO.File]::OpenRead($path)
  try {
    $gz = New-Object IO.Compression.GZipStream($fs, [IO.Compression.CompressionMode]::Decompress)
    $sr = New-Object IO.StreamReader($gz, [Text.Encoding]::UTF8)
    $p = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($sr)
    $p.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
    $p.SetDelimiters(',')
    $p.HasFieldsEnclosedInQuotes = $true
    $p.TrimWhiteSpace = $false
    $head = $p.ReadFields()
    $ix = @{}
    for ($i = 0; $i -lt $head.Length; $i++) { $ix[$head[$i].Trim().ToUpper()] = $i }
    $n = 0
    while (-not $p.EndOfData) {
      $f = $p.ReadFields()
      if ($null -eq $f) { continue }
      & $onRow $f $ix
      $n++
    }
    $p.Close()
    return $n
  } finally { $fs.Dispose() }
}
function G($f, $ix, [string]$name) {
  if (-not $ix.ContainsKey($name)) { return '' }
  $i = $ix[$name]
  if ($i -ge $f.Length) { return '' }
  return ($f[$i] + '').Trim()
}

# ------------------------------------------------------------------ main
try {
  # 1. newest complete pair
  $pairs = @()
  foreach ($b in Get-ChildItem -Path $Folder -Filter 'BILLED_DVVNL_*.csv.gz' -File) {
    $m = [regex]::Match($b.Name, '^BILLED_(DVVNL_[A-Z0-9]+)_(\d{2})(\d{2})(\d{4})\.csv\.gz$')
    if (-not $m.Success) { continue }
    $u = Join-Path $Folder ('UN' + $b.Name)
    if (-not (Test-Path $u)) { continue }
    $pairs += [pscustomobject]@{
      Tag = $m.Groups[2].Value + $m.Groups[3].Value + $m.Groups[4].Value
      Date = ('{0}-{1}-{2}' -f $m.Groups[4].Value, $m.Groups[3].Value, $m.Groups[2].Value)
      Billed = $b.FullName; Unbilled = $u
    }
  }
  if (-not $pairs.Count) { Log 'No BILLED_/UNBILLED_ pair in the folder yet - will try again on the next run.'; exit 0 }
  $pair = $pairs | Sort-Object Date -Descending | Select-Object -First 1
  # make sure the files are complete (not still being written by the fetch job)
  foreach ($p in @($pair.Billed, $pair.Unbilled)) {
    if (((Get-Date) - (Get-Item $p).LastWriteTime).TotalSeconds -lt 60) { Log ('File still being written: ' + $p + ' - next run.'); exit 0 }
  }

  $state = $null
  if (Test-Path $StateFile) { try { $state = Get-Content $StateFile -Raw | ConvertFrom-Json } catch {} }

  # 2. which master the update is for
  $token = $null; $master = $null
  if (-not $DryRun) {
    $token = Get-Login
    $info = Call-Server @{ action = 'daily_info'; token = $token } 120
    if (-not $info.success) { throw ('Server: ' + $info.message) }
    $master = @($info.masters | Where-Object { $_.scope -eq 'ALL' })[0]
    if (-not $master) { $master = @($info.masters)[0] }
    if (-not $master) { Log 'No master uploaded on the server yet - upload one in Admin > Master Data.'; exit 0 }
    if (-not $Force -and $master.last_tag -eq $pair.Tag) {
      Log ('Already uploaded: ' + $pair.Tag + ' for master ' + $master.file_name + ' - nothing to do.')
      exit 0
    }
  } else {
    $master = [pscustomobject]@{ master_file_id = 'DRYRUN'; base_id = ''; file_name = '(dry run)' }
  }
  Log ('Building daily update ' + $pair.Tag + ' for master ' + $master.file_name)

  # 3. change set (UNBILLED first, BILLED wins if a consumer is in both)
  $changes = New-Object 'System.Collections.Generic.Dictionary[string,string]'
  $billedSet = New-Object 'System.Collections.Generic.HashSet[string]'
  $sw = [Diagnostics.Stopwatch]::StartNew()

  $nU = Read-Gz $pair.Unbilled {
    param($f, $ix)
    $a = G $f $ix 'ACCT_ID'; if (-not $a) { return }
    $changes[$a] = '{"LAST_BILL_DATE":' + (J (To-Ymd (G $f $ix 'LAST_BILL_DATE'))) +
      ',"LAST_PAY_DATE":' + (J (To-Ymd (G $f $ix 'LAST_PAY_DATE'))) +
      ',"LAST_PAY_AMOUNT":' + (Num (G $f $ix 'LAST_PAY_AMT')) +
      ',"THIS_MONTH_TOTAL_PAY_AMT":' + (Num (G $f $ix 'TOTAL_PAY_AMT')) + '}'
  }
  $nB = Read-Gz $pair.Billed {
    param($f, $ix)
    $a = G $f $ix 'ACCT_ID'; if (-not $a) { return }
    [void]$billedSet.Add($a)
    # billed consumer: Total Outstanding = AMOUNT_PAYABLE of the billed report (kept from master if blank)
    $pay = G $f $ix 'AMOUNT_PAYABLE'
    $out = ''
    if ($pay -ne '') { $out = '"TOTAL_OUTSTANDING":' + (Num $pay) + ',"OUTSTANDING_AS_ON":' + (J $pair.Date) + ',' }
    $changes[$a] = '{' + $out + '"LAST_BILL_DATE":' + (J (To-Ymd (G $f $ix 'BILL_DATE'))) +
      ',"MR_SOURCE_CD":' + (J (G $f $ix 'MR_SOURCE_CD')) +
      ',"BILL_BASIS":' + (J (G $f $ix 'BILL_BASIS')) +
      ',"METER_READ_REMARK":' + (J (G $f $ix 'METER_READ_REMARK')) +
      ',"CONSUMPTION_CURR_MNTH":' + (Num (G $f $ix 'BILLED_UNITS')) +
      ',"CURRENT_READING":' + (Num (G $f $ix 'CUR_RDG_KWH')) +
      ',"LAST_PAY_DATE":' + (J (To-Ymd (G $f $ix 'LAST_PAY_DATE'))) +
      ',"LAST_PAY_AMOUNT":' + (Num (G $f $ix 'LAST_PAYMENT_AMOUNT')) +
      ',"THIS_MONTH_TOTAL_PAY_AMT":' + (Num (G $f $ix 'TOTAL_PAY_AMT')) + '}'
  }
  if ($nB -lt 1000 -or $changes.Count -lt 1000) { throw ('Files look incomplete (billed rows ' + $nB + ', unbilled rows ' + $nU + ') - skipped.') }
  $billed = $billedSet.Count
  $unbilled = $changes.Count - $billed

  $billingMonth = $pair.Date.Substring(0, 7)
  $sb = New-Object System.Text.StringBuilder (32MB)
  [void]$sb.Append('{"base_id":' + (J ([string]$master.base_id)) + ',"seq":0,"report_date":' + (J $pair.Date) +
                   ',"billing_month":' + (J $billingMonth) + ',"source":' + (J $pair.Tag) + ',"changes":{')
  $first = $true
  foreach ($kv in $changes.GetEnumerator()) {
    if (-not $first) { [void]$sb.Append(',') }
    $first = $false
    [void]$sb.Append((J $kv.Key)).Append(':').Append($kv.Value)
  }
  [void]$sb.Append('}}')
  $raw = [Text.Encoding]::UTF8.GetBytes($sb.ToString())
  $sb = $null
  $ms = New-Object IO.MemoryStream
  $gz = New-Object IO.Compression.GZipStream($ms, [IO.Compression.CompressionLevel]::Optimal)
  $gz.Write($raw, 0, $raw.Length); $gz.Close()
  $gzBytes = $ms.ToArray()
  Log ('Built in {0:N0} s: billed {1:N0}, unbilled {2:N0}, {3:N1} MB -> {4:N1} MB gzip' -f $sw.Elapsed.TotalSeconds, $billed, $unbilled, ($raw.Length / 1MB), ($gzBytes.Length / 1MB))

  if ($DryRun) {
    $out = Join-Path $Folder ('orai_daily_' + $pair.Tag + '.json.gz')
    [IO.File]::WriteAllBytes($out, $gzBytes)
    Log ('Dry run: saved ' + $out + ' (not uploaded).')
    exit 0
  }

  # 4. upload
  $r = Call-Server @{
    action = 'daily_update_upload'; token = $token
    master_file_id = $master.master_file_id; base_id = [string]$master.base_id
    tag = $pair.Tag; report_date = $pair.Date; billing_month = $billingMonth
    billed = $billed; unbilled = $unbilled; consumers = $changes.Count; size = $gzBytes.Length
    data = [Convert]::ToBase64String($gzBytes)
  } 600
  if (-not $r.success) { throw ('Upload refused: ' + $r.message) }
  @{ tag = $pair.Tag; master_file_id = $master.master_file_id; at = (Get-Date -Format s); seq = $r.seq } |
    ConvertTo-Json | Set-Content $StateFile -Encoding UTF8
  Log ('UPLOADED ' + $pair.Tag + ': ' + $r.message)
  exit 0
}
catch {
  Log ('ERROR: ' + $_.Exception.Message + ' - will try again on the next run.')
  exit 1
}
