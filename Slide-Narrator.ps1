<#
  Slide Narrator
  ------------------------------------------------------------------
  Turns PowerPoint speaker notes into natural-voice narration and
  exports narrated decks and videos.

  Drag one or more decks (or a folder) onto the window, pick a voice,
  and click Start. For every slide it:
    1. reads and cleans the speaker notes
    2. makes the audio in the background, several slides at a time
       (Windows speech 64-bit, a 32-bit helper, or balcon.exe -
       whichever works for the chosen voice), sets every slide to the
       same loudness, then converts it to AAC
    3. inserts the audio into the slide (plays automatically, icon
       hidden) and sets the slide to advance when the audio ends
  It saves  <deck>_narrated.pptx  next to the original (the original is
  never changed), plus a log:  <deck>_narration.log

  Video (optional):
    - Fast - ffmpeg H.264 or H.265: slide pictures + audio, built in
      the background, several videos at once. 720p to 4K, fixed frame
      rates or 1 frame per slide (VFR).
    - PowerPoint: PowerPoint's own export (keeps animations; slower).
    - Chapters by section or by slide, plus  <deck>_chapters.txt  ready
      to paste into a YouTube description.
    - Captions: subtitles inside the MP4, plus  <deck>.srt  and  <deck>.vtt
  Recommended: Fast - ffmpeg H.264, 4K, 1 per slide (VFR).

  Progress: the right-hand side shows every deck (audio, video, status)
  and everything being worked on right now (each slide being voiced, each
  video being built, with percent and time left).

  Start it by double-clicking  Slide-Narrator.bat
  You can also drag .pptx files onto Slide-Narrator.bat.

  Needs: Windows 10/11, PowerPoint (Microsoft 365).
  Recommended: NaturalVoiceSAPIAdapter (natural voices),
               ffmpeg.exe (even volume, AAC audio, fast video, chapters,
               captions).
  Optional:    balcon.exe (backup speech engine).
  Put ffmpeg.exe and a "balcon" subfolder next to this file.
  ------------------------------------------------------------------
#>

param([Parameter(ValueFromRemainingArguments = $true)][string[]]$StartFiles)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# balcon.exe renames the console window to "Balabolka" when it runs; this puts our name back
function Reset-ConsoleTitle { try { $Host.UI.RawUI.WindowTitle = "Slide Narrator - console (leave this open)" } catch {} }
Reset-ConsoleTitle

$AppDir       = $PSScriptRoot
$SettingsFile = Join-Path $AppDir "Slide-Narrator.settings.json"

# ---------- Helpers ----------
function Find-Balcon {
    $c = @(
        (Join-Path $AppDir "balcon.exe"),
        (Join-Path $AppDir "balcon\balcon.exe"),
        "C:\balcon\balcon.exe",
        "C:\Program Files (x86)\Balabolka\balcon\balcon.exe",
        "C:\Program Files (x86)\Balabolka\balcon.exe",
        "C:\Program Files\Balabolka\balcon\balcon.exe"
    )
    foreach ($p in $c) { if (Test-Path $p) { return (Resolve-Path $p).Path } }
    $cmd = Get-Command balcon.exe -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    return $null
}

function Find-Ffmpeg([string]$balconPath) {
    $cmd = Get-Command ffmpeg.exe -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    $c = @((Join-Path $AppDir "ffmpeg.exe"))
    if ($balconPath) { $c += (Join-Path (Split-Path $balconPath -Parent) "ffmpeg.exe") }
    foreach ($p in $c) { if (Test-Path $p) { return $p } }
    return $null
}

function Get-Voices([string]$balconPath) {
    $voices = New-Object System.Collections.Generic.List[string]

    # 1. Ask Windows (SAPI 5) directly - this includes NaturalVoiceSAPIAdapter voices
    try {
        $sp = New-Object -ComObject SAPI.SpVoice
        $tokens = $sp.GetVoices()
        for ($n = 0; $n -lt $tokens.Count; $n++) {
            $d = $tokens.Item($n).GetDescription()
            if ($d -and -not $voices.Contains($d)) { $voices.Add($d) }
        }
    } catch { Log "Note: could not list voices through SAPI ($($_.Exception.Message))" }

    # 2. Ask balcon (captured through a temp file, which works more reliably)
    if ($balconPath) {
        try {
            $tmp = Join-Path $env:TEMP "slide-narrator-voices.txt"
            if (Test-Path $tmp) { Remove-Item $tmp -Force }
            cmd /c "`"$balconPath`" -l > `"$tmp`" 2>&1" | Out-Null
            Reset-ConsoleTitle
            if (Test-Path $tmp) {
                foreach ($l in (Get-Content $tmp)) {
                    $t = "$l".Trim()
                    if ($t -and -not $t.EndsWith(":") -and -not $voices.Contains($t)) { $voices.Add($t) }
                }
            }
        } catch {}
    }

    # 3. Voices registered in the Windows registry
    $keys = @(
        "HKLM:\SOFTWARE\Microsoft\Speech\Voices\Tokens",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Speech\Voices\Tokens",
        "HKLM:\SOFTWARE\Microsoft\Speech Server\v11.0\Voices\Tokens",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Speech Server\v11.0\Voices\Tokens"
    )
    foreach ($k in $keys) {
        try {
            Get-ChildItem $k -ErrorAction Stop | ForEach-Object {
                $d = (Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue).'(default)'
                if ($d -and -not $voices.Contains($d)) { $voices.Add($d) }
            }
        } catch {}
    }
    # Drop short duplicate names (e.g. "Microsoft Poppy Native") when the
    # full name ("Microsoft Poppy Native - English (Australia)") is also listed
    $all = $voices.ToArray()
    $keep = $all | Where-Object {
        $short = $_
        -not ($all | Where-Object { $_ -ne $short -and ($_.StartsWith("$short - ") -or $_.StartsWith("$short (")) })
    }
    return ,@($keep)
}

# Clean text so speech engines don't choke on hidden characters
function Clean-Text([string]$text) {
    $t = $text
    $t = $t -replace [char]0x0B, ' '                     # PowerPoint soft line break (Shift+Enter)
    $t = $t -replace '[\u2018\u2019\u201B\u2032]', "'"     # curly single quotes
    $t = $t -replace '[\u201C\u201D\u201F\u2033]', '"'     # curly double quotes
    $t = $t -replace '\s*[\u2013\u2014\u2212]\s*', ', '          # en/em dashes, minus
    $t = $t -replace '\u2026', '...'                       # ellipsis
    $t = $t -replace '[\u00A0\u2007\u202F]', ' '           # non-breaking spaces
    $t = $t -replace '[\u2022\u25AA\u25CF\u2023]', ''      # bullet symbols
    $t = $t -replace '&', ' and '
    $t = $t -replace '[<>]', ' '
    $t = $t -replace '[\x00-\x08\x0C\x0E-\x1F\x7F]', ''  # other control characters
    $t = $t -replace '[ \t]{2,}', ' '
    return $t.Trim()
}

# Notes as they should appear in captions (no pronunciation changes)
function Clean-Caption([string]$text) {
    $t = $text -replace '[\u000B\r\n\t]+', ' '
    $t = $t -replace '[\u00A0\u2007\u202F]', ' '
    $t = $t -replace '[\u2022\u25AA\u25CF\u2023]', ''
    $t = $t -replace '[\x00-\x08\x0C\x0E-\x1F\x7F]', ''
    $t = $t -replace ' {2,}', ' '
    return $t.Trim()
}

# Run balcon once; returns $true if the wav was made
function Invoke-Balcon([string]$text, [string]$wavPath, [string]$voice, [int]$rate) {
    $tmp = [IO.Path]::ChangeExtension($wavPath, ".txt")
    [IO.File]::WriteAllText($tmp, $text, (New-Object System.Text.UTF8Encoding($true)))
    if (Test-Path $wavPath) { Remove-Item $wavPath -Force }
    $names = @($voice)
    $short = ($voice -replace '^Microsoft\s+', '') -split '\s+' | Select-Object -First 1
    if ($short -and $short -ne $voice) { $names += $short }
    foreach ($nm in $names) {
        $out = & $script:BalconExe -f $tmp -w $wavPath -n $nm -s $rate -enc utf8 2>&1
        Reset-ConsoleTitle
        $code = $LASTEXITCODE
        $ok = (Test-Path $wavPath) -and ((Get-Item $wavPath).Length -gt 1000)
        if ($ok) { return $true }
    }
    $script:LastSpeechError = "balcon (exit code $code): " + (($out | Out-String).Trim())
    return $false
}

# Join several WAV files (same voice/format) into one
function Join-Wav([string[]]$parts, [string]$outPath) {
    $fmt = $null; $data = New-Object System.IO.MemoryStream
    foreach ($p in $parts) {
        $b = [IO.File]::ReadAllBytes($p); $pos = 12
        while ($pos + 8 -le $b.Length) {
            $id = [Text.Encoding]::ASCII.GetString($b, $pos, 4)
            $size = [BitConverter]::ToInt32($b, $pos + 4)
            if ($size -lt 0 -or $pos + 8 + $size -gt $b.Length) { $size = $b.Length - $pos - 8 }
            if ($id -eq "fmt " -and -not $fmt) { $fmt = New-Object byte[] $size; [Array]::Copy($b, $pos + 8, $fmt, 0, $size) }
            elseif ($id -eq "data") { $data.Write($b, $pos + 8, $size) }
            $pos += 8 + $size + ($size % 2)
        }
    }
    if (-not $fmt) { throw "Could not read the audio pieces." }
    $fs = [IO.File]::Create($outPath); $w = New-Object IO.BinaryWriter($fs)
    $w.Write([Text.Encoding]::ASCII.GetBytes("RIFF")); $w.Write([int](4 + 8 + $fmt.Length + 8 + $data.Length))
    $w.Write([Text.Encoding]::ASCII.GetBytes("WAVE"))
    $w.Write([Text.Encoding]::ASCII.GetBytes("fmt ")); $w.Write([int]$fmt.Length); $w.Write($fmt)
    $w.Write([Text.Encoding]::ASCII.GetBytes("data")); $w.Write([int]$data.Length); $w.Write($data.ToArray())
    $w.Close(); $fs.Close()
}



# Load a single voice straight from the registry (works even when GetVoices fails)
function Get-RegistryVoiceToken([string]$voiceName) {
    $roots = @(
        @{ Ps = "HKLM:\SOFTWARE\Microsoft\Speech\Voices\Tokens";                    Id = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens" },
        @{ Ps = "HKLM:\SOFTWARE\Microsoft\Speech_OneCore\Voices\Tokens";            Id = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech_OneCore\Voices\Tokens" }
    )
    foreach ($r in $roots) {
        try {
            foreach ($k in (Get-ChildItem $r.Ps -ErrorAction Stop)) {
                $d = (Get-ItemProperty $k.PSPath -ErrorAction SilentlyContinue).'(default)'
                if ($d -eq $voiceName -or ($d -and $d -like "*$voiceName*")) {
                    $tok = New-Object -ComObject SAPI.SpObjectToken
                    $tok.SetId($r.Id + "\" + $k.PSChildName)
                    return $tok
                }
            }
        } catch {}
    }
    return $null
}

# Make audio directly through Windows (SAPI 5) - same voices the list shows
function Invoke-Sapi([string]$text, [string]$wavPath, [string]$voiceName, [int]$rate) {
    $stream = $null
    try {
        if (Test-Path $wavPath) { Remove-Item $wavPath -Force }
        $sp = New-Object -ComObject SAPI.SpVoice
        $pick = $null
        try {
            $tokens = $sp.GetVoices()
            for ($n = 0; $n -lt $tokens.Count; $n++) {
                $d = $tokens.Item($n).GetDescription()
                if ($d -eq $voiceName) { $pick = $tokens.Item($n); break }
                if (-not $pick -and $d -like "*$voiceName*") { $pick = $tokens.Item($n) }
            }
        } catch {
            # Listing all voices failed (often caused by a broken add-on voice).
            # Fall back to loading this one voice directly from the registry.
            $pick = Get-RegistryVoiceToken $voiceName
        }
        if (-not $pick) { $script:LastSpeechError = "SAPI: voice '$voiceName' not found"; return $false }
        $sp.Voice = $pick
        $sp.Rate  = [Math]::Max(-10, [Math]::Min(10, $rate))
        $fmt = New-Object -ComObject SAPI.SpAudioFormat
        $fmt.Type = 26                                   # 24 kHz, 16-bit, mono
        $stream = New-Object -ComObject SAPI.SpFileStream
        $stream.Format = $fmt
        $stream.Open($wavPath, 3, $false)                # 3 = create for write
        $sp.AudioOutputStream = $stream
        [void]$sp.Speak($text, 16)                       # 16 = plain text, not XML; waits until finished
        $stream.Close(); $stream = $null
        $ok = (Test-Path $wavPath) -and ((Get-Item $wavPath).Length -gt 1000)
        if (-not $ok) { $script:LastSpeechError = "SAPI: produced an empty audio file" }
        return $ok
    } catch {
        $script:LastSpeechError = "SAPI: " + $_.Exception.Message
        return $false
    } finally {
        if ($stream) { try { $stream.Close() } catch {} }
    }
}


# 32-bit Windows speech. Some voices (incl. Balabolka's natural voice packs) are
# only registered for 32-bit programs, so we run a tiny helper in 32-bit PowerShell.
$script:Ps32 = Join-Path $env:WINDIR "SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
$script:Sapi32Helper = Join-Path $env:TEMP "slide-narrator-sapi32.ps1"
$helperCode = @'
param([string]$TextFile, [string]$Wav, [string]$Voice, [int]$Rate)
try {
    $text = [IO.File]::ReadAllText($TextFile)
    $sp = New-Object -ComObject SAPI.SpVoice
    $pick = $null
    try {
        $t = $sp.GetVoices()
        for ($n = 0; $n -lt $t.Count; $n++) {
            $d = $t.Item($n).GetDescription()
            if ($d -eq $Voice) { $pick = $t.Item($n); break }
            if (-not $pick -and $d -like "*$Voice*") { $pick = $t.Item($n) }
        }
    } catch {}
    if (-not $pick) {
        foreach ($root in @("SOFTWARE\Microsoft\Speech\Voices\Tokens", "SOFTWARE\Microsoft\Speech_OneCore\Voices\Tokens")) {
            try {
                foreach ($k in (Get-ChildItem ("HKLM:\" + $root) -ErrorAction Stop)) {
                    $d = (Get-ItemProperty $k.PSPath -ErrorAction SilentlyContinue).'(default)'
                    if ($d -eq $Voice -or ($d -and $d -like "*$Voice*")) {
                        $pick = New-Object -ComObject SAPI.SpObjectToken
                        $pick.SetId("HKEY_LOCAL_MACHINE\" + $root + "\" + $k.PSChildName); break
                    }
                }
            } catch {}
            if ($pick) { break }
        }
    }
    if (-not $pick) { Write-Output "voice not found in 32-bit"; exit 2 }
    $sp.Voice = $pick
    $sp.Rate = [Math]::Max(-10, [Math]::Min(10, $Rate))
    $fmt = New-Object -ComObject SAPI.SpAudioFormat; $fmt.Type = 26
    $st = New-Object -ComObject SAPI.SpFileStream; $st.Format = $fmt
    $st.Open($Wav, 3, $false); $sp.AudioOutputStream = $st
    [void]$sp.Speak($text, 16); $st.Close()
    exit 0
} catch { Write-Output $_.Exception.Message; exit 1 }
'@
try { Set-Content -Path $script:Sapi32Helper -Value $helperCode -Encoding UTF8 } catch {}

# Video worker: builds an MP4 from slide pictures + audio with ffmpeg (runs as its own process)
$script:VideoWorker = Join-Path $env:TEMP "slide-narrator-video.ps1"
$videoWorkerCode = @'
param([string]$JobFile)
# Slide Narrator video worker: builds one MP4 from slide images + audio with ffmpeg,
# with chapters, a YouTube chapter list and captions. Mode "extras" adds chapters and
# captions to an MP4 that PowerPoint has already exported.
$ErrorActionPreference = "Continue"   # native tools write to stderr; we check exit codes instead
$job = Get-Content -LiteralPath $JobFile -Raw | ConvertFrom-Json
$log = $job.Log
$inv = [Globalization.CultureInfo]::InvariantCulture
$utf8 = New-Object System.Text.UTF8Encoding($false)
# Progress file is read by the app while we write it: retry briefly, and never fail the job over it
function SetProg([string]$v) {
    if (-not $job.Progress) { return }
    for ($i = 0; $i -lt 40; $i++) {
        try { [IO.File]::WriteAllText($job.Progress, $v); return } catch { Start-Sleep -Milliseconds 50 }
    }
}
# Progress for the app: "stage|percent|seconds left" (-1 = not known yet)
$script:Stage = "Starting"; $script:StageBase = 0.0; $script:StageSpan = 0.0
$script:StageClock = [Diagnostics.Stopwatch]::StartNew()
function Prog([string]$stage, [double]$frac, [double]$eta) {
    $pct = [int][Math]::Max(0, [Math]::Min(99, [Math]::Floor($script:StageBase + $script:StageSpan * $frac)))
    SetProg ("{0}|{1}|{2}" -f $stage, $pct, [int][Math]::Round($eta))
}
function Set-Stage([string]$name, [double]$base, [double]$span) {
    $script:Stage = $name; $script:StageBase = $base; $script:StageSpan = $span
    $script:StageClock.Restart(); Prog $name 0 -1
}
function Step-Prog {
    if ($script:progTotal -le 0) { return }
    $f = [Math]::Min(1.0, $script:doneCount / $script:progTotal); $eta = -1
    if ($f -gt 0.05) { $eta = $script:StageClock.Elapsed.TotalSeconds * (1 - $f) / $f }
    Prog $script:Stage $f $eta
}
function W([string]$m) { try { Add-Content -LiteralPath $log -Value ((Get-Date -Format "HH:mm:ss") + "  " + $m) -Encoding UTF8 } catch {} }
function F3([double]$x) { return $x.ToString("0.###", $inv) }
function Q2([string]$x) { return '"' + $x + '"' }
function QList([string]$x) { return "file '" + ($x -replace "'", "'\''") + "'" }

# ---- Small ffmpeg jobs, run several at once ----
$script:threads = 4; if ($job.Threads) { $script:threads = [Math]::Max(1, [int]$job.Threads) }
$script:running = New-Object System.Collections.ArrayList
$script:doneCount = 0; $script:progTotal = 0; $script:Cues = @()
function Wait-Slots([int]$max) {
    while ($script:running.Count -ge $max) {
        for ($k = $script:running.Count - 1; $k -ge 0; $k--) {
            $pr = $script:running[$k]
            if ($pr.P.HasExited) {
                $bad = ($pr.P.ExitCode -ne 0) -or ($pr.Check -and -not (Test-Path -LiteralPath $pr.Check))
                if ($bad) {
                    $msg = "$($pr.What): " + $pr.P.StandardError.ReadToEnd()
                    if ($pr.Soft) { W ("   note: " + $msg.Trim()) } else { throw $msg }
                }
                $script:running.RemoveAt($k); $script:doneCount++
                Step-Prog
            }
        }
        if ($script:running.Count -ge $max) { Start-Sleep -Milliseconds 50 }
    }
}
function Start-Ff([string]$argLine, [string]$check, [string]$what, [bool]$soft) {
    Wait-Slots $script:threads
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $job.Ffmpeg; $psi.Arguments = $argLine; $psi.WorkingDirectory = $job.WorkDir
    $psi.UseShellExecute = $false; $psi.CreateNoWindow = $true; $psi.RedirectStandardError = $true
    [void]$script:running.Add([pscustomobject]@{ P = [System.Diagnostics.Process]::Start($psi); Check = $check; What = $what; Soft = $soft })
}

# Run one long ffmpeg job, reporting its progress (ffmpeg's -progress file) as it goes
function Join-Args($list) {
    return (@($list | ForEach-Object { $x = [string]$_; if ($x -eq "" -or $x -match '[\s"]') { '"' + ($x -replace '"', '\"') + '"' } else { $x } }) -join ' ')
}
function Run-Ff($argList, [double]$total) {
    $pf = Join-Path $job.WorkDir "encode-progress.txt"
    Remove-Item -LiteralPath $pf -Force -ErrorAction SilentlyContinue
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $job.Ffmpeg; $psi.Arguments = (Join-Args (@("-progress", $pf, "-nostats") + @($argList)))
    $psi.UseShellExecute = $false; $psi.CreateNoWindow = $true; $psi.RedirectStandardError = $true; $psi.WorkingDirectory = $job.WorkDir
    $p = [System.Diagnostics.Process]::Start($psi)
    $errTask = $p.StandardError.ReadToEndAsync()
    $sw = [Diagnostics.Stopwatch]::StartNew()
    while (-not $p.WaitForExit(400)) {
        if ($total -le 0) { continue }
        $txt = ""
        try {
            $fs = [IO.File]::Open($pf, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete)
            try { $txt = (New-Object IO.StreamReader($fs)).ReadToEnd() } finally { $fs.Dispose() }
        } catch {}
        $m = [regex]::Matches($txt, 'out_time_(?:us|ms)=(\d+)')
        if ($m.Count -gt 0) {
            $f = [Math]::Min(1.0, ([double]$m[$m.Count - 1].Groups[1].Value / 1000000.0) / $total)
            $eta = -1; if ($f -gt 0.03) { $eta = $sw.Elapsed.TotalSeconds * (1 - $f) / $f }
            Prog $script:Stage $f $eta
        }
    }
    $p.WaitForExit()
    $err = [string]$errTask.Result
    Remove-Item -LiteralPath $pf -Force -ErrorAction SilentlyContinue
    return @{ Code = $p.ExitCode; Err = $err }
}

# ---- Captions ----
# Find the pauses in every narrated slide (silencedetect, written to a small text file per slide)
function Find-Pauses($segs) {
    $map = @{}
    foreach ($s in $segs) {
        if (-not ($s.Wav -and $s.Caption)) { continue }
        $f = "pauses_{0:D3}.txt" -f [int]$s.Index
        Remove-Item -LiteralPath (Join-Path $job.WorkDir $f) -Force -ErrorAction SilentlyContinue
        Start-Ff ("-hide_banner -nostats -loglevel error -i " + (Q2 $s.Wav) + " -af silencedetect=n=-40dB:d=0.15,ametadata=mode=print:file=$f -f null -") "" "slide $($s.Index) pauses" $true
    }
    Wait-Slots 1
    foreach ($s in $segs) {
        if (-not ($s.Wav -and $s.Caption)) { continue }
        $f = Join-Path $job.WorkDir ("pauses_{0:D3}.txt" -f [int]$s.Index)
        $list = @(); $st = $null
        if (Test-Path -LiteralPath $f) {
            foreach ($line in [IO.File]::ReadAllLines($f)) {
                if ($line -match 'lavfi\.silence_start=(-?[\d.]+)') { $st = [Math]::Max(0.0, [double]::Parse($matches[1], $inv)) }
                elseif ($line -match 'lavfi\.silence_end=([\d.]+)' -and $null -ne $st) {
                    $list += [pscustomobject]@{ S = $st; E = [double]::Parse($matches[1], $inv) }; $st = $null
                }
            }
            Remove-Item -LiteralPath $f -Force -ErrorAction SilentlyContinue
        }
        if ($null -ne $st) { $list += [pscustomobject]@{ S = $st; E = [double]$s.WavDur } }
        $map[[int]$s.Index] = $list
    }
    return $map
}

# Split a slide's text into caption-sized pieces (max ~84 characters, 2 lines of ~42)
function Split-Caption([string]$text) {
    $t = ($text -replace '\s+', ' ').Trim()
    $out = New-Object System.Collections.Generic.List[string]
    if (-not $t) { return ,$out }
    foreach ($sen in [regex]::Split($t, '(?<=[.!?]["\x27\u201D\u2019)\]]?)\s+')) {
        if (-not $sen) { continue }
        if ($sen.Length -le 84) { $out.Add($sen); continue }
        $cur = ""
        foreach ($w in ($sen -split ' ')) {
            if ($cur -and ($cur.Length + 1 + $w.Length) -gt 84) { $out.Add($cur); $cur = $w }
            elseif ($cur) { $cur = "$cur $w" } else { $cur = $w }
            if ($cur.Length -ge 40 -and $cur -match '[,;:]$') { $out.Add($cur); $cur = "" }
        }
        if ($cur) {
            # Don't leave a tiny piece on its own
            if ($cur.Length -lt 15 -and $out.Count -gt 0 -and ($out[$out.Count - 1].Length + 1 + $cur.Length) -le 100) { $out[$out.Count - 1] += " $cur" }
            else { $out.Add($cur) }
        }
    }
    return ,$out
}
function Wrap-Caption([string]$c) {
    if ($c.Length -le 42) { return $c }
    $mid = [int]($c.Length / 2); $best = -1
    for ($d = 0; $d -lt $mid; $d++) {
        if ($mid + $d -lt $c.Length -and $c[$mid + $d] -eq ' ') { $best = $mid + $d; break }
        if ($mid - $d -ge 0 -and $c[$mid - $d] -eq ' ') { $best = $mid - $d; break }
    }
    if ($best -lt 0) { return $c }
    return $c.Substring(0, $best) + "`n" + $c.Substring($best + 1)
}

# Time one slide's captions: share the speech between the pieces by length, then
# line each break up with a real pause in the audio when there is one nearby
function Get-SlideCues($s, [double]$offset, [double]$slideLen, $pauses) {
    $cues = @()
    $chunks = Split-Caption ([string]$s.Caption)
    if ($chunks.Count -eq 0) { return $cues }
    $wd = [double]$s.WavDur
    $s0 = 0.0; $s1 = $wd; $inner = @()
    foreach ($p in @($pauses)) {
        if ($p.S -le 0.05) { $s0 = [Math]::Max($s0, $p.E); continue }      # silence before the speech
        if ($p.E -ge $wd - 0.05) { $s1 = [Math]::Min($s1, $p.S); continue } # silence after it
        $inner += $p
    }
    if ($s1 -le $s0 + 0.2) { $s0 = 0.0; $s1 = $wd }
    $w = @(); foreach ($c in $chunks) { $w += ($c.Length + 4) }
    $total = 0.0; foreach ($x in $w) { $total += $x }
    $starts = @($s0); $ends = @()
    $anchorT = $s0; $anchorW = 0.0; $cum = 0.0; $next = 0
    for ($k = 0; $k -lt $chunks.Count - 1; $k++) {
        $cum += $w[$k]
        $est = $anchorT + ($cum - $anchorW) / [Math]::Max(1.0, $total - $anchorW) * ($s1 - $anchorT)
        $win = [Math]::Max(0.8, 0.35 * ($est - $starts[$k]))
        $pick = -1; $bestD = 1e9
        for ($m = $next; $m -lt $inner.Count; $m++) {
            $mid = ($inner[$m].S + $inner[$m].E) / 2
            if ($inner[$m].S -le $starts[$k] + 0.3) { continue }
            $dd = [Math]::Abs($mid - $est)
            if ($dd -le $win -and $dd -lt $bestD) { $bestD = $dd; $pick = $m }
            if ($mid -gt $est + $win) { break }
        }
        if ($pick -ge 0) {
            $ends += $inner[$pick].S; $starts += $inner[$pick].E; $next = $pick + 1
            $anchorT = $inner[$pick].E; $anchorW = $cum
        } else {
            $ends += $est; $starts += $est
        }
    }
    $ends += $s1
    for ($k = 0; $k -lt $chunks.Count; $k++) {
        $a = $starts[$k]; $b = $ends[$k]
        # Keep a caption up through short pauses; hold the last one a moment after speech ends
        if ($k + 1 -lt $chunks.Count) { if ($starts[$k + 1] - $b -lt 1.0) { $b = $starts[$k + 1] } else { $b += 0.5 } }
        else { $b = [Math]::Min($b + 0.5, $slideLen) }
        if ($b -lt $a + 0.5) { $b = [Math]::Min($a + 0.5, $slideLen) }
        if ($b -le $a) { continue }
        $cues += [pscustomobject]@{ Start = $offset + $a; End = $offset + $b; Text = (Wrap-Caption $chunks[$k]) }
    }
    return $cues
}

function Format-CueTime([double]$t, [string]$sep) {
    $ms = [long][Math]::Round($t * 1000)
    $h = [long][Math]::Floor($ms / 3600000); $ms -= $h * 3600000
    $m = [long][Math]::Floor($ms / 60000);   $ms -= $m * 60000
    $sec = [long][Math]::Floor($ms / 1000);  $ms -= $sec * 1000
    return ("{0:00}:{1:00}:{2:00}{3}{4:000}" -f $h, $m, $sec, $sep, $ms)
}

# Build captions for the whole video; writes <video>.srt and <video>.vtt. Returns the .srt path or "".
function Write-Captions($segs, $startsArr, $lensArr) {
    if (-not $job.Captions) { return "" }
    try {
        $pauses = Find-Pauses $segs
        $cues = @()
        for ($k = 0; $k -lt $segs.Count; $k++) {
            $s = $segs[$k]
            if (-not ($s.Wav -and $s.Caption)) { continue }
            $cues += Get-SlideCues $s $startsArr[$k] $lensArr[$k] $pauses[[int]$s.Index]
        }
        $script:Cues = $cues
        if ($cues.Count -eq 0) { return "" }
        $base = [IO.Path]::Combine([IO.Path]::GetDirectoryName($job.Out), [IO.Path]::GetFileNameWithoutExtension($job.Out))
        $srt = @(); $vtt = @("WEBVTT", ""); $n = 0
        foreach ($c in $cues) {
            $n++
            $srt += @("$n", ((Format-CueTime $c.Start ",") + " --> " + (Format-CueTime $c.End ",")), $c.Text, "")
            $vtt += @(((Format-CueTime $c.Start ".") + " --> " + (Format-CueTime $c.End ".")), $c.Text, "")
        }
        [IO.File]::WriteAllText("$base.srt", (($srt -join "`r`n") -replace "(?<!`r)`n", "`r`n"), $utf8)
        [IO.File]::WriteAllText("$base.vtt", ($vtt -join "`n"), $utf8)
        W "Captions: $($cues.Count) captions saved as $([IO.Path]::GetFileName($base)).srt and .vtt"
        return "$base.srt"
    } catch { W ("Captions skipped: " + $_.Exception.Message); return "" }
}

# Chapters: FFMETADATA for the MP4, plus <video>_chapters.txt ready to paste into a YouTube description
function Write-Chapters($chapters, [double]$clock) {
    if ($chapters.Count -eq 0) { return "" }
    $meta = Join-Path $job.WorkDir "chapters.txt"
    $ml = @(";FFMETADATA1"); $yt = @()
    for ($c = 0; $c -lt $chapters.Count; $c++) {
        $st = [long][Math]::Round($chapters[$c].Start * 1000)
        $en = if ($c + 1 -lt $chapters.Count) { [long][Math]::Round($chapters[$c + 1].Start * 1000) } else { [long][Math]::Round($clock * 1000) }
        $title = ([string]$chapters[$c].Title -replace '[\r\n\v]+', ' ').Trim()
        $ml += @("[CHAPTER]", "TIMEBASE=1/1000", "START=$st", "END=$en", ("title=" + ($title -replace '([\\=;#])', '\$1')))
        $sec = if ($c -eq 0) { 0 } else { [long][Math]::Floor($chapters[$c].Start) }   # YouTube needs the first at 0:00
        $hh = [long][Math]::Floor($sec / 3600); $mm = [long][Math]::Floor(($sec % 3600) / 60); $ss = $sec % 60
        $stamp = if ($clock -ge 3600) { "{0}:{1:00}:{2:00}" -f $hh, $mm, $ss } else { "{0}:{1:00}" -f $mm, $ss }
        $yt += "$stamp $title"
    }
    [IO.File]::WriteAllLines($meta, $ml, $utf8)
    try {
        $base = [IO.Path]::Combine([IO.Path]::GetDirectoryName($job.Out), [IO.Path]::GetFileNameWithoutExtension($job.Out))
        [IO.File]::WriteAllLines("${base}_chapters.txt", $yt, $utf8)
        $short = 0
        for ($c = 0; $c -lt $chapters.Count; $c++) {
            $en = if ($c + 1 -lt $chapters.Count) { $chapters[$c + 1].Start } else { $clock }
            if ($en - $chapters[$c].Start -lt 10) { $short++ }
        }
        $msg = "Chapter list for YouTube saved: $([IO.Path]::GetFileName($base))_chapters.txt"
        if ($chapters.Count -lt 3) { $msg += " (YouTube needs at least 3 chapters to show them)" }
        elseif ($short -gt 0) { $msg += " (YouTube needs each chapter to be 10 s or longer; $short are shorter)" }
        W $msg
    } catch { W ("Chapter list skipped: " + $_.Exception.Message) }
    return $meta
}

# Extra inputs and maps for chapters and captions, starting at input number $first
function Get-Extras([string]$meta, [string]$srt, [int]$first) {
    $in = @(); $map = @(); $i = $first
    if ($meta) { $in += @("-i", $meta); $map += @("-map_metadata", "$i", "-map_chapters", "$i"); $i++ }
    if ($srt)  { $in += @("-sub_charenc", "UTF-8", "-i", $srt); $map += @("-map", "$($i):s", "-c:s", "mov_text", "-metadata:s:s:0", "language=eng", "-metadata:s:s:0", "handler_name=English"); $i++ }
    return @{ In = $in; Map = $map }
}

try {
    $ff = $job.Ffmpeg; $fps = [int]$job.Fps
    $segs = @($job.Segments)
    if ($job.Codec -eq "h265") { $vc = @("-c:v", "libx265", "-preset", "medium", "-crf", "26", "-forced-idr", "1", "-x265-params", "log-level=error") }
    else                       { $vc = @("-c:v", "libx264", "-preset", "medium", "-crf", "22", "-tune", "stillimage") }

    # Where each slide starts and how long it lasts in the finished video
    $starts = @(); $lens = @(); $clock = 0.0; $chapters = @()
    foreach ($s in $segs) {
        if ($job.Mode -ne "extras" -and $fps -gt 0) {
            $frames = [Math]::Max(1, [Math]::Ceiling([double]$s.Dur * $fps - 0.000001))   # whole frames
            $secs = $frames / $fps
        } else { $secs = [Math]::Round([double]$s.Dur, 3) }
        if ($s.Chapter) { $chapters += [pscustomobject]@{ Title = [string]$s.Chapter; Start = $clock } }
        $starts += $clock; $lens += $secs; $clock += $secs
    }
    $withPauses = 0; if ($job.Captions) { $withPauses = @($segs | Where-Object { $_.Wav -and $_.Caption }).Count }

    if ($job.Mode -eq "extras") {
        # ---- Add chapters and captions to a video PowerPoint has already made ----
        $script:progTotal = $withPauses
        Set-Stage "Timing captions" 0 50
        $srt = Write-Captions $segs $starts $lens
        $meta = Write-Chapters $chapters $clock
        if (-not $srt -and -not $meta) { SetProg "done"; exit 0 }
        $ex = Get-Extras $meta $srt 1
        $tmp = [IO.Path]::ChangeExtension($job.Out, ".extras.mp4")
        $a = @("-y", "-loglevel", "error", "-i", $job.Out) + $ex.In + @("-map", "0:v", "-map", "0:a") + $ex.Map + @("-c:v", "copy", "-c:a", "copy", "-movflags", "+faststart", $tmp)
        Set-Stage "Adding to video" 50 50
        $res = Run-Ff $a $clock
        if ($res.Code -ne 0 -or -not (Test-Path -LiteralPath $tmp)) { throw "adding chapters/captions: " + $res.Err.Trim() }
        Move-Item -LiteralPath $tmp -Destination $job.Out -Force
        W "Added $(if ($meta) { "$($chapters.Count) chapters" })$(if ($meta -and $srt) { ' and ' })$(if ($srt) { 'captions' }) to the video."
        SetProg "done"; exit 0
    }

    if ($fps -le 0) {
        # ---- One frame per slide (variable frame rate) ----
        $script:progTotal = $segs.Count + $withPauses
        Set-Stage "Preparing audio" 0 8
        $imgLines = @(); $audLines = @(); $pads = @()
        # Prepare each slide's audio in parallel (one small ffmpeg per slide)
        for ($k = 0; $k -lt $segs.Count; $k++) {
            $s = $segs[$k]
            $dur = F3 $lens[$k]
            $pad = Join-Path $job.WorkDir ("pad_{0:D3}.wav" -f ($k + 1))
            if ($s.Wav) { $argLine = "-y -loglevel error -i " + (Q2 $s.Wav) + " -af apad -t $dur -ar 48000 -ac 1 -c:a pcm_s16le " + (Q2 $pad) }
            else        { $argLine = "-y -loglevel error -f lavfi -i anullsrc=r=48000:cl=mono -t $dur -c:a pcm_s16le " + (Q2 $pad) }
            Start-Ff $argLine $pad "slide $($s.Index) audio" $false
            $pads += $pad
            $audLines += (QList $pad)
        }
        Wait-Slots 1                                  # all audio pieces finished
        $srt = Write-Captions $segs $starts $lens
        # Players draw captions onto video frames, so repeat the slide's picture wherever a caption
        # starts or ends. The repeats are "no change" frames, so they add almost nothing to the size.
        $cuts = @()
        foreach ($c in $script:Cues) { $cuts += [Math]::Round($c.Start, 3); $cuts += [Math]::Round($c.End, 3) }
        $cuts = @($cuts | Sort-Object -Unique)
        $keys = @()
        for ($k = 0; $k -lt $segs.Count; $k++) {
            $a0 = [Math]::Round($starts[$k], 3); $b0 = [Math]::Round($starts[$k] + $lens[$k], 3); $prev = $a0
            $keys += F3 ([Math]::Max(0.0, $a0 - 0.005))           # keyframe at every slide start
            foreach ($t in $cuts) {
                if ($t -gt $prev + 0.04 -and $t -lt $b0 - 0.04) {
                    $imgLines += (QList $segs[$k].Img); $imgLines += ("duration " + (F3 ($t - $prev))); $prev = $t
                }
            }
            $imgLines += (QList $segs[$k].Img); $imgLines += ("duration " + (F3 ($b0 - $prev)))
        }
        $imgLines += (QList $segs[-1].Img)      # concat quirk: repeat the last picture
        # Millisecond timing for each picture (needs ffmpeg 5 or later; older versions use the plain list)
        $precise = @(); foreach ($l in $imgLines) { $precise += $l; if ($l -like "file *") { $precise += "option framerate 1000" } }
        $imgList = Join-Path $job.WorkDir "images.txt"; [IO.File]::WriteAllLines($imgList, $precise, $utf8)
        $imgPlain = Join-Path $job.WorkDir "images-plain.txt"; [IO.File]::WriteAllLines($imgPlain, $imgLines, $utf8)
        $audList = Join-Path $job.WorkDir "audio.txt";  [IO.File]::WriteAllLines($audList, $audLines, $utf8)
        $meta = Write-Chapters $chapters $clock
        $ex = Get-Extras $meta $srt 2
        $a = @("-y", "-loglevel", "error", "-f", "concat", "-safe", "0", "-i", $imgList, "-f", "concat", "-safe", "0", "-i", $audList) + $ex.In
        $map = @("-map", "0:v", "-map", "1:a") + $ex.Map
        # Fixed quality for each slide's picture (same quality as before captions), and the repeated
        # pictures get almost no data: they copy the previous frame exactly
        if ($job.Codec -eq "h265") { $vcv = @("-c:v", "libx265", "-preset", "medium", "-forced-idr", "1", "-x265-params", "log-level=error:qp=51:ipratio=8") }
        else                       { $vcv = @("-c:v", "libx264", "-preset", "medium", "-tune", "stillimage", "-qp", "39", "-x264-params", "ipratio=8") }
        $vcv += @("-g", "100000", "-force_key_frames", ($keys -join ","), "-pix_fmt", "yuv420p")
        $tail = @("-c:a", "aac", "-b:a", "64k", "-t", (F3 $clock), "-movflags", "+faststart")
        if ($job.Codec -eq "h265") { $tail += @("-tag:v", "hvc1") }
        Set-Stage "Encoding" 8 92
        $res = Run-Ff @($a + $map + $vcv + @("-fps_mode", "vfr") + $tail + @($job.Out)) $clock
        if ($res.Code -ne 0) {   # older ffmpeg: -vsync instead of -fps_mode, no per-picture options
            W ("   note: retrying for an older ffmpeg: " + $res.Err.Trim())
            $a = @($a | ForEach-Object { if ($_ -eq $imgList) { $imgPlain } else { $_ } })
            Set-Stage "Encoding" 8 92
            $res = Run-Ff @($a + $map + $vcv + @("-vsync", "vfr") + $tail + @($job.Out)) $clock
        }
        if ($res.Code -ne 0 -or -not (Test-Path -LiteralPath $job.Out)) { throw "building video: " + $res.Err.Trim() }
        foreach ($pd in $pads) { Remove-Item -LiteralPath $pd -Force -ErrorAction SilentlyContinue }
        SetProg "done"
        exit 0
    }

    # ---- Fixed frame rate: one short clip per slide, then join them ----
    $script:progTotal = $segs.Count + $withPauses
    Set-Stage "Encoding slides" 0 95
    $srt = Write-Captions $segs $starts $lens
    $clips = @()
    for ($k = 0; $k -lt $segs.Count; $k++) {
        $s = $segs[$k]
        $seg = Join-Path $job.WorkDir ("seg_{0:D3}.mkv" -f ($k + 1))
        $dur = ([double]$lens[$k]).ToString("0.######", $inv)
        $a = @("-y", "-loglevel", "error", "-loop", "1", "-framerate", "$fps", "-i", $s.Img)
        if ($s.Wav) { $a += @("-i", $s.Wav, "-af", "apad") } else { $a += @("-f", "lavfi", "-i", "anullsrc=r=48000:cl=mono") }
        $a += @("-t", $dur) + $vc + @("-pix_fmt", "yuv420p", "-r", "$fps", "-c:a", "pcm_s16le", "-ar", "48000", "-ac", "1", $seg)
        $out = & $ff @a 2>&1
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $seg)) { throw "slide $($s.Index): " + (($out | Out-String).Trim()) }
        $clips += $seg
        $script:doneCount++; Step-Prog
    }
    $list = Join-Path $job.WorkDir "segments.txt"
    [IO.File]::WriteAllLines($list, @($clips | ForEach-Object { QList $_ }), $utf8)
    $meta = Write-Chapters $chapters $clock
    $ex = Get-Extras $meta $srt 1
    $a = @("-y", "-loglevel", "error", "-f", "concat", "-safe", "0", "-i", $list) + $ex.In + @("-map", "0:v", "-map", "0:a") + $ex.Map
    $a += @("-c:v", "copy", "-c:a", "aac", "-b:a", "64k", "-movflags", "+faststart")
    if ($job.Codec -eq "h265") { $a += @("-tag:v", "hvc1") }
    $a += $job.Out
    Set-Stage "Joining" 95 5
    $res = Run-Ff $a $clock
    if ($res.Code -ne 0 -or -not (Test-Path -LiteralPath $job.Out)) { throw "joining: " + $res.Err.Trim() }
    foreach ($sg in $clips) { Remove-Item -LiteralPath $sg -Force -ErrorAction SilentlyContinue }
    SetProg "done"
    exit 0
} catch {
    W ("VIDEO FAILED: " + $_.Exception.Message)
    SetProg ("failed: " + $_.Exception.Message)
    exit 1
}
'@
try { Set-Content -Path $script:VideoWorker -Value $videoWorkerCode -Encoding UTF8 } catch {}

# Length of a WAV file in seconds
function Get-WavSeconds([string]$path) {
    try {
        $b = [IO.File]::ReadAllBytes($path); $pos = 12; $rate = 0; $len = 0
        while ($pos + 8 -le $b.Length) {
            $id = [Text.Encoding]::ASCII.GetString($b, $pos, 4); $size = [BitConverter]::ToInt32($b, $pos + 4)
            if ($size -lt 0 -or $pos + 8 + $size -gt $b.Length) { $size = $b.Length - $pos - 8 }
            if ($id -eq 'fmt ') { $rate = [BitConverter]::ToInt32($b, $pos + 16) }
            elseif ($id -eq 'data') { $len = $size }
            $pos += 8 + $size + ($size % 2)
        }
        if ($rate -gt 0) { return [double]$len / $rate }
    } catch {}
    return 5.0
}

function Invoke-Sapi32([string]$text, [string]$wavPath, [string]$voice, [int]$rate) {
    if (-not (Test-Path $script:Ps32)) { $script:LastSpeechError = "SAPI32: 32-bit PowerShell not found"; return $false }
    $tmp = Join-Path $env:TEMP "slide-narrator-text.txt"
    [IO.File]::WriteAllText($tmp, $text, (New-Object System.Text.UTF8Encoding($true)))
    if (Test-Path $wavPath) { Remove-Item $wavPath -Force }
    $out = & $script:Ps32 -NoProfile -ExecutionPolicy Bypass -File $script:Sapi32Helper $tmp $wavPath $voice $rate 2>&1
    $ok = (Test-Path $wavPath) -and ((Get-Item $wavPath).Length -gt 1000)
    if (-not $ok) { $script:LastSpeechError = "SAPI32: " + (($out | Out-String).Trim()) }
    return $ok
}

function Invoke-Engine([string]$eng, [string]$text, [string]$wavPath, [string]$voice, [int]$rate) {
    switch ($eng) {
        "sapi"   { return (Invoke-Sapi   $text $wavPath $voice $rate) }
        "sapi32" { return (Invoke-Sapi32 $text $wavPath $voice $rate) }
        "balcon" { if (-not $script:BalconExe) { $script:LastSpeechError = "balcon: not installed"; return $false }
                   return (Invoke-Balcon $text $wavPath $voice $rate) }
    }
    return $false
}

# Try each engine in turn. Remembers whichever works.
function Invoke-Speech([string]$text, [string]$wavPath, [string]$voice, [int]$rate) {
    $all = @("sapi", "sapi32", "balcon")
    $order = @()
    if ($script:Engine) { $order += $script:Engine }
    $order += ($all | Where-Object { $_ -ne $script:Engine })
    $errors = @()
    foreach ($eng in $order) {
        if (Invoke-Engine $eng $text $wavPath $voice $rate) {
            if ($script:Engine -ne $eng) { Log "   speech engine in use: $eng"; $script:Engine = $eng }
            return $true
        }
        $errors += $script:LastSpeechError
    }
    $script:LastSpeechError = ($errors -join " | ")
    return $false
}

# Make audio for one slide: try whole text (with retries), then sentence by sentence
function New-SlideAudio([string]$text, [string]$wav, [string]$voice, [int]$rate) {
    [IO.File]::WriteAllText([IO.Path]::ChangeExtension($wav, ".txt"), $text, (New-Object System.Text.UTF8Encoding($true)))
    for ($try = 1; $try -le 2; $try++) {
        if (Invoke-Speech $text $wav $voice $rate) { return $true }
        Log "   attempt $try failed: $($script:LastSpeechError)"
        Start-Sleep -Seconds 2
    }
    Log "   trying sentence by sentence..."
    $sentences = [regex]::Split($text, '(?<=[.!?])\s+') | Where-Object { $_.Trim() }
    $parts = @(); $n = 0
    foreach ($sen in $sentences) {
        $n++
        $pw = [IO.Path]::Combine([IO.Path]::GetDirectoryName($wav), [IO.Path]::GetFileNameWithoutExtension($wav) + "_part$n.wav")
        $ok = $false
        for ($try = 1; $try -le 2 -and -not $ok; $try++) { $ok = Invoke-Speech $sen $pw $voice $rate; if (-not $ok) { Start-Sleep -Seconds 2 } }
        if (-not $ok) { Log "   could not read this sentence: $sen"; Log "   reason: $($script:LastSpeechError)"; return $false }
        $parts += $pw
    }
    Join-Wav $parts $wav
    foreach ($p in $parts) { Remove-Item $p -Force -ErrorAction SilentlyContinue; Remove-Item ([IO.Path]::ChangeExtension($p, ".txt")) -Force -ErrorAction SilentlyContinue }
    return (Test-Path $wav)
}

function Load-Settings {
    try { if (Test-Path $SettingsFile) { return Get-Content $SettingsFile -Raw | ConvertFrom-Json } } catch {}
    return $null
}
function Save-Settings($obj) {
    try { $obj | ConvertTo-Json | Set-Content -Path $SettingsFile -Encoding UTF8 } catch {}
}

# ---------- Build the window ----------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Slide Narrator"
$form.Size = New-Object System.Drawing.Size(1170, 774)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$form.AllowDrop = $true
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false

$navy  = [System.Drawing.Color]::FromArgb(37, 51, 108)
$amber = [System.Drawing.Color]::FromArgb(255, 192, 0)

# Drop zone
$drop = New-Object System.Windows.Forms.Label
$drop.Location = New-Object System.Drawing.Point(20, 15)
$drop.Size = New-Object System.Drawing.Size(585, 50)
$drop.TextAlign = "MiddleCenter"
$drop.BorderStyle = "FixedSingle"
$drop.BackColor = [System.Drawing.Color]::FromArgb(238, 241, 250)
$drop.ForeColor = $navy
$drop.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$drop.Text = "Drag PowerPoint files (or a folder) here - as many as you like"
$drop.AllowDrop = $true
$form.Controls.Add($drop)

# File path + browse
$fileList = New-Object System.Windows.Forms.ListBox
$fileList.Location = New-Object System.Drawing.Point(20, 70)
$fileList.Size = New-Object System.Drawing.Size(480, 80)
$fileList.SelectionMode = "MultiExtended"
$fileList.HorizontalScrollbar = $true
$fileList.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$fileList.AllowDrop = $true
$form.Controls.Add($fileList)

$smallFont = New-Object System.Drawing.Font("Segoe UI", 9)
$browse = New-Object System.Windows.Forms.Button
$browse.Text = "Add files..."
$browse.Location = New-Object System.Drawing.Point(510, 69)
$browse.Size = New-Object System.Drawing.Size(95, 26)
$browse.Font = $smallFont
$form.Controls.Add($browse)

$removeBtn = New-Object System.Windows.Forms.Button
$removeBtn.Text = "Remove"
$removeBtn.Location = New-Object System.Drawing.Point(510, 97)
$removeBtn.Size = New-Object System.Drawing.Size(95, 26)
$removeBtn.Font = $smallFont
$form.Controls.Add($removeBtn)

$clearBtn = New-Object System.Windows.Forms.Button
$clearBtn.Text = "Clear list"
$clearBtn.Location = New-Object System.Drawing.Point(510, 125)
$clearBtn.Size = New-Object System.Drawing.Size(95, 26)
$clearBtn.Font = $smallFont
$form.Controls.Add($clearBtn)

# Voice
$lblVoice = New-Object System.Windows.Forms.Label
$lblVoice.Text = "Voice:"
$lblVoice.Location = New-Object System.Drawing.Point(20, 160)
$lblVoice.AutoSize = $true
$form.Controls.Add($lblVoice)

$voiceBox = New-Object System.Windows.Forms.ComboBox
$voiceBox.DropDownStyle = "DropDown"   # pick from the list, or type a name
$voiceBox.Location = New-Object System.Drawing.Point(110, 156)
$voiceBox.Size = New-Object System.Drawing.Size(390, 28)
$form.Controls.Add($voiceBox)

$preview = New-Object System.Windows.Forms.Button
$preview.Text = "Preview"
$preview.Location = New-Object System.Drawing.Point(510, 154)
$preview.Size = New-Object System.Drawing.Size(95, 30)
$form.Controls.Add($preview)

$testAll = New-Object System.Windows.Forms.Button
$testAll.Text = "Test voices"
$testAll.Location = New-Object System.Drawing.Point(510, 195)
$testAll.Size = New-Object System.Drawing.Size(95, 30)
$form.Controls.Add($testAll)

# Speed
$lblSpeed = New-Object System.Windows.Forms.Label
$lblSpeed.Text = "Speed: 0"
$lblSpeed.Location = New-Object System.Drawing.Point(20, 203)
$lblSpeed.AutoSize = $true
$form.Controls.Add($lblSpeed)

$speed = New-Object System.Windows.Forms.TrackBar
$speed.Location = New-Object System.Drawing.Point(110, 195)
$speed.Size = New-Object System.Drawing.Size(390, 45)
$speed.Minimum = -5; $speed.Maximum = 5; $speed.Value = 0; $speed.TickFrequency = 1
$form.Controls.Add($speed)

# Output options
$grp = New-Object System.Windows.Forms.GroupBox
$grp.Text = "Output"
$grp.Location = New-Object System.Drawing.Point(20, 245)
$grp.Size = New-Object System.Drawing.Size(585, 189)
$form.Controls.Add($grp)

$rbDeck = New-Object System.Windows.Forms.RadioButton
$rbDeck.Text = "Narrated PowerPoint only"
$rbDeck.Location = New-Object System.Drawing.Point(15, 25)
$rbDeck.AutoSize = $true
$grp.Controls.Add($rbDeck)

$rbVideo = New-Object System.Windows.Forms.RadioButton
$rbVideo.Text = "Narrated PowerPoint + MP4 video"
$rbVideo.Location = New-Object System.Drawing.Point(15, 55)
$rbVideo.AutoSize = $true
$rbVideo.Checked = $true                         # default: PowerPoint + video
$grp.Controls.Add($rbVideo)

$lblRes = New-Object System.Windows.Forms.Label
$lblRes.Text = "Resolution:"
$lblRes.Location = New-Object System.Drawing.Point(330, 57)
$lblRes.AutoSize = $true
$grp.Controls.Add($lblRes)

$resBox = New-Object System.Windows.Forms.ComboBox
$resBox.DropDownStyle = "DropDownList"
$resBox.Location = New-Object System.Drawing.Point(420, 53)
$resBox.Size = New-Object System.Drawing.Size(150, 28)
$resBox.Enabled = $false
$grp.Controls.Add($resBox)

$lblFps = New-Object System.Windows.Forms.Label
$lblFps.Text = "Frame rate:"
$lblFps.Location = New-Object System.Drawing.Point(330, 124)
$lblFps.AutoSize = $true
$grp.Controls.Add($lblFps)

$fpsBox = New-Object System.Windows.Forms.ComboBox
$fpsBox.DropDownStyle = "DropDownList"
$fpsBox.Location = New-Object System.Drawing.Point(420, 120)
$fpsBox.Size = New-Object System.Drawing.Size(150, 28)
$grp.Controls.Add($fpsBox)

# Choices: label, value, ffmpeg-only?
$script:ResChoices = @(
    @{ T = "720p";              V = 720;  F = $false },
    @{ T = "1080p";             V = 1080; F = $false },
    @{ T = "1440p (ffmpeg)";    V = 1440; F = $true },
    @{ T = "4K / 2160p";        V = 2160; F = $false }
)
$script:FpsChoices = @(
    @{ T = "1 per slide (VFR)"; V = 0;  F = $true },
    @{ T = "5 fps";             V = 5;  F = $false },
    @{ T = "10 fps";            V = 10; F = $false },
    @{ T = "15 fps";            V = 15; F = $false },
    @{ T = "24 fps";            V = 24; F = $false },
    @{ T = "25 fps";            V = 25; F = $false },
    @{ T = "30 fps";            V = 30; F = $false },
    @{ T = "60 fps";            V = 60; F = $false }
)
# Fill a combo with the choices allowed for the current method, keeping the selection if possible
function Fill-Choices($box, $choices, [bool]$ffmpeg, $wanted, $fallback) {
    $keep = if ($null -ne $wanted) { $wanted } elseif ($box.SelectedItem) { $box.SelectedItem } else { $null }
    $box.Items.Clear()
    foreach ($c in $choices) { if ($ffmpeg -or -not $c.F) { [void]$box.Items.Add($c.T) } }
    if ($keep -and $box.Items.Contains($keep)) { $box.SelectedItem = $keep }
    elseif ($box.Items.Contains($fallback)) { $box.SelectedItem = $fallback }
    else { $box.SelectedIndex = 0 }
}
function Choice-Value($choices, [string]$text) { foreach ($c in $choices) { if ($c.T -eq $text) { return $c.V } }; return $null }

# ---- Hardware, for the "Auto" settings ----
$script:LogicalCores  = [Environment]::ProcessorCount
$script:PhysicalCores = $script:LogicalCores
try { $script:PhysicalCores = [int](@(Get-CimInstance Win32_Processor -ErrorAction Stop | Measure-Object -Property NumberOfCores -Sum).Sum) } catch {}
if ($script:PhysicalCores -lt 1) { $script:PhysicalCores = $script:LogicalCores }
$script:TotalMemMB = 8192
try { $script:TotalMemMB = [int]((Get-CimInstance Win32_ComputerSystem -ErrorAction Stop).TotalPhysicalMemory / 1MB) } catch {}

# Slides voiced at once: Online voices are limited by Microsoft, Native voices by CPU threads and memory
function Get-AutoSlides([string]$voice) {
    if ($voice -match 'Online') { return 4 }
    # One slide per CPU thread: each voice is single-threaded, so hyper-threads add real throughput.
    # Capped so the voices use at most half the RAM, at ~150 MB each.
    $byMem = [int][Math]::Floor(($script:TotalMemMB / 2) / 150)
    return [Math]::Max(2, [Math]::Min(32, [Math]::Min($script:LogicalCores, $byMem)))
}
# Videos built at once: VFR jobs are light; fixed-frame-rate encodes already use most cores
function Get-AutoVideos([bool]$fast, [int]$fps, [int]$decks) {
    if (-not $fast) { return 1 }
    if ($fps -eq 0) {
        $n = [Math]::Floor($script:LogicalCores / 4)                      # VFR jobs are light
        $n = [Math]::Min($n, [Math]::Floor(($script:TotalMemMB / 2) / 600)) # ~600 MB each at 4K
    } else {
        if ($script:LogicalCores -ge 16) { $n = 2 } else { $n = 1 }       # fixed-rate encodes already use most cores
        $n = [Math]::Min($n, [Math]::Floor(($script:TotalMemMB / 2) / 2000))
    }
    # No point building more videos at once than there are decks (0 = list empty, don't cap)
    if ($decks -gt 0) { $n = [Math]::Min($n, $decks) }
    return [int][Math]::Max(1, [Math]::Min(12, $n))
}

$lblPar = New-Object System.Windows.Forms.Label
$lblPar.Text = "Slides at once:"
$lblPar.Location = New-Object System.Drawing.Point(330, 27)
$lblPar.AutoSize = $true
$grp.Controls.Add($lblPar)

$parallelBox = New-Object System.Windows.Forms.NumericUpDown
$parallelBox.Location = New-Object System.Drawing.Point(445, 23)
$parallelBox.Size = New-Object System.Drawing.Size(55, 26)
$parallelBox.Minimum = 1; $parallelBox.Maximum = 32; $parallelBox.Value = 4
$grp.Controls.Add($parallelBox)

$autoSlides = New-Object System.Windows.Forms.CheckBox
$autoSlides.Text = "Auto"
$autoSlides.Location = New-Object System.Drawing.Point(507, 25)
$autoSlides.AutoSize = $true
$autoSlides.Checked = $true
$grp.Controls.Add($autoSlides)

$lblMethod = New-Object System.Windows.Forms.Label
$lblMethod.Text = "Video method:"
$lblMethod.Location = New-Object System.Drawing.Point(15, 92)
$lblMethod.AutoSize = $true
$grp.Controls.Add($lblMethod)

$methodBox = New-Object System.Windows.Forms.ComboBox
$methodBox.DropDownStyle = "DropDownList"
$methodBox.Location = New-Object System.Drawing.Point(120, 88)
$methodBox.Size = New-Object System.Drawing.Size(200, 28)
[void]$methodBox.Items.AddRange(@("Fast - ffmpeg H.264", "Fast - ffmpeg H.265 (x265, smaller)", "PowerPoint (keeps animations)"))
$methodBox.SelectedIndex = 0
$grp.Controls.Add($methodBox)

$lblVidPar = New-Object System.Windows.Forms.Label
$lblVidPar.Text = "Videos at once:"
$lblVidPar.Location = New-Object System.Drawing.Point(330, 92)
$lblVidPar.AutoSize = $true
$grp.Controls.Add($lblVidPar)

$videoParBox = New-Object System.Windows.Forms.NumericUpDown
$videoParBox.Location = New-Object System.Drawing.Point(445, 88)
$videoParBox.Size = New-Object System.Drawing.Size(55, 26)
$videoParBox.Minimum = 1; $videoParBox.Maximum = 12; $videoParBox.Value = 2
$grp.Controls.Add($videoParBox)

$autoVideos = New-Object System.Windows.Forms.CheckBox
$autoVideos.Text = "Auto"
$autoVideos.Location = New-Object System.Drawing.Point(507, 90)
$autoVideos.AutoSize = $true
$autoVideos.Checked = $true
$grp.Controls.Add($autoVideos)

$lblChap = New-Object System.Windows.Forms.Label
$lblChap.Text = "Chapters:"
$lblChap.Location = New-Object System.Drawing.Point(15, 124)
$lblChap.AutoSize = $true
$grp.Controls.Add($lblChap)

$chapterBox = New-Object System.Windows.Forms.ComboBox
$chapterBox.DropDownStyle = "DropDownList"
$chapterBox.Location = New-Object System.Drawing.Point(120, 120)
$chapterBox.Size = New-Object System.Drawing.Size(200, 28)
[void]$chapterBox.Items.AddRange(@("By section (or by slide)", "By slide", "None"))
$chapterBox.SelectedIndex = 0
$grp.Controls.Add($chapterBox)

$captionBox = New-Object System.Windows.Forms.CheckBox
$captionBox.Text = "Captions (subtitles in the video, plus .srt and .vtt files)"
$captionBox.Location = New-Object System.Drawing.Point(15, 155)
$captionBox.AutoSize = $true
$captionBox.Checked = $true
$grp.Controls.Add($captionBox)

function Update-VideoChoices($wantRes, $wantFps) {
    $ff = $methodBox.SelectedIndex -lt 2
    Fill-Choices $resBox $script:ResChoices $ff $wantRes "4K / 2160p"
    Fill-Choices $fpsBox $script:FpsChoices $ff $wantFps $(if ($ff) { "1 per slide (VFR)" } else { "15 fps" })
}
Update-VideoChoices $null $null
$methodBox.Add_SelectedIndexChanged({ Update-VideoChoices $null $null })

# Start / Cancel
$start = New-Object System.Windows.Forms.Button
$start.Text = "Start"
$start.Location = New-Object System.Drawing.Point(20, 446)
$start.Size = New-Object System.Drawing.Size(470, 40)
$start.BackColor = $navy
$start.ForeColor = [System.Drawing.Color]::White
$start.FlatStyle = "Flat"
$start.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($start)

$cancel = New-Object System.Windows.Forms.Button
$cancel.Text = "Cancel"
$cancel.Location = New-Object System.Drawing.Point(500, 446)
$cancel.Size = New-Object System.Drawing.Size(105, 40)
$cancel.Enabled = $false
$form.Controls.Add($cancel)

# Progress + log
$lblAudio = New-Object System.Windows.Forms.Label
$lblAudio.Text = "Audio: -"
$lblAudio.Location = New-Object System.Drawing.Point(20, 491)
$lblAudio.Size = New-Object System.Drawing.Size(280, 20)
$lblAudio.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Controls.Add($lblAudio)

$lblVideo = New-Object System.Windows.Forms.Label
$lblVideo.Text = "Video: -"
$lblVideo.Location = New-Object System.Drawing.Point(300, 491)
$lblVideo.Size = New-Object System.Drawing.Size(305, 20)
$lblVideo.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Controls.Add($lblVideo)

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(20, 514)
$progress.Size = New-Object System.Drawing.Size(585, 20)
$form.Controls.Add($progress)

$log = New-Object System.Windows.Forms.TextBox
$log.Location = New-Object System.Drawing.Point(20, 542)
$log.Size = New-Object System.Drawing.Size(585, 150)
$log.Multiline = $true
$log.ScrollBars = "Vertical"
$log.ReadOnly = $true
$log.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($log)

# ---- Progress view (right-hand side): one row per deck, and everything being worked on now ----
$lblDecks = New-Object System.Windows.Forms.Label
$lblDecks.Text = "Decks"
$lblDecks.Location = New-Object System.Drawing.Point(625, 14)
$lblDecks.AutoSize = $true
$lblDecks.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblDecks)

function New-ProgressList([int]$y, [int]$h, $cols) {
    $lv = New-Object System.Windows.Forms.ListView
    $lv.Location = New-Object System.Drawing.Point(625, $y)
    $lv.Size = New-Object System.Drawing.Size(520, $h)
    $lv.View = [System.Windows.Forms.View]::Details
    $lv.FullRowSelect = $true; $lv.MultiSelect = $false; $lv.HideSelection = $true
    $lv.HeaderStyle = [System.Windows.Forms.ColumnHeaderStyle]::Nonclickable
    $lv.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    foreach ($c in $cols) { [void]$lv.Columns.Add($c[0], $c[1]) }
    # Draw progress bars in cells whose Tag is a percent (0-100); everything else is drawn normally
    $lv.OwnerDraw = $true
    try { $lv.GetType().GetProperty("DoubleBuffered", [Reflection.BindingFlags]"NonPublic,Instance").SetValue($lv, $true, $null) } catch {}
    $lv.Add_DrawColumnHeader({ param($sender, $e) $e.DrawDefault = $true })
    $lv.Add_DrawItem({ param($sender, $e) })
    $lv.Add_DrawSubItem({
        param($sender, $e)
        $tag = $e.SubItem.Tag
        if ($e.ColumnIndex -gt 0 -and $tag -is [int] -and $tag -ge 0) {
            $g = $e.Graphics; $r = $e.Bounds
            $g.FillRectangle([System.Drawing.SystemBrushes]::Window, $r)
            $in = New-Object System.Drawing.Rectangle(($r.X + 3), ($r.Y + 3), ($r.Width - 6), ($r.Height - 6))
            $g.FillRectangle($script:BarBack, $in)
            $w = [int][Math]::Round($in.Width * [Math]::Min(100, $tag) / 100.0)
            if ($w -gt 0) { $g.FillRectangle($(if ($tag -ge 100) { $script:BarDone } else { $script:BarFill }), $in.X, $in.Y, $w, $in.Height) }
            $flags = [System.Windows.Forms.TextFormatFlags]::HorizontalCenter -bor [System.Windows.Forms.TextFormatFlags]::VerticalCenter -bor [System.Windows.Forms.TextFormatFlags]::SingleLine
            [System.Windows.Forms.TextRenderer]::DrawText($g, $e.SubItem.Text, $sender.Font, $in, [System.Drawing.Color]::Black, $flags)
        } else { $e.DrawDefault = $true }
    })
    $form.Controls.Add($lv)
    return $lv
}
$script:BarBack = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(228, 232, 242))
$script:BarFill = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(140, 175, 235))
$script:BarDone = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(140, 205, 150))
$lvDecks = New-ProgressList 34 200 @(@("Deck", 170), @("Audio", 120), @("Video", 110), @("Status", 115))

$lblNow = New-Object System.Windows.Forms.Label
$lblNow.Text = "Now working"
$lblNow.Location = New-Object System.Drawing.Point(625, 246)
$lblNow.AutoSize = $true
$lblNow.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblNow)
$lvNow = New-ProgressList 266 426 @(@("Working on", 185), @("Stage", 120), @("Progress", 120), @("Time", 90))

$script:ViewClock = [Diagnostics.Stopwatch]::StartNew()
function Format-Secs([double]$s) {
    if ($s -lt 0) { return "" }
    if ($s -lt 60) { return ("{0:0.0} s" -f $s) }
    return ("{0}:{1:00}" -f [int][Math]::Floor($s / 60), [int]([Math]::Floor($s) % 60))
}
function Short-Name([string]$n) { if ($n.Length -gt 24) { return $n.Substring(0, 23) + "..." }; return $n }
function Set-Cell($item, [int]$col, [string]$text, [int]$pct = -1) {
    $sub = $item.SubItems[$col]
    if ($sub.Text -ne $text) { $sub.Text = $text }
    $sub.Tag = $pct
}
# Status text for one deck row
function Get-DeckStatus($d, $c, $v) {
    $r = $script:Run
    if ($d.Status) { return $d.Status }
    if ($r.Op -and $r.Op.Deck -eq $d) { return $r.Op.What }
    if ($d.Built) {
        if ($v -and -not $v.Done) { if ($v.Proc) { return $v.Stage } else { return "Video queued" } }
        if ($v -and -not $v.Ok) { return "Video failed" }
        if ($d.Note -and $d.Note -ne "OK") { return "Done (see log)" }
        return "Done"
    }
    if ($d.PSObject.Properties["Ready"] -and -not $d.Ready) { return "Could not read" }
    if ($c[1] -gt 0 -and $c[0] -lt $c[1]) { if ($c[2] -gt 0) { return "Voicing" } else { return "Waiting" } }
    if (-not $d.PSObject.Properties["Ready"]) { return "Waiting" }
    return "Audio ready"
}
# Refresh both lists (at most 4 times a second unless forced)
function Update-ProgressView([switch]$Force) {
    $r = $script:Run
    if (-not $r -or -not $r.Decks) { return }
    if (-not $Force -and $script:ViewClock.ElapsedMilliseconds -lt 250) { return }
    $script:ViewClock.Restart()
    $now = Get-Date
    # Audio per deck: done, total, running
    $cnt = @{}
    foreach ($j in $r.Jobs) {
        $k = $j.Deck.Path
        if (-not $cnt.ContainsKey($k)) { $cnt[$k] = [int[]]@(0, 0, 0) }
        $cnt[$k][1]++
        if ($j.Done) { $cnt[$k][0]++ } elseif ($j.Proc) { $cnt[$k][2]++ }
    }
    $lvDecks.BeginUpdate()
    foreach ($d in $r.Decks) {
        if (-not $d.Row) { continue }
        $c = $cnt[$d.Path]; if (-not $c) { $c = [int[]]@(0, 0, 0) }
        if ($c[1] -gt 0) { Set-Cell $d.Row 1 "$($c[0]) / $($c[1])" ([int][Math]::Floor(100 * $c[0] / $c[1])) }
        elseif ($d.PSObject.Properties["Ready"] -and $d.Ready) { Set-Cell $d.Row 1 "no notes" }
        else { Set-Cell $d.Row 1 "" }
        $v = $null; foreach ($x in $r.Videos) { if ($x.Deck -eq $d) { $v = $x } }
        if (-not $r.Video) { Set-Cell $d.Row 2 "-" }
        elseif ($v) {
            if ($v.Done) { if ($v.Ok) { Set-Cell $d.Row 2 "Done" 100 } else { Set-Cell $d.Row 2 "Failed" } }
            elseif ($v.Proc) { Set-Cell $d.Row 2 "$($v.Pct)%" $v.Pct }
            else { Set-Cell $d.Row 2 "Queued" }
        }
        else { Set-Cell $d.Row 2 $d.PptVideo }
        Set-Cell $d.Row 3 (Get-DeckStatus $d $c $v)
    }
    $lvDecks.EndUpdate()

    # Everything running right now
    $rows = New-Object System.Collections.ArrayList
    $nVid = 0
    foreach ($v in $r.Videos) {
        if (-not $v.Proc -or $v.Done) { continue }
        $nVid++
        $t = if ($v.Eta -ge 0) { (Format-Secs $v.Eta) + " left" } else { Format-Secs ($now - $v.Started).TotalSeconds }
        [void]$rows.Add(@(((Short-Name $v.Deck.Name) + "  -  video"), $v.Stage, "$($v.Pct)%", $t, $v.Pct))
    }
    if ($r.Op) {
        $o = $r.Op; $pct = -1; $pt = ""
        if ($o.Total -gt 0) { $pct = [int][Math]::Floor(100 * $o.N / $o.Total); $pt = "$($o.N) / $($o.Total)" }
        if ($o.What -like "*PowerPoint*") { $nVid++ }
        [void]$rows.Add(@((Short-Name $o.Deck.Name), $o.What, $pt, (Format-Secs ($now - $o.Start).TotalSeconds), $pct))
    }
    $rate = 0.0; if ($r.SpeakSecs -gt 0) { $rate = $r.SpeakChars / $r.SpeakSecs }
    $nSl = 0
    foreach ($j in $r.Jobs) {
        if ($j.Done -or -not $j.Proc -or -not $j.PhaseStart) { continue }
        $nSl++
        $el = ($now - $j.PhaseStart).TotalSeconds; $pct = -1; $pt = ""
        switch ($j.Phase) {
            "speak"   { $st = "Speaking"
                        if ($rate -gt 0) { $pct = [int][Math]::Min(99, [Math]::Floor(100 * $el * $rate / [Math]::Max(1, $j.Speech.Length))); $pt = "$pct%" } }
            "measure" { $st = "Measuring volume" }
            default   { $st = "Levelling + AAC" }
        }
        if ($j.Tries -gt 1) { $st += " (retry)" }
        [void]$rows.Add(@(("{0}  -  slide {1}" -f (Short-Name $j.Deck.Name), $j.Index), $st, $pt, (Format-Secs $el), $pct))
    }
    $lvNow.BeginUpdate()
    while ($lvNow.Items.Count -gt $rows.Count) { $lvNow.Items.RemoveAt($lvNow.Items.Count - 1) }
    for ($i = 0; $i -lt $rows.Count; $i++) {
        if ($i -ge $lvNow.Items.Count) {
            $it = New-Object System.Windows.Forms.ListViewItem("")
            for ($q = 0; $q -lt 3; $q++) { [void]$it.SubItems.Add("") }
            [void]$lvNow.Items.Add($it)
        }
        $it = $lvNow.Items[$i]; $row = $rows[$i]
        if ($it.Text -ne $row[0]) { $it.Text = $row[0] }
        Set-Cell $it 1 $row[1]; Set-Cell $it 2 $row[2] $row[4]; Set-Cell $it 3 $row[3]
    }
    $lvNow.EndUpdate()
    $lblNow.Text = "Now working:  $nSl slide$(if ($nSl -ne 1) { 's' }),  $nVid video$(if ($nVid -ne 1) { 's' })"
    $lvDecks.Invalidate(); $lvNow.Invalidate()
}

$script:LogFile      = ""
$script:LogToFile    = $false
$script:StartupLines = New-Object System.Collections.Generic.List[string]
$script:Engine       = ""

function Log([string]$msg) {
    $line = (Get-Date -Format "HH:mm:ss") + "  " + $msg
    $log.AppendText($msg + [Environment]::NewLine)
    if ($script:LogToFile) {
        try { Add-Content -Path $script:LogFile -Value $line -Encoding UTF8 } catch {}
    } else {
        $script:StartupLines.Add($line)
    }
    [System.Windows.Forms.Application]::DoEvents()
}

# Start a log file at the given path (replaces an older log of the same name)
function Start-LogFile($settings, [string]$path) {
    $script:LogFile = $path
    $script:LogToFile = $false
    $head = @(
        "Slide Narrator log - " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss"),
        "Windows: " + [Environment]::OSVersion.VersionString + "   PowerShell: " + $PSVersionTable.PSVersion + "   64-bit process: " + [Environment]::Is64BitProcess,
        "Settings: " + ($settings | ConvertTo-Json -Compress),
        "----- startup -----"
    ) + $script:StartupLines + @("----- run -----")
    try { Set-Content -Path $script:LogFile -Value $head -Encoding UTF8; $script:LogToFile = $true }
    catch { $log.AppendText("Could not write log file: $($_.Exception.Message)" + [Environment]::NewLine) }
}

# ---------- Startup checks ----------
$script:BalconExe = Find-Balcon
$script:Ffmpeg    = Find-Ffmpeg $script:BalconExe
$script:StopRequested    = $false

if ($script:BalconExe) { Log "balcon: $($script:BalconExe) (backup engine)" }
else { Log "balcon.exe not found. That's OK: Windows speech will be used directly." }
$voices = Get-Voices $script:BalconExe
foreach ($v in $voices) { [void]$voiceBox.Items.Add($v) }
if ($voiceBox.Items.Count -eq 0) { Log "No voices found automatically. You can type the voice name into the Voice box." }
else { Log "Voices found: $($voiceBox.Items.Count)"; foreach ($v in $voices) { $script:StartupLines.Add("   voice: $v") } }
Log ("CPU: {0} cores / {1} threads, RAM: {2:N0} GB" -f $script:PhysicalCores, $script:LogicalCores, ($script:TotalMemMB / 1024))
if ($script:Ffmpeg) { Log "ffmpeg: found (audio saved as AAC .m4a)" } else { Log "ffmpeg: not found (audio saved as WAV, which works fine)" }

# Restore last settings
$saved = Load-Settings
if ($saved) {
    if ($saved.Voice) { if ($voiceBox.Items.Contains($saved.Voice)) { $voiceBox.SelectedItem = $saved.Voice } else { $voiceBox.Text = $saved.Voice } }
    if ($null -ne $saved.Speed) { $speed.Value = [Math]::Max(-5, [Math]::Min(5, [int]$saved.Speed)) }
    if ($saved.Video) { $rbVideo.Checked = $true } else { $rbDeck.Checked = $true }

    if ($saved.Parallel) { $parallelBox.Value = [Math]::Max(1, [Math]::Min(32, [int]$saved.Parallel)) }
    if ($null -ne $saved.AutoSlides) { $autoSlides.Checked = [bool]$saved.AutoSlides }
    if ($null -ne $saved.AutoVideos) { $autoVideos.Checked = [bool]$saved.AutoVideos }
    if ($null -ne $saved.Method -and [int]$saved.Method -lt $methodBox.Items.Count) { $methodBox.SelectedIndex = [int]$saved.Method }
    if ($saved.VideoParallel) { $videoParBox.Value = [Math]::Max(1, [Math]::Min(12, [int]$saved.VideoParallel)) }
    if ($null -ne $saved.Chapters -and [int]$saved.Chapters -lt $chapterBox.Items.Count) { $chapterBox.SelectedIndex = [int]$saved.Chapters }
    # Captions always start ticked (not restored from last time)
    Update-VideoChoices $saved.Resolution $saved.FrameRate
}
if (-not $voiceBox.Text -and $voiceBox.Items.Count -gt 0) {
    $pref = $voiceBox.Items | Where-Object { $_ -match "Australia|Natasha|William" } | Select-Object -First 1
    if ($pref) { $voiceBox.SelectedItem = $pref } else { $voiceBox.SelectedIndex = 0 }
}
$lblSpeed.Text = "Speed: $($speed.Value)"

# ---------- Events ----------
function Add-Files($paths) {
    $added = 0
    foreach ($path in @($paths)) {
        if (-not $path) { continue }
        $path = "$path".Trim('"')
        if (-not (Test-Path -LiteralPath $path)) { continue }
        $items = @()
        if ((Get-Item -LiteralPath $path).PSIsContainer) {
            $items = Get-ChildItem -LiteralPath $path -File | Where-Object { $_.Extension -match '^\.pptx?$' } | ForEach-Object { $_.FullName }
        } else { $items = @($path) }
        foreach ($f in $items) {
            $name = [IO.Path]::GetFileName($f)
            if ($name -notmatch '\.pptx?$') { continue }
            if ($name -like '~$*') { continue }                      # PowerPoint temp/lock files
            if ($name -match '_narrated\.pptx?$') { continue }      # our own output
            $full = (Resolve-Path -LiteralPath $f).Path
            if (-not $fileList.Items.Contains($full)) { [void]$fileList.Items.Add($full); $added++ }
        }
    }
    Update-DropText
    if ($added -eq 0 -and @($paths).Count -gt 0) {
        [System.Windows.Forms.MessageBox]::Show("No new PowerPoint files were added. (Files ending in _narrated are skipped.)", "Slide Narrator") | Out-Null
    }
}

function Update-DropText {
    Update-AutoValues
    $n = $fileList.Items.Count
    if ($n -eq 0) {
        $drop.Text = "Drag PowerPoint files (or a folder) here - as many as you like"
        $drop.BackColor = [System.Drawing.Color]::FromArgb(238, 241, 250)
    } else {
        $drop.Text = "$n PowerPoint file(s) ready - drag more here to add them"
        $drop.BackColor = [System.Drawing.Color]::FromArgb(255, 236, 179)
    }
}

$dragEnter = {
    param($s, $e)
    if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) { $e.Effect = "Copy" } else { $e.Effect = "None" }
}
$dragDrop = {
    param($s, $e)
    $files = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
    if ($files -and $files.Count -gt 0) { Add-Files $files }
}
$form.Add_DragEnter($dragEnter); $form.Add_DragDrop($dragDrop)
$drop.Add_DragEnter($dragEnter); $drop.Add_DragDrop($dragDrop)
$fileList.Add_DragEnter($dragEnter); $fileList.Add_DragDrop($dragDrop)

$browse.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "PowerPoint (*.pptx;*.ppt)|*.pptx;*.ppt"
    $dlg.Multiselect = $true
    if ($fileList.Items.Count -gt 0) { $dlg.InitialDirectory = Split-Path $fileList.Items[$fileList.Items.Count - 1] -Parent }
    if ($dlg.ShowDialog() -eq "OK") { Add-Files $dlg.FileNames }
})
$removeBtn.Add_Click({
    $sel = @($fileList.SelectedItems)
    foreach ($x in $sel) { $fileList.Items.Remove($x) }
    Update-DropText
})
$clearBtn.Add_Click({ $fileList.Items.Clear(); Update-DropText })
$fileList.Add_KeyDown({ param($s, $e) if ($e.KeyCode -eq "Delete") { $removeBtn.PerformClick() } })
$drop.Add_Click({ $browse.PerformClick() })

$speed.Add_ValueChanged({ $lblSpeed.Text = "Speed: $($speed.Value)" })
$rbVideo.Add_CheckedChanged({ $resBox.Enabled = $rbVideo.Checked; $fpsBox.Enabled = $rbVideo.Checked; $methodBox.Enabled = $rbVideo.Checked; $videoParBox.Enabled = $rbVideo.Checked -and -not $autoVideos.Checked; $autoVideos.Enabled = $rbVideo.Checked; $chapterBox.Enabled = $rbVideo.Checked; $captionBox.Enabled = $rbVideo.Checked })
# Show the Auto numbers straight away, and keep them up to date as settings change
function Update-AutoValues {
    if ($autoSlides.Checked) {
        $v = Get-AutoSlides "$($voiceBox.Text)"
        $parallelBox.Value = [Math]::Max($parallelBox.Minimum, [Math]::Min($parallelBox.Maximum, $v))
    }
    if ($autoVideos.Checked) {
        $fast = $methodBox.SelectedIndex -lt 2
        $f = Choice-Value $script:FpsChoices "$($fpsBox.SelectedItem)"; if ($null -eq $f) { $f = 15 }
        $v = Get-AutoVideos $fast $f $fileList.Items.Count
        $videoParBox.Value = [Math]::Max($videoParBox.Minimum, [Math]::Min($videoParBox.Maximum, $v))
    }
}
$autoSlides.Add_CheckedChanged({ $parallelBox.Enabled = -not $autoSlides.Checked; Update-AutoValues })
$autoVideos.Add_CheckedChanged({ $videoParBox.Enabled = $rbVideo.Checked -and -not $autoVideos.Checked; Update-AutoValues })
$voiceBox.Add_SelectedIndexChanged({ Update-AutoValues })
$voiceBox.Add_TextChanged({ Update-AutoValues })
$methodBox.Add_SelectedIndexChanged({ Update-AutoValues })
$fpsBox.Add_SelectedIndexChanged({ Update-AutoValues })
Update-AutoValues
$resBox.Enabled = $rbVideo.Checked; $fpsBox.Enabled = $rbVideo.Checked; $methodBox.Enabled = $rbVideo.Checked; $videoParBox.Enabled = $rbVideo.Checked -and -not $autoVideos.Checked; $autoVideos.Enabled = $rbVideo.Checked; $chapterBox.Enabled = $rbVideo.Checked; $captionBox.Enabled = $rbVideo.Checked; $parallelBox.Enabled = -not $autoSlides.Checked

$preview.Add_Click({
    if (-not $voiceBox.Text.Trim()) { return }
    $preview.Enabled = $false
    Log "Previewing $($voiceBox.Text.Trim())..."
    $pv = Join-Path $env:TEMP "slide-narrator-preview.wav"
    if (Invoke-Speech "Welcome everyone. This is how your narrated slides will sound." $pv $voiceBox.Text.Trim() $speed.Value) {
        try { (New-Object System.Media.SoundPlayer $pv).PlaySync() } catch { Log "Could not play preview: $($_.Exception.Message)" }
    } else {
        Log "Preview failed: $($script:LastSpeechError)"
    }
    $preview.Enabled = $true
})


$testAll.Add_Click({
    $testAll.Enabled = $false; $start.Enabled = $false
    $log.Clear()
    Start-LogFile ([pscustomobject]@{ Action = "Test voices" }) (Join-Path $AppDir "Slide-Narrator voice test.log")
    Log "Log file: $($script:LogFile)"
    Log "Testing every voice with every engine (this takes a minute)..."
    $tw = Join-Path $env:TEMP "slide-narrator-test.wav"
    $working = @()
    foreach ($v in @($voiceBox.Items)) {
        $res = @()
        foreach ($eng in @("sapi", "sapi32", "balcon")) {
            if (Invoke-Engine $eng "Testing one two three." $tw "$v" 0) { $res += "$eng=OK"; if (-not ($working -contains "$v")) { $working += "$v" } }
            else { $res += "$eng=FAIL"; Log ("      {0} / {1}: {2}" -f $v, $eng, $script:LastSpeechError) }
        }
        Log ("   {0}:  {1}" -f $v, ($res -join "  "))
    }
    if ($working.Count -gt 0) { Log ("Voices that work: " + ($working -join ", ")) }
    else { Log "No voices worked. See the reasons above." }
    $testAll.Enabled = $true; $start.Enabled = $true
})

$cancel.Add_Click({ $script:StopRequested = $true; Log "Cancelling after the current slide..." })

# ---------- The main job ----------
# How it works (fast mode):
#   1. Read the speaker notes from every deck (quick).
#   2. Make the audio for ALL slides of ALL decks in the background, several
#      slides at a time (separate helper processes).
#   3. Meanwhile, deck by deck: wait for that deck's audio, insert it, save,
#      and export the video. While a video exports, audio for the next decks
#      keeps being made in the background.

function Log-To([string]$path, [string]$msg) {
    $log.AppendText($msg + [Environment]::NewLine)
    if ($path) { try { Add-Content -Path $path -Value ((Get-Date -Format "HH:mm:ss") + "  " + $msg) -Encoding UTF8 } catch {} }
}

function New-DeckLog([string]$path, $settings) {
    $head = @(
        "Slide Narrator log - " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss"),
        "Windows: " + [Environment]::OSVersion.VersionString + "   PowerShell: " + $PSVersionTable.PSVersion + "   64-bit process: " + [Environment]::Is64BitProcess,
        "Settings: " + ($settings | ConvertTo-Json -Compress),
        "----- startup -----"
    ) + $script:StartupLines + @("----- run -----")
    try { Set-Content -Path $path -Value $head -Encoding UTF8 } catch {}
}

# Read and clean the notes of every slide (PowerPoint, read-only, no window)
function Get-DeckNotes($ppt, [string]$pptx) {
    $list = @()
    $pres = $ppt.Presentations.Open($pptx, -1, 0, 0)
    try {
        foreach ($slide in $pres.Slides) {
            $notes = ""
            foreach ($shp in $slide.NotesPage.Shapes) {
                if ($shp.Type -eq 14 -and $shp.PlaceholderFormat.Type -eq 2 -and $shp.HasTextFrame) { $notes = $shp.TextFrame.TextRange.Text }
            }
            $notes = ($notes -replace "`r", "`r`n").Trim()
            $speech = ""; $caption = ""
            if ($notes) { $speech = Clean-Text $notes; $caption = Clean-Caption $notes }
            $list += [pscustomobject]@{ Index = $slide.SlideIndex; Speech = $speech; Caption = $caption }
        }
        return ,$list
    } finally { try { $pres.Close() } catch {} }
}

function Quote([string]$a) { return '"' + ($a -replace '"', '') + '"' }

# Start one background audio job
function Start-AudioJob($job) {
    [IO.File]::WriteAllText($job.Txt, $job.Speech, (New-Object System.Text.UTF8Encoding($true)))
    if (Test-Path $job.Wav) { Remove-Item $job.Wav -Force -ErrorAction SilentlyContinue }
    if (Test-Path $job.M4a) { Remove-Item $job.M4a -Force -ErrorAction SilentlyContinue }
    foreach ($f in @($job.NormWav, $job.Loud)) { if (Test-Path $f) { Remove-Item $f -Force -ErrorAction SilentlyContinue } }
    $job.Phase = "speak"
    $voice = $script:Run.Voice; $rate = $script:Run.Rate
    switch ($script:Engine) {
        "balcon" { $exe = $script:BalconExe
                   $argLine = "-f " + (Quote $job.Txt) + " -w " + (Quote $job.Wav) + " -n " + (Quote $voice) + " -s $rate -enc utf8" }
        "sapi32" { $exe = $script:Ps32
                   $argLine = "-NoProfile -ExecutionPolicy Bypass -File " + (Quote $script:Sapi32Helper) + " " + (Quote $job.Txt) + " " + (Quote $job.Wav) + " " + (Quote $voice) + " $rate" }
        default  { $exe = Join-Path $PSHOME "powershell.exe"
                   $argLine = "-NoProfile -ExecutionPolicy Bypass -File " + (Quote $script:Sapi32Helper) + " " + (Quote $job.Txt) + " " + (Quote $job.Wav) + " " + (Quote $voice) + " $rate" }
    }
    $job.Tries++
    $job.PhaseStart = Get-Date
    $job.Proc = Start-Process -FilePath $exe -ArgumentList $argLine -WindowStyle Hidden -PassThru
}

# ---- Even volume: every slide is measured and set to the same loudness ----
$script:LoudTarget = -19          # LUFS (the usual level for mono speech)
$script:LoudMeasure = "loudnorm=I=-19:TP=-1.5:LRA=11:print_format=json"
# Gain in dB from ffmpeg's loudness report (0 if it can't be read)
function Get-LoudGain([string]$report) {
    try {
        $m = [regex]::Match($report, '\{[^{}]*"input_i"[^{}]*\}')
        if (-not $m.Success) { return 0.0 }
        $i = [double]::Parse(($m.Value | ConvertFrom-Json).input_i, [Globalization.CultureInfo]::InvariantCulture)
        if ($i -lt -70) { return 0.0 }
        return [Math]::Max(-20.0, [Math]::Min(20.0, $script:LoudTarget - $i))
    } catch { return 0.0 }
}
# Gain, then a limiter so the loudest peaks stay just under -1.5 dB
function Get-LevelFilter([double]$gain) { return "volume=" + $gain.ToString("0.00", [Globalization.CultureInfo]::InvariantCulture) + "dB,alimiter=limit=0.8414:level=0" }
# One ffmpeg run: the levelled WAV (for the video) and the AAC copy (for the PowerPoint)
function Get-LevelArgs([string]$wav, [string]$norm, [string]$m4a, [double]$gain) {
    $af = Get-LevelFilter $gain
    return "-y -loglevel error -i " + (Quote $wav) + " -af $af -c:a pcm_s16le " + (Quote $norm) + " -af $af -c:a aac -b:a 64k -ac 1 " + (Quote $m4a)
}
# Same thing, done straight away (used when a slide's audio had to be remade)
function Set-SlideLevel([string]$wav, [string]$m4a) {
    $norm = [IO.Path]::ChangeExtension($wav, ".level.wav")
    $report = (& $script:Ffmpeg -hide_banner -nostats -i $wav -af $script:LoudMeasure -f null - 2>&1 | Out-String)
    $af = Get-LevelFilter (Get-LoudGain $report)
    & $script:Ffmpeg -y -loglevel error -i $wav -af $af -c:a pcm_s16le $norm -af $af -c:a aac -b:a 64k -ac 1 $m4a 2>$null | Out-Null
    if ((Test-Path $norm) -and (Get-Item $norm).Length -gt 1000) { Move-Item -LiteralPath $norm -Destination $wav -Force }
    if (-not ((Test-Path $m4a) -and (Get-Item $m4a).Length -gt 500)) {
        & $script:Ffmpeg -y -loglevel error -i $wav -c:a aac -b:a 64k -ac 1 $m4a 2>$null | Out-Null
    }
}

# Check running jobs, start new ones, update progress. Call this often.
# Read a small text file without locking it (the video worker may be writing it)
function Read-Shared([string]$path) {
    try {
        $fs = [IO.File]::Open($path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete)
        try { return (New-Object IO.StreamReader($fs)).ReadToEnd().Trim() } finally { $fs.Dispose() }
    } catch { return "" }
}

function Pump-Video {
    $r = $script:Run
    $running = 0; $queued = 0; $done = 0; $status = @()
    foreach ($v in $r.Videos) {
        if ($v.Done) { $done++; continue }
        if (-not $v.Proc) { $queued++; continue }
        if ($v.Proc.HasExited) {
            $v.Done = $true; $done++
            $p = Read-Shared $v.Progress
            $code = -1; try { $code = $v.Proc.ExitCode } catch {}
            $v.Ok = (($p -eq "done") -or ($code -eq 0)) -and (Test-Path -LiteralPath $v.Out)
            if ($v.Ok) { $v.Deck.VideoNote = "video OK"; Log-To $v.Deck.Log "Saved video: $($v.Out)" }
            else { $v.Deck.VideoNote = "video FAILED"; Log-To $v.Deck.Log "Video failed for $($v.Deck.Name): $p" }
        } else {
            $running++
            $p = Read-Shared $v.Progress
            if ($p -match '^(.*?)\|(\d+)\|(-?\d+)$') { $v.Stage = $matches[1]; $v.Pct = [int]$matches[2]; $v.Eta = [int]$matches[3] }
            $status += "$($v.Deck.Name) $($v.Pct)%"
        }
    }
    foreach ($v in $r.Videos) {
        if ($running -ge $r.VideoParallel) { break }
        if (-not $v.Done -and -not $v.Proc) {
            $exe = Join-Path $PSHOME "powershell.exe"
            $v.Proc = Start-Process -FilePath $exe -ArgumentList ("-NoProfile -ExecutionPolicy Bypass -File " + (Quote $script:VideoWorker) + " " + (Quote $v.JobFile)) -WindowStyle Hidden -PassThru
            $v.Started = Get-Date; $v.Stage = "Starting"
            Log-To $v.Deck.Log "Building video with ffmpeg: $($v.Deck.Name)"
            $running++; $queued--
        }
    }
    if ($r.Videos.Count -gt 0) {
        $lblVideo.Text = "Video: $running building, $queued waiting, $done done"
    }
}

function Pump-Audio {
    $r = $script:Run
    if ($script:StopRequested) {
        foreach ($j in $r.Jobs) { if ($j.Proc -and -not $j.Proc.HasExited) { try { $j.Proc.Kill() } catch {} } }
        foreach ($v in $r.Videos) { if ($v.Proc -and -not $v.Proc.HasExited) { try { $v.Proc.Kill() } catch {}; Get-Process ffmpeg -ErrorAction SilentlyContinue | Where-Object { $_.StartTime -ge $v.Proc.StartTime } | ForEach-Object { try { $_.Kill() } catch {} } } }
        throw "Cancelled."
    }
    Pump-Video
    $running = 0
    foreach ($j in $r.Jobs) {
        if ($j.Done -or -not $j.Proc) { continue }
        if ($j.Proc.HasExited) {
            if ($j.Phase -eq "convert") {
                # Levelled WAV + AAC done. If levelling failed, the original WAV is used as it is.
                if ((Test-Path $j.NormWav) -and (Get-Item $j.NormWav).Length -gt 1000) {
                    try { Move-Item -LiteralPath $j.NormWav -Destination $j.Wav -Force } catch {}
                }
                $j.Done = $true; $j.Ok = $true; $r.DoneCount++; continue
            }
            if ($j.Phase -eq "measure") {
                # Loudness measured: set the level and make the AAC copy for the PowerPoint
                $gain = Get-LoudGain (Read-Shared $j.Loud)
                $j.Phase = "convert"; $j.PhaseStart = Get-Date
                $j.Proc = Start-Process -FilePath $script:Ffmpeg -ArgumentList (Get-LevelArgs $j.Wav $j.NormWav $j.M4a $gain) -WindowStyle Hidden -PassThru
                $running++; continue
            }
            $ok = (Test-Path $j.Wav) -and ((Get-Item $j.Wav).Length -gt 1000)
            if ($ok -and $j.PhaseStart) {   # speaking speed, used to estimate progress of the others
                $r.SpeakChars += $j.Speech.Length; $r.SpeakSecs += ((Get-Date) - $j.PhaseStart).TotalSeconds
            }
            if ($ok -and $script:Ffmpeg) {
                # Speech done: measure its loudness in the background (report saved to a text file)
                $j.Phase = "measure"; $j.PhaseStart = Get-Date
                $argLine = '/d /c "' + (Quote $script:Ffmpeg) + ' -hide_banner -nostats -i ' + (Quote $j.Wav) + ' -af ' + $script:LoudMeasure + ' -f null - 2>' + (Quote $j.Loud) + '"'
                $j.Proc = Start-Process -FilePath $env:ComSpec -ArgumentList $argLine -WindowStyle Hidden -PassThru
                $running++
            }
            elseif ($ok) { $j.Done = $true; $j.Ok = $true; $r.DoneCount++ }
            elseif ($j.Tries -lt 2) { Start-AudioJob $j; $running++ }
            else { $j.Done = $true; $j.Ok = $false; $r.DoneCount++
                   Log-To $j.Deck.Log ("   [{0}] slide {1}: background audio failed, will retry when inserting" -f $j.Deck.Name, $j.Index) }
        } else { $running++ }
    }
    foreach ($j in $r.Jobs) {
        if ($running -ge $r.Parallel) { break }
        if (-not $j.Done -and -not $j.Proc) { Start-AudioJob $j; $running++ }
    }
    if ($r.Jobs.Count -gt 0) { $progress.Maximum = $r.Jobs.Count; $progress.Value = [Math]::Min($r.Jobs.Count, $r.DoneCount) }
    $lblAudio.Text = "Audio: $($r.DoneCount) of $($r.Jobs.Count) slides"
    Update-ProgressView
    [System.Windows.Forms.Application]::DoEvents()
}

function Wait-DeckAudio($deck) {
    while (@($script:Run.Jobs | Where-Object { $_.Deck -eq $deck -and -not $_.Done }).Count -gt 0) {
        Pump-Audio
        Start-Sleep -Milliseconds 150
    }
}

# Work out where chapters start: slide index -> chapter title
function Get-ChapterMap($pres, [int]$mode) {
    $map = @{}
    if ($mode -eq 2) { return $map }                       # None
    $useSections = $false
    if ($mode -eq 0) { try { $useSections = $pres.SectionProperties.Count -ge 2 } catch {} }
    if ($useSections) {
        $sp = $pres.SectionProperties
        for ($k = 1; $k -le $sp.Count; $k++) {
            if ($sp.SlidesCount($k) -gt 0) { $map[[int]$sp.FirstSlide($k)] = $sp.Name($k) }
        }
    } else {
        foreach ($slide in $pres.Slides) {
            $title = ""
            try { if ($slide.Shapes.HasTitle) { $title = $slide.Shapes.Title.TextFrame.TextRange.Text } } catch {}
            $title = ($title -replace '[\r\n\v]+', ' ').Trim()
            if (-not $title) { $title = "Slide $($slide.SlideIndex)" }
            $map[[int]$slide.SlideIndex] = "$($slide.SlideIndex). $title"
        }
    }
    return $map
}

# Add chapters, a chapter list and captions to a video PowerPoint has exported (uses the video worker)
function Add-VideoExtras($deck, $segments) {
    $r = $script:Run
    $vjob = [pscustomobject]@{
        Ffmpeg = $script:Ffmpeg; Fps = 0; Codec = ""; WorkDir = $deck.WorkDir; Log = $deck.Log; Progress = ""
        Out = $deck.OutMp4; Segments = $segments; Captions = $r.Captions; Mode = "extras"; Threads = [Math]::Max(2, $script:LogicalCores)
    }
    $jf = Join-Path $deck.WorkDir "video-extras.json"
    [IO.File]::WriteAllText($jf, ($vjob | ConvertTo-Json -Depth 5), (New-Object System.Text.UTF8Encoding($false)))
    $lblVideo.Text = "Video: adding chapters and captions to $($deck.Name)"
    $r.Op = [pscustomobject]@{ Deck = $deck; What = "Adding captions/chapters"; N = 0; Total = 0; Start = (Get-Date) }
    $p = Start-Process -FilePath (Join-Path $PSHOME "powershell.exe") -ArgumentList ("-NoProfile -ExecutionPolicy Bypass -File " + (Quote $script:VideoWorker) + " " + (Quote $jf)) -WindowStyle Hidden -PassThru
    while (-not $p.HasExited) { Pump-Audio; Start-Sleep -Milliseconds 200 }
    $r.Op = $null
    if ($p.ExitCode -ne 0) { Log "Could not add chapters or captions to the video (the video itself is fine). See the log for details." }
    $lblVideo.Text = "Video: -"
}

# Insert audio into one deck, save it, export the video
function Build-Deck($ppt, $deck) {
    $r = $script:Run
    $msoTrue = -1; $msoFalse = 0
    $script:LogFile = $deck.Log; $script:LogToFile = $true      # Log() now writes to this deck's log
    Log "Inserting audio into $($deck.Name)..."
    $pres = $ppt.Presentations.Open($deck.Path, $msoTrue, $msoFalse, $msoTrue)
    $failed = @(); $videoOk = $true; $segments = @()
    $chap = Get-ChapterMap $pres $r.Chapters
    $r.Op = [pscustomobject]@{ Deck = $deck; What = "Inserting audio"; N = 0; Total = $pres.Slides.Count; Start = (Get-Date) }
    try {
        foreach ($slide in $pres.Slides) {
            Pump-Audio
            $i = $slide.SlideIndex
            $r.Op.N = $i
            for ($s = $slide.Shapes.Count; $s -ge 1; $s--) {
                if ($slide.Shapes.Item($s).Name -eq "Narration") { $slide.Shapes.Item($s).Delete() }
            }
            $t = $slide.SlideShowTransition
            $t.AdvanceOnClick = $msoFalse; $t.AdvanceOnTime = $msoTrue
            $job = $r.Jobs | Where-Object { $_.Deck -eq $deck -and $_.Index -eq $i } | Select-Object -First 1
            $img = Join-Path $deck.WorkDir ("slide_{0:D2}.png" -f $i)
            if (-not $job) { $t.AdvanceTime = $r.NoNotes; $segments += [pscustomobject]@{ Index = $i; Img = $img; Wav = ""; WavDur = 0.0; Dur = [double]$r.NoNotes; Chapter = [string]$chap[[int]$i]; Caption = "" }; continue }

            if (-not $job.Ok) {
                Log "   slide ${i}: retrying audio..."
                $job.Ok = New-SlideAudio $job.Speech $job.Wav $r.Voice $r.Rate
            }
            if (-not $job.Ok) {
                $secs = [Math]::Max(5, [int](($job.Speech -split '\s+').Count / 2.5))
                $t.AdvanceTime = $secs; $failed += $i
                $segments += [pscustomobject]@{ Index = $i; Img = $img; Wav = ""; WavDur = 0.0; Dur = [double]$secs; Chapter = [string]$chap[[int]$i]; Caption = "" }
                Log ("   FAILED. Slide {0} has no audio and will show for {1} s. Reason: {2}" -f $i, $secs, $script:LastSpeechError)
                continue
            }
            # Use the levelled AAC file made in the background; make it now only if it's missing
            $audio = $job.Wav
            if ($script:Ffmpeg) {
                if (-not ((Test-Path $job.M4a) -and (Get-Item $job.M4a).Length -gt 500)) { Set-SlideLevel $job.Wav $job.M4a }
                if ((Test-Path $job.M4a) -and (Get-Item $job.M4a).Length -gt 500) { $audio = $job.M4a }
            }
            $media = $slide.Shapes.AddMediaObject2($audio, $msoFalse, $msoTrue, 10, 10, 40, 40)
            $media.Name = "Narration"
            $play = $media.AnimationSettings.PlaySettings
            $play.PlayOnEntry = $msoTrue; $play.HideWhileNotPlaying = $msoTrue
            $t.AdvanceTime = $r.Pause
            $wavSecs = Get-WavSeconds $job.Wav
            $segments += [pscustomobject]@{ Index = $i; Img = $img; Wav = $job.Wav; WavDur = [Math]::Round($wavSecs, 3); Dur = [Math]::Round($wavSecs + $r.Pause, 3); Chapter = [string]$chap[[int]$i]; Caption = [string]$job.Caption }
        }
        $r.Op = [pscustomobject]@{ Deck = $deck; What = "Saving PowerPoint"; N = 0; Total = 0; Start = (Get-Date) }
        Pump-Audio
        $pres.SaveAs($deck.OutPptx)
        Log "Saved: $($deck.OutPptx)"

        if ($r.Video -and $r.Fast) {
            # Export each slide as a picture, then hand the rest to ffmpeg in the background
            $h = [int]$r.VRes
            $w = [int]([Math]::Round($h * $pres.PageSetup.SlideWidth / $pres.PageSetup.SlideHeight / 2) * 2)
            Log "Exporting slide pictures ($w x $h)..."
            $r.Op = [pscustomobject]@{ Deck = $deck; What = "Exporting pictures"; N = 0; Total = $pres.Slides.Count; Start = (Get-Date) }
            foreach ($slide in $pres.Slides) {
                $r.Op.N = $slide.SlideIndex
                Pump-Audio
                $slide.Export((Join-Path $deck.WorkDir ("slide_{0:D2}.png" -f $slide.SlideIndex)), "PNG", $w, $h)
            }
            $vjob = [pscustomobject]@{
                Ffmpeg = $script:Ffmpeg; Fps = $r.Fps; Codec = $r.Codec; WorkDir = $deck.WorkDir; Log = $deck.Log
                Progress = (Join-Path $deck.WorkDir "video-progress.txt"); Out = $deck.OutMp4; Segments = $segments
                Captions = $r.Captions; Mode = "build"
                Threads = [Math]::Max(2, [int][Math]::Floor($script:LogicalCores / [Math]::Max(1, $r.VideoParallel)))
            }
            $jf = Join-Path $deck.WorkDir "video-job.json"
            [IO.File]::WriteAllText($jf, ($vjob | ConvertTo-Json -Depth 5), (New-Object System.Text.UTF8Encoding($false)))
            if (Test-Path $vjob.Progress) { Remove-Item $vjob.Progress -Force }
            if (Test-Path $deck.OutMp4) { Remove-Item $deck.OutMp4 -Force -ErrorAction SilentlyContinue }
            $script:Run.Videos += [pscustomobject]@{ Deck = $deck; JobFile = $jf; Progress = $vjob.Progress; Out = $deck.OutMp4; Proc = $null; Done = $false; Ok = $false; Stage = "Queued"; Pct = 0; Eta = -1; Started = $null }
            Log "Video queued ($($r.CodecName)). It will be built in the background."
            $deck.VideoNote = "video queued"
        }
        elseif ($r.Video) {
            Log "Exporting video ($($r.VRes)p, $($r.Fps) fps)... audio for the next decks keeps going meanwhile."
            $r.Op = [pscustomobject]@{ Deck = $deck; What = "PowerPoint video export"; N = 0; Total = 0; Start = (Get-Date) }
            $deck.PptVideo = "Exporting"
            $pres.CreateVideo($deck.OutMp4, $true, $r.NoNotes, $r.VRes, $r.Fps, 85)
            $lblVideo.Text = "Video: exporting $($deck.Name)"
            do {
                for ($k = 0; $k -lt 8; $k++) { Pump-Audio; Start-Sleep -Milliseconds 250 }
                $status = $pres.CreateVideoStatus      # 1 in progress, 2 queued, 3 done, 4 failed
            } while ($status -eq 1 -or $status -eq 2)
            $lblVideo.Text = "Video: -"
            $deck.PptVideo = $(if ($status -eq 3) { "Done" } else { "Failed" })
            if ($status -eq 3) { Log "Saved video: $($deck.OutMp4)"; if ($script:Ffmpeg -and ($r.Chapters -ne 2 -or $r.Captions)) { Add-VideoExtras $deck $segments } }
            else { $videoOk = $false; Log "Video export failed (status $status). Try 1080p, or open the narrated deck and use File > Export > Create a Video." }
        }
    } finally {
        $r.Op = $null
        try { $pres.Close() } catch {}
        $script:LogToFile = $false
    }
    $note = "OK"
    if ($failed.Count -gt 0) { $note = "no audio for slide(s) " + ($failed -join ", ") }
    if (-not $videoOk) { $note += "; video export failed" }
    Log-To $deck.Log "Finished: $note"
    return $note
}

$start.Add_Click({
    if ($fileList.Items.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Drag one or more PowerPoint files onto the window first.", "Slide Narrator") | Out-Null; return }
    if (-not $voiceBox.Text.Trim()) { [System.Windows.Forms.MessageBox]::Show("Choose or type a voice first.", "Slide Narrator") | Out-Null; return }

    $settings = [pscustomobject]@{ Voice = "$($voiceBox.Text.Trim())"; Speed = $speed.Value; Video = $rbVideo.Checked; Resolution = "$($resBox.SelectedItem)"; FrameRate = "$($fpsBox.SelectedItem)"; Parallel = [int]$parallelBox.Value; Method = $methodBox.SelectedIndex; VideoParallel = [int]$videoParBox.Value; Chapters = $chapterBox.SelectedIndex; Captions = $captionBox.Checked; AutoSlides = $autoSlides.Checked; AutoVideos = $autoVideos.Checked }
    $vres = Choice-Value $script:ResChoices "$($resBox.SelectedItem)"; if (-not $vres) { $vres = 1080 }
    $fps  = Choice-Value $script:FpsChoices "$($fpsBox.SelectedItem)"; if ($null -eq $fps) { $fps = 15 }
    $fast = $settings.Video -and ($methodBox.SelectedIndex -lt 2)
    if ($fast -and -not $script:Ffmpeg) {
        [System.Windows.Forms.MessageBox]::Show("The fast video method needs ffmpeg.exe, which wasn't found. PowerPoint will export the videos instead.", "Slide Narrator") | Out-Null
        $fast = $false
        if ($fps -eq 0) { $fps = 15 }
        if ($vres -eq 1440) { $vres = 1080 }
    }
    Save-Settings $settings

    $script:StopRequested = $false
    $controls = @($start, $browse, $removeBtn, $clearBtn, $fileList, $voiceBox, $preview, $testAll, $speed, $grp)
    foreach ($c in $controls) { $c.Enabled = $false }
    $cancel.Enabled = $true
    $form.AllowDrop = $false; $drop.AllowDrop = $false
    $log.Clear()

    $slidesAtOnce = [int]$settings.Parallel
    if ($settings.AutoSlides) { $slidesAtOnce = Get-AutoSlides $settings.Voice; $parallelBox.Value = $slidesAtOnce }
    $videosAtOnce = [int]$settings.VideoParallel
    if ($settings.AutoVideos) { $videosAtOnce = Get-AutoVideos $fast $fps $fileList.Items.Count; $videoParBox.Value = $videosAtOnce }

    $script:Run = [pscustomobject]@{
        Voice = $settings.Voice; Rate = $settings.Speed; Video = $settings.Video; VRes = $vres; Fps = $fps
        Parallel = $slidesAtOnce; Pause = 1; NoNotes = 4; Jobs = @(); DoneCount = 0
        Fast = $fast; Codec = $(if ($methodBox.SelectedIndex -eq 1) { "h265" } else { "h264" }); CodecName = $methodBox.Text
        Videos = @(); VideoParallel = $videosAtOnce; Chapters = $settings.Chapters
        Captions = [bool]($settings.Captions -and $settings.Video -and $script:Ffmpeg)
        Decks = @(); Op = $null; SpeakChars = 0.0; SpeakSecs = 0.0
    }
    $decks = @()
    foreach ($p in @($fileList.Items)) {
        $path = "$p"; $dir = Split-Path $path -Parent; $name = [IO.Path]::GetFileNameWithoutExtension($path)
        $decks += [pscustomobject]@{
            Path = $path; Name = $name; Log = (Join-Path $dir ($name + "_narration.log"))
            WorkDir = (Join-Path $dir ($name + "_narration"))
            OutPptx = (Join-Path $dir ($name + "_narrated.pptx")); OutMp4 = (Join-Path $dir ($name + ".mp4"))
            VideoNote = ""; Row = $null; Built = $false; Status = ""; PptVideo = ""
        }
    }
    $script:Run.Decks = $decks
    $lvDecks.Items.Clear(); $lvNow.Items.Clear()
    foreach ($d in $decks) {
        $it = New-Object System.Windows.Forms.ListViewItem((Short-Name $d.Name))
        for ($q = 0; $q -lt 3; $q++) { [void]$it.SubItems.Add("") }
        [void]$lvDecks.Items.Add($it); $d.Row = $it
    }
    Update-ProgressView -Force
    $results = @()
    $ppt = $null
    try {
        # Check the voice works and pick the fastest engine for it
        $log.AppendText("Checking the voice..." + [Environment]::NewLine)
        $pv = Join-Path $env:TEMP "slide-narrator-check.wav"
        if (-not (Invoke-Speech "Checking." $pv $script:Run.Voice $script:Run.Rate)) {
            throw "The voice '$($script:Run.Voice)' isn't working. Try 'Test voices'. Reason: $($script:LastSpeechError)"
        }
        $log.AppendText("Voice OK (engine: $($script:Engine)). CPU: $($script:PhysicalCores) cores / $($script:LogicalCores) threads." + [Environment]::NewLine)
        $log.AppendText("Slides at once: $($script:Run.Parallel)$(if ($settings.AutoSlides) { ' (auto)' })   Videos at once: $($script:Run.VideoParallel)$(if ($settings.AutoVideos) { ' (auto)' })" + [Environment]::NewLine)

        $ppt = New-Object -ComObject PowerPoint.Application
        $ppt.Visible = -1

        # 1. Read notes from every deck and queue the audio jobs
        foreach ($d in $decks) {
            New-DeckLog $d.Log ([pscustomobject]@{ File = $d.Path; Voice = $settings.Voice; Speed = $settings.Speed; Video = $settings.Video; Resolution = $settings.Resolution; FrameRate = $settings.FrameRate; Method = $methodBox.Text; SlidesAtOnce = $slidesAtOnce; VideosAtOnce = $videosAtOnce; Cores = "$($script:PhysicalCores)/$($script:LogicalCores)" })
            Log-To $d.Log "Reading notes: $($d.Name)"
            New-Item -ItemType Directory -Force -Path $d.WorkDir | Out-Null
            try {
                $notes = Get-DeckNotes $ppt $d.Path
                $d | Add-Member -NotePropertyName Ready -NotePropertyValue $true
                foreach ($n in $notes) {
                    if (-not $n.Speech) { continue }
                    $tag = "slide_{0:D2}" -f $n.Index
                    $script:Run.Jobs += [pscustomobject]@{
                        Deck = $d; Index = $n.Index; Speech = $n.Speech; Caption = $n.Caption
                        Txt = (Join-Path $d.WorkDir "$tag.txt"); Wav = (Join-Path $d.WorkDir "$tag.wav"); M4a = (Join-Path $d.WorkDir "$tag.m4a"); Phase = "speak"
                        NormWav = (Join-Path $d.WorkDir "$tag.level.wav"); Loud = (Join-Path $d.WorkDir "$tag.loudness.txt")
                        Proc = $null; Tries = 0; Done = $false; Ok = $false; PhaseStart = $null
                    }
                }
                Log-To $d.Log ("   {0} slides, {1} with notes" -f $notes.Count, @($notes | Where-Object { $_.Speech }).Count)
            } catch {
                $d | Add-Member -NotePropertyName Ready -NotePropertyValue $false -Force
                Log-To $d.Log "   Could not read this deck: $($_.Exception.Message)"
            }
            Pump-Audio      # start making audio straight away
        }

        # 2. Build each deck as soon as its audio is ready
        for ($k = 0; $k -lt $decks.Count; $k++) {
            $d = $decks[$k]
            $form.Text = "Slide Narrator - deck $($k + 1) of $($decks.Count)"
            if (-not $d.Ready) { $results += "FAILED  $($d.Name)  (could not read the deck)"; continue }
            Log-To $d.Log "===== Deck $($k + 1) of $($decks.Count): $($d.Name) - waiting for its audio ====="
            try {
                Wait-DeckAudio $d
                $note = Build-Deck $ppt $d
                $d | Add-Member -NotePropertyName Note -NotePropertyValue $note -Force
                $d.Built = $true
                $results += $d
            } catch {
                $msg = $_.Exception.Message
                if ($msg -eq "Cancelled.") { throw }
                $d.Status = "Failed"
                Log-To $d.Log "STOPPED: $msg"
                Log-To $d.Log ("Details: " + ($_ | Out-String).Trim())
                Log-To $d.Log ("Where: " + $_.ScriptStackTrace)
                $results += "FAILED  $($d.Name)  ($msg)"
            }
        }
        # 3. Wait for any background videos to finish
        if ($script:Run.Videos.Count -gt 0) {
            $form.Text = "Slide Narrator - finishing videos"
            $log.AppendText("Waiting for the videos to finish..." + [Environment]::NewLine)
            while (@($script:Run.Videos | Where-Object { -not $_.Done }).Count -gt 0) { Pump-Audio; Start-Sleep -Milliseconds 300 }
        }
    } catch {
        $msg = $_.Exception.Message
        if ($msg -eq "Cancelled.") {
            $results += "CANCELLED (remaining decks were not finished)"
            foreach ($d in $decks) { if (-not $d.Built -and -not $d.Status) { $d.Status = "Cancelled" } }
            foreach ($v in $script:Run.Videos) { if (-not $v.Done) { $v.Done = $true; $v.Ok = $false } }
        }
        else {
            $results += "STOPPED: $msg"
            foreach ($d in $decks) { Log-To $d.Log "STOPPED: $msg"; Log-To $d.Log ("Where: " + $_.ScriptStackTrace) }
        }
    } finally {
        foreach ($j in $script:Run.Jobs) { if ($j.Proc -and -not $j.Proc.HasExited) { try { $j.Proc.Kill() } catch {} } }
        foreach ($v in $script:Run.Videos) { if ($v.Proc -and -not $v.Proc.HasExited) { try { $v.Proc.Kill() } catch {} } }
        if ($ppt) { try { $ppt.Quit() } catch {}; try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) } catch {} }
        $script:LogToFile = $false
        foreach ($c in $controls) { $c.Enabled = $true }
        $resBox.Enabled = $rbVideo.Checked; $fpsBox.Enabled = $rbVideo.Checked; $methodBox.Enabled = $rbVideo.Checked; $videoParBox.Enabled = $rbVideo.Checked -and -not $autoVideos.Checked; $autoVideos.Enabled = $rbVideo.Checked; $chapterBox.Enabled = $rbVideo.Checked; $captionBox.Enabled = $rbVideo.Checked; $parallelBox.Enabled = -not $autoSlides.Checked
        $cancel.Enabled = $false
        $form.AllowDrop = $true; $drop.AllowDrop = $true
        $form.Text = "Slide Narrator"
        $lblVideo.Text = "Video: -"
        $script:Run.Op = $null
        foreach ($j in $script:Run.Jobs) { $j.Proc = $null }
        foreach ($v in $script:Run.Videos) { if (-not $v.Done) { $v.Done = $true } }
        try { Update-ProgressView -Force } catch {}
    }
    $lines = foreach ($x in $results) {
        if ($x -is [string]) { $x }
        else { $v = $x.VideoNote; if ($v -eq "video queued") { $v = "video not finished" }; "OK      $($x.Name)  ($($x.Note)$(if ($v) { '; ' + $v }))" }
    }
    $summary = "Summary:`r`n" + ($lines -join "`r`n")
    $log.AppendText("`r`n" + $summary + "`r`n")
    [System.Windows.Forms.MessageBox]::Show($summary + "`n`nEach deck's log is saved next to it as <deck name>_narration.log", "Slide Narrator") | Out-Null
})

if ($StartFiles) { Add-Files $StartFiles }
[void]$form.ShowDialog()
