<#
  Slide Narrator
  ------------------------------------------------------------------
  A small Windows app. Drag a PowerPoint file onto the window, pick a
  voice, and click Start. For every slide it:
    1. reads the speaker notes
    2. turns them into audio with Balabolka's balcon.exe
    3. inserts the audio into the slide (plays automatically, icon hidden)
    4. sets the slide to advance when the audio ends
  It saves a new file  <deck>_narrated.pptx  next to the original
  (the original is not changed), and can also export an MP4.

  Start it by double-clicking  Slide-Narrator.bat
  You can also drag a .pptx straight onto Slide-Narrator.bat.

  Needs: Windows, PowerPoint, and balcon.exe
  (https://www.cross-plus-a.com/bconsole.htm). Put balcon.exe next to
  this file, in a "balcon" subfolder, or in C:\balcon.
  ffmpeg is optional (smaller AAC audio instead of WAV, plus fast video).
  ------------------------------------------------------------------
#>

param([Parameter(ValueFromRemainingArguments = $true)][string[]]$StartFiles)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$AppDir       = $PSScriptRoot
$SettingsFile = Join-Path $AppDir "Slide-Narrator.settings.json"

# ---------- Words to say differently (edit freely) ----------
$Pronounce = [ordered]@{
    "e.g." = "for example"; "i.e." = "that is"
    "ISO"  = "I S O";  "PPE" = "P P E";  "OHS" = "O H S";  "WHS" = "W H S"
    "HSRs" = "H S Rs"; "HSR" = "H S R";  "HSC" = "H S C";  "PCBU" = "P C B U"
    "AT1"  = "A T 1";  "AT2" = "A T 2";  "MR"  = "M R";    "FER" = "F E R"; "SOC" = "S O C"
}

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

function Fix-Pronunciation([string]$text) {
    foreach ($k in $Pronounce.Keys) {
        $pattern = '(?<![A-Za-z])' + [regex]::Escape($k) + '(?![A-Za-z])'
        $text = [regex]::Replace($text, $pattern, $Pronounce[$k])
    }
    return $text
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
# Slide Narrator video worker: builds one MP4 from slide images + audio with ffmpeg
$ErrorActionPreference = "Continue"   # native tools write to stderr; we check exit codes instead
$job = Get-Content -LiteralPath $JobFile -Raw | ConvertFrom-Json
$log = $job.Log
function W([string]$m) { try { Add-Content -LiteralPath $log -Value ((Get-Date -Format "HH:mm:ss") + "  " + $m) -Encoding UTF8 } catch {} }
try {
    $ff = $job.Ffmpeg; $fps = [int]$job.Fps
    if ($job.Codec -eq "h265") { $vc = @("-c:v", "libx265", "-preset", "medium", "-crf", "26", "-x265-params", "log-level=error") }
    else                       { $vc = @("-c:v", "libx264", "-preset", "medium", "-crf", "22", "-tune", "stillimage") }
    if ($fps -le 0) {
        # ---- One frame per slide (variable frame rate) ----
        $total = @($job.Segments).Count; $n = 0; $clock = 0.0; $chapters = @()
        $imgLines = @(); $audLines = @(); $pads = @()
        foreach ($s in $job.Segments) {
            $n++
            $secs = [Math]::Round([double]$s.Dur, 3)
            if ($s.Chapter) { $chapters += [pscustomobject]@{ Title = [string]$s.Chapter; Start = $clock } }
            $clock += $secs
            $dur = $secs.ToString("0.###", [Globalization.CultureInfo]::InvariantCulture)
            $pad = Join-Path $job.WorkDir ("pad_{0:D3}.wav" -f $n)
            if ($s.Wav) { $a = @("-y", "-loglevel", "error", "-i", $s.Wav, "-af", "apad", "-t", $dur, "-ar", "48000", "-ac", "1", "-c:a", "pcm_s16le", $pad) }
            else        { $a = @("-y", "-loglevel", "error", "-f", "lavfi", "-i", "anullsrc=r=48000:cl=mono", "-t", $dur, "-c:a", "pcm_s16le", $pad) }
            $out = & $ff @a 2>&1
            if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $pad)) { throw "slide $($s.Index) audio: " + (($out | Out-String).Trim()) }
            $pads += $pad
            $q = { param($x) "file '" + ($x -replace "'", "'\''") + "'" }
            $imgLines += (& $q $s.Img); $imgLines += "duration $dur"
            $audLines += (& $q $pad)
            Set-Content -LiteralPath $job.Progress -Value "$n/$total"
        }
        $imgLines += (& $q (@($job.Segments)[-1].Img))      # concat quirk: repeat the last picture
        $utf8 = New-Object System.Text.UTF8Encoding($false)
        $imgList = Join-Path $job.WorkDir "images.txt"; [IO.File]::WriteAllLines($imgList, $imgLines, $utf8)
        $audList = Join-Path $job.WorkDir "audio.txt";  [IO.File]::WriteAllLines($audList, $audLines, $utf8)
        $a = @("-y", "-loglevel", "error", "-f", "concat", "-safe", "0", "-i", $imgList, "-f", "concat", "-safe", "0", "-i", $audList)
        $map = @("-map", "0:v", "-map", "1:a")
        if ($chapters.Count -gt 0) {
            $meta = Join-Path $job.WorkDir "chapters.txt"
            $ml = @(";FFMETADATA1")
            for ($c = 0; $c -lt $chapters.Count; $c++) {
                $st = [long][Math]::Round($chapters[$c].Start * 1000)
                $en = if ($c + 1 -lt $chapters.Count) { [long][Math]::Round($chapters[$c + 1].Start * 1000) } else { [long][Math]::Round($clock * 1000) }
                $tt = ($chapters[$c].Title -replace '[\r\n]+', ' ') -replace '([\\=;#])', '\$1'
                $ml += @("[CHAPTER]", "TIMEBASE=1/1000", "START=$st", "END=$en", "title=$tt")
            }
            [IO.File]::WriteAllLines($meta, $ml, $utf8)
            $a += @("-i", $meta); $map += @("-map_metadata", "2", "-map_chapters", "2")
        }
        $vcv = $vc + @("-g", "1", "-pix_fmt", "yuv420p")
        $tail = @("-c:a", "aac", "-b:a", "128k", "-t", $clock.ToString("0.###", [Globalization.CultureInfo]::InvariantCulture), "-movflags", "+faststart")
        if ($job.Codec -eq "h265") { $tail += @("-tag:v", "hvc1") }
        $out = & $ff @($a + $map + $vcv + @("-fps_mode", "vfr") + $tail + @($job.Out)) 2>&1
        if ($LASTEXITCODE -ne 0) {   # older ffmpeg: -vsync instead of -fps_mode
            $out = & $ff @($a + $map + $vcv + @("-vsync", "vfr") + $tail + @($job.Out)) 2>&1
        }
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $job.Out)) { throw "building video: " + (($out | Out-String).Trim()) }
        foreach ($pd in $pads) { Remove-Item -LiteralPath $pd -Force -ErrorAction SilentlyContinue }
        Set-Content -LiteralPath $job.Progress -Value "done"
        exit 0
    }

    $segs = @(); $n = 0; $total = @($job.Segments).Count
    $clock = 0.0; $chapters = @()
    foreach ($s in $job.Segments) {
        $n++
        $seg = Join-Path $job.WorkDir ("seg_{0:D3}.mkv" -f $n)
        # Round each slide to whole video frames so chapter times line up exactly
        $frames = [Math]::Max(1, [Math]::Ceiling([double]$s.Dur * $fps - 0.000001))
        $secs = $frames / $fps
        if ($s.Chapter) { $chapters += [pscustomobject]@{ Title = [string]$s.Chapter; Start = $clock } }
        $clock += $secs
        $dur = ([double]$secs).ToString("0.######", [Globalization.CultureInfo]::InvariantCulture)
        $a = @("-y", "-loglevel", "error", "-loop", "1", "-framerate", "$fps", "-i", $s.Img)
        if ($s.Wav) { $a += @("-i", $s.Wav, "-af", "apad") } else { $a += @("-f", "lavfi", "-i", "anullsrc=r=48000:cl=mono") }
        $a += @("-t", $dur) + $vc + @("-pix_fmt", "yuv420p", "-r", "$fps", "-c:a", "pcm_s16le", "-ar", "48000", "-ac", "1", $seg)
        $out = & $ff @a 2>&1
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $seg)) { throw "slide $($s.Index): " + (($out | Out-String).Trim()) }
        $segs += $seg
        Set-Content -LiteralPath $job.Progress -Value "$n/$total"
    }
    $list = Join-Path $job.WorkDir "segments.txt"
    $lines = $segs | ForEach-Object { "file '" + ($_ -replace "'", "'\''") + "'" }
    [IO.File]::WriteAllLines($list, $lines, (New-Object System.Text.UTF8Encoding($false)))
    $a = @("-y", "-loglevel", "error", "-f", "concat", "-safe", "0", "-i", $list)
    if ($chapters.Count -gt 0) {
        # Chapter markers (FFMETADATA format)
        $meta = Join-Path $job.WorkDir "chapters.txt"
        $ml = @(";FFMETADATA1")
        for ($c = 0; $c -lt $chapters.Count; $c++) {
            $st = [long][Math]::Round($chapters[$c].Start * 1000)
            $en = if ($c + 1 -lt $chapters.Count) { [long][Math]::Round($chapters[$c + 1].Start * 1000) } else { [long][Math]::Round($clock * 1000) }
            $tt = ($chapters[$c].Title -replace '[\r\n]+', ' ') -replace '([\\=;#])', '\$1'
            $ml += @("[CHAPTER]", "TIMEBASE=1/1000", "START=$st", "END=$en", "title=$tt")
        }
        [IO.File]::WriteAllLines($meta, $ml, (New-Object System.Text.UTF8Encoding($false)))
        $a += @("-i", $meta, "-map", "0:v", "-map", "0:a", "-map_metadata", "1", "-map_chapters", "1")
    }
    $a += @("-c:v", "copy", "-c:a", "aac", "-b:a", "128k", "-movflags", "+faststart")
    if ($job.Codec -eq "h265") { $a += @("-tag:v", "hvc1") }
    $a += $job.Out
    $out = & $ff @a 2>&1
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $job.Out)) { throw "joining: " + (($out | Out-String).Trim()) }
    foreach ($sg in $segs) { Remove-Item -LiteralPath $sg -Force -ErrorAction SilentlyContinue }
    Set-Content -LiteralPath $job.Progress -Value "done"
    exit 0
} catch {
    W ("VIDEO FAILED: " + $_.Exception.Message)
    Set-Content -LiteralPath $job.Progress -Value ("failed: " + $_.Exception.Message)
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
$form.Size = New-Object System.Drawing.Size(640, 744)
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
$grp.Size = New-Object System.Drawing.Size(585, 159)
$form.Controls.Add($grp)

$rbDeck = New-Object System.Windows.Forms.RadioButton
$rbDeck.Text = "Narrated PowerPoint only"
$rbDeck.Location = New-Object System.Drawing.Point(15, 25)
$rbDeck.AutoSize = $true
$rbDeck.Checked = $true
$grp.Controls.Add($rbDeck)

$rbVideo = New-Object System.Windows.Forms.RadioButton
$rbVideo.Text = "Narrated PowerPoint + MP4 video"
$rbVideo.Location = New-Object System.Drawing.Point(15, 55)
$rbVideo.AutoSize = $true
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

$lblPar = New-Object System.Windows.Forms.Label
$lblPar.Text = "Slides at once:"
$lblPar.Location = New-Object System.Drawing.Point(330, 27)
$lblPar.AutoSize = $true
$grp.Controls.Add($lblPar)

$parallelBox = New-Object System.Windows.Forms.NumericUpDown
$parallelBox.Location = New-Object System.Drawing.Point(445, 23)
$parallelBox.Size = New-Object System.Drawing.Size(55, 26)
$parallelBox.Minimum = 1; $parallelBox.Maximum = 8; $parallelBox.Value = 4
$grp.Controls.Add($parallelBox)

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
$videoParBox.Minimum = 1; $videoParBox.Maximum = 6; $videoParBox.Value = 2
$grp.Controls.Add($videoParBox)

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

function Update-VideoChoices($wantRes, $wantFps) {
    $ff = $methodBox.SelectedIndex -lt 2
    Fill-Choices $resBox $script:ResChoices $ff $wantRes "1080p"
    Fill-Choices $fpsBox $script:FpsChoices $ff $wantFps $(if ($ff) { "1 per slide (VFR)" } else { "15 fps" })
}
Update-VideoChoices $null $null
$methodBox.Add_SelectedIndexChanged({ Update-VideoChoices $null $null })

# Start / Cancel
$start = New-Object System.Windows.Forms.Button
$start.Text = "Start"
$start.Location = New-Object System.Drawing.Point(20, 416)
$start.Size = New-Object System.Drawing.Size(470, 40)
$start.BackColor = $navy
$start.ForeColor = [System.Drawing.Color]::White
$start.FlatStyle = "Flat"
$start.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($start)

$cancel = New-Object System.Windows.Forms.Button
$cancel.Text = "Cancel"
$cancel.Location = New-Object System.Drawing.Point(500, 416)
$cancel.Size = New-Object System.Drawing.Size(105, 40)
$cancel.Enabled = $false
$form.Controls.Add($cancel)

# Progress + log
$lblAudio = New-Object System.Windows.Forms.Label
$lblAudio.Text = "Audio: -"
$lblAudio.Location = New-Object System.Drawing.Point(20, 461)
$lblAudio.Size = New-Object System.Drawing.Size(280, 20)
$lblAudio.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Controls.Add($lblAudio)

$lblVideo = New-Object System.Windows.Forms.Label
$lblVideo.Text = "Video: -"
$lblVideo.Location = New-Object System.Drawing.Point(300, 461)
$lblVideo.Size = New-Object System.Drawing.Size(305, 20)
$lblVideo.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Controls.Add($lblVideo)

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(20, 484)
$progress.Size = New-Object System.Drawing.Size(585, 20)
$form.Controls.Add($progress)

$log = New-Object System.Windows.Forms.TextBox
$log.Location = New-Object System.Drawing.Point(20, 512)
$log.Size = New-Object System.Drawing.Size(585, 150)
$log.Multiline = $true
$log.ScrollBars = "Vertical"
$log.ReadOnly = $true
$log.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($log)

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
$script:Cancel    = $false

if ($script:BalconExe) { Log "balcon: $($script:BalconExe) (backup engine)" }
else { Log "balcon.exe not found. That's OK: Windows speech will be used directly." }
$voices = Get-Voices $script:BalconExe
foreach ($v in $voices) { [void]$voiceBox.Items.Add($v) }
if ($voiceBox.Items.Count -eq 0) { Log "No voices found automatically. You can type the voice name into the Voice box." }
else { Log "Voices found: $($voiceBox.Items.Count)"; foreach ($v in $voices) { $script:StartupLines.Add("   voice: $v") } }
if ($script:Ffmpeg) { Log "ffmpeg: found (audio saved as AAC .m4a)" } else { Log "ffmpeg: not found (audio saved as WAV, which works fine)" }

# Restore last settings
$saved = Load-Settings
if ($saved) {
    if ($saved.Voice) { if ($voiceBox.Items.Contains($saved.Voice)) { $voiceBox.SelectedItem = $saved.Voice } else { $voiceBox.Text = $saved.Voice } }
    if ($null -ne $saved.Speed) { $speed.Value = [Math]::Max(-5, [Math]::Min(5, [int]$saved.Speed)) }
    if ($saved.Video) { $rbVideo.Checked = $true }

    if ($saved.Parallel) { $parallelBox.Value = [Math]::Max(1, [Math]::Min(8, [int]$saved.Parallel)) }
    if ($null -ne $saved.Method -and [int]$saved.Method -lt $methodBox.Items.Count) { $methodBox.SelectedIndex = [int]$saved.Method }
    if ($saved.VideoParallel) { $videoParBox.Value = [Math]::Max(1, [Math]::Min(6, [int]$saved.VideoParallel)) }
    if ($null -ne $saved.Chapters -and [int]$saved.Chapters -lt $chapterBox.Items.Count) { $chapterBox.SelectedIndex = [int]$saved.Chapters }
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
$rbVideo.Add_CheckedChanged({ $resBox.Enabled = $rbVideo.Checked; $fpsBox.Enabled = $rbVideo.Checked; $methodBox.Enabled = $rbVideo.Checked; $videoParBox.Enabled = $rbVideo.Checked; $chapterBox.Enabled = $rbVideo.Checked })
$resBox.Enabled = $rbVideo.Checked; $fpsBox.Enabled = $rbVideo.Checked; $methodBox.Enabled = $rbVideo.Checked; $videoParBox.Enabled = $rbVideo.Checked; $chapterBox.Enabled = $rbVideo.Checked

$preview.Add_Click({
    if (-not $voiceBox.Text.Trim()) { return }
    $preview.Enabled = $false
    Log "Previewing $($voiceBox.Text.Trim())..."
    $pv = Join-Path $env:TEMP "slide-narrator-preview.wav"
    if (Invoke-Speech "Welcome everyone. This unit is about W H S communication and consultation." $pv $voiceBox.Text.Trim() $speed.Value) {
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

$cancel.Add_Click({ $script:Cancel = $true; Log "Cancelling after the current slide..." })

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
            $speech = ""
            if ($notes) { $speech = Fix-Pronunciation (Clean-Text $notes) }
            $list += [pscustomobject]@{ Index = $slide.SlideIndex; Speech = $speech }
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
    $job.Proc = Start-Process -FilePath $exe -ArgumentList $argLine -WindowStyle Hidden -PassThru
}

# Check running jobs, start new ones, update progress. Call this often.
function Pump-Video {
    $r = $script:Run
    $running = 0; $queued = 0; $done = 0; $status = @()
    foreach ($v in $r.Videos) {
        if ($v.Done) { $done++; continue }
        if (-not $v.Proc) { $queued++; continue }
        if ($v.Proc.HasExited) {
            $v.Done = $true; $done++
            $p = ""; try { $p = (Get-Content -LiteralPath $v.Progress -Raw).Trim() } catch {}
            $v.Ok = ($p -eq "done") -and (Test-Path -LiteralPath $v.Out)
            if ($v.Ok) { $v.Deck.VideoNote = "video OK"; Log-To $v.Deck.Log "Saved video: $($v.Out)" }
            else { $v.Deck.VideoNote = "video FAILED"; Log-To $v.Deck.Log "Video failed for $($v.Deck.Name): $p" }
        } else {
            $running++
            $p = ""; try { $p = (Get-Content -LiteralPath $v.Progress -Raw -ErrorAction Stop).Trim() } catch {}
            if ($p) { $status += "$($v.Deck.Name) $p" }
        }
    }
    foreach ($v in $r.Videos) {
        if ($running -ge $r.VideoParallel) { break }
        if (-not $v.Done -and -not $v.Proc) {
            $exe = Join-Path $PSHOME "powershell.exe"
            $v.Proc = Start-Process -FilePath $exe -ArgumentList ("-NoProfile -ExecutionPolicy Bypass -File " + (Quote $script:VideoWorker) + " " + (Quote $v.JobFile)) -WindowStyle Hidden -PassThru
            Log-To $v.Deck.Log "Building video with ffmpeg: $($v.Deck.Name)"
            $running++; $queued--
        }
    }
    if ($r.Videos.Count -gt 0) {
        $txt = "Video: $running building, $queued waiting, $done done"
        if ($status.Count -gt 0) { $txt += "  (" + ($status[0] -replace '^(.{0,25}).*?(\d+/\d+)$', '$1... $2') + ")" }
        $lblVideo.Text = $txt
    }
}

function Pump-Audio {
    $r = $script:Run
    if ($script:Cancel) {
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
                # AAC conversion finished (if it failed, the WAV is used instead)
                $j.Done = $true; $j.Ok = $true; $r.DoneCount++; continue
            }
            $ok = (Test-Path $j.Wav) -and ((Get-Item $j.Wav).Length -gt 1000)
            if ($ok -and $script:Ffmpeg) {
                # Speech done: compress it to AAC (.m4a) in the background for the PowerPoint
                $j.Phase = "convert"
                $argLine = "-y -loglevel error -i " + (Quote $j.Wav) + " -c:a aac -b:a 64k -ac 1 " + (Quote $j.M4a)
                $j.Proc = Start-Process -FilePath $script:Ffmpeg -ArgumentList $argLine -WindowStyle Hidden -PassThru
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

# Add chapters to an existing MP4 (used after a PowerPoint export)
function Add-ChaptersToMp4([string]$mp4, $segments, [string]$workDir) {
    $chapters = @(); $clock = 0.0
    foreach ($sg in $segments) {
        if ($sg.Chapter) { $chapters += [pscustomobject]@{ Title = [string]$sg.Chapter; Start = $clock } }
        $clock += [double]$sg.Dur
    }
    if ($chapters.Count -eq 0) { return }
    $ml = @(";FFMETADATA1")
    for ($c = 0; $c -lt $chapters.Count; $c++) {
        $st = [long][Math]::Round($chapters[$c].Start * 1000)
        $en = if ($c + 1 -lt $chapters.Count) { [long][Math]::Round($chapters[$c + 1].Start * 1000) } else { [long][Math]::Round($clock * 1000) }
        $tt = $chapters[$c].Title -replace '([\\=;#])', '\$1'
        $ml += @("[CHAPTER]", "TIMEBASE=1/1000", "START=$st", "END=$en", "title=$tt")
    }
    $meta = Join-Path $workDir "chapters.txt"
    [IO.File]::WriteAllLines($meta, $ml, (New-Object System.Text.UTF8Encoding($false)))
    $tmp = [IO.Path]::ChangeExtension($mp4, ".chapters.mp4")
    & $script:Ffmpeg -y -loglevel error -i $mp4 -i $meta -map 0 -map_metadata 1 -map_chapters 1 -c copy -movflags +faststart $tmp 2>$null | Out-Null
    if ((Test-Path $tmp) -and (Get-Item $tmp).Length -gt 1000) {
        Remove-Item $mp4 -Force; Rename-Item $tmp ([IO.Path]::GetFileName($mp4))
        Log "Added $($chapters.Count) chapters to the video."
    } else { Log "Could not add chapters to the video (the video itself is fine)." }
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
    try {
        foreach ($slide in $pres.Slides) {
            Pump-Audio
            $i = $slide.SlideIndex
            for ($s = $slide.Shapes.Count; $s -ge 1; $s--) {
                if ($slide.Shapes.Item($s).Name -eq "Narration") { $slide.Shapes.Item($s).Delete() }
            }
            $t = $slide.SlideShowTransition
            $t.AdvanceOnClick = $msoFalse; $t.AdvanceOnTime = $msoTrue
            $job = $r.Jobs | Where-Object { $_.Deck -eq $deck -and $_.Index -eq $i } | Select-Object -First 1
            $img = Join-Path $deck.WorkDir ("slide_{0:D2}.png" -f $i)
            if (-not $job) { $t.AdvanceTime = $r.NoNotes; $segments += [pscustomobject]@{ Index = $i; Img = $img; Wav = ""; Dur = [double]$r.NoNotes; Chapter = [string]$chap[[int]$i] }; continue }

            if (-not $job.Ok) {
                Log "   slide ${i}: retrying audio..."
                $job.Ok = New-SlideAudio $job.Speech $job.Wav $r.Voice $r.Rate
            }
            if (-not $job.Ok) {
                $secs = [Math]::Max(5, [int](($job.Speech -split '\s+').Count / 2.5))
                $t.AdvanceTime = $secs; $failed += $i
                $segments += [pscustomobject]@{ Index = $i; Img = $img; Wav = ""; Dur = [double]$secs; Chapter = [string]$chap[[int]$i] }
                Log ("   FAILED. Slide {0} has no audio and will show for {1} s. Reason: {2}" -f $i, $secs, $script:LastSpeechError)
                continue
            }
            # Use the AAC file made in the background; make it now only if it's missing
            $audio = $job.Wav
            if ($script:Ffmpeg) {
                if (-not ((Test-Path $job.M4a) -and (Get-Item $job.M4a).Length -gt 500)) {
                    & $script:Ffmpeg -y -loglevel error -i $job.Wav -c:a aac -b:a 64k -ac 1 $job.M4a 2>$null | Out-Null
                }
                if ((Test-Path $job.M4a) -and (Get-Item $job.M4a).Length -gt 500) { $audio = $job.M4a }
            }
            $media = $slide.Shapes.AddMediaObject2($audio, $msoFalse, $msoTrue, 10, 10, 40, 40)
            $media.Name = "Narration"
            $play = $media.AnimationSettings.PlaySettings
            $play.PlayOnEntry = $msoTrue; $play.HideWhileNotPlaying = $msoTrue
            $t.AdvanceTime = $r.Pause
            $segments += [pscustomobject]@{ Index = $i; Img = $img; Wav = $job.Wav; Dur = [Math]::Round((Get-WavSeconds $job.Wav) + $r.Pause, 3); Chapter = [string]$chap[[int]$i] }
        }
        $pres.SaveAs($deck.OutPptx)
        Log "Saved: $($deck.OutPptx)"

        if ($r.Video -and $r.Fast) {
            # Export each slide as a picture, then hand the rest to ffmpeg in the background
            $h = [int]$r.VRes
            $w = [int]([Math]::Round($h * $pres.PageSetup.SlideWidth / $pres.PageSetup.SlideHeight / 2) * 2)
            Log "Exporting slide pictures ($w x $h)..."
            foreach ($slide in $pres.Slides) {
                Pump-Audio
                $slide.Export((Join-Path $deck.WorkDir ("slide_{0:D2}.png" -f $slide.SlideIndex)), "PNG", $w, $h)
            }
            $vjob = [pscustomobject]@{
                Ffmpeg = $script:Ffmpeg; Fps = $r.Fps; Codec = $r.Codec; WorkDir = $deck.WorkDir; Log = $deck.Log
                Progress = (Join-Path $deck.WorkDir "video-progress.txt"); Out = $deck.OutMp4; Segments = $segments
            }
            $jf = Join-Path $deck.WorkDir "video-job.json"
            [IO.File]::WriteAllText($jf, ($vjob | ConvertTo-Json -Depth 5), (New-Object System.Text.UTF8Encoding($false)))
            if (Test-Path $vjob.Progress) { Remove-Item $vjob.Progress -Force }
            if (Test-Path $deck.OutMp4) { Remove-Item $deck.OutMp4 -Force -ErrorAction SilentlyContinue }
            $script:Run.Videos += [pscustomobject]@{ Deck = $deck; JobFile = $jf; Progress = $vjob.Progress; Out = $deck.OutMp4; Proc = $null; Done = $false; Ok = $false }
            Log "Video queued ($($r.CodecName)). It will be built in the background."
            $deck.VideoNote = "video queued"
        }
        elseif ($r.Video) {
            Log "Exporting video ($($r.VRes)p, $($r.Fps) fps)... audio for the next decks keeps going meanwhile."
            $pres.CreateVideo($deck.OutMp4, $true, $r.NoNotes, $r.VRes, $r.Fps, 85)
            $lblVideo.Text = "Video: exporting $($deck.Name)"
            do {
                for ($k = 0; $k -lt 8; $k++) { Pump-Audio; Start-Sleep -Milliseconds 250 }
                $status = $pres.CreateVideoStatus      # 1 in progress, 2 queued, 3 done, 4 failed
            } while ($status -eq 1 -or $status -eq 2)
            $lblVideo.Text = "Video: -"
            if ($status -eq 3) { Log "Saved video: $($deck.OutMp4)"; if ($script:Ffmpeg -and $r.Chapters -ne 2) { Add-ChaptersToMp4 $deck.OutMp4 $segments $deck.WorkDir } }
            else { $videoOk = $false; Log "Video export failed (status $status). Try 1080p, or open the narrated deck and use File > Export > Create a Video." }
        }
    } finally {
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

    $settings = [pscustomobject]@{ Voice = "$($voiceBox.Text.Trim())"; Speed = $speed.Value; Video = $rbVideo.Checked; Resolution = "$($resBox.SelectedItem)"; FrameRate = "$($fpsBox.SelectedItem)"; Parallel = [int]$parallelBox.Value; Method = $methodBox.SelectedIndex; VideoParallel = [int]$videoParBox.Value; Chapters = $chapterBox.SelectedIndex }
    $fast = $settings.Video -and ($methodBox.SelectedIndex -lt 2)
    if ($fast -and -not $script:Ffmpeg) {
        [System.Windows.Forms.MessageBox]::Show("The fast video method needs ffmpeg.exe, which wasn't found. PowerPoint will export the videos instead.", "Slide Narrator") | Out-Null
        $fast = $false
        if ($fps -eq 0) { $fps = 15 }
        if ($vres -eq 1440) { $vres = 1080 }
    }
    Save-Settings $settings
    $vres = Choice-Value $script:ResChoices "$($resBox.SelectedItem)"; if (-not $vres) { $vres = 1080 }
    $fps  = Choice-Value $script:FpsChoices "$($fpsBox.SelectedItem)"; if ($null -eq $fps) { $fps = 15 }

    $script:Cancel = $false
    $controls = @($start, $browse, $removeBtn, $clearBtn, $fileList, $voiceBox, $preview, $testAll, $speed, $grp)
    foreach ($c in $controls) { $c.Enabled = $false }
    $cancel.Enabled = $true
    $form.AllowDrop = $false; $drop.AllowDrop = $false
    $log.Clear()

    $script:Run = [pscustomobject]@{
        Voice = $settings.Voice; Rate = $settings.Speed; Video = $settings.Video; VRes = $vres; Fps = $fps
        Parallel = $settings.Parallel; Pause = 1; NoNotes = 4; Jobs = @(); DoneCount = 0
        Fast = $fast; Codec = $(if ($methodBox.SelectedIndex -eq 1) { "h265" } else { "h264" }); CodecName = $methodBox.Text
        Videos = @(); VideoParallel = $settings.VideoParallel; Chapters = $settings.Chapters
    }
    $decks = @()
    foreach ($p in @($fileList.Items)) {
        $path = "$p"; $dir = Split-Path $path -Parent; $name = [IO.Path]::GetFileNameWithoutExtension($path)
        $decks += [pscustomobject]@{
            Path = $path; Name = $name; Log = (Join-Path $dir ($name + "_narration.log"))
            WorkDir = (Join-Path $dir ($name + "_narration"))
            OutPptx = (Join-Path $dir ($name + "_narrated.pptx")); OutMp4 = (Join-Path $dir ($name + ".mp4"))
            VideoNote = ""
        }
    }
    $results = @()
    $ppt = $null
    try {
        # Check the voice works and pick the fastest engine for it
        $log.AppendText("Checking the voice..." + [Environment]::NewLine)
        $pv = Join-Path $env:TEMP "slide-narrator-check.wav"
        if (-not (Invoke-Speech "Checking." $pv $script:Run.Voice $script:Run.Rate)) {
            throw "The voice '$($script:Run.Voice)' isn't working. Try 'Test voices'. Reason: $($script:LastSpeechError)"
        }
        $log.AppendText("Voice OK (engine: $($script:Engine)). Making audio $($script:Run.Parallel) slides at a time." + [Environment]::NewLine)

        $ppt = New-Object -ComObject PowerPoint.Application
        $ppt.Visible = -1

        # 1. Read notes from every deck and queue the audio jobs
        foreach ($d in $decks) {
            New-DeckLog $d.Log ([pscustomobject]@{ File = $d.Path; Voice = $settings.Voice; Speed = $settings.Speed; Video = $settings.Video; Resolution = $settings.Resolution; FrameRate = $settings.FrameRate; Method = $methodBox.Text; Parallel = $settings.Parallel })
            Log-To $d.Log "Reading notes: $($d.Name)"
            New-Item -ItemType Directory -Force -Path $d.WorkDir | Out-Null
            try {
                $notes = Get-DeckNotes $ppt $d.Path
                $d | Add-Member -NotePropertyName Ready -NotePropertyValue $true
                foreach ($n in $notes) {
                    if (-not $n.Speech) { continue }
                    $tag = "slide_{0:D2}" -f $n.Index
                    $script:Run.Jobs += [pscustomobject]@{
                        Deck = $d; Index = $n.Index; Speech = $n.Speech
                        Txt = (Join-Path $d.WorkDir "$tag.txt"); Wav = (Join-Path $d.WorkDir "$tag.wav"); M4a = (Join-Path $d.WorkDir "$tag.m4a"); Phase = "speak"
                        Proc = $null; Tries = 0; Done = $false; Ok = $false
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
                $results += $d
            } catch {
                $msg = $_.Exception.Message
                if ($msg -eq "Cancelled.") { throw }
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
        if ($msg -eq "Cancelled.") { $results += "CANCELLED (remaining decks were not finished)" }
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
        $resBox.Enabled = $rbVideo.Checked; $fpsBox.Enabled = $rbVideo.Checked; $methodBox.Enabled = $rbVideo.Checked; $videoParBox.Enabled = $rbVideo.Checked; $chapterBox.Enabled = $rbVideo.Checked
        $cancel.Enabled = $false
        $form.AllowDrop = $true; $drop.AllowDrop = $true
        $form.Text = "Slide Narrator"
        $lblVideo.Text = "Video: -"
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
