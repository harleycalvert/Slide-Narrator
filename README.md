# Slide Narrator

Turn PowerPoint speaker notes into natural-voice narration and export as video. Windows, drag and drop, batch-friendly.

Slide Narrator reads the speaker notes on every slide, voices them with a natural-sounding text-to-speech voice, inserts the audio into each slide (set to play and advance automatically), and can export the result as an MP4. Your original PowerPoint is never changed.

## Features

- **Drag and drop:** add one deck, many decks, or a whole folder.
- **Natural voices:** Microsoft Natural voices, including offline Australian voices (Annette, Natasha, Poppy, William) and online Edge voices.
- **Fast:** audio is made several slides at a time, in the background.
- **Fast video with ffmpeg:** H.264 or H.265, from 720p to 4K, with fixed frame rates or **1 frame per slide (VFR)**. Several videos can be built at once.
- **Chapters:** taken from PowerPoint sections, or from slide titles.
- **Reliable:** cleans hidden characters from notes, retries failed slides, falls back to sentence-by-sentence, and tries three speech engines.
- **Logs:** one log per deck, saved next to it.

## Requirements

- Windows 10 or 11
- PowerPoint (Microsoft 365)
- Windows PowerShell 5.1 (built into Windows)

## Downloads

| File | Get it from | Needed? |
|---|---|---|
| `NaturalVoiceSAPIAdapter_v0.2.9_x86_x64.zip` | [github.com/gexgd0419/NaturalVoiceSAPIAdapter/releases](https://github.com/gexgd0419/NaturalVoiceSAPIAdapter/releases) | Required |
| `ms_natural_voice_en_au.zip` | [cross-plus-a.com/voice.htm](https://www.cross-plus-a.com/voice.htm) | Recommended (offline Australian voices) |
| `balcon.zip` | [cross-plus-a.com/bconsole.htm](https://www.cross-plus-a.com/bconsole.htm) | Recommended (backup speech engine) |
| `ffmpeg.exe` | [ffmpeg.org/download.html](https://ffmpeg.org/download.html) | Recommended (smaller audio, fast video, chapters) |
| `Slide-Narrator.ps1` and `Slide-Narrator.bat` | This repository | Required |

## Setup

### 1. Install the voices

1. Extract the adapter and the Australian voice pack into **permanent** folders. Moving or deleting them later breaks the voices.
2. Right-click the adapter's `Installer.exe` and choose **Run as administrator**.
3. Tick **Enable Microsoft Edge online voices**.
4. Set **Local voice path** to the Australian voices folder.
5. Click **Install 32-bit** and **Install 64-bit**. Both must show **Installed**.

### 2. Build the app folder

```
Slide-Narrator\
    Slide-Narrator.bat
    Slide-Narrator.ps1
    ffmpeg.exe
    balcon\
        balcon.exe
        ...
```

### 3. Open and test

1. Double-click `Slide-Narrator.bat`. A console window stays open behind the app. That's normal, and it shows any startup errors.
2. Click **Test voices**. Every voice is tried with every engine.
3. Pick a voice that shows `OK` and click **Preview**.

### 4. Narrate

1. Drag in your decks, or a folder of them.
2. Choose a voice and speed.
3. Choose your output and video settings.
4. Click **Start**.

**Recommended video settings:** `Fast - ffmpeg H.264` · `4K / 2160p` · `1 per slide (VFR)`

## Output

Everything is saved next to each original deck:

| File | Contents |
|---|---|
| `<deck>_narrated.pptx` | The deck with narration embedded and slides that advance automatically |
| `<deck>.mp4` | The video, if you chose one |
| `<deck>_narration.log` | What happened on every slide |
| `<deck>_narration\` | Working files: each slide's text, audio and picture. Safe to delete. |

## Options

| Setting | Choices | Notes |
|---|---|---|
| Slides at once | 1–8 | How many slides are voiced in parallel. Use 6–8 for Native voices and about 4 for Online voices. |
| Video method | Fast - ffmpeg H.264 · Fast - ffmpeg H.265 · PowerPoint | H.264 plays everywhere. H.265 files are smaller, but some browsers can't play them. PowerPoint is the only method that keeps animations, but it's slow and does one video at a time. |
| Resolution | 720p · 1080p · 1440p (ffmpeg only) · 4K | |
| Frame rate | 1 per slide (VFR, ffmpeg only) · 5–60 fps | VFR gives the smallest files and is quickest to build. Every frame is a keyframe. |
| Videos at once | 1–6 | Parallel ffmpeg video builds |
| Chapters | By section (or by slide) · By slide · None | |

### Pronunciation

Acronyms are spelled out for the voice, for example `WHS` becomes "W H S". The list is in the `$Pronounce` table near the top of `Slide-Narrator.ps1`, and you can edit it.

## Troubleshooting

| Problem | Fix |
|---|---|
| A voice fails with `Class not registered` | Run the adapter's `Installer.exe` as administrator and reinstall both the 32-bit and 64-bit versions. |
| The app closes as soon as it opens | Read the error in the console window behind it. |
| Online voices fail part-way through | Lower **Slides at once** to 2 or 3. |
| An H.265 video won't play in a browser | Rebuild it with **Fast - ffmpeg H.264**. |
| A VFR video stutters in a player | Change the frame rate to 5 or 10 fps. |
| You need animations in the video | Use the **PowerPoint** video method. |
| Anything else | Check `<deck>_narration.log`, or `Slide-Narrator voice test.log` for voice tests. |

## How it works

1. **Read:** speaker notes are read from each deck through PowerPoint automation, then cleaned and prepared for pronunciation.
2. **Voice:** background workers turn each slide's notes into a WAV file. The app uses SAPI 5 directly in 64-bit, a 32-bit SAPI helper, or `balcon.exe`, whichever works for the chosen voice. ffmpeg then converts each WAV to AAC.
3. **Build:** the audio is embedded in each slide, set to play automatically, and the slide advances one second after the audio ends. The deck is saved as `_narrated.pptx`.
4. **Video:** slides are exported as PNG pictures, and ffmpeg combines them with the WAV audio, adding chapters if you chose them. The PowerPoint method uses PowerPoint's own **Create a Video** export instead.

Audio for the next deck keeps being made while earlier videos are building.

## Version history

| Version | Changes |
|---|---|
| 1 | Script: one deck, command line, balcon voices |
| 2 | App: drag and drop, voice list, Preview, speed, video up to 4K |
| 3 | Reliability: text clean-up, retries, sentence fallback, log file |
| 4 | Voice engines: SAPI built in, 32-bit engine, Test voices, full voice names |
| 5 | Batch: many decks, one log per deck, parallel audio while videos export |
| 6 | Fast video: ffmpeg H.264 or H.265, several videos at once |
| 7 | Video options: chapters, AAC audio, resolution and frame-rate menus, 1 frame per slide (VFR) |

## Credits

- [NaturalVoiceSAPIAdapter](https://github.com/gexgd0419/NaturalVoiceSAPIAdapter) by gexgd0419
- [Balabolka and balcon](https://www.cross-plus-a.com/) by Ilya Morozov
- [FFmpeg](https://ffmpeg.org/)

These tools are not included in this repository. Download them from their sources above, and check each one's own licence.
