# Slide Narrator

Turn PowerPoint speaker notes into natural-voice narration and export as video. Windows, drag and drop, batch-friendly.

Slide Narrator reads the speaker notes on every slide, voices them with a natural-sounding text-to-speech voice, inserts the audio into each slide (set to play and advance automatically), and can export the result as an MP4. Your original PowerPoint is never changed.

It's built for speed: one slide is voiced per CPU thread, and several videos can be built in parallel. The videos are up to 4K resolution with tiny file sizes. Each slide is stored as a single frame for as long as its narration runs.

## Features

- **Drag and drop:** add one deck, many decks, or a whole folder.
- **Natural voices:** Microsoft Natural voices, including offline Australian voices (Annette, Natasha, Poppy, William) and online Edge voices.
- **Fast:** audio is made several slides at a time, in the background.
- **Fast video with ffmpeg:** H.264 or H.265, from 720p to 4K, with fixed frame rates or **1 frame per slide (VFR)**. Several videos can be built at once.
- **Chapters:** taken from PowerPoint sections, or from slide titles, plus a chapter list ready to paste into a YouTube description.
- **Captions:** timed from the narration's own pauses. They're built into the MP4 and also saved as `.srt` and `.vtt` files.
- **Even volume:** every slide is set to the same loudness, so there are no jumps between slides, voices or decks.
- **Reliable:** cleans hidden characters from notes, retries failed slides, falls back to sentence-by-sentence, and tries three speech engines.
- **Live progress:** a table of every deck (audio, video and status), plus a list of everything being worked on right now: each slide being voiced and each video being built, with percent and time left.
- **Logs:** one log per deck, saved next to it.

## Requirements

- Windows 10 or 11
- PowerPoint (Microsoft 365)
- Windows PowerShell 5.1 (built into Windows)

## Downloads

| File | Get it from | Needed? |
|---|---|---|
| `NaturalVoiceSAPIAdapter_v0.2.9_x86_x64.zip` | [github.com/gexgd0419/NaturalVoiceSAPIAdapter/releases](https://github.com/gexgd0419/NaturalVoiceSAPIAdapter/releases) | Recommended (needed for natural voices) |
| `ms_natural_voice_en_au.zip` | [cross-plus-a.com/voice.htm](https://www.cross-plus-a.com/voice.htm) | Recommended (offline Australian voices; needs the adapter) |
| `balcon.zip` | [cross-plus-a.com/bconsole.htm](https://www.cross-plus-a.com/bconsole.htm) | Optional (backup speech engine) |
| `ffmpeg.exe` | [ffmpeg.org/download.html](https://ffmpeg.org/download.html) | Recommended (fast video, chapters, captions, even volume, smaller files) |
| `Slide-Narrator.ps1` and `Slide-Narrator.bat` | This repository | Required |

Only the app is strictly required. Windows has basic built-in voices, but they sound robotic. For natural voices, install the adapter.

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

### Recommended settings (maximum performance and quality)

| Setting | Use | Why |
|---|---|---|
| Video method | `Fast - ffmpeg H.264` | Fast, and plays everywhere |
| Resolution | `4K / 2160p` | Sharpest result. With VFR, 4K costs very little extra time or file size. |
| Frame rate | `1 per slide (VFR)` | Smallest files and quickest builds |
| Slides at once | **The number of threads your CPU has**, as long as you have about **150 MB of free RAM per thread** (for example, 32 threads needs about 4.8 GB free) | Native voices run on your CPU, one slide per thread |
| Videos at once | **The number of PowerPoint files**, up to about **5–8** | Each VFR video is light, so several build side by side. 5 at once has been tested and works well. |

If you have less free RAM, lower **Slides at once** until it fits: free RAM in MB ÷ 150. **Online** voices are limited by Microsoft, not your CPU, so keep them at about 4 slides at once.

To find your thread count, check the line `CPU: X cores / Y threads` in the app's log when it opens, or look in **Task Manager → Performance → CPU → Logical processors**.

The **Auto** boxes choose sensible values for you. Untick them to enter the numbers above yourself.

## Output

Everything is saved next to each original deck:

| File | Contents |
|---|---|
| `<deck>_narrated.pptx` | The deck with narration embedded and slides that advance automatically |
| `<deck>.mp4` | The video, if you chose one, with chapters and captions built in |
| `<deck>.srt` and `<deck>.vtt` | Captions. Use `.srt` for YouTube and most video sites, and `.vtt` for web players and learning platforms such as Moodle. |
| `<deck>_chapters.txt` | Chapter times to paste into a YouTube description |
| `<deck>_narration.log` | What happened on every slide |
| `<deck>_narration\` | Working files: each slide's text, audio and picture. Safe to delete. |

## Options

| Setting | Choices | Notes |
|---|---|---|
| Slides at once | 1–32, or Auto | How many slides are voiced in parallel. Native voices: up to one per CPU thread (about 150 MB of RAM each). Online voices: about 4. |
| Video method | Fast - ffmpeg H.264 · Fast - ffmpeg H.265 · PowerPoint | H.264 plays everywhere. H.265 files are smaller, but some browsers can't play them. PowerPoint is the only method that keeps animations, but it's slow and does one video at a time. |
| Resolution | 720p · 1080p · 1440p (ffmpeg only) · 4K | 4K is the default |
| Frame rate | 1 per slide (VFR, ffmpeg only) · 5–60 fps | VFR gives the smallest files and is quickest to build. Every frame is a keyframe. |
| Videos at once | 1–12, or Auto | Parallel ffmpeg video builds. VFR: up to the number of decks (5–8 works well). Fixed frame rates: 1–2. |
| Chapters | By section (or by slide) · By slide · None | YouTube shows chapters only if there are at least 3, each 10 seconds or longer. |
| Captions | On or off | On by default. Captions are switched off when the video starts; viewers turn them on with the player's CC button. Needs ffmpeg. |

## Writing speaker notes for narration

The voice reads your speaker notes exactly as written, and the captions show them exactly as written. So write the notes as a script to be heard, not as notes to be glanced at.

| Do | Instead of |
|---|---|
| Write full sentences, the way you'd say them aloud. Use contractions such as *it's* and *you'll*. | Bullet points, fragments and headings |
| Keep sentences under about 25 words. | Long sentences full of clauses, which also make long captions |
| Say the full term the first time, then use the acronym: "Work health and safety, or WHS, ..." | An unexplained acronym |
| Use the singular acronym, or the full term in the plural: "health and safety representatives" | Plurals and possessives such as *HSRs* or *PCBU's* |
| Write out acronyms that look like words, such as *MR*, *SOC* or *IT*, or rephrase them. | Letting the voice say "Mister" or "sock" |
| Write words, not symbols: *and*, *per cent*, *to*, *dollars*, *section 58* | *&*, *%*, *–*, *$*, */*, *s 58* |
| Write *for example*, *that is*, *and so on*, *versus*. | *e.g.*, *i.e.*, *etc.*, *vs.* |
| Name a website: "the WorkSafe Victoria website". | A full URL |
| Use full stops and new paragraphs for pauses. | *...*, *[pause]* or *(click)* |
| Aim for 40 to 150 words per slide. The voice speaks about 140 words a minute. | Empty notes on content slides, or very long notes |

A few tips:

- **Test tricky words first.** Put them in the notes of a one-slide test deck, narrate it with **Narrated PowerPoint only**, and listen. It takes a few seconds. If a word sounds wrong, rephrase the sentence. The captions show your notes as written, so avoid phonetic spellings such as "W H S".
- **Notes don't need to repeat the slide.** Explain it, give an example, or say why it matters.
- **Each slide's notes stand alone.** There's a short pause between slides, so don't split a sentence across two slides.

### A prompt for writing narration-ready notes

If you use an AI assistant to write your slides, add this to your request:

```text
I'll turn this deck into a narrated video. A text-to-speech voice will read each slide's
speaker notes aloud, and the notes will also be shown as captions, so write them as a
spoken script:

- Full, natural sentences in a warm, conversational tone, as if a teacher is talking to
  one student. Use contractions. Keep most sentences under 25 words.
- Explain and add to the slide rather than reading it out. Aim for 40 to 150 words per
  slide, and give the title slide and any section-divider slides a short spoken
  introduction too.
- Acronyms: give the full term the first time it's used, with the acronym after it, for
  example "work health and safety, or WHS". Never use plural or possessive acronyms such
  as "HSRs" or "PCBU's"; use the full term instead, for example "health and safety
  representatives". If an acronym looks like a word, such as MR, SOC or IT, use the full
  term every time.
- No symbols or shorthand: write "and", "per cent", "to", "dollars", "section 58",
  "sections 20 to 26", "for example", "that is". No "e.g.", "i.e.", "etc.", "&", "%", "/",
  dashes or brackets.
- No URLs, bullet points, headings, emojis, stage directions or labels such as "Slide 3"
  or "[pause]". Use full stops and new paragraphs for pauses.
- Write dates and numbers the way they're spoken, for example "1 July 2026" and "about
  2 thousand workers".
- Each slide's notes must make sense on their own; don't split a sentence across slides.
- Use Australian English spelling.
```

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

1. **Read:** speaker notes are read from each deck through PowerPoint automation, then cleaned of hidden characters and symbols the voices stumble on.
2. **Voice:** background workers turn each slide's notes into a WAV file. The app uses SAPI 5 directly in 64-bit, a 32-bit SAPI helper, or `balcon.exe`, whichever works for the chosen voice. ffmpeg then measures each slide's loudness, sets it to −19 LUFS (the usual level for mono speech) with a limiter to catch peaks, and converts it to 64 kbps mono AAC for the PowerPoint. Videos are encoded separately from the original WAVs, also at 64 kbps, so the audio is only compressed once.
3. **Build:** the audio is embedded in each slide, set to play automatically, and the slide advances one second after the audio ends. The deck is saved as `_narrated.pptx`.
4. **Video:** slides are exported as PNG pictures, and ffmpeg combines them with the WAV audio, adding chapters if you chose them. Captions are split into short lines and timed by finding the pauses in each slide's narration. The PowerPoint method uses PowerPoint's own **Create a Video** export, and the chapters and captions are added afterwards.

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
| 8 | Captions (.srt, .vtt and built into the MP4), even volume across slides, YouTube chapter list |
| 9 | Progress view: deck table and live activity list with percent and time left |

## Credits

- [NaturalVoiceSAPIAdapter](https://github.com/gexgd0419/NaturalVoiceSAPIAdapter) by gexgd0419
- [Balabolka and balcon](https://www.cross-plus-a.com/) by Ilya Morozov
- [FFmpeg](https://ffmpeg.org/)

These tools are not included in this repository. Download them from their sources above, and check each one's own licence.
