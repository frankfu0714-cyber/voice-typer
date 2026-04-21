# Speaking Practice Transcriber

A small Windows desktop app for English speaking practice.

## What it does

- Press your chosen hotkey to start recording your voice.
- Press the same hotkey again to stop.
- The app transcribes your speech and either pastes it back into the app you were using or types it there phrase-by-phrase while you are still speaking — depending on the mode you pick.
- The latest transcript appears in the app and is copied to your clipboard.
- You can click `Change Hotkey` and then press the exact key or key combination you want.
- The app also shows recommended hotkeys you can apply with one click.
- The app window has a vertical scrollbar so you can reach the full interface on smaller screens.
- The app plays short sound cues when recording starts, recording stops, and a transcript is ready.
- Each transcript is also added to the session history.

## Modes

There are two transcription modes, switchable from the main window:

- **Record then transcribe** — the original flow. Press hotkey, speak, press again, and one final transcript is produced and pasted. Best when you want a single clean result.
- **Live dictation** — press hotkey and start speaking. Voice-activity detection watches for short pauses and transcribes each phrase as it completes, so text appears in your target app in real time as you talk. Closest to the macOS dictation experience. Because Whisper is not a true streaming model, text arrives phrase-by-phrase rather than letter-by-letter, and only committed phrases are pasted so you never see words flicker and get revised.

## Requirements

- Windows
- Python 3.10 or newer
- A working microphone

## Setup

Open PowerShell in this folder and run:

```powershell
py -3 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
python main.py
```

## Packaged EXE

A Windows build is available at:

```text
dist\SpeakingPracticeTranscriber\SpeakingPracticeTranscriber.exe
```

## Notes

- The first transcription may take longer because the local speech model needs to load.
- The app uses the `base.en` Whisper model, which is tuned for English.
- If your microphone is not detected, check Windows microphone permissions.
- Auto-paste works best when your text cursor is already in the target input box before you press your hotkey.
- You can change the hotkey inside the app by pressing the exact shortcut you want. Examples: `F6`, `Numpad 5`, `Ctrl+Alt+R`, `Ctrl+Shift+Space`.

## Next ideas

- Add grammar feedback after each transcript.
- Save transcripts to a file.
- Push-to-talk option (hold the hotkey to record).
- Microphone picker and a live input-level meter.
- GPU (CUDA) support for faster transcription.
