import ctypes
import json
import os
import queue
import re
import sys
import tempfile
import threading
import winsound
from ctypes import wintypes
from datetime import datetime
from pathlib import Path
from typing import Optional

import numpy as np
import pyperclip
import pystray
import sounddevice as sd
import soundfile as sf
import tkinter as tk
from faster_whisper import WhisperModel
from PIL import Image as PilImage
from PIL import ImageDraw, ImageFont
from pynput import keyboard
from tkinter import messagebox, scrolledtext, ttk


# ---------- Win32 structures for per-pixel-alpha layered windows ---------- #
#
# Used by _show_floating_indicator to paint a smoothly antialiased rounded
# pill (drawn by Pillow) straight onto the screen via UpdateLayeredWindow.
# Defined at module scope so the structure layouts are stable across calls.

class _POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]


class _SIZE(ctypes.Structure):
    _fields_ = [("cx", ctypes.c_long), ("cy", ctypes.c_long)]


class _BLENDFUNCTION(ctypes.Structure):
    _fields_ = [
        ("BlendOp", ctypes.c_ubyte),
        ("BlendFlags", ctypes.c_ubyte),
        ("SourceConstantAlpha", ctypes.c_ubyte),
        ("AlphaFormat", ctypes.c_ubyte),
    ]


class _BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [
        ("biSize", ctypes.c_uint32),
        ("biWidth", ctypes.c_long),
        ("biHeight", ctypes.c_long),
        ("biPlanes", ctypes.c_uint16),
        ("biBitCount", ctypes.c_uint16),
        ("biCompression", ctypes.c_uint32),
        ("biSizeImage", ctypes.c_uint32),
        ("biXPelsPerMeter", ctypes.c_long),
        ("biYPelsPerMeter", ctypes.c_long),
        ("biClrUsed", ctypes.c_uint32),
        ("biClrImportant", ctypes.c_uint32),
    ]


APP_NAME = "VoiceTyper"
APP_ID = "Onetranscribt.VoiceTyper"

# Mutex name for single-instance check (Windows). Only one process can own it.
SINGLE_INSTANCE_MUTEX_NAME = "Local\\VoiceTyper_SingleInstance"
ERROR_ALREADY_EXISTS = 183
SW_RESTORE = 9


def get_resource_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
    return Path(__file__).resolve().parent


def get_settings_dir() -> Path:
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        settings_dir = Path(local_app_data) / "VoiceTyper"
    else:
        settings_dir = Path.home() / ".onetranscribt"
    settings_dir.mkdir(parents=True, exist_ok=True)
    return settings_dir


def _bring_existing_window_to_front() -> bool:
    """Find the app's window by title and bring it to front. Returns True if found."""
    user32 = ctypes.windll.user32
    found: list[int] = []

    def enum_callback(hwnd: int, _lparam: int) -> int:
        if not user32.IsWindowVisible(hwnd):
            return 1
        length = user32.GetWindowTextLengthW(hwnd) + 1
        buf = ctypes.create_unicode_buffer(length)
        user32.GetWindowTextW(hwnd, buf, length)
        title = buf.value or ""
        if title.strip().startswith(APP_NAME):
            found.append(hwnd)
            return 0
        return 1

    WNDENUMPROC = ctypes.CFUNCTYPE(ctypes.c_int, wintypes.HWND, wintypes.LPARAM)
    user32.EnumWindows(WNDENUMPROC(enum_callback), 0)
    if found:
        hwnd = found[0]
        user32.ShowWindow(hwnd, SW_RESTORE)
        user32.SetForegroundWindow(hwnd)
        return True
    return False


def ensure_single_instance() -> Optional[object]:
    """
    Ensure only one instance of the app runs regardless of how it was launched
    (packaged exe, python main.py, setup_and_run.ps1, etc.).

    Strategy:
      1. Look for an existing window by title — catches any running instance,
         even legacy ones that predate the mutex.
      2. Grab a named mutex — catches race conditions where two new instances
         start at the same moment before either creates a window.

    Returns the mutex handle so the caller keeps it alive for the process lifetime.
    """
    if sys.platform != "win32":
        return None

    # Step 1: window-title check — works across all launch methods
    if _bring_existing_window_to_front():
        sys.exit(0)

    # Step 2: mutex — guards against simultaneous startup race
    kernel32 = ctypes.windll.kernel32
    mutex = kernel32.CreateMutexW(None, True, SINGLE_INSTANCE_MUTEX_NAME)
    if not mutex:
        return None
    if kernel32.GetLastError() == ERROR_ALREADY_EXISTS:
        kernel32.CloseHandle(mutex)
        _bring_existing_window_to_front()
        sys.exit(0)
    return mutex


class AudioRecorder:
    def __init__(self, sample_rate: int = 16000, channels: int = 1) -> None:
        self.sample_rate = sample_rate
        self.channels = channels
        self._frames: list[np.ndarray] = []
        self._stream: Optional[sd.InputStream] = None
        self._lock = threading.Lock()
        self.is_recording = False

    def start(self) -> None:
        if self.is_recording:
            return

        self._frames = []
        self._stream = sd.InputStream(
            samplerate=self.sample_rate,
            channels=self.channels,
            dtype="float32",
            callback=self._audio_callback,
        )
        self._stream.start()
        self.is_recording = True

    def stop(self) -> str:
        if not self.is_recording or self._stream is None:
            raise RuntimeError("Recorder is not running.")

        self._stream.stop()
        self._stream.close()
        self._stream = None
        self.is_recording = False

        with self._lock:
            if not self._frames:
                raise RuntimeError("No audio was captured.")
            audio = np.concatenate(self._frames, axis=0)

        fd, path = tempfile.mkstemp(suffix=".wav", prefix="speaking-practice-")
        os.close(fd)
        sf.write(path, audio, self.sample_rate)
        return path

    def _audio_callback(self, indata, frames, time_info, status) -> None:
        del frames, time_info
        if status:
            print(f"Audio warning: {status}")
        with self._lock:
            self._frames.append(indata.copy())


class StreamingAudioRecorder:
    """
    Live recorder with energy-based voice-activity detection.

    While running, it listens continuously. Whenever the microphone goes
    quiet for `min_silence_ms`, it emits the preceding speech as a WAV file
    via `on_segment(path)`, so a downstream transcriber can process it
    while the user keeps speaking. Short blips (< `min_speech_ms`) and
    sub-threshold noise are discarded. Very long utterances are force-cut
    every `max_segment_ms` so a user who keeps talking still sees progress.

    Design notes:
      * Audio-callback work is kept to math only — file I/O happens on a
        separate writer thread so PortAudio is never blocked.
      * A small pre-roll buffer is prepended to each segment so the first
        syllable isn't clipped when speech starts.
      * Segments are serialised through a single writer thread, so
        `on_segment` always fires in chronological order.
    """

    # Lowered from 0.02 → 0.008 so quieter / further-from-mic speech still
    # registers as speech. 0.008 is roughly "conversational voice on a
    # built-in laptop mic"; ambient room noise is usually 0.002 or below.
    DEFAULT_SILENCE_RMS = 0.008
    DEFAULT_MIN_SPEECH_MS = 200
    # Raised from 500 → 1000 ms so each committed phrase is a complete
    # sentence-sized chunk of audio (typical sentence-end pause). Whisper
    # gets much more context per call, so accuracy approaches batch mode.
    # Trade-off: text lands in sentence-sized bursts rather than mid-sentence.
    DEFAULT_MIN_SILENCE_MS = 1000
    # Raised from 12s → 20s so long run-on thoughts still commit eventually
    # without forcing a mid-sentence cut.
    DEFAULT_MAX_SEGMENT_MS = 20000
    DEFAULT_PRE_ROLL_MS = 300
    DEFAULT_BLOCK_MS = 30

    def __init__(
        self,
        on_segment,
        sample_rate: int = 16000,
        channels: int = 1,
        silence_threshold_rms: float = DEFAULT_SILENCE_RMS,
        min_speech_ms: int = DEFAULT_MIN_SPEECH_MS,
        min_silence_ms: int = DEFAULT_MIN_SILENCE_MS,
        max_segment_ms: int = DEFAULT_MAX_SEGMENT_MS,
        pre_roll_ms: int = DEFAULT_PRE_ROLL_MS,
        block_ms: int = DEFAULT_BLOCK_MS,
    ) -> None:
        self.on_segment = on_segment
        self.sample_rate = sample_rate
        self.channels = channels
        self.silence_threshold_rms = silence_threshold_rms
        self.min_speech_samples = int(min_speech_ms * sample_rate / 1000)
        self.min_silence_samples = int(min_silence_ms * sample_rate / 1000)
        self.max_segment_samples = int(max_segment_ms * sample_rate / 1000)
        self.pre_roll_samples = int(pre_roll_ms * sample_rate / 1000)
        self.block_size = max(1, int(block_ms * sample_rate / 1000))

        self._stream: Optional[sd.InputStream] = None
        self._lock = threading.Lock()
        self._state = "silent"              # "silent" | "speaking"
        self._pre_roll: list[np.ndarray] = []
        self._pre_roll_len = 0
        self._current: list[np.ndarray] = []
        self._current_len = 0
        # Size of the pre-roll padding that was promoted into the current
        # segment at speech onset. Tracked separately so the "minimum speech"
        # check doesn't count silence pre-roll as real speech.
        self._pre_roll_len_at_onset = 0
        self._trailing_silence = 0

        self._writer_queue: "queue.Queue[Optional[np.ndarray]]" = queue.Queue()
        self._writer_thread: Optional[threading.Thread] = None
        self.is_recording = False

    def start(self) -> None:
        if self.is_recording:
            return
        self._reset_segment_state()
        self._writer_queue = queue.Queue()
        self._writer_thread = threading.Thread(
            target=self._writer_loop, daemon=True
        )
        self._writer_thread.start()

        self._stream = sd.InputStream(
            samplerate=self.sample_rate,
            channels=self.channels,
            dtype="float32",
            blocksize=self.block_size,
            callback=self._audio_callback,
        )
        self._stream.start()
        self.is_recording = True

    def stop(self) -> None:
        if not self.is_recording:
            return
        if self._stream is not None:
            try:
                self._stream.stop()
                self._stream.close()
            except Exception:
                pass
            self._stream = None
        self.is_recording = False

        # Flush any in-progress phrase that is long enough to be useful.
        with self._lock:
            speech_len = (
                self._current_len
                - self._pre_roll_len_at_onset
                - self._trailing_silence
            )
            if self._state == "speaking" and speech_len >= self.min_speech_samples:
                audio = self._drain_current_locked()
            else:
                audio = None
                self._current = []
                self._current_len = 0
            self._pre_roll_len_at_onset = 0
            self._trailing_silence = 0
            self._state = "silent"

        if audio is not None:
            self._writer_queue.put(audio)
        # Sentinel — tells writer thread to drain its queue then exit.
        self._writer_queue.put(None)

        # Wait for the writer thread to flush so every on_segment call has
        # completed before stop() returns. This prevents a race where the
        # caller pushes its own sentinel onto its downstream queue before
        # the writer has delivered the final segment.
        if self._writer_thread is not None:
            self._writer_thread.join(timeout=3.0)
            self._writer_thread = None

    def _reset_segment_state(self) -> None:
        self._state = "silent"
        self._pre_roll = []
        self._pre_roll_len = 0
        self._current = []
        self._current_len = 0
        self._pre_roll_len_at_onset = 0
        self._trailing_silence = 0

    def _audio_callback(self, indata, frames, time_info, status) -> None:
        del frames, time_info
        if status:
            print(f"Audio warning: {status}")
        if indata.size == 0:
            return
        block = indata.copy()
        rms = float(np.sqrt(np.mean(block * block)))
        is_speech = rms >= self.silence_threshold_rms

        audio_to_emit: Optional[np.ndarray] = None

        with self._lock:
            if self._state == "silent":
                # Keep a rolling ~pre_roll_ms buffer of recent silence so the
                # first syllable isn't lost when speech starts.
                self._pre_roll.append(block)
                self._pre_roll_len += len(block)
                while (
                    self._pre_roll
                    and self._pre_roll_len - len(self._pre_roll[0]) >= self.pre_roll_samples
                ):
                    self._pre_roll_len -= len(self._pre_roll[0])
                    self._pre_roll.pop(0)

                if is_speech:
                    # Promote pre-roll into the new segment. Remember how
                    # much of the segment is pre-roll so the min-speech
                    # check below doesn't count that silence as speech.
                    self._current.extend(self._pre_roll)
                    self._current_len += self._pre_roll_len
                    self._pre_roll_len_at_onset = self._pre_roll_len
                    self._pre_roll = []
                    self._pre_roll_len = 0
                    self._current.append(block)
                    self._current_len += len(block)
                    self._trailing_silence = 0
                    self._state = "speaking"
            else:
                # state == "speaking"
                self._current.append(block)
                self._current_len += len(block)
                if is_speech:
                    self._trailing_silence = 0
                else:
                    self._trailing_silence += len(block)

                if self._trailing_silence >= self.min_silence_samples:
                    speech_len = (
                        self._current_len
                        - self._pre_roll_len_at_onset
                        - self._trailing_silence
                    )
                    if speech_len >= self.min_speech_samples:
                        audio_to_emit = self._drain_current_locked()
                    else:
                        # Blip — discard.
                        self._current = []
                        self._current_len = 0
                    self._pre_roll_len_at_onset = 0
                    self._trailing_silence = 0
                    self._state = "silent"
                elif self._current_len >= self.max_segment_samples:
                    # Force-cut a very long utterance and keep listening.
                    audio_to_emit = self._drain_current_locked()
                    # After a force-cut, the remaining live speech has no
                    # pre-roll of its own.
                    self._pre_roll_len_at_onset = 0
                    self._trailing_silence = 0
                    # Stay in "speaking" state so the next block extends the
                    # same thought rather than waiting for a fresh onset.

        if audio_to_emit is not None:
            self._writer_queue.put(audio_to_emit)

    def _drain_current_locked(self) -> np.ndarray:
        """Must be called with self._lock held."""
        audio = np.concatenate(self._current, axis=0)
        self._current = []
        self._current_len = 0
        return audio

    def _writer_loop(self) -> None:
        while True:
            item = self._writer_queue.get()
            if item is None:
                return
            try:
                fd, path = tempfile.mkstemp(suffix=".wav", prefix="voicetyper-live-")
                os.close(fd)
                sf.write(path, item, self.sample_rate)
            except OSError as exc:
                print(f"Live segment write failed: {exc}")
                continue
            try:
                self.on_segment(path)
            except Exception as exc:
                print(f"Live segment callback failed: {exc}")


class SpeakingPracticeApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_NAME)
        self.root.geometry("1040x760")
        self.root.minsize(920, 640)
        self.resource_dir = get_resource_dir()
        self.settings_path = get_settings_dir() / "app_settings.json"
        self.icon_ico_path = self.resource_dir / "assets" / "app_icon.ico"
        self.icon_path = self.resource_dir / "assets" / "app_icon.png"
        self.icon_image: Optional[tk.PhotoImage] = None

        self.recorder = AudioRecorder()
        self.ui_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.model: Optional[WhisperModel] = None
        self.model_lock = threading.Lock()
        self.hotkey_listener: Optional[keyboard.Listener] = None
        self.keyboard_controller = keyboard.Controller()
        self.transcribe_thread: Optional[threading.Thread] = None
        self.is_transcribing = False
        self.target_window_handle: Optional[int] = None
        self.auto_paste_enabled = True
        self.hotkey_tokens = self._load_hotkey_setting()
        self.hotkey_display = tk.StringVar(value=self._format_hotkey_for_display(self.hotkey_tokens))
        self._lang_options: dict[str, str] = {
            "English": "en",
            "Chinese / Mandarin (中文)": "zh",
        }
        self._lang_display_reverse: dict[str, str] = {v: k for k, v in self._lang_options.items()}
        self.language: str = self._load_language_setting()
        self.current_model_lang: Optional[str] = None
        self.pressed_keys: set[str] = set()
        self.hotkey_active = False
        self.capture_listener: Optional[keyboard.Listener] = None
        self.capture_dialog: Optional[tk.Toplevel] = None
        self.capture_preview_text = tk.StringVar(value="")
        self.is_capturing_hotkey = False
        self.capture_pressed_keys: set[str] = set()
        self.capture_modifiers: set[str] = set()
        self.capture_main_key: Optional[str] = None
        self.app_window_handle = 0
        self.sound_feedback_enabled = True
        self.recommended_hotkeys: list[tuple[str, tuple[str, ...], str]] = [
            ("F6", ("f6",), "Simple and easy to reach"),
            ("F8", ("f8",), "Best for most users"),
            ("Ctrl+Alt+R", ("ctrl", "alt", "r"), "Memorable and low conflict"),
            ("Numpad 5", ("numpad5",), "Good for numpad users"),
        ]

        # Transcription mode: "batch" (record → transcribe at stop) or
        # "live" (stream phrases while the user is still speaking).
        self.mode: str = self._load_mode_setting()
        self.streaming_recorder: Optional[StreamingAudioRecorder] = None
        self.live_transcribe_queue: Optional[queue.Queue] = None
        self.live_worker_thread: Optional[threading.Thread] = None
        self.live_phrase_count = 0
        self.live_target_window_handle: Optional[int] = None
        # "Listening..." indicator state (see _start_live_indicator).
        self._live_indicator_after_id: Optional[str] = None
        self._live_indicator_phase = 0
        # Floating always-on-top overlay state (see _show_floating_indicator).
        self._floating_indicator: Optional[tk.Toplevel] = None
        self._floating_indicator_text_label: Optional[tk.Label] = None
        self._floating_indicator_dot_label: Optional[tk.Label] = None
        self._floating_indicator_prefix_label: Optional[tk.Label] = None
        self._floating_indicator_lang_label: Optional[tk.Label] = None
        self._floating_indicator_after_id: Optional[str] = None
        self._floating_indicator_phase = 0

        self.status_text = tk.StringVar(value=self._initial_status_text())
        self.last_saved_text = ""
        self.tray_icon: Optional[pystray.Icon] = None

        self._configure_windows_app_id()
        self._configure_app_icon()
        self._configure_styles()
        self._build_ui()
        self.root.update_idletasks()
        self.app_window_handle = int(self.root.winfo_id())
        self._start_hotkey_listener()
        self._pump_ui_queue()
        # Close button hides to tray; tray Quit truly exits
        self.root.protocol("WM_DELETE_WINDOW", self._hide_to_tray)
        self._setup_tray()

        # Warm up the model in the background when we start in live mode so
        # the first phrase isn't blocked on a cold Whisper load.
        if self.mode == "live":
            threading.Thread(target=self._preload_model, daemon=True).start()

    def _configure_windows_app_id(self) -> None:
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_ID)
        except (AttributeError, OSError):
            pass

    def _configure_app_icon(self) -> None:
        icon_loaded = False
        try:
            if self.icon_ico_path.exists():
                self.root.iconbitmap(default=str(self.icon_ico_path))
                icon_loaded = True
        except tk.TclError:
            pass

        if icon_loaded or not self.icon_path.exists():
            return

        try:
            self.icon_image = tk.PhotoImage(file=str(self.icon_path))
            self.root.iconphoto(True, self.icon_image)
        except tk.TclError:
            self.icon_image = None

    def _configure_styles(self) -> None:
        self.root.configure(bg="#0f172a")
        style = ttk.Style(self.root)
        style.theme_use("clam")

        palette = {
            "bg": "#0f172a",
            "surface": "#111827",
            "surface_alt": "#1f2937",
            "card": "#ffffff",
            "muted": "#64748b",
            "text": "#0f172a",
            "accent": "#2563eb",
            "accent_hover": "#1d4ed8",
            "success_bg": "#dcfce7",
            "success_fg": "#166534",
            "badge_bg": "#dbeafe",
            "badge_fg": "#1d4ed8",
            "border": "#dbe3f0",
            "soft": "#f8fafc",
        }

        style.configure("App.TFrame", background=palette["bg"])
        style.configure("Card.TFrame", background=palette["card"], relief="flat")
        style.configure("HeaderCard.TFrame", background=palette["surface"])
        style.configure("Section.TLabelframe", background=palette["card"], borderwidth=0)
        style.configure("Section.TLabelframe.Label", background=palette["card"], foreground=palette["text"])

        style.configure(
            "Title.TLabel",
            background=palette["surface"],
            foreground="#ffffff",
            font=("Segoe UI", 24, "bold"),
        )
        style.configure(
            "Subtitle.TLabel",
            background=palette["surface"],
            foreground="#cbd5e1",
            font=("Segoe UI", 11),
        )
        style.configure(
            "CardTitle.TLabel",
            background=palette["card"],
            foreground=palette["text"],
            font=("Segoe UI", 12, "bold"),
        )
        style.configure(
            "Body.TLabel",
            background=palette["card"],
            foreground=palette["muted"],
            font=("Segoe UI", 10),
        )
        style.configure(
            "Footer.TLabel",
            background=palette["bg"],
            foreground="#475569",
            font=("Segoe UI", 9),
        )
        style.configure(
            "HeaderCredit.TLabel",
            background=palette["surface"],
            foreground="#64748b",
            font=("Segoe UI", 10),
        )
        style.configure(
            "Status.TLabel",
            background=palette["success_bg"],
            foreground=palette["success_fg"],
            font=("Segoe UI", 10, "bold"),
            padding=(12, 8),
        )
        style.configure(
            "Badge.TLabel",
            background=palette["badge_bg"],
            foreground=palette["badge_fg"],
            font=("Segoe UI", 10, "bold"),
            padding=(10, 6),
        )
        style.configure(
            "Hint.TLabel",
            background=palette["card"],
            foreground=palette["muted"],
            font=("Segoe UI", 9),
        )
        style.configure(
            "Primary.TButton",
            font=("Segoe UI", 10, "bold"),
            padding=(16, 12),
            background=palette["accent"],
            foreground="#ffffff",
            borderwidth=0,
        )
        style.map(
            "Primary.TButton",
            background=[("active", palette["accent_hover"]), ("pressed", palette["accent_hover"])],
            foreground=[("disabled", "#cbd5e1")],
        )
        style.configure(
            "Secondary.TButton",
            font=("Segoe UI", 10),
            padding=(14, 10),
            background=palette["soft"],
            foreground=palette["text"],
            borderwidth=1,
            relief="solid",
        )
        style.map(
            "Secondary.TButton",
            background=[("active", "#eef2ff"), ("pressed", "#e0e7ff")],
        )
        style.configure(
            "Preset.TButton",
            font=("Segoe UI", 9, "bold"),
            padding=(10, 8),
            background=palette["soft"],
            foreground=palette["text"],
            borderwidth=1,
            relief="solid",
        )
        style.map(
            "Preset.TButton",
            background=[("active", "#eff6ff"), ("pressed", "#dbeafe")],
        )
        style.configure(
            "App.TCheckbutton",
            background=palette["card"],
            foreground=palette["text"],
            font=("Segoe UI", 10),
        )

    def _build_ui(self) -> None:
        outer = ttk.Frame(self.root, style="App.TFrame")
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(0, weight=1)

        self.scroll_canvas = tk.Canvas(
            outer,
            bg="#0f172a",
            highlightthickness=0,
            borderwidth=0,
        )
        self.scroll_canvas.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=self.scroll_canvas.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.scroll_canvas.configure(yscrollcommand=scrollbar.set)

        shell = ttk.Frame(self.scroll_canvas, style="App.TFrame", padding=20)
        shell.columnconfigure(0, weight=1)
        shell.rowconfigure(3, weight=1)
        self.canvas_window = self.scroll_canvas.create_window((0, 0), window=shell, anchor="nw")
        shell.bind("<Configure>", self._update_scroll_region)
        self.scroll_canvas.bind("<Configure>", self._resize_scrollable_content)
        self.scroll_canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        header = ttk.Frame(shell, style="HeaderCard.TFrame", padding=24)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)

        title_row = ttk.Frame(header, style="HeaderCard.TFrame")
        title_row.grid(row=0, column=0, sticky="ew")
        title_row.columnconfigure(0, weight=0)
        title_row.columnconfigure(1, weight=1)

        ttk.Label(title_row, text="VoiceTyper", style="Title.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            title_row,
            text="by goldotaku  ·  goldotakutw.com",
            style="HeaderCredit.TLabel",
        ).grid(row=0, column=1, sticky="e", padx=(16, 0))

        ttk.Label(
            header,
            text="Record with one hotkey, get instant transcription, and paste it back where you are working.",
            style="Subtitle.TLabel",
            wraplength=720,
        ).grid(row=1, column=0, sticky="w", pady=(6, 18))
        ttk.Label(header, textvariable=self.status_text, style="Status.TLabel", wraplength=900).grid(
            row=2, column=0, sticky="ew"
        )

        top_grid = ttk.Frame(shell, style="App.TFrame")
        top_grid.grid(row=1, column=0, sticky="ew", pady=(18, 18))
        top_grid.columnconfigure(0, weight=3)
        top_grid.columnconfigure(1, weight=2)

        controls_card = ttk.Frame(top_grid, style="Card.TFrame", padding=20)
        controls_card.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        controls_card.columnconfigure(0, weight=1)

        ttk.Label(controls_card, text="Quick Actions", style="CardTitle.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            controls_card,
            text="Start talking, copy the result, or clear the current transcript.",
            style="Body.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(4, 16))

        control_buttons = ttk.Frame(controls_card, style="Card.TFrame")
        control_buttons.grid(row=2, column=0, sticky="w")

        self.toggle_button = ttk.Button(
            control_buttons,
            text=self._record_button_text("Start"),
            command=self.toggle_recording,
            style="Primary.TButton",
        )
        self.toggle_button.pack(side="left")

        self.copy_button = ttk.Button(
            control_buttons,
            text="Copy Transcript",
            command=self.copy_transcript,
            style="Secondary.TButton",
        )
        self.copy_button.pack(side="left", padx=(10, 0))

        self.clear_button = ttk.Button(
            control_buttons,
            text="Clear",
            command=self.clear_transcript,
            style="Secondary.TButton",
        )
        self.clear_button.pack(side="left", padx=(10, 0))

        mode_row = ttk.Frame(controls_card, style="Card.TFrame")
        mode_row.grid(row=3, column=0, sticky="w", pady=(18, 0))

        ttk.Label(mode_row, text="Mode:", style="Body.TLabel").pack(
            side="left", padx=(0, 10)
        )
        self.mode_var = tk.StringVar(value=self.mode)
        self.mode_radio_batch = ttk.Radiobutton(
            mode_row,
            text="Record then transcribe",
            variable=self.mode_var,
            value="batch",
            command=self._on_mode_change,
            style="App.TCheckbutton",
        )
        self.mode_radio_batch.pack(side="left", padx=(0, 14))
        self.mode_radio_live = ttk.Radiobutton(
            mode_row,
            text="Live dictation (type while speaking)",
            variable=self.mode_var,
            value="live",
            command=self._on_mode_change,
            style="App.TCheckbutton",
        )
        self.mode_radio_live.pack(side="left")

        self.auto_paste_var = tk.BooleanVar(value=True)
        self.auto_paste_check = ttk.Checkbutton(
            controls_card,
            text="Auto-paste transcript into the app I was using when I pressed the hotkey",
            variable=self.auto_paste_var,
            command=self._sync_options,
            style="App.TCheckbutton",
        )
        self.auto_paste_check.grid(row=4, column=0, sticky="w", pady=(12, 0))

        lang_row = ttk.Frame(controls_card, style="Card.TFrame")
        lang_row.grid(row=5, column=0, sticky="w", pady=(14, 0))

        ttk.Label(lang_row, text="Transcription Language:", style="Body.TLabel").pack(
            side="left", padx=(0, 10)
        )
        self.lang_display_var = tk.StringVar(
            value=self._lang_display_reverse.get(self.language, "English")
        )
        self.lang_combo = ttk.Combobox(
            lang_row,
            textvariable=self.lang_display_var,
            values=list(self._lang_options.keys()),
            state="readonly",
            width=28,
        )
        self.lang_combo.pack(side="left")
        self.lang_combo.bind("<<ComboboxSelected>>", self._on_language_change)

        hotkey_card = ttk.Frame(top_grid, style="Card.TFrame", padding=20)
        hotkey_card.grid(row=0, column=1, sticky="nsew")
        hotkey_card.columnconfigure(0, weight=1)

        ttk.Label(hotkey_card, text="Current Hotkey", style="CardTitle.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            hotkey_card,
            text="Change it anytime or use a recommended default.",
            style="Body.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(4, 14))

        hotkey_row = ttk.Frame(hotkey_card, style="Card.TFrame")
        hotkey_row.grid(row=2, column=0, sticky="ew")
        hotkey_row.columnconfigure(1, weight=1)

        ttk.Label(hotkey_row, text="Active shortcut", style="Body.TLabel").grid(
            row=0, column=0, sticky="w", padx=(0, 10)
        )
        self.hotkey_value_label = ttk.Label(
            hotkey_row,
            textvariable=self.hotkey_display,
            style="Badge.TLabel",
        )
        self.hotkey_value_label.grid(row=0, column=1, sticky="w")

        self.hotkey_button = ttk.Button(
            hotkey_card,
            text="Change Hotkey",
            command=self.start_hotkey_capture,
            style="Secondary.TButton",
        )
        self.hotkey_button.grid(row=3, column=0, sticky="w", pady=(16, 8))

        ttk.Label(
            hotkey_card,
            text="Tip: pick something easy to reach but unlikely to clash with your other apps.",
            style="Hint.TLabel",
            wraplength=280,
        ).grid(row=4, column=0, sticky="w")

        recommendations_card = ttk.Frame(shell, style="Card.TFrame", padding=20)
        recommendations_card.grid(row=2, column=0, sticky="ew", pady=(0, 18))
        recommendations_card.columnconfigure(0, weight=1)
        recommendations_card.columnconfigure(1, weight=1)

        ttk.Label(recommendations_card, text="Recommended Hotkeys", style="CardTitle.TLabel").grid(
            row=0, column=0, sticky="w", columnspan=2
        )
        ttk.Label(
            recommendations_card,
            text="If you are not sure what feels best, start with one of these safe choices.",
            style="Body.TLabel",
        ).grid(row=1, column=0, sticky="w", columnspan=2, pady=(4, 14))

        for index, (label, tokens, description) in enumerate(self.recommended_hotkeys):
            card = ttk.Frame(recommendations_card, style="Card.TFrame", padding=12)
            row = 2 + index // 2
            column = index % 2
            padx = (0, 8) if column == 0 else (8, 0)
            card.grid(row=row, column=column, sticky="ew", padx=padx, pady=6)
            card.columnconfigure(0, weight=1)

            ttk.Label(card, text=label, style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
            ttk.Label(card, text=description, style="Body.TLabel").grid(
                row=1, column=0, sticky="w", pady=(4, 10)
            )
            ttk.Button(
                card,
                text=f"Use {label}",
                command=lambda selected=tokens: self.apply_recommended_hotkey(selected),
                style="Preset.TButton",
            ).grid(row=2, column=0, sticky="w")

        content_grid = ttk.Frame(shell, style="App.TFrame")
        content_grid.grid(row=3, column=0, sticky="nsew")
        content_grid.columnconfigure(0, weight=3)
        content_grid.columnconfigure(1, weight=2)
        content_grid.rowconfigure(0, weight=1)

        transcript_card = ttk.Frame(content_grid, style="Card.TFrame", padding=20)
        transcript_card.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        transcript_card.columnconfigure(0, weight=1)
        transcript_card.rowconfigure(1, weight=1)

        ttk.Label(transcript_card, text="Latest Transcript", style="CardTitle.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        self.transcript_box = scrolledtext.ScrolledText(
            transcript_card,
            wrap="word",
            height=16,
            font=("Consolas", 11),
            relief="flat",
            borderwidth=0,
            bg="#f8fafc",
            fg="#0f172a",
            insertbackground="#0f172a",
            padx=12,
            pady=12,
        )
        self.transcript_box.grid(row=1, column=0, sticky="nsew", pady=(12, 0))

        history_card = ttk.Frame(content_grid, style="Card.TFrame", padding=20)
        history_card.grid(row=0, column=1, sticky="nsew")
        history_card.columnconfigure(0, weight=1)
        history_card.rowconfigure(1, weight=1)

        ttk.Label(history_card, text="Session History", style="CardTitle.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        self.history_box = scrolledtext.ScrolledText(
            history_card,
            wrap="word",
            height=16,
            font=("Consolas", 10),
            relief="flat",
            borderwidth=0,
            bg="#f8fafc",
            fg="#334155",
            insertbackground="#0f172a",
            padx=12,
            pady=12,
            state="disabled",
        )
        self.history_box.grid(row=1, column=0, sticky="nsew", pady=(12, 0))

        # Footer
        footer = ttk.Frame(shell, style="App.TFrame", padding=(0, 18, 0, 8))
        footer.grid(row=4, column=0, sticky="ew")
        footer.columnconfigure(0, weight=1)

        ttk.Label(
            footer,
            text="Made by goldotaku  ·  goldotakutw.com",
            style="Footer.TLabel",
        ).grid(row=0, column=0)

    def _update_scroll_region(self, _event=None) -> None:
        self.scroll_canvas.configure(scrollregion=self.scroll_canvas.bbox("all"))

    def _resize_scrollable_content(self, event) -> None:
        self.scroll_canvas.itemconfigure(self.canvas_window, width=event.width)

    def _on_mousewheel(self, event) -> None:
        widget = self.root.winfo_containing(event.x_root, event.y_root)
        if widget is None:
            return
        if widget in {self.transcript_box, self.history_box}:
            return
        self.scroll_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _start_hotkey_listener(self) -> None:
        if self.hotkey_listener is not None:
            self.hotkey_listener.stop()
        self.pressed_keys.clear()
        self.hotkey_active = False
        self.hotkey_listener = keyboard.Listener(
            on_press=self._on_global_key_press,
            on_release=self._on_global_key_release,
        )
        self.hotkey_listener.start()

    def _sync_options(self) -> None:
        self.auto_paste_enabled = self.auto_paste_var.get()

    def _capture_target_window(self) -> None:
        try:
            foreground_window = ctypes.windll.user32.GetForegroundWindow()
        except Exception:
            self.target_window_handle = None
            return

        app_window = self.app_window_handle
        self.target_window_handle = (
            int(foreground_window) if foreground_window and foreground_window != app_window else None
        )

    def _on_global_key_press(self, key) -> None:
        if self.is_capturing_hotkey:
            return

        token = self._key_to_token(key)
        if token is None:
            return

        self.pressed_keys.add(token)
        if self._hotkey_is_pressed() and not self.hotkey_active:
            self.hotkey_active = True
            self._capture_target_window()
            self.ui_queue.put(("toggle", None))

    def _on_global_key_release(self, key) -> None:
        token = self._key_to_token(key)
        if token is not None:
            self.pressed_keys.discard(token)
        if not self._hotkey_is_pressed():
            self.hotkey_active = False

    def _hotkey_is_pressed(self) -> bool:
        return set(self.hotkey_tokens).issubset(self.pressed_keys)

    def _pump_ui_queue(self) -> None:
        while not self.ui_queue.empty():
            event, payload = self.ui_queue.get()
            if event == "status":
                self.status_text.set(str(payload))
            elif event == "transcript":
                self._show_transcript(str(payload))
            elif event == "error":
                self._show_error(str(payload))
            elif event == "toggle":
                self.toggle_recording()
            elif event == "capture_preview":
                self.capture_preview_text.set(str(payload))
            elif event == "capture_complete":
                self._finish_hotkey_capture(tuple(payload))
            elif event == "live_phrase":
                self._handle_live_phrase(str(payload))
            elif event == "live_done":
                self._handle_live_done()
        self.root.after(150, self._pump_ui_queue)

    def toggle_recording(self) -> None:
        if self.mode == "live":
            if self.streaming_recorder is not None and self.streaming_recorder.is_recording:
                self._stop_live_recording()
            else:
                self._start_live_recording()
            return

        if self.is_transcribing:
            self.status_text.set("Transcription in progress. Please wait a moment.")
            return

        if self.recorder.is_recording:
            self._stop_recording()
        else:
            self._start_recording()

    def _start_recording(self) -> None:
        try:
            self.recorder.start()
        except Exception as exc:  # pragma: no cover - hardware dependent
            self._show_error(f"Could not start recording: {exc}")
            return

        self.status_text.set(
            f"Recording... Press {self.hotkey_display.get()} again to stop and transcribe."
        )
        self._set_window_title("Recording...")
        self.toggle_button.config(text=self._record_button_text("Stop"))
        self._play_feedback("recording_started")
        # Floating overlay mirrors the live-mode one, but says "Listening"
        # while we're still capturing audio. Flipped to "Transcribing" in
        # _stop_recording below.
        self._show_floating_indicator("🎤  Listening")

    def _stop_recording(self) -> None:
        try:
            audio_path = self.recorder.stop()
        except Exception as exc:
            self._show_error(f"Could not stop recording: {exc}")
            self.toggle_button.config(text=self._record_button_text("Start"))
            return

        self.toggle_button.config(text=self._record_button_text("Start"))
        self.status_text.set("Recording stopped. Transcribing your speech...")
        self._set_window_title("Transcribing...")
        self.is_transcribing = True
        self._play_feedback("recording_stopped")
        self._set_indicator_prefix("🎤  Transcribing")

        self.transcribe_thread = threading.Thread(
            target=self._transcribe_audio,
            args=(audio_path,),
            daemon=True,
        )
        self.transcribe_thread.start()

    # --------------------------------------------------------------- live ---
    #
    # Live dictation flow:
    #   1. Hotkey press → capture target window, start StreamingAudioRecorder,
    #      spin up a transcription worker thread.
    #   2. Recorder's VAD detects a pause → emits a WAV → pushed onto the
    #      live_transcribe_queue.
    #   3. Worker thread pops, transcribes with faster-whisper, and enqueues
    #      the text as ("live_phrase", text) on the main ui_queue.
    #   4. UI thread appends to the transcript box, history, and pastes into
    #      the target app via clipboard + Ctrl+V (focusing it once on the
    #      first phrase; subsequent phrases assume focus is still there).
    #   5. Hotkey press → recorder.stop() flushes a trailing phrase and
    #      pushes a sentinel; worker finishes pending items and emits a
    #      ("live_done", _) event that returns the UI to ready state.

    def _start_live_recording(self) -> None:
        if self.streaming_recorder is not None and self.streaming_recorder.is_recording:
            return

        # Normally _on_global_key_press has just captured the foreground
        # window, but if the user clicked the toggle button instead we need
        # to check again — otherwise we'd paste into a stale handle.
        if self.target_window_handle is None:
            self._capture_target_window()

        self.live_transcribe_queue = queue.Queue()
        self.live_phrase_count = 0
        self.live_target_window_handle = self.target_window_handle
        # Fresh transcript box for this dictation session so appended phrases
        # aren't mixed with a previous run.
        self.transcript_box.delete("1.0", "end")
        self.last_saved_text = ""
        # Insert an animated "🎤 Listening..." indicator that sits at the end
        # of the transcript box while we're recording. Phrases get inserted
        # BEFORE it (see _handle_live_phrase) so it stays trailing.
        self._start_live_indicator()

        self.streaming_recorder = StreamingAudioRecorder(
            on_segment=self._on_live_segment_ready,
        )

        try:
            self.streaming_recorder.start()
        except Exception as exc:  # pragma: no cover - hardware dependent
            self.streaming_recorder = None
            self.live_transcribe_queue = None
            self._show_error(f"Could not start live dictation: {exc}")
            return

        self.live_worker_thread = threading.Thread(
            target=self._live_transcribe_loop,
            daemon=True,
        )
        self.live_worker_thread.start()

        # Show a floating always-on-top "🎤 Listening..." widget so the user
        # can see from any app that VoiceTyper is active. Non-invasive — it
        # doesn't touch the target app's text at all.
        self._show_floating_indicator()

        self.toggle_button.config(text=self._record_button_text("Stop"))
        self._set_window_title("Recording...")
        self._play_feedback("recording_started")
        if self.model is None or self.current_model_lang != self.language:
            self.status_text.set(
                "Loading the speech model (first run only)... Your speech is already being captured."
            )
        else:
            self.status_text.set(
                f"Listening. Text will appear after each short pause. "
                f"Press {self.hotkey_display.get()} again to stop."
            )

    def _stop_live_recording(self) -> None:
        if self.streaming_recorder is None:
            return

        # Hide the floating overlay as soon as the user stops — no more
        # listening is happening after this.
        self._hide_floating_indicator()

        try:
            self.streaming_recorder.stop()
        except Exception as exc:
            self._show_error(f"Could not stop live dictation: {exc}")

        # Drop our reference — recorder will still have its writer thread
        # drain any queued audio via the sentinel. The worker thread below
        # will process the last phrase and then emit "live_done".
        self.streaming_recorder = None
        if self.live_transcribe_queue is not None:
            # Null sentinel stops the transcribe worker after remaining items.
            self.live_transcribe_queue.put(None)

        # Stop animating and remove the "Listening..." indicator; the status
        # bar already communicates the "finishing last phrase" state.
        self._stop_live_indicator()

        self.toggle_button.config(text=self._record_button_text("Start"))
        self._set_window_title("Transcribing...")
        self.status_text.set("Stopped. Finishing the last phrase...")
        self._play_feedback("recording_stopped")

    def _on_live_segment_ready(self, audio_path: str) -> None:
        """Recorder writer-thread callback — pushes audio onto the queue."""
        if self.live_transcribe_queue is not None:
            self.live_transcribe_queue.put(audio_path)

    def _live_transcribe_loop(self) -> None:
        """Drains live_transcribe_queue on a worker thread."""
        while True:
            item: Optional[str]
            if self.live_transcribe_queue is None:
                return
            item = self.live_transcribe_queue.get()
            if item is None:
                self.ui_queue.put(("live_done", None))
                return
            audio_path = item
            try:
                model = self._load_model()
                # Feed Whisper the most recent 200 chars of what's already
                # been transcribed this session. This is how we reclaim some
                # of the context that batch mode has for free — the model
                # uses the prompt to disambiguate homophones, continue
                # proper nouns, and match tense/style across phrases.
                # Strip ellipses from the context too so we don't teach
                # Whisper to reproduce them.
                prior = self._clean_live_text((self.last_saved_text or "").strip())
                initial_prompt = prior[-200:] if prior else None
                segments, _ = model.transcribe(
                    audio_path,
                    language=self.language,
                    vad_filter=True,   # second-pass Silero VAD inside Whisper
                    beam_size=3,        # larger beam for better short-phrase accuracy
                    condition_on_previous_text=False,
                    initial_prompt=initial_prompt,
                )
                text = " ".join(seg.text.strip() for seg in segments).strip()
                text = self._clean_live_text(text)
                if text:
                    self.ui_queue.put(("live_phrase", text))
            except Exception as exc:
                # Live errors go to status only — modal dialogs per failed
                # segment would be hostile. The session keeps going so a
                # transient glitch doesn't end dictation.
                self.ui_queue.put(
                    ("status", f"Skipped a phrase ({exc}). Continuing to listen…")
                )
            finally:
                try:
                    Path(audio_path).unlink(missing_ok=True)
                except OSError:
                    pass

    @staticmethod
    def _clean_live_text(text: str) -> str:
        """
        Remove Whisper's tendency to sprinkle ellipses at phrase boundaries
        when it's fed short audio chunks with pauses. Collapses runs of 2+
        periods and the Unicode ellipsis (…) to nothing, and normalises the
        resulting whitespace. Single periods (real sentence endings) are
        preserved.
        """
        if not text:
            return text
        # Kill runs of 2+ periods ("..", "...", "....").
        text = re.sub(r"\.{2,}", "", text)
        # Kill the Unicode ellipsis character.
        text = text.replace("…", "")
        # Collapse any double spaces we just created.
        text = re.sub(r"\s+", " ", text).strip()
        # A stray leading " ." or " ," can happen if we stripped inside
        # punctuation; clean that up too.
        text = re.sub(r"\s+([.,!?])", r"\1", text)
        return text

    def _handle_live_phrase(self, text: str) -> None:
        is_first = self.live_phrase_count == 0
        piece = text if is_first else " " + text

        # If a "Listening..." indicator is currently at the end of the box,
        # insert the new phrase BEFORE it so the indicator stays trailing.
        indicator_range = self.transcript_box.tag_ranges("live_indicator")
        if len(indicator_range) >= 2:
            self.transcript_box.insert(indicator_range[0], piece)
        else:
            self.transcript_box.insert("end", piece)
        self.transcript_box.see("end")
        self.last_saved_text += piece
        self._append_history(text)

        if self.auto_paste_enabled and self.live_target_window_handle is not None:
            self._paste_live_phrase(piece, focus_window=is_first)
        else:
            # No target app — keep clipboard useful at least.
            self._copy_transcript_to_clipboard(self.last_saved_text)

        self.live_phrase_count += 1
        # No per-phrase beep: the keystroke into the target app is feedback
        # enough, and beeping on every commit is too noisy during continuous
        # dictation. Start/stop beeps still fire once per session.
        preview = text if len(text) <= 80 else text[:80] + "…"
        self.status_text.set(
            f'Heard: "{preview}" — keep talking or press {self.hotkey_display.get()} to stop.'
        )

    def _paste_live_phrase(self, text: str, focus_window: bool) -> None:
        if not text:
            return
        pyperclip.copy(text)
        if focus_window:
            if not self._focus_target_window(self.live_target_window_handle):
                self.status_text.set(
                    "Transcript copied — Windows blocked auto-paste. "
                    "Click your app and press Ctrl+V."
                )
                return
            # Give the target window a beat to fully receive focus before
            # we send Ctrl+V (matches the batch-mode delay for Electron apps).
            self.root.after(150, self._send_ctrl_v)
        else:
            # Target should still be focused from the first paste. Firing
            # Ctrl+V directly keeps the flow tight between phrases.
            self._send_ctrl_v()

    def _handle_live_done(self) -> None:
        self.live_worker_thread = None
        self.live_transcribe_queue = None
        self.live_target_window_handle = None
        # Safety net — if stop didn't clean up (e.g. stopped via an error path)
        # make sure the indicator animation isn't still ticking.
        self._stop_live_indicator()
        self._set_window_title("Ready")
        if self.live_phrase_count == 0:
            self.status_text.set(
                f"No speech detected. Press {self.hotkey_display.get()} to try again."
            )
        else:
            self.status_text.set(
                f"Dictation complete ({self.live_phrase_count} phrases). "
                f"Press {self.hotkey_display.get()} to dictate again."
            )

    # "🎤 Listening..." animation shown inside the transcript box while live
    # dictation is running. Implemented with a Text tag so phrases can be
    # inserted before the indicator and leave it trailing.
    _LIVE_INDICATOR_PREFIX = "🎤  Listening"
    _LIVE_INDICATOR_TAG = "live_indicator"
    _LIVE_INDICATOR_TICK_MS = 500


    def _start_live_indicator(self) -> None:
        try:
            self.transcript_box.tag_configure(
                self._LIVE_INDICATOR_TAG,
                foreground="#94a3b8",
                font=("Consolas", 11, "italic"),
            )
        except tk.TclError:
            pass
        self._live_indicator_phase = 0
        self._insert_or_update_live_indicator()
        self._live_indicator_after_id = self.root.after(
            self._LIVE_INDICATOR_TICK_MS, self._tick_live_indicator
        )

    def _stop_live_indicator(self) -> None:
        after_id = getattr(self, "_live_indicator_after_id", None)
        if after_id is not None:
            try:
                self.root.after_cancel(after_id)
            except tk.TclError:
                pass
            self._live_indicator_after_id = None
        ranges = self.transcript_box.tag_ranges(self._LIVE_INDICATOR_TAG)
        if len(ranges) >= 2:
            self.transcript_box.delete(ranges[0], ranges[-1])

    def _tick_live_indicator(self) -> None:
        # If the recorder was torn down between ticks, bail out quietly.
        if self.streaming_recorder is None or not self.streaming_recorder.is_recording:
            self._live_indicator_after_id = None
            return
        self._live_indicator_phase = (self._live_indicator_phase + 1) % 4
        self._insert_or_update_live_indicator()
        self._live_indicator_after_id = self.root.after(
            self._LIVE_INDICATOR_TICK_MS, self._tick_live_indicator
        )

    def _insert_or_update_live_indicator(self) -> None:
        dots = "." * (self._live_indicator_phase + 1)  # 1..4 dots
        # Two-space left margin if there's already transcript text so the
        # indicator doesn't sit flush against the last word.
        box_has_text = self.transcript_box.get("1.0", "end-1c") != ""
        existing = self.transcript_box.tag_ranges(self._LIVE_INDICATOR_TAG)
        if len(existing) >= 2:
            # Already in the box — check if there's non-indicator text before
            # us, so we only add the spacer when needed.
            leading_text = self.transcript_box.get("1.0", existing[0])
            needs_spacer = bool(leading_text)
            self.transcript_box.delete(existing[0], existing[-1])
        else:
            needs_spacer = box_has_text

        spacer = "  " if needs_spacer else ""
        new_text = f"{spacer}{self._LIVE_INDICATOR_PREFIX}{dots}"
        self.transcript_box.insert("end", new_text, self._LIVE_INDICATOR_TAG)
        self.transcript_box.see("end")

    # Floating always-on-top overlay shown while live dictation is running.
    # It's a borderless Toplevel parked at the top-right of the primary
    # monitor, with an animated "🎤 Listening..." label. Fully non-invasive:
    # it never touches the target app's text — just sits on top so the user
    # can see at a glance that VoiceTyper is listening no matter where they
    # are. Click-through is set via Win32 so the overlay never steals focus.

    _FLOATING_INDICATOR_TICK_MS = 500

    # Floating always-on-top overlay shown while live dictation is running.
    # Frame-based layout — SetWindowRgn clips the corners to a rounded pill
    # shape. Corners aren't perfectly antialiased (that would require the
    # UpdateLayeredWindow + Pillow-premult dance), but this version is
    # reliable: the overlay is visible, typing keeps working, and we don't
    # accidentally steal foreground focus from the target app.

    def _show_floating_indicator(self, prefix_text: str = "🎤  Listening") -> None:
        if self._floating_indicator is not None:
            # Already showing — just refresh the prefix text (useful if
            # batch mode transitions Listening → Transcribing).
            self._set_indicator_prefix(prefix_text)
            return

        bg = "#0b1221"
        fg = "#f8fafc"
        muted = "#94a3b8"
        accent = "#ef4444"  # steady red "recording" dot
        border = "#1f2a44"

        win = tk.Toplevel(self.root)
        win.title("VoiceTyper listening")
        win.configure(bg=bg)
        win.overrideredirect(True)
        try:
            win.attributes("-topmost", True)
        except tk.TclError:
            pass

        screen_w = self.root.winfo_screenwidth()
        width, height = 210, 38
        margin_x, margin_y = 24, 44
        x = max(0, screen_w - width - margin_x)
        y = margin_y
        win.geometry(f"{width}x{height}+{x}+{y}")

        border_frame = tk.Frame(win, bg=border)
        border_frame.place(x=0, y=0, width=width, height=height)
        inner = tk.Frame(border_frame, bg=bg)
        inner.place(x=1, y=1, width=width - 2, height=height - 2)

        content = tk.Frame(inner, bg=bg)
        content.place(relx=0.5, rely=0.5, anchor="center")

        dot_label = tk.Label(
            content,
            text="●",
            bg=bg,
            fg=accent,
            font=("Segoe UI", 9, "bold"),
        )
        dot_label.pack(side="left", padx=(0, 7))

        prefix_label = tk.Label(
            content,
            text=prefix_text,
            bg=bg,
            fg=fg,
            font=("Segoe UI", 9, "bold"),
        )
        prefix_label.pack(side="left")

        dots_label = tk.Label(
            content,
            text="....",
            bg=bg,
            fg=fg,
            font=("Segoe UI", 9, "bold"),
            width=4,
            anchor="w",
        )
        dots_label.pack(side="left")

        lang_label = tk.Label(
            content,
            text=self._language_badge_text(),
            bg=bg,
            fg=muted,
            font=("Segoe UI", 9, "bold"),
        )
        lang_label.pack(side="left", padx=(6, 0))

        self._floating_indicator = win
        self._floating_indicator_text_label = dots_label
        self._floating_indicator_dot_label = dot_label
        self._floating_indicator_prefix_label = prefix_label
        self._floating_indicator_lang_label = lang_label
        self._floating_indicator_phase = 0

        win.update_idletasks()

        # Prevent the overlay from ever stealing foreground focus so that
        # SetForegroundWindow in the paste path keeps working.
        self._make_overlay_nonactivating(win)

        try:
            win.attributes("-alpha", 0.62)
        except tk.TclError:
            pass

        self._tick_floating_indicator()

    def _language_badge_text(self) -> str:
        """Short language indicator shown on the floating overlay."""
        if self.language == "zh":
            return "· 中"
        return "· EN"

    def _set_indicator_prefix(self, text: str) -> None:
        """Change the main label on the floating overlay (e.g. Listening → Transcribing)."""
        label = self._floating_indicator_prefix_label
        if label is None:
            return
        try:
            label.config(text=text)
        except tk.TclError:
            pass

    def _refresh_indicator_language(self) -> None:
        """If the overlay is up, sync the language badge to the current setting."""
        label = self._floating_indicator_lang_label
        if label is None:
            return
        try:
            label.config(text=self._language_badge_text())
        except tk.TclError:
            pass

    def _make_overlay_nonactivating(self, win: tk.Toplevel) -> None:
        """Windows: add WS_EX_NOACTIVATE + WS_EX_TOOLWINDOW so the overlay
        never takes keyboard focus or alters the foreground window."""
        if sys.platform != "win32":
            return
        try:
            hwnd = int(win.winfo_id())
            GWL_EXSTYLE = -20
            WS_EX_NOACTIVATE = 0x08000000
            WS_EX_TOOLWINDOW = 0x00000080
            user32 = ctypes.windll.user32
            user32.GetWindowLongW.restype = ctypes.c_long
            user32.GetWindowLongW.argtypes = [wintypes.HWND, ctypes.c_int]
            user32.SetWindowLongW.restype = ctypes.c_long
            user32.SetWindowLongW.argtypes = [
                wintypes.HWND, ctypes.c_int, ctypes.c_long,
            ]
            ex = user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
            user32.SetWindowLongW(
                hwnd, GWL_EXSTYLE, ex | WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW,
            )
        except Exception:
            pass

    def _apply_rounded_corners(self, win: tk.Toplevel, width: int, height: int,
                                radius: int = 18) -> None:
        """Clip the overlay to a rounded rectangle via SetWindowRgn.

        Correct ctypes argtypes matter on 64-bit Windows: HRGN is a pointer,
        so the return type of CreateRoundRectRgn must be c_void_p or the
        handle gets truncated to 32 bits and SetWindowRgn silently fails."""
        if sys.platform != "win32":
            return
        try:
            hwnd = int(win.winfo_id())
            gdi32 = ctypes.windll.gdi32
            user32 = ctypes.windll.user32
            gdi32.CreateRoundRectRgn.argtypes = [ctypes.c_int] * 6
            gdi32.CreateRoundRectRgn.restype = ctypes.c_void_p
            user32.SetWindowRgn.argtypes = [
                wintypes.HWND, ctypes.c_void_p, wintypes.BOOL,
            ]
            user32.SetWindowRgn.restype = ctypes.c_int
            # CreateRoundRectRgn takes exclusive (x2, y2) and the ellipse
            # dimensions for the corner arc. For a full pill, ellipse diameter
            # must equal window height — exactly radius*2.
            region = gdi32.CreateRoundRectRgn(
                0, 0, width + 1, height + 1, radius * 2, radius * 2,
            )
            if region:
                user32.SetWindowRgn(hwnd, region, True)
        except Exception:
            pass

    def _hide_floating_indicator(self) -> None:
        after_id = self._floating_indicator_after_id
        if after_id is not None:
            try:
                self.root.after_cancel(after_id)
            except tk.TclError:
                pass
            self._floating_indicator_after_id = None
        if self._floating_indicator is not None:
            try:
                self._floating_indicator.destroy()
            except tk.TclError:
                pass
            self._floating_indicator = None
            self._floating_indicator_text_label = None
            self._floating_indicator_dot_label = None
            self._floating_indicator_prefix_label = None
            self._floating_indicator_lang_label = None

    def _tick_floating_indicator(self) -> None:
        # Keep animating as long as the overlay exists — works for both
        # live mode and batch "Transcribing" state. Explicit hide tears it
        # down; there's no other exit condition.
        if self._floating_indicator is None:
            self._floating_indicator_after_id = None
            return

        self._floating_indicator_phase = (self._floating_indicator_phase + 1) % 4
        dots = "." * (self._floating_indicator_phase + 1)
        try:
            if self._floating_indicator_text_label is not None:
                self._floating_indicator_text_label.config(text=dots)
        except tk.TclError:
            self._floating_indicator_after_id = None
            return

        self._floating_indicator_after_id = self.root.after(
            self._FLOATING_INDICATOR_TICK_MS, self._tick_floating_indicator
        )

    def _on_mode_change(self) -> None:
        new_mode = self.mode_var.get()
        if new_mode == self.mode:
            return

        # Don't let the mode change mid-recording — it would leave state dangling.
        if self.recorder.is_recording or self.is_transcribing or (
            self.streaming_recorder is not None and self.streaming_recorder.is_recording
        ):
            self.status_text.set(
                "Stop the current recording or transcription before switching mode."
            )
            # Revert the radio button to reflect the actual mode.
            self.mode_var.set(self.mode)
            return

        self.mode = new_mode
        self._save_hotkey_setting()
        self.toggle_button.config(text=self._record_button_text("Start"))
        self.status_text.set(self._initial_status_text())

        # In live mode, kick off a background model preload so the first
        # phrase of the first dictation session doesn't stall on model load.
        if self.mode == "live" and (
            self.model is None or self.current_model_lang != self.language
        ):
            threading.Thread(target=self._preload_model, daemon=True).start()

    def _preload_model(self) -> None:
        try:
            self._load_model()
        except Exception as exc:
            self.ui_queue.put(("status", f"Model preload failed: {exc}"))

    def _load_model(self) -> WhisperModel:
        with self.model_lock:
            target_lang = self.language
            if self.model is None or self.current_model_lang != target_lang:
                self.model = None
                model_name = "base" if target_lang == "zh" else "base.en"
                self.ui_queue.put(
                    (
                        "status",
                        "Loading the speech model for the first time. This can take a minute.",
                    )
                )
                self.model = WhisperModel(model_name, device="cpu", compute_type="int8")
                self.current_model_lang = target_lang
            return self.model

    def _transcribe_audio(self, audio_path: str) -> None:
        try:
            model = self._load_model()
            segments, _ = model.transcribe(
                audio_path,
                language=self.language,
                vad_filter=True,
                beam_size=5,
            )
            transcript = " ".join(segment.text.strip() for segment in segments).strip()
            if not transcript:
                transcript = "[No speech detected]"
            self.ui_queue.put(("transcript", transcript))
        except Exception as exc:
            self.ui_queue.put(("error", f"Transcription failed: {exc}"))
        finally:
            self.is_transcribing = False
            try:
                Path(audio_path).unlink(missing_ok=True)
            except OSError:
                pass

    def _show_transcript(self, transcript: str) -> None:
        self.transcript_box.delete("1.0", "end")
        self.transcript_box.insert("1.0", transcript)
        self.last_saved_text = transcript
        self._copy_transcript_to_clipboard(transcript)
        self._append_history(transcript)
        if self.auto_paste_enabled:
            if self.target_window_handle is None:
                self.status_text.set(
                    "Transcript ready and copied to clipboard. Auto-paste works best when you start from the target app with your hotkey."
                )
            else:
                self.status_text.set("Transcript ready. Pasting it into your active app...")
                self.root.after(150, lambda text=transcript: self._paste_into_target_app(text))
        else:
            self.status_text.set("Transcript ready and copied to your clipboard.")
        self._set_window_title("Ready")
        self._play_feedback("transcript_ready")
        # Batch mode shows the overlay during recording + transcription;
        # tear it down now that the transcript has landed.
        self._hide_floating_indicator()

    def _append_history(self, transcript: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        entry = f"[{timestamp}] {transcript}\n\n"
        self.history_box.config(state="normal")
        self.history_box.insert("1.0", entry)
        self.history_box.config(state="disabled")

    def copy_transcript(self) -> None:
        text = self.transcript_box.get("1.0", "end").strip()
        if not text:
            self.status_text.set("There is no transcript to copy yet.")
            return
        self._copy_transcript_to_clipboard(text)
        self.status_text.set("Transcript copied to clipboard.")

    def _copy_transcript_to_clipboard(self, text: str) -> None:
        pyperclip.copy(text)

    def _paste_into_target_app(self, text: str) -> None:
        if not text or self.target_window_handle is None:
            return

        if not self._focus_target_window(self.target_window_handle):
            self.status_text.set("Transcript copied, but Windows blocked auto-paste. Click the target box and press Ctrl+V.")
            return

        # Delay the actual keystrokes so Electron/Chromium apps (like Cursor/VS Code)
        # finish their focus transition before receiving Ctrl+V — without this delay
        # those apps can process the paste event twice.
        self.root.after(200, self._send_ctrl_v)
        self.status_text.set("Transcript pasted into your active app and copied to the clipboard.")

    def _focus_target_window(self, window_handle: int) -> bool:
        try:
            user32 = ctypes.windll.user32
            user32.ShowWindow(window_handle, 5)
            return bool(user32.SetForegroundWindow(window_handle))
        except Exception:
            return False

    def _send_ctrl_v(self) -> None:
        self.keyboard_controller.press(keyboard.Key.ctrl)
        self.keyboard_controller.press("v")
        self.keyboard_controller.release("v")
        self.keyboard_controller.release(keyboard.Key.ctrl)

    def clear_transcript(self) -> None:
        self.transcript_box.delete("1.0", "end")
        self.status_text.set(
            f"Transcript cleared. Press {self.hotkey_display.get()} to record again."
        )

    def _show_error(self, message: str) -> None:
        self.status_text.set(message)
        self.is_transcribing = False
        self.toggle_button.config(text=self._record_button_text("Start"))
        self._set_window_title("Error")
        self._play_feedback("error")
        self._hide_floating_indicator()
        messagebox.showerror(APP_NAME, message)

    def _record_button_text(self, action: str) -> str:
        verb = "Dictating" if self.mode == "live" else "Recording"
        return f"{action} {verb} ({self.hotkey_display.get()})"

    def _initial_status_text(self) -> str:
        if self.mode == "live":
            return (
                f"Live dictation ready. Press {self.hotkey_display.get()} and start speaking — "
                f"text will appear phrase by phrase."
            )
        return f"Ready. Press {self.hotkey_display.get()} or click Start Recording."

    def apply_recommended_hotkey(self, hotkey_tokens: tuple[str, ...]) -> None:
        if self.recorder.is_recording or self.is_transcribing:
            self.status_text.set("Stop the current recording or transcription before changing the hotkey.")
            return
        if self.is_capturing_hotkey:
            self.cancel_hotkey_capture()

        self._set_hotkey(hotkey_tokens)
        self.status_text.set(f"Recommended hotkey applied: {self.hotkey_display.get()}.")

    def start_hotkey_capture(self) -> None:
        if self.recorder.is_recording or self.is_transcribing:
            self.status_text.set("Stop the current recording or transcription before changing the hotkey.")
            return
        if self.is_capturing_hotkey:
            return

        self.is_capturing_hotkey = True
        if self.hotkey_listener is not None:
            self.hotkey_listener.stop()
            self.hotkey_listener = None
        self.pressed_keys.clear()
        self.hotkey_active = False
        self.capture_pressed_keys.clear()
        self.capture_modifiers.clear()
        self.capture_main_key = None
        self.capture_preview_text.set("Waiting for your shortcut...")

        dialog = tk.Toplevel(self.root)
        dialog.title("Choose Recording Hotkey")
        dialog.geometry("420x180")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.protocol("WM_DELETE_WINDOW", self.cancel_hotkey_capture)

        frame = ttk.Frame(dialog, padding=16)
        frame.pack(fill="both", expand=True)

        ttk.Label(
            frame,
            text="Press the exact key or key combination you want to use.",
            wraplength=380,
        ).pack(anchor="w")
        ttk.Label(
            frame,
            text="Examples: F6, Numpad 5, Ctrl+Alt+R, Ctrl+Shift+Space",
            wraplength=380,
        ).pack(anchor="w", pady=(8, 12))
        ttk.Label(frame, text="Detected shortcut:").pack(anchor="w")
        ttk.Label(
            frame,
            textvariable=self.capture_preview_text,
            font=("Segoe UI", 12, "bold"),
        ).pack(anchor="w", pady=(4, 16))
        ttk.Button(frame, text="Cancel", command=self.cancel_hotkey_capture).pack(anchor="e")

        self.capture_dialog = dialog
        self.capture_listener = keyboard.Listener(
            on_press=self._on_capture_key_press,
            on_release=self._on_capture_key_release,
        )
        self.capture_listener.start()

    def cancel_hotkey_capture(self) -> None:
        if self.capture_listener is not None:
            self.capture_listener.stop()
            self.capture_listener = None
        if self.capture_dialog is not None:
            self.capture_dialog.grab_release()
            self.capture_dialog.destroy()
            self.capture_dialog = None
        self.is_capturing_hotkey = False
        self.capture_pressed_keys.clear()
        self.capture_modifiers.clear()
        self.capture_main_key = None
        self.capture_preview_text.set("")
        self._start_hotkey_listener()
        self.status_text.set(f"Hotkey unchanged. Still using {self.hotkey_display.get()}.")

    def _on_capture_key_press(self, key):
        token = self._key_to_token(key)
        if token is None:
            return

        self.capture_pressed_keys.add(token)
        if token in self._modifier_tokens():
            self.capture_modifiers.add(token)
        elif self.capture_main_key is None:
            self.capture_main_key = token

        preview_tokens = self._capture_tokens_in_order(self.capture_modifiers, self.capture_main_key)
        if preview_tokens:
            self.ui_queue.put(("capture_preview", self._format_hotkey_for_display(preview_tokens)))

    def _on_capture_key_release(self, key):
        token = self._key_to_token(key)
        if token is None:
            return

        self.capture_pressed_keys.discard(token)
        if token in self._modifier_tokens():
            self.capture_modifiers.discard(token)
            return

        if self.capture_main_key is not None and token == self.capture_main_key:
            final_tokens = self._capture_tokens_in_order(self.capture_modifiers, self.capture_main_key)
            self.ui_queue.put(("capture_complete", final_tokens))
            return False

    def _finish_hotkey_capture(self, hotkey_tokens: tuple[str, ...]) -> None:
        if self.capture_listener is not None:
            self.capture_listener.stop()
            self.capture_listener = None
        if self.capture_dialog is not None:
            self.capture_dialog.grab_release()
            self.capture_dialog.destroy()
            self.capture_dialog = None

        self.is_capturing_hotkey = False
        self.capture_pressed_keys.clear()
        self.capture_modifiers.clear()
        self.capture_main_key = None
        self.capture_preview_text.set("")

        if not hotkey_tokens:
            self._start_hotkey_listener()
            self._show_error("Could not detect a valid shortcut. Please try again.")
            return

        self._set_hotkey(hotkey_tokens)
        self.status_text.set(f"Hotkey updated to {self.hotkey_display.get()}.")

    def _set_hotkey(self, hotkey_tokens: tuple[str, ...]) -> None:
        self.hotkey_tokens = hotkey_tokens
        self.hotkey_display.set(self._format_hotkey_for_display(self.hotkey_tokens))
        self._save_hotkey_setting()
        self.toggle_button.config(text=self._record_button_text("Start"))
        self._start_hotkey_listener()
        self._set_window_title("Ready")

    def _set_window_title(self, state: str) -> None:
        if state == "Recording...":
            title = f"{APP_NAME} - {state}"
        elif state == "Transcribing...":
            title = f"{APP_NAME} - {state}"
        elif state == "Error":
            title = f"{APP_NAME} - Error"
        else:
            title = APP_NAME
        self.root.title(title)
        if self.tray_icon is not None:
            self.tray_icon.title = title

    def _play_feedback(self, event_name: str) -> None:
        if not self.sound_feedback_enabled:
            return

        threading.Thread(
            target=self._play_feedback_sync,
            args=(event_name,),
            daemon=True,
        ).start()

    def _play_feedback_sync(self, event_name: str) -> None:
        try:
            if event_name == "recording_started":
                winsound.Beep(880, 120)
            elif event_name == "recording_stopped":
                winsound.Beep(660, 90)
                winsound.Beep(520, 90)
            elif event_name == "transcript_ready":
                winsound.Beep(740, 90)
                winsound.Beep(988, 120)
            elif event_name == "error":
                winsound.MessageBeep(winsound.MB_ICONHAND)
        except RuntimeError:
            pass

    def _on_language_change(self, _event=None) -> None:
        selected_display = self.lang_display_var.get()
        new_lang = self._lang_options.get(selected_display, "en")
        if new_lang == self.language:
            return
        self.language = new_lang
        with self.model_lock:
            self.model = None
            self.current_model_lang = None
        self._save_hotkey_setting()
        lang_name = self._lang_display_reverse.get(new_lang, new_lang)
        self.status_text.set(
            f"Language set to {lang_name}. The model will reload on next recording."
        )
        # If the floating overlay is currently visible, sync its lang badge.
        self._refresh_indicator_language()

    def _load_language_setting(self) -> str:
        if not self.settings_path.exists():
            return "en"
        try:
            data = json.loads(self.settings_path.read_text(encoding="utf-8"))
            lang = data.get("language", "en")
            if lang in ("en", "zh"):
                return lang
        except (json.JSONDecodeError, OSError):
            pass
        return "en"

    def _load_mode_setting(self) -> str:
        if not self.settings_path.exists():
            return "batch"
        try:
            data = json.loads(self.settings_path.read_text(encoding="utf-8"))
            mode = data.get("mode", "batch")
            if mode in ("batch", "live"):
                return mode
        except (json.JSONDecodeError, OSError):
            pass
        return "batch"

    def _load_hotkey_setting(self) -> tuple[str, ...]:
        if not self.settings_path.exists():
            return ("f8",)

        try:
            data = json.loads(self.settings_path.read_text(encoding="utf-8"))
            saved_tokens = data.get("hotkey_tokens")
            if isinstance(saved_tokens, list):
                normalized_tokens = tuple(str(token) for token in saved_tokens if str(token).strip())
                if normalized_tokens:
                    return normalized_tokens

            legacy_hotkey = data.get("hotkey")
            if isinstance(legacy_hotkey, str):
                legacy_tokens = self._parse_legacy_hotkey(legacy_hotkey)
                if legacy_tokens:
                    return legacy_tokens
        except (json.JSONDecodeError, OSError):
            pass
        return ("f8",)

    def _save_hotkey_setting(self) -> None:
        settings = {
            "hotkey_tokens": list(self.hotkey_tokens),
            "language": self.language,
            "mode": self.mode,
        }
        self.settings_path.write_text(json.dumps(settings, indent=2), encoding="utf-8")

    def _parse_legacy_hotkey(self, hotkey: str) -> tuple[str, ...]:
        legacy_map = {
            "<ctrl>": "ctrl",
            "<alt>": "alt",
            "<shift>": "shift",
            "<cmd>": "win",
            "<space>": "space",
            "<enter>": "enter",
            "<tab>": "tab",
            "<esc>": "esc",
            "<backspace>": "backspace",
            "<delete>": "delete",
            "<up>": "up",
            "<down>": "down",
            "<left>": "left",
            "<right>": "right",
            "<home>": "home",
            "<end>": "end",
            "<page_up>": "pageup",
            "<page_down>": "pagedown",
            "<insert>": "insert",
        }
        tokens: list[str] = []
        for part in hotkey.split("+"):
            part = part.strip().lower()
            if not part:
                continue
            if part in legacy_map:
                tokens.append(legacy_map[part])
            elif part.startswith("<") and part.endswith(">"):
                tokens.append(part[1:-1])
            else:
                tokens.append(part)
        return tuple(tokens)

    def _format_hotkey_for_display(self, hotkey_tokens: tuple[str, ...]) -> str:
        pretty_map = {
            "ctrl": "Ctrl",
            "alt": "Alt",
            "shift": "Shift",
            "win": "Win",
            "space": "Space",
            "enter": "Enter",
            "tab": "Tab",
            "esc": "Esc",
            "backspace": "Backspace",
            "delete": "Delete",
            "up": "Up",
            "down": "Down",
            "left": "Left",
            "right": "Right",
            "home": "Home",
            "end": "End",
            "pageup": "PageUp",
            "pagedown": "PageDown",
            "insert": "Insert",
            "capslock": "CapsLock",
            "numlock": "NumLock",
            "numpaddecimal": "Numpad Decimal",
            "numpaddivide": "Numpad Divide",
            "numpadmultiply": "Numpad Multiply",
            "numpadsubtract": "Numpad Subtract",
            "numpadadd": "Numpad Add",
        }

        display_parts = []
        for token in hotkey_tokens:
            if token in pretty_map:
                display_parts.append(pretty_map[token])
            elif token.startswith("f") and token[1:].isdigit():
                display_parts.append(token.upper())
            elif token.startswith("numpad") and token[6:].isdigit():
                display_parts.append(f"Numpad {token[6:]}")
            elif len(token) == 1:
                display_parts.append(token.upper())
            else:
                display_parts.append(token.title())
        return "+".join(display_parts)

    def _capture_tokens_in_order(
        self,
        modifiers: set[str],
        main_key: Optional[str],
    ) -> tuple[str, ...]:
        ordered_tokens = [token for token in self._modifier_tokens() if token in modifiers]
        if main_key:
            ordered_tokens.append(main_key)
        return tuple(ordered_tokens)

    def _modifier_tokens(self) -> tuple[str, ...]:
        return ("ctrl", "alt", "shift", "win")

    def _key_to_token(self, key) -> Optional[str]:
        modifier_key_map = {
            keyboard.Key.ctrl: "ctrl",
            keyboard.Key.ctrl_l: "ctrl",
            keyboard.Key.ctrl_r: "ctrl",
            keyboard.Key.alt: "alt",
            keyboard.Key.alt_l: "alt",
            keyboard.Key.alt_r: "alt",
            keyboard.Key.alt_gr: "alt",
            keyboard.Key.shift: "shift",
            keyboard.Key.shift_l: "shift",
            keyboard.Key.shift_r: "shift",
            keyboard.Key.cmd: "win",
            keyboard.Key.cmd_l: "win",
            keyboard.Key.cmd_r: "win",
            keyboard.Key.space: "space",
            keyboard.Key.enter: "enter",
            keyboard.Key.tab: "tab",
            keyboard.Key.esc: "esc",
            keyboard.Key.backspace: "backspace",
            keyboard.Key.delete: "delete",
            keyboard.Key.up: "up",
            keyboard.Key.down: "down",
            keyboard.Key.left: "left",
            keyboard.Key.right: "right",
            keyboard.Key.home: "home",
            keyboard.Key.end: "end",
            keyboard.Key.page_up: "pageup",
            keyboard.Key.page_down: "pagedown",
            keyboard.Key.insert: "insert",
            keyboard.Key.caps_lock: "capslock",
            keyboard.Key.num_lock: "numlock",
            keyboard.Key.f1: "f1",
            keyboard.Key.f2: "f2",
            keyboard.Key.f3: "f3",
            keyboard.Key.f4: "f4",
            keyboard.Key.f5: "f5",
            keyboard.Key.f6: "f6",
            keyboard.Key.f7: "f7",
            keyboard.Key.f8: "f8",
            keyboard.Key.f9: "f9",
            keyboard.Key.f10: "f10",
            keyboard.Key.f11: "f11",
            keyboard.Key.f12: "f12",
        }
        if key in modifier_key_map:
            return modifier_key_map[key]

        if isinstance(key, keyboard.KeyCode):
            vk = getattr(key, "vk", None)
            numpad_map = {
                96: "numpad0",
                97: "numpad1",
                98: "numpad2",
                99: "numpad3",
                100: "numpad4",
                101: "numpad5",
                102: "numpad6",
                103: "numpad7",
                104: "numpad8",
                105: "numpad9",
                106: "numpadmultiply",
                107: "numpadadd",
                109: "numpadsubtract",
                110: "numpaddecimal",
                111: "numpaddivide",
            }
            if vk in numpad_map:
                return numpad_map[vk]

            if key.char and key.char.isprintable():
                return key.char.lower()

            if vk is not None and 48 <= vk <= 57:
                return chr(vk).lower()
            if vk is not None and 65 <= vk <= 90:
                return chr(vk).lower()

        return None

    # ------------------------------------------------------------------ tray ---

    def _make_tray_image(self) -> PilImage.Image:
        """Return a PIL image for the tray icon (uses app icon if available)."""
        if self.icon_path.exists():
            try:
                img = PilImage.open(str(self.icon_path)).convert("RGBA")
                img = img.resize((64, 64), PilImage.LANCZOS)
                return img
            except Exception:
                pass
        # Fallback: plain coloured square
        img = PilImage.new("RGBA", (64, 64), (37, 99, 235, 255))
        return img

    def _setup_tray(self) -> None:
        img = self._make_tray_image()
        menu = pystray.Menu(
            pystray.MenuItem("Show", self._tray_show, default=True),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Quit", self._tray_quit),
        )
        self.tray_icon = pystray.Icon(APP_ID, img, APP_NAME, menu)
        t = threading.Thread(target=self.tray_icon.run, daemon=True)
        t.start()

    def _hide_to_tray(self) -> None:
        """Hide the main window; the tray icon stays alive."""
        self.root.withdraw()
        if self.tray_icon is not None:
            self.tray_icon.notify(
                "Still running in the background.\nPress your hotkey to record.",
                APP_NAME,
            )

    def _tray_show(self, icon=None, item=None) -> None:
        """Restore / raise the main window from tray."""
        self.root.after(0, self._show_window)

    def _show_window(self) -> None:
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()

    def _tray_quit(self, icon=None, item=None) -> None:
        """Fully exit the application from tray."""
        self.root.after(0, self.on_close)

    # ---------------------------------------------------------------- close ---

    def on_close(self) -> None:
        if self.recorder.is_recording:
            try:
                self.recorder.stop()
            except Exception:
                pass
        if self.streaming_recorder is not None and self.streaming_recorder.is_recording:
            try:
                self.streaming_recorder.stop()
            except Exception:
                pass
        # Signal the live worker (if any) to exit.
        if self.live_transcribe_queue is not None:
            try:
                self.live_transcribe_queue.put(None)
            except Exception:
                pass
        self._hide_floating_indicator()
        if self.hotkey_listener is not None:
            self.hotkey_listener.stop()
        if self.tray_icon is not None:
            self.tray_icon.stop()
        self.scroll_canvas.unbind_all("<MouseWheel>")
        self.root.destroy()


def main() -> None:
    # Only one instance: if another is running, bring it to front and exit.
    _single_instance_mutex = ensure_single_instance()

    root = tk.Tk()
    root.style = ttk.Style()
    root.style.theme_use("clam")
    app = SpeakingPracticeApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
