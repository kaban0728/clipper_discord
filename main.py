import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
import os
import subprocess
import json
import threading
import sys
import time
import tempfile
import urllib.request
import zipfile
import shutil
import ctypes


def ensure_packages():
    """不足しているPythonパッケージを自動インストールする。"""
    if getattr(sys, 'frozen', False):
        return  # exe実行時はバンドル済みなのでスキップ

    missing = []
    try:
        import cv2  # noqa: F401
    except ImportError:
        missing.append('opencv-python')
    try:
        from PIL import Image  # noqa: F401
    except ImportError:
        missing.append('Pillow')
    try:
        import customtkinter  # noqa: F401
    except ImportError:
        missing.append('customtkinter')

    if not missing:
        return

    msg = (
        f"以下のパッケージが不足しています:\n"
        f"{', '.join(missing)}\n\n"
        f"自動インストールしますか？"
    )
    temp_root = tk.Tk()
    temp_root.withdraw()
    answer = messagebox.askyesno("パッケージ不足", msg, parent=temp_root)
    if not answer:
        messagebox.showwarning(
            "警告", "必要なパッケージがないため終了します。", parent=temp_root
        )
        temp_root.destroy()
        sys.exit(1)

    try:
        for pkg in missing:
            subprocess.check_call(
                [sys.executable, '-m', 'pip', 'install', pkg]
            )
        messagebox.showinfo(
            "完了",
            "パッケージのインストールが完了しました。\nアプリを再起動します。",
            parent=temp_root,
        )
        temp_root.destroy()
        os.execv(sys.executable, [sys.executable] + sys.argv)
    except Exception as e:
        messagebox.showerror(
            "エラー",
            f"インストールに失敗しました:\n{e}\n\n"
            f"手動で以下を実行してください:\n"
            f"pip install {' '.join(missing)}",
            parent=temp_root,
        )
        temp_root.destroy()
        sys.exit(1)


ensure_packages()

import cv2
from PIL import Image, ImageTk

# --- ユーティリティ関数 ---

def get_base_path():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

def get_tool_path(tool_name):
    if os.name == 'nt' and not tool_name.endswith('.exe'):
        tool_name += '.exe'
    local_path = os.path.join(get_base_path(), tool_name)
    if os.path.exists(local_path):
        return local_path
    return shutil.which(tool_name)

def format_time(seconds):
    h = int(seconds // 3600)
    m = int((seconds % 3600) // 60)
    s = seconds % 60
    return f"{h:02d}:{m:02d}:{s:05.2f}"


# --- Windows MCI オーディオプレイヤー ---

class AudioPlayer:
    """Windows MCI を使ったシンプルな音声プレイヤー。"""
    _counter = 0

    def __init__(self):
        AudioPlayer._counter += 1
        self.alias = f"clip_audio_{AudioPlayer._counter}"
        self.loaded = False

    def load(self, filepath):
        self.close()
        filepath = os.path.abspath(filepath).replace('/', '\\')
        err = self._mci(f'open "{filepath}" type waveaudio alias {self.alias}')
        self.loaded = (err == 0)
        return self.loaded

    def play_from(self, seconds):
        if self.loaded:
            self._mci(f'seek {self.alias} to {int(seconds * 1000)}')
            self._mci(f'play {self.alias}')

    def stop(self):
        if self.loaded:
            self._mci(f'stop {self.alias}')

    def close(self):
        if self.loaded:
            self._mci(f'stop {self.alias}')
            self._mci(f'close {self.alias}')
            self.loaded = False

    @staticmethod
    def _mci(cmd):
        buf = ctypes.create_unicode_buffer(255)
        return ctypes.windll.winmm.mciSendStringW(cmd, buf, 254, 0)


# --- カスタム レンジスライダー ---

class RangeSlider(tk.Canvas):
    """1本のバー上に始点・終点・再生位置の3つのハンドルを持つスライダー。"""

    def __init__(self, parent, min_val=0, max_val=100, width=620, height=50,
                 command=None):
        super().__init__(parent, width=width, height=height, highlightthickness=0)
        self.min_val = float(min_val)
        self.max_val = float(max_val)
        self.start_val = float(min_val)
        self.end_val = float(max_val)
        self.pos_val = float(min_val)  # 再生位置
        self.cw = width
        self.ch = height
        self.command = command  # callback(handle: 'start'|'end'|'pos', value)
        self.pad = 14
        self.hr = 10   # start/end handle radius
        self.pr = 7    # pos handle radius
        self.th = 6    # track height
        self.dragging = None
        self._draw()
        self.bind('<ButtonPress-1>', self._press)
        self.bind('<B1-Motion>', self._drag)
        self.bind('<ButtonRelease-1>', self._release)

    def _v2x(self, v):
        r = self.max_val - self.min_val
        return self.pad + (v - self.min_val) / max(r, 0.001) * (self.cw - 2 * self.pad)

    def _x2v(self, x):
        r = self.cw - 2 * self.pad
        v = self.min_val + (x - self.pad) / max(r, 1) * (self.max_val - self.min_val)
        return max(self.min_val, min(v, self.max_val))

    def _draw(self):
        self.delete('all')
        cy = self.ch // 2
        # トラック背景
        self.create_rectangle(self.pad, cy - self.th//2,
                              self.cw - self.pad, cy + self.th//2,
                              fill='#555555', outline='')
        sx, ex = self._v2x(self.start_val), self._v2x(self.end_val)
        # 選択範囲
        self.create_rectangle(sx, cy - self.th//2, ex, cy + self.th//2,
                              fill='#1f6aa5', outline='')
        # 再生位置ハンドル（白丸 + 青枠）
        px = self._v2x(self.pos_val)
        self.create_line(px, cy - 16, px, cy + 16, fill='#555', width=1)
        self.create_oval(px - self.pr, cy - self.pr, px + self.pr, cy + self.pr,
                         fill='white', outline='#007bff', width=2)
        # 始点ハンドル（緑）
        self.create_oval(sx - self.hr, cy - self.hr, sx + self.hr, cy + self.hr,
                         fill='#28a745', outline='white', width=2)
        # 終点ハンドル（赤）
        self.create_oval(ex - self.hr, cy - self.hr, ex + self.hr, cy + self.hr,
                         fill='#dc3545', outline='white', width=2)

    def _press(self, event):
        sx = self._v2x(self.start_val)
        ex = self._v2x(self.end_val)
        px = self._v2x(self.pos_val)
        ds, de, dp = abs(event.x - sx), abs(event.x - ex), abs(event.x - px)
        thr = self.hr * 3
        # 最も近いハンドルを選択
        candidates = [('start', ds), ('end', de), ('pos', dp)]
        candidates.sort(key=lambda c: c[1])
        best, dist = candidates[0]
        if dist < thr:
            self.dragging = best
        else:
            # バー上クリック → 再生位置を移動
            self.dragging = 'pos'
            self._drag(event)

    def _drag(self, event):
        if not self.dragging:
            return
        v = self._x2v(event.x)
        if self.dragging == 'start':
            self.start_val = max(self.min_val, min(v, self.end_val - 0.05))
            # 始点が再生位置を追い越した場合、再生位置も移動
            if self.pos_val < self.start_val:
                self.pos_val = self.start_val
        elif self.dragging == 'end':
            self.end_val = min(self.max_val, max(v, self.start_val + 0.05))
            if self.pos_val > self.end_val:
                self.pos_val = self.end_val
        else:  # pos
            self.pos_val = max(self.start_val, min(v, self.end_val))
        self._draw()
        if self.command:
            self.command(self.dragging,
                         {'start': self.start_val, 'end': self.end_val,
                          'pos': self.pos_val}[self.dragging])

    def _release(self, event):
        self.dragging = None

    def get_start(self):
        return self.start_val

    def get_end(self):
        return self.end_val

    def get_pos(self):
        return self.pos_val

    def set_pos(self, v):
        self.pos_val = max(self.start_val, min(v, self.end_val))
        self._draw()

    def set_playback_pos(self, v):
        self.pos_val = v
        self._draw()

    def clear_playback_pos(self):
        self._draw()


# --- トリミングウィンドウ ---

class TrimWindow:
    PW, PH = 640, 360

    def __init__(self, parent, video_path, callback):
        self.callback = callback
        self.video_path = video_path
        self.playing = False
        self.play_job = None
        self.current_pos = 0.0
        self.audio_player = AudioPlayer()
        self.audio_ready = False
        self.temp_audio = os.path.join(
            tempfile.gettempdir(), f"clipper_preview_{os.getpid()}.wav"
        )

        self.cap = cv2.VideoCapture(video_path)
        if not self.cap.isOpened():
            messagebox.showerror("エラー", "動画を開けませんでした。")
            callback(None)
            return

        self.total_frames = int(self.cap.get(cv2.CAP_PROP_FRAME_COUNT))
        self.fps = self.cap.get(cv2.CAP_PROP_FPS) or 30
        self.duration = self.total_frames / self.fps

        # ウィンドウ
        self.win = ctk.CTkToplevel(parent)
        self.win.title("トリミング — プレビュー")
        self.win.resizable(False, False)
        self.win.protocol("WM_DELETE_WINDOW", self.on_cancel)
        self.win.grab_set()

        # プレビュー
        self.canvas = tk.Label(self.win, bg="black", width=self.PW, height=self.PH)
        self.canvas.pack(padx=10, pady=(10, 5))

        # 時間情報
        fi = ctk.CTkFrame(self.win, fg_color="transparent")
        fi.pack(fill=tk.X, padx=15)
        self.time_label = ctk.CTkLabel(fi, text=f"00:00:00.00 / {format_time(self.duration)}",
                                       font=("Consolas", 12))
        self.time_label.pack(side=tk.LEFT)
        self.range_label = ctk.CTkLabel(fi, text=f"選択範囲: {format_time(self.duration)}",
                                        font=("Consolas", 12), text_color="gray")
        self.range_label.pack(side=tk.RIGHT)

        # レンジスライダー
        self.slider = RangeSlider(self.win, 0, self.duration, width=640, height=50,
                                  command=self.on_range_change)
        self.slider.config(bg="#242424")  # CTk dark mode background default
        self.slider.pack(padx=10)

        # 始点/終点ラベル
        fl = ctk.CTkFrame(self.win, fg_color="transparent")
        fl.pack(fill=tk.X, padx=25)
        self.start_lbl = ctk.CTkLabel(fl, text="● 始点: 00:00:00.00",
                                      text_color="#28a745", font=("Consolas", 12, "bold"))
        self.start_lbl.pack(side=tk.LEFT)
        self.end_lbl = ctk.CTkLabel(fl, text=f"● 終点: {format_time(self.duration)}",
                                    text_color="#dc3545", font=("Consolas", 12, "bold"))
        self.end_lbl.pack(side=tk.RIGHT)

        # 再生コントロール
        fc = ctk.CTkFrame(self.win, fg_color="transparent")
        fc.pack(pady=8)
        self.btn_play = ctk.CTkButton(fc, text="▶ 再生", width=120, height=32, command=self.toggle_play)
        self.btn_play.pack(side=tk.LEFT, padx=5)

        # 確定/キャンセル
        fb = ctk.CTkFrame(self.win, fg_color="transparent")
        fb.pack(pady=8)
        ctk.CTkButton(fb, text="この範囲で圧縮", width=180, height=36,
                      command=self.on_confirm, fg_color="#1f6aa5", font=("Meiryo", 12, "bold")).pack(side=tk.LEFT, padx=8)
        ctk.CTkButton(fb, text="トリミングなしで圧縮", width=180, height=36,
                      command=self.on_no_trim, fg_color="gray", font=("Meiryo", 12)).pack(side=tk.LEFT, padx=8)
        ctk.CTkButton(fb, text="キャンセル", width=100, height=36,
                      command=self.on_cancel, fg_color="transparent", border_width=1,
                      text_color=("gray10", "gray90")).pack(side=tk.LEFT, padx=8)

        # 音声抽出ステータス
        self.audio_status = ctk.CTkLabel(self.win, text="♪ 音声読込中...", text_color="gray",
                                         font=("Meiryo", 11))
        self.audio_status.pack()

        # 初期フレーム
        self._seek_show(0)
        # 音声を裏で抽出
        threading.Thread(target=self._extract_audio, daemon=True).start()

    def _extract_audio(self):
        """ffmpegで音声をWAVに抽出する（バックグラウンド）。"""
        try:
            cmd = [get_tool_path('ffmpeg'), '-y', '-i', self.video_path,
                   '-vn', '-ac', '2', '-ar', '44100', '-acodec', 'pcm_s16le',
                   self.temp_audio]
            si = subprocess.STARTUPINFO()
            si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            subprocess.run(cmd, startupinfo=si, check=True,
                           capture_output=True)
            # MCI はスレッド親和性があるため、メインスレッドでロード
            self.win.after(0, self._load_audio_on_main)
        except Exception:
            self.audio_ready = False
            self.win.after(0, lambda: self.audio_status.configure(
                text="♪ 音声なし", text_color="gray"))

    def _load_audio_on_main(self):
        """メインスレッドで音声をMCIにロードする。"""
        ok = self.audio_player.load(self.temp_audio)
        self.audio_ready = ok
        self.audio_status.configure(
            text="♪ 音声準備完了" if ok else "♪ 音声なし",
            text_color="#28a745" if ok else "gray")

    def _seek_show(self, seconds):
        """指定秒にシークしてフレーム表示。"""
        seconds = max(0, min(seconds, self.duration))
        fn = max(0, min(int(seconds * self.fps), self.total_frames - 1))
        self.cap.set(cv2.CAP_PROP_POS_FRAMES, fn)
        ret, frame = self.cap.read()
        if ret:
            self.current_pos = seconds
            self._show(frame)
            self.time_label.configure(
                text=f"{format_time(seconds)} / {format_time(self.duration)}")

    def _show(self, frame):
        """cv2フレームをtkinterに表示。"""
        frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        frame = cv2.resize(frame, (self.PW, self.PH))
        img = ImageTk.PhotoImage(Image.fromarray(frame))
        self.canvas.config(image=img)
        self.canvas.imgtk = img

    def on_range_change(self, handle, val):
        """スライダー操作時のプレビュー更新。"""
        self.start_lbl.configure(text=f"● 始点: {format_time(self.slider.get_start())}")
        self.end_lbl.configure(text=f"● 終点: {format_time(self.slider.get_end())}")
        dur = self.slider.get_end() - self.slider.get_start()
        self.range_label.configure(text=f"選択範囲: {format_time(max(0, dur))}")
        # 再生中にハンドルを操作したら停止してシーク
        if self.playing:
            self.stop_play()
        # プレビューは常に再生位置ハンドルの位置を表示
        self._seek_show(self.slider.get_pos())

    def toggle_play(self):
        if self.playing:
            self.stop_play()
        else:
            self.start_play()

    def start_play(self):
        self.playing = True
        self.btn_play.configure(text="■ 停止")
        # 再生位置ハンドルの現在位置から再生開始
        start_pos = self.slider.get_pos()
        # 再生位置が終点にいたら始点に戻す
        if start_pos >= self.slider.get_end() - 0.1:
            start_pos = self.slider.get_start()
            self.slider.set_pos(start_pos)
        self.cap.set(cv2.CAP_PROP_POS_FRAMES, int(start_pos * self.fps))
        self.play_wall_start = time.time()
        self.play_pos_start = start_pos
        # 音声再生
        if self.audio_ready:
            self.audio_player.play_from(start_pos)
        self._play_tick()

    def stop_play(self):
        self.playing = False
        if hasattr(self, 'btn_play'):
            self.btn_play.configure(text="▶ 再生")
        if self.play_job:
            self.win.after_cancel(self.play_job)
            self.play_job = None
        if self.audio_ready:
            self.audio_player.stop()
        self.slider.clear_playback_pos()

    def _play_tick(self):
        if not self.playing:
            return

        elapsed = time.time() - self.play_wall_start
        target = self.play_pos_start + elapsed
        end = self.slider.get_end()

        if target >= end:
            self._seek_show(end)
            self.stop_play()
            return

        # フレームを順次読み込み、遅れている場合はスキップ
        target_frame = int(target * self.fps)
        cur_frame = int(self.cap.get(cv2.CAP_PROP_POS_FRAMES))

        if target_frame > cur_frame + 5:
            # 大幅に遅れ→シーク
            self.cap.set(cv2.CAP_PROP_POS_FRAMES, target_frame)
            ret, frame = self.cap.read()
        else:
            # 順次読み込み（不要フレームはスキップ）
            ret, frame = True, None
            while cur_frame < target_frame and ret:
                ret, frame = self.cap.read()
                cur_frame += 1
            if frame is None:
                ret, frame = self.cap.read()

        if ret and frame is not None:
            self.current_pos = target
            self._show(frame)
            self.time_label.configure(
                text=f"{format_time(target)} / {format_time(self.duration)}")
            self.slider.set_playback_pos(target)

        self.play_job = self.win.after(33, self._play_tick)  # ~30fps表示

    def on_confirm(self):
        self.stop_play()
        s, e = self.slider.get_start(), self.slider.get_end()
        self._cleanup()
        self.callback((s, e))

    def on_no_trim(self):
        self.stop_play()
        self._cleanup()
        self.callback((0, self.duration))

    def on_cancel(self):
        self.stop_play()
        self._cleanup()
        self.callback(None)

    def _cleanup(self):
        self.audio_player.close()
        self.cap.release()
        self.win.destroy()
        try:
            if os.path.exists(self.temp_audio):
                os.remove(self.temp_audio)
        except Exception:
            pass


# --- メインGUIクラス ---

class DiscordCompressorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Discord用 動画圧縮ツール")
        self.root.geometry("450x390")
        self.root.resizable(False, False)

        self.label_title = ctk.CTkLabel(root, text="動画をDiscordサイズに圧縮",
                                        font=("Meiryo", 16, "bold"))
        self.label_title.pack(pady=15)

        # 目標サイズ
        fs = ctk.CTkFrame(root, fg_color="transparent")
        fs.pack(pady=10)
        ctk.CTkLabel(fs, text="目標サイズ (MB):").pack(side=tk.LEFT, padx=5)
        self.entry_size = ctk.CTkEntry(fs, width=70)
        self.entry_size.insert(0, "9.5")
        self.entry_size.pack(side=tk.LEFT, padx=5)

        # 音声チャンネル
        fa = ctk.CTkFrame(root, fg_color="transparent")
        fa.pack(pady=10)
        ctk.CTkLabel(fa, text="音声チャンネル:").pack(side=tk.LEFT, padx=5)
        self.audio_channel = tk.StringVar(value="stereo")
        ctk.CTkRadioButton(fa, text="ステレオ", variable=self.audio_channel,
                           value="stereo").pack(side=tk.LEFT, padx=10)
        ctk.CTkRadioButton(fa, text="モノラル", variable=self.audio_channel,
                           value="mono").pack(side=tk.LEFT, padx=10)

        # ボタン
        self.btn_select = ctk.CTkButton(root, text="動画ファイルを選択して開始",
                                        command=self.select_file, height=45,
                                        font=("Meiryo", 14, "bold"))
        self.btn_select.pack(pady=20, fill=tk.X, padx=40)
        self.btn_select.configure(state="disabled")

        # ステータス
        self.status_label = ctk.CTkLabel(root, text="起動中...", text_color="gray")
        self.status_label.pack(pady=5)
        self.progress = ctk.CTkProgressBar(root, width=350)
        self.progress.pack(pady=10)
        self.progress.set(0)

        self.check_ffmpeg_setup()

    def update_status(self, message, color="gray", progress_mode=None):
        self.status_label.configure(text=message, text_color=color)
        if progress_mode == 'start':
            self.progress.configure(mode='indeterminate')
            self.progress.start()
        elif progress_mode == 'stop':
            self.progress.stop()
            self.progress.configure(mode='determinate')
            self.progress.set(0)

    def check_ffmpeg_setup(self):
        threading.Thread(target=self._setup_ffmpeg_thread, daemon=True).start()

    def _setup_ffmpeg_thread(self):
        fp = get_tool_path('ffmpeg')
        pp = get_tool_path('ffprobe')
        if fp and pp and os.path.exists(fp) and os.path.exists(pp):
            self.root.after(0, lambda: self.update_status("準備完了", "gray"))
            self.root.after(0, lambda: self.btn_select.configure(state="normal"))
            return
        self.root.after(0, lambda: self.update_status(
            "FFmpegダウンロード中... (初回のみ)", "#1f6aa5", 'start'))
        try:
            url = "https://www.gyan.dev/ffmpeg/builds/ffmpeg-release-essentials.zip"
            base_dir = get_base_path()
            zip_path = os.path.join(base_dir, "ffmpeg_temp.zip")
            urllib.request.urlretrieve(url, zip_path)
            self.root.after(0, lambda: self.update_status("展開中...", "blue"))
            with zipfile.ZipFile(zip_path, 'r') as zf:
                for f in zf.namelist():
                    if f.endswith("bin/ffmpeg.exe"):
                        with open(os.path.join(base_dir, "ffmpeg.exe"), "wb") as out:
                            out.write(zf.read(f))
                    elif f.endswith("bin/ffprobe.exe"):
                        with open(os.path.join(base_dir, "ffprobe.exe"), "wb") as out:
                            out.write(zf.read(f))
            if os.path.exists(zip_path):
                os.remove(zip_path)
            self.root.after(0, lambda: self.update_status("セットアップ完了！", "#28a745", 'stop'))
            self.root.after(0, lambda: self.btn_select.configure(state="normal"))
            messagebox.showinfo("完了", "セットアップが完了しました。")
        except Exception as e:
            self.root.after(0, lambda: self.update_status("セットアップ失敗", "#dc3545", 'stop'))
            messagebox.showerror("エラー", f"エラー: {e}")

    def select_file(self):
        path = filedialog.askopenfilename(
            title="圧縮する動画を選択",
            filetypes=[("Video Files", "*.mp4 *.mov *.avi *.mkv"), ("All Files", "*.*")])
        if not path:
            return
        self.current_input = path
        TrimWindow(self.root, path, self.on_trim_done)

    def on_trim_done(self, result):
        if result is None:
            return
        trim_start, trim_end = result
        default_fn = os.path.splitext(os.path.basename(self.current_input))[0] + "_discord.mp4"
        output = filedialog.asksaveasfilename(
            title="保存先を指定してください", initialfile=default_fn,
            defaultextension=".mp4", filetypes=[("MP4 Files", "*.mp4")])
        if not output:
            return
        threading.Thread(target=self.run_compression,
                         args=(self.current_input, output, trim_start, trim_end),
                         daemon=True).start()

    def get_duration(self, path):
        cmd = [get_tool_path('ffprobe'), '-v', 'error', '-show_entries',
               'format=duration', '-of', 'json', path]
        si = subprocess.STARTUPINFO()
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        r = subprocess.run(cmd, capture_output=True, text=True, startupinfo=si, check=True)
        return float(json.loads(r.stdout)['format']['duration'])

    def run_compression(self, input_path, output_path, trim_start=0, trim_end=None):
        self.btn_select.configure(state="disabled")
        self.update_status("解析中...", "#1f6aa5", 'start')
        try:
            ffmpeg = get_tool_path('ffmpeg')
            target_mb = float(self.entry_size.get())
            full_dur = self.get_duration(input_path)
            if trim_end is None or trim_end >= full_dur:
                trim_end = full_dur
            duration = trim_end - trim_start
            if duration <= 0:
                raise ValueError("トリミング後の長さが0以下です。")

            while True:
                audio_kbps = 128
                total_bits = target_mb * 8 * 1024 * 1024
                audio_bits = audio_kbps * 1024 * duration
                vbr = (total_bits - audio_bits) / duration
                if vbr < 10000:
                    raise ValueError("動画が長すぎます。")

                self.update_status(f"エンコード中... ({int(vbr/1000)}kbps)", "#f69c0d")
                cmd = [ffmpeg, '-y']
                if trim_start > 0:
                    cmd.extend(['-ss', str(trim_start)])
                cmd.extend(['-i', input_path])
                if trim_end < full_dur:
                    cmd.extend(['-t', str(duration)])
                cmd.extend([
                    '-c:v', 'libx264', '-b:v', f'{int(vbr)}',
                    '-maxrate', f'{int(vbr * 1.5)}', '-bufsize', f'{int(vbr * 2)}',
                    '-c:a', 'aac', '-b:a', f'{audio_kbps}k',
                ])
                if self.audio_channel.get() == "mono":
                    cmd.extend(['-ac', '1'])
                cmd.append(output_path)

                si = subprocess.STARTUPINFO()
                si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                subprocess.run(cmd, check=True, startupinfo=si)

                file_size_mb = os.path.getsize(output_path) / (1024 * 1024)
                if file_size_mb >= 10.0:
                    try:
                        os.remove(output_path)
                    except Exception:
                        pass
                    target_mb -= 0.5
                    if target_mb <= 0:
                        raise ValueError("目標サイズを下げて再試行しましたが、10MB以下にできませんでした。")
                    self.update_status(f"サイズ超過({file_size_mb:.2f}MB)。目標サイズを{target_mb:.1f}MBに下げて再試行中...", "#f69c0d")
                    continue
                else:
                    break

            self.update_status("完了！", "#28a745", 'stop')
            messagebox.showinfo("成功", f"保存しました:\n{output_path}\nサイズ: {file_size_mb:.2f}MB")
        except Exception as e:
            self.update_status("エラー", "#dc3545", 'stop')
            messagebox.showerror("エラー", str(e))
        finally:
            self.btn_select.configure(state="normal")
            txt = self.status_label.cget("text")
            if "エラー" not in txt and "完了" not in txt:
                self.update_status("待機中...", "gray")

if __name__ == "__main__":
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")
    root = ctk.CTk()
    app = DiscordCompressorApp(root)
    root.mainloop()