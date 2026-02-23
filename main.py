import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import subprocess
import json
import threading
import sys
import urllib.request
import zipfile
import shutil

# --- ユーティリティ関数（ここを修正！） ---

def get_base_path():
    """
    EXEファイルがある「本来の場所」を取得する。
    PyInstallerの _MEIPASS (一時フォルダ) ではなく、
    sys.executable (EXE本体のパス) のディレクトリを使う。
    """
    if getattr(sys, 'frozen', False):
        # EXEとして実行されている場合
        return os.path.dirname(sys.executable)
    else:
        # 普通のPythonスクリプトとして実行されている場合
        return os.path.dirname(os.path.abspath(__file__))

def get_tool_path(tool_name):
    if os.name == 'nt' and not tool_name.endswith('.exe'):
        tool_name += '.exe'
    
    # EXEの横にあるか探す
    local_path = os.path.join(get_base_path(), tool_name)
    if os.path.exists(local_path):
        return local_path
    
    # 環境変数パスも確認
    return shutil.which(tool_name)

# --- GUIクラス ---

class DiscordCompressorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Discord用 動画圧縮ツール")
        self.root.geometry("450x340")
        self.root.resizable(False, False)

        # UI構築
        self.label_title = tk.Label(root, text="動画をDiscordサイズに圧縮", font=("Meiryo", 14, "bold"))
        self.label_title.pack(pady=10)

        # 設定エリア
        self.frame_size = tk.Frame(root)
        self.frame_size.pack(pady=5)
        tk.Label(self.frame_size, text="目標サイズ (MB):").pack(side=tk.LEFT)
        self.entry_size = tk.Entry(self.frame_size, width=5)
        self.entry_size.insert(0, "9.5")
        self.entry_size.pack(side=tk.LEFT, padx=5)

        # 音声チャンネル設定
        self.frame_audio = tk.Frame(root)
        self.frame_audio.pack(pady=5)
        tk.Label(self.frame_audio, text="音声チャンネル:").pack(side=tk.LEFT)
        self.audio_channel = tk.StringVar(value="stereo")
        tk.Radiobutton(self.frame_audio, text="ステレオ（変更なし）", variable=self.audio_channel, value="stereo").pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(self.frame_audio, text="モノラル", variable=self.audio_channel, value="mono").pack(side=tk.LEFT, padx=5)

        # ボタンエリア
        self.btn_select = tk.Button(root, text="動画ファイルを選択して開始", command=self.select_file, height=2, bg="#e1e1e1")
        self.btn_select.pack(pady=15, fill=tk.X, padx=30)
        self.btn_select.config(state=tk.DISABLED)

        # ステータスエリア
        self.status_label = tk.Label(root, text="起動中...", fg="gray")
        self.status_label.pack(pady=5)

        self.progress = ttk.Progressbar(root, orient=tk.HORIZONTAL, length=350, mode='determinate')
        self.progress.pack(pady=10)
        self.progress['value'] = 0

        # 起動時にFFmpegチェック
        self.check_ffmpeg_setup()

    def update_status(self, message, color="black", progress_mode=None):
        self.status_label.config(text=message, fg=color)
        
        if progress_mode == 'start':
            self.progress.config(mode='indeterminate')
            self.progress.start(10)
        elif progress_mode == 'stop':
            self.progress.stop()
            self.progress.config(mode='determinate')
            self.progress['value'] = 0

    # --- 自動インストール処理 ---
    def check_ffmpeg_setup(self):
        threading.Thread(target=self._setup_ffmpeg_thread, daemon=True).start()

    def _setup_ffmpeg_thread(self):
        ffmpeg_path = get_tool_path('ffmpeg')
        ffprobe_path = get_tool_path('ffprobe')

        if ffmpeg_path and ffprobe_path and os.path.exists(ffmpeg_path) and os.path.exists(ffprobe_path):
            self.root.after(0, lambda: self.update_status("準備完了", "gray"))
            self.root.after(0, lambda: self.btn_select.config(state=tk.NORMAL))
            return

        self.root.after(0, lambda: self.update_status("FFmpegダウンロード中... (初回のみ)", "blue", 'start'))
        
        try:
            url = "https://www.gyan.dev/ffmpeg/builds/ffmpeg-release-essentials.zip"
            zip_name = "ffmpeg_temp.zip"
            base_dir = get_base_path() # ← ここが修正されたので、EXEの隣に保存される

            urllib.request.urlretrieve(url, os.path.join(base_dir, zip_name))
            self.root.after(0, lambda: self.update_status("展開中...", "blue"))

            with zipfile.ZipFile(os.path.join(base_dir, zip_name), 'r') as zip_ref:
                for file in zip_ref.namelist():
                    if file.endswith("bin/ffmpeg.exe"):
                        with open(os.path.join(base_dir, "ffmpeg.exe"), "wb") as f_out:
                            f_out.write(zip_ref.read(file))
                    elif file.endswith("bin/ffprobe.exe"):
                        with open(os.path.join(base_dir, "ffprobe.exe"), "wb") as f_out:
                            f_out.write(zip_ref.read(file))

            if os.path.exists(os.path.join(base_dir, zip_name)):
                os.remove(os.path.join(base_dir, zip_name))

            self.root.after(0, lambda: self.update_status("セットアップ完了！", "green", 'stop'))
            self.root.after(0, lambda: self.btn_select.config(state=tk.NORMAL))
            messagebox.showinfo("完了", "セットアップが完了しました。\nEXEファイルと同じ場所にffmpeg.exeが保存されました。")

        except Exception as e:
            self.root.after(0, lambda: self.update_status("セットアップ失敗", "red", 'stop'))
            messagebox.showerror("エラー", f"エラー: {e}")

    # --- 圧縮処理 ---

    def select_file(self):
        input_path = filedialog.askopenfilename(
            title="圧縮する動画を選択",
            filetypes=[("Video Files", "*.mp4 *.mov *.avi *.mkv"), ("All Files", "*.*")]
        )
        if not input_path:
            return

        default_filename = os.path.splitext(os.path.basename(input_path))[0] + "_discord.mp4"
        
        output_path = filedialog.asksaveasfilename(
            title="保存先を指定してください",
            initialfile=default_filename,
            defaultextension=".mp4",
            filetypes=[("MP4 Files", "*.mp4")]
        )

        if not output_path:
            return

        threading.Thread(target=self.run_compression, args=(input_path, output_path), daemon=True).start()

    def get_duration(self, input_path):
        ffprobe_cmd = get_tool_path('ffprobe')
        cmd = [ffprobe_cmd, '-v', 'error', '-show_entries', 'format=duration', '-of', 'json', input_path]
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        
        result = subprocess.run(cmd, capture_output=True, text=True, startupinfo=startupinfo, check=True)
        return float(json.loads(result.stdout)['format']['duration'])

    def run_compression(self, input_path, output_path):
        self.btn_select.config(state=tk.DISABLED)
        self.update_status("解析中...", "blue", 'start')

        try:
            ffmpeg_cmd = get_tool_path('ffmpeg')
            
            target_mb = float(self.entry_size.get())
            duration = self.get_duration(input_path)

            audio_kbps = 128
            target_total_bits = target_mb * 8 * 1024 * 1024
            audio_bits = audio_kbps * 1024 * duration
            video_bits_available = target_total_bits - audio_bits
            video_bitrate = video_bits_available / duration

            if video_bitrate < 10000:
                raise ValueError("動画が長すぎます。")

            self.update_status(f"エンコード中... ({int(video_bitrate/1000)}kbps)", "orange")

            cmd = [
                ffmpeg_cmd, '-y',
                '-i', input_path,
                '-c:v', 'libx264', '-b:v', f'{int(video_bitrate)}',
                '-maxrate', f'{int(video_bitrate * 1.5)}', '-bufsize', f'{int(video_bitrate * 2)}',
                '-c:a', 'aac', '-b:a', f'{audio_kbps}k',
            ]
            if self.audio_channel.get() == "mono":
                cmd.extend(['-ac', '1'])
            cmd.append(output_path)
            
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            subprocess.run(cmd, check=True, startupinfo=startupinfo)

            self.update_status("完了！", "green", 'stop')
            messagebox.showinfo("成功", f"保存しました:\n{output_path}")

        except Exception as e:
            self.update_status("エラー", "red", 'stop')
            messagebox.showerror("エラー", str(e))
        finally:
            self.btn_select.config(state=tk.NORMAL)
            if "エラー" not in self.status_label.cget("text") and "完了" not in self.status_label.cget("text"):
                self.update_status("待機中...", "gray")

if __name__ == "__main__":
    root = tk.Tk()
    app = DiscordCompressorApp(root)
    root.mainloop()