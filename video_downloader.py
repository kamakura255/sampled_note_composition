import tkinter as tk
from tkinter import ttk, messagebox
import yt_dlp
import os
import re
from datetime import datetime
import pytz
import sys

class VideoDownloader:
    def __init__(self, root):
        self.root = root
        self.root.title("動画ダウンローダー")
        self.root.geometry("600x500")
        
        # メインフレーム
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # URL入力
        url_frame = ttk.Frame(main_frame)
        url_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(url_frame, text="URLを入力してください:").grid(row=0, column=0, sticky=tk.W)
        self.url_var = tk.StringVar()
        self.url_entry = ttk.Entry(url_frame, textvariable=self.url_var, width=50)
        self.url_entry.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=2)
        ttk.Button(url_frame, text="URL読み込み", command=self.load_video_info).grid(row=1, column=1, padx=5)
        
        # 動画情報表示
        info_frame = ttk.LabelFrame(main_frame, text="動画情報", padding="5")
        info_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        self.title_var = tk.StringVar(value="タイトル: ")
        ttk.Label(info_frame, textvariable=self.title_var, wraplength=550).grid(row=0, column=0, sticky=tk.W)
        
        self.duration_var = tk.StringVar(value="長さ: ")
        ttk.Label(info_frame, textvariable=self.duration_var).grid(row=1, column=0, sticky=tk.W)
        
        # 時間指定
        time_frame = ttk.LabelFrame(main_frame, text="時間指定", padding="5")
        time_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(time_frame, text="開始時間:").grid(row=0, column=0, sticky=tk.W)
        self.start_time_var = tk.StringVar(value="00:00")
        self.start_time_entry = ttk.Entry(time_frame, textvariable=self.start_time_var, width=10)
        self.start_time_entry.grid(row=0, column=1, padx=5)
        
        ttk.Label(time_frame, text="終了時間:").grid(row=0, column=2, sticky=tk.W)
        self.end_time_var = tk.StringVar(value="00:00")
        self.end_time_entry = ttk.Entry(time_frame, textvariable=self.end_time_var, width=10)
        self.end_time_entry.grid(row=0, column=3, padx=5)
        
        ttk.Label(time_frame, text="(形式: MM:SS または HH:MM:SS)").grid(row=0, column=4, padx=5)
        
        # ダウンロードタイプ
        ttk.Label(main_frame, text="ダウンロード形式:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.download_type = tk.StringVar(value="video")
        ttk.Radiobutton(main_frame, text="動画 (MP4)", variable=self.download_type, 
                       value="video").grid(row=4, column=0, sticky=tk.W)
        ttk.Radiobutton(main_frame, text="音声のみ (MP3)", variable=self.download_type,
                       value="audio").grid(row=5, column=0, sticky=tk.W)
        
        # ダウンロードボタン
        ttk.Button(main_frame, text="ダウンロード", command=self.download).grid(row=6, column=0, pady=20)
        
        # 進行状況
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, length=400, variable=self.progress_var)
        self.progress_bar.grid(row=7, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # ステータスラベル
        self.status_var = tk.StringVar(value="準備完了")
        ttk.Label(main_frame, textvariable=self.status_var).grid(row=8, column=0, sticky=tk.W, pady=5)

    def validate_youtube_url(self, url):
        patterns = [
            r'^https?://(?:www\.)?youtube\.com/watch\?v=[\w-]+',
            r'^https?://(?:www\.)?youtu\.be/[\w-]+',
            r'^https?://(?:www\.)?youtube\.com/shorts/[\w-]+',
        ]
        return any(bool(re.match(pattern, url)) for pattern in patterns)

    def load_video_info(self):
        url = self.url_var.get().strip()
        if not url:
            messagebox.showerror("エラー", "URLを入力してください")
            return
            
        if not self.validate_youtube_url(url):
            messagebox.showerror("エラー", "無効なYouTube URLです")
            return
            
        try:
            self.status_var.set("動画情報を取得中...")
            print(f"[情報] 動画情報を取得中: {url}")
            
            ydl_opts = {
                'quiet': True,
                'no_warnings': True,
            }
            
            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                info = ydl.extract_info(url, download=False)
                
                # タイトルを設定
                self.title_var.set(f"タイトル: {info.get('title', '不明')}")
                
                # 動画の長さを設定
                duration = info.get('duration', 0)
                hours = duration // 3600
                minutes = (duration % 3600) // 60
                seconds = duration % 60
                
                if hours > 0:
                    duration_str = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
                else:
                    duration_str = f"{minutes:02d}:{seconds:02d}"
                    
                self.duration_var.set(f"長さ: {duration_str}")
                
                # 終了時間を動画の長さに設定
                self.end_time_var.set(duration_str)
                print(f"[情報] 動画情報の取得完了")
                
        except Exception as e:
            error_message = f"動画情報の取得に失敗しました: {str(e)}"
            print(f"[エラー] {error_message}")
            messagebox.showerror("エラー", error_message)
            return

    def format_time_for_download(self, seconds):
        hours = seconds // 3600
        minutes = (seconds % 3600) // 60
        secs = seconds % 60
        return f"{hours:02d}:{minutes:02d}:{secs:02d}"

    def parse_time(self, time_str):
        try:
            parts = time_str.split(':')
            if len(parts) == 2:  # MM:SS
                return int(parts[0]) * 60 + int(parts[1])
            elif len(parts) == 3:  # HH:MM:SS
                return int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])
            else:
                raise ValueError("Invalid time format")
        except:
            raise ValueError("Invalid time format")

    def progress_hook(self, d):
        if d['status'] == 'downloading':
            # ダウンロード進捗を計算
            if 'total_bytes' in d:
                percent = (d['downloaded_bytes'] / d['total_bytes']) * 100
                self.progress_var.set(percent)
                self.status_var.set(f"ダウンロード中... {percent:.1f}%")
            elif 'downloaded_bytes' in d:
                self.status_var.set(f"ダウンロード中... {d['downloaded_bytes'] / 1024 / 1024:.1f}MB")
        elif d['status'] == 'finished':
            self.status_var.set("変換中...")
            print(f"[成功] ダウンロード完了: {d['filename']}")

    def download(self):
        url = self.url_var.get().strip()
        if not url:
            messagebox.showerror("エラー", "URLを入力してください")
            return

        try:
            # 時間をパース
            start_seconds = self.parse_time(self.start_time_var.get())
            end_seconds = self.parse_time(self.end_time_var.get())
            
            if start_seconds >= end_seconds:
                messagebox.showerror("エラー", "終了時間は開始時間より後である必要があります")
                return
                
            # 現在の日本時間を取得してフォルダ名を生成
            jst = pytz.timezone('Asia/Tokyo')
            now = datetime.now(jst)
            download_dir = f"downloads_{now.strftime('%Y%m%d%H%M')}"
            
            # ダウンロードディレクトリの作成
            if not os.path.exists(download_dir):
                os.makedirs(download_dir)
                print(f"[情報] 保存フォルダを作成: {download_dir}")

            # 時間指定文字列を作成（HH:MM:SS形式）
            start_time = self.format_time_for_download(start_seconds)
            end_time = self.format_time_for_download(end_seconds)

            # ffmpegのパスを設定
            ffmpeg_dir = r'C:\ffmpeg\bin'
            ffmpeg_exe = os.path.join(ffmpeg_dir, 'ffmpeg.exe')
            
            if os.path.exists(ffmpeg_exe):
                has_ffmpeg = True
                print(f"[情報] ffmpegを検出: {ffmpeg_exe}")
                os.environ["PATH"] = os.pathsep.join([ffmpeg_dir, os.environ.get("PATH", "")])
                yt_dlp.utils.std_headers['PATH'] = os.environ["PATH"]
            else:
                has_ffmpeg = False
                print(f"[警告] ffmpegが見つかりません。以下の場所にffmpeg.exeを配置してください:")
                print(f"- {ffmpeg_exe}")

            # yt-dlp オプションの設定
            ydl_opts = {
                'progress_hooks': [self.progress_hook],
                'outtmpl': os.path.join(download_dir, '%(title)s.%(ext)s'),
                'quiet': True,
                'no_warnings': True,
            }

            if has_ffmpeg:
                ydl_opts.update({
                    'postprocessors': [{
                        'key': 'FFmpegVideoConvertor',
                        'preferedformat': 'mp4',
                    }],
                    'postprocessor_args': {
                        'FFmpegVideoConvertor': [
                            '-ss', self.format_time_for_download(start_seconds),
                            '-t', str(end_seconds - start_seconds)
                        ]
                    }
                })
                print("[情報] 時間指定ダウンロードを使用")

            if self.download_type.get() == "audio":
                if has_ffmpeg:
                    ydl_opts.update({
                        'format': 'bestaudio/best',
                        'postprocessors': [{
                            'key': 'FFmpegExtractAudio',
                            'preferredcodec': 'mp3',
                            'preferredquality': '192',
                        }],
                    })
                    print("[情報] 音声モードで準備中（MP3変換あり）...")
                else:
                    ydl_opts.update({
                        'format': 'bestaudio/best',
                    })
                    print("[情報] 音声モードで準備中（MP3変換なし）...")
                    messagebox.showwarning(
                        "警告",
                        "FFmpegが見つかりません。音声ファイルはMP3に変換されず、元のフォーマットでダウンロードされます。"
                    )
            else:
                # 動画モードの設定
                ydl_opts.update({
                    'format': 'best[ext=mp4]/best',  # 単一フォーマットを指定
                })
                print("[情報] 動画モードで準備中...")


            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                print(f"[情報] ダウンロード開始: {url}")
                print(f"[情報] 時間指定: {self.start_time_var.get()} - {self.end_time_var.get()}")
                ydl.download([url])

            self.status_var.set("ダウンロード完了!")
            self.progress_var.set(100)
            messagebox.showinfo("成功", "ダウンロードが完了しました")
            
        except ValueError as e:
            error_message = f"時間形式が正しくありません: {str(e)}"
            print(f"[エラー] {error_message}")
            messagebox.showerror("エラー", error_message)
        except Exception as e:
            error_message = str(e)
            print(f"[エラー] {type(e).__name__}: {error_message}")
            self.status_var.set("エラーが発生しました")
            messagebox.showerror("エラー", f"ダウンロード中にエラーが発生しました: {error_message}")
        finally:
            self.progress_var.set(0)

def main():
    root = tk.Tk()
    app = VideoDownloader(root)
    root.mainloop()

if __name__ == "__main__":
    main()