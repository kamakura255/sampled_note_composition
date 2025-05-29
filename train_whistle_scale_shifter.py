import tkinter as tk
from tkinter import ttk, filedialog
import pygame
import os
import numpy as np
from scipy.io import wavfile
from scipy.fft import fft
from scipy.signal import find_peaks
import traceback
from pydub import AudioSegment
import librosa
import pandas as pd
from datetime import datetime

class WhistleScaleShifter:
    def __init__(self, root):
        self.root = root
        self.root.title("警笛音階シフター")
        self.root.geometry("400x350")
        
        pygame.mixer.init()
        
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        ttk.Label(main_frame, text="警笛音声ファイルを選択", font=('メイリオ', 12, 'bold')).grid(row=0, column=0, columnspan=2, pady=10)
        
        self.file_path = tk.StringVar()
        file_entry = ttk.Entry(main_frame, textvariable=self.file_path, width=40)
        file_entry.grid(row=1, column=0, padx=5, pady=5)
        
        browse_btn = ttk.Button(main_frame, text="参照", command=self.browse_file)
        browse_btn.grid(row=1, column=1, padx=5, pady=5)
        
        analyze_btn = ttk.Button(main_frame, text="分析して音階を生成", command=self.analyze_and_generate)
        analyze_btn.grid(row=2, column=0, columnspan=2, pady=20)
        
        ttk.Label(main_frame, text="再生コントロール", font=('メイリオ', 12, 'bold')).grid(row=3, column=0, columnspan=2, pady=10)
        
        self.play_original_btn = ttk.Button(main_frame, text="元の警笛を再生", command=self.play_original, state='disabled')
        self.play_original_btn.grid(row=4, column=0, columnspan=2, pady=5)
        
        self.play_scale_btn = ttk.Button(main_frame, text="生成した音階を再生", command=self.play_scale, state='disabled')
        self.play_scale_btn.grid(row=5, column=0, columnspan=2, pady=5)
        
        stop_btn = ttk.Button(main_frame, text="停止", command=self.stop_sound)
        stop_btn.grid(row=6, column=0, columnspan=2, pady=5)
        
        export_btn = ttk.Button(main_frame, text="音階情報をExcelに出力", command=self.export_to_excel, state='disabled')
        export_btn.grid(row=7, column=0, columnspan=2, pady=10)
        self.export_btn = export_btn
        
        self.status_var = tk.StringVar()
        self.status_var.set("警笛音声ファイルを選択してください")
        status_label = ttk.Label(main_frame, textvariable=self.status_var)
        status_label.grid(row=8, column=0, columnspan=2, pady=10)

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[
                ("Audio files", "*.wav;*.mp3;*.m4a"),
                ("Wave files", "*.wav"),
                ("MP3 files", "*.mp3"),
                ("M4A files", "*.m4a"),
                ("All files", "*.*")
            ]
        )
        if file_path:
            self.file_path.set(file_path)
            self.play_original_btn.config(state='normal')
            print(f"ファイルを読み込みました: {file_path}")
            self.status_var.set("ファイルを読み込みました")

    def convert_to_wav(self, input_file):
        print(f"音声ファイルをWAVに変換中: {input_file}")
        
        if input_file.lower().endswith('.wav'):
            return input_file
        
        output_file = os.path.splitext(input_file)[0] + "_converted.wav"
        audio = AudioSegment.from_file(input_file)
        audio.export(output_file, format="wav")
        print(f"変換完了: {output_file}")
        return output_file

    def find_nearest_note(self, freq):
        notes = ['ド', 'レ', 'ミ', 'ファ', 'ソ', 'ラ', 'シ', 'ド']
        ratios = [1.0, 9/8, 5/4, 4/3, 3/2, 5/3, 15/8, 2.0]
        base_c4 = 261.63  # C4 (ド)の周波数
        
        octave = int(np.log2(freq/base_c4))
        norm_freq = freq / (2**octave)
        
        min_diff = float('inf')
        nearest_note = None
        nearest_ratio = None
        
        for note, ratio in zip(notes, ratios):
            note_freq = base_c4 * ratio
            diff = abs(norm_freq - note_freq)
            if diff < min_diff:
                min_diff = diff
                nearest_note = note
                nearest_ratio = ratio
        
        return nearest_note, nearest_ratio

    def analyze_whistle(self, filename):
        print(f"音声ファイルを分析中: {filename}")
        sample_rate, data = wavfile.read(filename)
        print(f"サンプリングレート: {sample_rate}Hz")
        
        if len(data.shape) > 1:
            data = data[:, 0]
        
        data = data.astype(np.float32) / np.max(np.abs(data))
        
        n = len(data)
        freq = np.fft.fftfreq(n, d=1/sample_rate)
        fft_data = np.abs(fft(data))
        
        pos_freq = freq[:n//2]
        pos_fft = fft_data[:n//2]
        
        max_freq_idx = np.argmax(pos_fft)
        base_freq = pos_freq[max_freq_idx]
        
        return base_freq, data, sample_rate

    def get_international_note(self, note_name, freq):
        base_c4 = 261.63  # C4の周波数
        octave = int(np.log2(freq/base_c4)) + 4
        
        note_map = {
            'ド': 'C',
            'レ': 'D',
            'ミ': 'E',
            'ファ': 'F',
            'ソ': 'G',
            'ラ': 'A',
            'シ': 'B'
        }
        return f"{note_map[note_name]}{octave}"

    def generate_scale(self, original_data, sample_rate, base_freq, base_note, base_ratio, output_filename):
        print(f"\n音階を生成中... 検出周波数: {base_freq:.1f}Hz ({base_note})")
        
        # 出力フォルダの作成（年月日時分形式）
        output_dir = os.path.dirname(output_filename)
        date_str = datetime.now().strftime("%Y%m%d%H%M")
        scale_dir = os.path.join(output_dir, f"{date_str}_onkai")
        os.makedirs(scale_dir, exist_ok=True)
        
        notes = ['ドー', 'レー', 'ミー', 'ファー', 'ソー', 'ラー', 'シー', 'ドー']
        notes_base = ['ド', 'レ', 'ミ', 'ファ', 'ソ', 'ラ', 'シ', 'ド']
        ratios = [1.0, 9/8, 5/4, 4/3, 3/2, 5/3, 15/8, 2.0]
        
        self.scale_info = []
        current_time = 0.0
        
        shift_ratios = []
        for ratio in ratios:
            shift_ratios.append(ratio / base_ratio)
        
        print("\n音階の周波数:")
        scale_complete = np.array([], dtype=np.int16)
        
        # 音声の長さと処理パラメータ
        duration = 10.0  # 各音の長さを10秒に設定
        target_length = int(duration * sample_rate)
        fade_time = 0.3  # フェードイン/アウトの時間
        fade_samples = int(fade_time * sample_rate)
        
        # 元の音声から最も強い部分を見つける（全体で1回だけ実行）
        window_size = int(0.1 * sample_rate)  # 100ms窓
        energy = np.array([np.sum(original_data[i:i+window_size]**2) 
                          for i in range(0, len(original_data)-window_size)])
        max_energy_start = np.argmax(energy)
        
        # 最も強い部分から1秒分のデータを取得
        segment_length = int(1.0 * sample_rate)
        if max_energy_start + segment_length > len(original_data):
            max_energy_start = len(original_data) - segment_length
        
        best_segment = original_data[max_energy_start:max_energy_start+segment_length]
        
        for note, note_base, ratio in zip(notes, notes_base, shift_ratios):
            freq = base_freq * ratio
            print(f"{note}: {freq:.1f}Hz")
            
            self.scale_info.append({
                '時刻 (秒)': f"{current_time:.1f}",
                '周波数 (Hz)': f"{freq:.1f}",
                '振幅': 1.0,
                '音階（国際式）': self.get_international_note(note_base, freq),
                '音階（ドレミ式）': note
            })
            
            # シフト量をセント値で計算
            cents = 1200 * np.log2(ratio)
            
            # ベストセグメントにピッチシフトを適用
            shifted_segment = librosa.effects.pitch_shift(
                best_segment.astype(np.float32),
                sr=sample_rate,
                n_steps=cents/100
            )
            
            # 1秒のセグメントを10秒に拡張（音質を維持）
            shifted = np.tile(shifted_segment, 10)[:target_length]
            
            # 冒頭と末尾のみフェードイン/アウト
            fade_in = np.linspace(0, 1, fade_samples)
            fade_out = np.linspace(1, 0, fade_samples)
            
            shifted[:fade_samples] *= fade_in
            shifted[-fade_samples:] *= fade_out
            
            # int16に変換
            shifted = (shifted * 32767).astype(np.int16)
            
            # 個別の音階ファイルとして保存
            note_filename = os.path.join(scale_dir, f"{self.get_international_note(note_base, freq)}.wav")
            wavfile.write(note_filename, sample_rate, shifted)
            print(f"保存: {note_filename}")
            
            # 全体の音階に追加
            silence = np.zeros(int(0.5 * sample_rate), dtype=np.int16)  # 0.5秒の無音
            scale_complete = np.concatenate([scale_complete, shifted, silence])
            
            current_time += 10.5  # 音の長さ(10.0秒) + 無音区間(0.5秒)
        
        # 完全な音階をWAVファイルとして保存
        complete_scale_file = os.path.join(scale_dir, "complete_scale.wav")
        wavfile.write(complete_scale_file, sample_rate, scale_complete)
        print(f"\n完全な音階を保存しました: {complete_scale_file}")
        
        return complete_scale_file

    def analyze_and_generate(self):
        try:
            input_file = self.file_path.get()
            if not input_file:
                print("エラー: ファイルが選択されていません")
                return

            wav_file = self.convert_to_wav(input_file)
            base_freq, original_data, sample_rate = self.analyze_whistle(wav_file)
            note, ratio = self.find_nearest_note(base_freq)
            
            print(f"\n検出された基本周波数: {base_freq:.1f}Hz")
            print(f"最も近い音階: {note} ({base_freq:.1f}Hz)")
            
            self.status_var.set(f"検出音: {note} ({base_freq:.1f}Hz)")

            output_file = os.path.splitext(input_file)[0] + "_scale.wav"
            output_file = self.generate_scale(original_data, sample_rate, base_freq, note, ratio, output_file)
            
            self.play_scale_btn.config(state='normal')
            self.export_btn.config(state='normal')
            self.generated_file = output_file
            print(f"音階の生成が完了しました: {output_file}")

        except Exception as e:
            print("エラーが発生しました:")
            print(traceback.format_exc())

    def export_to_excel(self):
        try:
            if not hasattr(self, 'scale_info'):
                print("音階情報がありません。先に音階を生成してください。")
                return
            
            excel_file = os.path.join(
                os.path.dirname(self.generated_file),
                "scale_info.xlsx"
            )
            
            df = pd.DataFrame(self.scale_info)
            df.to_excel(excel_file, index=False)
            print(f"音階情報をExcelファイルに出力しました: {excel_file}")
            self.status_var.set("音階情報をExcelに出力しました")
            
        except Exception as e:
            print("Excel出力エラー:")
            print(traceback.format_exc())

    def play_original(self):
        self._play_file(self.file_path.get())

    def play_scale(self):
        if hasattr(self, 'generated_file'):
            self._play_file(self.generated_file)
        else:
            print("音階ファイルが生成されていません")

    def _play_file(self, filename):
        try:
            if os.path.exists(filename):
                pygame.mixer.music.stop()
                pygame.mixer.music.load(filename)
                pygame.mixer.music.play()
                print(f"再生中: {filename}")
            else:
                print(f"エラー: ファイルが見つかりません: {filename}")
        except Exception as e:
            print(f"再生エラー: {str(e)}")
            print(traceback.format_exc())

    def stop_sound(self):
        pygame.mixer.music.stop()
        print("再生を停止しました")

if __name__ == "__main__":
    root = tk.Tk()
    app = WhistleScaleShifter(root)
    root.mainloop()