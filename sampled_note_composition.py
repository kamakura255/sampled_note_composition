import pandas as pd
import os
import re
from tkinter import Tk, filedialog
from pydub import AudioSegment
import math

# 音階と半音差のマッピング（基準: A=0）
NOTE_SEMITONE = {
    "C": -9, "C#": -8, "D": -7, "D#": -6,
    "E": -5, "F": -4, "F#": -3, "G": -2,
    "G#": -1, "A": 0, "A#": 1, "B": 2
}

def get_semitone_distance(base_note, target_note):
    """半音距離を計算"""
    # base_noteの解析
    if '#' in base_note:
        base_letter = base_note[:2]  # 例: "C#4" -> "C#"
        base_octave = int(base_note[2:])  # 例: "C#4" -> 4
    else:
        base_letter = base_note[0]  # 例: "C4" -> "C"
        base_octave = int(base_note[1:])  # 例: "C4" -> 4
    
    # target_noteの解析
    if len(target_note) == 1:
        # 例: "C" -> "C", 4
        target_letter, target_octave = target_note, 4
    elif len(target_note) == 2 and target_note[1] == '#':
        # 例: "C#" -> "C#", 4
        target_letter, target_octave = target_note, 4
    elif '#' in target_note:
        # 例: "C#4" -> "C#", 4
        target_letter = target_note[:2]
        target_octave = int(target_note[2:])
    else:
        # 例: "C4" -> "C", 4
        target_letter = target_note[0]
        target_octave = int(target_note[1:])

    semitone_base = NOTE_SEMITONE[base_letter] + 12 * base_octave
    semitone_target = NOTE_SEMITONE[target_letter] + 12 * target_octave
    return semitone_target - semitone_base

def change_pitch(sound, semitone_diff):
    """ピッチ変更（改良版）"""
    if semitone_diff == 0:
        return sound
    
    # 音程変更の倍率を計算
    pitch_ratio = 2.0 ** (semitone_diff / 12.0)
    
    # 新しいサンプリングレートを計算
    new_sample_rate = int(sound.frame_rate * pitch_ratio)
    
    # サンプリングレート変更による速度変化を補正
    pitched_sound = sound._spawn(sound.raw_data, overrides={'frame_rate': new_sample_rate})
    
    # 元のサンプリングレートに戻す（ピッチは変更されたまま）
    return pitched_sound.set_frame_rate(sound.frame_rate)

# GUIでファイル選択
root = Tk()
root.withdraw()

# 音階ファイル選択
paths = filedialog.askopenfilenames(title="音階ファイル（例：C4.wav, A5.wavなど）を選択")
note_files_raw = {}

print("選択されたファイル:")
for path in paths:
    filename = os.path.splitext(os.path.basename(path))[0].upper()
    print(f"  ファイル名: {filename}")
    match = re.match(r"^([A-G]#?)[0-9]$", filename)
    if match:
        note_letter = match.group(1)
        print(f"    抽出された音階: {note_letter}")
        if note_letter not in note_files_raw:
            note_files_raw[note_letter] = path
        else:
            print(f"    {note_letter} は既に存在するためスキップ")
    else:
        print(f"    マッチしませんでした（正規表現: ^([A-G]#?)[0-9]$）")

print(f"\n音階ファイル辞書の内容: {note_files_raw}")

note_files = {}
for k, v in note_files_raw.items():
    try:
        audio = AudioSegment.from_wav(v)
        note_files[k] = audio
        print(f"  {k}: 読み込み成功 - {len(audio)}ms, {audio.frame_rate}Hz, {audio.channels}ch")
    except Exception as e:
        print(f"  {k}: 読み込み失敗 - {e}")

available_notes = list(note_files.keys())
print("読み込んだ音階:", available_notes)

# 必要な音階（A〜G#）があるか確認
required_notes = ["A", "B", "C", "D", "E", "F", "G"]
missing = [note for note in required_notes if note not in note_files]
if missing:
    raise Exception(f"次の音階ファイルが足りません: {', '.join(missing)}")

# Excel読み込み
excel_path = filedialog.askopenfilename(title="Excelファイルを選択")
df = pd.read_excel(excel_path)
output_dir = os.path.dirname(excel_path)

# 「hh:mm:ss:fff」形式のミリ秒部分（fff）だけピリオドに変換
df["時刻修正"] = df["時刻(hh:mm:ss:fff)"].str.replace(r"(?<=\d{2}:\d{2}:\d{2}):", ".", regex=True)

# 正しく変換された文字列を使ってミリ秒に変換
df["ms"] = pd.to_timedelta(df["時刻修正"]).dt.total_seconds() * 1000

# 音声合成
output = AudioSegment.silent(duration=0)
print(f"\n音声合成を開始します。データ行数: {len(df)}")

for i in range(len(df)):
    start_ms = int(df.iloc[i]["ms"])
    end_ms = int(df.iloc[i + 1]["ms"]) if i < len(df) - 1 else start_ms + 500
    duration = end_ms - start_ms
    note_full = df.iloc[i]["音階（国際式）"].upper().strip()
    
    print(f"\n行 {i+1}: 音階='{note_full}', 開始={start_ms}ms, 長さ={duration}ms")

    match = re.match(r"([A-G]#?)([0-9]?)", note_full)
    if not match:
        print(f"  無効な音階形式: {note_full}")
        continue

    note_name = match.group(1)
    note_octave = match.group(2)
    if note_octave:
        target_note = note_name + note_octave
    else:
        target_note = note_name + "4"  # オクターブがなければ4を仮定
    
    print(f"  解析結果: note_name='{note_name}', target_note='{target_note}'")
    print(f"  利用可能な音階: {list(note_files.keys())}")

    # 基本音階を取得（シャープの場合は元の音階から）
    if '#' in note_name:
        base_note_name = note_name[0]  # F# -> F, A# -> A
        base_sound = note_files.get(base_note_name)
        if base_sound is None:
            print(f"  {base_note_name} の基本音が読み込まれていません。スキップ。")
            continue
        print(f"  {note_name} を {base_note_name} から計算で作成します")
    else:
        base_sound = note_files.get(note_name)
        if base_sound is None:
            print(f"  {note_name} の基本音が読み込まれていません。スキップ。")
            continue
        print(f"  {note_name} の基本音を使用します")

    # ピッチ補正：基本音階のオクターブ4から目標音階への変換
    base_reference = (note_name[0] if '#' in note_name else note_name) + "4"
    semitone_diff = get_semitone_distance(base_reference, target_note)
    
    # シャープの場合は追加で+1半音
    if '#' in note_name:
        semitone_diff += 1
    
    print(f"  半音差: {semitone_diff} (基準: {base_reference} -> {target_note})")
    
    try:
        adjusted_sound = change_pitch(base_sound, semitone_diff)
        print(f"  ピッチ調整完了: {len(adjusted_sound)}ms")
    except Exception as e:
        print(f"  ピッチ調整エラー: {e}")
        continue

    # 長さ調整
    if duration <= 0:
        print(f"  警告: 無効な長さ {duration}ms をスキップ")
        continue
        
    if len(adjusted_sound) < duration:
        # 音声を繰り返して必要な長さにする
        repeats = (duration // len(adjusted_sound)) + 1
        adjusted_sound = (adjusted_sound * repeats)[:duration]
    else:
        adjusted_sound = adjusted_sound[:duration]
    
    print(f"  長さ調整完了: {len(adjusted_sound)}ms")

    # 無音追加（必要なら）
    if start_ms > len(output):
        silence_duration = start_ms - len(output)
        output += AudioSegment.silent(duration=silence_duration)
        print(f"  無音追加: {silence_duration}ms")

    # 合成
    try:
        output = output.overlay(adjusted_sound, position=start_ms)
        print(f"  合成完了 - 総長: {len(output)}ms")
    except Exception as e:
        print(f"  合成エラー: {e}")
        continue

# 保存（より安全な形式で）
output_path = os.path.join(output_dir, "romantic_railway_警笛完成版.wav")

# 音声データの正規化（音量調整）
if len(output) > 0:
    # 音量を適切なレベルに調整
    output = output.normalize()
    
    # 16bit PCM形式で保存
    output.export(output_path, format="wav", parameters=["-acodec", "pcm_s16le"])
    print(f"完成しました！ファイル名: {output_path}")
    print(f"ファイルサイズ: {os.path.getsize(output_path) / 1024:.1f} KB")
    print(f"再生時間: {len(output) / 1000:.2f} 秒")
else:
    print("エラー: 音声データが生成されませんでした。")