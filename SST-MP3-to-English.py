import moonshine_onnx as moonshine # or import moonshine_onnx
from pydub import AudioSegment
import os

#To-Do: add italian languange. Moonshine maybe don't support italian, check it

def split_mp3(file_path, segment_length_minutes=1):
    # Carica il file audio
    audio = AudioSegment.from_mp3(file_path)

    # Calcola la durata del segmento in millisecondi
    segment_length_ms = segment_length_minutes * 60 * 1000

    # Lunghezza totale dell'audio
    total_length_ms = len(audio)

    # Crea una cartella per i segmenti
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    output_dir = f"{base_name}_segments"
    os.makedirs(output_dir, exist_ok=True)

    segment_paths = []

    # Divide e salva i segmenti
    for i in range(0, total_length_ms, segment_length_ms):
        segment = audio[i:i + segment_length_ms]
        segment_file_name = os.path.join(output_dir, f"{base_name}_part{i // segment_length_ms + 1}.mp3")
        segment.export(segment_file_name, format="mp3")
        print(f"Saved segment: {segment_file_name}")
        segment_paths.append(segment_file_name)

    return segment_paths

file_path = "C:\\Users\\AdminAkro\\Music\\Riunione WMATA 10-11 luglio 2025\\6.mp3"
segment_paths = split_mp3(file_path)
transcription = []

for index in range(0, len(segment_paths)):
    with open("transcrizione-6.txt", "a") as file:
        print(f"Sto trascrivendo il file")
        # or moonshine_onnx.transcribe(...)
        transcription.append(moonshine.transcribe(segment_paths[index], 'moonshine/base'))
        print(transcription[index], file=file)


print("Fatto")
