import os
import subprocess
import argparse


VIDEO_EXTS = [".mp4", ".avi", ".mov", ".mkv", ".wmv", ".flv"]

def compress_videos_in_folder(folder_path, ffmpeg_path, delete_original=True, crf="23"):
    print(f"[INFO] Старт обработки папки: {folder_path}")

    if not os.path.exists(folder_path):
        print(f"[ERROR] Папка не найдена: {folder_path}")
        return

    for dirpath, dirnames, filenames in os.walk(folder_path):
        print(f"[DEBUG] Обработка папки: {dirpath} (файлов: {len(filenames)})")
        for file in filenames:
            print(f"[DEBUG] Проверяю файл: {file}")

            if any(file.lower().endswith(ext) for ext in VIDEO_EXTS) and "compressed" not in file.lower():
                input_path = os.path.join(dirpath, file)
                name, ext = os.path.splitext(file)
                output_path = os.path.join(dirpath, f"{name}_compressed{ext}")

                print(f"[INFO] Сжимаю: {input_path}")

                cmd = [
                    ffmpeg_path,
                    "-i", input_path,
                    "-vcodec", "libx264",
                    "-crf", crf,
                    output_path
                ]

                result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

                if result.returncode == 0:
                    print(f"[OK] Сжатие прошло успешно: {output_path}")
                    if delete_original:
                        try:
                            os.remove(input_path)
                            print(f"[INFO] Удалён оригинал: {input_path}")
                        except Exception as e:
                            print(f"[ERROR] Не удалось удалить оригинал: {e}")
                else:
                    print(f"[ERROR] Сжатие не удалось для: {input_path}")
                    print(result.stderr.decode('cp1251', errors='ignore'))

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Сжимает видеофайлы в папке с помощью ffmpeg.")
    parser.add_argument("folder", help="Путь к папке для обработки видео.")
    parser.add_argument("--ffmpeg", default=r"k:\DOP\OED\METHOD&TOOLS\3 - PROJECTS\2 - ON GOING\PYTHON\ffmpeg-2025-06-02-git-688f3944ce-essentials_build\bin\ffmpeg.exe", help="Путь к исполняемому файлу ffmpeg.")
    parser.add_argument("--crf", default="23", help="Параметр CRF для сжатия (меньше — лучше качество, по умолчанию 23).")
    parser.add_argument("--keep", action="store_true", help="Не удалять оригинальные файлы после сжатия.")

    args = parser.parse_args()

    compress_videos_in_folder(
        folder_path=args.folder,
        ffmpeg_path=args.ffmpeg,
        delete_original=not args.keep,
        crf=args.crf
    )
