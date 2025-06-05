import os
import subprocess

ffmpeg_path = r"k:\DOP\OED\METHOD&TOOLS\3 - PROJECTS\2 - ON GOING\PYTHON\ffmpeg-2025-06-02-git-688f3944ce-essentials_build\bin\ffmpeg.exe"
root_folder = r"K:\DOP\OED"
video_exts = [".mp4", ".avi", ".mov", ".mkv", ".wmv", ".flv"]

for dirpath, dirnames, filenames in os.walk(root_folder):
    for file in filenames:
        if any(file.lower().endswith(ext) for ext in video_exts) and "compressed" not in file.lower():
            input_path = os.path.join(dirpath, file)
            name, ext = os.path.splitext(file)
            output_path = os.path.join(dirpath, f"{name}_compressed{ext}")

            cmd = [
                ffmpeg_path,
                "-i", input_path,
                "-vcodec", "libx264",
                "-crf", "23",
                output_path
            ]
            print(f"Compressing: {input_path}")
            result = subprocess.run(cmd)

            if result.returncode == 0:
                print(f"Compression succeeded, deleting original: {input_path}")
                os.remove(input_path)
            else:
                print(f"Compression failed for: {input_path}")
