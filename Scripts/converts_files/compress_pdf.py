import os
import subprocess
import time
import stat

# Укажи здесь путь к Ghostscript
GS_PATH = r"k:\DOP\OED\METHOD&TOOLS\3 - PROJECTS\2 - ON GOING\PYTHON\PYTHON\gs10051w64\bin\gswin64c.exe"

def safe_remove(path, retries=5, delay=1):
    """
    Безопасное удаление файла с повторами, если файл занят другим процессом.
    """
    for attempt in range(retries):
        try:
            if os.path.exists(path):
                # Убираем флаг read-only, если установлен
                os.chmod(path, stat.S_IWRITE)
                os.remove(path)
                print(f"🗑 Удалён оригинальный файл: {path}")
            return
        except PermissionError:
            print(f"⚠️ Попытка {attempt+1}: файл занят, ждём {delay} сек...")
            time.sleep(delay)
        except Exception as e:
            print(f"❌ Неожиданная ошибка при удалении: {e}")
            return
    print(f"❌ Не удалось удалить файл: {path}")

def compress_pdf_ghostscript(input_path, output_path, quality='ebook'):
    """
    Сжимает PDF с помощью Ghostscript.
    quality: один из ['screen', 'ebook', 'printer', 'prepress', 'default']
    """
    quality_map = {
        'screen': '/screen',
        'ebook': '/ebook',
        'printer': '/printer',
        'prepress': '/prepress',
        'default': '/default'
    }
    gs_quality = quality_map.get(quality, '/ebook')

    if not os.path.isfile(input_path):
        print(f"❌ Файл не найден: {input_path}")
        return

    gs_command = [
        GS_PATH,
        '-sDEVICE=pdfwrite',
        '-dCompatibilityLevel=1.4',
        f'-dPDFSETTINGS={gs_quality}',
        '-dNOPAUSE',
        '-dQUIET',
        '-dBATCH',
        f'-sOutputFile={output_path}',
        input_path
    ]

    try:
        subprocess.run(gs_command, check=True, creationflags=subprocess.CREATE_NO_WINDOW)
        print(f"✅ Сжат с помощью Ghostscript: {output_path}")
        safe_remove(input_path)
    except subprocess.CalledProcessError as e:
        print(f"❌ Ошибка Ghostscript: {e}")

def process_folder(folder, quality='ebook'):
    """
    Обходит папку рекурсивно и сжимает все PDF-файлы.
    """
    for root, _, files in os.walk(folder):
        for file in files:
            if file.lower().endswith('.pdf') and not file.startswith('~$') and '_compressed' not in file:
                full_path = os.path.join(root, file)
                compressed_path = full_path.replace('.pdf', '_compressed.pdf')
                compress_pdf_ghostscript(full_path, compressed_path, quality=quality)

if __name__ == "__main__":
    folder_path = r"k:/EXCHANGE/Yevgeniy Karabekov/"  # Укажи путь к папке с PDF
    process_folder(folder_path, quality='ebook')
