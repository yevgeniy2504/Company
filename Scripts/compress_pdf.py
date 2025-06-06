import os
import subprocess

# Укажи здесь полный путь к Ghostscript exe
GS_PATH = r"k:\DOP\OED\METHOD&TOOLS\3 - PROJECTS\2 - ON GOING\PYTHON\PYTHON\gs10051w64\bin\gswin64c.exe"

def compress_pdf_ghostscript(input_path, output_path, quality='ebook'):
    """
    Сжимает PDF с помощью Ghostscript.
    quality: один из ['screen', 'ebook', 'printer', 'prepress', 'default']
    """
    quality_map = {
        'screen': '/screen',       # низкое качество, самый маленький размер
        'ebook': '/ebook',         # хорошее качество, средний размер
        'printer': '/printer',     # высокое качество
        'prepress': '/prepress',   # лучшее качество (для печати)
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

        os.remove(input_path)
        print(f"🗑 Удалён оригинальный файл: {input_path}")

    except subprocess.CalledProcessError as e:
        print(f"❌ Ошибка Ghostscript: {e}")

def process_folder(folder, quality='ebook'):
    for root, _, files in os.walk(folder):
        for file in files:
            if file.lower().endswith('.pdf') and not file.startswith('~$'):
                full_path = os.path.join(root, file)
                compressed_path = full_path.replace('.pdf', '_compressed.pdf')
                compress_pdf_ghostscript(full_path, compressed_path, quality=quality)

if __name__ == "__main__":
    folder_path = r"k:\DOP\OED"  # Укажи свою папку с PDF
    process_folder(folder_path, quality='ebook')
