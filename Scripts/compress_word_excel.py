import os
import zipfile
import shutil
from PIL import Image
import io

def compress_image(data, ext, quality=70):
    try:
        img = Image.open(io.BytesIO(data))
        if img.mode in ("P", "RGBA"):
            img = img.convert("RGB")
        buf = io.BytesIO()
        if ext.lower() == ".png":
            img.save(buf, format="PNG", optimize=True)
        else:
            img.save(buf, format="JPEG", optimize=True, quality=quality)
        return buf.getvalue()
    except Exception as e:
        print(f"Ошибка сжатия изображения: {e}")
        return data  # возвращаем оригинал, если ошибка

def should_process(file_path):
    filename = os.path.basename(file_path)
    return (
        filename.lower().endswith(('.docx', '.xlsx')) and
        not filename.startswith('~$')
    )

def process_office_file(file_path):
    print(f"Обработка: {file_path}")
    tmp_dir = file_path + "_tmp"

    try:
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(tmp_dir)
    except zipfile.BadZipFile:
        print(f"  ❌ Пропущено — файл не является zip-архивом: {file_path}")
        return

    media_path = os.path.join(tmp_dir, 'word', 'media') if 'doc' in file_path.lower() else os.path.join(tmp_dir, 'xl', 'media')
    if not os.path.exists(media_path):
        print("  ℹ️ Нет встроенных изображений.")
    else:
        for img_name in os.listdir(media_path):
            ext = os.path.splitext(img_name)[1].lower()
            if ext in ['.jpg', '.jpeg', '.png']:
                img_path = os.path.join(media_path, img_name)
                try:
                    with open(img_path, 'rb') as f:
                        data = f.read()
                    new_data = compress_image(data, ext)
                    with open(img_path, 'wb') as f:
                        f.write(new_data)
                    print(f"  ✅ Сжато: {img_name}")
                except Exception as e:
                    print(f"  ⚠️ Ошибка при обработке {img_name}: {e}")

    # Перезапись файла (можно сделать резервную копию, если нужно)
    try:
        with zipfile.ZipFile(file_path, 'w', zipfile.ZIP_DEFLATED) as zip_out:
            for root, _, files in os.walk(tmp_dir):
                for file in files:
                    full_path = os.path.join(root, file)
                    arcname = os.path.relpath(full_path, tmp_dir)
                    zip_out.write(full_path, arcname)
        print("  ✅ Готово")
    except Exception as e:
        print(f"  ❌ Ошибка при сохранении файла: {e}")
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)

def process_folder(folder):
    for root, _, files in os.walk(folder):
        for file in files:
            full_path = os.path.join(root, file)
            if should_process(full_path):
                process_office_file(full_path)

if __name__ == "__main__":
    folder_path = r"k:\DOP\OED\METHOD&TOOLS\3 - PROJECTS\1 - COMPLETED\4 - STANDARDS\2 - 5S CHECK LIST"
    process_folder(folder_path)
