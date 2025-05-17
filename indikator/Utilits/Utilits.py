import os
import subprocess
import platform


# Функция для отображения изображения
def open_image(image_path):
    if os.path.exists(image_path):
        # Определяем операционную систему и выбираем команду для открытия файла
        if platform.system() == "Windows":
            os.startfile(image_path)  # Открыть с помощью программы по умолчанию на Windows
        elif platform.system() == "Darwin":  # macOS
            subprocess.call(["open", image_path])  # Открыть на macOS
        else:  # предполагается Linux
            subprocess.call(["xdg-open", image_path])  # Открыть на Linux
    else:
        print("Изображение не найдено:", image_path)

# Функция для отображения EXCEL файла
def open_excel(file_path):
    if os.path.exists(file_path):
        # Определяем операционную систему и выбираем команду для открытия файла
        if platform.system() == "Windows":
            os.startfile(file_path)  # Открыть с помощью программы по умолчанию на Windows
        elif platform.system() == "Darwin":  # macOS
            subprocess.call(["open", file_path])  # Открыть на macOS
        else:  # предполагается Linux
            subprocess.call(["xdg-open", file_path])  # Открыть на Linux
    else:
        print("Файл не найден:", file_path)
