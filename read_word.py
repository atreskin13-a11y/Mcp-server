import os
from docx import Document

def read_docx(file_path):
    """
    Читает содержимое .docx файла и возвращает текст.
    
    :param file_path: Путь к .docx файлу
    :return: Извлечённый текст из документа
    """
    try:
        doc = Document(file_path)
        full_text = []
        for paragraph in doc.paragraphs:
            full_text.append(paragraph.text)
        return '\n'.join(full_text)
    except Exception as e:
        return f"Ошибка при чтении файла {file_path}: {e}"

def process_directory(directory):
    """
    Обходит директорию и извлекает текст из всех .docx файлов.
    
    :param directory: Путь к директории
    :return: Словарь с именами файлов и их содержимым
    """
    results = {}
    for root, _, files in os.walk(directory):
        for file in files:
            if file.lower().endswith('.docx'):
                full_path = os.path.join(root, file)
                text = read_docx(full_path)
                results[file] = text
    return results

# Пример использования
if __name__ == "__main__":
    # Укажите путь к вашей директории с Word-документами
    directory_path = r"C:\Users\Kachanov_IA\Desktop\стандарты"
    
    extracted_data = process_directory(directory_path)
    
    for filename, content in extracted_data.items():
        print(f"=== {filename} ===")
        print(content[:500] + "..." if len(content) > 500 else content)
