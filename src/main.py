import os
import logging
from docx import Document
from part1 import process_part1
from part2 import process_part2

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def merge_documents(doc1_path, doc2_path, output_path):
    logging.info(f"Попытка объединения документов: {doc1_path} и {doc2_path}")
    
    if not os.path.exists(doc1_path):
        raise FileNotFoundError(f"Файл {doc1_path} не найден")
    if not os.path.exists(doc2_path):
        raise FileNotFoundError(f"Файл {doc2_path} не найден")
    
    doc1 = Document(doc1_path)
    doc2 = Document(doc2_path)

    #doc1.add_page_break()

    for element in doc2.element.body:
        doc1.element.body.append(element)

    doc1.save(output_path)
    logging.info(f"Объединенный документ сохранен как {output_path}")

def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    logging.info(f"Базовая директория: {base_dir}")

    input_directory_part1 = os.path.join(base_dir, "..", "data", "part1")
    input_directory_part2 = os.path.join(base_dir, "..", "data", "part2/")
    
    output_file_part1 = os.path.join(base_dir, "output.docx")
    output_file_part2 = os.path.join(base_dir, "result.docx")
    
    logging.info(f"Обработка part1 с входной директорией {input_directory_part1}")
    process_part1(input_directory_part1, output_file_part1)
    
    logging.info(f"Обработка part2 с входной директорией {input_directory_part2}")
    process_part2(input_directory_part2, output_file_part2)
    
    logging.info(f"Проверка наличия файла {output_file_part1}: {os.path.exists(output_file_part1)}")
    logging.info(f"Проверка наличия файла {output_file_part2}: {os.path.exists(output_file_part2)}")

    merged_output_file = os.path.join(base_dir, "справка.docx")
    merge_documents(output_file_part1, output_file_part2, merged_output_file)
    os.remove(output_file_part1)
    os.remove(output_file_part2)
    
if __name__ == "__main__":
        main()