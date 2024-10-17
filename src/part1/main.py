import os
from docx import Document
from part1 import process_part1
from part2 import process_part2

def merge_documents(doc1_path, doc2_path, output_path):
    doc1 = Document(doc1_path)
    doc2 = Document(doc2_path)

    # Добавляем разрыв страницы перед содержимым второго документа
    doc1.add_page_break()

    # Копируем содержимое второго документа в первый
    for element in doc2.element.body:
        doc1.element.body.append(element)

    # Сохраняем объединенный документ
    doc1.save(output_path)

def main():
    # Пути к директориям с данными
    input_directory_part1 = "C:/Users/HUAWEI/Desktop/eco_detector/data/part1"
    input_directory_part2 = "C:/Users/HUAWEI/Desktop/eco_detector/data/part2/"
    
    # Обработка первой части
    output_file_part1 = "output.docx"
    process_part1(input_directory_part1)
    
    # Обработка второй части
    output_file_part2 = "result.docx"
    process_part2(input_directory_part2)
    
    # Объединение документов
    merged_output_file = "merged_output.docx"
    merge_documents(output_file_part1, output_file_part2, merged_output_file)
    
    print(f"Объединенный документ сохранен как {merged_output_file}")
    
    # Опционально: удаление исходных файлов
    # os.remove(output_file_part1)
    # os.remove(output_file_part2)

if __name__ == "__main__":
    main()