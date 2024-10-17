import os
import logging
from reader import main as read_data
from analyzer import main as analyze_data
from formatter import main as format_data
from writer import main as write_document  # Будет использоваться позже

# Настройка логирования
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def main():
    # Путь к директории с Excel файлами УКАЖИТЕ ПУТЬ
    input_directory = "C:/Users/HUAWEI/Desktop/eco_detector/data/part1"
    
    # Проверка существования директории
    if not os.path.exists(input_directory):
        logging.error(f"Директория не найдена: {input_directory}")
        return

    # Чтение данных
    logging.info("Начало чтения данных")
    data = read_data(input_directory)
    logging.info(f"Прочитано {len(data)} файлов")

    if not data:
        logging.warning("Нет данных для анализа")
        return

    # Анализ данных
    logging.info("Начало анализа данных")
    analysis_results = analyze_data(data)
    logging.info("Анализ данных завершен")

    # Форматирование данных
    logging.info("Начало форматирования данных")
    formatted_results = format_data(analysis_results)
    logging.info("Форматирование данных завершено")

    # Вывод отформатированных результатов
    print("\nОтформатированные результаты:")
    print(formatted_results)

    # Запись результатов в документ (будет реализовано позже)
    output_file = "output.docx"
    write_document(formatted_results, output_file)
    logging.info(f"Результаты записаны в файл: {output_file}")

if __name__ == "__main__":
    main()