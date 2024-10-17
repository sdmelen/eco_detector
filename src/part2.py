import pandas as pd
from docx.shared import Pt
from docx import Document
from datetime import datetime, timedelta
import logging
import os


def format_datetime_range(start_str, end_str, minutes_to_subtract):
    """Форматирует диапазон дат и времени в нужный формат."""
    date_time_format = "%d/%m/%Y %H:%M"
    original_start = datetime.strptime(start_str, date_time_format)
    original_end = datetime.strptime(end_str, date_time_format)

    adjusted_start = original_start - timedelta(minutes=minutes_to_subtract)

    start_time_str = adjusted_start.strftime("%H:%M")
    end_time_str = original_end.strftime("%H:%M")

    if adjusted_start.date() == original_end.date():
        date_str = adjusted_start.strftime("%d.%m.%Y")
        return f"с {start_time_str} по {end_time_str} {date_str}"
    else:
        start_date_str = adjusted_start.strftime("%d.%m.%Y")
        end_date_str = original_end.strftime("%d.%m.%Y")
        return f"с {start_time_str} {start_date_str} по {end_time_str} {end_date_str}"


def get_duration_string(total_minutes):
    """Возвращает строку длительности в часах и минутах с правильным склонением."""

    def get_hour_string(hours):
        if hours == 1:
            return "1 час"
        elif 2 <= hours <= 4:
            return f"{hours} часа"
        else:
            return f"{hours} часов"

    def get_minute_string(minutes):
        if minutes == 1:
            return "1 минута"
        elif 2 <= minutes <= 4:
            return f"{minutes} минуты"
        else:
            return f"{minutes} минут"

    hours = total_minutes // 60
    minutes = total_minutes % 60

    hours_str = get_hour_string(hours) if hours > 0 else ""
    minutes_str = get_minute_string(minutes) if minutes > 0 else ""

    return f"{hours_str} {minutes_str}".strip()


def add_paragraph_to_document(document, gas_name, duration_str, details):
    """Добавляет абзац в документ с заданными параметрами."""
    paragraph = document.add_paragraph()
    run = paragraph.add_run(f"по {gas_name} – {duration_str} {details}")
    run.font.name = 'Times New Roman'
    run.font.size = Pt(14)


def record_max_excess_duration(path, file_name, document, gas_names):
    """Находит строки с максимальным количеством точек и записывает результаты в одну строку в документ."""
    df = pd.read_excel(f"{path}{file_name}.xlsx")
    max_points = df['Количество точек'].max()
    max_excess_rows = df[df['Количество точек'] == max_points].values.tolist()

    station_names = [line[0] for line in max_excess_rows]
    stations_str = ', '.join(station_names)

    first_line = max_excess_rows[0]
    duration_minutes = int(first_line[2]) * 20
    duration_str = get_duration_string(duration_minutes)
    formatted_range = format_datetime_range(first_line[3], first_line[4], 20)
    details = f"{formatted_range} ({stations_str})"

    add_paragraph_to_document(document, gas_names[file_name], duration_str, details)


def record_total_excess_duration(path, file_name, document, gas_names):
    """Находит суммарную длительность превышений для каждой точки и записывает результаты в одну строку в документ."""

    df = pd.read_excel(f"{path}{file_name}.xlsx")
    points_map = df.groupby(df.columns[0])['Количество точек'].sum().to_dict()
    max_value = max(points_map.values())

    max_keys = [key for key, value in points_map.items() if value == max_value]
    stations_str = ', '.join(max_keys)

    total_duration_minutes = max_value * 20
    duration_str = get_duration_string(total_duration_minutes)
    details = f"({stations_str})"

    add_paragraph_to_document(document, gas_names[file_name], duration_str, details)


def add_custom_text(document, text, font_name='Times New Roman', font_size=14, bold=False):
    """Добавляет заголовок с указанным шрифтом, размером и жирным стилем."""
    heading = document.add_paragraph()
    run = heading.add_run(text)
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.bold = bold


def process_multiple_files(path, file_names, document):
    """Обрабатывает несколько файлов и записывает результаты в один документ."""
    gas_names = {
        "CO_п": "оксиду углерода",
        "H2S_период": "сероводороду",
        "NO_п": "оксиду азота",
        "NO2_п": "диоксиду азота",
        "PM10_п": "PM₁₀"
    }

    add_custom_text(document, 'Максимальная непрерывная длительность превышений:', font_name='Times New Roman',
                    font_size=14, bold=True)
    for file_name in file_names:
        record_max_excess_duration(path, file_name, document, gas_names)

    add_custom_text(document, 'Максимальная общая длительность превышений:', font_name='Times New Roman', font_size=14,
                    bold=True)
    for file_name in file_names:
        record_total_excess_duration(path, file_name, document, gas_names)


def main():
    """Основная функция для выполнения всех операций."""
    path = "C:/Users/HUAWEI/Desktop/eco_detector/data/part2/"

    document = process_part2(path)
    document.save('result.docx')


def process_part2(directory_name, output_file='result.docx'):
    logging.info(f"Начало обработки part2 с входной директорией {directory_name}")
    document = Document()
    file_names = [
        'CO_п',
        'NO2_п',
        'NO_п',
        'PM10_п',
        'H2S_период'
    ]
    process_multiple_files(directory_name, file_names, document)
    
    # Сохранение документа
    document.save(output_file)
    logging.info(f"Документ Word сохранен: {output_file}")
    logging.info(f"Файл {output_file} создан: {os.path.exists(output_file)}")
    
    return output_file

def main():
    """Основная функция для выполнения всех операций."""
    path = "C:/Users/HUAWEI/Desktop/eco_detector/data/part2/"
    output_file = process_part2(path)
    logging.info(f"Обработка part2 завершена. Результат сохранен в {output_file}")

if __name__ == "__main__":
    main()

