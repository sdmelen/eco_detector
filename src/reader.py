import os
import pandas as pd
import logging
from typing import Dict, Any
import re

# Настройка логирования
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def simplify_station_name(station_name: str) -> str:
    """
    Упрощает название станции, обрабатывая специальные случаи и удаляя ненужные обозначения.
    
    :param station_name: Исходное название станции
    :return: Упрощенное название станции
    """
    # Словарь специальных случаев
    special_cases = {
        "М2 (Жулебино) (С)": "Жулебино",
        "М1 (Очаковское) (С)": "Очаковское",
        "М (Балашиха-Речная) ()": "Балашиха-Речная",
        "МКАД 105 восток (Сп)": "МКАД 105 восток",
        "МКАД 52 запад (Сп)": "МКАД 52 запад"
    }

    # Проверяем, является ли станция специальным случаем
    if station_name in special_cases:
        return special_cases[station_name]

    # Общая обработка для остальных случаев
    # Удаляем обозначения в скобках: (С), (Ж), (А), (Сп), ()
    station_name = re.sub(r'\s*\([СЖСА]?п?\)\s*', '', station_name)
    
    # Удаляем пустые скобки
    station_name = re.sub(r'\s*\(\)\s*', '', station_name)
    
    # Если название начинается с "М1 " или "М2 ", удаляем эту часть
    station_name = re.sub(r'^М[12]?\s+\(([^)]+)\)', r'\1', station_name)

    return station_name.strip()

def read_excel_files(directory: str) -> Dict[str, pd.DataFrame]:
    """
    Читает все Excel файлы в указанной директории и возвращает словарь с обработанными данными.
    
    :param directory: Путь к директории с Excel файлами
    :return: Словарь, где ключ - имя файла (газ), значение - DataFrame с данными
    """
    data_dict = {}
    
    # Проверяем существование директории
    if not os.path.exists(directory):
        logging.error(f"Директория не найдена: {directory}")
        return data_dict

    # Получаем список файлов в директории
    files = [f for f in os.listdir(directory) if f.endswith('.xlsx')]
    
    if not files:
        logging.warning(f"В директории {directory} не найдено Excel файлов")
        return data_dict
    
    for file in files:
        file_path = os.path.join(directory, file)
        gas_name = os.path.splitext(file)[0]
        
        try:
            df = pd.read_excel(file_path)
            processed_df = process_dataframe(df, gas_name)
            if not processed_df.empty:
                data_dict[gas_name] = processed_df
            else:
                logging.warning(f"Файл {file} не содержит данных, превышающих ПДКмр.")
        except Exception as e:
            logging.error(f"Ошибка при обработке файла {file}: {str(e)}")
            continue
    
    if not data_dict:
        logging.warning("Не удалось обработать ни один файл с данными")
    else:
        logging.info(f"Успешно обработано файлов: {len(data_dict)}")
    
    return data_dict

def process_dataframe(df: pd.DataFrame, gas_name: str) -> pd.DataFrame:
    """
    Обрабатывает DataFrame: фильтрует данные, добавляет категорию станции 
    и упрощает названия станций.
    
    :param df: Исходный DataFrame
    :param gas_name: Название газа (имя файла)
    :return: Обработанный DataFrame
    """
    required_columns = ["Макс раз знач (в ПДКмр)", "Макс раз знач (дата и вр)", "Станция"]
    
    # Проверка наличия необходимых столбцов
    if not all(col in df.columns for col in required_columns):
        logging.error(f"В файле {gas_name} отсутствуют необходимые столбцы")
        return pd.DataFrame()
    
    # Фильтрация данных
    df = df[df["Макс раз знач (в ПДКмр)"] > 1.00]
    
    # Добавление категории станции и упрощение названий
    mo_stations = ["МО", "Звенигород", "Балашиха-Салтыковка", "Реутов-2", "М (Балашиха-Речная)"]
    df["Категория"] = df["Станция"].apply(lambda x: "Московская область" if any(station in x for station in mo_stations) else "Москва")
    
    # Упрощение названий станций
    df["Станция"] = df["Станция"].apply(lambda x: simplify_station_name(x))
    
    return df[required_columns + ["Категория"]]