import os
import pandas as pd
import logging
from typing import Dict, Any
import re

# Настройка логирования
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def read_excel_files(directory: str) -> Dict[str, pd.DataFrame]:
    """
    Читает все Excel файлы в указанной директории и возвращает словарь с обработанными данными.
    
    :param directory: Путь к директории с Excel файлами
    :return: Словарь, где ключ - имя файла (газ), значение - DataFrame с данными
    """
    data_dict = {}
    
    for root, _, files in os.walk(directory):
        for file in files:
            if file.endswith('.xlsx'):
                file_path = os.path.join(root, file)
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
    
    # Добавление категории станции
    mo_stations = ["МО", "Звенигород", "Балашиха-Салтыковка", "Реутов-2", "М (Балашиха-Речная)"]
    df["Категория"] = df["Станция"].apply(lambda x: "МО" if any(x.startswith(station) for station in mo_stations) else "Москва")
    
    # Упрощение названий станций
    df["Станция"] = df["Станция"].apply(lambda x: simplify_station_name(x))
    
    return df[required_columns + ["Категория"]]

def simplify_station_name(station_name: str) -> str:
    """
    Упрощает название станции, удаляя ненужные скобки и их содержимое.
    
    :param station_name: Исходное название станции
    :return: Упрощенное название станции
    """

    station_name = re.sub(r'\s*\([СЖСА]?\)\s*', '', station_name) # Удаляем (С), (Ж), (А), (Сп), () 
    return station_name

def main(directory: str) -> Dict[str, Any]:
    """
    Основная функция для чтения и обработки Excel файлов.
    
    :param directory: Путь к директории с Excel файлами
    :return: Словарь с обработанными данными
    """
    logging.info(f"Начало обработки файлов в директории: {directory}")
    data = read_excel_files(directory)
    logging.info(f"Обработка завершена. Прочитано {len(data)} файлов.")
    return data

if __name__ == "__main__":
    # Пример использования
    input_directory = "C:/Users/HUAWEI/Desktop/eco_detector/data/part1"
    result = main(input_directory)
    print(f"Обработано газов: {list(result.keys())}")
    for gas, df in result.items():
        print(f"\n{gas}:")
        print(df.head())