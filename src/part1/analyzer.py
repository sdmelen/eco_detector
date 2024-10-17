import pandas as pd
import logging
from typing import Dict, Any
from datetime import datetime

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def analyze_data(data_dict: Dict[str, pd.DataFrame]) -> Dict[str, Any]:
    results = {}
    for gas, df in data_dict.items():
        logging.info(f"Анализ данных для газа: {gas}")
        results[gas] = analyze_gas_data(df)
    return results

def parse_datetime(date_string: str) -> datetime:
    formats = ["%d/%m/%Y %H:%M", "%Y-%m-%d %H:%M:%S"]
    for fmt in formats:
        try:
            return datetime.strptime(date_string, fmt)
        except ValueError:
            continue
    raise ValueError(f"Не удалось распознать формат даты: {date_string}")

def analyze_gas_data(df: pd.DataFrame) -> Dict[str, Any]:
    categories = ["Москва", "МО"]
    gas_results = {}

    for category in categories:
        category_df = df[df["Категория"] == category]

        if category_df.empty:
            gas_results[category] = {
                "количество_станций": 0,
                "превышения": []
            }
        else:
            # Сортировка по убыванию значения ПДКмр
            category_df = category_df.sort_values("Макс раз знач (в ПДКмр)", ascending=False)

            превышения = []
            current_pdkmr = None
            current_stations = []

            for _, row in category_df.iterrows():
                pdkmr = round(float(row["Макс раз знач (в ПДКмр)"]), 1)
                
                try:
                    date_time = parse_datetime(row["Макс раз знач (дата и вр)"])
                    formatted_time = date_time.strftime("%H:%M %d.%m.%Y")
                except ValueError as e:
                    logging.error(f"Ошибка при обработке даты: {e}")
                    formatted_time = "Неизвестное время"

                if pdkmr == current_pdkmr:
                    current_stations.append(f"в {formatted_time} ({row['Станция']})")
                else:
                    if current_pdkmr is not None:
                        превышения.append({
                            "пдкмр": current_pdkmr,
                            "станции": current_stations
                        })
                    current_pdkmr = pdkmr
                    current_stations = [f"в {formatted_time} ({row['Станция']})"]

            # Добавляем последнюю группу превышений
            if current_pdkmr is not None:
                превышения.append({
                    "пдкмр": current_pdkmr,
                    "станции": current_stations
                })

            gas_results[category] = {
                "количество_станций": category_df["Станция"].nunique(),
                "превышения": превышения
            }

    return gas_results

def format_results(results: Dict[str, Any]) -> Dict[str, Any]:
    formatted_results = {}

    for gas, data in results.items():
        formatted_results[gas] = {
            "Москва": format_category_data(data["Москва"]),
            "МО": format_category_data(data["МО"])
        }

    return formatted_results

def format_category_data(category_data: Dict[str, Any]) -> str:
    if category_data["количество_станций"] == 0:
        return "нет превышений"

    превышения = category_data["превышения"]
    превышения_str = []
    for п in превышения:
        if п['пдкмр'] == 1.0:
            превышения_str.append(f"на уровне 1 ПДКмр " + ", ".join(п['станции']))
        else:
            превышения_str.append(f"до {п['пдкмр']:.1f} ПДКмр " + ", ".join(п['станции']))

    return f"на {category_data['количество_станций']} АСКЗА:\n" + "\n".join(превышения_str)

def main(data_dict: Dict[str, pd.DataFrame]) -> Dict[str, Any]:
    logging.info("Начало анализа данных")
    analysis_results = analyze_data(data_dict)
    formatted_results = format_results(analysis_results)
    logging.info("Анализ данных завершен")
    return formatted_results