import logging
from typing import Dict, Any

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Add this dictionary to map chemical formulas to full names
gas_names = {
    "CO": "оксиду углерода",
    "H2S": "сероводороду",
    "NO": "оксиду азота",
    "NO2": "диоксиду азота",
    "PM10": "PM₁₀",
    "C10H8": "нафталину",
    "C6H5OH": "фенолу",
    "C6H6": "бензолу",
    "C7H8": "толуолу",
    "C8H8": "стиролу",
    "CH2O": "формальдегиду"
}

def format_data(analysis_results: Dict[str, Any]) -> str:
    """
    Форматирует результаты анализа в строку.

    :param analysis_results: Словарь с результатами анализа
    :return: Отформатированная строка
    """
    formatted_output = "Превышения ПДКмр:\n"

    for category in ["Москва", "МО"]:
        formatted_output += f"{category}\n"
        category_data = []

        for gas, data in analysis_results.items():
            if data[category] != "нет превышений":
                # Use the full name of the gas instead of the formula
                full_name = gas_names.get(gas, gas)
                category_data.append(f"по {full_name} {data[category]}")

        if category_data:
            formatted_output += "\n".join(category_data) + "\n"
        else:
            formatted_output += "нет превышений\n"

    return formatted_output.strip()


def main(analysis_results: Dict[str, Any]) -> str:
    """
    Основная функция для форматирования данных.
    
    :param analysis_results: Словарь с результатами анализа
    :return: Отформатированная строка
    """
    logging.info("Начало форматирования данных")
    formatted_output = format_data(analysis_results)
    logging.info("Форматирование данных завершено")
    return formatted_output