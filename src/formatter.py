import logging
from typing import Dict, Any

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

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
    Форматирует результаты анализа в строку с правильной пунктуацией.
    """
    formatted_output = "Превышения ПДКмр:\n"

    for category_index, category in enumerate(["Москва", "Московская область"]):
        formatted_output += f"{category}\n"
        category_data = []

        # Собираем все газы с превышениями для текущей категории
        valid_gases = [(gas, data) for gas, data in analysis_results.items() 
                      if data[category] != "нет превышений"]

        if not valid_gases:
            formatted_output += "нет превышений\n"
            continue

        # Обрабатываем каждый газ
        for gas_index, (gas, data) in enumerate(valid_gases):
            gas_data = data[category]
            
            # Получаем строки для каждого уровня превышения
            lines = gas_data.split('\n')
            first_line = lines[0]  # строка с количеством станций
            превышения = lines[1:]  # строки с превышениями
            
            # Форматируем вывод для газа
            gas_output = f"по {gas_names.get(gas, gas)} {first_line}:\n"
            
            # Добавляем каждое превышение с правильной пунктуацией
            for i, превышение in enumerate(превышения):
                gas_output += превышение
                if i < len(превышения) - 1:
                    gas_output += ","  # Запятая между превышениями одного газа
                
                gas_output += "\n"
            
            # Добавляем точку с запятой после всех превышений газа
            # Если это последний газ в последней категории, ставим точку
            if gas_index == len(valid_gases) - 1 and category_index == 1:
                gas_output = gas_output.rstrip() + "."  # Точка в конце всего документа
            else:
                gas_output = gas_output.rstrip() + ";\n"  # Точка с запятой после газа
                
            category_data.append(gas_output)

        formatted_output += "".join(category_data)

    return formatted_output.rstrip()

def main(analysis_results: Dict[str, Any]) -> str:
    """
    Основная функция для форматирования данных.
    """
    logging.info("Начало форматирования данных")
    formatted_output = format_data(analysis_results)
    logging.info("Форматирование данных завершено")
    return formatted_output