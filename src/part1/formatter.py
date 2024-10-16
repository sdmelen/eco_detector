import logging
from typing import Dict, Any

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def format_data(analysis_results: Dict[str, Any]) -> str:
    """
    Форматирует результаты анализа в строку с тегами форматирования.
    
    :param analysis_results: Словарь с результатами анализа
    :return: Отформатированная строка
    """
    formatted_output = "<b>Превышения ПДКмр:</b>\n"

    for category in ["Москва", "МО"]:
        formatted_output += f"<i>{category}</i>\n"
        category_data = []

        for gas, data in analysis_results.items():
            if data[category] != "нет превышений":
                gas_data = format_gas_data(gas, data[category])
                if gas_data:
                    category_data.append(gas_data)

        if category_data:
            formatted_output += "\n".join(category_data) + "\n"
        else:
            formatted_output += "нет превышений\n"

    return formatted_output.strip()

def format_gas_data(gas: str, data: str) -> str:
    """
    Форматирует данные для конкретного газа.
    
    :param gas: Название газа
    :param data: Данные о превышениях для газа
    :return: Отформатированная строка для газа
    """
    lines = data.split('\n')
    station_count = lines[0].split(':')[0].split()[-2]
    formatted_gas = f"<b>по {gas} на {station_count} АСКЗА:</b>\n"

    for line in lines[1:]:
        parts = line.split(' ', 5)  # Разделяем строку на 6 частей
        pdkmr = parts[1]
        time = parts[3]
        date = parts[4]
        station = parts[5].strip('()')  # Убираем скобки вокруг названия станции
        formatted_gas += f"<b>до {pdkmr} ПДКмр</b> {time} {date} {station}),\n"

    return formatted_gas.rstrip(',\n') + ';'

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

if __name__ == "__main__":
    # Пример использования
    sample_results = {
        "H2S": {
            "Москва": "на 2 АСКЗА:\nдо 4.5 ПДКмр в 01:00 19.09.2024 (Гурьянова),\nдо 2.9 ПДКмр в 08:20 18.09.2024 (Жулебино)",
            "МО": "нет превышений"
        },
        "CO": {
            "Москва": "на 1 АСКЗА:\nдо 1.1 ПДКмр в 22:40 18.09.2024 (Жулебино)",
            "МО": "нет превышений"
        }
    }
    
    formatted_output = main(sample_results)
    print(formatted_output)