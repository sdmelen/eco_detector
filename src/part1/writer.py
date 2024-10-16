import logging
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def create_word_document(formatted_results: str, output_file: str):
    """
    Создает документ Word на основе отформатированных результатов.
    
    :param formatted_results: Отформатированная строка с результатами
    :param output_file: Путь для сохранения документа Word
    """
    doc = Document()
    
    # Установка стилей
    styles = doc.styles
    style_normal = styles['Normal']
    style_normal.font.name = 'Times New Roman'
    style_normal.font.size = Pt(12)
    
    # Добавление заголовка
    title = doc.add_paragraph()
    title_run = title.add_run("Данные АСКЗА")
    title_run.bold = True
    title_run.font.size = Pt(14)
    title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # Разбор отформатированных результатов и добавление в документ
    lines = formatted_results.split('\n')
    for line in lines:
        p = doc.add_paragraph()
        if line.startswith('<b>'):
            # Жирный текст
            run = p.add_run(line.replace('<b>', '').replace('</b>', ''))
            run.bold = True
        elif line.startswith('<i>'):
            # Курсив
            run = p.add_run(line.replace('<i>', '').replace('</i>', ''))
            run.italic = True
        else:
            # Обычный текст
            parts = line.split('<b>')
            for i, part in enumerate(parts):
                if i == 0:
                    p.add_run(part)
                else:
                    bold_part, normal_part = part.split('</b>')
                    run = p.add_run(bold_part)
                    run.bold = True
                    p.add_run(normal_part)

    # Сохранение документа
    doc.save(output_file)
    logging.info(f"Документ Word сохранен: {output_file}")

def main(formatted_results: str, output_file: str):
    """
    Основная функция для создания документа Word.
    
    :param formatted_results: Отформатированная строка с результатами
    :param output_file: Путь для сохранения документа Word
    """
    logging.info("Начало создания документа Word")
    create_word_document(formatted_results, output_file)
    logging.info("Создание документа Word завершено")

if __name__ == "__main__":
    # Пример использования
    sample_results = """<b>Превышения ПДКмр:</b>
<i>Москва</i>
<b>по H2S на 2 АСКЗА:</b>
<b>до 4,5 ПДКмр</b> в 01:00 19.09.2024 (Гурьянова),
<b>до 2,9 ПДКмр</b> в 08:20 18.09.2024 (Жулебино);
<b>по CO на 1 АСКЗА:</b>
<b>до 1,1 ПДКмр</b> в 22:40 18.09.2024 (Жулебино);
<i>МО</i>
нет превышений"""
    
    main(sample_results, "output.docx")