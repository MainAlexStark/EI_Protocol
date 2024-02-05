from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt
import docx
import os

# Доделать форматы в Условиях проведения поверки
def extract_value(text, start_text, end_text):
    if text.find(start_text) != -1:
        start = text.find(start_text) + len(start_text)

        end = text.find(end_text)

        if end_text == '\n':
            value = text[start:]
        else:
            value = text[start:end]

        return value.strip()
def create_template(path:str):
    file_path = os.path.abspath(__file__)
    main_path = os.path.dirname(os.path.dirname(file_path))

    try:

        doc = Document(path)
        print('Файл успешно открыт')

        # Получение первой таблицы
        table = doc.tables[0]

        # Получение текста из клетки 0 0
        cell = table.cell(0, 0)
        cell_text = cell.text

        # Получить все таблицы в документе
        tables = doc.tables

        # Удалить таблицу по индексу (нумерация начинается с 0)
        table_index = 0
        table = tables[table_index]

        # Удалить все строки в таблице
        for row in table.rows:
            table._tbl.remove(row._tr)

        # Удалить таблицу из контейнера (документа)
        doc.element.body.remove(table._element)


        # Словарь для переменных
        var = {}

        index = 0
        for element in doc.element.body:
            # Проверяем, является ли элемент абзацем
            if isinstance(element, docx.oxml.text.paragraph.CT_P):
                # Получаем абзац
                paragraph = docx.text.paragraph.Paragraph(element, doc)



                if 'УСЛОВИЯ ПРОВЕДЕНИЯ ПОВЕРКИ' in paragraph.text:
                    paragraph.text += '\n' + cell_text
                    paragraph.paragraph_format.alignment = doc.styles['Normal'].paragraph_format.alignment
                    paragraph.paragraph_format.space_after = doc.styles['Normal'].paragraph_format.space_after
                    paragraph.paragraph_format.line_spacing = doc.styles['Normal'].paragraph_format.line_spacing
                    paragraph.paragraph_format.space_before = doc.styles['Normal'].paragraph_format.space_before
                    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
                    for run in paragraph.runs:
                        run.font.size = Pt(12)
                        run.font.all_caps = False

                value = extract_value(paragraph.text, "ПРОТОКОЛ №", "\n")
                if value:
                    print(value)
                    var['номер протокола'] = value

                    paragraph.text = paragraph.text.replace(value, 'НОМЕР_ПРОТОКОЛА')
                    paragraph.text = paragraph.text.replace('№', '#', 1)

                value = extract_value(paragraph.text, "СИ -", "№")
                if value:
                    print(value)
                    var['весы'] = value

                value = extract_value(paragraph.text, "СИ –", "№")
                if value:
                    print(value)
                    var['весы'] = value

                value = extract_value(paragraph.text, "№", "\n")
                if value:
                    print(value)
                    var['номер весов'] = value
                    paragraph.text = paragraph.text.replace(value, 'НОМЕР_ВЕСОВ')

                    for temp_paragraph in doc.paragraphs:
                        temp_paragraph.text = temp_paragraph.text.replace('№','#')

                value = extract_value(paragraph.text, "Принадлежащего:", ", ИНН")
                if value:
                    print(value)
                    var['компания'] = value
                    paragraph.text = paragraph.text.replace(value, "КОМПАНИЯ")
                    paragraph.text = paragraph.text.replace(',', '', 1)

                value = extract_value(paragraph.text, "ИНН", ",")
                if value:
                    print(value)
                    var['инн'] = value
                    paragraph.text = paragraph.text.replace(value, "НОМЕР_ИНН")

                    print(extract_value(paragraph.text, ",", "\n"))
                    var['юридический адрес'] = extract_value(paragraph.text, ",", "\n")
                    paragraph.text = paragraph.text.replace(extract_value(paragraph.text, ",", "\n"), "ЮРИДИЧЕСКИЙ_АДРЕС")

                value = extract_value(paragraph.text, "Место поверки:", "\n")
                if value:
                    print(value)
                    var['место поверки'] = value
                    paragraph.text = paragraph.text.replace(value, "МЕСТО_ПОВЕРКИ")

                value = extract_value(paragraph.text, "ФИФ ОЕИ:", "\n")
                if value:
                    print(value)
                    var['фиф'] = value

                value = extract_value(paragraph.text, "начале поверки:", "°C")
                if value:
                    print(value)
                    var['температура'] = value
                    paragraph.text = paragraph.text.replace(value, "ТЕМПЕРАТУРА")

                value = extract_value(paragraph.text, "конце поверки:", "°C")
                if value:
                    print(value)
                    var['температура'] = value
                    paragraph.text = paragraph.text.replace(value, "ТЕМПЕРАТУРА")

                value = extract_value(paragraph.text, "окружающего воздуха:", "°C")
                if value:
                    print(value)
                    var['температура'] = value
                    paragraph.text = paragraph.text.replace(value, "ТЕМПЕРАТУРА")

                value = extract_value(paragraph.text, "влажность воздуха:", "%")
                if value:
                    print(value)
                    var['влажность'] = value
                    paragraph.text = paragraph.text.replace(value, "ВЛАЖНОСТЬ")

                value = extract_value(paragraph.text, "давление:", "кПа")
                if value:
                    print(value)
                    var['давление'] = value
                    paragraph.text = paragraph.text.replace(value, "ДАВЛЕНИЕ")

                value = extract_value(paragraph.text, "Напряжение сети", "В")
                if value:
                    print(value)
                    var['Напряжение'] = value
                    paragraph.text = paragraph.text.replace(value, "НАПРЯЖЕНИЕ")

                value = extract_value(paragraph.text, "Частота", "Гц")
                if value:
                    print(value)
                    var['Частота'] = value
                    paragraph.text = paragraph.text.replace(value, "ЧАСТОТА")

                if 'ЭТАЛОНЫ, применяемые при поверке' in paragraph.text:
                    var['эталоны поверки'] = doc.paragraphs[index + 1].text
                    doc.paragraphs[index + 1].text = "ЭТАЛОНЫ_ПОВЕРКИ"

                if 'Поверитель' in paragraph.text and '__' in paragraph.text:
                    value = extract_value(paragraph.text, "__", "Дата")
                    if value:
                        print(value)
                        var['поверитель'] = value.replace('_', '')
                        paragraph.text = paragraph.text.replace(value.replace('_', ''), "ПОВЕРИТЕЛЬ")

                        print(f'!!!!!{extract_value(paragraph.text, "Дата поверки", " г.")}')
                        var['дата поверки'] = extract_value(paragraph.text, "Дата поверки", "г.")
                        paragraph.text = paragraph.text.replace(extract_value(paragraph.text, "Дата поверки", "г."), "ДАТА_ПОВЕРКИ")


                index += 1

        # Loop through paragraphs and tables
        for element in doc.element.body:
            # Проверяем, является ли элемент абзацем
            if isinstance(element, docx.oxml.text.paragraph.CT_P):
                # Получаем абзац
                paragraph = docx.text.paragraph.Paragraph(element, doc)

                # Проверяем наличие символа '#' в тексте абзаца
                if '#' in paragraph.text:
                    paragraph.text = paragraph.text.replace('#', '№')

        print(var)

        # Проверяем все ли переменные заполнены
        for i in var:
            if not i: return var[i]

        # Сохраняем файл
        doc.save(f"{main_path}/templates/{var["весы"]} {var['фиф']}.docx")
        print('Файл сохранен')


    except Exception as e:
        print('Ошибка при создании шаблона:', e)
        return True

    return False



def make_new_protocol(path,
                      name_scale,
                      num_protocol,
                      num_scale,
                      company,
                      INN,
                      legal_place,
                      ver_place,
                      temperature,
                      humidity,
                      pressure,
                      standards,
                      verificationer,
                      day,
                      month,
                      year,
                      unfit):
    print(path,
          name_scale,
          num_protocol,
          num_scale,
          company,
          INN,
          legal_place,
          ver_place,
          temperature,
          humidity,
          pressure,
          standards,
          verificationer,
          day,
          month,
          year,
          unfit)