from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt
import docx
import os

from loguru import logger


def extract_value(text, start_text, end_text):
    if text.find(start_text) != -1:
        start = text.find(start_text) + len(start_text)

        end = text.find(end_text)

        if end_text == '\n':
            value = text[start:]
        else:
            value = text[start:end]

        return value.strip()


def make_new_protocol(args):
    try:
        logger.debug('start')
        
        num_protocol = args['num_protocol'].split('-',1)[0] + '-' + args['work_place'].split('-')[0].strip() + '-' + args['num_protocol'].split('-',1)[1]

        for key in args.keys():
            argument = args[key]
            logger.debug(f"{key}: {argument}")

        doc = Document(f'app/templates\\{args['scale']}.docx')
        logger.info(f'File for protocol Open: {args['scale']}.docx')

        index = 0
        for element in doc.element.body:
            # Проверяем, является ли элемент абзацем
            if isinstance(element, docx.oxml.text.paragraph.CT_P):
                # Получаем абзац
                paragraph = docx.text.paragraph.Paragraph(element, doc)

                paragraph.text = paragraph.text.replace('НОМЕР_ПРОТОКОЛА_ПЕРЕМЕННАЯ',num_protocol)\
                    .replace('НОМЕР_ВЕСОВ_ПЕРЕМЕННАЯ',args['num_scale']) \
                    .replace('КОМПАНИЯ_ПЕРЕМЕННАЯ',args['company']) \
                    .replace('НОМЕР_ИНН_ПЕРЕМЕННАЯ', args['INN']) \
                    .replace('ЮРИДИЧЕСКИЙ_АДРЕС_ПЕРЕМЕННАЯ', args['legal_address']) \
                    .replace('МЕСТО_ПОВЕРКИ_ПЕРЕМЕННАЯ', args['inspection_address']) \
                    .replace('ТЕМПЕРАТУРА_ПЕРЕМЕННАЯ', args['temperature']) \
                    .replace('ВЛАЖНОСТЬ_ПЕРЕМЕННАЯ', args['humidity']) \
                    .replace('ДАВЛЕНИЕ_ПЕРЕМЕННАЯ', args['pressure']) \
                    .replace('ЭТАЛОНЫ_ПОВЕРКИ_ПЕРЕМЕННАЯ', args['standarts']) \
                    .replace('ПОВЕРИТЕЛЬ_ПЕРЕМЕННАЯ', args['verificationer']) \
                    .replace('ДАТА_ПОВЕРКИ_ПЕРЕМЕННАЯ', args['inspection_date']) \

                if 'voltage' in args.keys():
                    paragraph.text = paragraph.text.replace('НАПРЯЖЕНИЕ_ПЕРЕМЕННАЯ', args['voltage']) \
                    .replace('ЧАСТОТА_ПЕРЕМЕННАЯ', args['frequency'])

                if 'соответствует установленным в описании типа метрологическим требованиям' in paragraph.text and args['unfit'] == True:
                    paragraph.text = paragraph.text.replace('соответствует','несоответствует')\
                    .replace('пригодно','непригодно')

                index += 1


        scale = args['scale'].rsplit(' ',1)[0]
        fif = args['scale'].rsplit(' ',1)[1]

        full = f'{args['path']}/{args['num_protocol']} {scale} №{args['num_scale']} ({fif}).docx'

        doc.save(full)
        logger.debug(full)

        if args['use_excel']:
            ...

        logger.debug('end')

        return full
    
    except Exception as e:
        logger.error(f'Error create protocol: {e}')

        logger.debug('end')

        return e