
# Открываем файл для чтения
# with open('C:/ООО ПТЦ Медтехника/КОГБУЗ Детский диагностический центр/Карла Маркса 42/2024/04.04/Весы/172 ВМЭН-200 №00278 (16605-15).doc', 'r', encoding='utf-8') as file:
#     filedata = file.read()
#     print(filedata)
# Заменяем значение на новое значение
#newdata = filedata.replace('старое значение', 'новое значение')

# # Открываем файл для записи и записываем измененные данные
# with open('file.doc', 'w') as file:
#     file.write(newdata)


from docx import Document

# Путь к файлу doc
doc_file = 'C:/ООО ПТЦ Медтехника/КОГБУЗ Детский диагностический центр/Карла Маркса 42/2024/04.04/Весы/173 ВЭМ-150 №АЗ-14268 (16720-09).docx'

# Открываем файл
doc = Document(doc_file)

# Читаем содержимое файла
text = ''
for paragraph in doc.paragraphs:
    #text += paragraph.text
    print(paragraph.text)

print(text)