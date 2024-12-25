import re, os, datetime, docx
from docx.shared import Pt

def repl(pattern, replace):
    #Необходимо отсечь первую страницу документа по ключевому слову "Образовательные"
    all_paras = []
    for i in doc.paragraphs:
        if 'Образовательные' in i.text:
            break
        else:
            all_paras.append(i)
    
        for paras in all_paras:
            f = re.findall(pattern, paras.text) # список строк с найденными совпадениями
            for i in f:
                if i:
                    paras.text = re.sub(i, replace, paras.text)
                    log.write(str(datetime.datetime.now()) + ' В файле ' + path + 'внесены изменения:   ' + paras.text + '\n')
                

log = open('log.txt', 'a')
log.write('\n\n\nЗАПУСК ПРИЛОЖЕНИЯ ---- ' + str(datetime.datetime.now()) + '\n\n')
#path = os.getcwd()

# ПОИСК ФАЙЛОВ ФОРМАТА *.DOCX В ТЕКУЩЕЙ И НИЖНИХ ДИРЕКТОРИЯХ
paths = []
folder = os.getcwd()

for root, dirs, files in os.walk(folder):
    for file in files:
        if file.endswith('docx') and not file.startswith('~'):
            paths.append(os.path.join(root, file))


for path in paths:
    doc = docx.Document(path)

# ИЗМЕНЕНИЕ СТИЛЯ ПО УМОЛЧАНИЮ
    style = doc.styles['Normal']
    style.font.name = 'PT Astra Serif'
    style.font.size = Pt(14)

    pattern = ['2023', 'старший преподаватель', 'майор']
    replace = ['2024', 'доцент', 'подполковник']
    for i in range(len(pattern)):
        repl(pattern[i], replace[i])

    doc.save(path + '!!!.docx')  #- ПЕРЕЗАПИСЬ ИСХОДНЫХ ФАЙЛОВ

log.close()