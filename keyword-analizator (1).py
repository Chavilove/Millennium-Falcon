import openpyxl
import re
import sys
import string


stop_list = []


def get_text():
    with open(sys.argv[1], "r", encoding="utf-8") as file:
        text = file.read()
    text = clean_text(text)
    return text


def clean_text(text): # Удаляем лишние символы    
    text = re.sub('[,.();:!?]', '', text).replace(' - ', ' ').replace(' — ', ' ')    
    return text


def get_frequencies_dict(text):
    frequencies = {}
    words_list = text.lower().split()
    for word in words_list:
        if word in stop_list:
            continue
        frequencies.setdefault(word, 0)
        frequencies[word] = frequencies[word] + 1
    return frequencies
    

def sort(dictionary):
    items_list = list(dictionary.items())
    items_list.sort(key=lambda i: i[1])
    dictionary = {}
    for item in items_list:
        key, val = item
        dictionary.update({key:val})    
    return dictionary


def write_to_excel(frequencies):
    # создаем новый excel-файл
    wb = openpyxl.Workbook()
    # добавляем новый лист
    wb.create_sheet(title = 'Частота слов', index = 0)
    # получаем лист, с которым будем работать
    sheet = wb['Частота слов']
    items_list = frequencies.items()
    for item in items_list:
        key, val = item
        sheet.append([key, val])
    wb.save('keywords.xlsx')


def main():
    text = get_text()
    raw_frequencies = get_frequencies_dict(text)
    frequencies = sort(raw_frequencies)
    write_to_excel(frequencies)
    


if __name__ == "__main__":
    main()