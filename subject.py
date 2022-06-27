import re


# Функция определения периодичности предмета
def turn(input_data) -> int:
    return 1 if input_data.find('b', class_='up') else 2 if input_data.find('b', class_='dn') else 0


# Вид проведения предмета
def subj_type(input_data) -> str:
    return input_data.text if not input_data.has_attr('class') else input_data.find_next('b').text


# Форматирование названия дисциплины
def format_subject_name(subject_name: str) -> str:
    if '(' in subject_name:
        subject_name = subject_name.replace(
            subject_name[subject_name.find('('):subject_name.find(')')+1], ' ')

    if ' ' not in subject_name:
        return subject_name[:6] + '.' if len(subject_name) >= 8 else subject_name
    else:
        subject_name = subject_name.split(' ')
        result = ''
        for sub in subject_name:
            if len(sub) > 2:
                if '-' in sub:
                    sub = sub.split('-')
                    for s in sub:
                        result += s[:1].upper()
                    continue
                result += sub[:1].upper()
        return result


# Название дисциплины
def subject_name(input_data) -> str:
    return format_subject_name(''.join(re.findall(r'– ([.\S\s-]+) – ', input_data)).strip())


# Формирматирование ФИО преподавателей
def lector_name(input_data):
    if input_data != None:
        return ''.join(re.findall(r'[\w]+ \w[.]\w[.]', input_data.find('a').text))


# Объект предмета
class Subject:
    def __init__(self, period):
        self.turn = turn(period)
        self.type = subj_type(period.find('b'))
        self.name = subject_name(period.find('span').text)
        self.lector = lector_name(period.find(class_='preps'))
        self.groups = [i.text for i in period.find(
            class_='groups').find_all('a')]
