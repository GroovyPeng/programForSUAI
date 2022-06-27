from subject import Subject


# Формирование списка предметов
def subjects(period, study_days, WEEKDAYS) -> list:
    result = []
    while period.next_sibling and (period.next_sibling not in study_days) and (period.next_sibling.text not in WEEKDAYS):
        period = period.next_sibling
        result.append(Subject(period))
    return result


# Объект пары
class Period:
    def __init__(self, period, study_days, day, WEEKDAYS):
        self.day = day
        self.number = int(period.text[0])
        self.subjects = subjects(period, study_days, WEEKDAYS)
        self.count = len(self.subjects)
