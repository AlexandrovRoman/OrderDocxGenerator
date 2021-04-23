import itertools
import os
import re
from os.path import join

from config import TASKS_FOLDER
import docx

import camelot


def convert_to_int(num: str, file_path='где то'):
    mask = r"\d"
    res = re.findall(mask, num)
    try:
        return int(''.join(res))
    except ValueError:
        return int(input(f'В {file_path} был указано не число: {num}. Введите верное число: '))


def get_salary_from_docx(file_path: str):
    return sum(convert_to_int(row.cells[-1].text, file_path) for row in docx.Document(file_path).tables[0].rows[1:])


def get_salary_from_pdf(file_path: str):
    return sum(convert_to_int(salary, file_path) for salary in camelot.read_pdf(file_path)[0].df[7][1:])


def get_salary_amount(act_number: str) -> int:
    file_path = join(TASKS_FOLDER, act_number, f"Акт №{act_number}.docx")
    if os.path.exists(file_path):
        return get_salary_from_docx(file_path)

    file_path = join(TASKS_FOLDER, act_number, f"Акт №{act_number}.pdf")
    if os.path.exists(file_path):
        return get_salary_from_pdf(file_path)

    file_path = join(TASKS_FOLDER, act_number, f"ЗН №{act_number}.docx")
    if os.path.exists(file_path):
        return get_salary_from_docx(file_path) // 2

    for file_name in os.listdir(join(TASKS_FOLDER, act_number)):
        if 'акт' in file_name.lower() or 'зн' in file_name.lower():
            file_path = join(TASKS_FOLDER, act_number, file_name)
            if file_name.endswith('.pdf'):
                return get_salary_from_pdf(file_path)
            elif file_name.endswith('.docx'):
                return get_salary_from_docx(file_path)
            print(file_path)
            break

    return 0


def main():
    total = 0
    print('---------------------')
    print(f'|\tЗН\t|\tStonks\t|')
    print('---------------------')
    for i in itertools.count(1):
        folder_path = join(TASKS_FOLDER, str(i))
        if not os.path.isdir(folder_path):
            break
        salary = get_salary_amount(str(i))
        print(f'|\t{i}\t|\t{salary}\t|')
        total += salary
    print('---------------------')
    print(f'Всего заработано: {total} руб')


if __name__ == '__main__':
    main()
