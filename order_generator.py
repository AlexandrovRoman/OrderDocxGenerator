from datetime import datetime
from os import mkdir
from os.path import join
from typing import Union
from warnings import warn
from box import Box
from docxtpl import DocxTemplate
from num2t4ru import num2text
from config import TASKS_FOLDER, TEMPLATES_FOLDER, CONTRACT_NUMBER, CONTRACT_SIGN_AT, PRICE, USERNAME
from utils import remove_row, month2str, doc2pdf


def create_order(context: Union[dict, Box]) -> None:
    doc = DocxTemplate(join(TEMPLATES_FOLDER, "Order.docx"))

    target_folder = join(TASKS_FOLDER, context.order_number)

    try:
        mkdir(target_folder)
    except FileExistsError:
        warn(f"Directory {target_folder} exist")

    doc.render(context)
    remove_row(doc.get_docx().tables[0], 1)
    filename = f"ЗН №{context.order_number}"
    file_path = join(target_folder, f"{filename}.docx")
    doc.save(file_path)
    # doc2pdf(file_path, join(target_folder, f"{filename}.pdf"))


if __name__ == '__main__':
    context = Box({"order_number": input("Введите номер ЗН: "),
                   "contract_number": CONTRACT_NUMBER,
                   "today": datetime.today().strftime("%d.%m.%Y"),
                   "total": 0,
                   "tasks": [],
                   "username": USERNAME})
    months = ('января', 'февраля', 'марта', 'апреля', 'мая', 'июня',
              'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря')
    context.contract_sign_at = CONTRACT_SIGN_AT.strftime(f"%d {month2str(CONTRACT_SIGN_AT.month, months)} %Y г.")
    for _ in range(int(input("Введите количество задач: "))):
        task = Box({
            "task_number": input("Номер задачи: "),
            "start": input("Начало: "),
            "end": input("Конец: "),
            "time": int(input("Время выполнения: ")),
            "price": PRICE
        })
        task.total = task.time * task.price
        context.tasks.append(task)
        context.total += task.total
    context.str_total = num2text(context.total, ((u'рубль', u'рубля', u'рублей'), 'm'))

    create_order(context)
