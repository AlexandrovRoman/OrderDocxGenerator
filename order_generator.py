from datetime import datetime
from os import mkdir
from os.path import join
from typing import Union
from warnings import warn
from box import Box
from docxtpl import DocxTemplate
from num2t4ru import num2text
from config import tasks_folder, templates_folder, contract_number
from utils import remove_row


def create_order(context: Union[dict, Box]) -> None:
    doc = DocxTemplate(join(templates_folder, "Order.docx"))

    target_folder = join(tasks_folder, context.order_number)

    try:
        mkdir(target_folder)
    except FileExistsError:
        warn(f"Directory {target_folder} exist")

    doc.render(context)
    remove_row(doc.get_docx().tables[0], 1)
    doc.save(join(target_folder, f"ЗН №{context.order_number}.docx"))


if __name__ == '__main__':
    context = Box({"order_number": input("Введите номер ЗН: "),
                   "contract_number": contract_number,
                   "today": datetime.today().strftime("%d.%m.%Y"),
                   "total": 0,
                   "tasks": []})
    for _ in range(int(input("Введите количество задач: "))):
        task = Box({
            "task_number": input("Номер задачи: "),
            "start": input("Начало: "),
            "end": input("Конец: "),
            "time": int(input("Время выполнения: ")),
            "price": 400
        })
        task.total = task.time * task.price
        context.tasks.append(task)
        context.total += task.total
    context.str_total = num2text(context.total, ((u'рубль', u'рубля', u'рублей'), 'm'))

    create_order(context)
