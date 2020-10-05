from os.path import join
from docxtpl import DocxTemplate
from config import TASKS_FOLDER, TEMPLATES_FOLDER
from utils import get_context_from_docx, remove_row


def create_act(order_number):
    target_folder = join(TASKS_FOLDER, order_number)

    try:
        context = get_context_from_docx(join(TASKS_FOLDER, order_number, f"ЗН №{order_number}.docx"))
        context.order_number = order_number
    except FileNotFoundError:
        print("Заказа с данным номером нет")
        exit()

    doc = DocxTemplate(join(TEMPLATES_FOLDER, "Act.docx"))
    doc.render(context)
    remove_row(doc.get_docx().tables[0], 1)
    doc.save(join(target_folder, f"Акт №{order_number}.docx"))


if __name__ == '__main__':
    create_act(input("Введите номер ЗН: "))
