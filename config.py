from datetime import datetime
from os import getenv
from os.path import join, split, abspath

TASKS_FOLDER = getenv('TASKS_FOLDER', 'tasks')
TEMPLATES_FOLDER = getenv('TEMPLATES_FOLDER', join(split(abspath(__file__))[0], 'templates'))
CONTRACT_NUMBER = getenv('CONTRACT_NUMBER', 17052020)
CONTRACT_SIGN_AT = datetime.strptime(getenv('CONTRACT_SIGN_AT', '17.05.2020'), '%d.%m.%Y')
PRICE = int(getenv('PRICE', 400))
