import xlsxwriter
from mimesis import Person
from mimesis import Generic
from mimesis import locales
import random


# Создаем фейковые данные российского формата
g = Generic('ru')

def get_date():
    """
    Функция, которая создает дату в нужном формате
    :return: дата в формате дд.мм.гггг
    """
    date = str(g.datetime.date()).split('-')
    day = date[2]
    month = date[1]
    year = date[0]
    return(day + '.' + month + '.' + year)

def create_xlsx(number=None):
    """
    Функция, которая генерирует таблицу excel с форматированным именем и заполняет их случайными данными.
    :param number: номер таблицы
    :return:
    """
    # Создаем рабочую книгу и добавляем рабочую таблицу
    workbook = xlsxwriter.Workbook('Employee_plans{}.xlsx'.format('_' + str(number + 1)))
    worksheet = workbook.add_worksheet()

    # Количество сотрудников определяется рандомно
    employees_count = random.randint(1, 11)

    # Количество проектов определяется рандомно
    projects_count = random.randint(1, 11)

    # Создаем фейковых пользователей
    person = Person(locales.RU)

    # пустой список под имена сотрудников
    employees = []

    # пустой список под проекты
    projects = []

    # Массив со списками данных, которыми будет заполняться excel-таблица
    data = []

    # Наполняем список сотрудников рандомными именами
    for i in range(employees_count):
        full_name = person.full_name()
        employees.append(full_name)

    # Наполняем список проектов рандомными словами
    for i in range(projects_count):
        project = g.text.word().title()
        projects.append(project)

    # Начало шапки таблицы
    table_hat = ['Название проекта', 'Руководитель', 'Дата сдачи план.', 'Дата сдачи факт.']

    # Дополняем table_hat полями для сотрудника в зависимости от количества сотрудников и добавляем в список списков
    for employee in employees:
        table_hat.append(employee + ' план.')
        table_hat.append(employee + ' факт.')
    data.append(table_hat)

    # Создаем списки с данными соответствующими каждому столбцу и добавляем их в список списков
    for project in projects:
        project_info = [
            project,
            random.choice(employees),
            get_date(),
            get_date(),
            *[random.randint(1, 11) for i in range(employees_count * 2)]
        ]
        data.append(project_info)

    # Начало таблицы с отступом в одну ячейку и одну строку
    row = 0
    col = 0

    # создаем форматы значений для ячеек таблицы
    number_format = workbook.add_format({'num_format': '#,#'})
    date_format = workbook.add_format({'num_format': 'dd.mm.yyyy'})

    # Заполняем рабочую книгу данными
    for row_data in data:
        for value in row_data:
            if row_data.index(value) == (2 or 3):
                worksheet.write(row, col, value, date_format)
            elif row_data.index(value) > 3:
                worksheet.write(row, col, value, number_format)
            else:
                worksheet.write(row, col, value)
            col += 1
        row += 1
        col = 0

    # закрываем рабочую книгу
    workbook.close()

if __name__ == '__main__':
    try:
        number_of_files = int(input('Введите количество таблиц, которые необходимо создать: '))
        for i in range(number_of_files):
            create_xlsx(i)
    except:
        print('Не получилось создать таблицы, возможно были введены некорректные значения или таблицы с такими именами уде созданы')
