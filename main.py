import re
from Excel import Excel
from Email import Email
from Ozon import OzonUtils


def get_end(l):
    l = l % 100
    if (l > 20):
        return get_end(l % 10)
    if (l == 1):
        return 'а'
    elif (2 <= l <= 4):
        return 'и'
    else:
        return ''


if __name__ == '__main__':
    while True:
        user_query = str(input("Введите запрос: "))
        data, total = OzonUtils().parse_query(user_query)
        if data is not None:
            break

    data = [x for sublist in data for x in sublist]
    excel = Excel()
    excel.ws.append(("Наименование", "Описание", "Количество отзывов", "Цена со скидкой", "Цена без скидки"))
    for i in data:
        excel.ws.append(i)
    excel.ws.append([f'Общее количество товаров: {total}'])

    excel.ws.merge_cells(f'A{len(data) + 2}:E{len(data) + 2}')

    price = [int(re.sub('[^0-9]', '', i[4])) for i in data]
    all_mins = [index for index, value in enumerate(price) if value == min(price)]
    for row in all_mins:
        excel.highlight_row(row + 2, len(data[0]) + 1, 'FF0000')

    excel.adjust_column_width()
    excel.add_border(cell_range=f'A1:E{len(data) + 1}')

    excel.make_bold(len(data[0]), 1)
    excel.make_bold(len(data[0]), len(data) + 2)
    excel.wb.save("Шаблон для заполнения.xlsx")
    email = Email(['email2@gmail.com'], 'Парсинг Озона',
                  f"Запрос пользователя: {user_query}\nВ файле {len(data)+2} строк" + get_end(len(data)+2))
    # заменить ['email2@gmail.com'] на список почт получателей
    email.generate_mail()
    email.add_file("Шаблон для заполнения.xlsx")
    email.send_mail()
