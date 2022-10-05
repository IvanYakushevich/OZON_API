# библиотеки
import json
from datetime import datetime
import pandas as pd
import requests

# получаю данные по OZON API
# необходымый url из OZON API, отпределнный путь запроса
url = "http://api-seller.ozon.ru/v3/posting/fbs/list"
url = url.strip()

# ввожу ключ и id от кабинета OZON
head = {
    "Client-Id": "***ID***",
    "Api-Key": "***API_KEY***"
}
# дата конца периода, за который надо получить список отправлений (для параметра body 'to'), на текущий момент
time = datetime.utcnow().strftime('%Y-%m-%dT%H:%M:%S.%fZ'[:-4] + 'Z')
# параметры запроса в OZON_API
body = {
    "dir": "DESC",
    "filter": {
        "since": "2021-07-13T00:00:00.000Z",
        "status": str(),
        "to": time
    },
    "limit": 500,
    "offset": 0,
    "translit": True,
    "with": {
        "analytics_data": True,
        "financial_data": True
    }
}
# данные из OZON API
body = json.dumps(body)  # Надо передавать в API озон именно так
response = requests.post(url, headers=head, data=body)
# проверка корректности подключения и получения ответа от OZON_API
if response.status_code == 200:
    print('Код:200, соединение с OZON API завершено успешно!')
# сохраняю результат от OZON API в .json
with open('data_fbs.json', 'w') as convert_file:
    convert_file.write(json.dumps(response.json()))
# создаю pretty json для разбора ключей и параметров, для создания датафремов и итоговой таблицы (в итоге можно удалить)
with open('data_fbs.json', 'r') as read_file:
    object = json.load(read_file)
    pretty_object = json.dumps(object, indent=4)
with open('pretty_data_fbs.json', 'w') as convert_file:
    convert_file.write(pretty_object)

# конвертирую json в список  и словарь, без первых ключей, для упрощения развертывания (создания коллекции)
with open(r"C:\Users\DCE\Documents\OZON_TEST\data_fbs.json") as read_file:
    object = json.load(read_file)
    del_res = object['result']
    del_post = del_res['postings']

with open(r"C:\Users\DCE\Documents\OZON_TEST\dict_json_fbs.json") as f:
    data = json.load(f)
    df_all = pd.DataFrame(data)

# из списка словаря без первых ключей создаю json
with open("dict_json_fbs.json", "w", encoding="utf-8") as file:
    json.dump(del_post, file)
# создаю dataframe из json, с первыми 4я колонками
with open(r"C:\Users\DCE\Documents\OZON_TEST\dict_json_fbs.json") as f:
    data = json.load(f)
df = pd.DataFrame(columns=["Номер отправления", "Принят в обработку", "Дата отгрузки", "Статус"])
for i in range(0, len(data)):
    current_item = data[i]
    df.loc[i] = [data[i]["posting_number"], data[i]["in_process_at"], data[i]["delivering_date"], data[i]["status"]]

# создаю таблицу с другими 4я колонками раскрываю список products с ценой
products_dict = []
for i in del_post:
    #         print(i) # enter each list
    for k in i['products']:  # enter each dictionary
        products_dict.append(k)
df_products = pd.DataFrame(products_dict)  # full data with nested column
df_products_end = pd.DataFrame(
    columns=["Сумма отправления", "Наименование товара", "Итоговая стоимость товара", "Количество"])
for i in range(0, len(products_dict)):
    current_item = products_dict[i]
    df_products_end.loc[i] = [products_dict[i]["price"], products_dict[i]["name"], products_dict[i]["price"],
                              products_dict[i]["quantity"]]

# создаю таблицу с колонками про скидку
discount = []
for j in del_post:
    for k in j['financial_data']['products']:
        discount.append(k)
df_discount = pd.DataFrame(discount)
df_discount_end = pd.DataFrame(columns=["Цена товара до скидок", "Скидка %", "Скидка руб"])
for i in range(0, len(discount)):
    current_item = discount[i]
    df_discount_end.loc[i] = [discount[i]["old_price"], discount[i]["total_discount_percent"],
                              discount[i]["total_discount_value"]]

# добавляю столбец со стоимостью доставки
picking = []
for i in df_discount['picking']:
    if i is None:
        picking.append(i)
    else:
        picking.append(i['amount'])
df_picking = pd.DataFrame(picking, columns=["Cтоимость доставки"])

# создаю 2 датафрейма с колонками про город доставки и способ доставки
# город доставки
city = []
for i in df_all['analytics_data'].items():
    k = i[1:]
    h = list(k)
    for j in h:
        if j is None:
            city.append(j)
        else:
            city.append(j['city'])
df_city = pd.DataFrame(city, columns=["Город доставки"])
# способ доставки
deliv_type = []
for i in df_all['analytics_data'].items():
    k = i[1:]
    h = list(k)
    for j in h:
        if j is None:
            deliv_type.append(j)
        else:
            deliv_type.append(j['delivery_type'])
df_deliv = pd.DataFrame(deliv_type, columns=["Способ доставки"])

# соединяем 6 датафреймов df, df_products_end, df_picking, df_discount_end, df_city, df_deliv
df_ozon_fbs = pd.concat([df, df_products_end, df_picking, df_discount_end, df_city, df_deliv], axis=1)

# создаю запись в exel, записываю dataframe в excel, сохраняю excel файл в папку
ozon_fbs_data = pd.ExcelWriter('ozon_fbs_data.xlsx')
df_ozon_fbs.to_excel(ozon_fbs_data)
ozon_fbs_data.save()

print('DataFrame записан в Excel')
print('Exel таблица с данными создана, дата:', time)
