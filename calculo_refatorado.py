import pandas as pd
from math import isnan
import json
from datetime import datetime


def total_per_person(divided, divisor):
    try:
        return divided / divisor
    except ZeroDivisionError as zde:
        return 0


def add_shipping_cost_to_total(cell, valor):
    if not isnan(cell):
        cell += valor
    return cell


path = '/home/dev2/Downloads/marmitas ballke.ods'
months = {
    1: 'Janeiro',
    2: 'Fevereiro',
    3: 'Março',
    4: 'Abril',
    5: 'Maio',
    6: 'Junho',
    7: 'Julho',
    8: 'Agosto',
    9: 'Setembro',
    10: 'Outubro',
    11: 'Novembro',
    12: 'Dezembro',
}

sheet_name = months[datetime.now().month]  # watchout for the month!

df = pd.read_excel(path, engine="odf", sheet_name=sheet_name)

shipping_cost = 2  # R$2,00
columns = list(df.columns)[1:-1]
df.set_index('Nomes', inplace=True)

for day in columns:

    people_who_bought_lunch = df[df[day].notna()]
    sum_of_people_who_bought_lunch = len(people_who_bought_lunch.loc[:, day]) - 1
    total_pay_per_person = total_per_person(shipping_cost, sum_of_people_who_bought_lunch)
    day_total = df.at['Total Dia', day] + shipping_cost

    df[day] = df.loc[:, day].apply(add_shipping_cost_to_total, args=(total_pay_per_person,))

    df.at['Total Dia', day] = day_total


df['Total'] = df['Total'].apply(lambda x: 0)
df.loc[:, 'Total'] = df.sum(axis=1)

df.at['Total Dia', 'Total'] = df['Total'].sum() - df.at['Total Dia', 'Total']


print(df.to_string())

json_file = df.to_json(orient='index')
parsed = json.loads(json_file)

with open('/home/dev2/Downloads/marmitastotal.json', 'w') as file:
    file.write(json.dumps(parsed, indent=4))

with pd.ExcelWriter('/home/dev2/Downloads/marmitastotal.ods', engine='odf', mode='w', if_sheet_exists='replace') as writer:
    df.to_excel(writer, sheet_name=sheet_name)
