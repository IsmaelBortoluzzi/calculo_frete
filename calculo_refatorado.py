import pandas as pd
from math import isnan
import json
from datetime import datetime


def total_por_pessoa(divided, divisor):
    try:
        return divided / divisor
    except ZeroDivisionError as zde:
        return 0


def somar_frete_ao_total(cell, valor):
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

nome_da_planilha = months[datetime.now().month]

df = pd.read_excel(path, engine="odf", sheet_name=nome_da_planilha)

valor_frete = 2
colunas = list(df.columns)[1:-1]
df.set_index('Nomes', inplace=True)

for dia in colunas:

    pessoas_q_pegaram_marmita = df[df[dia].notna()]
    num_pes_q_pegaram_marmita = len(pessoas_q_pegaram_marmita.loc[:, dia]) - 1
    total_pagar_por_pess = total_por_pessoa(valor_frete, num_pes_q_pegaram_marmita)
    total_dia = df.at['Total Dia', dia] + valor_frete

    df[dia] = df.loc[:, dia].apply(somar_frete_ao_total, args=(total_pagar_por_pess,))

    df.at['Total Dia', dia] = total_dia


df['Total'] = df['Total'].apply(lambda x: 0)
df.loc[:, 'Total'] = df.sum(axis=1)

df.at['Total Dia', 'Total'] = df['Total'].sum() - df.at['Total Dia', 'Total']


print(df.to_string())

arquivo_json = df.to_json(orient='index')
parsed = json.loads(arquivo_json)

with open('/home/dev2/Downloads/marmitastotal.json', 'w') as file:
    file.write(json.dumps(parsed, indent=4))

with pd.ExcelWriter('/home/dev2/Downloads/marmitastotal.ods', engine='odf', mode='w', if_sheet_exists='replace') as writer:
    df.to_excel(writer, sheet_name=nome_da_planilha)
