import pandas as pd
from math import isnan
import json
from datetime import datetime


def check_if_is_nan(arg):
    try:
        if isnan(arg):
            return True
        else:
            return False
    except TypeError as te:
        return False


# def row_style(row):
#     cor = ['background: ' + cores[color_index]] * len(row)
#     color_index += 1
#     print(cor)
#     return row.Series(cor, row.index)


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

print(datetime.now().month)
nome_da_planilha = months[datetime.now().month]

data_frame = pd.read_excel(path, engine="odf", sheet_name=nome_da_planilha)
planilha = data_frame.values

dias = planilha[0].__len__() - 2  # -2 pra tirar o nome e o total que tem na planilha
valor_frete = 2
total_a_pagar = planilha[-1][-1]

for dia in range(1, dias + 1):
    # total_de_cada = list()
    num_pes_q_pegaram_marmita = 0
    total_pagar_por_pess = 0
    pes_q_pegaram_marmita = list()

    for pessoa in planilha:
        if not check_if_is_nan(pessoa[dia]):
            num_pes_q_pegaram_marmita += 1
            pes_q_pegaram_marmita.append(pessoa[0])

    pes_q_pegaram_marmita.pop()

    try:
        total_pagar_por_pess = valor_frete / (num_pes_q_pegaram_marmita - 1)  # -1 pq ele considera o Total Dia tbm como "pessoa"

        for pessoa in planilha:
            if pessoa[0] in pes_q_pegaram_marmita:
                pessoa[-1] += total_pagar_por_pess

        total_a_pagar += 2
                # total_de_cada.append(pessoa[-1])

    except ZeroDivisionError as zde:
        pass

    # print(f'{dia}° dia: \n\t{num_pes_q_pegaram_marmita-1}\n\t{pes_q_pegaram_marmita}\n\t{total_pagar_por_pess:.2f}\n\t{total_de_cada}')
    # print()


planilha[-1][-1] = total_a_pagar
total = list()
index = 0

for pessoa in planilha:
    print(f'{pessoa[-1]:.2f}')
    data_frame.loc[[index], 'TOTAL'] = pessoa[-1]
    index += 1
    total.append(pessoa[-1])


with pd.ExcelWriter('/home/dev2/Downloads/marmitastotal.ods', engine='odf', mode='w', if_sheet_exists='replace') as writer:
    data_frame.to_excel(writer, sheet_name=nome_da_planilha)


arquivo_json = data_frame.to_json(orient='index')
parsed = json.loads(arquivo_json)

with open('/home/dev2/Downloads/marmitastotal.json', 'w') as file:
    file.write(json.dumps(parsed, indent=4))

with open('/home/dev2/Downloads/marmitastotal.txt', 'w') as file:
    file.write(json.dumps(parsed, indent=4))


# print(data_frame)
# print(total)

