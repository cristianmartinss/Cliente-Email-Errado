import pandas as pd
import pymssql
import re
data = []
#lista de emails aqui.
lista = []

for i in lista:
    conn = pymssql.connect(server='', user='', password='', database='')
    query = f"SELECT RTRIM(a1_vend),RTRIM(a1_nome),RTRIM(a1_cgc),RTRIM(a1_cod),RTRIM(a1_loja) FROM sa1010 WHERE a1_email = '{i}'" 
    cursor = conn.cursor()  
    cursor.execute(query)
    retorno = cursor.fetchall()
    if not retorno: retorno == 'No Result'
    print(retorno)
    processed_row = []
    for row in retorno:
        processed_row.append([re.sub(r'[^\x00-\x7F]+', '', value) for value in row])
    
    data.extend(processed_row)
    conn.close()
columns = ['Cod_Vend', 'Nome_Cliente', 'CNPJ_Cliente', 'Cod_Cliente', 'Loja_Cliente']
df = pd.DataFrame(data, columns=columns)
path = 'results.xlsx'
df.to_excel(path, index=False)


