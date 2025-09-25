import pandas as pd
import glob
import os

folder_path = "src\\data\\raw"

excel_files = glob.glob(os.path.join(folder_path, '*.xlsx'))

if not excel_files:
    print("Nenhum arquivo compátivel encontrado")
else:
    dfs = []

    for excel_file in excel_files:
        try:
            dt_temp = pd.read_excel(excel_file)
            file_name = os.path.basename(excel_file)

            dt_temp['Arquivo'] = os.path.basename(excel_file)

            if 'brasil' in file_name.lower():
                dt_temp['location'] = 'br'
            elif 'france' in file_name.lower():
                dt_temp['location'] = 'fr'
            elif 'italian' in file_name.lower():
                dt_temp['location'] =  'it'

            dt_temp['campaign'] = dt_temp['utm_link'].str.extract(r'utm_campaign=(.*)')

            dfs.append(dt_temp) #guardar dados tratados
        except Exception as e:
            print(f"Erro ao ler o arquivo {excel_file} - erro: {e}")

        
if dfs:
    result = pd.concat(dfs, ignore_index=True)

    output_file = os.path.join('src', 'data', 'ready', 'clean.xlsx')
    
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter') #Configuração do motor de escrita no Excel

    result.to_excel(writer, index=False)

    writer._save()
else:
    print("Nenhum dado para ser salvo!")



