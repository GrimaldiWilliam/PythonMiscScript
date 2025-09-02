import os
import threading

from concurrent.futures import ThreadPoolExecutor

import pandas as pd
from sentence_transformers import SentenceTransformer, util
from typing import Literal

from transformers.data.data_collator import tolist

model = SentenceTransformer('all-MiniLM-L6-v2')

file_path1 = "C:\\Users\\AdminAkro\\Downloads\\[WMATA] Trigger Condition Faults List\\FE010019308_FaultsList.xlsx"
s_name1 = 0 #nome foglio da confrontare
col_name1 ='TriggerCondition' #nome colonna da confrontare
file_path2 = "C:\\Users\\AdminAkro\\Downloads\\[WMATA] Trigger Condition Faults List\\FE010019308_FaultsList_old.xlsx"
s_name2 = 0 #nome foglio da confrontare
col_name2 ='TriggerCondition' #nome colonna da confrontare
file_path_output = "C:\\Users\\AdminAkro\\Downloads\\out.xlsx"
col3 = []
cell3 = ""
threshold = 0.3

def are_conceptually_similar(sentence1, sentence2):
    if sentence1 == sentence2:
        return True  # Salta il confronto
    else:
        embedding1 = model.encode(sentence1, convert_to_tensor=True)
        embedding2 = model.encode(sentence2, convert_to_tensor=True)
        cosine_similarity = util.pytorch_cos_sim(embedding1, embedding2).item()
        return cosine_similarity

def read_excel_column(file_path, s_name, col_name):
    xls = pd.read_excel(file_path, s_name)
    column = xls[col_name]
    col = column.astype(str).tolist()
    return col

def read_excel_row(file_path, s_name, row_index):
    xls = pd.read_excel(file_path, s_name)
    #print(row_index)
    riga = xls.iloc[row_index]
    r = riga.astype(str).tolist()
    #print(r)
    return r

def list_of_sheets(file_path):
    xls = pd.ExcelFile(file_path, engine="openpyxl")
    sheet_names = xls.sheet_names
    return sheet_names

def write_excel_with_dataframe(file_path, df):
    try:
        if not os.path.exists(file_path):
            with pd.ExcelWriter(file_path, mode='w', engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name=sheet1)
        else:
            with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
                df.to_excel(writer, sheet_name=sheet1)
        print(f"Scritto il file {file_path} in {sheet1}")
    except Exception as e:
        print(f"Errore durante la scrittura del file {file_path}")

def cell_union(cell1, cell2):
    if are_conceptually_similar(cell1, cell2) < threshold:
        return f"{cell1} \r\n pippo \r\n {cell2}"
    else:
        return f"{cell1}"

def number_of_row(file_path, sheet):
    try:
        xls = pd.read_excel(file_path, sheet)
        return xls.shape[0]  # Restituisce il numero di righe
    except Exception as e:
        print(f"Errore durante il conteggio delle righe nel foglio {sheet}: {e}")
        return -1  # Indicatore di errore

def get_cell(row_index, col_index, sheet, file_path):
    try:
        # Leggi il file Excel
        xls = pd.read_excel(file_path, sheet_name=sheet)
        # Recupera il valore della cella
        cell_value = xls.at[row_index, col_index]
        return str(cell_value)  # Restituisce il valore della cella come stringa
    except Exception as e:
        print(f"Errore durante l'accesso alla cella nella riga {row_index}, colonna {col_index} nel foglio {sheet}: {e}")
        return None

def cut_row(row, index_row_cut=9):
    return row[:index_row_cut]

def process_sheets(sheet1, sheet2):
    already_compared = set()
    try:
        col1 = read_excel_column(file_path1, sheet1, col_name1)
        col2 = read_excel_column(file_path2, sheet2, col_name2)

        for col_index1 in range(len(col1)):
            for col_index2 in range(len(col2)):
                for row_index1 in range(1, row_number1 + 1):
                    for row_index2 in range(1, row_number2 + 1):
                        row1 = cut_row(read_excel_row(file_path1, sheet1, row_index=row_index1))
                        row2 = cut_row(read_excel_row(file_path1, sheet1, row_index=row_index2))
                        if row1 == row2:
                            cell1 = get_cell(row_index1, col_index1, sheet1, file_path1)
                            cell2 = get_cell(row_index2, col_index2, sheet2, file_path2)
                            if (cell1, cell2) not in already_compared or (cell2, cell1) not in already_compared:
                                cell3 = cell_union(cell1, cell2)
                                already_compared.add((cell1, cell2))
                                print(cell3)
                                col3.append(cell3)
                    write_excel_with_dataframe(file_path_output, pd.DataFrame({"TriggerCondition" :col3}))
    except Exception as e:
        print(f"La pagina {sheet2} nel file {file_path2} non ha una colonna chiamata {col_name2}")

if __name__ == '__main__':
    sheet_names1 = list_of_sheets(file_path1)
    sheet_names2 = list_of_sheets(file_path2)
    threads = []
    with ThreadPoolExecutor() as executor:

        for sheet1 in sheet_names1:
            row_number1 = number_of_row(file_path1, sheet1)
            for sheet2 in sheet_names2:
                row_number2 = number_of_row(file_path2, sheet2)
                if sheet1 == sheet2:
                    print(f"Analizzo il foglio {sheet1}")
                    executor.submit(process_sheets, sheet1, sheet2)
                    #thread = threading.Thread(target=process_sheets, args=(sheet1, sheet2))
                    #threads.append(thread)
                    #thread.start()
        #for thread in threads:
         #   thread.join()

print("Ho finito")