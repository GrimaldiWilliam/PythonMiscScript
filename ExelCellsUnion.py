import os

import pandas as pd
from sentence_transformers import SentenceTransformer, util
from typing import Literal

model = SentenceTransformer('all-MiniLM-L6-v2')

file_path1 = "C:\\Users\\AdminAkro\\Downloads\\[WMATA] Trigger Condition Faults List\\FE010019308_FaultsList.xlsx"
s_name1 = 0 #nome foglio da confrontare
col_name1 ='TriggerCondition' #nome colonna da confrontare
file_path2 = "C:\\Users\\AdminAkro\\Downloads\\[WMATA] Trigger Condition Faults List\\FE010019308_FaultsList_old.xlsx"
s_name2 = 0 #nome foglio da confrontare
col_name2 ='TriggerCondition' #nome colonna da confrontare
file_path_output = "C:\\Users\\AdminAkro\\Downloads\\out.xlsx"
col1 = ""
col2 = ""
col3 = []
threshold = 0.3

def are_conceptually_similar(sentence1, sentence2):
 embedding1 = model.encode(sentence1, convert_to_tensor=True)
 embedding2 = model.encode(sentence2, convert_to_tensor=True)
 cosine_similarity = util.pytorch_cos_sim(embedding1, embedding2).item()
 return cosine_similarity

def leggi_excel(file_path, s_name, col_name):
    xls = pd.read_excel(file_path, s_name)
    colonna = xls[col_name]
    col = colonna.astype(str).tolist()
    return col

def list_of_sheets(file_path):
    xls = pd.ExcelFile(file_path, engine="openpyxl")
    sheet_names = xls.sheet_names
    return sheet_names

if __name__ == '__main__':
    sheet_names1 = list_of_sheets(file_path1)
    sheet_names2 = list_of_sheets(file_path2)
    mode: Literal["w", "a"] = "w"
    for sheet1 in sheet_names1:
        try:
            col1 = leggi_excel(file_path1, sheet1, col_name1)
            for sheet2 in sheet_names2:
                try:
                    col2 = leggi_excel(file_path2, sheet2, col_name2)
                    print("Inizio il confronto tra le celle")
                    if sheet1 == sheet2 :
                        col3 = []
                        for i, (cell1, cell2) in enumerate(zip(col1, col2)):
                            if are_conceptually_similar(cell1, cell2) < threshold:
                                cell3 = f"{cell1} \r\n\r\n {cell2}"
                            else:
                                cell3 = f"{cell1}"
                            col3.append(cell3)
                            if i == 1:
                                mode = 'w'
                            else:
                                mode = 'a'
                            print(f"{sheet2}")
                            df1 = pd.DataFrame({"TriggerCondition" :col3})
                            with pd.ExcelWriter(file_path_output, mode='w', engine='openpyxl', if_sheet_exists='new') as writer:
                                df1.to_excel(writer, sheet_name=sheet1)
                except Exception as e:
                    print(f"La pagina {sheet2} nel file {file_path2} non ha una colonna chiamata {col_name2}")
        except Exception as e:
            print(f"La pagina {sheet1} nel file {file_path1} non ha una colonna chiamata {col_name1}")
    print("Ho finito")