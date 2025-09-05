import os

from concurrent.futures import ThreadPoolExecutor

import pandas as pd
from sentence_transformers import SentenceTransformer, util

model = SentenceTransformer('all-MiniLM-L6-v2')

file_path1 = "C:\\Users\\AdminAkro\\Downloads\\[WMATA] Trigger Condition Faults List\\FE010019308_FaultsList.xlsx"
s_name1 = 0 #nome foglio da confrontare
col_name1 ='TriggerCondition' #nome colonna da confrontare
file_path2 = "C:\\Users\\AdminAkro\\Downloads\\[WMATA] Trigger Condition Faults List\\FE010019308_FaultsList_old.xlsx"
s_name2 = 0 #nome foglio da confrontare
col_name2 ='TriggerCondition' #nome colonna da confrontare
file_path_output = "C:\\Users\\AdminAkro\\Downloads\\out.xlsx"
col3 = []
threshold = 0.3
index_row_to_cut=9

def are_conceptually_similar(sentence1, sentence2):
    if sentence1 == sentence2:
        return True
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
    try:
        xls = pd.read_excel(file_path, s_name)
        riga = xls.iloc[row_index -1]
        r = riga.astype(str).tolist()
        return r
    except Exception as e:
        print(f"row index out of range. I'm going on with the next row. Error: {e}")
        r = ["", ""]
        return r

def list_of_sheets(file_path):
    xls = pd.ExcelFile(file_path, engine="openpyxl")
    sheet_names = xls.sheet_names
    return sheet_names

def write_excel_with_dataframe(file_path, column_to_write, sheet_name):
    try:
        df = pd.DataFrame({col_name1 :column_to_write})
        if not os.path.exists(file_path):
            with pd.ExcelWriter(file_path, mode='w', engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name=sheet_name)
                print("I'm in w")
        else:
            with pd.ExcelWriter(file_path, mode='a', engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name=sheet_name)
                print("I'm in a")
        print(f"I wrote the file {file_path} in {sheet_name}")
    except Exception as e:
        print(f"Errore writing the file {file_path}")

def cell_union(cell1, cell2):
    if are_conceptually_similar(cell1, cell2) < threshold:
        string = f"{cell1} \pippo {cell2}"
        string = string.replace("\n", " \pippo ")
    else:
        string = f"{cell1}"
        string = string.replace("\n", " \pippo ")
    return string

def number_of_row(file_path, sheet):
    try:
        xls = pd.read_excel(file_path, sheet)
        return xls.shape[0]  # Restituisce il numero di righe
    except Exception as e:
        print(f"Error counting row in the sheet {sheet}: {e}")
        return -1  # Indicatore di errore

def get_cell(row_index, col_index, sheet, file_path):
    try:
        xls = pd.read_excel(file_path, sheet_name=sheet)
        cell_value = xls.at[row_index, col_index]
        #print(str(cell_value))
        return str(cell_value)
    except Exception as e:
        print(f"Error reading cell at row {row_index}, column {col_index} in the sheet {sheet}: {e}")
      #  return None

def cut_row(row, index_row_cut):
    return row[:index_row_cut]

def there_is_a_column(sheet, col_name):
    try:
        read_excel_column(file_path1, sheet, col_name)
        return True
    except Exception as e:
        print(f"The page {sheet} doesn't have a column named {col_name}")
        return False

def process_sheets(sheet1, sheet2):
    #try:
    col1 = read_excel_column(file_path1, sheet1, col_name1)
    col2 = read_excel_column(file_path2, sheet2, col_name2)

    for col_index1 in range(len(col1)):
        for col_index2 in range(len(col2)):
            if col1[col_index1] == col2[col_index2]:
                for row_index1 in range(1, row_number1 + 1):
                    for row_index2 in range(1, row_number2 + 1):
                        row1 = cut_row(read_excel_row(file_path1, sheet1, row_index=row_index1), index_row_to_cut)
                        row2 = cut_row(read_excel_row(file_path2, sheet2, row_index=row_index2), index_row_to_cut)
                        if row1 == row2:
                            print(f"I'm in {row1} in the sheet {sheet1} who is == {row2} of the sheet {sheet2}")
                            cell1 = get_cell(row_index1, col_name1, sheet1, file_path1)
                            cell2 = get_cell(row_index2, col_name2, sheet2, file_path2)
                            cell3 = cell_union(cell1, cell2)
                            #print(cell3)
                            col3.append(cell3)
                        else:
                            a = 1
                            #print(f"{row1} in the sheet {sheet1} != {row2} in the sheet {sheet2}")
    #except Exception as e:
     #   print(f"Sheet {sheet2} in the file {file_path2} doesn't has a column named {col_name2}")

if __name__ == '__main__':
    sheet_names1 = list_of_sheets(file_path1)
    sheet_names2 = list_of_sheets(file_path2)
    threads = []

    if os.path.exists(file_path_output):
        os.remove(file_path_output)

    with ThreadPoolExecutor() as executor:
        futures = []  # Lista per tracciare i task
        for sheet_name1 in sheet_names1:
            row_number1 = number_of_row(file_path1, sheet_name1)
            for sheet_name2 in sheet_names2:
                row_number2 = number_of_row(file_path2, sheet_name2)
                if sheet_name1 == sheet_name2:
                    print(f"Analyzing {sheet_name1} and {sheet_name2}")
                    if there_is_a_column(sheet_name1, col_name1):
                        if there_is_a_column(sheet_name2, col_name2):
                            futures.append(executor.submit(process_sheets, sheet_name1, sheet_name2))

        # Attendi che tutti i task vengano completati
        for future in futures:
            future.result()  # Questo assicura che eventuali eccezioni vengano sollevate

        print(f"col3: {col3}")
        write_excel_with_dataframe(file_path_output, col3, sheet_name1)

                # process_sheets(sheet_name1, sheet_name2)

                    #thread = threading.Thread(target=process_sheets, args=(sheet1, sheet2))
                    #threads.append(thread)
                    #thread.start()
        #for thread in threads:
         #   thread.join()

print("I've finished")