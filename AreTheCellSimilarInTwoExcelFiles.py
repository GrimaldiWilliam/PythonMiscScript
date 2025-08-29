import pandas as pd
from sentence_transformers import SentenceTransformer, util

model = SentenceTransformer('all-MiniLM-L6-v2')

file_path1 = "C:\\Users\\AdminAkro\\Downloads\\WMATA\\FE010019308_FaultsList.xlsx"
s_name1 = 0 #nome foglio da confrontare
col_name1 ='MaintDescription' #nome colonna da confrontare
file_path2 = "C:\\Users\\AdminAkro\\Downloads\\WMATA\\Cartel1.xlsx"
s_name2 = 0 #nome foglio da confrontare
col_name2 ='TriggerCondition' #nome colonna da confrontare

def are_conceptually_similar(sentence1, sentence2):
 embedding1 = model.encode(sentence1, convert_to_tensor=True)
 embedding2 = model.encode(sentence2, convert_to_tensor=True)
 cosine_similarity = util.pytorch_cos_sim(embedding1, embedding2).item()
 threshold = 0.5
 return cosine_similarity >= threshold, cosine_similarity

def leggi_excel(file_path, s_name, col_name):
    df = pd.read_excel(file_path, s_name)
    colonna = df[col_name]
    righe = colonna.astype(str).tolist()
    return righe

if __name__ == '__main__':
    numeri1 = [3,4,5,6,7,8,9,10] #celle da controllare sul primo file
    numeri2 = [1,2,3,4,5,6,7,8] #celle da controllare sul secondo file
    for index1 in numeri1:
        for index2 in numeri2:
            testi1 = leggi_excel(file_path1, index1, col_name1)
            testi2 = leggi_excel(file_path2, index2, col_name2)

            for s1 in testi1:
                for s2 in testi2:
                    similar, score = are_conceptually_similar(s1, s2)
                    if similar:
                        print(f"Le frasi sono concettualmente simili? {similar}")
                        print(f"Punteggio di similarità: {score:.2f}")
                        print("Frase1: " + s1 + "\n" +"Frase2: " + s2 + "\n")

    print("Ho finito")