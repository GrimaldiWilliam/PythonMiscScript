import pandas as pd
from sentence_transformers import SentenceTransformer, util

model = SentenceTransformer('all-MiniLM-L6-v2')

file_path = "C:\\Users\\AdminAkro\\Downloads\\WMATA Riunionone.md"
threshold = 0.7

def are_conceptually_similar(sentence1, sentence2):
 embedding1 = model.encode(sentence1, convert_to_tensor=True)
 embedding2 = model.encode(sentence2, convert_to_tensor=True)
 cosine_similarity = util.pytorch_cos_sim(embedding1, embedding2).item()
 return cosine_similarity >= threshold, cosine_similarity

def leggi_file(file_path):
    # Apri il file in modalità lettura
    with open(file_path, 'r', encoding='utf-8') as file:
        righe = file.readlines()

    righe = [riga.strip() for riga in righe]

    return righe

if __name__ == '__main__':
    righe = leggi_file(file_path)
    i = 1
    for index1 in righe:
        for index2 in righe:
            similar, score = are_conceptually_similar(index1, index2)
            if similar:
                    print(f"Le frasi sono concettualmente simili? {similar}")
                    print(f"Punteggio di similarità: {score:.2f}")
                    print("Frase1: " + index1 + "\n" +"Frase2: " + index2 + "\n")
            else:
                    with open("WMATA-Riunionone-2.md", "a") as file:
                        if i>1:
                            file_last_row = righe[i-1]
                            if index1 != file_last_row:
                                print(index1, file=file)
        i += 1

    print("Ho finito")