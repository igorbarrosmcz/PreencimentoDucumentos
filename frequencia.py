from docxtpl import DocxTemplate
import pandas as pd
import os

# Caminhos dos modelos
MODELO_30 = "MODELO1.docx"
MODELO_31 = "MODELO2.docx"
PLANILHA = "professores.xlsx"
PASTA_SAIDA = "frequencias_geradas"
os.makedirs(PASTA_SAIDA, exist_ok=True)

# Leitura da planilha (espera uma coluna 'dia')
df = pd.read_excel(PLANILHA)

# Validação
if "dia" not in df.columns:
    raise ValueError("A planilha deve conter uma coluna chamada 'dia' com valores 30 ou 31.")

# Loop para gerar documentos
for _, row in df.iterrows():
    dia = int(row["dia"])
    if dia == 30:
        modelo = MODELO_30
    elif dia == 31:
        modelo = MODELO_31
    else:
        print(f"⚠️ Dia inválido para {row['Nome']}: {dia}")
        continue

    doc = DocxTemplate(modelo)
    context = {
        "NOME": row["Nome"],
        "CPF": row["CPF"],
        "CURSO": row["Curso"]
    }
    doc.render(context)

    nome_limpo = row["Nome"].replace(" ", "_")
    caminho_saida = os.path.join(PASTA_SAIDA, f"frequencia_{nome_limpo}_dia{dia}.docx")
    doc.save(caminho_saida)

print(f"Documentos gerados na pasta: {PASTA_SAIDA}")
