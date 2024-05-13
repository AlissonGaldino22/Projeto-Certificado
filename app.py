import openpyxl
from PIL import Image, ImageDraw, ImageFont
import datetime
  
workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')  
sheet_alunos = workbook_alunos['Sheet1']  

# Iterar sobre as linhas da planilha
for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2,max_row=2)):
    # Extrair os dados de cada célula
    nome_participante = linha[0].value 
    data_final = linha[1].value 
    nome_professor = linha[2].value 

    # Verificar se as datas são do tipo datetime e formatá-las
    if isinstance(data_final, datetime.datetime):
        data_final = data_final.strftime('%d/%m/%Y')

    # Definir as fontes a serem usadas
    fonte_nome = ImageFont.truetype('./tahomabd.ttf', 90)
    fonte_geral = ImageFont.truetype('./tahoma.ttf', 80)
    fonte_data = ImageFont.truetype('./tahoma.ttf', 55)
 
    # Abrir a imagem do certificado
    image = Image.open('./certificado.png')
    desenhar = ImageDraw.Draw(image)

    # Adicionar os dados do participante ao certificado
    desenhar.text((545, 670), str(nome_participante), fill='black', font=fonte_nome)
    desenhar.text((1248, 1015), str(data_final), fill='black', font=fonte_data)
    desenhar.text((371, 1015), str(nome_professor), fill='black', font=fonte_data)
    
    # Salvar o certificado com um nome único
    image.save(f'./{indice}_certificado.png')



