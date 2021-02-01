# metodo para ler arquivo word
# importar biblioteca
import os
import get_file_word as gf
import datetime
import pandas as pd
from datetime import date
import shutil
import subprocess
import re
import docx


# montagem de diretório na rede a partir de um diretório local (pastalocal)
#mkdir pastaLocal
dir_destino= "pastaLocal"
if not os.path.exists(dir_destino):
    os.makedirs(dir_destino)
    
# achando o caminho dentro da rede - MONTANDO DIRETÓRIO NA REDE
try:
    # achando o caminho dentro da rede - MONTANDO DIRETÓRIO NA REDE
    mount_smbfs smb://ensi-filer02/gerpubprop$ pastaLocal/
    ##### se der ERRO nesse comando ejetar a pasta gerpubprop$ antes e rodar o codigo novamente
except:       
    print("Erro na montagem do Diretório ejetar a pasta gerpubprop$ antes e rodar o codigo novamente")
    

# capturar data do sistema
data_atual = date.today()

#setar variáveis onde estão arquivos da triagem e tipo do arquivo word (docx)
fileExt = r".docx"
dir_origem = r"/Users/wpessoa/repositorios/portifolio_nge/pastaLocal/Coordenacao de Gestao Editorial/2021/TRIAGEM/"

# verificar se existe o diretório
print("Existe esse diretório? ",os.path.exists(dir_origem)) 

#criar variável contendo todos os arquivos com extensão DOCX no diretório especifico do dir_origem 
files_array = [_ for _ in os.listdir(dir_origem) if _.endswith(fileExt)]

print("Quantidade de arquivos do word no diretório: ",len(files_array))
# print(files_array)
print("-------------------------------------------------------------")

# rodar método get_info para ler aquivos do Word e retornar features de cada arquivo
# argumentos (diretório e array dos arquivos word)
# retorna (NºID,Qtd_caracteres,Qtd_tabela,Qtd_image,data) e gera variável triagem_docx

triagem_docx=gf.get_info(dir_origem,files_array)


# gerar um Dataframe a partir do método get_file_word (gf)
df_triagem_docx = pd.DataFrame(triagem_docx,columns=['Nun_ID',"Qtd_PG_word","Qtd_carac","Qtd_tabela","Qtd_image","Qtd_estilos", "Tamanho","Data"])

# inserindo data da triagem
df_triagem_docx['Dt_triagem'] = data_atual
df_triagem_docx["Apresentação"]= None
df_triagem_docx["pag_final"]= 0
df_triagem_docx=df_triagem_docx[['Nun_ID',"Qtd_PG_word","Qtd_carac","Qtd_tabela","Qtd_image","Qtd_estilos", "pag_final","Tamanho","Data","Dt_triagem","Apresentação"]]


# OUTPUT das variáveis do método
for i in range(0,len(df_triagem_docx.Nun_ID)):
    #print("i= ",i)
    print(files_array[i])
    print("Número de ID: ",df_triagem_docx.Nun_ID[i])
    print("Qtd de páginas: ",df_triagem_docx.Qtd_PG_word[i])
    #print("Qtd de caracteres: ",df_triagem_docx.Qtd_carac[i])
    #print("Qtd de tabelas: ",df_triagem_docx.Qtd_tabela[i])
    #print("Qtd de imagens: ",df_triagem_docx.Qtd_image[i])
    #print("Qtd de estilos no Word: ",df_triagem_docx.Qtd_estilos[i])
    #print("Data de criação do documento: ",df_triagem_docx.Data[i])
    #df_triagem_docx.loc[[i]].to_excel("/Users/wpessoa/repositorios/portifolio_nge/pastaLocal/Coordenacao de Gestao Editorial/2021/TRIAGEM/pre-triagem-ID"+df_triagem_docx.Nun_ID[i]+".xlsx",index=False,header=True )
    
    
    #for file in files_array:
    # metodo para veriricar se o texto word tem Apresentação
    try: 
        #print(files_array[i])
        doc = docx.Document(dir_origem+files_array[i])
        # ler em cada parágrafo dentro do arquivo Word
        paragra= [p.text for p in doc.paragraphs]
        for paragrafo in list(paragra):
            # str_extract_all(text_1, regex(pattern = 'f.*',ignore_case = TRUE, multiline = FALSE))
            # print(paragrafo)
            if (re.search('^Apresentação$', paragrafo, re.IGNORECASE))or(re.search('^Apresentação $', paragrafo, re.IGNORECASE)):
                print("APRESENTAÇÃO ENCONTRADA")
                #print(paragrafo)
                df_triagem_docx.loc[i, 'Apresentação']="sim"
            if (re.search('\f', paragrafo)):
                print(paragrafo)
    except:       
        print("erro leitura de arquivo")
    print("-------------------------------------------------------------")

# recuperar conteudo do arquivo Excel
excelFile_old = dir_origem+"pre-triagem"+"-"+str(data_atual.year)+".xlsx"

if os.path.exists(excelFile_old):
    #print("EXISTE ARQUIVO EXCEL")
    excelFile = pd.read_excel(excelFile_old)    
    #df_triagem_docx.append(excelFile)
    # adicionar dataframe do arquivo excel anterior no dataframe da saida do método  get_file_word
    
    excelFile = excelFile.append(df_triagem_docx)
    
    # --------------------------------
    # alterar o formato da data
    excelFile['Dt_triagem'] = pd.to_datetime(excelFile.Dt_triagem)
    excelFile['Dt_triagem'] = excelFile['Dt_triagem'].dt.strftime('%d/%m/%Y')

    # --------------------------------
    # alterar o formato da data
    excelFile['Data'] = pd.to_datetime(excelFile.Data)
    excelFile['Data'] = excelFile['Data'].dt.strftime('%d/%m/%Y')
    try:
        # deletar arquvi antigo
        # fazer metodo para inserir linha e não excluir
        os.remove(excelFile_old)
    except OSError:
        print("Oops!  O ARQUIVO EXCEL EM USO POR UM USUÁRIO. FAVOR FECHAR ESSE ARQUIVO!")
        #break
    # --------------------------------
    excelFile.to_excel(dir_origem+"pre-triagem"+"-"+str(data_atual.year)+".xlsx",index=False,header=True )
    
else:
    #print("(((  NÃO  )))  EXISTE ARQUIVO EXCEL")
    # gerar arquivo excel para arquivo tratado pela primeira vez
    df_triagem_docx.to_excel(dir_origem+"pre-triagem"+"-"+str(data_atual.year)+".xlsx",index=False,header=True )

# copiara os arquivo word para pasta de processo feito
dir_destino = dir_origem+"/FEITO/"
if len(files_array)>0:
    try:
        for linha in files_array:
            if not os.path.exists(dir_destino):
                os.makedirs(dir_destino)
            shutil.move(dir_origem+ linha, dir_destino + linha)
    except OSError:
            print("Oops!  erro no método copiara arquivo - ERRO NO S.O.")
            

#DESMONTANDO DIRETÓRIO NA REDE ====== precisa fechar todos arquivos que estiver usando a Pata que foi MONTADA
#try:
#    umount pastaLocal/
#except OSError:       
#    print("Erro na Desmontagem do Diretório")
  