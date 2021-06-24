import gspread
from oauth2client.service_account import ServiceAccountCredentials
import streamlit as st

#https://pt.linkedin.com/pulse/manipulando-planilhas-do-google-usando-python-renan-pessoa
#documentação -> https://gspread.readthedocs.io/en/latest/api.html#models
#scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
scope = ['https://spreadsheets.google.com/feeds']
creds = ServiceAccountCredentials.from_json_keyfile_name("controle.json", scope)

cliente = gspread.authorize(creds)

#sheet = cliente.open("Ciente Limpeza").sheet1 # Open the spreadhseet
sheet=cliente.open_by_key('1J4DAYirEo3fvxMva6iDXx9mpRH6e8HTBIX2CK3svJUQ').get_worksheet(0)
dados = sheet.get_all_records()  # Get a list of all records

#nomesServ = ['Aquiles Rhuan Bandeira Neres Pinheiro','Higor Eurípedes Pimentel Fernandes de Araújo','Ismael de Souza Martins Junior','Marcelo Paulino Galhardo','Mônica Regina Vieira Santos','Paulo Cesar de Castro Filho','Samuel de Paula Faria']
#servid=easygui.choicebox("Servidor:","Escolha",nomesServ,0)

#texto=''
# for k in range(len(nomesServ)):
#     texto=texto+'\n'+str(k)+')'+nomesServ[k]
#     #print(texto)
# numero=input('Digite o número correspondente ao seu nome e pressione enter:'+ texto)
# servid=nomesServ[int(numero)]
# print('Nome selecionado:'+servid)
datando=[]
n_solicitacao=[]
nome=[]
telefone=[]
predio=[]
sala=[]
data=[]
observacao=[]

for dic in dados:
    if dic['Ciente']=='VERDADEIRO':
        print(dic['Nº da Solicitação'])
        n_solicitacao.append(dic['Nº da Solicitação'])
        nome.append(dic['Nome do Solicitante'])
        telefone.append(dic['Telefone'])
        predio.append(dic['Prédio'])
        sala.append(dic['Sala/Local'])
        data.append(dic['Data da Limpeza'])
        observacao.append(dic['Observações'])

st.title('Controle Limpeza')
selecionado = st.selectbox('Nº da solicitação:',n_solicitacao)
#print(nome[n_solicitacao.index(selecionado)])
n=n_solicitacao.index(selecionado)

#apresentar dados da solicitação
padrao='<p style="font-family:Courier; color:Blue; font-size: 20px;"'
infor='<p style="font-family:Courier; color:Blue; font-size: 20px;"'
info2='<p style="font-family:Courier; color:Red; font-size: 20px;"'
titulo='<p style="font-family:Courier; color:Blue; font-size: 25px;"'
st.markdown(titulo+'<strong>Dados da Solicitação</strong></p>',unsafe_allow_html=True)
#st.text('<p style="font-family:Courier; color:Blue; font-size: 20px;">Nome: '+ nome[n]+'</p>',unsafe_allow_html=True)

st.markdown(padrao+'<strong>Nome</strong>: '+ str(nome[n])+'</p>',unsafe_allow_html=True)
st.markdown(padrao+'<strong>Telefone</strong>: '+ str(telefone[n])+'</p>',unsafe_allow_html=True)
st.markdown(padrao+'<strong>Prédio</strong>: '+ str(predio[n])+'</p>',unsafe_allow_html=True)
st.markdown(padrao+'<strong>Sala</strong>: '+ str(sala[n])+'</p>',unsafe_allow_html=True)
st.markdown(padrao+'<strong>Data</strong>: '+ str(data[n])+'</p>',unsafe_allow_html=True)
st.markdown(padrao+'<strong>Descrição</strong>: '+ observacao[n]+'</p>',unsafe_allow_html=True)

status=st.selectbox('Selecione o Status',['Selecionar','Ciente','Não é possível atender'])
print(status)
celula = sheet.find(n_solicitacao[n])

if (status=='Ciente'):
    #print('Selecionou ciente')
    #st.markdown(infor+'Status alterado para ' + status + '</p>',unsafe_allow_html=True)
    st.text('Status alterado para ' + status)
    sheet.update_acell('R'+str(celula.row),'VERDADEIRO')
elif(status=='Não é possível atender'):
    #st.markdown(infor + 'Status alterado para ' + status + '</p>', unsafe_allow_html=True)
    st.text('Status alterado para ' + status)
    sheet.update_acell('R'+str(celula.row),'FALSO')
        #listinha.append([dic['item'],dic['atividade'],dic['descrição do termo de adesão'],dic['descrição de trabalho realizado'],str(dic['tempo / dia (min)']),dic['comprovação (se houver)'],dic['data início'],dic['data fim']])
        #datando.append(dic['data início'])
#print(listinha[-1])
#listinha.append(listinha[-1])
# dataEscolhida=easygui.choicebox("Depois de que data deseja gerar o relatório?","Escolha a data",datando,0)
# #datetime.strptime(dataEscolhida, '%d/%m/%Y').date()
# de = datetime.strptime(dataEscolhida, '%d/%m/%Y').date()  # data escolhida
# print(dataEscolhida)
# for dic in dados:
#
#     inicio=dic['data início']
#     print(inicio)
#     if inicio!='':
#         da1=datetime.strptime(inicio, '%d/%m/%Y').date() #data atual da linha
#     else:
#         inicio='01/01/1900'
#         da1=datetime.strptime(inicio, '%d/%m/%Y').date() #data atual da linha
#     if dic['Servidor']==servid and da1>=de:
#
#         listinha.append([dic['item'],dic['atividade'],dic['descrição do termo de adesão'],dic['descrição de trabalho realizado'],str(dic['tempo / dia']),dic['comprovação'],dic['data início'],dic['data fim']])
#         #datando.append(dic['data início'])

#fim de importação de dados da web
# bar = Bar('Processing', max=len(listinha)-1)
#
# #diretórios
# adesao = Document(os.path.dirname(os.path.realpath(__file__)) + "\\adesao.docx")
# atividade = Document(os.path.dirname(os.path.realpath(__file__)) + "\\atividade.docx")
#
# #importação de dados CSV
# # import csv
# # with open('dados.csv', newline='') as f:
# #     reader = csv.reader(f)
# #     lista = list(reader)
# #     #listao=lista[0].split(";")
# #
# # # #print(lista)
# # i=0
# # listinha=[]
# # for dado in lista:
# #     if (i<=(len(lista))):
# #         print(i)
# #         listinha.append(dado[0].split(";"))
# #     else:
# #         break
# #     i=i+1
#
# # print(listinha[1])
# #sublista=lista[1].split(";")
# #print(sublista)
#
# item=0
# data1_anterior=""
# data1=""
# for sublista in range(1,len(listinha)-1):
#     #sublista=sublista.split(";")
#     j=item
#     n=1
#     k=1
#     while(n<=3 and j<len(listinha)-1):
#         at=listinha[item][1]
#         #print("at: " + at)
#         descr=listinha[item][2]
#         descrr=listinha[item][3]
#         tempo=listinha[item][4]
#         comprov=listinha[item][5]
#         data1=listinha[item][6]
#         data2=listinha[item][7]
#        # print("item: " + str(item))
#        # print("j: " + str(j))
#        #print("data1: " + data1)
#        # print("data1_anterior: " + data1_anterior)
#         if (data1_anterior!=listinha[j][6] and n>1 and k==0):
#             at = ""
#             descr = ""
#             descrr = ""
#             tempo = ""
#             comprov = ""
#
#         for table in adesao.tables:
#             for row in table.rows:
#                 for cell in row.cells:
#                     for paragraph in cell.paragraphs:
#                         if 'at'+str(n) in paragraph.text:
#                             paragraph.text = paragraph.text.replace("at"+str(n), at)
#                         if 'descr'+str(n) in paragraph.text:
#                             paragraph.text = paragraph.text.replace("descr"+str(n), descr)
#                         if 'tempo'+str(n) in paragraph.text:
#                             paragraph.text = paragraph.text.replace("tempo"+str(n), tempo)
#                         if 'comprov'+str(n) in paragraph.text:
#                             paragraph.text = paragraph.text.replace("comprov"+str(n), comprov)
#                         if 'data1' in paragraph.text:
#                             paragraph.text = paragraph.text.replace("data1", data1)
#                         if 'data2' in paragraph.text:
#                             paragraph.text = paragraph.text.replace("data2", data2)
#         #sleep(0.5)
#         for table in atividade.tables:
#             for row in table.rows:
#                 for cell in row.cells:
#                     for paragraph in cell.paragraphs:
#                         if 'at'+str(n) in paragraph.text:
#                             paragraph.text = paragraph.text.replace("at"+str(n), at)
#                         if 'descrr'+str(n) in paragraph.text:
#                             paragraph.text = paragraph.text.replace("descrr"+str(n), descr)
#                         if 'data1' in paragraph.text:
#                             paragraph.text = paragraph.text.replace("data1", data1)
#                         if 'data2' in paragraph.text:
#                             paragraph.text = paragraph.text.replace("data2", data2)
#         if ((j+1)<len(listinha)):
#             j=j+1
#             k=0
#             #data1_anterior = ""
#         if(listinha[j][6]==data1):
#             item=item+1
#             data1_anterior=data1
#             k=1
#         n=n+1
#
#
#     periodo=data1.replace("/","-") + " a " + data2.replace("/","-")
#     #periodo=periodo.replace("-2021","")
#     diretorio = os.path.dirname(os.path.realpath(__file__)) + '\\Relatórios\\('+ periodo +') Termo de Adesão.docx'
#     diretorio2 = os.path.dirname(os.path.realpath(__file__)) + '\\Relatórios\\(' + periodo + ') Relatório de Atividade.docx'
#     #print(diretorio)
#     sleep(1)
#     adesao.save(diretorio)
#     atividade.save(diretorio2)
#
#     #exportando PDF
#     def convert_to_pdf(filepath:str):
#         """Save a pdf of a docx file."""
#         try:
#             word = client.DispatchEx("Word.Application")
#             target_path = filepath.replace(".docx", r".pdf")
#             word_doc = word.Documents.Open(filepath)
#             word_doc.SaveAs(target_path, FileFormat=17)
#             word_doc.Close()
#         except Exception as e:
#             raise e
#         finally:
#             word.Quit()
#
#     #print(diretorio)
#     #print(diretorio2)
#     convert_to_pdf(diretorio)
#     convert_to_pdf(diretorio2)
#
#     data1_anterior=data1
#     item = item + 1
#
#     adesao = Document(os.path.dirname(os.path.realpath(__file__)) + "\\adesao.docx")
#     atividade = Document(os.path.dirname(os.path.realpath(__file__)) + "\\atividade.docx")
#     bar.next()
# bar.finish()
