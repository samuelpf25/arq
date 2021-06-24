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

st.sidebar.title('Gestão Limpeza')
pg=st.sidebar.selectbox('Selecione a Página',['Solicitações em Aberto','Solicitações a Finalizar','Consulta'])
if (pg=='Solicitações em Aberto'):
    st.title('Controle Limpeza')
    selecionado = st.selectbox('Nº da solicitação:',n_solicitacao)
    #print(nome[n_solicitacao.index(selecionado)])
    n=n_solicitacao.index(selecionado)

    #apresentar dados da solicitação
    padrao='<p style="font-family:Courier; color:Blue; font-size: 15px;"'
    infor='<p style="font-family:Courier; color:Blue; font-size: 15px;"'
    info2='<p style="font-family:Courier; color:Red; font-size: 15px;"'
    titulo='<p style="font-family:Courier; color:Blue; font-size: 20px;"'
    st.markdown(titulo+'<b>Dados da Solicitação</b></p>',unsafe_allow_html=True)
    #st.text('<p style="font-family:Courier; color:Blue; font-size: 20px;">Nome: '+ nome[n]+'</p>',unsafe_allow_html=True)

    st.markdown(padrao+'<b>Nome</b>: '+ str(nome[n])+'</p>',unsafe_allow_html=True)
    st.markdown(padrao+'<b>Telefone</b>: '+ str(telefone[n])+'</p>',unsafe_allow_html=True)
    st.markdown(padrao+'<b>Prédio</b>: '+ str(predio[n])+'</p>',unsafe_allow_html=True)
    st.markdown(padrao+'<b>Sala</b>: '+ str(sala[n])+'</p>',unsafe_allow_html=True)
    st.markdown(padrao+'<b>Data</b>: '+ str(data[n])+'</p>',unsafe_allow_html=True)
    st.markdown(padrao+'<b>Descrição</b>: '+ observacao[n]+'</p>',unsafe_allow_html=True)

    #status=st.selectbox('Selecione o Status',['Selecionar','Ciente','Não é possível atender'])
    #print(status)
    celula = sheet.find(n_solicitacao[n])
    status=st.radio('Selecione o status:',['-','Ciente','Não é possível atender'])
    botao=st.button('Registrar')
    if (botao==True):
        if (status=='Ciente'):
            #print('Selecionou ciente')
            #st.markdown(infor+'Status alterado para ' + status + '</p>',unsafe_allow_html=True)
            st.text('Status alterado para ' + status)
            sheet.update_acell('R'+str(celula.row),'VERDADEIRO')
        elif(status=='Não é possível atender'):
            #st.markdown(infor + 'Status alterado para ' + status + '</p>', unsafe_allow_html=True)
            st.text('Status alterado para ' + status)
            sheet.update_acell('R'+str(celula.row),'FALSO')
elif pg=='Solicitações a Finalizar':
    st.title('Solicitações a Finalizar')
elif pg=='Consulta':
    st.title('Consulta')
