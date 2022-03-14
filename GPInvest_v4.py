from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import tkinter
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import pandas as pd
import locale
import datetime
from tkcalendar import DateEntry
import gspread

########################NECESSÁRIO############################
#Data
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
today = datetime.date.today()
data = today.strftime('%d %B %Y')
data2 = today.strftime('%d/%m/%Y')
att_data = today + datetime.timedelta(days = 1)
data3 = att_data.strftime('%d/%m/%Y')

#Ambiente
chave = '1ooYr6Vo5oCnBi-M-_uksu8gPj7uo6FGru6dosCUN1VY'
gc = gspread.service_account(filename='key.json')

########################FUNÇÕES###############################
def cadastro_investimento():
    sh = gc.open_by_key(chave)
    ws = sh.get_worksheet(0)
    df = pd.DataFrame(ws.get_all_records())

    #Dados de entrada
    nome = (str(vnome.get())).title()
    valor = locale.atof(vvalor.get())
    juros = locale.atof(vjuros.get())
    dias = vdias.get()
    parcela = combobox1.get()
    observação = str(vobservação.get("1.0", END))
    
    #Error campo não preenchido e valores não correspondidos
    if nome=="" or valor=="" or juros=="" or dias=="":
        return messagebox.showinfo("Error", "Preencha todos os campos para realizar\na entrada!")
    try:
        val = float(valor)
        jur = float(juros)
        dias = int(dias)
    except ValueError:
        messagebox.showinfo("Error", "As variáveis ( valor, juros e dias )\nsó aceitam números!")
    if parcela=="":
        return messagebox.showinfo("Error", "Selecione um valor válido\npara parcela!")

    #Inserção dataframe
    if int(parcela)==1:
        #Entrada
        att_data = today + datetime.timedelta(days = int(dias))
        data_pagamento1 = att_data.strftime('%d/%m/%Y')
        dados_entrada = [(data2, nome, float(valor), data_pagamento1, (float(valor)+float(juros)), 'A PAGAR', observação)]
        inputh = pd.DataFrame(dados_entrada, columns = ['D_ENTRADA', 'CLIENTE', 'VALOR', 'D_SAIDA', 'JUROS', 'STATUS', 'OBSERVAÇÃO'])
        dados_to_inserir_entrada = df.append(inputh, ignore_index=True)

    if int(parcela)>=2:
        data_att = today
        for i in range(0,(int(parcela))+1):
            if i==1:
                att_data1 = data_att + datetime.timedelta(days = int(dias))
                data_pagamento2 = att_data1.strftime('%d/%m/%Y')
                #Entrada
                dados_entrada1 = [(data2, nome, float(valor), data_pagamento2, (float(valor)+float(juros)), 'A PAGAR', observação)]
                inputh1 = pd.DataFrame(dados_entrada1, columns = ['D_ENTRADA', 'CLIENTE', 'VALOR', 'D_SAIDA', 'JUROS', 'STATUS', 'OBSERVAÇÃO'])
                dados_to_inserir_entrada = df.append(inputh1, ignore_index=True)
    
            elif i>1:
                data_att1 = data_att + datetime.timedelta(days = (int(dias)*i))
                data_att2 = data_att1.strftime('%d/%m/%Y')
                #Entrada
                dados_entrada2 = [(data2, nome, float(valor), data_att2, (float(valor)+float(juros)), 'A PAGAR', observação)]
                inputh2 = pd.DataFrame(dados_entrada2, columns = ['D_ENTRADA', 'CLIENTE', 'VALOR', 'D_SAIDA', 'JUROS', 'STATUS', 'OBSERVAÇÃO'])
                dados_to_inserir_entrada = dados_to_inserir_entrada.append(inputh2, ignore_index=True)

    #Salvar dados
    ws.update([df.columns.values.tolist()] + dados_to_inserir_entrada.values.tolist())
    #Limpar os campos
    vnome.delete(0, END)
    vvalor.delete(0, END)
    vjuros.delete(0, END)
    vdias.delete(0, END)
    combobox1.set('')
    vobservação.delete('1.0',END)
    #Mensagem salvar
    messagebox.showinfo("Sucess", "Entrada realizada!")

def filtro_movimentacao_dia():
    sh = gc.open_by_key(chave)
    ws = sh.get_worksheet(0)
    df = pd.DataFrame(ws.get_all_records())

    tree.delete(*tree.get_children())
    format_date1 = datetime.datetime.strptime((data_entry.get()), '%d/%m/%Y').date()
    data_relatorio = format_date1.strftime('%d/%m/%Y')
    df_mask=df['D_SAIDA']==data_relatorio
    filtered_df = df[df_mask]
    if filtered_df.empty==True:
        messagebox.showinfo("Error", 'Não há movimentações para o dia!')
    else:
        lista_filter = filtered_df.values.tolist()
        for i in lista_filter:
            tree.insert("", END, values=i, tag='1')

def modificações():
    def alterar():
        sh = gc.open_by_key(chave)
        ws = sh.get_worksheet(0)
        df = pd.DataFrame(ws.get_all_records())
        try:
            itemSelecionado2 = tree2.selection()[0]
            valores2 = tree2.item(itemSelecionado2, 'values')

            #Adicionando "PAGO"
            índice2 = (df.index[(df['D_ENTRADA'] == valores2[0]) & (df['CLIENTE'] == valores2[1]) & (df['VALOR'] == float(valores2[2])) & (df['D_SAIDA'] == valores2[3]) & (df['JUROS'] == float(valores2[4])) & (df['STATUS'] == valores2[5]) & (df['OBSERVAÇÃO'] == valores2[6])].tolist())
            índice2 = índice2[0]
            
            #Alterando valores juros
            valor1 = locale.atof(vvalor1.get())
            juros1 = locale.atof(vjuros1.get())
            ws.update_cell((índice2+2), 3, float(valor1))
            ws.update_cell((índice2+2), 5, (float(juros1)+float(valor1)))

            #Atualizando
            tree2.delete(*tree2.get_children())
            sh = gc.open_by_key(chave)
            ws = sh.get_worksheet(0)
            df = pd.DataFrame(ws.get_all_records())
            índice22 = (df.index[(df['CLIENTE'] == valores2[1])].tolist())

            for ss in índice22:
                cl1 = list(((df.loc[[ss]]).values)[0])
                print(cl1)
                tree2.insert("", END, values=cl1, tag='1')

            vvalor1.delete(0, END)
            vjuros1.delete(0, END)
        except:
            messagebox.showinfo("ERRO", "Selecione um registro na tabela.")

    def excluir_registro():
        sh = gc.open_by_key(chave)
        ws = sh.get_worksheet(0)
        df = pd.DataFrame(ws.get_all_records())
        try:
            itemSelecionado2 = tree2.selection()[0]
            valores2 = tree2.item(itemSelecionado2, 'values')

            #Adicionando "PAGO"
            índice2 = (df.index[(df['D_ENTRADA'] == valores2[0]) & (df['CLIENTE'] == valores2[1]) & (df['VALOR'] == float(valores2[2])) & (df['D_SAIDA'] == valores2[3]) & (df['JUROS'] == float(valores2[4])) & (df['STATUS'] == valores2[5]) & (df['OBSERVAÇÃO'] == valores2[6])].tolist())
            índice2 = índice2[0]

            #deletando
            ws.delete_rows(índice2+2)

            #Atualizando
            tree2.delete(*tree2.get_children())
            sh = gc.open_by_key(chave)
            ws = sh.get_worksheet(0)
            df = pd.DataFrame(ws.get_all_records())
            índice22 = (df.index[(df['CLIENTE'] == valores2[1])].tolist())

            for ss in índice22:
                cl1 = list(((df.loc[[ss]]).values)[0])
                tree2.insert("", END, values=cl1, tag='1')
            vvalor1.delete(0, END)
            vjuros1.delete(0, END)

        except:
            messagebox.showinfo("ERRO", "Selecione um registro na tabela.")
    
    def pagar():
        sh = gc.open_by_key(chave)
        ws = sh.get_worksheet(0)
        df = pd.DataFrame(ws.get_all_records())
        try:
            itemSelecionado2 = tree2.selection()[0]
            valores2 = tree2.item(itemSelecionado2, 'values')

            #Adicionando "PAGO"
            índice2 = (df.index[(df['D_ENTRADA'] == valores2[0]) & (df['CLIENTE'] == valores2[1]) & (df['VALOR'] == float(valores2[2])) & (df['D_SAIDA'] == valores2[3]) & (df['JUROS'] == float(valores2[4])) & (df['STATUS'] == valores2[5]) & (df['OBSERVAÇÃO'] == valores2[6])].tolist())
            índice2 = índice2[0]
            
            #Alterando valores juros
            ws.update_cell((índice2+2), 6, 'RECEBIDO')

            #Atualizando
            tree2.delete(*tree2.get_children())
            sh = gc.open_by_key(chave)
            ws = sh.get_worksheet(0)
            df = pd.DataFrame(ws.get_all_records())
            índice22 = (df.index[(df['CLIENTE'] == valores2[1])].tolist())

            for ss in índice22:
                cl1 = list(((df.loc[[ss]]).values)[0])
                print(cl1)
                tree2.insert("", END, values=cl1, tag='1')
            vvalor1.delete(0, END)
            vjuros1.delete(0, END)
        except:
            messagebox.showinfo("ERRO", "Selecione um registro na tabela.")
    
######################NOVA JANELA############################
    new_window = Tk()
    new_window.title("Aplicações gerais")
    new_window.geometry("800x500+310+109")
    new_window.configure(background='#566573')
    new_window.wm_attributes('-toolwindow', 1)
    new_window.wm_resizable(width=False, height=False)

    LabelFrame(new_window, text='Modificação', font=('Times New Roman', 12, 'bold'), background='#566573', foreground='white').place(x=4,y=7, width=794, height=491)

    Label(new_window, text="Valor", font=('Times New Roman', 12, 'bold'), background="#566573", foreground='white', anchor=CENTER).place(x=350,y=30,width=100, height=20)
    vvalor1 = Entry(new_window)
    vvalor1.place(x=250, y=55, width=300, height=25)

    Label(new_window, text="Juros", font=('Times New Roman', 12, 'bold'), background="#566573", foreground='white', anchor=CENTER).place(x=350,y=90,width=100, height=20)
    vjuros1 = Entry(new_window)
    vjuros1.place(x=250, y=115, width=300, height=25)

    btn_edt1 = Button(new_window, text="Alterar", font=('Arial', 12, 'bold'), background="#3498DB", foreground='white', command=alterar)
    btn_edt1.place(x=300, y=175, width=200, height=26)

    btn_edt2 = Button(new_window, text="Pagar", font=('Arial', 12, 'bold'), background="#1CA424", foreground='white', command=pagar)
    btn_edt2.place(x=20, y=175, width=200, height=26)

    btn_edt3 = Button(new_window, text="Excluir", font=('Arial', 12, 'bold'), background="#DB1C1C", foreground='white', command=excluir_registro)
    btn_edt3.place(x=580, y=175, width=200, height=26)

    sh = gc.open_by_key(chave)
    ws = sh.get_worksheet(0)
    df = pd.DataFrame(ws.get_all_records())

    try:
        ########################## TREEVIEW #######################
        tree2 = ttk.Treeview(new_window, selectmode='browse', column=('D_ENTRADA', 'CLIENTE', 'VALOR', 'D_SAIDA', 'JUROS', 'STATUS', 'OBSERVAÇÃO'), show='headings')
        tree2.column('D_ENTRADA', width=74, minwidth=50, anchor="center", stretch=True)
        tree2.heading('#1', text='D_ENTRADA')
        tree2.column('CLIENTE', width=290, minwidth=50, anchor="center", stretch=NO)
        tree2.heading('#2', text='CLIENTE')
        tree2.column('VALOR', width=71, minwidth=50, anchor="center", stretch=NO)
        tree2.heading('#3', text='VALOR')
        tree2.column('D_SAIDA', width=72, minwidth=50, anchor="center", stretch=NO)
        tree2.heading('#4', text='D_SAIDA')
        tree2.column('JUROS', width=56, minwidth=50, anchor="center", stretch=NO)
        tree2.heading('#5', text='JUROS')
        tree2.column('STATUS', width=66, minwidth=50, anchor="center", stretch=NO)
        tree2.heading('#6', text='STATUS')
        tree2.column('OBSERVAÇÃO', width=155, minwidth=50, anchor="center", stretch=NO)
        tree2.heading('#7', text='OBSERVAÇÃO')
        tree2.place(x=8, y=226, width=786, height=268)

        scrollbar = ttk.Scrollbar(new_window, orient=tkinter.VERTICAL, command=tree2.yview)
        tree2.configure(yscroll=scrollbar.set)
        scrollbar.place(x=788, y=227, width=4, height=266)

        itemSelecionado = tree.selection()[0]
        valores = tree.item(itemSelecionado, 'values')

        #Adicionando "PAGO"
        índice = (df.index[(df['CLIENTE'] == valores[1])].tolist())

        for s in índice:
            cl = list(((df.loc[[s]]).values)[0])
            tree2.insert("", END, values=cl, tag='1')
        
    except:
        messagebox.showinfo("ERRO", "Selecione um cliente na tabela\npara verificar movimentações!")

def relatório_geral():
    ano_gerar = str(combobox2.get())
    if ano_gerar!="":
        sh = gc.open_by_key(chave)
        ws = sh.get_worksheet(0)
        df = pd.DataFrame(ws.get_all_records())
        
        #Filters
        filter1=df['STATUS']=='RECEBIDO'
        filtered_graf2 = df[filter1]

        filtered_graf2['DATA1'] = pd.to_datetime(filtered_graf2['D_SAIDA'])
        filtered_graf2['DATA1'] = filtered_graf2['DATA1'].dt.strftime('%d/%m/%Y')
        filtered_graf2['DATA1'] = pd.to_datetime(filtered_graf2['DATA1'])
        filtered_graf2['Mes'] = filtered_graf2['DATA1'].dt.strftime('%b')
        filtered_graf2['Ano'] = filtered_graf2['DATA1'].dt.strftime('%Y')
        lista_anos_verifique = list(filtered_graf2['Ano'].unique())

        if ano_gerar in lista_anos_verifique:
            df_ano = filtered_graf2['Ano']== ano_gerar
            filtered_ano = filtered_graf2[df_ano]

            valor_mes_entrada = filtered_ano.groupby('Mes')['VALOR'].sum().reset_index().sort_values('VALOR', ascending=True)
            valor_ano_entrada = filtered_ano.groupby('Ano')['VALOR'].sum().reset_index().sort_values('VALOR', ascending=True)
            valor_mes_entrada1 = filtered_ano.groupby('Ano')['JUROS'].sum().reset_index().sort_values('JUROS', ascending=True)
            mov_anos = filtered_graf2.groupby('Ano')['JUROS'].sum().reset_index().sort_values('JUROS', ascending=True)
            #####################Gráfico1###########################
            # Define o stilo para ggplot
            plt.style.use("fivethirtyeight")

            figure = plt.figure(figsize=None, dpi=75)
            ax1 = figure.add_subplot(111)
            canva = FigureCanvasTkAgg(figure, tb4)
            canva.get_tk_widget().place(x=0, y=232, width=399, height=244)
            # Pie chart, where the slices will be ordered and plotted counter-clockwise:
            labels = 'Entrada', 'Saida'
            sizes = [sum(valor_ano_entrada['VALOR']), sum(valor_mes_entrada1['JUROS'])]
            explode = (0, 0)  # only "explode" the 2nd slice (i.e. 'Hogs')

            ax1.pie(sizes, explode=explode, autopct='%1.1f%%',
                    shadow=True, startangle=345)
            ax1.legend(labels,
                        loc="lower left",
                        title="Movimentação",
                        bbox_to_anchor=(-0.05,0,1,0.5),
                        fontsize='medium')
            ax1.set_title("Tipo de movimentação (%)", fontsize=13)
            ax1.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
            #####################Gráfico2###########################
            #Escolha o ano
            figure1 = plt.figure(figsize=None, dpi=50)
            ax_mes = figure1.add_subplot(111)
            canva1 = FigureCanvasTkAgg(figure1, tb4)
            canva1.get_tk_widget().place(x=0, y=28, width=800, height=202)
            # Dados para cada subplot
            ax_mes.bar(valor_mes_entrada['Mes'], (valor_mes_entrada['VALOR']/1000), color='#00BFFF', alpha=1, align='center')
            ax_mes.set_title("Movimentação mensal (x1000)", fontsize=20)
            ax_mes.grid(True)
            #####################Gráfico 3######################
            figure2 = plt.figure(figsize=None, dpi=50)
            ax_ano = figure2.add_subplot(111)
            canva2 = FigureCanvasTkAgg(figure2, tb4)
            canva2.get_tk_widget().place(x=401, y=232, width=399, height=244)
            # Dados para cada subplot
            ax_ano.barh(mov_anos['Ano'], (mov_anos['JUROS']/1000), color='#00FF00', alpha=1, align='center')
            ax_ano.set_title("Movimentação anual (x1000)", fontsize=20)
            ax_ano.grid(True)
        else:
            Label(tb4, background="#C6C6C6").place(x=0, y=28, width=800, height=202)
            Label(tb4, background="#C6C6C6").place(x=0, y=232, width=399, height=244)
            Label(tb4, background="#C6C6C6").place(x=401, y=232, width=399, height=244)
            messagebox.showinfo("Erro", "Não há dados para o ano selecionado!")
    else:
        messagebox.showinfo("Erro", "Selecione o ano para gerar\no relatório!")
###########################TELA PRINCIPAL########################
app = Tk()
app.title("GP INVEST")
app.geometry("800x500+280+79")
app.wm_attributes('-toolwindow', 1)
app.wm_resizable(width=False, height=False)
nb = ttk.Notebook(app)
nb.place(x=0, y=0, width=800, height=500)
###########################TELA ABAS########################
lista_combobox = [1,2,3,4,5,6,7,8,9,10,11,12]
lista_ano = ['2020', '2021', '2022', '2023', '2024', '2025', '2026', '2027', '2028', '2029', '2030', '2031', '2032', '2033', '2034', '2035', '2036', '2037', '2038', '2039', '2040']

tb2 = Frame(nb, background='#566573')
tb3 = Frame(nb, background='#566573')
tb4 = Frame(nb, background='#566573')

nb.add(tb2, text="ENTRADA")
nb.add(tb3, text="SAIDA")
nb.add(tb4, text="RELATÓRIO")

Label(tb2, text="Nome:", font=('Times New Roman', 12, 'bold'), background="#566573", foreground='black', anchor=W).place(x=247,y=20,width=100, height=20)
vnome = Entry(tb2)
vnome.place(x=250, y=45, width=300, height=25)

Label(tb2, text="Valor:", font=('Times New Roman', 12, 'bold'), background="#566573", foreground='black', anchor=W).place(x=247,y=80,width=100, height=20)
vvalor = Entry(tb2)
vvalor.place(x=250, y=105, width=300, height=25)

Label(tb2, text="Juros:", font=('Times New Roman', 12, 'bold'), background="#566573", foreground='black', anchor=W).place(x=247,y=140,width=100, height=20)
vjuros = Entry(tb2)
vjuros.place(x=250, y=165, width=300, height=25)

Label(tb2, text="Dia(s):", font=('Times New Roman', 12, 'bold'), background="#566573", foreground='black', anchor=W).place(x=247,y=200,width=100, height=20)
vdias = Entry(tb2)
vdias.place(x=250, y=225, width=300, height=25)

Label(tb2, text="N° parcela(s):", font=('Times New Roman', 12, 'bold'), background="#566573", foreground='black', anchor=W).place(x=247,y=260,width=100, height=20)
combobox1 = ttk.Combobox(tb2, values=lista_combobox, state="readonly")
combobox1.place(x=250, y=285, width=300, height=25)

Label(tb2, text="Observação:", font=('Times New Roman', 12, 'bold'), background="#566573", foreground='black', anchor=W).place(x=247,y=320,width=100, height=20)
vobservação = Text(tb2)
vobservação.place(x=250, y=345, width=300, height=60)

Label(tb4, background="#C6C6C6").place(x=0, y=28, width=800, height=202)
Label(tb4, background="#C6C6C6").place(x=0, y=232, width=399, height=244)
Label(tb4, background="#C6C6C6").place(x=401, y=232, width=399, height=244)

data_entry = DateEntry(tb3, font=('Arial', 11, 'bold'), locale='pt_br', background="#3498DB", foreground='white')
data_entry.place(x=350, y=7, width=100, height=26)

LabelFrame(tb3, text='Planilha de controle').place(x=0,y=40, width=796, height=435)
########################### BOTOES ########################
#ENTRADA
btn2 = Button(tb2, text="Salvar", font=('Arial', 12, 'bold'), background="#3498DB", foreground='white', command=cadastro_investimento)
btn2.place(x=300, y=430, width=200, height=26)

#RELATÓRIO
btn3 = Button(tb3, text="Atualizar dados", font=('Arial', 10, 'bold'), background="#3498DB", foreground='white', command=filtro_movimentacao_dia)
btn3.place(x=20, y=7, width=150, height=26)

btn4 = Button(tb3, text="Dados cliente", font=('Arial', 10, 'bold'), background="#3498DB", foreground='white', command=modificações)
btn4.place(x=600, y=7, width=180, height=26)

btn5 = Button(tb4, text="Gerar relatório", font=('Arial', 10, 'bold'), background="#3498DB", foreground='white', command=relatório_geral)
btn5.place(x=110, y=4, width=120, height=20)
########################## TREEVIEW #######################
tree = ttk.Treeview(tb3, selectmode='browse', column=('D_ENTRADA', 'CLIENTE', 'VALOR', 'D_SAIDA', 'JUROS', 'STATUS', 'OBSERVAÇÃO'), show='headings')
tree.column('D_ENTRADA', width=74, minwidth=50, anchor="center", stretch=True)
tree.heading('#1', text='D_ENTRADA')
tree.column('CLIENTE', width=290, minwidth=50, anchor="center", stretch=NO)
tree.heading('#2', text='CLIENTE')
tree.column('VALOR', width=71, minwidth=50, anchor="center", stretch=NO)
tree.heading('#3', text='VALOR')
tree.column('D_SAIDA', width=72, minwidth=50, anchor="center", stretch=NO)
tree.heading('#4', text='D_SAIDA')
tree.column('JUROS', width=56, minwidth=50, anchor="center", stretch=NO)
tree.heading('#5', text='JUROS')
tree.column('STATUS', width=66, minwidth=50, anchor="center", stretch=NO)
tree.heading('#6', text='STATUS')
tree.column('OBSERVAÇÃO', width=155, minwidth=50, anchor="center", stretch=NO)
tree.heading('#7', text='OBSERVAÇÃO')
tree.place(x=4, y=60, height=408)

scrollbar0 = ttk.Scrollbar(tb3, orient=tkinter.VERTICAL, command=tree.yview)
tree.configure(yscroll=scrollbar0.set)
scrollbar0.place(x=784, y=61, width=4, height=406)

#Combobox para ano
Label(tb4, text="Ano:", font=('Times New Roman', 12, 'bold'), background="#566573", foreground='white', anchor=W).place(x=0,y=2,width=50, height=20)
combobox2 = ttk.Combobox(tb4, values=lista_ano, state="readonly")
combobox2.place(x=40, y=4, width=60, height=20)

app.mainloop()