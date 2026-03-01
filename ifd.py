import tkinter
from tkinter import ttk
import customtkinter as  ct
import ttkbootstrap as tb
from openpyxl.workbook import Workbook

with open("Exames.txt","r",encoding="utf-8")as arquivo:
    lista_interna=arquivo.read()
    lista_interna=lista_interna.replace("\n","").strip()
    lista_exames = [item.strip().strip('"') for item in lista_interna.split(",") if item.strip()]
    arquivo.close()


















































































































with open("Fonos.txt","r",encoding="utf-8")as arquivo:
    lista_interna1=arquivo.read()
    lista_interna1=lista_interna1.replace("\n","").strip()
    lista_fonos = [item.strip().strip('"') for item in lista_interna1.split(",") if item.strip()]
    arquivo.close()

with open("Medicos.txt","r",encoding="utf-8")as arquivo:
    lista_interna2=arquivo.read()
    lista_interna2=lista_interna2.replace("\n","").strip()
    lista_medicos = [item.strip().strip('"') for item in lista_interna2.split(",") if item.strip()]
    arquivo.close()

#-------------Lista e dicionarios----------
dicionario={} #dicionario onde fica toda ficha (empresa/colaborador/ocupação/exames)
contador_exames=0 #contador
lista_contador_len=[] # index do dicionario
lista_recebe_indices_exames=[]# lista dos indices exames
tipo_de_exame=["Adm","Dem","Per","Rt","Mro","Ass"]
lista_retorna_ocupacao=[] #lista que retorna index de "tipo_de_exame

#-------------Fucoes----------
def funcao_checkboc (i):
  if lista_estado_checkbox[i].get()==1:
      lista_retorna_ocupacao.append(tipo_de_exame[i])
  elif len(lista_retorna_ocupacao)>=1 :
      lista_retorna_ocupacao.pop(len(lista_retorna_ocupacao)-2)
  else:pass
def button_next():
    lista_recebe_exames=[listbox_exames.get(idx) for idx in listbox_exames.curselection()]
    try: #tratando error quando caso esqucer de ticar checkbox, retorna "Esqueceu_de_Colocar"
        d={"empresa":empresa.get(),"colaborador":name.get(),"ocu":lista_retorna_ocupacao[0],"exames":lista_recebe_exames}
    except IndexError:
        d={"empresa":empresa.get(),"colaborador":name.get(),"ocu":"Esqueceu_de_Colocar","exames":lista_recebe_exames}

    dicionario[int(len(lista_contador_len))]=d
    contactor_exams=+1 #contador de exames
    lista_contador_len.append(contactor_exams)
    circular_prosgrebar.step(1)


    trev.insert("", "end", values=(empresa.get(), name.get(), lista_retorna_ocupacao,lista_recebe_exames))
def button_before():
    try:
        remover_ultimo_index=len(dicionario)-1
        del dicionario[remover_ultimo_index]
        remover_trev = trev.get_children()
        trev.delete(remover_trev[0])
    except KeyError:
        pass
    circular_prosgrebar.step(-1)
def acrescentar_exame_listbox():
    opcao = ct.CTkInputDialog(
        text="(Logo apos acrescentar qualquer item, é necessario reiniciar para salvar)\n""Escolha uma das opções:\n\n"
             "1 - Medico\n2 - Fono\n3 - Exames",
        title="Acrescentar Item"
    )
    resposta = opcao.get_input()
    if resposta is None:
        return
    if int(resposta) == 1:
        opcao_escolhida = ct.CTkInputDialog(
            text="Digite o Nome do(a) Medico(a): ",title="Acrescentando Medico(a)")
        adicao = opcao_escolhida.get_input()
        if adicao:
            with open("Medicos.txt", "a", encoding="utf-8") as file:
                file.write(f',"{adicao}"')
    elif int(resposta) == 2:
        opcao_escolhida = ct.CTkInputDialog(
            text="Digite o Nome do(a) Fono(a): ",title="Acrescentando Fono(a)")
        adicao = opcao_escolhida.get_input()
        if adicao:
            with open("Fonos.txt", "a", encoding="utf-8") as file:
                file.write(f',"{adicao}"')
    elif int(resposta) == 3:
        opcao_escolhida = ct.CTkInputDialog(
            text="Digite o exame: ",title="Acrescentando exame")
        adicao = opcao_escolhida.get_input()
        if adicao:
            with open("Exames.txt", "a", encoding="utf-8") as file:
                file.write(f',"{adicao}"')
def salvar_planilha():
    wb = Workbook()
    ws = wb.active
    coluna=["A","B","C","D","E","F"]
    linha=4
    ws.row_dimensions[6].height = 19
    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 33
    ws.column_dimensions['C'].width = 33
    ws.column_dimensions['D'].width = 6
    ws.column_dimensions['E'].width = 95
    ws.column_dimensions['F'].width = 10
    ws["E1"]=combobox_medicos.get()
    ws["E2"] = combobox_fonos.get()
    ws["A1"] = "Data"
    ws["B3"]="Empresa"
    ws["C3"] = "Colaborador"
    ws.merge_cells("D3:E3")
    ws["D3"] = "Exames"


    for index_excell in dicionario:
        posicao = coluna[0] + str(linha)
        ws[posicao]=date.entry.get()
        exames_separado = "+".join(dicionario[index_excell]["exames"])  #usando o join para converter a lista em ua unica str
        posicao=coluna[1]+str(linha)
        ws[posicao]=dicionario[index_excell]["empresa"]
        posicao = coluna[2] + str(linha)
        ws[posicao] = dicionario[index_excell]["colaborador"]
        posicao = coluna[3] + str(linha)
        ws[posicao] = dicionario[index_excell]["ocu"]
        posicao = coluna[4] + str(linha)
        ws[posicao] = exames_separado
        if dicionario[index_excell]["exames"][0]== "Audiometria" or dicionario[index_excell]["exames"][1]== "Audiometria":
            posicao = coluna[5] + str(linha)
            ws[posicao]=1
        linha+=1
    ws["F3"]="Audios"
    ws["F2"]="=Soma(F5:F50)"
    saves=str(date.entry.get().replace("/","-"))+".xlsx" #trocando "/" poe "-", pois o windows não aceita arquivos com "/"
    wb.save(saves)
#-------------Janela principal-----------
Interface=ct.CTk()
Interface.geometry("1200x900")
Interface.title("Automatizador Planilha")
Interface.configure(fg_color="black",)
Interface.iconbitmap("suporte.ico")

#--------------------Frames-----------
frame1=ct.CTkFrame(Interface,fg_color="#008080",width=680,height=40,border_color="grey")
frame1.place(x=5,y=125)

frame2=ct.CTkFrame(Interface,fg_color="#008B8B",width=1550,height=275,corner_radius=2,border_width=3,
                   border_color="black")
frame2.place(x=5,y=550)

frame3=ct.CTkFrame(Interface,fg_color="white",width=680,height=180,corner_radius=2,border_width=1,
                   border_color="#008B8B")
frame3.place(x=5,y=275)

#---------------------Input-----------
date=tb.DateEntry(Interface,bootstyle="success")
date.place(x=405,y=5)

empresa=ct.CTkEntry(Interface,placeholder_text="Digite o nome da Empresa:",width=450,height=35,fg_color="white",
                    text_color="black",border_color="#008B8B",placeholder_text_color="#008B8B")
empresa.place(x=5,y=45)

name=ct.CTkEntry(Interface,placeholder_text="Digite o nome do colaborador:",width=450,height=35,fg_color="white",
                    text_color="black",border_color="#008B8B",placeholder_text_color="#008B8B")
name.place(x=5,y=85)

#-----------------CheckBox------
contador_posicao=5
lista_estado_checkbox=[]

for index,x in enumerate(tipo_de_exame):
   estado_checkbox = ct.IntVar()
   lista_estado_checkbox.append(estado_checkbox)
   check=ct.CTkCheckBox(frame1,text=tipo_de_exame[index],variable=lista_estado_checkbox[index],
                        command=lambda parametro_lambda=index:funcao_checkboc(parametro_lambda),corner_radius=50,
                       hover_color="#800000",border_color="black",fg_color="green")
   check.place(x=contador_posicao,y=5)
   contador_posicao+=90

#-----------------ListBox---------
listbox_exames=tkinter.Listbox(Interface, selectmode="extended", font="Calibria", borderwidth=5, height=28, width=93,
                               yscrollcommand="0,1")
numero_d_elementos=len(lista_exames)
for y in lista_exames:
   numero_d_elementos-=1
   listbox_exames.insert(0,lista_exames[numero_d_elementos])
listbox_exames.place(x=700,y=3)

#-----------------Treeview---------
trev=ttk.Treeview(frame2,columns=("1","2","3","4"),show="headings",height=15,yscrollcommand="0,1" )
trev.column("1",minwidth=5,width=150,anchor="center",)
trev.column("2",minwidth=5,width=150,anchor="center")
trev.column("3",minwidth=5,width=150,anchor="center")
trev.column("4",minwidth=5,width=1085,anchor="w")
trev.heading("1",text="Empresa",)
trev.heading("2",text="Colaborador")
trev.heading("3",text="Ocupação")
trev.heading("4",text="Exame",anchor="w")
trev.place(x=5,y=5)

#-----------------botoes---------

botao_proximo=ct.CTkButton(Interface,text="Proximo",command=button_next,fg_color="grey",text_color="black",
                           border_color="#008B8B",bg_color="#DCDCDC",hover_color="#2F4F4F")
botao_proximo.place(x=500,y=180)

botao_anterior=ct.CTkButton(Interface,text="Anterior",command=button_before,fg_color="grey",text_color="black",
                            border_color="#008B8B",bg_color="#DCDCDC",hover_color="#2F4F4F")
botao_anterior.place(x=350,y=180)

botao_salvar=ct.CTkButton(Interface,text="Salvar",fg_color="grey",text_color="black",border_color="#008B8B",
                          bg_color="#DCDCDC",hover_color="#2F4F4F",command=salvar_planilha)
botao_salvar.place(x=10,y=230)

botao_acrescentar_exame=ct.CTkButton(frame3,text="+",width=120,height=100,corner_radius=150,border_width=3,
                                     border_color="#008B8B",text_color="white",fg_color="#008B8B",font=("Arial",100),
                                     command=acrescentar_exame_listbox)
botao_acrescentar_exame.place(x=350,y=30)


#-----------------combobox---------
combobox_medicos=ct.CTkComboBox(frame3,values=lista_medicos,width=200,border_color="#006400",fg_color="#008B8B",
                                bg_color="black",text_color="black")
combobox_medicos.place(x=15,y=15)

combobox_fonos=ct.CTkComboBox(frame3,values=lista_fonos,width=200,border_color="#006400",fg_color="#008B8B",
                              bg_color="black",text_color="black")
combobox_fonos.place(x=15,y=130)

#--------------Prosgresbar----------
circular_prosgrebar=tb.Meter(frame3,metersize=100,metertype="semi",amountused=0,amounttotal=50,stripethickness=10,
                             meterthickness=5,bootstyle="success",textright="%")
circular_prosgrebar.place(x=550,y=5)

#--------------lABEL---------------
finale=ct.CTkLabel(Interface,text="Create by:\nFelipe Grassi",font=("Arial_Black",10))
finale.place(x=590,y=520)

Interface.mainloop()