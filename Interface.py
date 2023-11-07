import PySimpleGUI as sg
import Sistema_AutoXLSX_InterfaceCode as sa


config_button={ 'size': (18,2), 'font':('Helvetica', 15),}
X = 3


layout =[
    [sg.Text("Codigo",font= ('Helvetica', 20)),sg.Text(" ",size=(7, X)),sg.Input(size=(20, 5),font= ('Helvetica', 20),key='-COD-')],
    [sg.Text("Vencimento",font= ('Helvetica', 20)),sg.Text(" ",size=(0, X)),sg.Input(size=(20, 5),font= ('Helvetica', 20),key='-Vencimento-')],
    [sg.Text("Valor",font= ('Helvetica', 20)),sg.Text(" ",size=(10, X)),sg.Input(size=(20, 5),font= ('Helvetica', 20),key='-Valor-')],
    [sg.Text("Imposto",font= ('Helvetica', 20)),sg.Text(" ",size=(6, X)),sg.Input(size=(20, 5),font= ('Helvetica', 20),key='-TAX-')],
    [sg.Text(" ",size=(26, 5)),sg.Button('Enviar',**config_button,key='-Send-')]]

window = sg.Window("AutoXLSX",layout,size=(800,500))

while True:
    event, values = window.read(timeout=100)
    
    if event == sg.WINDOW_CLOSED:
        break
    
    elif event == '-Send-':
         cod=values['-COD-']
         cod = int(cod)
         vencimento=values['-Vencimento-']
         valor=values['-Valor-']
         tax=values['-TAX-']
         sa.AutoXLSX(cod,vencimento,valor,tax)
         