from functions import *
import tkinter as tk
from tkinter import ttk, messagebox
from credentials import DB_PASSWORD, DB_USER
import psycopg


def menu_equipamentos():
    janela = tk.Toplevel()
    janela.wm_title("Equipamentos")
    janela.geometry("600x400")
    sub_btn=tk.Button(janela,text = 'Cadastrar Equipamento', command = inserir_equipamento_janela)
    sub_btn.grid(row=2, column=1)
    
def inserir_equipamento(params):
    with psycopg.connect(f"dbname=postgres user={DB_USER} password={DB_PASSWORD}") as conn:
        with conn.cursor() as cur:
            cur.execute(
                "INSERT INTO equipamentos (nome, descricao, fabricante, n_serie, localizacao, status, categoria) VALUES (%s, %s, %s, %s, %s, %s, %s)", (params)   
            )
            
            conn.commit()
    
def inserir_equipamento_janela():
    janela = tk.Toplevel()
    janela.wm_title("Inserir Equipamento")
    janela.geometry("600x400")
    
    name_var=tk.StringVar()
    desc_var=tk.StringVar()
    fabr_var=tk.StringVar()
    nserie_var=tk.StringVar()
    loc_var=tk.StringVar()
    status_var=tk.StringVar()
    categ_var=tk.StringVar()
    
    name_label = tk.Label(janela, text = 'Username', font=('calibre',10, 'bold'))
    name_entry = tk.Entry(janela,textvariable = name_var, font=('calibre',10,'normal'))
    desc_label = tk.Label(janela, text = 'Username', font=('calibre',10, 'bold'))
    desc_entry = tk.Entry(janela,textvariable = desc_var, font=('calibre',10,'normal'))
    fabr_label = tk.Label(janela, text = 'Username', font=('calibre',10, 'bold'))
    fabr_entry = tk.Entry(janela,textvariable = fabr_var, font=('calibre',10,'normal'))
    nserie_label = tk.Label(janela, text = 'Username', font=('calibre',10, 'bold'))
    nserie_entry = tk.Entry(janela,textvariable = nserie_var, font=('calibre',10,'normal'))
    loc_label = tk.Label(janela, text = 'Username', font=('calibre',10, 'bold'))
    loc_entry = tk.Entry(janela,textvariable = loc_var, font=('calibre',10,'normal'))
    status_label = tk.Label(janela, text = 'Username', font=('calibre',10, 'bold'))
    status_entry = tk.Entry(janela,textvariable = status_var, font=('calibre',10,'normal'))
    categ_label = tk.Label(janela, text = 'Username', font=('calibre',10, 'bold'))
    categ_entry = tk.Entry(janela,textvariable = categ_var, font=('calibre',10,'normal'))
    
    params = name_entry.get(), desc_entry.get(), fabr_entry.get(), nserie_entry.get(), loc_entry.get(), status_entry.get(), categ_entry.get()
    
    sub_btn=tk.Button(janela,text = 'Cadastrar', command = lambda: inserir_equipamento(params))
    sub_btn.grid(row=15, column=1)
    
    name_label.grid(row=0, column=0)
    name_entry.grid(row=0, column=1)
    desc_label.grid(row=1, column=0)
    desc_entry.grid(row=1, column=1)
    fabr_label.grid(row=2, column=0)
    fabr_entry.grid(row=2, column=1)
    nserie_label.grid(row=3, column=0)
    nserie_entry.grid(row=3, column=1)
    loc_label.grid(row=4, column=0)
    loc_entry.grid(row=4, column=1)
    status_label.grid(row=5, column=0)
    combo_status = ttk.Combobox(
        state='readonly',
        values=['Equipamentos Principais', 'Equipamentos de Rede', 'Dispositivos de Campo', 'Componentes elétricos e de painel', 'Cabos', 'Manuais e Guias'],
    )
    combo_status.place(x=500, y=500)
    categ_label.grid(row=6, column=0)
    categ_entry.grid(row=6, column=1)

 
root=tk.Tk()
root.geometry("1000x700")



sub_btn=tk.Button(root,text = 'Menu Equipamentos', command = menu_equipamentos)
sub_btn.grid(row=2,column=1)
 
root.mainloop()
