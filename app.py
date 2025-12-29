from functions import *
import tkinter as tk
from tkinter import ttk, messagebox
from credentials import DB_PASSWORD, DB_USER
import psycopg
from PIL import Image, ImageTk

def menu_equipamentos():
    janela = tk.Toplevel()
    janela.title("Equipamentos")
    janela.geometry("900x600")
    janela.configure(bg="#f4f6f9")

    container = tk.Frame(janela, bg="#f4f6f9")
    container.place(relx=0.5, rely=0.5, anchor="center")

    img = Image.open("assets/icon.png")
    img = img.resize((120, 120))
    logo = ImageTk.PhotoImage(img)
    logo_label = tk.Label(container, image=logo)
    logo_label.image = logo
    logo_label.pack(side="top", pady=20)

    titulo = tk.Label(
        container,
        text="Gerenciamento de Equipamentos",
        font=("Segoe UI", 20, "bold"),
        fg="#00468e",
        bg="#f4f6f9"
    )
    titulo.pack(pady=(0, 30))

    def criar_botao(texto, comando):
        return tk.Button(
            container,
            text=texto,
            width=30,
            height=2,
            font=("Segoe UI", 13, "bold"),
            bg="#00468e",
            fg="white",
            activebackground="#003366",
            activeforeground="white",
            relief="flat",
            cursor="hand2",
            command=comando
        )

    create_btn = criar_botao(
        "Cadastrar Equipamento",
        lambda: inserir_equipamento_janela(janela)
    )
    create_btn.pack(pady=10)

    update_btn = criar_botao(
        "Atualizar Equipamento",
        lambda: atualizar_equipamento_janela(janela)
    )
    update_btn.pack(pady=10)

    delete_btn = criar_botao(
        "Excluir Equipamento",
        lambda: excluir_equipamento_janela(janela)
    )
    delete_btn.pack(pady=10)

    
def inserir_equipamento(params):
    with psycopg.connect(f"dbname=postgres user={DB_USER} password={DB_PASSWORD}") as conn:
        with conn.cursor() as cur:
            try:
                print(params)
                cur.execute(
                    "INSERT INTO equipamentos (nome, quantidade, descricao, fabricante, n_serie, localizacao, status, categoria) VALUES (%s, %s, %s, %s, %s, %s, %s, %s)", (params)   
                )
                cur.execute(
                    f"SELECT id_equipamento FROM equipamentos WHERE nome = %s AND n_serie = %s", (params[0], params[4])
                )
                gerar_log("Entrada", params[1], params[4])
                conn.commit()
            except Exception as e:
                return e
        
def inserir_equipamento_tk(params, janela, janela_pai):
    produto_criado = True
    try:
        e = inserir_equipamento(params)
        if(e) != None:
            produto_criado = False
            messagebox.showerror("Erro", f"Ocorreu um erro ao Cadastrar Produto: {e}. Tente novamente!")
        janela.destroy()
    finally:
        if(produto_criado):
            messagebox.showinfo("Sucesso!", "Produto cadastrado com sucesso!")
        janela_pai.deiconify()
        janela.destroy() 
        
def inserir_equipamento_api(params):
    try:
       inserir_equipamento(params)
    except Exception as e:
       with open("logs.txt", "a") as f:
           f.write("Erro: " + e) 

           
    
def inserir_equipamento_janela(janela_pai):
    janela_pai.withdraw()

    janela = tk.Toplevel()
    janela.title("Inserir Equipamento")
    janela.geometry("1000x780")
    janela.configure(bg="#f4f6f8")
    janela.focus_set()

    name_var = tk.StringVar()
    qntd_var = tk.StringVar()
    fabr_var = tk.StringVar()
    loc_var = tk.StringVar()

    main_frame = tk.Frame(janela, bg="#f4f6f8")
    main_frame.pack(expand=True, fill="both")

    header = tk.Frame(main_frame, bg="#00468e", height=80)
    header.pack(fill="x")

    img = Image.open("assets/icon.png")
    img = img.resize((120, 120))
    logo = ImageTk.PhotoImage(img)
    logo_label = tk.Label(header, image=logo, bg="#FFFFFF")
    logo_label.image = logo
    logo_label.pack(side="left")

    tk.Label(
        header,
        text="Cadastro de Equipamento",
        bg="#00468e",
        fg="white",
        font=("Segoe UI", 20, "bold")
    ).pack(pady=22)

    form_card = tk.Frame(
        main_frame,
        bg="white",
        highlightbackground="#d0d7de",
        highlightthickness=1
    )
    form_card.pack(pady=30, padx=50)

    label_style = {
        "bg": "white",
        "font": ("Segoe UI", 12, "bold"),
        "anchor": "e"
    }

    entry_style = {
        "font": ("Segoe UI", 11),
        "width": 35,
        "relief": "solid",
        "bd": 1
    }

    tk.Label(form_card, text="Nome:", **label_style).grid(row=0, column=0, padx=20, pady=12, sticky="e")
    name_entry = tk.Entry(form_card, textvariable=name_var, **entry_style)
    name_entry.grid(row=0, column=1, pady=12)

    tk.Label(form_card, text="Quantidade:", **label_style).grid(row=1, column=0, padx=20, pady=12, sticky="e")
    qntd_entry = tk.Entry(form_card, textvariable=qntd_var, **entry_style)
    qntd_entry.grid(row=1, column=1, pady=12)

    tk.Label(form_card, text="Fabricante:", **label_style).grid(row=2, column=0, padx=20, pady=12, sticky="e")
    fabr_entry = tk.Entry(form_card, textvariable=fabr_var, **entry_style)
    fabr_entry.grid(row=2, column=1, pady=12)

    tk.Label(form_card, text="Nº Série:", **label_style).grid(row=3, column=0, padx=20, pady=12, sticky="e")
    combo_nserie = ttk.Combobox(
        form_card,
        values=["Não se aplica"],
        width=35
    )
    combo_nserie.grid(row=3, column=1, pady=12)

    tk.Label(form_card, text="Localização:", **label_style).grid(row=4, column=0, padx=20, pady=12, sticky="e")
    loc_entry = tk.Entry(form_card, textvariable=loc_var, **entry_style)
    loc_entry.grid(row=4, column=1, pady=12)

    tk.Label(form_card, text="Status:", **label_style).grid(row=5, column=0, padx=20, pady=12, sticky="e")
    combo_status = ttk.Combobox(
        form_card,
        state="readonly",
        values=["Disponível", "Emprestado", "Com Falha"],
        width=35
    )
    combo_status.grid(row=5, column=1, pady=12)

    tk.Label(form_card, text="Categoria:", **label_style).grid(row=6, column=0, padx=20, pady=12, sticky="e")
    combo_categ = ttk.Combobox(
        form_card,
        state="readonly",
        values=[
            "Equipamentos Principais",
            "Equipamentos de Rede",
            "Dispositivos de Campo",
            "Componentes elétricos e de painel",
            "Cabos",
            "Manuais e Guias"
        ],
        width=35
    )
    combo_categ.grid(row=6, column=1, pady=12)

    tk.Label(form_card, text="Descrição:", **label_style).grid(
        row=7, column=0, padx=20, pady=12, sticky="ne"
    )

    desc_text = tk.Text(
        form_card,
        width=35,
        height=7,
        font=("Segoe UI", 10),
        relief="solid",
        bd=1,
        wrap="word"
    )
    desc_text.grid(row=7, column=1, pady=12)

    sub_btn = tk.Button(
        main_frame,
        text="Cadastrar Equipamento",
        font=("Segoe UI", 12, "bold"),
        bg="#00468e",
        fg="white",
        width=36,
        height=2,
        bd=0,
        cursor="hand2",
        activebackground="#003366",
        command=lambda: inserir_equipamento_tk(
            (
                name_entry.get().capitalize(),
                qntd_entry.get(),
                desc_text.get("1.0", "end").strip().capitalize(),
                fabr_entry.get(),
                None if combo_nserie.get() == "Não se aplica" else combo_nserie.get(),
                loc_entry.get(),
                combo_status.get(),
                combo_categ.get()
            ),
            janela,
            janela_pai
        )
    )
    sub_btn.pack(pady=10)




def listar_equipamentos():
    with psycopg.connect(f"dbname=postgres user={DB_USER} password={DB_PASSWORD}") as conn:
        with conn.cursor() as cur:
            try:
                cur.execute(
                    "SELECT * FROM equipamentos WHERE ativo = true",   
                )
                conn.commit()

                return cur.fetchall()
            
            except Exception as e:
                return e
            
def excluir_equipamento_tk(tabela):
    selecionados = tabela.selection()
    iid = selecionados[0]

    if not selecionados:
        return
    try:
        excluir_equipamento(iid)
    except Exception as e:
        print(e)
    finally:
        if(tabela.exists(iid)):
            tabela.delete(iid)

def excluir_equipamento(id_equipamento):
    with psycopg.connect(f"dbname=postgres user={DB_USER} password={DB_PASSWORD}") as conn:
        with conn.cursor() as cur:
            try:
                print(id_equipamento)
                cur.execute(
                    f"SELECT quantidade FROM equipamentos WHERE id_equipamento = %s", (id_equipamento,)
                )       
                gerar_log("Retirada", int(cur.fetchone()), id_equipamento)

                cur.execute(
                    f"UPDATE equipamentos SET ativo = FALSE WHERE id_equipamento = %s", (id_equipamento,),
                )
                conn.commit()
            except Exception as e:
                print(e)
                return e    
                
'''
def excluir_equipamento_janela(janela_pai):
    janela_pai.withdraw()
    janela = tk.Toplevel(janela_pai)
    janela.wm_title("Excluir Equipamento")
    janela.geometry("900x600")
    janela.focus_set()
    
    
    frame = tk.Frame(janela)
    frame.pack(fill=tk.X, expand=True, padx=10, pady=40)
    
    label = tk.Label(janela, text="Equipamentos")
    label.place(x=10, y=10)
    
    scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL)
    lista = tk.Listbox(frame, yscrollcommand=scrollbar.set, height=50, width=50, font='calibre', selectmode='single')
    
    rows = listar_equipamentos()
    
    # rows[i][1] = nome, rows[i][3] = fabricante, rows[i][4] = n_serie, rows[i][5] = localizacao, rows[i][6] = status
    for i in range(len(rows)):
        lista.insert(i, f"Nome: {rows[i][1]} # Quantidade: {rows[i][2]} # Fabricante: {rows[i][4]} # Nº Série: {rows[i][5]} # Localização: {rows[i][6]}")

    delete_btn = tk.Button(janela, width=30, text="Excluir", command = lambda: excluir_equipamento_tk(lista))
    delete_btn.place(x=350, y=567)
            
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    lista.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
'''

def excluir_equipamento_janela(janela_pai):
    janela_pai.withdraw()

    janela = tk.Toplevel(janela_pai)
    janela.title("Excluir Equipamento")
    janela.geometry("1000x650")
    janela.configure(bg="#f4f6f8")
    janela.focus_set()

    main_frame = tk.Frame(janela, bg="#f4f6f8")
    main_frame.pack(expand=True, fill="both")

    header = tk.Frame(main_frame, bg="#00468e", height=80)
    header.pack(fill="x")
    header.pack_propagate(False)

    tk.Label(
        header,
        text="Excluir Equipamento",
        bg="#00468e",
        fg="white",
        font=("Segoe UI", 20, "bold")
    ).pack(side="left", padx=30)

    card = tk.Frame(
        main_frame,
        bg="white",
        highlightbackground="#d0d7de",
        highlightthickness=1
    )
    card.pack(padx=40, pady=30, fill="both", expand=True)
    content = tk.Frame(card, bg="white")
    content.pack(padx=20, pady=20, fill="both", expand=True)

    tk.Label(
        content,
        text="Selecione um equipamento para excluir",
        bg="white",
        font=("Segoe UI", 12, "bold")
    ).pack(anchor="w", pady=(0, 10))

    table_frame = tk.Frame(content, bg="white")
    table_frame.pack(fill="both", expand=True)

    scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL)

    colunas = ("nome", "quantidade", "fabricante", "n_serie", "localizacao", "status")

    tabela = ttk.Treeview(
        table_frame,
        columns=colunas,
        show="headings",
        yscrollcommand=scrollbar.set,
        selectmode="browse",
        height=12
    )

    scrollbar.config(command=tabela.yview)

    tabela.heading("nome", text="Nome")
    tabela.heading("quantidade", text="Qtd")
    tabela.heading("fabricante", text="Fabricante")
    tabela.heading("n_serie", text="Nº Série")
    tabela.heading("localizacao", text="Localização")
    tabela.heading("status", text="Status")

    tabela.column("nome", width=220)
    tabela.column("quantidade", width=60, anchor="center")
    tabela.column("fabricante", width=160)
    tabela.column("n_serie", width=130)
    tabela.column("localizacao", width=120, anchor="center")
    tabela.column("status", width=120, anchor="center")

    rows = listar_equipamentos()
    for row in rows:
        tabela.insert(
            "",
            tk.END,
            iid=row[0],
            values=(
                row[1],
                row[2],
                row[4],
                row[5],
                row[6],
                row[7]
            )
        )

    tabela.pack(side=tk.LEFT, fill="both", expand=True)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    actions = tk.Frame(content, bg="white")
    actions.pack(fill="x", pady=15)

    delete_btn = tk.Button(
        actions,
        text="🗑  Excluir Equipamento",
        font=("Segoe UI", 11, "bold"),
        bg="#c0392b",
        fg="white",
        width=28,
        height=2,
        bd=0,
        cursor="hand2",
        activebackground="#922b21",
        activeforeground="white",
        command=lambda: excluir_equipamento_tk(tabela)
    )
    delete_btn.pack(side="right")





def atualizar_equipamento_janela(janela_pai):
    print()

def gerar_log(operacao, quantidade, id_equip):
    with psycopg.connect(f"dbname=postgres user={DB_USER} password={DB_PASSWORD}") as conn:
        with conn.cursor() as cur:
            try:
                cur.execute(
                    "INSERT INTO operacoes (tipo_operacao, quantidade_op, id_funcionario, id_equipamento) VALUES (%s, %s, %s, %s)", (operacao, quantidade, 1, id_equip)
                )
                conn.commit()
            except Exception as e:
                print(e)
                return e
    

root = tk.Tk()
root.title("Sistema de Controle")
root.geometry("900x600")
root.configure(bg="#f4f6f8")

main_frame = tk.Frame(root, bg="#f4f6f8")
main_frame.pack(expand=True, fill="both")

header = tk.Frame(main_frame, bg="#00468e", height=90)
header.pack(fill="x")

header.pack_propagate(False)
img = Image.open("assets/icon.png")
img = img.resize((120, 120))
logo = ImageTk.PhotoImage(img)
logo_label = tk.Label(header, image=logo, bg="#FFFFFF")
logo_label.image = logo
logo_label.pack(side="right")

title = tk.Label(
    header,
    text="Controle WinMOD PRO",
    bg="#00468e",
    fg="white",
    font=("Segoe UI", 22, "bold")
)
title.pack(side="left", padx=30)

card = tk.Frame(
    main_frame,
    bg="white",
    highlightbackground="#d0d7de",
    highlightthickness=1
)
card.pack(pady=80, padx=200)

card_content = tk.Frame(card, bg="white")
card_content.pack(padx=40, pady=40)

subtitle = tk.Label(
    card_content,
    text="Menu Principal",
    bg="white",
    font=("Segoe UI", 16, "bold")
)
subtitle.pack(pady=(0, 30))

sub_btn = tk.Button(
    card_content,
    text="⚙️  Menu de Equipamentos",
    font=("Segoe UI", 14, "bold"),
    bg="#00468e",
    fg="white",
    width=28,
    height=2,
    bd=0,
    cursor="hand2",
    activebackground="#003366",
    activeforeground="white",
    command=menu_equipamentos
)
sub_btn.pack()

footer = tk.Label(
    main_frame,
    text="© Sistema Interno • Controle WinMOD PRO",
    bg="#f4f6f8",
    fg="#6b7280",
    font=("Segoe UI", 9)
)
footer.pack(side="bottom", pady=10)

root.mainloop()
