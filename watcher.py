from credentials import DB_PASSWORD, DB_USER, EMAIL, EMAIL_PASSWORD, EMAIL_ENVIO
import psycopg
import pandas as pd
import os
import time
import email
import json
import imaplib
from redmail import EmailSender
from imap_tools import MailBox
from pathlib import Path

def gerar_xlsx():
    from openpyxl import Workbook
    from openpyxl.utils.dataframe import dataframe_to_rows
    from openpyxl.worksheet.table import Table, TableStyleInfo
    from openpyxl.utils import get_column_letter
    
    sql_query = "SELECT * FROM equipamentos WHERE ativo = true"
    
    with psycopg.connect(f"dbname=postgres user={DB_USER} password={DB_PASSWORD}") as conn:
        df = pd.read_sql_query(sql_query, con=conn)
        wb = Workbook()
        ws = wb.active
        ws.cell(row=1, column=1, value="ID")
        ws.cell(row=1, column=2, value="Nome")
        ws.cell(row=1, column=3, value="Quantidade")
        ws.cell(row=1, column=4, value="Descrição")
        ws.cell(row=1, column=5, value="Fabricante")
        ws.cell(row=1, column=6, value="Nº Série/Nº Modelo")
        ws.cell(row=1, column=7, value="Localização")
        ws.cell(row=1, column=8, value="Status")
        ws.cell(row=1, column=9, value="Categoria")
        ws.cell(row=1, column=10, value="Ativo")
        ws.column_dimensions['B'].width = 56
        ws.column_dimensions['C'].width = 14
        ws.column_dimensions['D'].width = 120
        ws.column_dimensions['E'].width = 25
        ws.column_dimensions['F'].width = 28
        ws.column_dimensions['G'].width = 13
        ws.column_dimensions['H'].width = 12
        ws.column_dimensions['I'].width = 34
        ws.column_dimensions['J'].width = 13
        ws.title = "Inventário WinMOD PRO"
        
        for r in dataframe_to_rows(df, header=False, index=False):
            ws.append(r)
            
        end_cell = get_column_letter(ws.max_column) + str(ws.max_row)
        table_ref = f"A1:{end_cell}"
        
        style = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
        table = Table(displayName="WinMOD_PRO", ref=table_ref)
        table.tableStyleInfo = style
        
        ws.add_table(table)
        
        file_name = 'WinMOD PRO - Inventário.xlsx'
        wb.save(file_name) 

def alteracao_bd_json():
    with open('dados.json', 'r', encoding='utf-8-sig') as f:
        data = json.load(f)
        
    df = pd.DataFrame([data]) # cria um dataframe com os dados que vierem do email (forms -> email -> python)
    
    if(not esta_cadastrado(df['email_funcionario'][0])): # chama função com valor do e-mail presente no forms
        try: 
            with psycopg.connect(f"dbname=postgres user={DB_USER} password={DB_PASSWORD}") as conn:
                with conn.cursor() as cur:
                        cur.execute(
                            "INSERT INTO funcionarios (nome, email) VALUES (%s, %s)", (df['nome_funcionario'][0], df['email_funcionario'][0]) # insere o funcionario no banco de dados
                        )
                        conn.commit()
                        
        except Exception as e:
            print(f"Erro ao cadastrar user no banco de dados. Erro: {e}")
            return
    
    if(df['operacao'][0] == "Entrada/Devolução"): # caso da operação ser entrada
        print(f"""
            =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
              Operação: Entrada/Devolução
              
              Nome: {df['nome_funcionario'][0]}
              E-mail: {df['email_funcionario'][0]}
              Equipamento: {df['nome_equipamento'][0]}
              Quantidade: {df['quantidade'][0]}
              Armário: {df['nome_armario'][0]}
            =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
              """)
        if(not nome_ja_cadastrado(df["nome_equipamento"][0])):
            params = (df['nome_equipamento'][0], int(df['quantidade'][0]), '', '', '', df['nome_armario'][0], "Disponível", df['categoria'][0])
            try: 
                with psycopg.connect(f"dbname=postgres user={DB_USER} password={DB_PASSWORD}") as conn:
                    with conn.cursor() as cur:
                            cur.execute(
                                "INSERT INTO equipamentos (nome, quantidade, descricao, fabricante, n_serie, localizacao, status, categoria) VALUES (%s, %s, %s, %s, %s, %s, %s, %s)", (params)   
                            )
            except Exception as e:
                print("erro 1")
                enviar_email_erro(e, df['email_funcionario'][0])
                return
                    
        else:
            try: 
                with psycopg.connect(f"dbname=postgres user={DB_USER} password={DB_PASSWORD}") as conn:
                    with conn.cursor() as cur:
                            cur.execute(
                                "UPDATE equipamentos SET quantidade = quantidade + %s WHERE nome = %s", (int(df['quantidade'][0]), df['nome_equipamento'][0])
                            )
            except Exception as e:
                enviar_email_erro(e, df['email_funcionario'][0])
                return
        
    elif(df['operacao'][0] == "Retirada"): # caso da operação ser retirada
        print(f"""
            =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
              Operação: Retirada
              
              Nome: {df['nome_funcionario'][0]}
              E-mail: {df['email_funcionario'][0]}
              Equipamento: {df['nome_equipamento'][0]}
              Quantidade: {df['quantidade'][0]}
              Armário: {df['nome_armario'][0]}
            =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
              """)
        try: 
            with psycopg.connect(f"dbname=postgres user={DB_USER} password={DB_PASSWORD}") as conn:
                with conn.cursor() as cur:
                    cur.execute(
                        "SELECT quantidade FROM equipamentos WHERE id_equipamento = %s", df['id_equipamento'][0]  # query para pegar a quantidade disponivel no bd
                    )
                    quantidade_bd = cur.fetchone()
                    
                    if quantidade_bd == df['quantidade'][0]: # caso 1: quantidade do equipamento no banco de dados ser exatamente a mesma da solicitada para retirada
                        cur.execute(
                            "UPDATE equipamentos SET ativo = true WHERE id_equipamento = %s", df['id_equipamento'][0] 
                        )
                    elif quantidade_bd < df['quantidade'][0]: # caso 2: quantidade do equipamento no banco de dados ser insuficiente se comparada com a solicitada para retirada
                        raise Exception
                    elif quantidade_bd > df['quantidade'][0]: # caso 3: quantidade do equipamento no banco de dados ser maior que a solicitada para retirada
                        cur.execute(
                            "UPDATE equipamento SET quantidade = quantidade - %s", (df["quantidade"][0],)
                        )
                        
        except Exception as e:
            enviar_email_erro(e, df['email_funcionario'][0])
            return
            
        gerar_xlsx()
        enviar_email_sucesso(df['operacao'][0], df['nome_equipamento'][0], df['email_funcionario'][0])
        insert_operacao_bd(df['email_funcionario'][0], df['operacao'][0], df['quantidade'][0], df['id_equipamento']) 
        return
    
def esperar_email():
    with MailBox('imap.gmail.com').login(EMAIL, EMAIL_PASSWORD) as mailbox:
        print("Aguardando e-mail...")
        responses = mailbox.idle.wait(timeout=600) # entra em modo idle (aguarda qualquer e-mail novo)
        
        if responses:
            for resp in responses:
                if 'EXISTS' in str(resp): # recebe qualquer e-mail
                    print("Novo e-mail chegou!")
                    ler_anexo()


def ler_anexo():
    mail = imaplib.IMAP4_SSL('imap.gmail.com')
    mail.login(EMAIL, EMAIL_PASSWORD)
    mail.select('inbox')
    
    criterio_busca = f"UNSEEN FROM {EMAIL_ENVIO}"
    status, data = mail.search(None, criterio_busca)
    
    email_ids = data[0].split()
    
    for email_id in email_ids: # separa os emails em IDs
        status, msg_data = mail.fetch(email_id, '(RFC822)')
        for response_part in msg_data:
            if(isinstance(response_part, tuple)):
                msg = email.message_from_bytes(response_part[1])
            
            for part in msg.walk():
                if(part.get_content_maintype == "multipart"):
                    continue
                if(part.get_content_disposition is None):
                    continue
                
                filename = part.get_filename()
                if filename:
                    filepath = os.path.join(os.getcwd(), filename)
                    with open(filepath, "wb") as f:
                        f.write(part.get_payload(decode=True))
                        
    mail.close()
    mail.logout()
    
    alteracao_bd_json()

def esta_cadastrado(email):
    try:
        with psycopg.connect(f"dbname=postgres user={DB_USER} password={DB_PASSWORD}") as conn:
            with conn.cursor() as cur:
                cur.execute(
                    "SELECT id_funcionario FROM funcionarios WHERE email = %s", (email,) # busca o id que esta cadastrado com o email
                )    
                conn.commit()
                if(cur.fetchone() == None):
                    return False
                else:
                    return True
    except Exception as e:
        print("erro 4")
        print(email)
        print(e)
        return
            
def nome_ja_cadastrado(nome):
    try:
        with psycopg.connect(f"dbname=postgres user={DB_USER} password={DB_PASSWORD}") as conn:
            with conn.cursor() as cur:
                    cur.execute(
                        "SELECT id_equipamento FROM equipamentos WHERE nome = %s", (nome,) # busca o id que esta cadastrado com o email
                    )    
                    conn.commit()
                    if(cur.fetchone() == None):
                        return False
                    else:
                        return True
    except Exception as e:
        print("erro 3")
        print(nome)
        print(e)
        return

def insert_operacao_bd(email, operacao, quantidade, id_equip):
    try:
        with psycopg.connect(f"dbname=postgres user={DB_USER} password={DB_PASSWORD}") as conn:
            with conn.cursor() as cur:
                if(operacao == "Entrada/Devolução"):
                    operacao = "Entrada"
                    cur.execute(
                        "SELECT id_equipamento FROM equipamentos ORDER BY id_equipamento DESC LIMIT 1"
                    )
                    id_equip = cur.fetchone()[0]
                    
                cur.execute(
                    "SELECT id_funcionario FROM funcionarios WHERE email = %s", (email,)
                )
                id_func = cur.fetchone()[0]
                
                cur.execute(
                    "INSERT INTO operacoes (tipo_operacao, id_funcionario, id_equipamento, quantidade_op) VALUES (%s, %s, %s, %s)", (operacao, id_func, id_equip, quantidade)
                )
                
    
    except Exception as e:
        print("erro 10")
        print(e)
        return

def enviar_email_erro(erro, email_user): # envia e-mail para user detalhando o erro causado no processo
    print("Enviando Email: Erro")
    email = EmailSender(host="smtp.gmail.com", port=587)
    
    email.username = EMAIL
    email.password = EMAIL_PASSWORD
    email.send(
        receivers=[email_user],
        subject="Sem Sucesso",
        text=f"Ocorreu um erro ao processar sua solicitação. Erro {erro}"
    )
    
def enviar_email_sucesso(operacao, equipamento, email_user): #  envia e-mail para user de confirmação
    print("Enviando email: Sucesso")
    email = EmailSender(host="smtp.gmail.com", port=587)
    email.username = EMAIL
    email.password = EMAIL_PASSWORD
    email.send( # email com anexo do arquivo do inventario
        receivers=[EMAIL_ENVIO],
        subject="Sucesso",
        attachments={'WinMOD PRO - Inventário.xlsx':Path('WinMOD PRO - Inventário.xlsx')}
    )
    
    email.send( # email p/ user
        subject=f"Pedido concluído | WinMOD PRO",
        receivers=[email_user],
        text=f"Seu pedido de {operacao} do equipamento {equipamento} foi devidamente concluído."
    )


while True:
    if esperar_email():
        print("Processando anexo")
        ler_anexo()
        
    time.sleep(1)