from credentials import DB_PASSWORD, DB_USER, EMAIL, EMAIL_PASSWORD, EMAIL_ENVIO
import psycopg
import pandas as pd
import os
import email
import json
import imaplib
from redmail import EmailSender
from pathlib import Path
from app import gerar_xlsx

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
        print("Entrada/devolução")
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
                print("erro 2")
                enviar_email_erro(e, df['email_funcionario'][0])
                return
                    
        gerar_xlsx()
        enviar_email_sucesso(df['operacao'][0], df['nome_equipamento'][0], df['email_funcionario'][0])
        insert_operacao_bd(df['email_funcionario'][0], df['operacao'][0], df['quantidade'][0], df['id_equipamento']) 
        return
    
def esperar_email():
    with imaplib.IMAP4_SSL('imap.gmail.com') as mail:
        mail.login(EMAIL, EMAIL_PASSWORD)
        mail.select('inbox')
        
        print("Esperando email")
        
        with mail.idle() as idler:
            for response in idler:
                if b'EXISTS' in response:
                    print("Email novo chegou")
                    break
                
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


if __name__ == "__main__":
    esperar_email()
