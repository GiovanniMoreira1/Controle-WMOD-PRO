from credentials import DB_PASSWORD, DB_USER
import tkinter as tk
from tkinter import ttk
import psycopg


with psycopg.connect(f"dbname=postgres user={DB_USER} password={DB_PASSWORD}") as conn:
    with conn.cursor() as cur:

        cur.execute(
            "INSERT INTO funcionarios (nome, email) VALUES (%s, %s)", ("teste", "teste@durr.com.br")
        )

        cur.execute("SELECT * FROM funcionarios")
        print(cur.fetchone())

        conn.commit()


root = tk.Tk()
ttk.Button(root, text="hello world").grid()

root.mainloop()