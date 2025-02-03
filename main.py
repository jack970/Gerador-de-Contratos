import tkinter as tk
import requests
import locale
from tkinter import ttk
from datetime import date
from docx import Document

def busca_cep(cep):
    req = requests.get(f'https://viacep.com.br/ws/{cep}/json/')
    response = req.json()
    return response

def busca_cep_locador(args):
    cep = cep_locador.get()
    cep_buscado = busca_cep(cep.replace('-', ''))
    bairro_locador_var.set(cep_buscado['bairro'])
    rua_locador_var.set(cep_buscado['logradouro'])
    cidade_locador_var.set(cep_buscado['localidade'])
    print(cep_buscado)

def busca_cep_imovel(args):
    cep = cep_imovel_var.get()
    cep_buscado = busca_cep(cep.replace('-', ''))
    bairro_imovel_var.set(cep_buscado['bairro'])
    rua_imovel_var.set(cep_buscado['logradouro'])
    cidade_imovel_var.set(cep_buscado['localidade'])
    print(cep_buscado)

def submit_form():
    dados = {
        "NOME_LOCADOR": nome_locador_entry.get(),
        "ESTADO_LOCADOR": estado_civil_locador_var.get(),
        "CPF_LOCADOR": locador_cpf_entry.get(),
        "RG_LOCADOR": locador_rg_entry.get(),
        "PROFISSAO_LOCADOR": profissao_locador.get(),
        "RUA_LOCADOR": rua_locador_var.get(),
        "NUMERO_LOCADOR": numero_locador.get(),
        "BAIRRO_LOCADOR": bairro_locador_var.get(),
        "CIDADE_LOCADOR": cidade_locador_var.get(),
        "CEP_LOCADOR": cep_locador_var.get(),
        ###### LOCATÁRIO
        "NOME_LOCATARIO": nome_locatario.get(),
        "ESTADO_LOCATARIO": estado_civil_locatario_var.get(),
        "PROFISSAO_LOCATARIO": profissao_locatario.get(),
        "RG_LOCATARIO": rg_locatario.get(),
        "CPF_LOCATARIO": cpf_locatario.get(),
        ###### IMOVEL
        "BAIRRO_IMOVEL": bairro_imovel_var.get(),
        "NUMERO_IMOVEL": numero_imovel.get(),
        "RUA_IMOVEL": rua_imovel_var.get(),
        "CIDADE_IMOVEL": cidade_imovel_var.get(),
        "CEP_IMOVEL": cep_imovel_var.get(),
        "PRAZO_LOCACAO": prazo_locacao.get(),
        "VALOR_IMOVEL": valor_imovel.get(),
        "DIA_PAGAMENTO": dia_pagamento.get(),
        ##### COMARCA
        "CIDADE_COMARCA": cidade_comarca.get(),
        "UF_COMARCA": uf_comarca.get(),
        "DATA_ATUAL": data_atual_comarca.get()
    }
    print("Dados do Formulário:", dados)
    preencher_template(dados)
    status_label.config(text="Formulário enviado com sucesso!")

def preencher_template(dados):
    template_path = "template_contrato_residencial.docx"
    output_docx = "contrato_residencial_preenchido.docx"
    output_pdf = "contrato_locacao.pdf"
    
    doc = Document(template_path)
    for p in doc.paragraphs:
        for chave, valor in dados.items():
            if f"{{{chave}}}" in p.text:
                print(chave, valor)
                p.text = p.text.replace(f"{{{chave}}}", valor)
    
    doc.save(output_docx)
    # converter_para_pdf(output_docx, output_pdf)

def carrega_data_atual():
    locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
    data_formatada = date.today().strftime('%d de %B de %Y')
    # Definir um valor inicial no Entry
    data_atual_comarca.set(data_formatada)


estado_civil_opcoes = ["solteiro", "casado"]

root = tk.Tk()
root.title("Gerador de Contrato Residencial/Comercial")

#-- stringVAr---
tipo_contrato = tk.StringVar()

#locador
estado_civil_locador_var = tk.StringVar()
cep_locador_var = tk.StringVar()
bairro_locador_var = tk.StringVar()
rua_locador_var = tk.StringVar()
cidade_locador_var = tk.StringVar()

#locatario
estado_civil_locatario_var = tk.StringVar()

#imovel
cep_imovel_var = tk.StringVar()
bairro_imovel_var = tk.StringVar()
rua_imovel_var = tk.StringVar()
cidade_imovel_var = tk.StringVar()

#comarca
data_atual_comarca = tk.StringVar()
#####


# Combobox no topo
ttk.Label(root, text="Tipo de Contrato:", textvariable=tipo_contrato).pack(anchor="n", padx=10, pady=2)
escolha_contrato = ttk.Combobox(root, values=["Residencial", "Comercial"], state='readonly')
escolha_contrato.current(0)
escolha_contrato.pack(fill="x", padx=10, pady=2)


# Locador
frame = tk.LabelFrame(root, text="Locador")
frame.pack(padx=10, pady=10, fill="both")

# Configurar as colunas do frame para serem redimensionáveis
frame.grid_columnconfigure(0, weight=1)
frame.grid_columnconfigure(1, weight=1)
frame.grid_columnconfigure(2, weight=1)


# Widgets dentro do frame
ttk.Label(frame, text="Nome:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
nome_locador_entry = ttk.Entry(frame)
nome_locador_entry.grid(row=1, column=0, padx=5, pady=5, sticky="ew")

tk.Label(frame, text="CPF:").grid(row=0, column=1, padx=5, pady=5, sticky="w")
locador_cpf_entry = tk.Entry(frame)
locador_cpf_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

tk.Label(frame, text="RG:").grid(row=0, column=2, padx=5, pady=5, sticky="w")
locador_rg_entry = tk.Entry(frame)
locador_rg_entry.grid(row=1, column=2, padx=5, pady=5, sticky="ew")

ttk.Label(frame, text="Estado Civil:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
estado_civil_locador = ttk.Combobox(frame, values=estado_civil_opcoes, state='readonly', textvariable=estado_civil_locador_var)
estado_civil_locador.current(0)
estado_civil_locador.grid(row=3, column=0, padx=5, pady=5, sticky="ew")

ttk.Label(frame, text="Profissão:").grid(row=2, column=1, padx=10, pady=2, sticky="w")
profissao_locador = ttk.Entry(frame)
profissao_locador.grid(row=3, column=1, padx=5, pady=5, sticky="ew")

ttk.Label(frame, text="CEP:").grid(row=2, column=2, padx=10, pady=2, sticky="w")
cep_locador = ttk.Entry(frame, textvariable=cep_locador_var)
cep_locador.bind("<Return>", busca_cep_locador)
cep_locador.grid(row=3, column=2, padx=5, pady=5, sticky="ew")

ttk.Label(frame, text="Bairro:").grid(row=4, column=0, padx=10, pady=2, sticky="w")
ttk.Entry(frame, textvariable=bairro_locador_var).grid(row=5, column=0, padx=5, pady=5, sticky="ew")

ttk.Label(frame, text="Cidade:").grid(row=4, column=1, padx=10, pady=2, sticky="w")
ttk.Entry(frame, textvariable=cidade_locador_var).grid(row=5, column=1, padx=5, pady=5, sticky="ew")

ttk.Label(frame, text="Número:").grid(row=4, column=2, padx=10, pady=2, sticky="w")
numero_locador = ttk.Entry(frame)
numero_locador.grid(row=5, column=2, padx=5, pady=5, sticky="ew")

#Locatario
frame1 = tk.LabelFrame(root, text="Locatário")
frame1.pack(padx=10, pady=10, fill="x")
# Configurar as colunas do frame para serem redimensionáveis
frame1.grid_columnconfigure(0, weight=1)
frame1.grid_columnconfigure(1, weight=1)
frame1.grid_columnconfigure(2, weight=1)


ttk.Label(frame1, text="Nome:").grid(row=0, column=0, padx=10, pady=2, sticky="w")
nome_locatario = ttk.Entry(frame1)
nome_locatario.grid(row=1, column=0, padx=5, pady=5, sticky="ew")

ttk.Label(frame1, text="CPF:").grid(row=0, column=1, padx=10, pady=2, sticky="w")
cpf_locatario = ttk.Entry(frame1)
cpf_locatario.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

ttk.Label(frame1, text="RG:").grid(row=0, column=2, padx=10, pady=2, sticky="w")
rg_locatario = ttk.Entry(frame1)
rg_locatario.grid(row=1, column=2, padx=5, pady=5, sticky="ew")

ttk.Label(frame1, text="Estado Civil").grid(row=2, column=0, padx=5, pady=5, sticky="w")
escolha_contrato = ttk.Combobox(frame1, values=estado_civil_opcoes, state='readonly')
escolha_contrato.current(0)
escolha_contrato.grid(row=3, column=0, padx=5, pady=5, sticky="ew")

ttk.Label(frame1, text="Profissão:").grid(row=2, column=1, padx=10, pady=2, sticky="w")
profissao_locatario = ttk.Entry(frame1)
profissao_locatario.grid(row=3, column=1, padx=5, pady=5, sticky="ew")

# IMÓVEL
frame2 = tk.LabelFrame(root, text="Imóvel")
frame2.pack(padx=10, pady=10, fill="both")

# Configurar as colunas do frame para serem redimensionáveis
frame2.grid_columnconfigure(0, weight=1)
frame2.grid_columnconfigure(1, weight=1)
frame2.grid_columnconfigure(2, weight=1)

ttk.Label(frame2, text="Valor do Imóvel:").grid(row=0, column=0, padx=10, pady=2, sticky="w")
valor_imovel = ttk.Entry(frame2)
valor_imovel.grid(row=1, column=0, padx=5, pady=5, sticky="ew")

ttk.Label(frame2, text="Prazo de Locação:").grid(row=0, column=1, padx=10, pady=2, sticky="w")
prazo_locacao = ttk.Entry(frame2)
prazo_locacao.grid(row=1, column=1, padx=10, pady=2, sticky="ew")

ttk.Label(frame2, text="Dia Pagamento:").grid(row=0, column=2, padx=10, pady=2, sticky="w")
dia_pagamento = ttk.Entry(frame2)
dia_pagamento.grid(row=1, column=2, padx=10, pady=2, sticky="ew")

ttk.Label(frame2, text="CEP:").grid(row=2, column=0, padx=10, pady=2, sticky="w")
ttk.Entry(frame2,  textvariable=cep_imovel_var).grid(row=3, column=0, padx=10, pady=2, sticky="ew")

ttk.Label(frame2, text="BAIRRO:").grid(row=2, column=1, padx=10, pady=2, sticky="w")
ttk.Entry(frame2, textvariable=bairro_imovel_var).grid(row=3, column=1, padx=10, pady=2, sticky="ew")

ttk.Label(frame2, text="CIDADE:").grid(row=2, column=2, padx=10, pady=2, sticky="w")
ttk.Entry(frame2, textvariable=cidade_imovel_var).grid(row=3, column=2, padx=5, pady=5, sticky="ew")

ttk.Label(frame2, text="Número:").grid(row=4, column=0, padx=10, pady=2, sticky="w")
numero_imovel = ttk.Entry(frame2)
numero_imovel.grid(row=5, column=0, padx=5, pady=5, sticky="ew")

# COMARCA
frame3 = tk.LabelFrame(root, text="Comarca")
frame3.pack(padx=10, pady=10, fill="both")

# Configurar as colunas do frame para serem redimensionáveis
frame3.grid_columnconfigure(0, weight=1)
frame3.grid_columnconfigure(1, weight=1)

ttk.Label(frame3, text="Cidade").grid(row=0, column=0, padx=10, pady=2, sticky="w")
cidade_comarca = ttk.Entry(frame3)
cidade_comarca.grid(row=1, column=0, padx=5, pady=5, sticky="ew")

ttk.Label(frame3, text="UF").grid(row=0, column=1, padx=10, pady=2, sticky="w")
uf_comarca = ttk.Entry(frame3)
uf_comarca.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

ttk.Label(frame3, text="Data Atual").grid(row=0, column=2, padx=10, pady=2, sticky="w")
input_data_atual = ttk.Entry(frame3, textvariable=data_atual_comarca, state='readonly')
input_data_atual.grid(row=1, column=2, padx=5, pady=5, sticky="ew")

ttk.Button(root, text="Enviar", command=submit_form).pack(pady=10)
status_label = ttk.Label(root,  text="")
status_label.pack()


carrega_data_atual()
root.mainloop()
