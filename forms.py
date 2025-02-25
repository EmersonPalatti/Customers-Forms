import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
import shutil
import os
import re
import tempfile
import uuid
from io import BytesIO

# Definir caminhos relativos para os templates e logo (assumindo que estão na raiz do repositório)
TEMPLATE_PATH1 = "test1.xlsx"
TEMPLATE_PATH2 = "test2.xlsx"
MERCK_LOGO = "merck1.jpg"

# Função para sanitizar o nome da empresa para uso em nomes de arquivos
def sanitize_filename(name):
    name = re.sub(r'[\/:*?"<>|]', '', name)
    name = name.replace(' ', '_')
    return name

# Função para gerar um nome de arquivo baseado no nome da empresa
def generate_unique_name(base_name, empresa):
    sanitized_empresa = sanitize_filename(empresa)
    return f"{base_name}_{sanitized_empresa}.xlsx"

# Função para salvar dados nos arquivos Excel e retornar o arquivo como bytes
def save_to_excel(template_path, data, cells, image_keys):
    wb = load_workbook(template_path)
    ws = wb.active
    temp_files = []

    for key, cell in cells.items():
        if key in image_keys and data[key] is not None:
            temp_img_path = os.path.join(tempfile.gettempdir(), f"{key}_{uuid.uuid4()}.png")
            with open(temp_img_path, "wb") as f:
                f.write(data[key].getvalue())
            img = OpenpyxlImage(temp_img_path)
            ws.add_image(img, cell)
            temp_files.append(temp_img_path)
        else:
            ws[cell] = data[key]

    # Salvar em um buffer de memória em vez de no disco
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    # Limpar arquivos temporários
    for temp_file in temp_files:
        if os.path.exists(temp_file):
            os.remove(temp_file)

    return buffer

# Logo
st.logo(MERCK_LOGO, size='large')

# Título do formulário
st.title("Registration Form")

with st.expander('Company Details', expanded=True):
    nome_empresa = st.text_input("Company Name", placeholder='Merck S/A')
    col1, col2 = st.columns(2)

    with col1:
        cnpj = st.text_input("Tax ID", placeholder='33.069.212/0038-76', help='CPF or CNPJ, Format: XX.XXX.XXX/XXXX-XX')
        inscricao_estadual = st.text_input('State Tax ID')
        n_suframa = st.text_input('Suframa Number')
        cod_df = st.text_input('DF Code')
    with col2:
        telefone_fixo = st.text_input("Landline", placeholder='(00) 0000-0000')
        celular = st.text_input("Cell phone", placeholder='(00) 00000-0000')
        email = st.text_input("Email for XML", help='In the case of an individual, it must be the email of the person being registered.')

with st.expander('Address Details'):
    col3, col4 = st.columns(2)

    with col3:
        endereco = st.text_input("Address", placeholder='Alameda Xingu')
        endereco_n = st.text_input("Address Number", placeholder='350')
        endereco_bairro = st.text_input("Neighborhood", placeholder='Alphaville Industrial')
        cep = st.text_input("ZIP Code", placeholder='06455-030', help='Format: 00000-000')
    
    with col4:
        cidade = st.text_input("City", placeholder='Barueri')
        uf = st.text_input("State", placeholder='SP')
        caixa_postal = st.text_input("Postal Box")
    
    st.divider()
    st.write('Complement')
    col5, col6 = st.columns(2)

    with col5:    
        sigla_universidade = st.text_input("University Short Name")
        sigla_instituto = st.text_input("Institute Short Name")
        departamento = st.text_input("Department")
        laboratorio = st.text_input("Laboratory")
    with col6:
        bloco_predio = st.text_input("Building")
        andar = st.text_input("Floor")
        sala = st.text_input("Room")

with st.expander('Contact Details'):
    nome_contato = st.text_input("Contact Name")
    cargo = st.text_input("Position")
    email_contato = st.text_input("Contact Email")
    telefone_contato = st.text_input("Contact Phone", placeholder='(00) 00000-0000')

with st.expander('Contribution Details'):
    col7, col8 = st.columns(2)

    with col7:
        tipo_empresa = st.text_input("Company Type")
        uso_produtos = st.text_input("Use for Products")

    with col8:
        area_atuacao_empresa = st.text_input("Area of Activity")
        tipo_contribuicao = st.text_input("Contribution Type")
        
    st.divider()
    st.write('Tax Incentive')
    col9, col10 = st.columns(2)

    with col9:
        icms = st.selectbox("ICMS", ('Isento', 'Contribuinte'), placeholder='Isento', index=None)
        ipi = st.selectbox("IPI", ('Isento', 'Contribuinte'), placeholder='Contribuinte', index=None)
    with col10:    
        pis = st.selectbox("PIS", ('Isento', 'Contribuinte'), placeholder='Contribuinte', index=None)
        cofins = st.selectbox("COFINS", ('Isento', 'Contribuinte'), placeholder='Contribuinte', index=None)
    observacao_incentivo_geral = st.text_area("Observation", placeholder='Put here any observation about the Tax Incentive...')

with st.expander('Associated Companies'):
    n_associated_companies = st.number_input("Number of Associated Companies", min_value=0, max_value=3, value=0)
    for i in range(n_associated_companies):
        st.divider()
        associated_name = st.text_input("Associated Company Name", key=f"company_name_{i}")
        associated_tax_id = st.text_input("Associated Company Tax ID", key=f"tax_id_{i}")

with st.expander('Comprovants'):
    col11, col12 = st.columns(2)

    with col11:
        comprovante_endereco = st.file_uploader("Address Proof", type=['jpg', 'jpeg', 'png'])
        cartao_receita_federal = st.file_uploader("Federal Tax Card", type=['jpg', 'jpeg', 'png'], help='www.receita.fazenda.gov.br')
        exclusivo_pessoa_fisica = st.file_uploader("Exclusive for Individuals", type=['jpg', 'jpeg', 'png'], help='http://buscatextual.cnpq.br/buscatextual/busca.do?metodo=apresentar')
    
    with col12:
        cartao_sintegra = st.file_uploader("Sintegra Card", type=['jpg', 'jpeg', 'png'], help='www.sintegra.gov.br')
        cartao_suframa = st.file_uploader("Suframa Card", type=['jpg', 'jpeg', 'png'], help='https://servicos.suframa.gov.br/servicos')

# Dicionário com os dados do formulário
data = {
    "nome_empresa": nome_empresa,
    "cnpj": cnpj,
    "inscricao_estadual": inscricao_estadual,
    "n_suframa": n_suframa,
    "cod_df": cod_df,
    "telefone_fixo": telefone_fixo,
    "celular": celular,
    "email": email,
    "endereco": endereco,
    "endereco_n": endereco_n,
    "endereco_bairro": endereco_bairro,
    "cep": cep,
    "cidade": cidade,
    "uf": uf,
    "caixa_postal": caixa_postal,
    "sigla_universidade": sigla_universidade,
    "sigla_instituto": sigla_instituto,
    "departamento": departamento,
    "laboratorio": laboratorio,
    "bloco_predio": bloco_predio,
    "andar": andar,
    "sala": sala,
    "nome_contato": nome_contato,
    "cargo": cargo,
    "email_contato": email_contato,
    "telefone_contato": telefone_contato,
    "tipo_empresa": tipo_empresa,
    "uso_produtos": uso_produtos,
    "area_atuacao_empresa": area_atuacao_empresa,
    "tipo_contribuicao": tipo_contribuicao,
    "icms": icms,
    "ipi": ipi,
    "pis": pis,
    "cofins": cofins,
    "observacao_incentivo_geral": observacao_incentivo_geral,
    "comprovante_endereco": comprovante_endereco,
    "cartao_receita_federal": cartao_receita_federal,
    "exclusivo_pessoa_fisica": exclusivo_pessoa_fisica,
    "cartao_sintegra": cartao_sintegra,
    "cartao_suframa": cartao_suframa
}

# Dicionários com as células correspondentes para cada arquivo Excel
cells_test1 = {
    "nome_empresa": "C11",
    "cnpj": "D11",
    "inscricao_estadual": "C14",
    "n_suframa": "D14",
    "cod_df": "C15",
    "telefone_fixo": "D15",
    "celular": "C16",
    "email": "D16",
    "endereco": "C17",
    "endereco_n": "D17",
    "endereco_bairro": "C18",
    "cep": "D18",
    "cidade": "C19",
    "uf": "D19",
    "caixa_postal": "C20",
    "sigla_universidade": "D20",
    "sigla_instituto": "C21",
    "departamento": "D21",
    "laboratorio": "C22",
    "bloco_predio": "D22",
    "andar": "C23",
    "sala": "D23",
    "nome_contato": "C24",
    "cargo": "D24",
    "email_contato": "C25",
    "telefone_contato": "D25",
    "tipo_empresa": "C26",
    "uso_produtos": "D26",
    "area_atuacao_empresa": "C27",
    "tipo_contribuicao": "D27",
    "icms": "C28",
    "ipi": "D28",
    "pis": "C29",
    "cofins": "D29",
    "observacao_incentivo_geral": "C30",
    "comprovante_endereco": "E11",
    "cartao_receita_federal": "E12",
    "exclusivo_pessoa_fisica": "E13",
    "cartao_sintegra": "E14",
    "cartao_suframa": "E15"
}

cells_test2 = {
    "nome_empresa": "C11",
    "cnpj": "D11",
    "inscricao_estadual": "C14",
    "n_suframa": "D14",
    "cod_df": "C15",
    "telefone_fixo": "D15",
    "celular": "C16",
    "email": "D16",
    "endereco": "C17",
    "endereco_n": "D17",
    "endereco_bairro": "C18",
    "cep": "D18",
    "cidade": "C19",
    "uf": "D19",
    "caixa_postal": "C20",
    "sigla_universidade": "D20",
    "sigla_instituto": "C21",
    "departamento": "D21",
    "laboratorio": "C22",
    "bloco_predio": "D22",
    "andar": "C23",
    "sala": "D23",
    "nome_contato": "C24",
    "cargo": "D24",
    "email_contato": "C25",
    "telefone_contato": "D25",
    "tipo_empresa": "C26",
    "uso_produtos": "D26",
    "area_atuacao_empresa": "C27",
    "tipo_contribuicao": "D27",
    "icms": "C28",
    "ipi": "D28",
    "pis": "C29",
    "cofins": "D29",
    "observacao_incentivo_geral": "C30",
    "comprovante_endereco": "F11",
    "cartao_receita_federal": "F12",
    "exclusivo_pessoa_fisica": "F13",
    "cartao_sintegra": "F14",
    "cartao_suframa": "F15"
}

# Lista de chaves que correspondem a imagens
image_keys = [
    "comprovante_endereco",
    "cartao_receita_federal",
    "exclusivo_pessoa_fisica",
    "cartao_sintegra",
    "cartao_suframa"
]

# Botão para enviar os dados
if st.button("Send"):
    if not nome_empresa:
        st.error("Por favor, preencha o nome da empresa.")
    else:
        # Gerar nomes de arquivos baseados no nome da empresa
        filename1 = generate_unique_name("test1", nome_empresa)
        filename2 = generate_unique_name("test2", nome_empresa)

        # Gerar os arquivos Excel em memória
        file1_buffer = save_to_excel(TEMPLATE_PATH1, data, cells_test1, image_keys)
        file2_buffer = save_to_excel(TEMPLATE_PATH2, data, cells_test2, image_keys)

        # Exibir botões de download
        st.success("Information saved successfully! Download the files below:")
        st.download_button(
            label="Download test1.xlsx",
            data=file1_buffer,
            file_name=filename1,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.download_button(
            label="Download test2.xlsx",
            data=file2_buffer,
            file_name=filename2,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
