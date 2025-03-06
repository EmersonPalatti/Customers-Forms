import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
import shutil
import os
import re
import tempfile
import uuid
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Caminhos relativos dos templates Excel
TEMPLATE_PATH1 = "templates/Merck.xlsx"
TEMPLATE_PATH2 = "templates/Sigma.xlsx"
LOGO_PATH = "merck1.jpg"

# Função para sanitizar o nome da empresa para uso em nomes de arquivos
def sanitize_filename(name):
    name = re.sub(r'[\/:*?"<>|]', '', name)
    name = name.replace(' ', '_')
    return name

# Função para gerar um nome de arquivo baseado no nome da empresa
def generate_unique_name(base_name, empresa):
    sanitized_empresa = sanitize_filename(empresa)
    return f"{base_name}_{sanitized_empresa}.xlsx"

# Função para salvar os dados no Excel
def save_to_excel(path, data, cells_sold_to, cells_ship_to, image_keys, sheet_sold_to="FICHA CADASTRAL (Sold-to)", sheet_ship_to="FICHA CADASTRAL (Ship-to)"):
    wb = load_workbook(path)
    temp_files = []

    # Aba Sold-to (principal)
    ws_sold_to = wb[sheet_sold_to]
    for key, cell in cells_sold_to.items():
        if key in image_keys and data.get(key) is not None:
            temp_img_path = os.path.join(tempfile.gettempdir(), f"{key}_{uuid.uuid4()}.png")
            with open(temp_img_path, "wb") as f:
                f.write(data[key].getvalue())
            img = OpenpyxlImage(temp_img_path)
            ws_sold_to.add_image(img, cell)
            temp_files.append(temp_img_path)
        elif key in ["associated_names", "associated_tax_ids"] and data.get(key):
            values = data[key].split("; ")
            for i, value in enumerate(values):
                if i < len(cell):
                    ws_sold_to[cell[i]] = value.strip()
        elif isinstance(cell, str):
            ws_sold_to[cell] = data.get(key)

    # Aba Ship-to (endereço de entrega, se aplicável)
    if data.get("shipping_address", False):
        ws_ship_to = wb[sheet_ship_to]
        for key, cell in cells_ship_to.items():
            if key in image_keys and data.get(key) is not None:
                temp_img_path = os.path.join(tempfile.gettempdir(), f"{key}_{uuid.uuid4()}.png")
                with open(temp_img_path, "wb") as f:
                    f.write(data[key].getvalue())
                img = OpenpyxlImage(temp_img_path)
                ws_ship_to.add_image(img, cell)
                temp_files.append(temp_img_path)
            else:
                ws_ship_to[cell] = data.get(key)

    wb.save(path)
    for temp_file in temp_files:
        if os.path.exists(temp_file):
            os.remove(temp_file)

# Função para enviar e-mail com anexos
def send_email(sender_email, receiver_email, subject, body, files, password):
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    for file_path in files:
        with open(file_path, 'rb') as f:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(file_path)}')
            msg.attach(part)

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, password)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        st.error(f"Erro ao enviar e-mail: {str(e)}")
        return False

# Logo
st.logo(LOGO_PATH, size='large')

# Título do formulário
st.title("Formulário para Cadastro")

with st.expander('Dados de Cadastro', expanded=True):
    nome_empresa = st.text_input("Razão Social", placeholder='Merck S/A')
    col1, col2 = st.columns(2)
    with col1:
        cnpj = st.text_input("CNPJ/CPF", placeholder='33.069.212/0038-76', help='CPF or CNPJ, Format: XX.XXX.XXX/XXXX-XX')
        inscricao_estadual = st.text_input('Inscrição Estadual')
        n_suframa = st.text_input('Número Suframa')
        cod_df = st.text_input('Código DF')
    with col2:
        telefone_fixo = st.text_input("Telefone Fixo", placeholder='(00) 0000-0000')
        celular = st.text_input("Telefone Celular", placeholder='(00) 00000-0000')
        email = st.text_input("Email para envio do XML", help='No caso de pessoa física deverá ser o e-mail da pessoa que está sendo cadastrada.')

with st.expander('Endereço'):
    col3, col4 = st.columns(2)
    with col3:
        endereco = st.text_input("Endereço", placeholder='Alameda Xingu')
        endereco_n = st.text_input("Número", placeholder='350')
        endereco_bairro = st.text_input("Bairro", placeholder='Alphaville Industrial')
        cep = st.text_input("CEP", placeholder='06455-030', help='Format: 00000-000')
    with col4:
        cidade = st.text_input("Cidade", placeholder='Barueri')
        uf = st.text_input("Estado", placeholder='SP')
        caixa_postal = st.text_input("Caixa Postal")
    
    complement = st.toggle('Endereço necessita de complemento?', key='complement', help='Complemento (Ex: Bloco, Sala, etc.)')
    if complement:
        st.write('Complementos')
        col5, col6 = st.columns(2)
        with col5:    
            sigla_universidade = st.text_input("Sigla da Universidade", key='sold_to_universidade')
            sigla_instituto = st.text_input("Sigla do Instituto", key='sold_to_instituto')
            departamento = st.text_input("Departamento", key='sold_to_departamento')
            laboratorio = st.text_input("Laboratório", key='sold_to_laboratorio')
        with col6:
            bloco_predio = st.text_input("Bloco do Prédio", key='sold_to_bloco')
            andar = st.text_input("Andar", key='sold_to_andar')
            sala = st.text_input("Sala", key='sold_to_sala')
    else:
        sigla_universidade = sigla_instituto = departamento = laboratorio = bloco_predio = andar = sala = None

    shipping_address = st.toggle('Endereço de entrega é em outro local?', key='shipping_address', help='Ative a opção caso o endereço de entrega seja diferente do endereço de faturamento.')
    with st.container(border=True):
        if shipping_address:
            st.write('Endereço de Entrega')
            col13, col14 = st.columns(2)
            with col13:
                shipping_endereco = st.text_input("Endereço", placeholder='Alameda Xingu', key='shipping_endereco')
                shipping_endereco_n = st.text_input("Número", placeholder='350', key='shipping_endereco_n')
                shipping_endereco_bairro = st.text_input("Bairro", placeholder='Alphaville Industrial', key='shipping_endereco_bairro')
                shipping_cep = st.text_input("CEP", placeholder='06455-030', help='Format: 00000-000', key='shipping_cep')
            with col14:
                shipping_cidade = st.text_input("Cidade", placeholder='Barueri', key='shipping_cidade')
                shipping_uf = st.text_input("Estado", placeholder='SP', key='shipping_uf')
                shipping_caixa_postal = st.text_input("Caixa Postal", key='shipping_caixa_postal')
            shipping_address_complement = st.toggle('Endereço de entrega necessita de complemento?', key='shipping_address_complement', help='Complemento (Ex: Bloco, Sala, etc.)')
            if shipping_address_complement:
                st.write('Complementos')
                col15, col16 = st.columns(2)
                with col15:    
                    shipping_sigla_universidade = st.text_input("Sigla da Universidade", key='ship_to_universidade')
                    shipping_sigla_instituto = st.text_input("Sigla do Instituto", key='ship_to_instituto')
                    shipping_departamento = st.text_input("Departamento", key='ship_to_departamento')
                    shipping_laboratorio = st.text_input("Laboratório", key='ship_to_laboratorio')
                with col16:
                    shipping_bloco_predio = st.text_input("Bloco do Prédio", key='ship_to_bloco')
                    shipping_andar = st.text_input("Andar", key='ship_to_andar')
                    shipping_sala = st.text_input("Sala", key='ship_to_sala')
            else:
                shipping_sigla_universidade = shipping_sigla_instituto = shipping_departamento = shipping_laboratorio = shipping_bloco_predio = shipping_andar = shipping_sala = None
            
            # Novo campo para comprovante de endereço de entrega dentro de um container
            with st.container(border=True):
                st.write("Comprovante de Endereço de Entrega")
                shipping_comprovante_endereco = st.file_uploader("Comprovante de Endereço de Entrega", type=['jpg', 'jpeg', 'png'], key='shipping_comprovante_endereco')
                st.caption("O comprovante de endereço deve ser obtido em https://buscacepinter.correios.com.br/app/endereco/index.php. Tire um print e anexe aqui.")
        else:
            shipping_endereco = shipping_endereco_n = shipping_endereco_bairro = shipping_cep = None
            shipping_cidade = shipping_uf = shipping_caixa_postal = None
            shipping_sigla_universidade = shipping_sigla_instituto = shipping_departamento = shipping_laboratorio = shipping_bloco_predio = shipping_andar = shipping_sala = None
            shipping_comprovante_endereco = None

with st.expander('Informações de contato - Solicitante do Cadastro'):
    nome_contato = st.text_input("Nome", key='nome_contato')
    cargo = st.text_input("Função", key='cargo')
    email_contato = st.text_input("Email", key='email_contato')
    telefone_contato = st.text_input("Telefone", placeholder='(00) 00000-0000', key='telefone_contato')

with st.expander('Informações de Contribuição'):
    col7, col8 = st.columns(2)
    with col7:
        tipo_empresa = st.selectbox("Tipo de Empresa", ('Publica', 'Privada', 'Mista'), placeholder='Escolha uma opção.', index=None, key='tipo_empresa')
        uso_produtos = st.selectbox("Uso dos Produtos", (
            'C3 = Consumidor Final: ICMS + IPI',
            'I3 = Industrialização: ICMS + IPI',
            'C5 = Consumidor Final: IPI', 
            'C0 = Consumidor Final: S/ Impostos',
            'C1 = Consumidor Final: ICMS',
            'C2 = Consumidor Final: ICMS + Sub.Trib.',
            'C4 = Consumidor Final: ICMS + Sub.Trib. + IPI',
            'CX = Consumidor Final: ICMS somente',
            'I0 = Industrialização: S/ Impostos',
            'I1 = Industrialização: ICMS',
            'I2 = Industrialização: ICMS + Sub.Trib.',
            'I4 = Industrialização: ICMS + Sub.Trib. + IPI',
            'I5 = Industrialização: IPI',
            'I9 = ISS',
            'IX = Industrialização: ICMS somente'
        ), placeholder='Escolha uma opção.', index=None, key='uso_produtos')
    with col8:
        area_atuacao_empresa = st.selectbox("Área de Atuação da Empresa", (
            'Ind./Comm./General',
            'Chemical Industry',
            'Government/nonprofit',
            'Hospitals/Physicians',
            'Biotechnology',
            'Service/testing',
            'Pharmaceutical',
            'Universities/Schools',
            'Diagnostics',
            'Dealers/Exp./Resell',
            'Distribuidor'
        ), key='area_atuacao_empresa', placeholder='Escolha uma opção.', index=None)
        tipo_contribuicao = st.text_input("Tipo de Contribuição", key='tipo_contribuicao')

    with st.container(border=True):    
        st.write('Incentivo Fiscal')
        col9, col10 = st.columns(2)
        with col9:
            icms = st.selectbox("ICMS", ('Isento', 'Contribuinte'), placeholder='Escolha uma opção.', index=None, key='icms')
            ipi = st.selectbox("IPI", ('Isento', 'Contribuinte'), placeholder='Escolha uma opção.', index=None, key='ipi')
        with col10:    
            pis = st.selectbox("PIS", ('Isento', 'Contribuinte'), placeholder='Escolha uma opção.', index=None, key='pis')
            cofins = st.selectbox("COFINS", ('Isento', 'Contribuinte'), placeholder='Escolha uma opção.', index=None, key='cofins')
        observacao_incentivo_geral = st.text_area("Observação", placeholder='Observação(ões) sobre Incentivo Fiscal', key='observacao_incentivo_geral')

with st.expander('Empresas Coligadas (preencher somente se necessário)'):
    n_associated_companies = st.number_input("Quantidade de Empresas Coligadas", min_value=0, max_value=4, value=0, key='n_associated_companies', help='Pode-se adicionar até mais 4 empresas coligadas.')
    associated_names = []
    associated_tax_ids = []
    for i in range(n_associated_companies):
        with st.container(border=True, key=f"associated_company_{i}"):
            associated_name = st.text_input(f"__Razão Social__ - Empresa Coligada{i+1}", key=f"company_name_{i}")
            associated_tax_id = st.text_input(f"__CNPJ__ - Empresa Coligada {i+1}", key=f"tax_id_{i}")
            associated_names.append(associated_name)
            associated_tax_ids.append(associated_tax_id)

with st.expander('Comprovantes'):
    col11, col12 = st.columns(2)
    with col11:
        comprovante_endereco = st.file_uploader("Comprovante de Endereço", type=['jpg', 'jpeg', 'png'], key='comprovante_endereco')
        cartao_receita_federal = st.file_uploader("Cartão da Receita Federal", type=['jpg', 'jpeg', 'png'], help='www.receita.fazenda.gov.br', key='cartao_receita_federal')
        exclusivo_pessoa_fisica = st.file_uploader("Exclusivo para Pessoa Física - Vínculo com uma Instituição", type=['jpg', 'jpeg', 'png'], help='http://buscatextual.cnpq.br/buscatextual/busca.do?metodo=apresentar. Se a pessoa física não tiver um curriculo Lates, deverá apresentar outra comprovação.', key='exclusivo_pessoa_fisica')
    with col12:
        cartao_sintegra = st.file_uploader("Cartão do Sintegra", type=['jpg', 'jpeg', 'png'], help='Somente para empresas que possui I.E. (www.sintegra.gov.br)', key='cartao_sintegra')
        cartao_suframa = st.file_uploader("Cartão Suframa", type=['jpg', 'jpeg', 'png'], help='Somente para empresa que possui Inscrição Suframa, https://servicos.suframa.gov.br/servicos', key='cartao_suframa')
        comprovante_faturamento = st.file_uploader("Contrato Social, Balanço Patrimonial, DRE ou Cartão CNPJ", type=['jpg', 'jpeg', 'png'], key='comprovante_faturamento')

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
    "shipping_address": shipping_address,
    "shipping_endereco": shipping_endereco,
    "shipping_endereco_n": shipping_endereco_n,
    "shipping_endereco_bairro": shipping_endereco_bairro,
    "shipping_cep": shipping_cep,
    "shipping_cidade": shipping_cidade,
    "shipping_uf": shipping_uf,
    "shipping_caixa_postal": shipping_caixa_postal,
    "shipping_sigla_universidade": shipping_sigla_universidade,
    "shipping_sigla_instituto": shipping_sigla_instituto,
    "shipping_departamento": shipping_departamento,
    "shipping_laboratorio": shipping_laboratorio,
    "shipping_bloco_predio": shipping_bloco_predio,
    "shipping_andar": shipping_andar,
    "shipping_sala": shipping_sala,
    "shipping_comprovante_endereco": shipping_comprovante_endereco,
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
    "associated_names": "; ".join(associated_names) if associated_names else None,
    "associated_tax_ids": "; ".join(associated_tax_ids) if associated_tax_ids else None,
    "comprovante_endereco": comprovante_endereco,
    "cartao_receita_federal": cartao_receita_federal,
    "exclusivo_pessoa_fisica": exclusivo_pessoa_fisica,
    "cartao_sintegra": cartao_sintegra,
    "cartao_suframa": cartao_suframa,
    "comprovante_faturamento": comprovante_faturamento
}

# Dicionários com as células correspondentes para cada aba
cells_sold_to = {
    "nome_empresa": "C11",
    "cnpj": "H11",
    "inscricao_estadual": "I11",
    "n_suframa": "J11",
    "cod_df": "K11",
    "telefone_fixo": "D15",
    "celular": "E15",
    "email": "F15",
    "endereco": "C13",
    "endereco_n": "G13",
    "endereco_bairro": "H13",
    "cep": "I13",
    "cidade": "J13",
    "uf": "L13",
    "caixa_postal": "C15",
    "sigla_universidade": "C19",
    "sigla_instituto": "E19",
    "departamento": "F19",
    "laboratorio": "H19",
    "bloco_predio": "J19",
    "andar": "K19",
    "sala": "L19",
    "nome_contato": "C22",
    "cargo": "F22",
    "email_contato": "G22",
    "telefone_contato": "J22",
    "tipo_empresa": "C27",
    "uso_produtos": "D27",
    "area_atuacao_empresa": "F27",
    "tipo_contribuicao": "H27",
    "icms": "C31",
    "ipi": "D31",
    "pis": "E31",
    "cofins": "F31",
    "observacao_incentivo_geral": "G29",
    "associated_names": ["C35", "C36", "C37", "C38"],
    "associated_tax_ids": ["I35", "I36", "I37", "I38"],
    "comprovante_endereco": "C55",
    "cartao_receita_federal": "C86",
    "exclusivo_pessoa_fisica": "C187",
    "cartao_sintegra": "C111",
    "cartao_suframa": "C147"
}

cells_ship_to = {
    "shipping_endereco": "C13",
    "shipping_endereco_n": "G13",
    "shipping_endereco_bairro": "H13",
    "shipping_cep": "I13",
    "shipping_cidade": "J13",
    "shipping_uf": "L13",
    "shipping_caixa_postal": "C15",
    "shipping_sigla_universidade": "C19",
    "shipping_sigla_instituto": "E19",
    "shipping_departamento": "F19",
    "shipping_laboratorio": "H19",
    "shipping_bloco_predio": "J19",
    "shipping_andar": "K19",
    "shipping_sala": "L19",
    "shipping_comprovante_endereco": "C43"
}

# Lista de chaves que correspondem a imagens
image_keys = [
    "comprovante_endereco",
    "cartao_receita_federal",
    "exclusivo_pessoa_fisica",
    "cartao_sintegra",
    "cartao_suframa",
    "shipping_comprovante_endereco"
]

# Configurações de e-mail usando st.secrets
SENDER_EMAIL = st.secrets["SENDER_EMAIL"]
RECEIVER_EMAIL = st.secrets["RECEIVER_EMAIL"]
EMAIL_PASSWORD = st.secrets["EMAIL_PASSWORD"]

# Botão para enviar os dados
if st.button("Enviar"):
    if not nome_empresa:
        st.error("Por favor, preencha o nome da empresa.")
    else:
        with st.spinner("Processando e enviando os dados..."):
            with tempfile.TemporaryDirectory() as temp_dir:
                filename1 = generate_unique_name("Merck", nome_empresa)
                filename2 = generate_unique_name("Sigma", nome_empresa)
                temp_path1 = os.path.join(temp_dir, filename1)
                temp_path2 = os.path.join(temp_dir, filename2)

                shutil.copy(TEMPLATE_PATH1, temp_path1)
                shutil.copy(TEMPLATE_PATH2, temp_path2)
                save_to_excel(temp_path1, data, cells_sold_to, cells_ship_to, image_keys)
                save_to_excel(temp_path2, data, cells_sold_to, cells_ship_to, image_keys)

                files = [temp_path1, temp_path2]
                if data.get("comprovante_faturamento") is not None:
                    temp_faturamento_path = os.path.join(temp_dir, f"Comprovante_Faturamento_{uuid.uuid4()}.png")
                    with open(temp_faturamento_path, "wb") as f:
                        f.write(data["comprovante_faturamento"].getvalue())
                    files.append(temp_faturamento_path)

                subject = f"Formulário de Cadastro - {nome_empresa}"
                body = f"Segue em anexo os arquivos preenchidos para {nome_empresa}.\n\nEnviado automaticamente pelo formulário Streamlit."
                if send_email(SENDER_EMAIL, RECEIVER_EMAIL, subject, body, files, EMAIL_PASSWORD):
                    st.success("Formulário enviado com sucesso!")
                    st.markdown(
                        """
                        <div style='text-align: center; padding: 20px; border: 2px solid rgb(235, 60, 150); border-radius: 10px; background-color: rgb(79, 53, 140);'>
                            <h2 style='color: white;'>Formulário de cadastro enviado com sucesso!</h2>
                            <p style='color: white;'>Obrigada pelo interesse em nossos produtos!<br>Estamos comprometidos com a sustentabilidade e impulsionados pela paixão por inovação.<br>Obrigado por fazer parte dessa jornada conosco!<br>Seu formulário está nas mãos do nosso Time de Cadastros. Assim que o cadastro for concluído, entraremos em contato.</p>
                            <h1 style='color: #2dbecd; font-size: 40px;'>TWC</h1>
                            <p style='color: white;'>Aproveite seu desconto em nossa loja online: <a href='https://www.sigmaaldrich.com/BR/pt' target='_blank' style='color: #2dbecd;'>https://www.sigmaaldrich.com/BR/pt</a></p>
                        </div>
                        """,
                        unsafe_allow_html=True
                    )
                    st.balloons()
                else:
                    st.error("Falha ao enviar o e-mail. Verifique as credenciais ou a conexão.")
