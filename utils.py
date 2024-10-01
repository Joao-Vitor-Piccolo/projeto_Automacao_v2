from pptx import Presentation
from pptx.dml.color import RGBColor
from openpyxl import load_workbook
import os
import fitz
from pptxtopdf import convert
import win32com.client as win32
import json

config_path = 'config.json'


def load_config(config_p):
    if os.path.exists(config_p):
        with open(config_p, 'r') as file:
            return json.load(file)
    else:
        return {
            "horario": "15:00",
            "email": "jvpiccolo13@gmail.com",
            "planilha": "planilha_cliente.xlsx",
            "pptx0_file": "Informativo - Email Especialistas.pptx"
        }


config = load_config(config_path)

email_list = []
name_list = []

x = 0

wb = load_workbook(config['planilha'])
ws = wb.active
lista = []
lista_copy = lista
path1 = config['pptx0_file']
diretorio = os.path.join(os.getcwd(), 'Slides')
outlook = win32.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

ppt = Presentation(path1)
slide = ppt.slides[0]

for row in ws.iter_rows(values_only=True):
    if any(cell is not None for cell in row):
        lista.append(row)

for indice, tupla in enumerate(lista):
    email_funcionario = lista[indice][-1].split(":")[1].replace(" ", "")
    name_funcionario = lista_copy[indice][-1].split(":")[0]
    if " " in name_funcionario:
        name_funcionario = name_funcionario.split(" ")[0]
    email_list.append(email_funcionario)
    name_list.append(name_funcionario)

"""
Abaixo jás uma função do código que pediram para não implementar, deixarei aqui via as duvidas


# Cheva se "Default" se encontra no primeiro bloco do Excel
def check0():
    if 'Defaut' in ws['A1'].value:
        # se sim, retorne True
        return True
    # Se não, retorne False
    return False

# Faz um template inicial para a planilha, deixando-a inteira em "Default-Mode"
def default_p():
    ws.delete_rows(1, ws.max_row)  # Deleta a planilha inteira 
    dados_default = ['Empresa_Default', '11.111.111/0000-00', '(11) 11111-1111', 'contato@default.com', 'nome_default',
                     'Funcionario Default: funcionario_default@gmail.com']
    for coluna, valor in enumerate(dados_default, 1):
        ws.cell(1, coluna, valor)  # Insere os dados "Default"
    wb.save(config["planilha"])  # Salva a planilha
"""


def list_s():
    text_boxes = []
    for shape in slide.shapes:
        if shape.has_text_frame and shape.text:
            text_boxes.append(shape.text)
    return text_boxes

def change_text(txt_id, new_txt):
    count = 0
    for shape in slide.shapes:
        if shape.has_text_frame and shape.text:
            count += 1
            if count == txt_id:
                text_frame = shape.text_frame
                first_paragraph = text_frame.paragraphs[0]
                first_run = first_paragraph.runs[0] if first_paragraph.runs else first_paragraph.add_run()
                # Preserve formatting of the first run
                font = first_run.font
                font_name = font.name
                font_size = font.size
                font_bold = font.bold
                # Clear existing text and apply new text with preserved formatting
                text_frame.clear()  # Clears all text and formatting
                new_run = text_frame.paragraphs[0].add_run()  # New run in first paragraph
                new_run.text = new_txt
                # Apply the new run
                new_run.font.name = font_name
                new_run.font.size = font_size
                new_run.font.bold = font_bold
                new_run.font.color.rgb = RGBColor(255, 255, 255)
                return


class Cliente:
    def __init__(self, empresa, cnpj, celular, email, nome_s):
        self.empresa = empresa
        self.cnpj = cnpj
        self.celular = celular
        self.email = email
        self.nome_s = nome_s

async def make_slide():
    global x
    nome_funcionario = lista_copy[0][5].split(":")[0]
    if " " in nome_funcionario:
        nome_funcionario = nome_funcionario.split(" ")[0]
    cliente = Cliente(lista_copy[0][0], lista_copy[0][1], lista_copy[0][2], lista_copy[0][3], lista_copy[0][4])
    change_text(1, nome_funcionario + ', TEM UM CLIENTE NOVO ESPERANDO O SEU BOAS VINDAS!!!')
    change_text(3, 'Empresa: ' + cliente.empresa)
    change_text(4, 'CNPJ: ' + cliente.cnpj)
    change_text(5, 'Telefone: ' + cliente.celular)
    change_text(6, 'Email: ' + cliente.email)
    change_text(7, 'Nome dos sócios: ' + cliente.nome_s)
    path = os.path.join(diretorio, f'Email_{str(x + 1)}_Especialistas.pptx')
    ppt.save(path)
    x += 1
    print(f'Slides: {x} PRONTO!')
    lista_copy.pop(0)
    return path


def clear_files(file: str):
    try:
        os.remove(file)
        print(f"Deleted: {file}")
    except Exception as error:
        print(f"Failed to delete {file}: {error}")


async def convert_to_pdf(file):
    try:
        convert(diretorio, diretorio)
        print("Conversion done!")
        clear_files(file)
        return True
    except Exception as e:
        print(f"(PPTX) Conversion failed: {e}")


async def convert_to_img(file):
    try:
        if ".pdf" in file:
            pdf = fitz.open(file)
            page = pdf.load_page(0)
            pix = page.get_pixmap()
            file2 = file.replace(".pdf", ".jpg")
            pix.save(file2)
            pdf.close()
            clear_files(file)
            return file2
    except Exception as e:
        print(f"(PDF) Conversion failed: {e}")


def check_conta(mail):
    for myEmailAddress in outlook.Session.Accounts:
        if config['email'] in str(myEmailAddress):
            mail._oleobj_.Invoke(*(64209, 0, 8, 0, myEmailAddress))
            return True

async def send_email(file):
    nome_funcionario = name_list.pop(0)
    mail = outlook.CreateItem(0)
    if check_conta(mail):
        mail.Subject = f"{nome_funcionario}, Veja seus clientes!!"
        mail.HTMLBody = f"""<html>
                    <body>
                        <h2 style="border: 2px solid black; padding: 10px; color: white; text-align: center;">
                            {nome_funcionario}, Abra para ver seus clientes:
                        </h2>
                        <img src="cid:image1" width="500" style="display: block; margin: 0 auto;">
                    </body>
                    </html>
                    """
        attachment = mail.Attachments.Add(file)
        attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F",
                                                "image1")
        mail.To = email_list[0]
        try:
            mail.Send()
            print('Enviado para:', email_list[0])
            email_list.pop(0)
        except Exception as error:
            print('Não enviado: ', error)
        clear_files(file)
    #   default_p()


def get_time():
    return config['horario']
