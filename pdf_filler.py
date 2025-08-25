
# pdf_filler.py

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from pypdf import PdfReader, PdfWriter
import io
import os

# Importa o mapeamento dos campos do PDF
from pdf_mapping import pdf_fields

def fill_pdf_document(input_pdf_path, output_dir_path, data_to_fill):
    """
    Preenche um documento PDF com os dados fornecidos.

    Args:
        input_pdf_path (str): Caminho completo para o arquivo PDF de entrada.
        output_dir_path (str): Caminho do diretório onde o PDF preenchido será salvo.
        data_to_fill (dict): Dicionário com os dados a serem preenchidos nos campos do PDF.
                              As chaves devem corresponder aos nomes dos campos em pdf_fields.
    Returns:
        str: Caminho completo do PDF preenchido, se bem-sucedido.
        None: Se ocorrer um erro.
    """
    try:
        # Criar um novo PDF em memória para sobrepor o texto
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=letter)

        # Definir a fonte e o tamanho
        can.setFont("Helvetica", 8) # Ajuste o tamanho da fonte conforme necessário

        # Preencher os campos com os dados
        for field_name, text_value in data_to_fill.items():
            if field_name in pdf_fields and text_value:
                field_info = pdf_fields[field_name]
                x = field_info["x"]
                y = field_info["y"]
                can.drawString(x, y, text_value)

        can.save()

        # Mover para o início do stream
        packet.seek(0)
        new_pdf = PdfReader(packet)

        # Ler o PDF original
        existing_pdf = PdfReader(open(input_pdf_path, "rb"))
        output = PdfWriter()

        # Adicionar a primeira página do PDF original
        page = existing_pdf.pages[0]
        page.merge_page(new_pdf.pages[0])
        output.add_page(page)

        # Salvar o PDF preenchido
        output_pdf_path_full = os.path.join(output_dir_path, f"feriado preenchido para {data_to_fill.get('razao_social', 'empresa')}.pdf")
        with open(output_pdf_path_full, "wb") as outputStream:
            output.write(outputStream)

        return output_pdf_path_full

    except Exception as e:
        print(f"Erro ao preencher o PDF: {e}")
        return None


