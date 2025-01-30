import os
import win32print
import win32ui
from win32con import DM_OUT_BUFFER, DM_IN_BUFFER

# Caminho da pasta com os PDFs
pasta_pdfs = r"T:\Publico\04255015000151\tst"


def imprimir_pdf_diretamente(caminho_pdf, impressora):
    try:
        # Abre uma conexão com a impressora
        handle = win32print.OpenPrinter(impressora)
        printer_info = win32print.GetPrinter(handle, 2)

        # Configurações padrão do job de impressão
        dev_mode = printer_info["pDevMode"]
        dev_mode.Fields |= DM_OUT_BUFFER | DM_IN_BUFFER

        # Cria o job de impressão
        win32print.StartDocPrinter(handle, 1, ("Impressão de PDF", None, "RAW"))
        win32print.StartPagePrinter(handle)

        with open(caminho_pdf, "rb") as pdf_file:
            data = pdf_file.read()
            win32print.WritePrinter(handle, data)

        win32print.EndPagePrinter(handle)
        win32print.EndDocPrinter(handle)
        win32print.ClosePrinter(handle)

        print(f"Impressão concluída: {os.path.basename(caminho_pdf)}")
    except Exception as e:
        print(f"Erro ao imprimir {os.path.basename(caminho_pdf)}: {e}")


# Obter a impressora padrão
impressora = win32print.GetDefaultPrinter()
print(f"Impressora padrão detectada: {impressora}")

# Verificar e imprimir os PDFs
if os.path.exists(pasta_pdfs):
    pdfs = [f for f in os.listdir(pasta_pdfs) if f.lower().endswith('.pdf')]

    if pdfs:
        for pdf in pdfs:
            caminho_pdf = os.path.join(pasta_pdfs, pdf)
            imprimir_pdf_diretamente(caminho_pdf, impressora)
    else:
        print("Não há arquivos PDF na pasta especificada.")
else:
    print(f"A pasta especificada não foi encontrada: {pasta_pdfs}")