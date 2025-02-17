import os
import time
import re
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# Configuração da pasta onde os certificados serão monitorados
PASTA_CERTIFICADOS = r"T:\Tecnologia Informação\2025\CNPJ"


# Função para criar a instância do Chrome com o perfil
def criar_driver_com_perfil():
    print("Criando uma instância do Chrome com o perfil do usuário...")
    options = Options()
    options.add_argument(r"user-data-dir=C:\Users\gabriel.lima\AppData\Local\Google\Chrome\User Data")
    options.add_argument(r"--profile-directory=Default")  # Referência ao perfil "Default" do Chrome
    driver = webdriver.Chrome(options=options)
    print("Navegador criado com sucesso!")
    return driver


# Função para parsear o arquivo e extrair o nome da empresa e a senha
def extrair_nome_e_senha_do_certificado(caminho_certificado):
    """
    Extrai o nome da empresa e a senha do certificado com base no padrão de nome do arquivo.
    Exemplo: "Empresa ABC (senha123).pfx" -> Nome: "Empresa ABC", Senha: "senha123"
    """
    nome_arquivo = os.path.basename(caminho_certificado)
    print(f"Processando o arquivo: {nome_arquivo}")

    # Usar regex para extrair o nome e a senha
    padrao = r"^(.*?)\s*\((.*?)\)\.pfx$"  # Pega o texto antes dos parênteses como Nome e dentro dos parênteses como Senha
    match = re.match(padrao, nome_arquivo)
    if match:
        nome_empresa = match.group(1).strip()  # Texto antes dos parênteses
        senha = match.group(2).strip()  # Texto dentro dos parênteses
        print(f"Nome da empresa extraído: {nome_empresa}")
        print(f"Senha extraída: {senha}")
        return nome_empresa, senha
    else:
        raise ValueError(f"Não foi possível extrair nome e senha do arquivo: {nome_arquivo}")


# Função para pesquisar pela empresa no campo de busca
def buscar_empresa(driver, nome_empresa):
    """
    Busca pelo nome da empresa no campo de busca do site.
    """
    print(f"Pesquisando pela empresa '{nome_empresa}'...")
    try:
        # Localiza o campo de busca e insere o nome da empresa
        campo_busca = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "input.custom-input[placeholder='Filtre por nome ou CNPJ...']"))
        )
        campo_busca.clear()
        campo_busca.send_keys(nome_empresa)
        print(f"Nome da empresa '{nome_empresa}' inserido no campo de busca.")

        # Aguarda a tabela filtrar os resultados
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "tr.linha-empresa"))
        )
        print("Tabela filtrada com sucesso.")
        return True
    except Exception as e:
        print(f"Erro ao pesquisar pela empresa: {str(e)}")
        return False


# Função para encontrar a linha correspondente após a busca
def encontrar_linha_empresa(driver):
    """
    Encontra a linha da tabela correspondente após a busca.
    """
    try:
        # Pega a primeira linha da tabela (após o filtro)
        linha = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "tr.linha-empresa"))
        )
        print("Linha da empresa localizada.")
        return linha
    except Exception as e:
        print(f"Erro ao localizar a linha da empresa: {str(e)}")
        return None


# Função para atualizar o certificado no site
def atualizar_certificado_no_site(driver, senha, caminho_certificado, nome_empresa):
    print("Iniciando o processo de atualização do certificado no site...")
    try:
        # Navegar até a página de clientes
        print("Acessando a página de clientes...")
        driver.get("https://app.simpleshub.com.br/customers")
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "input.custom-input[placeholder='Filtre por nome ou CNPJ...']"))
        )
        print("Página de clientes carregada.")

        # Buscar pela empresa no campo de busca
        if not buscar_empresa(driver, nome_empresa):
            print(f"Não foi possível localizar a empresa '{nome_empresa}' no campo de busca. Processo encerrado.")
            return

        # Encontrar a linha correspondente
        linha_empresa = encontrar_linha_empresa(driver)
        if not linha_empresa:
            print(
                f"Não foi possível identificar nenhuma empresa correspondente a '{nome_empresa}'. Processo encerrado.")
            return

        # Clicar no ícone para excluir o certificado anterior
        print("Excluindo o certificado atual...")
        try:
            icone_lixo = linha_empresa.find_element(By.CSS_SELECTOR,
                                                    "button.custom-button-destructive:not(.v-popper--has-tooltip)")
            icone_lixo.click()
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "button.custom-button-destructive"))
            ).click()
            print("Certificado anterior excluído.")
        except Exception as e:
            print(f"Erro ao excluir o certificado anterior: {str(e)}")
            return

        # Clique no botão de adicionar certificado
        print("Adicionando novo certificado...")
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "button.custom-button-secondary"))
        ).click()

        # Enviar o arquivo do certificado
        campo_upload = driver.find_element(By.CSS_SELECTOR, "input[type='file']")
        campo_upload.send_keys(caminho_certificado)
        print("Certificado enviado.")

        # Preencher a senha do certificado
        campo_senha = driver.find_element(By.CSS_SELECTOR, "input.custom-input")
        campo_senha.send_keys(senha)
        print("Senha do certificado preenchida.")

        # Clicar no botão para salvar
        botao_salvar = driver.find_element(By.CSS_SELECTOR, "button.custom-button-brand")
        botao_salvar.click()
        print("Certificado salvo com sucesso!")
    except Exception as e:
        print(f"Erro ao atualizar o certificado: {str(e)}")


# Função para monitorar a pasta e iniciar o processo ao encontrar um novo arquivo
def monitorar_pasta():
    print(f"Iniciando monitoramento da pasta: {PASTA_CERTIFICADOS}")

    # Estado inicial: arquivos existentes na pasta
    arquivos_existentes = set(os.listdir(PASTA_CERTIFICADOS))

    while True:
        # Verifica os arquivos atuais
        arquivos_atual = set(os.listdir(PASTA_CERTIFICADOS))

        # Identifica novos arquivos adicionados
        novos_arquivos = arquivos_atual - arquivos_existentes
        if novos_arquivos:
            for arquivo in novos_arquivos:
                if arquivo.lower().endswith(".pfx"):
                    print(f"Novo arquivo detectado: {arquivo}")
                    return os.path.join(PASTA_CERTIFICADOS, arquivo)

        # Atualiza o estado dos arquivos existentes
        arquivos_existentes = arquivos_atual

        # Intervalo para verificar novamente (em segundos)
        time.sleep(3)


# Script principal
if __name__ == "__main__":
    print("Aguardando novos certificados serem adicionados...")

    while True:
        try:
            # Monitora a pasta até detectar um novo certificado
            caminho_certificado = monitorar_pasta()

            # Extrai o nome da empresa e a senha do arquivo
            nome_empresa, senha_certificado = extrair_nome_e_senha_do_certificado(caminho_certificado)

            # Abre o navegador e processa o certificado
            driver = criar_driver_com_perfil()
            try:
                atualizar_certificado_no_site(driver, senha_certificado, caminho_certificado, nome_empresa)
            finally:
                print("Encerrando o navegador...")
                driver.quit()

        except Exception as e:
            print(f"Erro durante o processamento do certificado: {str(e)}")

        print("Voltando a monitorar a pasta para novos certificados...")
