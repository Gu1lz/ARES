import glob
import sys
import os
import pygetwindow as gw
import time
import random
from colorama import init, Fore, Back, Style
import MySQLdb
import uuid
import getpass
import shutil
import configparser
import json
import pyperclip
import configparser
from time import sleep
import pandas as pd
import threading
import pyperclip
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
from selenium.webdriver.chrome.options import Options  # Change this line
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service


sessoes = {}


class App:
    def __init__(self, email= "", password= "", 
                 path="", language="", main_url="", marketplace_url="", binary_location="", driver_location="", time_to_sleep=""):
            try:
                
                os.system('cls')
                time.sleep(0.1)
                banner = [
                "        .8.          8 888888888o.   8 8888888888      d888888o.",
                "       .888.         8 8888    `88.  8 8888          .`8888:' `88.",
                "      :88888.        8 8888     `88  8 8888          8.`8888.   Y8",
                "     . `88888.       8 8888     ,88  8 8888          `8.`8888.",
                "    .8. `88888.      8 8888.   ,88'  8 888888888888   `8.`8888.",
                "   .8`8. `88888.     8 888888888P'   8 8888            `8.`8888.",
                "  .8' `8. `88888.    8 8888`8b       8 8888             `8.`8888.",
                " .8'   `8. `88888.   8 8888 `8b.     8 8888         8b   `8.`8888.",
                ".888888888. `88888.  8 8888   `8b.   8 8888         `8b.  ;8.`8888",
                ".8'       `8. `88888.8 8888     `88. 8 888888888888  `Y8888P ,88P'"
                ]
                

                for bannerzinho in banner:
                    print(Fore.RED + " " * 15 + bannerzinho)
                    time.sleep(0.05)

                time.sleep(1)
                menu = [
                    f"",
                    f"{Fore.YELLOW}Para sair utilize a tecla CTRL+C{Fore.RESET} "
                ]

                for letra in menu:
                    print(letra)
                    time.sleep(0.05)

                self.email = email
                self.password = password
                self.path = path
                self.language = language
                self.marketplace_options = None
                self.posts = None
                self.time_to_sleep = float(time_to_sleep)
                with open('marketplace_options.json', encoding='utf-8') as f:
                    self.marketplace_options = json.load(f)
                    self.marketplace_options = self.marketplace_options[self.language]
               
                chrome_options = Options()
                chrome_options.binary_location = binary_location
                chrome_options.add_argument("--disable-notifications")
                chrome_options.add_argument("--disable-infobars")
                chrome_options.add_argument("--no-sandbox")
                chrome_options.add_argument("--disable-notifications")
                chrome_options.add_argument("--disable-infobars")
                chrome_options.add_argument("--disable-gpu")
                
                original_stdout = sys.stdout
                with open(os.devnull, 'w') as null_file:
                    sys.stdout, sys.stderr = null_file, null_file

                    self.driver = webdriver.Chrome(executable_path=driver_location, options=chrome_options)
                    
                sys.stdout = original_stdout
                os.system('cls')
                
                for bannerzinho in banner:
                    print(Fore.RED + " " * 15 + bannerzinho)

                for item in menu:
                    print(item)
                self.driver.maximize_window()
                self.main_url = str(main_url)
                self.marketplace_url = marketplace_url
                self.driver.get(self.main_url)
                self.log_in()
                self.posts = self.fetch_all_posts()
                for post in self.posts:
                    self.move_from_home_to_marketplace_create_item()
                    print(f"{Fore.GREEN}Estamos no post:{Fore.RESET} {str(post[0])}")
                    self.create_post(post)
                    os.system('cls')
                    for bannerzinho in banner:
                        print(Fore.RED + " " * 15 + bannerzinho)

                    for item in menu:
                        print(item)

                print(f"{Fore.GREEN} Finalizado!!! Aguarde para voltar para o menu {Fore.RESET}")

            except KeyboardInterrupt:
                print(f"{Fore.YELLOW}Programa interrompido pelo usuário. Aguarde para voltar {Fore.RESET}")
                
            finally:
                self.driver.quit()
             
            
    def log_in(self):
        email_input = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.ID, "email")))
        email_input.send_keys(self.email)
        password_input = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.ID, "pass")))
        password_input.send_keys(self.password)
        login_button = WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//*[@type='submit']")))
        login_button.click()
        

    def move_from_home_to_marketplace_create_item(self):
        WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//a[@aria-label="Facebook"]')))
        self.driver.get(self.marketplace_url)

    def fotos(self, folder):
        script_directory = get_script_folder()
        fotos = os.path.join(script_directory, "fotos")
        image_path = os.path.join(fotos, folder)

        extensoes_permitidas = ['.jpeg', '.jpg', '.png']

        arquivos_imagem = []
        for extensao in extensoes_permitidas:
            padrao_busca = os.path.join(image_path, '*' + extensao)
            arquivos_imagem.extend(glob.glob(padrao_busca))

        print(f"{Fore.GREEN}Arquivos de imagem encontrados: {Fore.RESET}", arquivos_imagem)

        if not arquivos_imagem:
            print("Nenhum arquivo de imagem encontrado.")
            return
        
        arquivos_formatados = "\n".join([os.path.abspath(arquivo) for arquivo in arquivos_imagem])

        upload_input_xpath = '//input[@type="file"]'
        upload_input = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, upload_input_xpath)))
        upload_input.send_keys(arquivos_formatados)

        sleep(self.time_to_sleep)

        
    def add_text_to_post(self, title, price):
        pyperclip.copy(title)
        title_input = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.XPATH, "//label[@aria-label='" + self.marketplace_options["labels"]["Title"] + "']/div/div/input")))
        title_input.click()
        title_input.send_keys(Keys.CONTROL, 'v')
        title_input.send_keys(Keys.ENTER)
        pyperclip.copy(price)
        price_input = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.XPATH, "//label[@aria-label='" + self.marketplace_options["labels"]["Price"] +  "']/div/div/input")))
        price_input.send_keys(Keys.CONTROL, 'v')
        price_input.send_keys(Keys.ENTER)

    def fetch_all_posts(self):
        posts = None
        try:
            script_directory = get_script_folder()
            data = os.path.join(script_directory, "data.xlsx")
            df = pd.read_excel(data, "ARES")
            posts = []
            posts.extend(df.to_records(index=False).tolist())
        except pd.errors.EmptyDataError as empty_error:
            print("O DataFrame está vazio. Nenhum dado encontrado.", empty_error)
        except Exception as e:
            print(f"Falha ao ler dados do arquivo Excel: {e}")
        
        return posts

    
    def create_post(self, post):
        self.fotos(str(post[8]))
        titulo = str(post[1])
        preço = str(post[2])
        descrição = str(post[7])
        try:
            print(f"{Fore.GREEN}Título:{Fore.RESET} {titulo}")
            self.add_text_to_post(titulo, preço)
        except Exception as e:
            print(f"Erro: {e}")
            self.driver.quit()
        try:
            category_input = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.XPATH, "//label[@aria-label='" + self.marketplace_options["labels"]["Category"] +  "']")))
            category_input.click()
            category_option = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.XPATH, "//div[@role='dialog']/div/div/div/span/div/div[" + self.get_element_position("categories", str(post[3])) + "]")))
            category_option.click()
        except Exception as e:
            print(f"Erro: {e}")
            self.driver.quit()
            
        try:
            state_input = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.XPATH, "//label[@aria-label='" + self.marketplace_options["labels"]["State"] +  "']")))
            state_input.click()
            state_option = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.XPATH, '//div[@role="listbox"]/div/div/div/div/div[1]/div/div[' + self.get_element_position("states", str(post[4])) + ']')))
            state_option.click()
        except Exception as e:
            print(f"Erro: {e}")
            self.driver.quit()
        try:
            print(f"{Fore.GREEN}Descrição:{Fore.RESET} {descrição}")
            pyperclip.copy(descrição)
            description_input = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.XPATH, "//label[@aria-label='Descrição']/div/div/textarea")))
            description_input.click()
            description_input.send_keys(Keys.CONTROL, 'v')
            description_input.send_keys(Keys.ENTER)
            
        
        except Exception as e:
            print(f"Erro: {e}")
            
        if post[5] == "platforms":
            type_input = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.XPATH, "//label[@aria-label='" + self.marketplace_options["labels"]["Platform"] +  "']")))
            type_input.click()
            type_option = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.XPATH, '//div[@role="menu"]/div/div/div[1]/div/div[' + self.get_element_position("platforms", post[6]) + ']')))
            type_option.click()
        
        if post[5] == "devices":
            type_input = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.XPATH, "//label[@aria-label='" + self.marketplace_options["labels"]["Device Name"] +  "']")))
            type_input.click()
            type_option = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.XPATH, '//div[@role="menu"]/div/div/div[1]/div/div[' + self.get_element_position("devices", post[6]) + ']')))
            type_option.click()
        try:
            next_button = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.XPATH, "//div[@aria-label='" + self.marketplace_options["labels"]["Next Button"] +  "']")))
            next_button.click()
        except Exception as e:
            print(f"Erro: {e}")
            self.driver.quit()
            
        pasta = str(post[8])
        self.excluir_pasta_linhas(pasta)
        self.post_in_more_places(str(post[9]))
        try:
            post_button = WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//div[@aria-label='" + self.marketplace_options["labels"]["Post"] +  "']")))
            post_button.click()
        except Exception as e:
            print(f"Erro: {e}")

        sleep(self.time_to_sleep)
        print(" ")


    def get_element_position(self, key, specific):
        if specific in self.marketplace_options[key]:
            return str(self.marketplace_options[key][specific])
        return -1


    def post_in_more_places(self, groups):
        groups_positions = groups.split(",")
        try:
            for group_position in groups_positions:
                group_input = WebDriverWait(self.driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[@aria-label='" + self.marketplace_options["labels"]["Marketplace"] +  "']/div/div/div/div[4]/div/div/div/div/div/div/div[2]/div[" + group_position + "]")))
                group_input.click()
        except Exception as a:
            print(a)
            sleep(self.time_to_sleep)

    def excluir_pasta_linhas(self, pasta):
        try:
            script_directory = get_script_folder()
            data = os.path.join(script_directory, "data.xlsx")
            df = pd.read_excel(data, "ARES")
            df = df.drop(0)
            df.to_excel(data, "ARES" ,index=False)
            print(f"{Fore.GREEN}ID do excel apagado{Fore.RESET}")
        except Exception as e:
            print(f"Erro: {e}")
   
        try:
            pasta_desejada = "fotos"
            script_directory = get_script_folder()
            caminho_pasta_desejada = os.path.join(script_directory, pasta_desejada)
            caminho_pasta = os.path.join(caminho_pasta_desejada, pasta)
            shutil.rmtree(caminho_pasta)
            print(f'{Fore.GREEN}Fotos de anúncio apagada{Fore.RESET}')
        except Exception as e:
            print(f"Erro: {e}")

            
def get_script_folder():
    if getattr(sys, 'frozen', False):
        script_path = os.path.dirname(sys.executable)
    else:
        script_path = os.path.dirname(
            os.path.abspath(sys.modules['__main__'].__file__)
        )
    return script_path


def depedencias():
    script_directory = get_script_folder()
    config_file_path = os.path.join(script_directory, 'config.ini')

    if os.path.exists(config_file_path):
        return

    config = configparser.ConfigParser()
    config['FACEBOOK'] = {
        'email': "",
        'password': "",
        'main_url': 'https://www.facebook.com',
        'marketplace_url': 'https://web.facebook.com/marketplace/create/item',
        'marketplace_your_posts': 'https://www.facebook.com/marketplace/you/selling'
    }

    images_path = os.path.join(script_directory, 'fotos\\')

    binary_location = r'C:\Program Files\Google\Chrome\Application\chrome.exe'

    config['CONFIG'] = {
        'language': 'pt',
        'images_path': images_path,
        'binary_location': binary_location,
        'driver_location': 'chromedriver.exe',  
        'time_to_sleep': '0.6'
    }

    with open(config_file_path, 'w') as configfile:
        config.write(configfile)

    pasta_alvo = 'fotos'
    caminho_pasta = os.path.join(script_directory, pasta_alvo)

    if os.path.exists(caminho_pasta) and os.path.isdir(caminho_pasta):
        return
    else:
        os.makedirs(caminho_pasta)

    nome_arquivo_excel = 'data.xlsx'

    if os.path.exists(nome_arquivo_excel):
        return
    else:
        df = pd.DataFrame(columns=['id', 'titulo', 'preço', 'categoria', 'estado', 'tipo', 'opção', 'descrição', 'diretorio', 'grupos'])

        df.to_excel(nome_arquivo_excel, sheet_name="ARES" ,index=False)
        
def config():
    script_directory = get_script_folder()
    config = configparser.ConfigParser()

    email = input("Digite o seu email: ")
    password = input("Digite a sua senha: ")

    config['FACEBOOK'] = {
        'email': email,
        'password': password,
        'main_url': 'https://www.facebook.com',
        'marketplace_url': 'https://web.facebook.com/marketplace/create/item',
        'marketplace_your_posts': 'https://www.facebook.com/marketplace/you/selling'
    }

    config_file_path = os.path.join(script_directory, 'config.ini')
    images_path = os.path.join(script_directory, 'fotos\\')

    
    binary_location = r'C:\Program Files\Google\Chrome\Application\chrome.exe'
        
    config['CONFIG'] = {
            'language': 'pt',
            'images_path': images_path,
            'binary_location': binary_location,
            'driver_location': 'chromedriver.exe', 
            'time_to_sleep': '0.6'
    }    

    with open(config_file_path, 'w') as configfile:
        config.write(configfile)

    print("Arquivo config.ini criado com sucesso!")

def excluir():
    try:
        pasta_desejada = "fotos"

        script_directory = get_script_folder()
        caminho_pasta_desejada = os.path.join(script_directory, pasta_desejada)

        if os.path.exists(caminho_pasta_desejada):
            os.chdir(caminho_pasta_desejada) 

            pastas_existentes = [pasta for pasta in os.listdir() if os.path.isdir(pasta)]
            pastas_numeradas = [pasta for pasta in pastas_existentes if pasta.isdigit()]

            if pastas_numeradas:
                while True:
                    try:
                        quantidade_pastas_a_excluir = int(input(f"{Fore.RED}>{Fore.RESET} "))
                        break  
                    except ValueError:
                        print("{Fore.YELLOW}Isso não é um número. Tente novamente{Fore.RESET}")

                pastas_a_excluir = pastas_numeradas[:quantidade_pastas_a_excluir]

                if pastas_a_excluir:
                    for pasta in pastas_a_excluir:
                        caminho_pasta_a_excluir = os.path.join(caminho_pasta_desejada, pasta)
                        shutil.rmtree(caminho_pasta_a_excluir)
                        print(f"A pasta '{pasta}' foi excluída.")
                else:
                    print(f"{Fore.YELLOW}Não há pastas numeradas para excluir.{Fore.RESET}")
            else:
                print(f"{Fore.YELLOW}Não há pastas numeradas para excluir.{Fore.RESET}")
        else:
            print(f"{Fore.YELLOW}A pasta '{pasta_desejada}' não existe em {diretorio_atual}.{Fore.RESET}")
    except Exception as e:
        print(f"{Fore.RED}ERRO LOCALIZADO: {e}{Fore.RESET}")
    
def criar():
    try:
        pasta_desejada = "fotos"
        quantidade_pastas = 0
        while True:
            try:
                quantidade_pastas = int(input(f"{Fore.RED}>{Fore.RESET} "))
                break  
            except ValueError:
                print("Isso não é um número. Tente novamente")

        script_directory = get_script_folder()

     
        caminho_pasta_desejada = os.path.join(script_directory, pasta_desejada)

        if os.path.exists(caminho_pasta_desejada):
            os.chdir(caminho_pasta_desejada)  

            pastas_existentes = [pasta for pasta in os.listdir() if os.path.isdir(pasta)]
            pastas_numeradas = [pasta for pasta in pastas_existentes if pasta.startswith('pasta_')]

            if pastas_numeradas:
                ultima_pasta_numerada = max(pastas_numeradas, key=lambda x: int(x.split('_')[1]))
                numero_ultima_pasta = int(ultima_pasta_numerada.split('_')[1])
            else:
                numero_ultima_pasta = 0
                print(f"{Fore.YELLOW}Número de pasta já alcançada{Fore.RESET}")

        else:
            os.makedirs(caminho_pasta_desejada)
            os.chdir(caminho_pasta_desejada)  
            print(f"{Fore.YELLOW}Pasta Criada {Fore.RESET}")
            numero_ultima_pasta = 0

        for i in range(numero_ultima_pasta + 1, numero_ultima_pasta + quantidade_pastas + 1):
            nome_pasta = f'{i}'
            if not os.path.exists(nome_pasta):
                os.makedirs(nome_pasta)
                print(f"Criada a pasta '{nome_pasta}' em '{pasta_desejada}'")
    except:
        print(f"{Fore.RED}ERRO LOCALIZADO, TENTE NOVAMENTE{Fore.RESET}")

def menu():
    thread = threading.Thread(target=depedencias)
    thread.start()
    try:
        os.system('cls')
        time.sleep(0.1)
        banner = [
        "        .8.          8 888888888o.   8 8888888888      d888888o.",
        "       .888.         8 8888    `88.  8 8888          .`8888:' `88.",
        "      :88888.        8 8888     `88  8 8888          8.`8888.   Y8",
        "     . `88888.       8 8888     ,88  8 8888          `8.`8888.",
        "    .8. `88888.      8 8888.   ,88'  8 888888888888   `8.`8888.",
        "   .8`8. `88888.     8 888888888P'   8 8888            `8.`8888.",
        "  .8' `8. `88888.    8 8888`8b       8 8888             `8.`8888.",
        " .8'   `8. `88888.   8 8888 `8b.     8 8888         8b   `8.`8888.",
        ".888888888. `88888.  8 8888   `8b.   8 8888         `8b.  ;8.`8888",
        ".8'       `8. `88888.8 8888     `88. 8 888888888888  `Y8888P ,88P'"
    ]

                                        
        for bannerzinho in banner:
            print(Fore.RED +" " * 15+ bannerzinho)
            time.sleep(0.05) 
     
        time.sleep(1)
        
        menu = [
        f"",
        f"{Fore.GREEN}Seja Bem-Vindo ao ARES!{Fore.RESET}",
        " ",
        f"                    ___=[{Fore.RED}        ARES V1.O{Fore.RESET}       ]=_",
        f"                    ___=[{Fore.YELLOW}[*]{Fore.RESET} 1 - Iniciar programa]=_",
        f"                    ___=[{Fore.YELLOW}[*]{Fore.RESET} 2 - Criar Pasta     ]=_",
        f"                    ___=[{Fore.YELLOW}[*]{Fore.RESET} 3 - Excluir Pasta   ]=_",
        f"                    ___=[{Fore.YELLOW}[*]{Fore.RESET} 4 - Configurar      ]=_",
        f"                    ___=[{Fore.YELLOW}[*]{Fore.RESET} 5 - Limpar tela     ]=_",
        " ",
        f"Por favor, escolha a opção desejada digitando o número correspondente.",
        f"Agradecemos por confiar em nosso software para suas necessidades de upar anúncios. Tenha uma ótima experiência!",
        ""
    ]
        for letra in menu:
            print(letra)
            time.sleep(0.05)
            
        thread.join()

        while True:
            
            escolha = input(f"{Fore.RED}>{Fore.RESET} ")
            
            if escolha == '1':
                config_object = configparser.ConfigParser()
                config_object.read("config.ini")
                facebook = config_object["FACEBOOK"]
                configuration = config_object["CONFIG"]
                app = App(facebook["email"], facebook["password"], configuration["images_path"], configuration["language"], facebook["main_url"], facebook["marketplace_url"], configuration["binary_location"], configuration["driver_location"], configuration["time_to_sleep"])
                break
            elif escolha == '2':
                print(f"{Fore.YELLOW}Quantas pastas você deseja CRIAR?{Fore.RESET}")
                criar()
            elif escolha == '3':
                print(f"{Fore.YELLOW}Quantas pastas você deseja EXCLUIR?{Fore.RESET}")
                excluir()
            elif escolha == '4':
                config()
            elif escolha == '5':
                  
                break  
            else:
                print("Escolha inválida. Tente novamente.")
    except Exception as erro:            
        print(f"{Fore.RED}ERRO LOCALIZADO, TENTE NOVAMENTE ERRO:{erro}{Fore.RESET}")

if __name__ == "__main__":
    init()
    time.sleep(1)
    while True: 
        menu()
    else:
        print("ERRO")
    

