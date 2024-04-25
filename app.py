import os
import time
import zipfile
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from pathlib import Path
import datetime
import win32event
import win32api
import winerror
import json
import pystray
from pystray import MenuItem as item
from PIL import Image

# Função para encerrar a aplicação
def exit_action(icon, item): 
    icon.visible = False 
    os._exit(0)    

def abrir_arquivo_log():
    os.startfile(os.path.expandvars('%appdata%\\zipador_tdm\\Log.log'))    

# Criar ícone na bandeja do sistema
def create_systray():
    # Carregar uma imagem para o ícone
    try:
        image = Image.open(os.path.expandvars('%appdata%\\zipador_tdm\\icone.png'))
    except FileNotFoundError:
        blank_image = Image.new('RGBA', (16, 16), (0, 0, 0, 0))
        blank_image.save(os.path.expandvars('%appdata%\\zipador_tdm\\icone.png'), 'PNG')
        image = Image.open(os.path.expandvars('%appdata%\\zipador_tdm\\icone.png')) 

    # Criar o ícone na bandeja do sistema
    icon = pystray.Icon('zipador_tdm', image, 'Zipador de TDM')

    # Adicionar item de menu para sair
    abrir_log = item('Abrir Log', abrir_arquivo_log)
    exit_item = item('Encerrar', exit_action)
    icon.menu = (abrir_log, exit_item,)

    # Exibir o ícone
    icon.run()

class SingleInstance:
    def __init__(self, name):
        self.mutexname = name
        try:
            self.mutex = win32event.CreateMutex(None, False, self.mutexname)
        except:
            self.mutex = None

    def already_running(self):
        return win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS

    def cleanup(self):
        if self.mutex:
            win32api.CloseHandle(self.mutex)

# Função para compactar arquivo em um arquivo zip
def zip_arquivo(arquivo):
    nome_zip = os.path.splitext(arquivo)[0] + '.zip'
    with zipfile.ZipFile(nome_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
        zipf.write(arquivo, os.path.basename(arquivo))

def arquivo_log(texto):
    log = os.path.expandvars('%appdata%\\zipador_tdm\\Log.log')
    with open(log, 'a') as arquivo:
        arquivo.write(f'{(datetime.date.today()).strftime("%d/%m/%Y")} - {(datetime.datetime.now()).strftime("%H:%M")}\n{texto}\n\n') 
        arquivo.close()   
            
def modifica_arquivo_tdm(arquivo_txt):
    caminho = Path(arquivo_txt) 
    leitura = []
    
    if caminho.exists():    
        try:
            with open(arquivo_txt, 'r', encoding='ansi') as arquivo:
                for _ in range(2):
                    leitura.append(arquivo.readline())

                arquivo.close()

            novo_nome = f'{caminho.parent}\\{leitura[1][58:67]}_{leitura[1][3:23]}_TDM.txt'                                                    
            
            if (leitura[0].startswith('E01')) and novo_nome != '__TDM.txt':                                                 
                os.rename(arquivo_txt, novo_nome)  # Renomeia o arquivo
                zip_arquivo(novo_nome)  # Compacta o arquivo renomeado            
                os.remove(novo_nome)
                arquivo_log(f'Criado arquivo com Sucesso: {novo_nome}')
            else:
                arquivo_log(f'Arquivo não é TDM válido: {novo_nome}')
            
            return True
            
        except PermissionError as e:
            if e.errno == 13:
                return False   
            
# Classe para manipular os eventos do sistema de arquivos
class MonitorarPasta(FileSystemEventHandler):
    def __init__(self):
        super().__init__()

    def on_created(self, event):
        if event.is_directory:
            return
        elif event.src_path.endswith('.txt'): #and event.event_type == 'modified':  # Se o arquivo criado for um txt
            time.sleep(1)  # Pequeno atraso para garantir que o arquivo esteja completamente escrito no disco
            # Verificar se o arquivo foi modificado após o último processamento
            #if os.path.exists(event.src_path) and os.stat(event.src_path).st_mtime > self.ultimo_processamento:
            try:
                while not modifica_arquivo_tdm(event.src_path):
                    time.sleep(10)
            except Exception as e:
                arquivo_log(f'Erro ao criar arquivo: {e}')


if __name__ == "__main__":
    # Defina um nome único para o seu aplicativo
    app_name = "zipador_tdm"    
    
    caminho_pasta = ''
    
    instance = SingleInstance(app_name)
    
    if instance.already_running():
        instance.cleanup()
    else:        
        # Caminho para a pasta que você deseja monitorar
        try:
            with open (str(os.path.expandvars('%appdata%\\zipador_tdm\\settings.json')), encoding='ANSI', mode='r') as arquivo_json:
                settings_json = arquivo_json.read()
                settings_json = settings_json.replace('\\','/')  
                arquivo_json.close()
                with open (str(os.path.expandvars('%appdata%\\zipador_tdm\\settings.json')), 'w') as arquivo_json:
                    arquivo_json.write(settings_json)
                    arquivo_json.close()
                with open (str(os.path.expandvars('%appdata%\\zipador_tdm\\settings.json')), mode='r') as arquivo_json:                                
                    settings = json.load(arquivo_json)
                    caminho_pasta = str(settings[0]['caminho_pasta']).replace('/','\\')
                    arquivo_json.close()                    
        except FileNotFoundError:
            arquivo_log(f'Arquivo de Configuração \'settings.json\' não Encontrado.\nCriado Novo Arquivo de Configuração Zerado!')
            with open (str(os.path.expandvars('%appdata%\\zipador_tdm\\settings.json')), 'w') as arquivo_json:
                arquivo_json.write('''[{
  "caminho_pasta": ""
}]''')
            instance.cleanup()  
            os._exit(0) 
        
        observer = Observer()
        # Inicializar o observador  
        
        try:
            observer.schedule(MonitorarPasta(), path=caminho_pasta, recursive=True)
            observer.start()
        except FileNotFoundError:
            arquivo_log(f'Pasta de monitoramento não encontrada!')
            instance.cleanup()
            observer.stop() 
            os._exit(0)           
            
        create_systray()                   
    
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            observer.stop()
            instance.cleanup()
        except SystemExit:
            observer.stop()
            instance.cleanup()       
            
        observer.join()                                 
            
        
