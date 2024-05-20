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
import sys
import ctypes
from PyQt5 import QtWidgets
from PyQt5 import QtWinExtras
ctypes.windll.shell32.ExtractIconW.restype = ctypes.c_size_t
from tkinter import filedialog
from plyer import notification

# Para Pegar o ícone do executável
def displayIcon(path: str, index=0):
    hicon = ctypes.windll.shell32.ExtractIconW(0, path, index)
    qpixmap = QtWinExtras.QtWin.fromHICON(hicon)
    ctypes.windll.user32.DestroyIcon(ctypes.c_void_p(hicon))
    name = os.path.expandvars('%appdata%\\zipador_tdm\\icone.ico')
    qpixmap.save(name, 'ICO')

# Encerrar a aplicação pelo Systray
def fechar_app(icon, item): 
    icon.visible = False 
    os._exit(0)    

# Abre Configuração pelo Systray
def selecionar_pasta():
    pasta_monitoramento = filedialog.askdirectory()
    if (pasta_monitoramento) != '':
        arquivo_log(f'Seleciona Nova Pasta de Monitoramento: {pasta_monitoramento}')
        with open (str(os.path.expandvars('%appdata%\\zipador_tdm\\settings.json')), 'w') as arquivo_json:
            arquivo_json.write(f'''[
    {{
        "caminho_pasta": "{pasta_monitoramento}"
    }}
]''')
        reiniciar_app()    

# Abrir o arquivo de Log pelo Systray
def abrir_arquivo_log():
    os.startfile(os.path.expandvars('%appdata%\\zipador_tdm\\Log.log'))    

# Reiniciar Aplicação
def reiniciar_app():
    python = sys.executable
    os.execl(python, python, *sys.argv)

# Mostrar notificação
def notificacao(titulo,mensagem):
    notification.notify(
        title=titulo,
        message=mensagem,
        app_name='Zipador de TDM',
        app_icon=os.path.expandvars('%appdata%\\zipador_tdm\\icone.ico'), 
        timeout=10
    )

# Criar ícone no Systray
def create_systray():
    # Tentar carregar uma imagem para o ícone do systray, caso não encontre cria uma em branco
    try:
        image = Image.open(os.path.expandvars('%appdata%\\zipador_tdm\\icone.ico'))
    except FileNotFoundError:
        app = QtWidgets.QApplication(sys.argv)  
        displayIcon(sys.executable)
        image = Image.open(os.path.expandvars('%appdata%\\zipador_tdm\\icone.ico')) 

    icon = pystray.Icon('zipador_tdm', image, f'Zipador de TDM - Monitorante Pasta: {ler_arquivo_json()}')
        
    abrir_log = item('Abrir Log', abrir_arquivo_log)
    selecionar = item('Selecionar Pasta Monitoramento', selecionar_pasta)
    reiniciar = item('Reiniciar', reiniciar_app)
    fechar = item('Encerrar', fechar_app)
    icon.menu = (abrir_log, selecionar, pystray.Menu.SEPARATOR, reiniciar, fechar,)

    icon.run()

# Classe para verificar se aplicação já está aberta
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

# Compactar arquivo em um arquivo zip
def zip_arquivo(arquivo):
    nome_zip = os.path.splitext(arquivo)[0] + '.zip'
    with zipfile.ZipFile(nome_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
        zipf.write(arquivo, os.path.basename(arquivo))

# Adicionar texto no arquivo de Log
def arquivo_log(texto):
    log = os.path.expandvars('%appdata%\\zipador_tdm\\Log.log')
    with open(log, 'a') as arquivo:
        arquivo.write(f'{(datetime.date.today()).strftime("%d/%m/%Y")} - {(datetime.datetime.now()).strftime("%H:%M")}\n{texto}\n\n') 
        arquivo.close()   
            
# Lê o arquivo txt e altera o nome para o padrão compacta e depois exclui o txt            
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
                os.rename(arquivo_txt, novo_nome) 
                zip_arquivo(novo_nome)            
                os.remove(novo_nome)
                nome_zip = os.path.splitext(novo_nome)[0] + '.zip'
                arquivo_log(f'Arquivo criado com Sucesso: {nome_zip}')
                notificacao('Compactação de TDM',f'Arquivo criado com Sucesso: {nome_zip}')
            else:
                arquivo_log(f'Arquivo não é TDM válido: {nome_zip}')
                notificacao('Problema ao Compactar', f'Arquivo não é TDM válido: {nome_zip}')
            
            return True
            
        except PermissionError as e:
            if e.errno == 13:
                return False   
            
# Classe que monitora quando têm um arquivo.txt novo na pasta
class MonitorarPasta(FileSystemEventHandler):
    def __init__(self):
        super().__init__()

    def on_created(self, event):
        if event.is_directory:
            return
        elif (event.src_path.lower()).endswith('.txt'):
            time.sleep(1) 
            try:
                while not modifica_arquivo_tdm(event.src_path):
                    time.sleep(10)
            except Exception as e:
                arquivo_log(f'Erro ao criar arquivo: {e}')

# Tenta ler o caminho da pasta de monitoramento da arquivo settings.json e inverte a \ para / para evitir erro no json
# Caso não encontre ele pega o caminho do executável
def ler_arquivo_json():
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
            arquivo_json.close()  
            return str(settings[0]['caminho_pasta']).replace('/','\\')
    except FileNotFoundError:
            arquivo_log(f'Arquivo de Configuração \'settings.json\' não Encontrado.\nCriado Novo Arquivo de Configuração!')
            with open (str(os.path.expandvars('%appdata%\\zipador_tdm\\settings.json')), 'w') as arquivo_json:
                arquivo_json.write(f'''[
    {{
        "caminho_pasta": "{os.getcwd()}"
    }}
]''')
            return ler_arquivo_json()                      

if __name__ == "__main__":
    # Nome único para a aplicação, para não executar 2 ao mesmo tempo
    app_name = "zipador_tdm" 
    instance = SingleInstance(app_name) 
       
    # Cria a pasta no %appdata%
    if not os.path.exists(os.path.expandvars('%appdata%\\zipador_tdm\\')):
        os.mkdir(os.path.expandvars('%appdata%\\zipador_tdm\\'))
    
    # Verificar se a aplicação já está aberta
    if instance.already_running():
        instance.cleanup()
    else:        
             
        # Inicia o monitoramento
        observer = Observer()          
        
        try:
            observer.schedule(MonitorarPasta(), path=ler_arquivo_json(), recursive=True)
            observer.start()
        except FileNotFoundError:
            arquivo_log(f'Pasta de monitoramento não encontrada!')
            instance.cleanup()
            observer.stop() 
            os._exit(0)           
        
        #Cria o icone no Systray    
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
            
        
