import os
import time
import zipfile
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from pathlib import Path
from tkinter import messagebox


# Função para compactar arquivo em um arquivo zip
def zip_arquivo(arquivo):
    nome_zip = os.path.splitext(arquivo)[0] + '.zip'
    with zipfile.ZipFile(nome_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
        zipf.write(arquivo, os.path.basename(arquivo))
        
def modifica_arquivo_tdm(arquivo_txt):
    caminho = Path(arquivo_txt) 
    
    if caminho.exists():    
        try:
            with open(arquivo_txt, 'r', encoding='ansi') as arquivo:
                leitura = arquivo.readlines()                           
                novo_nome = f'{caminho.parent}\\{leitura[1][58:67]}_{leitura[1][3:23]}_TDM.txt'                                                
                arquivo.close() 
                               
            if novo_nome != '__TDM.txt':                    
                os.rename(arquivo_txt, novo_nome)  # Renomeia o arquivo
                zip_arquivo(novo_nome)  # Compacta o arquivo renomeado            
                os.remove(novo_nome)
            
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
                messagebox.showwarning(title='Zipador de Arquivo TDM',message=f'Erro: {e}')
            
# Caminho para a pasta que você deseja monitorar
caminho_pasta = 'D:\\Cessações de Uso'

# Inicializar o observador
observer = Observer()
observer.schedule(MonitorarPasta(), path=caminho_pasta, recursive=True)
observer.start()

try:
    while True:
        time.sleep(1)
except KeyboardInterrupt:
    observer.stop()
observer.join()
