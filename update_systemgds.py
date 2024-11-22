import os
import shutil
import psutil
import xml.etree.ElementTree as ET
import tkinter as tk
import threading
import subprocess
import json
import ctypes
import sys

from tkinter import messagebox, filedialog
from tkinter import ttk

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if not is_admin():
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, " ".join(sys.argv), None, 1
    )
    sys.exit()

# Função para ler o arquivo XML e retornar as configurações
def read_config(xml_path):
    tree = ET.parse(xml_path)
    root = tree.getroot()

    config = {
        'pingServer': root.find('pingServer').text,
        'sourcePath': root.find('sourcePath').text,
        'destinations': []
    }

    # Lendo os destinos
    for destination in root.findall('destination'):
        destination_data = {
            'path': destination.find('destinationPath').text,
            'deleteFiles': ['SistemaGDS.exe', 'Temp\\*.mdb']  # Arquivos fixos para excluir
        }
        config['destinations'].append(destination_data)

    return config


# Função para salvar as configurações em um arquivo XML
def save_config(config, xml_path):
    configuracao = ET.Element("configuracao")

    ping_server_elem = ET.SubElement(configuracao, "pingServer")
    ping_server_elem.text = config['pingServer']

    source_path_elem = ET.SubElement(configuracao, "sourcePath")
    source_path_elem.text = config['sourcePath']

    for destination in config['destinations']:
        destination_elem = ET.SubElement(configuracao, "destination")

        destination_path_elem = ET.SubElement(destination_elem, "destinationPath")
        destination_path_elem.text = destination['path']

        delete_files_elem = ET.SubElement(destination_elem, "deleteFiles")
        for file in destination['deleteFiles']:
            file_elem = ET.SubElement(delete_files_elem, "file")
            file_elem.text = file

    tree = ET.ElementTree(configuracao)
    tree.write(xml_path)

    messagebox.showinfo("Sucesso", f"Arquivo XML '{xml_path}' gerado com sucesso!")


# Função para verificar a conectividade com o servidor
def check_connectivity(ping_server):
    print(f"O sistema vai verificar a conectividade com o servidor '{ping_server}'...\n")
    try:
        response = subprocess.run(['ping', '-n', '1', ping_server], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        if response.returncode != 0:
            print(f"ERRO: Não foi possível conectar ao servidor '{ping_server}'. Verifique a conectividade e tente novamente.\n")
            return False
        else:
            print(f"Conectividade com o servidor '{ping_server}' verificada com sucesso!\n")
            return True
    except Exception as e:
        print(f"Erro ao tentar verificar conectividade: {e}")
        return False


# Função para abrir o explorador de arquivos e selecionar um diretório
def browse_directory(entry):
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        entry.delete(0, tk.END)
        entry.insert(tk.END, folder_selected)


# Função para adicionar um destino
def add_destination(destination_entries):
    dest_frame = tk.Frame(frame_destinations)
    dest_frame.pack(fill=tk.X, pady=5)

    # Campo para o caminho do destino
    label_dest_path = tk.Label(dest_frame, text="Pasta Destino:")
    label_dest_path.pack(side=tk.LEFT, padx=5)
    entry_dest_path = tk.Entry(dest_frame, width=50)
    entry_dest_path.pack(side=tk.LEFT, padx=5)

    button_select_dest = tk.Button(dest_frame, text="Selecionar Destino", command=lambda: browse_directory(entry_dest_path))
    button_select_dest.pack(side=tk.LEFT, padx=5)

    destination_entries.append(entry_dest_path)


# Função para gerar o arquivo XML
def generate_xml(entry_ping_server, entry_source_path, destination_entries):
    config = {
        'pingServer': entry_ping_server.get(),
        'sourcePath': entry_source_path.get(),
        'destinations': []
    }

    for dest_entry in destination_entries:
        destination_data = {
            'path': dest_entry.get(),
            'deleteFiles': ['SistemaGDS.exe', 'Temp\\*.mdb']  # Arquivos fixos para excluir
        }
        config['destinations'].append(destination_data)

    save_config(config, 'updateConfig.xml')


# Função para executar a atualização do sistema
import time

# Função para calcular o tamanho total dos arquivos na origem
def calculate_total_size(source_path):
    total_size = 0
    for dirpath, dirnames, filenames in os.walk(source_path):
        for filename in filenames:
            file_path = os.path.join(dirpath, filename)
            if os.path.isfile(file_path):
                total_size += os.path.getsize(file_path)
    return total_size


# Função para excluir arquivos antigos com simulação de tempo
def delete_old_files(destination_path, delete_files):
    print(f"\nExcluindo arquivos antigos em {destination_path}...\n")
    for file_path in delete_files:
        full_path = os.path.join(destination_path, file_path)
        try:
            if os.path.isdir(full_path):
                shutil.rmtree(full_path)
                print(f"Diretório {file_path} removido em {destination_path}.")
            elif os.path.isfile(full_path):
                os.remove(full_path)
                print(f"Arquivo {file_path} removido em {destination_path}.")
            else:
                print(f"Arquivo ou diretório não encontrado: {file_path}")
        except Exception as e:
            print(f"Erro ao remover arquivos antigos em {destination_path}: {e}")
        # Simula o tempo para exclusão
        time.sleep(0.5)


def copy_new_files_with_progress(source_path, destination_path):
    print(f"\nCopiando os novos arquivos para {destination_path}...\n")

    # Caminhos específicos a serem copiados
    files_to_copy = ['SistemaGDS.exe']  # Arquivos fixos
    folders_to_copy = ['Relatorios']  # Pastas fixas
    dll_files_to_copy = []  # Lista para armazenar arquivos .dll

    # Adicionando todos os arquivos .dll da origem
    for file_name in os.listdir(source_path):
        if file_name.endswith('.dll'):
            dll_files_to_copy.append(file_name)

    files_to_copy.extend(dll_files_to_copy)

    # Calcular tamanho total dos arquivos a serem copiados
    total_size = 0
    for file_name in files_to_copy:
        file_path = os.path.join(source_path, file_name)
        if os.path.isfile(file_path):
            total_size += os.path.getsize(file_path)

    for folder_name in folders_to_copy:
        folder_path = os.path.join(source_path, folder_name)
        if os.path.isdir(folder_path):
            for dirpath, _, filenames in os.walk(folder_path):
                for filename in filenames:
                    file_path = os.path.join(dirpath, filename)
                    total_size += os.path.getsize(file_path)

    copied_size = 0

    # Copiar arquivos específicos
    for file_name in files_to_copy:
        src_file = os.path.join(source_path, file_name)
        dest_file = os.path.join(destination_path, file_name)

        if os.path.isfile(src_file):
            try:
                shutil.copy2(src_file, dest_file)
                copied_size += os.path.getsize(src_file)

                # Atualiza a barra de progresso
                progress = (copied_size / total_size) * 100
                progress_bar['value'] = progress
                root.update_idletasks()
                print(f"Arquivo copiado: {file_name}")
            except Exception as e:
                print(f"Erro ao copiar o arquivo {src_file}: {e}")

    # Copiar pasta "Relatorios"
    for folder_name in folders_to_copy:
        src_folder = os.path.join(source_path, folder_name)
        dest_folder = os.path.join(destination_path, folder_name)

        if os.path.isdir(src_folder):
            for dirpath, _, filenames in os.walk(src_folder):
                for filename in filenames:
                    src_file = os.path.join(dirpath, filename)
                    rel_path = os.path.relpath(src_file, src_folder)
                    dest_file = os.path.join(dest_folder, rel_path)

                    # Garante que o diretório de destino exista
                    os.makedirs(os.path.dirname(dest_file), exist_ok=True)

                    try:
                        shutil.copy2(src_file, dest_file)
                        copied_size += os.path.getsize(src_file)

                        # Atualiza a barra de progresso
                        progress = (copied_size / total_size) * 100
                        progress_bar['value'] = progress
                        root.update_idletasks()
                        print(f"Arquivo copiado: {filename}")
                    except Exception as e:
                        print(f"Erro ao copiar o arquivo {src_file}: {e}")

    print(f"Todos os arquivos necessários foram copiados para {destination_path}.")


# Função principal de atualização com estimativa de tempo
def update_system():
    def run_update():
        try:
            # Ler configurações do arquivo XML
            config = read_config('updateConfig.xml')

            # Verificar conectividade com o servidor
            if not check_connectivity(config['pingServer']):
                messagebox.showerror("Erro", "Não foi possível conectar ao servidor.")
                return

            # Fechar o SistemaGDS se estiver aberto
            close_systemgds()

            # Processar cada destino
            for destination in config['destinations']:
                # Excluir arquivos antigos
                delete_old_files(destination['path'], destination['deleteFiles'])

                # Copiar novos arquivos com progressão real
                copy_new_files_with_progress(config['sourcePath'], destination['path'])

            # Concluir a atualização
            messagebox.showinfo("Sucesso", "Atualização concluída com sucesso!")

        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro durante a atualização: {str(e)}")
            print(f"Erro: {str(e)}")

        # Resetar a barra de progresso
        progress_bar['value'] = 0

    # Cria um thread para rodar a função de atualização
    update_thread = threading.Thread(target=run_update)
    update_thread.start()


# Função para fechar o SistemaGDS se estiver em execução
def close_systemgds():
    print("\nO SistemaGDS será fechado automaticamente, caso esteja aberto...\n")
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'].lower() == 'sistemagds.exe':
            print(f"O SistemaGDS está aberto. Fechando o processo (PID: {proc.info['pid']})...")
            proc.kill()
            print("SistemaGDS fechado com sucesso.")
            return True
    print("SistemaGDS não está aberto.")
    return False


# Função para excluir arquivos antigos
def delete_old_files(destination_path, delete_files):
    print(f"\nExcluindo arquivos antigos em {destination_path}...\n")

    for file_path in delete_files:
        full_path = os.path.join(destination_path, file_path)
        try:
            if os.path.isdir(full_path):
                shutil.rmtree(full_path)
                print(f"Diretório {file_path} removido em {destination_path}.")
            elif os.path.isfile(full_path):
                os.remove(full_path)
                print(f"Arquivo {file_path} removido em {destination_path}.")
            else:
                print(f"Arquivo ou diretório não encontrado: {file_path}")
        except Exception as e:
            print(f"Erro ao remover arquivos antigos em {destination_path}: {e}")


# Função para sobrescrever a pasta 'Relatorios' se existir
def overwrite_relatorios(source_path, destination_path):
    print(f"\nSobrescrevendo a pasta 'Relatorios' em {destination_path}...\n")

    relatorios_path = os.path.join(destination_path, 'Relatorios')

    # Se a pasta 'Relatorios' existe, removemos
    if os.path.isdir(relatorios_path):
        try:
            shutil.rmtree(relatorios_path)
            print(f"Pasta 'Relatorios' removida em {destination_path}.")
        except Exception as e:
            print(f"Erro ao remover a pasta 'Relatorios' em {destination_path}: {e}")

    # Copia a nova pasta 'Relatorios'
    try:
        shutil.copytree(os.path.join(source_path, 'Relatorios'), relatorios_path)
        print(f"Pasta 'Relatorios' copiada com sucesso para {destination_path}.")
    except Exception as e:
        print(f"Erro ao copiar a pasta 'Relatorios' para {destination_path}: {e}")


# Função para copiar os arquivos necessários (SistemaGDS.exe e Relatorios)
def copy_new_files(source_path, destination_path):
    print(f"\nCopiando os novos arquivos para {destination_path}...\n")
    try:
        # Copiar o SistemaGDS.exe
        shutil.copy(os.path.join(source_path, 'SistemaGDS.exe'), destination_path)
        print(f"SistemaGDS.exe copiado com sucesso para {destination_path}.")

        # Sobrescrever a pasta Relatorios
        overwrite_relatorios(source_path, destination_path)
    except Exception as e:
        print(f"Erro ao copiar novos arquivos para {destination_path}: {e}")

CONFIG_FILE = "config.json"

def load_config():
    """Carregar configurações do arquivo JSON."""
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    return {"autorun": False}  # Valor padrão

def save_config(config):
    """Salvar configurações no arquivo JSON."""
    with open(CONFIG_FILE, "w") as f:
        json.dump(config, f)

def set_autorun(enable):
    task_name = "MyAppAutoStart"
    app_path = os.path.abspath(__file__)
    config = load_config()
    if enable:
        # Altere o argumento /SC ONLOGON para /SC ONSTART
        command = f'SchTasks /Create /SC ONSTART /TN "{task_name}" /TR "{app_path}" /F'
        config["autorun"] = True
    else:
        command = f'SchTasks /Delete /TN "{task_name}" /F'
        config["autorun"] = False

    try:
        subprocess.run(command, shell=True, check=True)
        save_config(config)
        if enable:
            messagebox.showinfo("Sucesso", "O aplicativo será executado sempre que o Windows for iniciado!")
        else:
            messagebox.showinfo("Sucesso", "A inicialização automática foi desativada!")
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Erro", f"Erro ao configurar a inicialização: {str(e)}")

def check_autorun():
    """Carregar estado da inicialização automática das configurações salvas."""
    config = load_config()
    return config.get("autorun", False)  # Valor padrão é False

def create_config_window(parent_window):
    config_window = tk.Toplevel(parent_window)
    config_window.title("Configurar Caminho XML")
    config_window.transient(parent_window)  # Janela é filha da janela principal
    config_window.grab_set()  # Foco permanece nesta janela

    # Campo para o servidor de ping
    label_ping_server = tk.Label(config_window, text="Servidor Ping:")
    label_ping_server.pack(padx=5, pady=5)
    entry_ping_server = tk.Entry(config_window, width=50)
    entry_ping_server.pack(padx=5)

    # Campo para o caminho de origem
    label_source_path = tk.Label(config_window, text="Caminho de Origem:")
    label_source_path.pack(padx=5, pady=5)
    entry_source_path = tk.Entry(config_window, width=50)
    entry_source_path.pack(padx=5)

    # Botão para selecionar o diretório de origem
    button_select_source = tk.Button(config_window, text="Selecionar Origem", command=lambda: browse_directory(entry_source_path))
    button_select_source.pack(pady=5)

    # Área para múltiplos destinos
    label_destinations = tk.Label(config_window, text="Destinos:")
    label_destinations.pack(padx=5, pady=5)

    # Frame para os destinos
    global frame_destinations
    frame_destinations = tk.Frame(config_window)
    frame_destinations.pack(padx=5, pady=5)

    destination_entries = []

    # Botão para adicionar novos destinos
    button_add_destination = tk.Button(config_window, text="Adicionar Destino", command=lambda: add_destination(destination_entries))
    button_add_destination.pack(pady=5)

    # Função de salvar e fechar a janela
    def save_and_close():
        save_config({
            'pingServer': entry_ping_server.get(),
            'sourcePath': entry_source_path.get(),
            'destinations': [{'path': dest.get(), 'deleteFiles': ['SistemaGDS.exe', 'Temp\\*.mdb']} for dest in
                             destination_entries]
        }, 'updateConfig.xml')
        config_window.destroy()  # Fecha a janela após salvar

    # Botão para salvar as configurações
    button_save = tk.Button(config_window, text="Salvar", command=save_and_close)
    button_save.pack(pady=10)

    # Botão para fechar a tela de configuração
    button_cancel = tk.Button(config_window, text="Fechar", command=config_window.destroy)
    button_cancel.pack(pady=10)

def create_main_window():
    global progress_bar, root, button_update_system

    root = tk.Tk()
    root.title("Atualiza SistemaGDS")
    root.iconbitmap("icone.ico")

    # Frame para os botões lado a lado
    button_frame = tk.Frame(root)
    button_frame.pack(pady=10)

    # Botão para atualizar sistema
    button_update_system = tk.Button(button_frame, text="Atualizar Sistema", command=update_system)
    button_update_system.pack(side="left", padx=5)

    # Botão para configurar XML
    button_configure = tk.Button(button_frame, text="Configurar XML", command=lambda: create_config_window(root))
    button_configure.pack(side="left", padx=5)

    # Barra de progresso
    progress_bar = ttk.Progressbar(root, orient="horizontal", length=400, mode="determinate")
    progress_bar.pack(pady=10)

    # Verificar se a execução ao iniciar o Windows está habilitada
    autorun_var = tk.BooleanVar(value=check_autorun())
    checkbox_autorun = tk.Checkbutton(
        root,
        text="Executar atualização ao iniciar o Windows",
        variable=autorun_var,
        command=lambda: set_autorun(autorun_var.get()),
    )
    checkbox_autorun.pack(pady=10)

    # Verifica se a execução automática foi ativada e se o script foi iniciado
    if check_autorun():
        # Simula o clique no botão de atualizar assim que a janela for carregada
        root.after(1000, button_update_system.invoke)

    # Função para sair do programa
    def on_close():
        if messagebox.askokcancel("Sair", "Deseja realmente fechar o sistema?"):
            root.destroy()
            os._exit(0)  # Força o encerramento

    root.protocol("WM_DELETE_WINDOW", on_close)
    root.geometry("600x180")
    root.mainloop()

if __name__ == "__main__":
    create_main_window()

