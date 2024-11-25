import os
import shutil
import psutil
import xml.etree.ElementTree as ET
import tkinter as tk
import threading
import subprocess
import json
import time
from PIL import Image, ImageTk  # Importar a biblioteca PIL
from tkinter import messagebox, filedialog
from tkinter import ttk

CONFIG_FILE = "config.json"


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

# Função para salvar as configurações em arquivo XML
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
    status_label['text'] = f"verificando a conectividade com o servidor '{ping_server}'...\n"
    try:
        response = subprocess.run(['ping', '-n', '1', ping_server], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        if response.returncode != 0:
            status_label['text'] = f"ERRO: Não foi possível conectar ao servidor '{ping_server}'.\nVerifique a conectividade e tente novamente."
            return False
        else:
            status_label['text'] = f"Conectividade com o servidor '{ping_server}'verificada com sucesso!"
            return True
    except Exception as e:
        status_label['text'] = f"Erro ao tentar verificar conectividade: {e}"
        return False

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

# Função para excluir arquivos antigos com simulação de tempo
def delete_old_files(destination_path, delete_files):
    print(f"\nExcluindo arquivos antigos em {destination_path}...\n")
    for file_path in delete_files:
        full_path = os.path.join(destination_path, file_path)
        try:
            if os.path.isdir(full_path):
                shutil.rmtree(full_path)
                status_label['text'] = f"Diretório {file_path} removido em {destination_path}."
            elif os.path.isfile(full_path):
                os.remove(full_path)
                status_label['text'] = f"Arquivo {file_path} removido em {destination_path}."
            else:
                status_label['text'] = f"Arquivo ou diretório não encontrado: {file_path}"
        except Exception as e:
            status_label['text'] = f"Erro ao remover arquivos antigos em {destination_path}: {e}"
        # Simula o tempo para exclusão
        time.sleep(0.5)

def copy_new_files_with_progress(source_path, destination_path):
    """
    Copia novos arquivos e pastas de uma origem para um destino, exibindo o progresso em uma barra e atualizando o status.
    """
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
                # Atualiza o status_label
                status_label['text'] = f"Copiando arquivo: {file_name}"
                root.update_idletasks()

                shutil.copy2(src_file, dest_file)
                copied_size += os.path.getsize(src_file)

                # Atualiza a barra de progresso
                progress = (copied_size / total_size) * 100
                progress_bar['value'] = progress
                root.update_idletasks()
            except Exception as e:
                status_label['text'] = f"Erro ao copiar o arquivo {src_file}: {e}"

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
                        # Atualiza o status_label
                        status_label['text'] = f"Copiando arquivo: {filename}"
                        root.update_idletasks()

                        shutil.copy2(src_file, dest_file)
                        copied_size += os.path.getsize(src_file)

                        # Atualiza a barra de progresso
                        progress = (copied_size / total_size) * 100
                        progress_bar['value'] = progress
                        root.update_idletasks()
                    except Exception as e:
                        status_label['text'] = f"Erro ao copiar o arquivo {src_file}: {e}"
    # Atualiza o status final
    status_label['text'] = "Cópia concluída!"

# Função para fechar o SistemaGDS se estiver em execução
def close_systemgds():
    status_label['text'] = "Alerta", "O SistemaGDS será fechado automaticamente, caso esteja aberto..."
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'].lower() == 'sistemagds.exe':
            status_label['text'] = f"O SistemaGDS está aberto. Fechando o processo (PID: {proc.info['pid']})..."
            proc.kill()
            return True
    return False

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

def load_config():
    """Carregar configurações do arquivo JSON."""
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    return {"autorun": False}  # Valor padrão

def save_config_json(config):
    """Salvar configurações no arquivo JSON."""
    with open(CONFIG_FILE, "w") as f:
        json.dump(config, f)

def set_autorun(enable):
    """Configura a inicialização automática e salva o estado no arquivo de configuração."""
    config = load_config()
    if enable:
        config["autorun"] = True  # Habilita o autorun
        messagebox.showinfo("Sucesso", "A aplicação será executada automaticamente ao abrir e atualizará o sistema.")
    else:
        config["autorun"] = False  # Desabilita o autorun
        messagebox.showinfo("Sucesso", "A inicialização automática foi desativada!")

    try:
        save_config_json(config)  # Salva a configuração no arquivo
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao salvar a configuração: {str(e)}")

def check_autorun():
    """Carrega o estado do autorun das configurações salvas."""
    config = load_config()
    return config.get("autorun", False)

def add_destination(destination_entries, frame_destinations):
    """Adiciona um campo de destino à tela com um botão para selecionar o caminho."""
    dest_frame = tk.Frame(frame_destinations)
    dest_frame.pack(padx=5, pady=5)

    # Campo para o destino
    dest_entry = tk.Entry(dest_frame, width=50)
    dest_entry.pack(side="left", padx=5)

    # Botão "Selecionar" para escolher o diretório do destino
    button_select_dest = tk.Button(dest_frame, text="Selecionar", width=12,
                                   command=lambda: browse_directory(dest_entry))
    button_select_dest.pack(side="left", padx=5)

    # Adiciona o campo de destino à lista
    destination_entries.append(dest_entry)

def browse_directory(entry_widget):
    """Abre uma janela de seleção de diretório e insere o caminho no campo de entrada."""
    selected_dir = filedialog.askdirectory()  # Abre o diálogo para selecionar diretório
    if selected_dir:
        entry_widget.delete(0, tk.END)  # Limpa o campo de entrada
        entry_widget.insert(0, selected_dir)  # Insere o diretório selecionado

def show_paths_from_xml():
    try:
        # Ler as configurações do arquivo XML
        config = read_config('updateConfig.xml')

        # Criar a janela para exibição
        paths_window = tk.Toplevel(root)
        paths_window.title("Caminhos Configurados")
        paths_window.geometry("350x350")  # Ajustando a altura para acomodar a logo e o texto

        # Remover botões de minimizar e maximizar
        paths_window.resizable(False, False)  # Desativa redimensionamento
        paths_window.attributes('-toolwindow', True)  # Remove minimizar/maximizar

        # Centralizar a janela na tela
        paths_window.update_idletasks()
        width = paths_window.winfo_width()
        height = paths_window.winfo_height()
        x = (paths_window.winfo_screenwidth() // 2) - (width // 2)
        y = (paths_window.winfo_screenheight() // 2) - (height // 2)
        paths_window.geometry(f"{width}x{height}+{x}+{y}")

        # Cabeçalho com logo e texto "Parâmetros Salvos"
        header_frame = tk.Frame(paths_window)
        header_frame.pack(fill="x", padx=10, pady=10)

        # Logo (substitua o caminho correto da imagem)
        logo_image = tk.PhotoImage(file="logo.png")  # Altere o caminho da imagem conforme necessário
        logo_label = tk.Label(header_frame, image=logo_image)
        logo_label.image = logo_image  # Necessário para manter a referência da imagem
        logo_label.pack(side="top", pady=5)

        # Texto do cabeçalho "Parâmetros Salvos"
        header_text = tk.Label(header_frame, text="Parâmetros Salvos", font=("Arial", 14, "bold"))
        header_text.pack(side="top", pady=5)

        # Exibir o servidor de ping
        label_ping_server = tk.Label(paths_window, text=f"Servidor de Ping: {config['pingServer']}", anchor="w", wraplength=320)
        label_ping_server.pack(fill="x", padx=10, pady=5)

        # Exibir o caminho de origem
        label_source_path = tk.Label(paths_window, text=f"Caminho de Origem: {config['sourcePath']}", anchor="w", wraplength=320)
        label_source_path.pack(fill="x", padx=10, pady=5)

        # Exibir os destinos
        label_destinations = tk.Label(paths_window, text="Destinos Configurados:", anchor="w")
        label_destinations.pack(fill="x", padx=10, pady=5)

        # Criar um frame com rolagem para destinos
        frame_scroll = tk.Frame(paths_window)
        frame_scroll.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        canvas = tk.Canvas(frame_scroll)
        scrollbar = ttk.Scrollbar(frame_scroll, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        for i, destination in enumerate(config['destinations'], start=1):
            label_destination = tk.Label(
                scrollable_frame,
                text=f"Destino {i}: {destination['path']}",
                anchor="w",
                wraplength=300
            )
            label_destination.pack(fill="x", padx=10, pady=5)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    except FileNotFoundError:
        messagebox.showerror("Erro", "Arquivo XML de configuração não encontrado!")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro ao exibir os caminhos: {e}")

def create_config_window(parent_window):
    # Cria a janela de configuração
    config_window = tk.Toplevel(parent_window)
    config_window.title("Configurar Caminho XML")
    config_window.transient(parent_window)  # Janela é filha da janela principal
    config_window.grab_set()  # Foco permanece nesta janela

    # Configuração do tamanho fixo da janela
    window_width = 500
    window_height = 520

    # Centralizar a janela
    screen_width = config_window.winfo_screenwidth()
    screen_height = config_window.winfo_screenheight()
    position_top = int((screen_height / 2) - (window_height / 2))
    position_left = int((screen_width / 2) - (window_width / 2))
    config_window.geometry(f"{window_width}x{window_height}+{position_left}+{position_top}")

    # Cabeçalho com logo e texto
    header_frame = tk.Frame(config_window)
    header_frame.pack(fill="x", padx=10, pady=10)

    # Logo GDS (substitua 'logo.png' pelo caminho correto)
    logo_image = tk.PhotoImage(file="logo.png")
    logo_label = tk.Label(header_frame, image=logo_image)
    logo_label.image = logo_image  # Necessário para evitar garbage collection
    logo_label.pack(side="top", pady=5)

    # Texto do cabeçalho
    header_text = tk.Label(header_frame, text="Parâmetros de Configuração", font=("Arial", 16))
    header_text.pack(side="top", pady=5)

    # Campo para servidor de ping
    fields_frame = tk.Frame(config_window)
    fields_frame.pack(padx=10, pady=10, fill="x")

    label_ping_server = tk.Label(fields_frame, text="Servidor Ping:")
    label_ping_server.grid(row=0, column=0, sticky="w", pady=5)
    entry_ping_server = tk.Entry(fields_frame, width=40)
    entry_ping_server.grid(row=0, column=1, pady=5, padx=5)

    # Campo para o caminho de origem
    label_source_path = tk.Label(fields_frame, text="Caminho de Origem:")
    label_source_path.grid(row=1, column=0, sticky="w", pady=5)
    entry_source_path = tk.Entry(fields_frame, width=40)
    entry_source_path.grid(row=1, column=1, pady=5, padx=5)
    button_select_source = tk.Button(fields_frame, text="Selecionar", width=12,
                                     command=lambda: browse_directory(entry_source_path))
    button_select_source.grid(row=1, column=2, padx=5)

    # Área para múltiplos destinos
    label_destinations = tk.Label(config_window, text="Destinos:")
    label_destinations.pack(padx=5, pady=5)

    destinations_frame = tk.Frame(config_window)
    destinations_frame.pack(fill="both", expand=True, padx=10, pady=5)

    # Configuração do Canvas e barra de rolagem com tamanho fixo
    canvas = tk.Canvas(destinations_frame, height=50)  # Altura fixa
    scrollbar = tk.Scrollbar(destinations_frame, orient="vertical", command=canvas.yview)
    scrollbar.pack(side="right", fill="y")

    frame_destinations = tk.Frame(canvas)
    canvas.create_window((0, 0), window=frame_destinations, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)

    def on_frame_configure(event):
        canvas.configure(scrollregion=canvas.bbox("all"))

    frame_destinations.bind("<Configure>", on_frame_configure)

    destination_entries = []

    # Adiciona pelo menos um campo de destino visível
    add_destination(destination_entries, frame_destinations)

    # Botão para adicionar novos destinos
    button_add_destination = tk.Button(config_window, text="Novo Destino", width=15,
                                       command=lambda: add_destination(destination_entries, frame_destinations))
    button_add_destination.pack(pady=5)

    # Frame para os botões Salvar e Fechar
    button_frame = tk.Frame(config_window)
    button_frame.pack(pady=10)

    # Botão Fechar
    button_cancel = tk.Button(button_frame, text="Fechar", width=15, command=config_window.destroy)
    button_cancel.pack(side="left", padx=10)

    # Botão Salvar
    button_save = tk.Button(button_frame, text="Salvar", width=15,
                            command=lambda: save_and_close(config_window, entry_ping_server, entry_source_path,
                                                           destination_entries))
    button_save.pack(side="left", padx=10)

def save_and_close(config_window, entry_ping_server, entry_source_path, destination_entries):
    # Função para salvar as configurações
    save_config({
        'pingServer': entry_ping_server.get(),
        'sourcePath': entry_source_path.get(),
        'destinations': [{'path': dest.get(), 'deleteFiles': ['SistemaGDS.exe', 'Temp\\*.mdb']} for dest in
                         destination_entries]
    }, 'updateConfig.xml')
    config_window.destroy()

# Função para verificar a senha
def check_password():
    # Crie a janela para solicitar a senha
    password_window = tk.Toplevel(root)
    password_window.title("Verificação de Senha")

    # Defina o tamanho da janela
    window_width = 300
    window_height = 150

    # Obtenha o tamanho da tela
    screen_width = password_window.winfo_screenwidth()
    screen_height = password_window.winfo_screenheight()

    # Calcule a posição para centralizar a janela
    position_top = int((screen_height / 2) - (window_height / 2))
    position_right = int((screen_width / 2) - (window_width / 2))

    # Defina a geometria com a posição calculada
    password_window.geometry(f"{window_width}x{window_height}+{position_right}+{position_top}")
    password_window.resizable(False, False)

    # Adiciona um título explicativo
    label = tk.Label(password_window, text="Digite a senha para acessar 'Configurar XML':")
    label.pack(pady=10)

    # Campo de entrada para a senha
    password_entry = tk.Entry(password_window, show="*", width=20)
    password_entry.pack(pady=5)

    # Função para verificar a senha
    def verify_password():
        password = password_entry.get()
        correct_password = "israel2015"  # Defina a senha que você deseja
        if password == correct_password:
            password_window.destroy()  # Fechar a janela de senha
            create_config_window(root)  # Chama a função para abrir a janela de configuração
        else:
            messagebox.showerror("Erro", "Senha incorreta. Tente novamente.")

    # Vincula a tecla Enter à função de verificação da senha
    password_window.bind('<Return>', lambda event: verify_password())  # Ao pressionar Enter, chama verify_password

    # Botão para verificar a senha
    button_verify = tk.Button(password_window, text="Entrar", command=verify_password)
    button_verify.pack(pady=10)

    # Focar o cursor na entrada da senha ao abrir a janela
    password_entry.focus()

def update_progress():
    # Simulação de progresso
    for value in range(101):  # De 0 a 100
        progress_bar['value'] = value  # Atualiza a barra
        progress_text_label.config(text=f"{value}%")  # Atualiza o texto
        root.update_idletasks()  # Atualiza a interface gráfica
        time.sleep(0.05)  # Simula o tempo de espera
    status_label.config(text="Atualização concluída!")  # Mensagem final

def create_main_window():
    global progress_bar, root, status_label, progress_text_label

    root = tk.Tk()
    root.title("Atualiza SistemaGDS")
    root.iconbitmap("icone.ico")

    # Remover botões de minimizar e maximizar
    root.resizable(False, False)  # Desabilita redimensionamento
    root.attributes('-toolwindow', True)  # Remove o botão de maximizar no Windows

    # Carregar o logo
    logo_image = Image.open("logo.png")  # Substitua "logo.png" pelo caminho do seu logo
    logo_image = logo_image.resize((73, 70), Image.Resampling.LANCZOS)  # Ajuste o tamanho do logo conforme necessário
    logo_photo = ImageTk.PhotoImage(logo_image)

    # Criar um label para o logo (acima dos botões)
    logo_label = tk.Label(root, image=logo_photo, bg="white")
    logo_label.pack(pady=10)  # Ajuste o espaçamento acima do logo se necessário

    # Frame para os botões lado a lado
    button_frame = tk.Frame(root)
    button_frame.pack(pady=10)

    # Largura padrão para os botões
    button_width = 20

    # Botão para atualizar sistema
    button_update_system = tk.Button(
        button_frame, text="Atualizar Sistema", width=button_width, command=update_system
    )
    button_update_system.pack(side="top", pady=5)

    # Botão para configurar XML (agora solicita a senha)
    button_configure = tk.Button(
        button_frame, text="Configurar XML", width=button_width, command=check_password
    )
    button_configure.pack(side="top", pady=5)

    # Botão para exibir caminhos do XML
    button_show_paths = tk.Button(
        button_frame, text="Exibir Caminhos do XML", width=button_width, command=show_paths_from_xml
    )
    button_show_paths.pack(side="top", pady=5)

    # Barra de progresso
    progress_bar = ttk.Progressbar(root, orient="horizontal", length=280, mode="determinate")
    progress_bar.pack(pady=10)

    # Status abaixo da barra de progresso
    status_label = tk.Label(root, text="", fg="red", font=("Arial", 8))
    status_label.pack(pady=5)

    # Verificar se a execução ao iniciar o Windows está habilitada
    autorun_var = tk.BooleanVar(value=check_autorun())
    checkbox_autorun = tk.Checkbutton(
        root,
        text="Autorun",
        variable=autorun_var,
        command=lambda: set_autorun(autorun_var.get()),
    )
    checkbox_autorun.pack(pady=10)

    # Executar "Atualizar Sistema" automaticamente se autorun estiver habilitado
    if check_autorun():
        update_system()

    # Centralizar a janela ao abrir
    window_width = 320
    window_height = 350
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    position_top = int((screen_height / 2) - (window_height / 2))
    position_right = int((screen_width / 2) - (window_width / 2))
    root.geometry(f"{window_width}x{window_height}+{position_right}+{position_top}")

    # Função para sair do programa
    def on_close():
        if messagebox.askokcancel("Sair", "Deseja realmente fechar o sistema?"):
            root.destroy()
            os._exit(0)  # Força o encerramento



    root.protocol("WM_DELETE_WINDOW", on_close)
    root.mainloop()

if __name__ == "__main__":
    create_main_window()

