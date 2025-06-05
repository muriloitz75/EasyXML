import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import logging
import traceback
from datetime import datetime

# Configuração de logging
log_file = f"easyxml_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file),
        logging.StreamHandler(sys.stdout)
    ]
)

# Importa o módulo principal
try:
    import main as easyxml
    logging.info("Módulo EasyXML importado com sucesso")
except ImportError as e:
    logging.error(f"Erro ao importar o módulo EasyXML: {str(e)}")
    messagebox.showerror("Erro", f"Erro ao carregar o módulo principal: {str(e)}")
    sys.exit(1)

class EasyXMLApp:
    def __init__(self, root):
        self.root = root
        self.root.title("EasyXML - Processador de Notas Fiscais")
        self.root.geometry("600x400")
        self.root.resizable(True, True)
        
        # Configura o ícone se disponível
        try:
            self.root.iconbitmap("icon.ico")
        except:
            pass
        
        self.setup_ui()
    
    def setup_ui(self):
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Título
        title_label = ttk.Label(main_frame, text="EasyXML - Processador de Notas Fiscais", font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        
        # Descrição
        desc_label = ttk.Label(main_frame, text="Processa arquivos XML de notas fiscais e gera relatório Excel", font=("Arial", 10))
        desc_label.pack(pady=5)
        
        # Frame para seleção de diretório
        dir_frame = ttk.LabelFrame(main_frame, text="Diretório de XMLs", padding="10")
        dir_frame.pack(fill=tk.X, pady=10)
        
        self.dir_var = tk.StringVar(value=os.path.join(os.getcwd(), "nfs"))
        dir_entry = ttk.Entry(dir_frame, textvariable=self.dir_var, width=50)
        dir_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        browse_btn = ttk.Button(dir_frame, text="Procurar...", command=self.browse_directory)
        browse_btn.pack(side=tk.RIGHT)
        
        # Frame para saída
        output_frame = ttk.LabelFrame(main_frame, text="Diretório de Saída", padding="10")
        output_frame.pack(fill=tk.X, pady=10)
        
        self.output_var = tk.StringVar(value=os.path.join(os.getcwd(), "Notas_Processadas"))
        output_entry = ttk.Entry(output_frame, textvariable=self.output_var, width=50)
        output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        output_btn = ttk.Button(output_frame, text="Procurar...", command=self.browse_output)
        output_btn.pack(side=tk.RIGHT)
        
        # Barra de progresso
        progress_frame = ttk.Frame(main_frame)
        progress_frame.pack(fill=tk.X, pady=10)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X)
        
        self.status_var = tk.StringVar(value="Pronto para processar")
        status_label = ttk.Label(progress_frame, textvariable=self.status_var, font=("Arial", 9))
        status_label.pack(anchor=tk.W, pady=(5, 0))
        
        # Botões
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=10)
        
        process_btn = ttk.Button(btn_frame, text="Processar XMLs", command=self.process_xml)
        process_btn.pack(side=tk.LEFT, padx=5)
        
        open_output_btn = ttk.Button(btn_frame, text="Abrir Pasta de Saída", command=self.open_output_folder)
        open_output_btn.pack(side=tk.LEFT, padx=5)
        
        exit_btn = ttk.Button(btn_frame, text="Sair", command=self.root.destroy)
        exit_btn.pack(side=tk.RIGHT, padx=5)
        
        # Informações
        info_frame = ttk.Frame(main_frame)
        info_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        info_text = "Instruções:\n\n"
        info_text += "1. Selecione o diretório contendo os arquivos XML de notas fiscais\n"
        info_text += "2. Selecione o diretório de saída para os arquivos Excel\n"
        info_text += "3. Clique em 'Processar XMLs' para iniciar o processamento\n\n"
        info_text += "O aplicativo processará todos os arquivos XML e gerará um relatório Excel."
        
        info_label = ttk.Label(info_frame, text=info_text, justify=tk.LEFT)
        info_label.pack(anchor=tk.W)
        
        # Versão
        version_label = ttk.Label(main_frame, text="Versão 1.0", font=("Arial", 8))
        version_label.pack(side=tk.RIGHT, pady=(5, 0))
    
    def browse_directory(self):
        directory = filedialog.askdirectory(initialdir=self.dir_var.get())
        if directory:
            self.dir_var.set(directory)
    
    def browse_output(self):
        directory = filedialog.askdirectory(initialdir=self.output_var.get())
        if directory:
            self.output_var.set(directory)
    
    def open_output_folder(self):
        output_dir = self.output_var.get()
        if os.path.exists(output_dir):
            os.startfile(output_dir)
        else:
            messagebox.showinfo("Informação", "O diretório de saída ainda não existe.")
    
    def update_progress(self, value, status):
        self.progress_var.set(value)
        self.status_var.set(status)
        self.root.update_idletasks()
    
    def process_xml(self):
        input_dir = self.dir_var.get()
        output_dir = self.output_var.get()
        
        # Verifica se o diretório de entrada existe
        if not os.path.exists(input_dir):
            messagebox.showwarning("Aviso", f"O diretório de entrada '{input_dir}' não existe. Deseja criá-lo?")
            try:
                os.makedirs(input_dir)
                messagebox.showinfo("Informação", f"Diretório '{input_dir}' criado com sucesso.")
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível criar o diretório: {str(e)}")
                return
        
        # Verifica se o diretório de saída existe
        if not os.path.exists(output_dir):
            messagebox.showwarning("Aviso", f"O diretório de saída '{output_dir}' não existe. Deseja criá-lo?")
            try:
                os.makedirs(output_dir)
                messagebox.showinfo("Informação", f"Diretório '{output_dir}' criado com sucesso.")
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível criar o diretório: {str(e)}")
                return
        
        # Verifica se há arquivos XML no diretório
        xml_files = [f for f in os.listdir(input_dir) if f.lower().endswith('.xml')]
        if not xml_files:
            messagebox.showwarning("Aviso", f"Não foram encontrados arquivos XML no diretório '{input_dir}'.")
            return
        
        # Inicia o processamento em uma thread separada
        self.update_progress(0, f"Iniciando processamento de {len(xml_files)} arquivos XML...")
        
        def process_thread():
            try:
                # Modifica as variáveis globais para os diretórios selecionados
                original_nfs_dir = "nfs"
                original_output_dir = "Notas_Processadas"
                
                # Substitui os diretórios no código
                easyxml.diretorio_saida = output_dir
                
                # Copia os arquivos XML para o diretório nfs se necessário
                if input_dir != os.path.join(os.getcwd(), "nfs"):
                    # Se o diretório de entrada não for o padrão, copia os arquivos
                    if not os.path.exists("nfs"):
                        os.makedirs("nfs")
                    
                    self.update_progress(10, "Preparando arquivos XML...")
                    for i, xml_file in enumerate(xml_files):
                        src_path = os.path.join(input_dir, xml_file)
                        dst_path = os.path.join("nfs", xml_file)
                        
                        # Se o arquivo já existe no destino, não copia
                        if not os.path.exists(dst_path):
                            import shutil
                            shutil.copy2(src_path, dst_path)
                        
                        progress = 10 + (i / len(xml_files) * 20)
                        self.update_progress(progress, f"Preparando arquivo {i+1} de {len(xml_files)}...")
                
                # Executa o processamento
                self.update_progress(30, "Processando arquivos XML...")
                
                # Chama a função principal do módulo
                easyxml.main()
                
                self.update_progress(100, "Processamento concluído com sucesso!")
                messagebox.showinfo("Sucesso", "Processamento concluído com sucesso!")
                
            except Exception as e:
                logging.error(f"Erro durante o processamento: {str(e)}")
                logging.error(traceback.format_exc())
                self.update_progress(0, f"Erro: {str(e)}")
                messagebox.showerror("Erro", f"Ocorreu um erro durante o processamento: {str(e)}")
        
        threading.Thread(target=process_thread).start()

if __name__ == "__main__":
    root = tk.Tk()
    app = EasyXMLApp(root)
    root.mainloop()
