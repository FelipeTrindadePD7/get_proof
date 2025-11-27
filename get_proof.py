import os
import re
import sys
import threading
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    os.system("pip install pandas openpyxl xlrd")
    import pandas as pd

try:
    import PyPDF2
except ImportError:
    os.system("pip install PyPDF2")
    import PyPDF2

try:
    import pdfplumber
except ImportError:
    os.system("pip install pdfplumber")
    import pdfplumber

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, scrolledtext
except ImportError:
    print("Erro: tkinter não instalado")
    sys.exit(1)


def normalize_account(conta):
    """Normaliza conta removendo caracteres. Ex: '52938-2' -> '529382'"""
    if conta is None:
        return ""
    return re.sub(r'[^0-9]', '', str(conta))


def extract_pdf_pages(pdf_path):
    """Extrai texto de cada página do PDF"""
    pages = {}
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            pages[i] = {'text': text, 'numbers': normalize_account(text)}
    return pages


def find_account_pages(conta, pages):
    """Busca conta nas páginas do PDF"""
    found = []
    conta_norm = normalize_account(conta)
    
    if not conta_norm or len(conta_norm) < 3:
        return found
    
    conta_sem_dv = conta_norm[:-1] if len(conta_norm) > 1 else conta_norm
    
    for num, data in pages.items():
        if conta_norm in data['numbers'] or conta_sem_dv in data['numbers'] or str(conta) in data['text']:
            found.append(num)
    
    return found


def create_pdf(pdf_path, page_numbers, output_path):
    """Cria PDF com páginas específicas"""
    if not page_numbers:
        return False
    
    try:
        with open(pdf_path, 'rb') as f:
            reader = PyPDF2.PdfReader(f)
            writer = PyPDF2.PdfWriter()
            
            for num in page_numbers:
                if num < len(reader.pages):
                    writer.add_page(reader.pages[num])
            
            if writer.pages:
                with open(output_path, 'wb') as out:
                    writer.write(out)
                return True
    except Exception as e:
        print(f"Erro criar PDF: {e}")
    return False


def clean_filename(name):
    """Remove caracteres inválidos"""
    if not name or str(name).lower() == 'nan':
        return "sem_nome"
    name = str(name)
    for c in '<>:"/\\|?*\n\r\t':
        name = name.replace(c, '_')
    return ' '.join(name.split())[:100].strip()


def find_column(df, names):
    """Encontra coluna pelo nome - busca exata primeiro, depois parcial"""
    # Primeira passada: busca exata
    for col in df.columns:
        for name in names:
            if str(col).lower().strip() == name.lower().strip():
                return col
    
    # Segunda passada: busca parcial
    for col in df.columns:
        for name in names:
            if name.lower() in str(col).lower():
                return col
    return None


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Extrator de Comprovantes")
        self.root.geometry("800x600")
        
        self.pdf_var = tk.StringVar()
        self.excel_var = tk.StringVar()
        self.out_var = tk.StringVar(value="comprovantes_extraidos")
        self.df = None
        self.conta_col = None
        self.nome_col = None
        self.ccusto_col = None
        self.last_dir = os.path.expanduser("~")  # Lembra último diretório
        
        self.setup_ui()
    
    def setup_ui(self):
        main = ttk.Frame(self.root, padding=10)
        main.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main, text="Extrator de Comprovantes", font=('Arial', 16, 'bold')).pack(pady=15)
        
        # Arquivos
        files = ttk.LabelFrame(main, text="📁 Arquivos", padding=15)
        files.pack(fill=tk.X, pady=5)
        
        # PDF
        f1 = ttk.Frame(files)
        f1.pack(fill=tk.X, pady=5)
        ttk.Label(f1, text="PDF Comprovantes:", width=18, font=('Arial', 10)).pack(side=tk.LEFT)
        pdf_entry = ttk.Entry(f1, textvariable=self.pdf_var, font=('Arial', 9))
        pdf_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        pdf_entry.bind('<Return>', lambda e: self.validate_pdf())
        ttk.Button(f1, text="Procurar...", width=12, command=self.get_pdf).pack(side=tk.LEFT)
        
        # Excel
        f2 = ttk.Frame(files)
        f2.pack(fill=tk.X, pady=5)
        ttk.Label(f2, text="Planilha Excel:", width=18, font=('Arial', 10)).pack(side=tk.LEFT)
        excel_entry = ttk.Entry(f2, textvariable=self.excel_var, font=('Arial', 9))
        excel_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        excel_entry.bind('<Return>', lambda e: self.validate_excel())
        ttk.Button(f2, text="Procurar...", width=12, command=self.get_excel).pack(side=tk.LEFT)
        
        # Saída
        f3 = ttk.Frame(files)
        f3.pack(fill=tk.X, pady=5)
        ttk.Label(f3, text="Pasta de Saída:", width=18, font=('Arial', 10)).pack(side=tk.LEFT)
        out_entry = ttk.Entry(f3, textvariable=self.out_var, font=('Arial', 9))
        out_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        out_entry.bind('<Return>', lambda e: self.validate_out())
        ttk.Button(f3, text="Procurar...", width=12, command=self.get_out).pack(side=tk.LEFT)
        
        # Botão
        btn_frame = ttk.Frame(main)
        btn_frame.pack(pady=15)
        self.btn = ttk.Button(btn_frame, text="▶ PROCESSAR COMPROVANTES", command=self.start, width=30)
        self.btn.pack()
        
        self.prog = ttk.Progressbar(main, mode='indeterminate', length=400)
        self.prog.pack(fill=tk.X, pady=5)
        
        # Log
        logf = ttk.LabelFrame(main, text="📋 Log de Processamento", padding=5)
        logf.pack(fill=tk.BOTH, expand=True, pady=5)
        self.log = scrolledtext.ScrolledText(logf, height=12, state='disabled', font=('Courier', 9))
        self.log.pack(fill=tk.BOTH, expand=True)
    
    def get_pdf(self):
        try:
            # Tenta usar zenity (GTK file chooser) se disponível
            import subprocess
            try:
                result = subprocess.run(
                    ['zenity', '--file-selection', '--title=Selecionar PDF de Comprovantes', '--file-filter=Arquivos PDF | *.pdf'],
                    capture_output=True,
                    text=True,
                    timeout=300
                )
                if result.returncode == 0:
                    f = result.stdout.strip()
                    if f and os.path.exists(f):
                        self.pdf_var.set(f)
                        self.last_dir = os.path.dirname(f)
                        self.write_log(f"✓ PDF: {os.path.basename(f)}")
                        return
            except:
                pass
            
            # Fallback para tkinter
            init_dir = self.last_dir if hasattr(self, 'last_dir') and os.path.exists(self.last_dir) else os.path.expanduser("~/Downloads")
            
            f = filedialog.askopenfilename(
                parent=self.root,
                title="Selecionar PDF de Comprovantes",
                filetypes=[
                    ("Arquivos PDF", "*.pdf"),
                    ("Todos os arquivos", "*.*")
                ],
                initialdir=init_dir
            )
            if f:
                self.pdf_var.set(f)
                self.last_dir = os.path.dirname(f)
                self.write_log(f"✓ PDF: {os.path.basename(f)}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao selecionar PDF: {e}")
    
    def get_excel(self):
        try:
            # Tenta usar zenity primeiro
            import subprocess
            try:
                result = subprocess.run(
                    ['zenity', '--file-selection', '--title=Selecionar Planilha Excel', '--file-filter=Arquivos Excel | *.xlsx *.xls'],
                    capture_output=True,
                    text=True,
                    timeout=300
                )
                if result.returncode == 0:
                    f = result.stdout.strip()
                    if f and os.path.exists(f):
                        self.excel_var.set(f)
                        self.last_dir = os.path.dirname(f)
                        self.write_log(f"✓ Excel: {os.path.basename(f)}")
                        self.load_excel(f)
                        return
            except:
                pass
            
            # Fallback para tkinter
            init_dir = self.last_dir if hasattr(self, 'last_dir') and os.path.exists(self.last_dir) else os.path.expanduser("~/Downloads")
            
            f = filedialog.askopenfilename(
                parent=self.root,
                title="Selecionar Planilha Excel",
                filetypes=[
                    ("Arquivos Excel", "*.xlsx"),
                    ("Arquivos Excel antigo", "*.xls"),
                    ("Todos os arquivos", "*.*")
                ],
                initialdir=init_dir
            )
            if f:
                self.excel_var.set(f)
                self.last_dir = os.path.dirname(f)
                self.write_log(f"✓ Excel: {os.path.basename(f)}")
                self.load_excel(f)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao selecionar Excel: {e}")
    
    def load_excel(self, path):
        try:
            self.df = pd.read_excel(path)
            cols = list(self.df.columns)
            
            # Auto-detectar colunas (hardcoded)
            self.conta_col = find_column(self.df, ['conta', 'account'])
            self.nome_col = find_column(self.df, ['nome social', 'nome', 'funcionario'])
            self.ccusto_col = find_column(self.df, ['descrição ccusto', 'descricao ccusto', 'descrição de ccusto', 'descricao de ccusto', 'desc ccusto', 'ccusto', 'centro de custo', 'setor'])
            
            self.write_log(f"Colunas: {len(cols)} | Registros: {len(self.df)}")
            self.write_log(f"✓ Detectadas: Conta={self.conta_col}, Nome={self.nome_col}, CCusto={self.ccusto_col}")
        except Exception as e:
            self.write_log(f"Erro: {e}")
    
    def get_out(self):
        try:
            # Tenta usar zenity primeiro
            import subprocess
            try:
                result = subprocess.run(
                    ['zenity', '--file-selection', '--directory', '--title=Selecionar Pasta de Saída'],
                    capture_output=True,
                    text=True,
                    timeout=300
                )
                if result.returncode == 0:
                    d = result.stdout.strip()
                    if d:
                        self.out_var.set(d)
                        self.last_dir = d
                        self.write_log(f"✓ Pasta: {d}")
                        return
            except:
                pass
            
            # Fallback para tkinter
            init_dir = self.last_dir if hasattr(self, 'last_dir') and os.path.exists(self.last_dir) else os.path.expanduser("~")
            
            d = filedialog.askdirectory(
                parent=self.root,
                title="Selecionar Pasta de Saída",
                initialdir=init_dir,
                mustexist=False
            )
            if d:
                self.out_var.set(d)
                self.last_dir = d
                self.write_log(f"✓ Pasta: {d}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao selecionar pasta: {e}")
    
    def validate_pdf(self):
        """Valida caminho do PDF digitado"""
        path = self.pdf_var.get().strip()
        if path and os.path.exists(path) and path.endswith('.pdf'):
            self.last_dir = os.path.dirname(path)
            self.write_log(f"✓ PDF: {os.path.basename(path)}")
        elif path:
            messagebox.showwarning("Aviso", "Arquivo PDF não encontrado!")
    
    def validate_excel(self):
        """Valida caminho do Excel digitado"""
        path = self.excel_var.get().strip()
        if path and os.path.exists(path) and (path.endswith('.xlsx') or path.endswith('.xls')):
            self.last_dir = os.path.dirname(path)
            self.write_log(f"✓ Excel: {os.path.basename(path)}")
            self.load_excel(path)
        elif path:
            messagebox.showwarning("Aviso", "Arquivo Excel não encontrado!")
    
    def validate_out(self):
        """Valida pasta de saída"""
        path = self.out_var.get().strip()
        if path:
            self.write_log(f"✓ Pasta: {path}")
    
    def write_log(self, msg):
        self.log.config(state='normal')
        self.log.insert(tk.END, msg + "\n")
        self.log.see(tk.END)
        self.log.config(state='disabled')
        self.root.update()
    
    def start(self):
        if not self.pdf_var.get() or not self.excel_var.get():
            messagebox.showerror("Erro", "Selecione PDF e Excel!")
            return
        if self.df is None:
            messagebox.showerror("Erro", "Carregue Excel!")
            return
        if not self.conta_col or not self.nome_col or not self.ccusto_col:
            messagebox.showerror("Erro", "Colunas não encontradas no Excel!\nVerifique se existem as colunas: Conta, Nome e Descrição Ccusto")
            return
        
        self.btn.config(state='disabled')
        self.prog.start()
        threading.Thread(target=self.process, daemon=True).start()
    
    def process(self):
        try:
            pdf_path = self.pdf_var.get()
            out_dir = self.out_var.get()
            conta_col = self.conta_col
            nome_col = self.nome_col
            ccusto_col = self.ccusto_col
            
            Path(out_dir).mkdir(parents=True, exist_ok=True)
            
            self.write_log("\n" + "="*50)
            self.write_log("Extraindo PDF...")
            pages = extract_pdf_pages(pdf_path)
            self.write_log(f"Total páginas: {len(pages)}")
            
            self.write_log("\nBuscando comprovantes...")
            ok = 0
            nok = 0
            
            for _, row in self.df.iterrows():
                conta = row[conta_col]
                nome = row[nome_col]
                ccusto = row[ccusto_col]
                
                if pd.isna(conta) or str(conta).strip() == '':
                    continue
                
                conta_str = str(conta).strip()
                nome_str = clean_filename(nome)
                ccusto_str = clean_filename(ccusto)
                
                paginas = find_account_pages(conta_str, pages)
                
                if paginas:
                    out = os.path.join(out_dir, f"{ccusto_str}_{nome_str}.pdf")
                    i = 1
                    while os.path.exists(out):
                        out = os.path.join(out_dir, f"{ccusto_str}_{nome_str}_{i}.pdf")
                        i += 1
                    
                    if create_pdf(pdf_path, paginas, out):
                        self.write_log(f"✓ {ccusto_str}_{nome_str} (pág {[p+1 for p in paginas]})")
                        ok += 1
                    else:
                        nok += 1
                else:
                    self.write_log(f"- {ccusto_str}_{nome_str} [{conta_str}]: não encontrado")
                    nok += 1
            
            self.write_log("\n" + "="*50)
            self.write_log(f"Extraídos: {ok} | Não encontrados: {nok}")
            
            self.root.after(0, lambda: messagebox.showinfo("OK", f"Concluído!\n✓ {ok}\n✗ {nok}"))
            
        except Exception as e:
            self.write_log(f"ERRO: {e}")
            self.root.after(0, lambda: messagebox.showerror("Erro", str(e)))
        finally:
            self.root.after(0, self.finish)
    
    def finish(self):
        self.prog.stop()
        self.btn.config(state='normal')


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
