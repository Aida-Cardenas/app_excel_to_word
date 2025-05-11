import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from docx import Document
import os

class ContratosApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Generador de Contratos")
        self.root.geometry("600x400")
        
        self.excel_path = tk.StringVar()
        self.template_path = tk.StringVar()
        self.output_dir = tk.StringVar()
        
        # interfaz
        self.create_widgets()
    
    def create_widgets(self):
        # cuadro principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # seleccion excel
        ttk.Label(main_frame, text="Archivo Excel:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.excel_path, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(main_frame, text="Buscar", command=self.browse_excel).grid(row=0, column=2)
        
        # seleccion word
        ttk.Label(main_frame, text="Plantilla Word:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.template_path, width=50).grid(row=1, column=1, padx=5)
        ttk.Button(main_frame, text="Buscar", command=self.browse_template).grid(row=1, column=2)
        
        # crear carpeta para guardar los contratos
        ttk.Label(main_frame, text="Carpeta de salida:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.output_dir, width=50).grid(row=2, column=1, padx=5)
        ttk.Button(main_frame, text="Buscar", command=self.browse_output).grid(row=2, column=2)
        
        # boton de "generar contratos"
        ttk.Button(main_frame, text="Generar Contratos", command=self.generate_contracts).grid(row=3, column=1, pady=20)
        
        # barra progreso por si son muchos y tarda
        self.progress = ttk.Progressbar(main_frame, length=400, mode='determinate')
        self.progress.grid(row=4, column=0, columnspan=3, pady=10)
    
    def browse_excel(self):
        filename = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if filename:
            self.excel_path.set(filename)
    
    def browse_template(self):
        filename = filedialog.askopenfilename(
            filetypes=[("Word files", "*.docx")]
        )
        if filename:
            self.template_path.set(filename)
    
    def browse_output(self):
        dirname = filedialog.askdirectory()
        if dirname:
            self.output_dir.set(dirname)
    
    def generate_contracts(self):
        if not all([self.excel_path.get(), self.template_path.get(), self.output_dir.get()]):
            messagebox.showerror("Error", "Por favor, complete todos los campos")
            return
        
        try:
            # leer excel
            df = pd.read_excel(self.excel_path.get())
            
            # Verificar columnas requeridas
            required_columns = ['nombre', 'apellido', 'cedula', 'num_casas', 'prestamo']
            if not all(col in df.columns for col in required_columns):
                messagebox.showerror("Error", "El archivo Excel debe contener las columnas: nombre, apellido, cedula, num_casas, prestamo")
                return
            
            # barra de carga
            self.progress['maximum'] = len(df)
            self.progress['value'] = 0
            
            # leer cada fila
            for index, row in df.iterrows():
                # crear nuevo doc desde plantilla
                doc = Document(self.template_path.get())
                
                # reemplazar marcadores en el doc
                replacements = {
                    '{{nombre}}': str(row['nombre']),
                    '{{apellido}}': str(row['apellido']),
                    '{{cedula}}': str(row['cedula']),
                    '{{num_casas}}': str(row['num_casas']),
                    '{{prestamo}}': str(row['prestamo'])
                }
                
                for paragraph in doc.paragraphs:
                    for key, value in replacements.items():
                        if key in paragraph.text:
                            paragraph.text = paragraph.text.replace(key, value)
                
                # guardar doc
                output_file = os.path.join(
                    self.output_dir.get(),
                    f"contrato_{row['cedula']}.docx"
                )
                doc.save(output_file)
                
                # actualizar barra de progreso
                self.progress['value'] = index + 1
                self.root.update_idletasks()
            
            messagebox.showinfo("Éxito", "Contratos generados correctamente")
            
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error: {str(e)}")
        
        finally:
            self.progress['value'] = 0

if __name__ == "__main__":
    root = tk.Tk()
    app = ContratosApp(root)
    root.mainloop() 