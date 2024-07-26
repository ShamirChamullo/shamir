import os
import tkinter as tk
from tkinter import filedialog, ttk
import pandas as pd
import re
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
import matplotlib.pyplot as plt
from openpyxl.drawing.image import Image

class ETLApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Proceso ETL")
        self.master.geometry("600x400")

        # Crear y configurar widgets
        self.folder_path = tk.StringVar()
        self.start_column = tk.StringVar()
        self.end_column = tk.StringVar()
        self.start_row = tk.StringVar()

        tk.Label(self.master, text="Carpeta de datos:").pack()
        tk.Entry(self.master, textvariable=self.folder_path, width=50).pack()
        tk.Button(self.master, text="Seleccionar carpeta", command=self.select_folder).pack()

        tk.Label(self.master, text="Columna inicial (ej. A):").pack()
        tk.Entry(self.master, textvariable=self.start_column).pack()

        tk.Label(self.master, text="Columna final (ej. P):").pack()
        tk.Entry(self.master, textvariable=self.end_column).pack()

        tk.Label(self.master, text="Fila inicial:").pack()
        tk.Entry(self.master, textvariable=self.start_row).pack()

        tk.Button(self.master, text="Procesar archivos", command=self.process_files).pack()

        self.progress_bar = ttk.Progressbar(self.master, orient="horizontal", length=200, mode="determinate")
        self.progress_bar.pack()

        self.result_text = tk.Text(self.master, height=10, width=70)
        self.result_text.pack()

    def select_folder(self):
        folder_selected = filedialog.askdirectory()
        self.folder_path.set(folder_selected)

    def process_files(self):
        try:
            folder = self.folder_path.get()
            start_col = self.start_column.get().upper()
            end_col = self.end_column.get().upper()
            start_row = int(self.start_row.get())

            start_col_index = column_index_from_string(start_col)
            end_col_index = column_index_from_string(end_col)

            # Obtener lista de archivos Excel en la carpeta
            excel_files = [f for f in os.listdir(folder) if f.endswith('.xlsx') and f.startswith('AvanceVentasINTI')]

            all_data = []
            total_files = len(excel_files)

            for i, file in enumerate(excel_files):
                file_path = os.path.join(folder, file)
                
                # Extraer año, mes y día del nombre del archivo
                match = re.search(r'AvanceVentasINTI\.(\d{4})\.(\d{2})\.(\d{2})\.', file)
                if match:
                    year, month, day = match.groups()
                else:
                    year, month, day = "", "", ""

                # Cargar el libro de trabajo
                wb = load_workbook(filename=file_path, read_only=True)
                sheet = wb['ITEM_O']

                # Obtener los datos de las celdas especificadas
                data = []
                for row in sheet.iter_rows(min_row=start_row, min_col=start_col_index, max_col=end_col_index):
                    data.append([cell.value for cell in row])

                # Crear DataFrame
                df = pd.DataFrame(data)
                
                # Añadir columnas de año, mes y día
                df['ANIO'] = year
                df['MES'] = month
                df['DIA'] = day

                all_data.append(df)

                # Actualizar barra de progreso
                self.progress_bar["value"] = (i + 1) / total_files * 100
                self.master.update_idletasks()

            # Combinar todos los DataFrames
            final_df = pd.concat(all_data, ignore_index=True)

            # Generar nombres de columnas basados en las letras de las columnas
            column_names = [get_column_letter(i) for i in range(start_col_index, end_col_index + 1)]
            column_names.extend(['ANIO', 'MES', 'DIA'])
            final_df.columns = column_names

            # Exportar a Excel en la misma carpeta de entrada
            output_path = os.path.join(folder, 'Out.xlsx')
            final_df.to_excel(output_path, index=False)

            # Generar gráficos y guardarlos en una hoja aparte
            self.generate_charts(final_df, output_path)

            # Mostrar resultados
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, f"Proceso completado. Archivo guardado en: {output_path}\n\n")
            self.result_text.insert(tk.END, final_df.head().to_string())

        except Exception as e:
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, f"Error: {str(e)}")

    def generate_charts(self, df, output_path):
        with pd.ExcelWriter(output_path, engine='openpyxl', mode='a') as writer:
            workbook = writer.book

            # Crear una nueva hoja para los gráficos
            charts_sheet = workbook.create_sheet(title='Charts')

            for column in df.columns:
                if df[column].dtype in ['int64', 'float64']:
                    # Generar histograma
                    plt.figure()
                    df[column].hist()
                    plt.title(f'Histograma de {column}')
                    hist_path = os.path.join(self.folder_path.get(), f'{column}_hist.png')
                    plt.savefig(hist_path)
                    plt.close()

                    # Agregar histograma al archivo Excel
                    img = Image(hist_path)
                    charts_sheet.add_image(img, f'A{charts_sheet.max_row + 2}')

                # Generar gráfico de torta
                if df[column].dtype == 'object' or df[column].dtype.name == 'category':
                    plt.figure()
                    df[column].value_counts().plot.pie(autopct='%1.1f%%')
                    plt.title(f'Torta de {column}')
                    pie_path = os.path.join(self.folder_path.get(), f'{column}_pie.png')
                    plt.savefig(pie_path)
                    plt.close()

                    # Agregar gráfico de torta al archivo Excel
                    img = Image(pie_path)
                    charts_sheet.add_image(img, f'A{charts_sheet.max_row + 20}')

if __name__ == "__main__":
    root = tk.Tk()
    app = ETLApp(root)
    root.mainloop()
