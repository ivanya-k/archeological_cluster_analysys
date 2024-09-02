import tkinter as tk
from tkinter import ttk, filedialog as fd, messagebox
from matplotlib import pyplot as plt
from scipy.cluster.hierarchy import dendrogram, linkage
import numpy as np
import pandas as pd
import scipy
import os

# Define the fields and their default values
fields = ('Порог выделения цветом', 'Размер шрифта подписей оси X', 'Размер шрифта подписей оси Y', 'Размер шрифта названия оси X',
          'Размер шрифта названия оси Y', 'Размер картинки по оси Х', 'Размер картинки по оси Y')
fields_default = ('0.38', '15', '15', '20', '20', '25', '10')

# Global variable to store the file path
filepath_global = None

def select_file():
    global filepath_global
    file = fd.askopenfile(mode='r', filetypes=[('Excel files', '*.xls;*.xlsx'), ('All files', '*.*')])
    if file:
        filepath = os.path.abspath(file.name)
        filepath_global = filepath
        file_label.config(text=f"Выбранный файл: {os.path.basename(filepath)}")
        #messagebox.showinfo("Файл выбран", f"Файл находится по адресу: {filepath}")

def num_from_letter(letter):
    return (- (65-ord(letter)))

def toggle_postfix_field():
    if complex_name_var.get():
        postfix_entry.config(state='normal')
    else:
        postfix_entry.config(state='disabled')

def create_dendrogram(entries):
    try:
        if filepath_global is None:
            raise ValueError("Файл не выбран. Пожалуйста, выберите файл.")

        input_table = filepath_global
        file_base_name = os.path.splitext(os.path.basename(input_table))[0]
        dendrogramm_capture = str(caption_entry.get())
        title_fontsize = str(title_fontsize_entry.get())
        first_line = int(first_line_entry.get())
        first_column = num_from_letter(first_column_entry.get()) + 1
        names_column = num_from_letter(names_column_entry.get())
        is_complex_name = complex_name_var.get()
        postfix_column = num_from_letter(postfix_entry.get())
        color_threshold = float(entries['Порог выделения цветом'].get())
        xtick_fontsize = int(entries['Размер шрифта подписей оси X'].get())
        ytick_fontsize = int(entries['Размер шрифта подписей оси Y'].get())
        xlabel_fontsize = int(entries['Размер шрифта названия оси X'].get())
        ylabel_fontsize = int(entries['Размер шрифта названия оси Y'].get())
        figure_xsize = int(entries['Размер картинки по оси Х'].get())
        figure_ysize = int(entries['Размер картинки по оси Y'].get())

        # Read the Excel file
        if input_table.endswith('.xls'):
            tabl2 = pd.read_excel(input_table, engine='xlrd')
        else:
            tabl2 = pd.read_excel(input_table, engine='openpyxl')

        a = tabl2.to_numpy()
        matrix = a[first_line - 2:, first_column - 1:]
        pam_names = a[first_line - 2:, names_column]

        # Ensure matrix contains numeric data
        matrix = pd.DataFrame(matrix).apply(pd.to_numeric, errors='coerce').fillna(0).to_numpy()

        if is_complex_name:
            num_mogily = a[first_line - 2:, postfix_column]
            pam_labels = []
            for i in range(len(pam_names)):
                if not pd.isna(pam_names[i]):
                    name = pam_names[i]
                pam_labels.append(str(name) + ' - ' + str(int(num_mogily[i])))
        else:
            pam_labels = pam_names

        Z = linkage(matrix, method='single', metric='dice')
        cond_den = scipy.spatial.distance.pdist(matrix, metric='dice')
        b = np.zeros([matrix.shape[0] + 1, matrix.shape[0] + 1])
        j0 = 2
        j = 2
        k = 1
        for i in range(len(cond_den)):
            if j == matrix.shape[0] + 1:
                j0 += 1
                j = j0
                k += 1
            b[k, j] = cond_den[i]
            j += 1
        den_matrix = b

        c = list(b)
        for i in range(len(c)):
            c[i] = list(c[i])

        for i in range(len(pam_labels)):
            c[0][1 + i] = c[1 + i][0] = pam_labels[i]

        # Calculate full dendrogram
        plt.figure(figsize=(figure_xsize, figure_ysize))
        plt.title(dendrogramm_capture, fontsize=title_fontsize)
        plt.xlabel('Памятник', fontsize=xlabel_fontsize)
        plt.ylabel('Расстояние по Гауэру', fontsize=ylabel_fontsize)
        plt.rc('ytick', labelsize=ytick_fontsize)
        dendrogram(
            Z,
            labels=pam_labels,
            leaf_rotation=90.,  # rotates the x axis labels
            leaf_font_size=xtick_fontsize,  # font size for the x axis labels
            color_threshold=color_threshold
        )

        save_path = fd.asksaveasfilename(defaultextension=".png", initialfile=f"{file_base_name} дендрограмма.png", filetypes=[("PNG files", "*.png"), ("All files", "*.*")])
        if save_path:
            plt.savefig(save_path, bbox_inches='tight')
            #messagebox.showinfo("Успех", f"Дендрограмма сохранена в {save_path}")

        df = pd.DataFrame(c)
        save_path = fd.asksaveasfilename(defaultextension=".xlsx", initialfile=f"{file_base_name} таблица расстояний.xlsx", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if save_path:
            df.to_excel(save_path, index=False, header=False)
            #messagebox.showinfo("Успех", f"Таблица расстояний сохранена в {save_path}")

        messagebox.showinfo("Успех", "Процесс успешно завершен!")

    except Exception as e:
        messagebox.showerror("Ошибка", str(e))

def makeform(root, fields):
    entries = {}
    for i, field in enumerate(fields):
        row = tk.Frame(root)
        lab = tk.Label(row, width=32, text=field+": ", anchor='w')
        ent = tk.Entry(row, width=32)
        ent.insert(0, fields_default[i])
        row.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
        lab.pack(side=tk.LEFT)
        ent.pack(side=tk.RIGHT, expand=tk.YES, fill=tk.X)
        entries[field] = ent
    return entries


if __name__ == '__main__':
    root = tk.Tk()
    root.title('Генератор дендрограмм')

    open_button = ttk.Button(root, text='Выберите файл', command=select_file)
    open_button.pack(fill=tk.X, padx=5, pady=5)

    file_label = tk.Label(root, text="Файл не выбран", anchor='w')
    file_label.pack(fill=tk.X, padx=5, pady=5)

    # Add the entry for "Первая строчка с информацией"
    first_line_frame = tk.Frame(root)
    first_line_label = tk.Label(first_line_frame, width=32, text="Первая строчка с информацией: ", anchor='w')
    first_line_entry = tk.Entry(first_line_frame, width=32)
    first_line_entry.insert(0, '3')
    first_line_frame.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
    first_line_label.pack(side=tk.LEFT)
    first_line_entry.pack(side=tk.RIGHT, expand=tk.YES, fill=tk.X)

    # Add the entry for "Первый столбец с информацией"
    first_column_frame = tk.Frame(root)
    first_column_label = tk.Label(first_column_frame, width=32, text="Первый столбец с информацией: ", anchor='w')
    first_column_entry = tk.Entry(first_column_frame, width=32)
    first_column_entry.insert(0, 'C')
    first_column_frame.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
    first_column_label.pack(side=tk.LEFT)
    first_column_entry.pack(side=tk.RIGHT, expand=tk.YES, fill=tk.X)
    # Add the entry for "Столбец с именами объектов"
    names_column_frame = tk.Frame(root)
    names_column_label = tk.Label(names_column_frame, width=32, text="Столбец с именами объектов: ", anchor='w')
    names_column_entry = tk.Entry(names_column_frame, width=32)
    names_column_entry.insert(0, 'B')
    names_column_frame.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
    names_column_label.pack(side=tk.LEFT)
    names_column_entry.pack(side=tk.RIGHT, expand=tk.YES, fill=tk.X)

    # Add the Checkbutton for "Составное название признака?"
    complex_name_var = tk.BooleanVar()
    complex_name_checkbutton = tk.Checkbutton(root, text="Составное название объекта?", variable=complex_name_var, command=toggle_postfix_field)
    complex_name_checkbutton.pack(fill=tk.X, padx=5, pady=5)

    # Add the entry for "Столбец с постфиксами"
    postfix_frame = tk.Frame(root)
    postfix_label = tk.Label(postfix_frame, width=32, text="Столбец с постфиксами: ", anchor='w')
    postfix_entry = tk.Entry(postfix_frame, width=32)
    postfix_entry.insert(0, 'C')
    postfix_frame.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
    postfix_label.pack(side=tk.LEFT)
    postfix_entry.pack(side=tk.RIGHT, expand=tk.YES, fill=tk.X)
    postfix_entry.config(state='disabled')  # Initially disabled

    # Add a small separation
    separator = tk.Frame(root, height=2, bd=1, relief=tk.SUNKEN)
    separator.pack(fill=tk.X, padx=5, pady=10)

    # Add the text "Настройки отображения дендрограммы"
    settings_label = tk.Label(root, text="Настройки отображения дендрограммы", font=('Arial', 12, 'bold'))
    settings_label.pack(fill=tk.X, padx=5, pady=5)

    # Add the entry for "Подпись к рисунку"
    caption_frame = tk.Frame(root)
    caption_label = tk.Label(caption_frame, width=32, text="Подпись к рисунку: ", anchor='w')
    caption_entry = tk.Entry(caption_frame, width=32)
    caption_entry.insert(0, 'Подпись к рисунку')
    caption_frame.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
    caption_label.pack(side=tk.LEFT)
    caption_entry.pack(side=tk.RIGHT, expand=tk.YES, fill=tk.X)

    # Add the entry for "Размер шрифта подписи к рисунку"
    title_fontsize_frame = tk.Frame(root)
    title_fontsize_label = tk.Label(title_fontsize_frame, width=32, text="Размер шрифта подписи к рисунку: ", anchor='w')
    title_fontsize_entry = tk.Entry(title_fontsize_frame, width=32)
    title_fontsize_entry.insert(0, '20')
    title_fontsize_frame.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
    title_fontsize_label.pack(side=tk.LEFT)
    title_fontsize_entry.pack(side=tk.RIGHT, expand=tk.YES, fill=tk.X)

    # Add the remaining entries
    ents = makeform(root, fields)

    b1 = tk.Button(root, text='Сгенерировать дендрограмму и таблицу расстояний', command=lambda e=ents: create_dendrogram(e))
    b1.pack(side=tk.LEFT, padx=5, pady=5)

    b3 = tk.Button(root, text='Выход', command=root.destroy)
    b3.pack(side=tk.LEFT, padx=5, pady=5)

    root.mainloop()