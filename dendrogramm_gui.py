import tkinter as tk
from matplotlib import pyplot as plt
from scipy.cluster.hierarchy import dendrogram, linkage
import numpy as np
import pandas as pd
import scipy

#modules for file selector
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
import os

#input_table = 'Матрица 57 Керамический комплекс закр комплексов один культ.xls'
#dendrogramm_capture = 'Дендрограмма 29.  Кластер-анализ керамического комплекса из закрытых комплексов одинцовской культуры.'
#first_line = 3
#first_column = 3
#names_column = 1
#is_complex_name = False
#postfix_column = names_column + 1


fields = ('Подпись к рисунку', 'Размер шрифта подписи к рисунку', 'Первая строчка с информацией', 'Первый столбец с информацией', 'Столбец с именами объектов',
          'Сложное название? (1-да, 0-нет)', 'Столбец с постфиксами', 'Порог выделения цветом', 'Размер шрифта подписей оси X',
          'Размер шрифта подписей оси Y','Размер шрифта названия оси X',
          'Размер шрифта названия оси Y', 'Размер картинки по оси Х', 'Размер картинки по оси Y', )
fields_default = ('Подпись к рисунку','20', '3', 'C', 'B',
          '0', 'C', '0.38','15',
          '15', '20',
          '20', '25', '10')

def select_file():
    global filepath_global
    file = fd.askopenfile(mode='r', filetypes=[('Only Excel .xls files', '*.xls'), ('All files', '*.*')])
    if file:
        filepath = os.path.abspath(file.name)
        filepath_global = filepath
        #Label(win, text="The File is located at : " + str(filepath), font=('Aerial 11')).pack()
        #entries['Имя файла'].append(filepath)

def num_from_letter(letter):
    return (- (65-ord(letter)))

def create_dendrogram(entries):

    #print(filepath_global)
    input_table = filepath_global #str(entries['Имя файла'].get())
    dendrogramm_capture = str(entries['Подпись к рисунку'].get())
    title_fontsize = str(entries['Размер шрифта подписи к рисунку'].get())
    first_line = int(entries['Первая строчка с информацией'].get())
    first_column = num_from_letter(entries['Первый столбец с информацией'].get()) + 1
    names_column = num_from_letter(entries['Столбец с именами объектов'].get())
    is_complex_name = bool(int(entries['Сложное название? (1-да, 0-нет)'].get()))
    postfix_column = num_from_letter(entries['Столбец с постфиксами'].get())
    color_threshold = float(entries['Порог выделения цветом'].get())
    xtick_fontsize = int(entries['Размер шрифта подписей оси X'].get())
    ytick_fontsize = int(entries['Размер шрифта подписей оси Y'].get())
    xlabel_fontsize = int(entries['Размер шрифта названия оси X'].get())
    ylabel_fontsize = int(entries['Размер шрифта названия оси Y'].get())
    figure_xsize = int(entries['Размер картинки по оси Х'].get())
    figure_ysize = int(entries['Размер картинки по оси Y'].get())


    tabl2 = pd.read_excel(input_table)
    a = tabl2.to_numpy()
    matrix = a[first_line - 2:, first_column - 1:]
    pam_names = a[first_line - 2:, names_column]
    if is_complex_name:
        # make labels for complex with number
        num_mogily = a[first_line - 2:, postfix_column]
        pam_labels = []
        for i in range(len(pam_names)):
            if not pd.isna(pam_names[i]):
                name = pam_names[i]
            pam_labels.append(str(name) + ' - ' + str(int(num_mogily[i])))
    else:
        pam_labels = pam_names

    # generate the linkage matrix
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

    # calculate full dendrogram
    plt.figure(figsize=(figure_xsize, figure_ysize))
    tit = dendrogramm_capture
    plt.title(tit, fontsize=title_fontsize)
    plt.xlabel('Памятник', fontsize = xlabel_fontsize)
    plt.ylabel('Расстояние по Гауэру', fontsize= ylabel_fontsize)
    plt.rc('ytick', labelsize = ytick_fontsize)
    dendrogram(
        Z,
        labels=pam_labels,
        leaf_rotation = 90.,  # rotates the x axis labels
        leaf_font_size = xtick_fontsize,  # font size for the x axis labels
        color_threshold = color_threshold
    )

    filepath = input_table[:-4] + ' дендрограмма'
    plt.savefig(filepath, bbox_inches='tight')
    #plt.show()

    df = pd.DataFrame(c)

    filepath = input_table[:-4] + ' таблица расстояний' + '.xlsx'
    df.to_excel(filepath, index=False, header=False)

def makeform(root, fields):
    entries = {}
    i = 0
    for field in fields:
        #print(field)
        row = tk.Frame(root)
        lab = tk.Label(row, width=32, text=field+": ", anchor='w')
        ent = tk.Entry(row, width=32)
        ent.insert(0, fields_default[i])
        row.pack(side=tk.TOP,
                 fill=tk.X,
                 padx=5,
                 pady=5)
        lab.pack(side=tk.LEFT)
        ent.pack(side=tk.RIGHT,
                 expand=tk.YES,
                 fill=tk.X)
        entries[field] = ent
        i += 1

    #entries['Подпись к рисунку'].insert(0, "This is the default text")

    return entries




if __name__ == '__main__':
    root = tk.Tk()

    open_button = ttk.Button(
        root,
        text='Выберите файл',
        command=select_file
    )

    open_button.pack(expand=True)

    ents = makeform(root, fields)

    #entries['Подпись к рисунку'].insert(0, "This is the default text")

    b1 = tk.Button(root, text='Сгенегировать дендрограмму и таблицу расстояний',
           command=(lambda e=ents: create_dendrogram(e)))
    b1.pack(side=tk.LEFT, padx=5, pady=5)



    b3 = tk.Button(root, text='Выход', command=root.quit)
    b3.pack(side=tk.LEFT, padx=5, pady=5)
    root.mainloop()