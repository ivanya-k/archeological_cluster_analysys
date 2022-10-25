from matplotlib import pyplot as plt
from scipy.cluster.hierarchy import dendrogram, linkage
import numpy as np
import pandas as pd
import scipy

def create_dendrogram(input_table, dendrogramm_capture, first_line, first_column, names_column, is_complex_name, postfix_column):
    tabl2 = pd.read_excel(input_table)
    a = tabl2.to_numpy()
    matrix = a[first_line - 2:,first_column - 1:]
    pam_names = a[first_line - 2:,names_column]
    if is_complex_name:
        #make labels for complex with number
        num_mogily = a[first_line - 2:,postfix_column]
        pam_labels=[]
        for i in range(len(pam_names)):
            if not pd.isna(pam_names[i]):
                name = pam_names[i]
            pam_labels.append(str(name)+' - '+ str(num_mogily[i]))
    else:
        pam_labels = pam_names
    
    # generate the linkage matrix
    Z = linkage(matrix, method = 'single', metric='dice')
    
    cond_den = scipy.spatial.distance.pdist(matrix, metric='dice')
    
    b = np.zeros([matrix.shape[0]+1,matrix.shape[0]+1])
    j0=2
    j = 2
    k = 1
    for i in range(len(cond_den)):
        if j == matrix.shape[0]+1:
            j0+=1
            j=j0
            k+=1
        b[k,j] = cond_den[i]
        j+=1
    den_matrix = b
    
    c=list(b)
    for i in range(len(c)):
        c[i] = list(c[i])
        
    for i in range(len(pam_labels)):    
        c[0][1+i]=c[1+i][0]=pam_labels[i]
        
    # calculate full dendrogram
    plt.figure(figsize=(25, 10))
    tit = dendrogramm_capture
    plt.title(tit, fontsize=20)
    plt.xlabel('Памятник', fontsize=20)
    plt.ylabel('Расстояние по Гауэру', fontsize=20)
    plt.rc('ytick', labelsize=10)
    dendrogram(
        Z,
        labels = pam_labels,
        leaf_rotation=90.,  # rotates the x axis labels
        leaf_font_size=15.,  # font size for the x axis labels
    )


    filepath = input_table[:-4] + ' дендрограмма'
    plt.savefig(filepath, bbox_inches='tight')
    plt.show()

    #color_threshold = 0.38
    
    df = pd.DataFrame(c)
    
    filepath = input_table[:-4] + ' таблица расстояний'+ '.xlsx'
    df.to_excel(filepath, index=False, header=False)
