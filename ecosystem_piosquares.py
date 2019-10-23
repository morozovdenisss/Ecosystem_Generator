import matplotlib.pyplot as plt
from matplotlib.offsetbox import TextArea, DrawingArea, OffsetImage, AnnotationBbox
from matplotlib.gridspec import GridSpec
import matplotlib.image as mpimg
import numpy as np
import pandas as pd
from xml.etree import ElementTree
import json, xlrd, csv, urllib.request

# Step 1 - Extract variables from Json file and create dictionaries
# Convert from Excel to CSV and remove unneeded columns,replace industries, where > 1
def csv_from_excel():
    wb = xlrd.open_workbook('pioneers.xls')
    sh = wb.sheet_by_index(0)
    csv_file = open('csv_file.csv', 'w')
    wr = csv.writer(csv_file, quoting=csv.QUOTE_ALL)
    for rownum in range(sh.nrows):
        wr.writerow(sh.row_values(rownum))
    csv_file.close()
    
def remove_cols():
    f=pd.read_csv("csv_file.csv")
    keep_col = ['Company Name', 'Domain', 'Logo', 'Your industry']
    new_f = f[keep_col]
    new_f.to_csv("new_csv.csv", index=False)
    
csv_from_excel()
remove_cols()
file = pd.read_csv('new_csv.csv')

def coma_remove(item):
    item = item.split(',', 1)[0].replace(',', '')
    return str(item)

file['Your industry'] = file['Your industry'].apply(coma_remove)
print(file.head(20))

global industries
industries = []
for i in file['Your industry']:
    if len(industries) == 4:
        break
    if i not in industries:
        industries.append(i)
print(industries)


#Step 2 - Plot the graphs, iterate dictionaries, get logos, input into table
class graph():
    def boxes(self):
        count = 1
        x, y  = 0.2, 0.2
        fig = plt.figure(constrained_layout=True)
        gs = GridSpec(2, 2, figure=fig)
        
        ax1 = fig.add_subplot(gs[0, 0], title = industries[0])
        file.loc[file['Your industry'] == industries[0]]
        for i in file['Logo']:
            urllib.request.urlretrieve(i, industries[0] + str(count) + ".png")
            arr_AS = mpimg.imread(industries[0] + str(count) + ".png")
            imagebox = OffsetImage(arr_AS, zoom=0.2)
            if count < 4:
                x += 0.3
            if count  == 4:
                x = 0.2
                y = 0.5
            if count > 4 and count < 7:
                x += 0.3
            if count  == 7:
                x = 0.2
                y = 0.8
            if count > 8 and count < 10:
                x += 0.3
            if count == 10:
                x, y  = 0.2, 0.2
                count = 1
                break
            AS = AnnotationBbox(imagebox, (x, y))
            count += 1
            ax1.add_artist(AS)
            
        ax2 = fig.add_subplot(gs[0, 1], title = industries[1])
        file.loc[file['Your industry'] == industries[1]]
        for i in file['Logo']:
            urllib.request.urlretrieve(i, industries[1] + str(count) + ".png")
            arr_AS = mpimg.imread(industries[1] + str(count) + ".png")
            imagebox = OffsetImage(arr_AS, zoom=0.2)
            if count < 4:
                x += 0.3
            if count  == 4:
                x = 0.2
                y = 0.5
            if count > 4 and count < 7:
                x += 0.3
            if count  == 7:
                x = 0.2
                y = 0.8
            if count > 8 and count < 10:
                x += 0.3
            if count == 10:
                x, y  = 0.2, 0.2
                count = 1
                break
            AS = AnnotationBbox(imagebox, (x, y))
            count += 1
            ax2.add_artist(AS)

        ax3 = fig.add_subplot(gs[1, 0], title = industries[2])
        ax4 = fig.add_subplot(gs[1, 1], title = industries[3])
        fig.suptitle("Ecosystem Map")
        for i, ax in enumerate(fig.axes):
            ax.tick_params(labelbottom=False, labelleft=False, bottom = False, left = False)
        plt.savefig('add_picture_matplotlib_figure.png',bbox_inches='tight')
        plt.show()


g = graph()
g.boxes()

# Step 3 - Create Text and input all files into Word document