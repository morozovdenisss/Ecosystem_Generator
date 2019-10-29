import matplotlib.pyplot as plt
from matplotlib.offsetbox import OffsetImage, AnnotationBbox
from matplotlib.gridspec import GridSpec
import matplotlib.image as mpimg
import numpy as np
import pandas as pd
import xlrd, csv, requests, shutil, os
from docx import Document
from docx.shared import Inches
from PIL import Image
from docx.shared import Cm, Pt, Mm
from docx.enum.text import WD_BREAK
from docx.enum.style import WD_STYLE_TYPE

# Step 1 - Extract variables from Json file and create dictionaries
# Convert from Excel to CSV, remove unneeded columns, only leave 1 industry
def csv_from_excel():
    wb = xlrd.open_workbook('pioneers.xls')
    sh = wb.sheet_by_index(0)
    csv_file = open('csv_files/csv_file.csv', 'w')
    wr = csv.writer(csv_file, quoting=csv.QUOTE_ALL)
    for rownum in range(sh.nrows):
        wr.writerow(sh.row_values(rownum))
    csv_file.close()
    
def remove_cols():
    f=pd.read_csv("csv_files/csv_file.csv")
    keep_col = ['Company Name', 'Domain', 'Description', 'Logo', 'Your industry', 'What is the current stage of your startup?', 'Total funding received in €', 'Product focus']
    new_f = f[keep_col]
    new_f = new_f.dropna()
    new_f.to_csv("csv_files/new_csv.csv", index=False)
    
def coma_remove(item):
    item = item.split(',', 1)[0].replace(',', '')
    return str(item)

#Step 2 - Plot the graphs, iterate dictionaries, get logos, input into table
class graph():
    def boxes(self):        
        fig = plt.figure(constrained_layout=True, dpi=200)
        gs = GridSpec(2, 2, figure=fig)
        
        ax1 = fig.add_subplot(gs[0, 0], title = industries[0])
        industry = 0
        self.image_placer(ax1, industry)
            
        ax2 = fig.add_subplot(gs[0, 1], title = industries[1])
        industry = 1
        self.image_placer(ax2, industry)

        ax3 = fig.add_subplot(gs[1, 0], title = industries[2])
        industry = 2
        self.image_placer(ax3, industry)
        
        ax4 = fig.add_subplot(gs[1, 1], title = industries[3])
        industry = 3
        self.image_placer(ax4, industry)
        
        #fig.suptitle("Ecosystem Map")
        for i, ax in enumerate(fig.axes):
            ax.tick_params(labelbottom=False, labelleft=False, bottom = False, left = False)
        plt.savefig('images/ecosystem_map.png',bbox_inches='tight')
        plt.show()
        
    def stages_graph(self):
        plt.clf()
        plt.close()
        labels = ['Concept\n Stage', 'Seed\n Stage', 'Early\n Stage', 'Growth\n Stage', 'Established\n Business']
        real_labels = ['Concept Stage (got an idea)', 'Seed Stage (working on a product)', 'Early Stage (prototype ready and close to market)', 'Growth Stage (we\'re out there and making money)', 'Established Business (achieved break-even point operationally)']
        numbers = []
        x = np.arange(len(labels))
        for i in real_labels:
            if len(file[file['What is the current stage of your startup?'] == i]) > 0:
                numbers.append(len(file[file['What is the current stage of your startup?'] == i]))
        width = 0.7
        fig, ax = plt.subplots()
        rect = ax.bar(x, numbers, width, label='Startups')
        ax.set_ylabel('Number of Startups')
        #ax.set_title('Stages of Startups')
        ax.set_xticks(x)
        ax.set_xticklabels(labels)
        for i in rect:
            height = i.get_height()
            ax.annotate('{}'.format(height),
                        xy=(i.get_x() + i.get_width() / 2, height),
                        xytext=(0, 0.2),
                        textcoords="offset points",
                        ha='center', va='bottom')
        ax.legend()
        fig.tight_layout()
        plt.savefig('images/stages_bar.png', bbox_inches='tight')
        plt.show()
        
    def funding(self):
        plt.clf()
        plt.close()
        labels = ['25k', '75k', '125k', '200k', '300k', '500k', '800k', '1M', '1.5M', '2M', '2.5M', '3M', '5M', '10M', '10+M']
        real_labels = ['1-25k', '26 - 75k', '76 - 125k', '126 - 200k', '201 - 300k', '301 - 500k', '501 - 800k', '801k - 1 M', '1 - 1.5 M', '1.5 - 2 M', '2 - 2.5 M', '2.5 - 3 M', '3 - 5 M', '5 - 10 M', '10+ M']
        x = np.arange(len(labels))
        numbers = []
        for i in real_labels:
            if len(file[file['Total funding received in €'] == i]) > 0:
                numbers.append(len(file[file['Total funding received in €'] == i]))
            if len(file[file['Total funding received in €'] == i]) == 0:
                numbers.append(0)
        width = 0.5
        fig, ax = plt.subplots()
        rect = ax.bar(x, numbers, width, label='Funding')
        ax.set_ylabel('Number of Startups')
        #ax.set_title('Funding of Startups')
        ax.set_xticks(x)
        ax.set_xticklabels(labels)
        for i in rect:
            height = i.get_height()
            ax.annotate('{}'.format(height),
                        xy=(i.get_x() + i.get_width() / 2, height),
                        xytext=(0, 0.2),
                        textcoords="offset points",
                        ha='center', va='bottom')
        ax.legend()
        plt.tick_params(axis='x', which='major', labelsize = 9)
        fig.tight_layout()
        plt.savefig('images/funding_bar.png', bbox_inches='tight')
        plt.show()
        
    def product_focus(self):
        plt.clf()
        plt.close()
        labels = ['Software', 'Hardware', 'Other']
        real_labels = ['Software application', 'Physical product', 'Something else']
        x = np.arange(len(labels))
        numbers = []
        for i in real_labels:
            if len(file[file['Product focus'] == i]) > 0:
                numbers.append(len(file[file['Product focus'] == i]))
            if len(file[file['Product focus'] == i]) == 0:
                numbers.append(0)
        width = 0.5
        fig, ax = plt.subplots()
        rect = ax.bar(x, numbers, width, label='Product')
        ax.set_ylabel('Number of Startups')
        #ax.set_title('Funding of Startups')
        ax.set_xticks(x)
        ax.set_xticklabels(labels)
        for i in rect:
            height = i.get_height()
            ax.annotate('{}'.format(height),
                        xy=(i.get_x() + i.get_width() / 2, height),
                        xytext=(0, 0.2),
                        textcoords="offset points",
                        ha='center', va='bottom')
        ax.legend()
        plt.tick_params(axis='x', which='major', labelsize = 9)
        fig.tight_layout()
        plt.savefig('images/product_focus.png', bbox_inches='tight')
        plt.show()

    def image_placer(self, _, industry):
        count = 1
        new = file.loc[file['Your industry'] == industries[industry]]
        basewidth = 25
        for i in new['Logo']:
            png = i.split('/', 1)[-1].replace('/', '')
            r = requests.get(i, stream  = True)
            with open('logos/'+png, 'wb') as out_file:
                shutil.copyfileobj(r.raw, out_file) 
            del r
            img = Image.open('logos/'+png)
            wpercent = (basewidth/float(img.size[0]))
            hsize = int((float(img.size[1])*float(wpercent)))
            img = img.resize((basewidth,hsize), Image.ANTIALIAS)
            img.save('logos/'+png) 
            arr_AS = mpimg.imread('logos/'+png)
            imagebox = OffsetImage(arr_AS)
            if count == 1:
                x, y  = 0.2, 0.2
            if count > 1 and count < 4:
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
            _.add_artist(AS)

# Step 3 - Create Text and input all files and text into Word document 
#r.add_text('Ecosystem Report for {0}, {1}, {2} and {3}'.format(*industries))
def ecosystem_map():
    global document
    document = Document()
    p = document.add_paragraph()
    p.alignment = 1
    r = p.add_run()
    r.add_text('Ecosystem Report')
    r.font.size = Pt(20)
    r.bold = True
    p1 = document.add_paragraph()
    p1.alignment = 1
    r = p1.add_run()
    r.add_picture('images/ecosystem_map.png') 
    document.add_paragraph()

def stages():
    p = document.add_paragraph()
    p.alignment = 1
    r = p.add_run()
    r.add_text('Stages of Startups')
    r.font.size = Pt(14)
    r.bold = True
    document.add_paragraph()
    p1 = document.add_paragraph()
    p1.alignment = 1
    r = p1.add_run()
    r.add_picture('images/stages_bar.png', width=Inches(5), height=Inches(3)) 
    document.add_paragraph()

def funding():
    p = document.add_paragraph()
    p.alignment = 1
    r = p.add_run()
    r.add_text('Funding of Startups')
    r.font.size = Pt(14)
    r.bold = True
    document.add_paragraph()
    p1 = document.add_paragraph()
    p1.alignment = 1
    r = p1.add_run()
    r.add_picture('images/funding_bar.png', width=Inches(5), height=Inches(3)) 
    document.add_paragraph()

def product():
    p = document.add_paragraph()
    p.alignment = 1
    r = p.add_run()
    r.add_text('Product Focus')
    r.font.size = Pt(14)
    r.bold = True
    document.add_paragraph()
    p1 = document.add_paragraph()
    p1.alignment = 1
    r = p1.add_run()
    r.add_picture('images/product_focus.png', width=Inches(5), height=Inches(3)) 
    document.add_paragraph()

def set_column_width(column, width):
    column.width = width
    for cell in column.cells:
        cell.width = width

def table_maker(_):
    p = document.add_paragraph()
    p.alignment = 1
    r = p.add_run()
    r.font.size = Pt(12)
    r.bold = True
    r.add_text(industries[_])
    table = document.add_table(1, 3)
    heading_cells = table.rows[0].cells
    heading_cells[0].text = 'Company Name  '
    heading_cells[1].text = 'Domain  '
    heading_cells[2].text = 'Description'
    new = file.loc[file['Your industry'] == industries[_]]
    keep = ['Company Name', 'Domain', 'Description']
    for name, domain, descr in new[keep].itertuples(index=False): 
        cells = table.add_row().cells
        if len(name) > 16:
            name = name[0:16] + '..'
            cells[0].text = name
        if len(name) <= 16:
            cells[0].text = name
        cells[1].text = domain  
        if len(descr) > 60:
            descr = descr[0:60] + '...'
            cells[2].text = descr
        if len(descr) <= 60:
            cells[2].text = descr   
    set_column_width(table.columns[0], Inches(2.5))
    set_column_width(table.columns[1], Inches(2.5))
    set_column_width(table.columns[2], Inches(5))
    table.style = 'Table Grid'
    document.add_paragraph()

def launch():
    csv_from_excel()
    remove_cols()
    file = pd.read_csv('csv_files/new_csv.csv')
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
    g = graph()
    g.boxes()
    g.stages_graph()
    g.funding()
    g.product_focus()
    ecosystem_map()
    stages()
    funding()
    product()
    for i in range(0, 3):
        table_maker(i)
    document.save('Report.docx')

# Need to figure out how to change .docx into .pdf
launch()


