import matplotlib.pyplot as plt
from matplotlib.offsetbox import OffsetImage, AnnotationBbox, TextArea, DrawingArea 
import matplotlib.gridspec as gsa
import matplotlib.image as mpimg
import numpy as np
import pandas as pd
import matplotlib.image as mpimg
import numpy as np
import pandas as pd
import gmpy2
import os
import re
from PIL import Image

class graph():
    def adapt_csv(self):
        f=pd.read_csv("csv_files/csv_file.csv")
        keep_col = ['Startup Stage', 'Sub-Stage II', 'Name', 'Website']
        new_f = f[keep_col]
        new_f = new_f.dropna()
        new_f.to_csv("csv_files/as_new_csv.csv", index=True)
        
    def idea(self): 
        global directory_real
        global folders
        global count
        count = 0
        folders = ['01_01_Inspirational_Events', '01_01_Startup_Media', '01_02_Best_Practices', '01_02_Training_Feedback', '01_03_Build_First_Product', '01_03_Team_Formation', '02_01_Establish', '02_01_Workspace', '02_02_Formalize','02_02_Prepare for Seed',  '02_03_Pitch _ Demo', '02_03_Seed Accelerators', '03_01_Investor Networking', '03_01_Major Media', '03_02_Angels - Micro-VCs', '03_02_Venture Capitalists', '03_03_Expansion', '03_03_Infrastructure']
        for i in folders:
            if count < 6:
                stage = '01_Idea_Stage/'
            elif count > 5 and count < 12:
                stage = '02_Launch_Stage/'
            elif count > 11:
                stage = '03_Growth_Stage/'
            directory = os.getcwd() + '/logos/' + stage
            directory_real = directory + i + '/'
            self.image_placer(directory_real)
            count += 1
        
        fig = plt.figure(constrained_layout=False, dpi=300, figsize=(15,10))
        gs = gsa.GridSpec(22, 6, figure=fig)
        ax1 = fig.add_subplot(gs[0, 0:])
        ax1.text(0.5, 0.5, r'01. Idea Stage', fontsize=15, va="center", ha="center")
        ax2 = fig.add_subplot(gs[1, 0:2])
        ax2.text(0.5, 0.5, r'Inspire', fontsize=15, va="center", ha="center")
        ax3 = fig.add_subplot(gs[1, 2:4])
        ax3.text(0.5, 0.5, r'Educate', fontsize=15, va="center", ha="center")
        ax4 = fig.add_subplot(gs[1, 4:6])
        ax4.text(0.5, 0.5, r'Validate', fontsize=15, va="center", ha="center")

        for i in range(0,6):
            ax5 = fig.add_subplot(gs[3:7, i])
            directory_real = os.getcwd() + '/images/' + folders[i] + '.png'
            #self.place_created(ax5, directory_real, i)
            new = re.sub("\d", "", folders[i])
            new = re.sub("_", " ", new)
            ax5.set_title(new, fontsize = 8)
            img = Image.open(directory_real)
            ax5.imshow(img)
            ax5.axis('off')
        
        ax11 = fig.add_subplot(gs[7, 0:])
        ax11.text(0.5, 0.5, r'02. Launch Stage', fontsize=15, va="center", ha="center")
        ax12 = fig.add_subplot(gs[8, 0:2])
        ax12.text(0.5, 0.5, r'Start', fontsize=15, va="center", ha="center")
        ax13 = fig.add_subplot(gs[8, 2:4])
        ax13.text(0.5, 0.5, r'Develop', fontsize=15, va="center", ha="center")
        ax14 = fig.add_subplot(gs[8, 4:6])
        ax14.text(0.5, 0.5, r'Launch', fontsize=15, va="center", ha="center")
        
        count = 6
        for i in range(0,6):
            ax6 = fig.add_subplot(gs[10:14, i])
            directory_real = os.getcwd() + '/images/' + folders[count] + '.png'
            #self.place_created(ax5, directory_real, i)
            new = re.sub("\d", "", folders[count])
            new = re.sub("_", " ", new)
            ax6.set_title(new, fontsize = 8)
            img = Image.open(directory_real)
            ax6.imshow(img)
            ax6.axis('off')
            count+=1
        
        ax21 = fig.add_subplot(gs[15, 0:])
        ax21.text(0.5, 0.5, r'03. Growth Stage', fontsize=15, va="center", ha="center")
        ax22 = fig.add_subplot(gs[16, 0:2])
        ax22.text(0.5, 0.5, r'Recognition', fontsize=15, va="center", ha="center")
        ax23 = fig.add_subplot(gs[16, 2:4])
        ax23.text(0.5, 0.5, r'Funding', fontsize=15, va="center", ha="center")
        ax24 = fig.add_subplot(gs[16, 4:6])
        ax24.text(0.5, 0.5, r'Growth', fontsize=15, va="center", ha="center")
        
        count = 12
        for i in range(0,6):
            ax7 = fig.add_subplot(gs[18:, i])
            directory_real = os.getcwd() + '/images/' + folders[count] + '.png'
            #self.place_created(ax5, directory_real, i)
            new = re.sub("\d", "", folders[count])
            new = re.sub("_", " ", new)
            ax7.set_title(new, fontsize = 8)
            img = Image.open(directory_real)
            ax7.imshow(img)
            ax7.axis('off')
            count+=1
                
        #fig.suptitle("Ecosystem Map")
        for i, ax in enumerate(fig.axes):
            ax.tick_params(labelbottom=False, labelleft=False, bottom = False, left = False)
        plt.savefig('images/austrian_map.png',bbox_inches='tight')
        plt.show()
        
    def image_placer(self, directory_real):
        #_.set_title(industries[industry], fontsize = 8)
        filenames = []
        for i in os.listdir(directory_real):
            filename = os.fsdecode(i)
            if filename.endswith(('.jpeg', '.png', '.jpg', '.gif')): 
                filenames.append(filename)
        square, rem = gmpy2.isqrt_rem(len(filenames))
        fig1, axs = plt.subplots(square+1, square+1, dpi=300, figsize=(10, 10))
        axs = axs.flatten()
        for i, ax in zip(filenames,axs):
            img = np.array(Image.open(str(directory_real) + i))
            # If something doesn't work, delete the faulty logo
            #print(i)
            ax.axis('off')
            ax.imshow(img)
        for i, ax in enumerate(fig1.axes):
            ax.tick_params(labelbottom=False, labelleft=False, bottom = False, left = False)
            ax.axis('off')
        plt.savefig('images/' + folders[count] + '.png', bbox_inches='tight')
        plt.close
            
def launch():
    g = graph()
    g.idea()
launch()

'''
This code was intended forthe creationof PieChart 
- it is currently postponed in favor of a simpler graph.

file = pd.read_csv('EcoSystem.csv')
file.head()
#Iterate columns, get the relevant data from Sub-Type II
fig, ax = plt.subplots()

size = 0.3
vals = np.array([[60., 60.], [60., 60.], [60., 60.]])
vals2 = np.array([[60., 60., 60.], [60., 60., 60.], [60., 60., 60.]])
vals3 = np.array([[60., 60., 60., 60., 60., 60.], [60., 60., 60., 60., 60., 60.], [60., 60., 60., 60., 60., 60.]])

cmap = plt.get_cmap("tab20c")
layer1 = cmap(np.arange(3)*4)
labels1 = ['01 Idea', '02 Launch', '03 Growth']
layer2 = cmap(np.array([1, 2, 3, 5, 6, 7, 9, 10, 11]))
labels2 = ['Inspire', 'Educate', 'Validate', 'Start', 'Develop', 'Launch', 'Recognition', 'Finding', 'Growth']
#Need to fix the colors
layer3 = cmap(np.arange(6)*3)
#layer3 = cmap(np.array([0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]))

ax.pie(vals.sum(axis=1), radius=1-size, colors=layer1, labels = labels1, labeldistance = 0.92-size,
       wedgeprops=dict(width=size, edgecolor='w'))

ax.pie(vals2.flatten(), radius=1, colors=layer2, labels = labels2, labeldistance = 1.05  - size,
       wedgeprops=dict(width=size, edgecolor='w'))

ax.pie(vals3.flatten(), radius=1+size, colors=layer3,
       wedgeprops=dict(width=size, edgecolor='w'))

ax.pie(vals3.flatten(), radius=1+size*5, colors=layer3,
       wedgeprops=dict(width=size*4, edgecolor='w'))

ax.set(aspect="equal")

# Austrian Startups Logo
arr_AS = mpimg.imread('AS.jpg')
imagebox = OffsetImage(arr_AS, zoom=0.58)
AS = AnnotationBbox(imagebox, (0, 0))
ax.add_artist(AS)

# For each zone we need to ax.set, to set zone. How to make it non rectangle shape?
ax.set_xlim(0, 1)
ax.set_ylim(0, 1)
#  Defining png path - so I need to download it first
arr_lena = mpimg.imread('Lenna.jpg')
imagebox = OffsetImage(arr_lena, zoom=0.2)
# We define where this is placed, can iterate numbers  based  on number of images
ab = AnnotationBbox(imagebox, (0.4, 0.6))
ax.add_artist(ab)
plt.grid()
plt.draw()
plt.savefig('add_picture_matplotlib_figure.png',bbox_inches='tight')

# Helps to identify zones
xy = (-1.18, 0.685)
ax.plot(xy[0], xy[1], ".r")

xy = (0.5, 0.7)
ax.plot(xy[0], xy[1], ".r")

plt.show()
'''
