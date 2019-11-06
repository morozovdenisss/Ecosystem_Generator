import matplotlib.pyplot as plt
from matplotlib.offsetbox import OffsetImage, AnnotationBbox, TextArea, DrawingArea 
from matplotlib.gridspec import GridSpec
import matplotlib.image as mpimg
import numpy as np
import pandas as pd
import matplotlib.image as mpimg
import numpy as np
import pandas as pd
import os
from PIL import Image

class graph():
    def idea(self): 
        fig = plt.figure(constrained_layout=False, dpi=300, figsize=(15,10))
        gs = GridSpec(18, 6, figure=fig)
        ax1 = fig.add_subplot(gs[0, 0:])
        ax1.text(0.5, 0.5, r'01. Idea Stage', fontsize=15, va="center", ha="center")
        ax2 = fig.add_subplot(gs[1, 0:2])
        ax2.text(0.5, 0.5, r'Inspire', fontsize=15, va="center", ha="center")
        ax3 = fig.add_subplot(gs[1, 2:4])
        ax3.text(0.5, 0.5, r'Educate', fontsize=15, va="center", ha="center")
        ax4 = fig.add_subplot(gs[1, 4:6])
        ax4.text(0.5, 0.5, r'Validate', fontsize=15, va="center", ha="center")
        ax5 = fig.add_subplot(gs[2:6, 0])
        directory = os.fsencode('/Users/denismorozov/Desktop/git/ecosystem_generator/images/as/')
        self.image_placer(ax5, directory)
        ax6 = fig.add_subplot(gs[2:6, 1])
        self.image_placer(ax6, directory)
        ax7 = fig.add_subplot(gs[2:6, 2])
        self.image_placer(ax7, directory)
        ax8 = fig.add_subplot(gs[2:6, 3])
        self.image_placer(ax8, directory)
        ax9 = fig.add_subplot(gs[2:6, 4])
        self.image_placer(ax9, directory)
        ax10 = fig.add_subplot(gs[2:6, 5])
        self.image_placer(ax10, directory)

        ax11 = fig.add_subplot(gs[6, 0:])
        ax11.text(0.5, 0.5, r'02. Launch Stage', fontsize=15, va="center", ha="center")
        ax12 = fig.add_subplot(gs[7, 0:2])
        ax12.text(0.5, 0.5, r'Start', fontsize=15, va="center", ha="center")
        ax13 = fig.add_subplot(gs[7, 2:4])
        ax13.text(0.5, 0.5, r'Develop', fontsize=15, va="center", ha="center")
        ax14 = fig.add_subplot(gs[7, 4:6])
        ax14.text(0.5, 0.5, r'Launch', fontsize=15, va="center", ha="center")
        ax15 = fig.add_subplot(gs[8:12, 0])
        directory = os.fsencode('/Users/denismorozov/Desktop/git/ecosystem_generator/images/as/')
        self.image_placer(ax15, directory)
        ax16 = fig.add_subplot(gs[8:12, 1])
        self.image_placer(ax16, directory)
        ax17 = fig.add_subplot(gs[8:12, 2])
        self.image_placer(ax17, directory)
        ax18 = fig.add_subplot(gs[8:12, 3])
        self.image_placer(ax18, directory)
        ax19 = fig.add_subplot(gs[8:12, 4])
        self.image_placer(ax19, directory)
        ax20 = fig.add_subplot(gs[8:12, 5])
        self.image_placer(ax20, directory)
        
        ax21 = fig.add_subplot(gs[12, 0:])
        ax21.text(0.5, 0.5, r'03. Growth Stage', fontsize=15, va="center", ha="center")
        ax22 = fig.add_subplot(gs[13, 0:2])
        ax22.text(0.5, 0.5, r'Recognition', fontsize=15, va="center", ha="center")
        ax23 = fig.add_subplot(gs[13, 2:4])
        ax23.text(0.5, 0.5, r'Funding', fontsize=15, va="center", ha="center")
        ax24 = fig.add_subplot(gs[13, 4:6])
        ax24.text(0.5, 0.5, r'Growth', fontsize=15, va="center", ha="center")
        ax25 = fig.add_subplot(gs[14:, 0])
        directory = os.fsencode('/Users/denismorozov/Desktop/git/ecosystem_generator/images/as/')
        self.image_placer(ax25, directory)
        ax26 = fig.add_subplot(gs[14:, 1])
        self.image_placer(ax26, directory)
        ax27 = fig.add_subplot(gs[14:, 2])
        self.image_placer(ax27, directory)
        ax28 = fig.add_subplot(gs[14:, 3])
        self.image_placer(ax28, directory)
        ax29 = fig.add_subplot(gs[14:, 4])
        self.image_placer(ax29, directory)
        ax30 = fig.add_subplot(gs[14:, 5])
        self.image_placer(ax30, directory)
    
        #fig.suptitle("Ecosystem Map")
        for i, ax in enumerate(fig.axes):
            ax.tick_params(labelbottom=False, labelleft=False, bottom = False, left = False)
        plt.savefig('images/austrian_map.png',bbox_inches='tight')
        plt.show()
        
    
    def image_placer(self, _, directory):
        #_.set_title(industries[industry], fontsize = 8)
        count = 1
        basewidth = 25
        filenames = []
        for i in os.listdir(directory):
            filename = os.fsdecode(i)
            if filename.endswith( ('.jpeg', '.png', '.jpg', '.gif') ): 
                filenames.append(filename)
        for i in filenames:
            img = Image.open('images/as/' + i)
            wpercent = (basewidth/float(img.size[0]))
            hsize = int((float(img.size[1])*float(wpercent)))
            img = img.resize((basewidth,hsize), Image.ANTIALIAS)
            img.save('images/as/' + i) 
            arr_AS = mpimg.imread('images/as/' + i)
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

def launch():
    g = graph()
    g.idea()
launch()




'''
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
