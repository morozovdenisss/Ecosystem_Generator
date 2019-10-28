import matplotlib.pyplot as plt
from matplotlib.offsetbox import TextArea, DrawingArea, OffsetImage, AnnotationBbox
import matplotlib.image as mpimg
import numpy as np
import pandas as pd
'''
file = pd.read_csv('EcoSystem.csv')
file.head()
#Iterate columns, get the relevant data from Sub-Type II
'''
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

