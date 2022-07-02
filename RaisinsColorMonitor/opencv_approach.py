from sklearn.cluster import KMeans
import matplotlib.pyplot as plt
import numpy as np
import cv2
from collections import Counter
from skimage.color import rgb2lab, deltaE_cie76
import os
from time import sleep
from picamera import PiCamera
def RGB2HEX(color):
    return "#{:02x}{:02x}{:02x}".format(int(color[0]), int(color[1]), int(color[2]))
def get_image(image_path):
    image = cv2.imread(image_path)
    image = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
    return image
def get_colors(image, number_of_colors, show_chart): 
    modified_image = cv2.resize(image, (600, 400), interpolation = cv2.INTER_AREA)
    modified_image = modified_image.reshape(modified_image.shape[0]*modified_image.shape[1], 3)
    clf = KMeans(n_clusters = number_of_colors)
    labels = clf.fit_predict(modified_image)
    counts = Counter(labels)
    center_colors = clf.cluster_centers_
    ordered_colors = [center_colors[i]/255 for i in counts.keys()]
    hex_colors = [RGB2HEX(ordered_colors[i]*255) for i in counts.keys()]
    rgb_colors = [ordered_colors[i]*255 for i in counts.keys()]
    if (show_chart):
        fig=plt.figure(figsize = (8, 6))
        plt.pie(counts.values(), labels = hex_colors, colors = ordered_colors)
        fig.savefig('/home/pi/RaisinsColorMonitor/plot.png')
    #return rgb_colors
with open('/home/pi/RaisinsColorMonitor/status.txt', 'w') as f:
    f.write('WIP')
print('editied Status file to WIP')
camera = PiCamera()
sleep(2)
camera.capture("/home/pi/RaisinsColorMonitor/Raisins.png")
with open('/home/pi/RaisinsColorMonitor/bands.txt') as f:
    lines = f.readlines()
temp=get_colors(get_image('/home/pi/RaisinsColorMonitor/Raisins.png'),int(lines[0]), True)
with open('/home/pi/RaisinsColorMonitor/status.txt', 'w') as f:
    f.write('DONE')
print('editied Status file to DONE')
