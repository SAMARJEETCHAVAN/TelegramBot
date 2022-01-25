import cv2
from PIL import Image,ImageChops
import numpy as np
import math
import socket
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import os
import leafcenter# import findmedian
leafcenter
gauth = GoogleAuth()
drive = GoogleDrive(gauth)

def findleafarea(img_path):
    #print('1:comes here')
    img = cv2.imread(img_path)
    #print('2:comes here')
    hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
    #print('3:comes here')
    mask1 = cv2.inRange(hsv, (36, 0, 0), (70, 255,255))
    #print('4:comes here')
    mask2 = cv2.inRange(hsv, (15,0,0), (36, 255, 255))
    #print('5:comes here')
    mask = cv2.bitwise_or(mask1, mask2)
    #print('6:comes here')
    target = cv2.bitwise_and(img,img, mask=mask)
    #print('7:comes here')
    cv2.imwrite("onlyleaf.jpeg", target)
    #print('8:comes here')
    image = Image.open("onlyleaf.jpeg")
    #print('9:comes here')
    image_data = image.load()
    #print('10:comes here')
    height1,width1 = image.size
    #print('11:comes here')
    ############################################################
    #print('12:comes here')
    image = Image.open(img_path)
    #print('13:comes here')
    height2,width2 = image.size
    ##print(height1,width1,height2,width2)
    if height1==width2:
        image = image.rotate(90,expand=True)
        height2,width2 = image.size
    image_data = image.load()
    #print('14:comes here')
    

    for loop1 in range(height2):
        for loop2 in range(width2):
            r,g,b = image_data[loop1,loop2]
            if((b>r)and(b>g)):
                image_data[loop1,loop2] = 0,0,0
    image.save('leafandcoin.jpeg')
    #print('15:comes here')
    ############################################################
    img2 = Image.open("onlyleaf.jpeg")
    #print('16:comes here')
    img1 = Image.open("leafandcoin.jpeg")
    #print('17:comes here')
    diff = ImageChops.difference(img1, img2)
    #print('18:comes here')
    diff = diff.save("diffbet2.jpeg")
    #print('19:comes here')
    image = Image.open('diffbet2.jpeg')
    #print('20:comes here')
    image_data = image.load()
    #print('21:comes here')
    height,width = image.size
    #print('22:comes here')
    for loop1 in range(height):
        for loop2 in range(width):
            r,g,b = image_data[loop1,loop2]
            if((r>200)or((r<50)and(g<50)and(b<50))):
                image_data[loop1,loop2] = 0,0,0
    image.save('diffbet3.jpeg')
    #print('23:comes here')
    ############################################################
    input_image = cv2.imread('diffbet3.jpeg', cv2.IMREAD_COLOR)
    #print('24:comes here')
    kernel = np.ones((3,3), np.uint8)       # set kernel as 3x3 matrix from numpy
    #print('25:comes here')
    erosion_image = cv2.erode(input_image, kernel, iterations=2)
    #print('26:comes here')
    dilation_image = cv2.dilate(erosion_image, kernel, iterations=2)
    #print('27:comes here')
    erosion_image = cv2.erode(dilation_image, kernel, iterations=1)
    #print('28:comes here')
    dilation_image = cv2.dilate(erosion_image, kernel, iterations=1)
    #print('29:comes here')
    erosion_image = cv2.erode(dilation_image, kernel, iterations=2)
    #print('30:comes here')
    dilation_image = cv2.dilate(erosion_image, kernel, iterations=5)
    #print('31:comes here')
    erosion_image = cv2.erode(dilation_image, kernel, iterations=10)
    #print('32:comes here')
    dilation_image = cv2.dilate(erosion_image, kernel, iterations=10)
    #print('33:comes here')
    #cv2.waitKey(0)
    #print('34:comes here')
    cv2.imwrite("coin.jpeg", dilation_image)
    #print('35:comes here')
    ############################################################
    def find_area(img):
        image = Image.open(img)
        image_data = image.load()
        height,width = image.size
        zeros = 0
        values = 1
        for loop1 in range(height):
            for loop2 in range(width):
                r,g,b = image_data[loop1,loop2]
                if((r==0)and(g==0)and(b==0)):
                    zeros = zeros + 1
                else:
                    values = values + 1
        return(values)
    #print('36:comes here')
    leafvalue = find_area("onlyleaf.jpeg")
    #print('37:comes here')
    coinvalue = find_area("coin.jpeg")
    #print('38:comes here')
    leafarea = leafvalue*(math.pi)/coinvalue
    #print('39:comes here')
    #print('40:comes here')
    image = cv2.imread(img_path,cv2.IMREAD_UNCHANGED)
    position = (leafcenter.findmedian("onlyleaf.jpeg"))
    #print(position)
    leafarea = "%.3f"%float(leafarea)
    cv2.putText(
     image, #numpy array on which text is written
     ("%.3f sq.cm"%float(leafarea)), #text
     position, #position at which writing has to start
     cv2.FONT_HERSHEY_SIMPLEX, #font family
     1, #font size
     (255, 255, 255, 255), #font color
     3) #font stroke
    cv2.imwrite(img_path, image)
    upload_file_list = [img_path]
    for upload_file in upload_file_list:
        gfile = drive.CreateFile({'parents': [{'id': '1FwRcVk2u6JR-cqmnBGTOm8fD-wpat0QV'}]})
        # Read file and set it as a content of this instance.
        gfile.SetContentFile(upload_file)
        gfile.Upload()
    
    os.remove('leafandcoin.jpeg')
    os.remove('diffbet2.jpeg')
    os.remove("onlyleaf.jpeg")
    os.remove("coin.jpeg")
    os.remove('diffbet3.jpeg')
    return("Area of leaf is %.3f sq.cm"%float(leafarea))
