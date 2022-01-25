import cv2
import numpy as np
import math
from scipy import ndimage

def sort_contours(cnts, method="left-to-right"):
    # initialize the reverse flag and sort index
    reverse = False
    i = 0

    # handle if we need to sort in reverse
    if method == "right-to-left" or method == "bottom-to-top":
        reverse = True

    # handle if we are sorting against the y-coordinate rather than
    # the x-coordinate of the bounding box
    if method == "top-to-bottom" or method == "bottom-to-top":
        i = 1

    # construct the list of bounding boxes and sort them from top to
    # bottom
    boundingBoxes = [cv2.boundingRect(c) for c in cnts]
    (cnts, boundingBoxes) = zip(*sorted(zip(cnts, boundingBoxes),key=lambda b: b[1][i], reverse=reverse))

    # return the list of sorted contours and bounding boxes
    return (cnts, boundingBoxes)


#Functon for extracting the box
def box_extraction(img_for_box_extraction_path):
    img = cv2.imread(img_for_box_extraction_path, 0)
    (thresh, img_bin) = cv2.threshold(img, 128, 255,cv2.THRESH_BINARY | cv2.THRESH_OTSU)  # Thresholding the image
    img_bin = 255-img_bin  # Invert the image
    cv2.imwrite("Images/Image_bin.jpg",img_bin)
    kernel_length =(np.array(img).shape[1]//100)#np.array(img).shape[1])
    print(np.array(img).shape[1],kernel_length)
    verticle_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, kernel_length))
    hori_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (kernel_length, 1))
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
    img_temp1 = cv2.erode(img_bin, verticle_kernel, iterations=3)
    verticle_lines_img = cv2.dilate(img_temp1, verticle_kernel, iterations=10)
    kernel2 = np.ones((20,2), np.uint8)  # note this is a horizontal kernel
    d_im = cv2.dilate(verticle_lines_img, kernel2, iterations=2)     
    verticle_lines_img = cv2.erode(d_im, kernel2, iterations=2)
    cv2.imwrite("Images/verticle_lines.jpg",verticle_lines_img)
    img_temp2 = cv2.erode(img_bin, hori_kernel, iterations=3)
    horizontal_lines_img = cv2.dilate(img_temp2, hori_kernel, iterations=10)
    kernel3 = np.ones((2,20), np.uint8)  # note this is a horizontal kernel
    d_im = cv2.dilate(horizontal_lines_img, kernel3, iterations=2)     
    horizontal_lines_img = cv2.erode(d_im, kernel3, iterations=2)
    cv2.imwrite("Images/horizontal_lines.jpg",horizontal_lines_img)
    alpha = 0.5
    beta = 1.0 - alpha
    img_final_bin = cv2.addWeighted(verticle_lines_img, alpha, horizontal_lines_img, beta, 0.0)
    img_final_bin = cv2.erode(~img_final_bin, kernel, iterations = 1)
    (thresh, img_final_bin) = cv2.threshold(img_final_bin, 128, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
    cv2.imwrite("Images/img_final_bin.jpg",img_final_bin)
    contours, hierarchy = cv2.findContours(img_final_bin, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)[-2:]
    (contours, boundingBoxes) = sort_contours(contours, method="top-to-bottom")
    idx = 0
    for c in contours:
        x, y, w, h = cv2.boundingRect(c)
        if (w > 80 and h > 20) and w > 3*h:
            idx += 1
            new_img = img[y:y+h, x:x+w]
            cv2.imwrite("./Output/"+str(idx) + '.png', new_img)
    cv2.drawContours(img, contours, -1, (0, 0, 255), 1)
    cv2.imwrite("./Temp/img_contour.jpg", img)
#box_extraction(r'/home/pi/20210526/New/1.jpeg')
