import cv2 
import os
import numpy as np
import threading

data_folder = os.path.join(os.getcwd(), 'data')
image_pool_folder = os.path.join(data_folder, 'image_pool')
parallel_folder = os.path.join(data_folder, 'image_pool_2')

def main(list_files, thread_no):
    count = 0
    for each_img in list_files:
        image_path_processed = os.path.join(parallel_folder, each_img)
        if(os.path.exists(image_path_processed)):
            continue
        image_path = os.path.join(image_pool_folder, each_img)
        img = cv2.imread(image_path) #load rgb image
        hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV) #convert it to hsv
        h, s, v = cv2.split(hsv)
        value = 50
        lim = 255 - value
        v[v > lim] = 255
        v[v <= lim] += value
        final_hsv = cv2.merge((h, s, v))
        img = cv2.cvtColor(final_hsv, cv2.COLOR_HSV2BGR)
        image_path_processed = os.path.join(parallel_folder, each_img)
        cv2.imwrite(image_path_processed, img)
        count += 1
        print(str(thread_no)," number of files = ",len(os.listdir(parallel_folder)))

def create_directory(directory):
    if not os.path.exists(directory):
        os.makedirs(directory)

if __name__ == "__main__":
    create_directory(parallel_folder)
    list_files = os.listdir(image_pool_folder)
    count = 0

    quarter = int(len(list_files) /10)

    thread_list = []

    for i in range(10):
        thread_list.append(threading.Thread(target=main, args=(list_files[quarter*i:i*quarter+quarter],i, )))
        thread_list[i].start()
