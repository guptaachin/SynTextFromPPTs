# -*- coding: utf-8 -*-

######################### Arguments ##########################################
# the transcription file, which is the output from extract_bounding_box.py
# transcription = r'C:\Users\ismailej\Desktop\slides\slide\manual_part9\unique\transcription_sample.txt'
# the final transcriptions in csv format
# outfile = open(r'C:\Users\ismailej\Desktop\OCR\annotatation.csv', 'w')
# the images extracted by extract_bounding_box.py
# images_folder = r'C:\Users\ismailej\Desktop\OCR\images'
# the annotated images will be stored in the following location
# images_annotated_folder = r'C:\Users\ismailej\Desktop\OCR\images_annotated'
# the below is used for removing the duplicates, so before adding the image
# annotated_folder = os.path.join(os.getcwd(), 'annotated_folder')
# we will check if the image is already contained in the existing annotation list
###############################################################################

import csv # write in csv format
import os
import hashlib
from PIL import Image, ImageDraw


annotated_folder = os.path.join(os.getcwd(), 'annotated_folder')
data_folder = os.path.join(os.getcwd(), 'data')
images_annotated_folder = os.path.join(os.getcwd(), 'images_annotated_folder')
images_folder = os.path.join(os.getcwd(), 'images')
transcription = os.path.join(data_folder, 'transcription.txt')
outfile = open(os.path.join(data_folder, 'annotation.csv'), 'w')


fields = ['file', 'x0','y0', 'width', 'height', 'trans', 'md5hash']
writer = csv.DictWriter(outfile, delimiter=',', lineterminator='\n', fieldnames=fields)
writer.writeheader()   
annotation = {} # contains all the annotation md5hash

if annotated_folder:
  for files in os.listdir(annotated_folder):
      if files.endswith('csv'):
          with open(os.path.join(annotated_folder, files)) as f:
              reader = csv.DictReader(f)
              for row in reader:
                  if 'md5hash' in row:
                    annotation[row['md5hash']] = row['file']

def crop_image(image, rectangle):
    cropped = image.crop([rectangle[0], rectangle[1], rectangle[0] + rectangle[2], rectangle[1] + rectangle[3]])
    bw = cropped.convert('1')
    width, height = bw.size
    pix = bw.load() # getting the pixel values
    horzhist = [0]*height
    
    for i in range(height):
        total = 0
        for j in range(width):
            total += pix[j, i]
        horzhist[i] = total
    
    first = horzhist[0] # Get the first row intensity
    last = horzhist[-1]
    j = 0
    for i in range(1, height):
        if horzhist[i] != first:
            break
    for j in range(height - 2, -1, -1):
        if horzhist[j] != last:
            break
    # Return the new values
            
    new_rect = [0]*4
    new_rect[0] = rectangle[0]
    
    if j > 2:
        new_rect[1] = rectangle[1] + j
    else:
        new_rect[1] = rectangle[1]
        j = 0
        
    if i < height - 2:
        i = height - (i + 1)
        new_rect[3] = rectangle[3] - i - j 
    else:
        new_rect[3] = rectangle[3] - j 
    new_rect[2] = rectangle[2]
    
    return new_rect
    

first = True
first_below = True
ppt = ''
name = image = ''

def pause_n_print(log):
    print(log)
    input('paused')

with open(transcription) as fi:
    for lines in fi:
        # pause_n_print('lines = '+lines)
        lines = lines.strip()
        if 'SlideName' in lines:
            elements = lines.split()
            ppt_name = elements[2]
            #outfile.write(lines + '\n')
            # if first == False:
            #     pause_n_print('ouiouiouiouiouiouiouiouiouiouiouioui')
            #     image.save(os.path.join(images_annotated_folder, name))
            first = False
            first_below = True
        elif 'Slide' in lines:
            elements = lines.split()
            slide_num = elements[1]
            #outfile.write(lines + '\n')
            if first_below == False:
                image.save(os.path.join(images_annotated_folder, name))
            first_below = False
            name = ppt_name + slide_num + '.jpg'
            image_file = os.path.join(images_folder, name)
            image = Image.open(image_file)
            dig = hashlib.md5(image.tobytes()).hexdigest()
            drawable = ImageDraw.Draw(image)
        else:
            elements = lines.split()
            rectangle = elements[:4]
            orig = rectangle[:]
            rectangle = list(map(int, rectangle))
            
            # Now we have to do the processing to make the bounding box
            # tight
            # new_rect = crop_image(image, rectangle)
            # new_rect[2] = new_rect[0] + new_rect[2]
            # new_rect[3] = new_rect[1] + new_rect[3]
            # new_rect[1], new_rect[3] = new_rect[3], new_rect[1]

            # new_rect = crop_image(image, rectangle)
            new_rect = rectangle
            new_rect[2] = new_rect[0] + new_rect[2]
            new_rect[3] = new_rect[1] + new_rect[3]
            new_rect[1], new_rect[3] = new_rect[3], new_rect[1]

            drawable.rectangle(new_rect, outline=(255, 0, 0, 255)) # red box
            trans = ' '.join(elements[4:])
            if dig not in annotation or name == annotation[dig]:
                annotation[dig] = name
                dic = {}
                dic['file'] = name
                dic['x0'] = new_rect[0]
                dic['y0'] = new_rect[1]
                dic['width'] = new_rect[2] - new_rect[0] # width
                dic['height'] = new_rect[3] - new_rect[1] # height
                dic['trans'] = trans
                dic['md5hash'] = dig
                writer.writerow(dic)
                
outfile.close()
           
            
                

