# -*- coding: utf-8 -*-
# Credits - Ismail.
# modified by achin a little bit

import csv # write in csv format
import os
import hashlib
from PIL import Image, ImageDraw
import argparse

data_folder = os.path.join(os.getcwd(), 'data')

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

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("language", help="lang_ja,lang_ko,lang_es")
    parser.add_argument("batch_no", help="enter batch number")
    args = parser.parse_args()
    CURR_LANG = args.language
    counter = str(args.batch_no)

    lang_folder = os.path.join(data_folder, CURR_LANG)

    images_annotated_folder = os.path.join(lang_folder, 'images_annotated_folder')
    images_folder = os.path.join(lang_folder, "images")

    create_directory(images_annotated_folder)
    
    # for each_file in os.listdir(lang_folder):
    #     if 'transcription_' in each_file:
    process_transcription_file(lang_folder, images_annotated_folder, images_folder, counter)

def process_transcription_file(lang_folder, images_annotated_folder, images_folder, counter):
    transcription = os.path.join(lang_folder, 'transcription_'+counter+'.txt')

    outfile = open(os.path.join(lang_folder, 'annotation_'+counter+'.csv'), 'w')

    fields = ['file', 'x0', 'y0', 'width', 'height', 'trans', 'md5hash']
    writer = csv.DictWriter(outfile, delimiter=',', lineterminator='\n', fieldnames=fields)
    writer.writeheader()
    annotation = {}  # contains all the annotation md5hash


    if images_annotated_folder:
        for files in os.listdir(images_annotated_folder):
            if files.endswith('csv'):
                with open(os.path.join(images_annotated_folder, files)) as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        if 'md5hash' in row:
                            annotation[row['md5hash']] = row['file']


    first = True
    first_below = True
    ppt = ''
    name = image = ''

    with open(transcription) as fi:
        for line in fi:
            # pause_n_print('line = '+line)
            line = line.strip()
            if 'SlideName' in line:
                elements = line.split()
                ppt_name = elements[2]
                # outfile.write(line + '\n')
                if first == False:
                    print('saving = ', name)
                    image.save(os.path.join(images_annotated_folder, name))
                first = False
                first_below = True
            elif 'Slide' in line:
                elements = line.split()
                slide_num = elements[1]
                # outfile.write(line + '\n')
                if first_below == False:
                    print('saving = ', name)
                    image.save(os.path.join(images_annotated_folder, name))
                first_below = False
                name = ppt_name + "_"+slide_num+"_"+str(counter) + '.jpg'
                image_file = os.path.join(images_folder, name)
                image = Image.open(image_file)
                dig = hashlib.md5(image.tobytes()).hexdigest()
                drawable = ImageDraw.Draw(image)
            else:
                elements = line.split()
                rectangle = elements[:4]
                rectangle = list(map(int, rectangle))

                # Now we have to do the processing to make the bounding box
                # tight
                new_rect = crop_image(image, rectangle)
                new_rect[2] = new_rect[0] + new_rect[2]
                new_rect[3] = new_rect[1] + new_rect[3]
                new_rect[1], new_rect[3] = new_rect[3], new_rect[1]
                drawable.rectangle(new_rect, outline=(255, 0, 0, 255))
                trans = ' '.join(elements[4:])
                if dig not in annotation or name == annotation[dig]:
                    annotation[dig] = name
                    dic = {}
                    dic['file'] = name
                    dic['x0'] = new_rect[0]
                    dic['y0'] = new_rect[1]
                    dic['width'] = new_rect[2] - new_rect[0]  # width
                    dic['height'] = new_rect[3] - new_rect[1]  # height
                    dic['trans'] = trans
                    dic['md5hash'] = dig
                    writer.writerow(dic)
        print('saving last image = ' , name)
        image.save(os.path.join(images_annotated_folder, name))

    outfile.close()

def create_directory(directory):
    if not os.path.exists(directory):
        os.makedirs(directory)

if __name__=='__main__':
    main()
