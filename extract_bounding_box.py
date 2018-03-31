# -*- coding: utf-8 -*-


################### Arguments ################################################
# folder points to the directory where slides are present, 1 files will be
# created in this folder 
# 1) transcription.txt => which contains bounding box information + transcription
# slides_folder = r'C:\Users\ismailej\Desktop\slides\slide\manual_part9\unique'
#images_folder will contain the location of stored ppt slide images

###############################################################################

import win32com.client, sys
import os

# slides folder contains all the slides
data_folder = os.path.join(os.getcwd(), 'data')

# images folder is the place where all the images will be written down
images_folder = os.path.join(os.getcwd(), 'images')

def charwise_hex_string(item):
    final = ''
    first_time = True
    for elem in range(len(item)):
        dec_value = ord(item[elem])
        hex_value = hex(dec_value)
        hex_value = hex_value[2:]  #0xff
        if len(hex_value) < 4:
            hex_value = '0'*(4 - len(hex_value)) + hex_value
        hex_value = 'u' + hex_value
        if first_time:
            final = hex_value
            first_time = False
        else:
            final = final + '_' + hex_value
    split_final = final.split('_u0020_')
    split_final = ' u0020 '.join(split_final)

    return split_final
    
Application = win32com.client.Dispatch("PowerPoint.Application")
Application.Visible = True

# open the transcription in the writable mode
transcription = open(os.path.join(data_folder, 'transcription.txt'), 'w')

for each_ppt in os.listdir(data_folder):
    if each_ppt.endswith('ppt') or each_ppt.endswith('pptx'):

        # create an object for the powerpoint file
        try:
            presentation_object = Application.Presentations.Open(os.path.join(data_folder, each_ppt))
        except:
            # corrupt slide
            print('could not open ',each_ppt)
            input('waiting')
            # os.remove(os.path.join(slides_folder, each_ppt))
            continue

        # provide the heading for each slide
        trans = ["SlideName - " + each_ppt]
        transcription.write(trans[0] + '\n')
        # print("Working on " + trans[0])

        # keeps a count for the images
        count = 1

        # for each opened slide in the ppt
        print('working for = ',each_ppt)




        for each_slide_object in presentation_object.Slides:
            trans = []
            #bound = []
            #bound.append("Slide " + str(count))
            trans.append("Slide " + str(count))
            count += 1

            for each_shape in each_slide_object.Shapes:

                print('each_shape.TextFrame.HasText = ',each_shape.TextFrame.HasText)

                if each_shape.TextFrame.HasText:
                    try:
                        elems = each_shape.TextFrame.TextRange.Lines()
                        for elem in elems:
                            # elem.Text is the complete string
                            if elem.Text not in ("\r", "\n", " ", u"\u000D", u"\u000A"):
                                result = charwise_hex_string(elem.Text)
                                trans.append(str(int(elem.BoundLeft)) + ' ' +  str(int(elem.BoundTop)) + ' ' + str(int(elem.BoundWidth)) + ' ' + str(int(elem.BoundHeight)) + ' ' + result)
                    except:
                        try:
                            smart = each_shape.GroupItems
                            for i in range(smart.Count):
                                elem = smart[i].TextFrame.TextRange.Lines()
                                for s in elem:
                                    result = charwise_hex_string(elem.Text)
                                    trans.append(str(int(elem.BoundLeft)) + ' ' + str(int(elem.BoundTop)) + ' ' + str(int(elem.BoundWidth)) + ' ' + str(int(elem.BoundHeight)) + ' ' + result)
                        except:
                            break
                else:
                    try:
                        smart = each_shape.GroupItems

                        print('each_shape.GroupItems = ', len(each_shape.GroupItems))
                        input()

                        for i in range(smart.Count):
                            elem = smart[i].TextFrame.TextRange.Lines()
                            for s in elem:
                                result = charwise_hex_string(elem.Text)
                                trans.append(str(int(elem.BoundLeft)) + ' ' + str(int(elem.BoundTop)) + ' ' + str(int(elem.BoundWidth)) + ' ' + str(int(elem.BoundHeight)) + ' ' + result)
                    except Exception as e:
                        print('Exception--------------', e)
                        break
            else:
                # Everything good store the slide as image
                name = each_ppt + str(count - 1) + '.jpg'
                print('name ======= ',name)
                each_slide_object.export(os.path.join(images_folder, name), 'JPG')
                transcription.write('\n'.join(trans) + '\n')
                trans = []

        # work for just one slide

        # break
        presentation_object.Close()
        
Application.Quit()
transcription.close()
