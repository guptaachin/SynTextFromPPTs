# -*- coding: utf-8 -*-


################### Arguments ################################################
# folder points to the directory where slides are present, 1 files will be
# created in this folder 
# 1) transcription.txt => which contains bounding box information + transcription
folder = r'C:\Users\ismailej\Desktop\slides\slide\manual_part9\unique'
#images_folder will contain the location of stored ppt slide images
images_folder = r'C:\Users\ismailej\Desktop\OCR\images'
###############################################################################

import win32com.client, sys
import os


def decode_string(item):
    final = ''
    first_time = True
    for elem in range(len(item)):
        dec_value = ord(item[elem])
        hex_value = hex(dec_value)
        hex_value = hex_value[2:]
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


transcription = open(os.path.join(folder, 'transcription.txt'), 'w')

for files in os.listdir(folder):
    if files.endswith('ppt') or files.endswith('pptx'):
        try: Presentation = Application.Presentations.Open(os.path.join(folder, files))
        except:
            os.remove(os.path.join(folder, files))
            continue
        trans = ["SlideName - " + files]
        transcription.write(trans[0] + '\n')
        print("Working on " + trans[0])
        count = 1
        for Slide in Presentation.Slides:
            trans = []
            #bound = []
            #bound.append("Slide " + str(count))
            trans.append("Slide " + str(count))
            count += 1
            for Shape in Slide.Shapes:
                if Shape.TextFrame.HasText:
                    try: 
                        elems = Shape.TextFrame.TextRange.Lines()
                        for elem in elems: 
                            if elem.Text not in ("\r", "\n", " ", u"\u000D", u"\u000A"):
                                result = decode_string(elem.Text)
                                trans.append(str(int(elem.BoundLeft)) + ' ' +  str(int(elem.BoundTop)) + ' ' + str(int(elem.BoundWidth)) + ' ' + str(int(elem.BoundHeight)) + ' ' + result)
                    except:
                        try:
                            smart = Shape.GroupItems
                            for i in range(smart.Count):
                                elem = smart[i].TextFrame.TextRange.Lines()
                                for s in elem:
                                    result = decode_string(elem.Text)
                                    trans.append(str(int(elem.BoundLeft)) + ' ' +  str(int(elem.BoundTop)) + ' ' + str(int(elem.BoundWidth)) + ' ' + str(int(elem.BoundHeight)) + ' ' + result)
                        except:
                            break
                else:
                    try:
                        smart = Shape.GroupItems
                        for i in range(smart.Count):
                            elem = smart[i].TextFrame.TextRange.Lines()
                            for s in elem:
                                result = decode_string(elem.Text)
                                trans.append(str(int(elem.BoundLeft)) + ' ' +  str(int(elem.BoundTop)) + ' ' + str(int(elem.BoundWidth)) + ' ' + str(int(elem.BoundHeight)) + ' ' + result)
                    except:
                        break
            else:
                # Everything good store the slide as image
                name = files + str(count - 1) + '.jpg'
                Slide.export(os.path.join(images_folder, name), 'JPG')
                transcription.write('\n'.join(trans) + '\n')
                trans = []

        Presentation.Close()
        
Application.Quit()
transcription.close()
