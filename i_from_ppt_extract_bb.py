# -*- coding: utf-8 -*-
################### Arguments ################################################
# https://support.microsoft.com/en-us/help/827745/how-to-change-the-export-resolution-of-a-powerpoint-slide
# 72 as Decimal
###############################################################################

import win32com.client, sys
import os
import argparse
import shutil
import i_draw_bb
from PIL import Image
import numpy as np
import i_draw_bb

BATCH = 250
images_folder = ""
data_folder = os.path.join(os.getcwd(), 'data')

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


# def create_the_h5_py(path, lang, count):
#     OUT_FILE = path + '/SynthText_'+ lang+'_'+str(count) + '.h5'
#     out_db = h5py.File(OUT_FILE,'w')
#     out_db.create_group('/data')

#     return out_db


CURR_LANG = ""
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("language", help="lang_ja,lang_ko,lang_es")
    args = parser.parse_args()

    global CURR_LANG
    CURR_LANG = args.language

    lang_folder = os.path.join(data_folder, CURR_LANG)

    con_slides_file_path = os.path.join(lang_folder, 'considered.txt')
    con_set = populate_links_have(con_slides_file_path)
    con_slide_file = open(con_slides_file_path, 'a')

    images_folder = os.path.join(lang_folder, 'images') # data/lang_ja/images
    
    ppt_folder = os.path.join(lang_folder, 'ppts')
    
    create_directory(images_folder)

    Application = win32com.client.Dispatch("PowerPoint.Application")
    Application.Visible = True

    BATCH_COUNTER = -1
    # open the transcription in the writable mode
    
    transcription = None

    folder_for_ppt = ppt_folder
    ppt_count = 0
    # current_out_db = None

    for each_ppt in os.listdir(folder_for_ppt):
        # each_ppt not in con_set and
            
        if  ( each_ppt.endswith('ppt') or each_ppt.endswith('pptx')):

            print('PPTS processed = ', ppt_count)
            if ppt_count % BATCH == 0:
                if transcription:
                    transcription.close()
                    i_draw_bb.main(BATCH_COUNTER, CURR_LANG)
                    input('done with first batch')
                BATCH_COUNTER += 1
                try:
                    transcription = open(os.path.join(lang_folder, 'transcription_'+str(BATCH_COUNTER)+'.txt'), 'a')
                except IOError:
                    transcription = open(os.path.join(lang_folder, 'transcription_'+str(BATCH_COUNTER)+'.txt'), 'w')
                
            ppt_count += 1

            if(each_ppt in con_set):
                continue
                

            con_slide_file.write(each_ppt+ '\n')
            print("working for = ",each_ppt)
            # create an object for the powerpoint file
            try:
                presentation_object = Application.Presentations.Open(os.path.join(folder_for_ppt, each_ppt))
            except Exception as e:
                # corrupt slide
                print(each_ppt, 'could not open ',e)
                continue

            trans = ["SlideName - " + each_ppt]
            transcription.write(trans[0] + '\n')

            count = 1
            for each_slide_object in presentation_object.Slides:

                # coord_matrix = [[[],[],[],[]],[[],[],[],[]]]

                # lines_list = []
                trans = []

                print('============================ slide no =========================== '+str(count))
                trans.append("Slide " + str(count))
                count += 1
                was_anything_found = False
                to_be_deleted_shape = []

                print('BEFORE number of shapes in the current slide = ', len(each_slide_object.Shapes))
                ungroup_all_shapes(each_slide_object.Shapes)
                print('REVISED number of shapes in the current slide = ', len(each_slide_object.Shapes))


                for i in range(len(each_slide_object.Shapes)):
                    each_shape = each_slide_object.Shapes[i]
                    if each_shape.HasTextFrame and each_shape.TextFrame.HasText and not each_shape.HasSmartArt:
                        elems = each_shape.TextFrame.TextRange.Lines()
                        for elem in elems:
                            # elem.Text is the complete string
                            if elem.Text not in ("\r", "\n", " ", u"\u000D", u"\u000A"): # , u"\u000B", u"\u0009"
                                was_anything_found = True
                                print(elem.Text)
                                result = charwise_hex_string(elem.Text)
                                top_left = 0
                                top_right = 1
                                bottom_right = 2
                                bottom_left = 3
                                x = 0
                                y = 1

                                # lines_list.append(result)

                                # print(lines_list)

                                # coord_matrix[x][top_left].append(int(elem.BoundLeft))
                                # coord_matrix[y][top_left].append(int(elem.BoundTop))

                                # coord_matrix[x][top_right].append(int(elem.BoundLeft) + int(elem.BoundWidth))
                                # coord_matrix[y][top_right].append(int(elem.BoundTop))

                                # coord_matrix[x][bottom_right].append(int(elem.BoundLeft)+ int(elem.BoundWidth))
                                # coord_matrix[y][bottom_right].append(int(elem.BoundTop) + int(elem.BoundHeight))

                                # coord_matrix[x][bottom_left].append(int(elem.BoundLeft))
                                # coord_matrix[y][bottom_left].append(int(elem.BoundTop) + int(elem.BoundHeight))

                                # point_top_left = (int(elem.BoundLeft), int(elem.BoundTop))
                                # point_top_right = (int(elem.BoundLeft) + int(elem.BoundWidth) , int(elem.BoundTop))
                                # point_bot_left = (int(elem.BoundLeft) , int(elem.BoundTop) + int(elem.BoundHeight))
                                # point_bot_right = (int(elem.BoundLeft)+ int(elem.BoundWidth) , int(elem.BoundTop) + int(elem.BoundHeight))
                                # box = [point_top_left, point_top_right, point_bot_left, point_bot_right]
                                # list_boxes.append(box)
            
                                trans.append(str(int(elem.BoundLeft)) + ' ' + str(int(elem.BoundTop)) + ' ' + str(
                                        int(elem.BoundWidth)) + ' ' + str(int(elem.BoundHeight)) + ' ' + result)
                    else:
                        # print('adding for delete a single image')
                        to_be_deleted_shape.append(each_shape)
                else:
                    # Everything good store the slide as image
                    name = each_ppt +"_"+ str(count - 1) + "_" + str(BATCH_COUNTER) + '.jpg'
                    
                    if was_anything_found:

                        try:
                            [temp.Delete() for temp in to_be_deleted_shape]
                        except:
                            print('exception during delete shape')

                        print('saving ======= ', name)
                        each_slide_object.export(os.path.join(images_folder, name), 'JPG')
                        transcription.write('\n'.join(trans) + '\n')
            try:
                presentation_object.Close()
            except Exception as e:
                print(e)

    Application.Quit()
    transcription.close()
    con_slide_file.close()

def create_directory(directory):
    if not os.path.exists(directory):
        os.makedirs(directory)


def ungroup_all_shapes(shapes):
    for each_shape in shapes:
        try:
            each_shape.Ungroup()
        except:
            # print('Sorry could not do that ungrouping')
            pass

def populate_links_have(links_path):
    s = set()
    try:
        lnk_file = open(links_path, 'r')
    except IOError:
        lnk_file = open(links_path, 'w')
        return s

    present_lines = lnk_file.readlines()
    for each_line in present_lines:
        curr_entry = each_line.rstrip()
        s.add(curr_entry)
    lnk_file.close()
    return s


if __name__=='__main__':
    main()
    
