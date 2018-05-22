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

# slides folder contains all the slides
# images folder is the place where all the images will be written down
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
    
    # remove only for texting
    # [os.remove(images_folder+'/'+r) for r in os.listdir(images_folder)]   
    
    ppt_folder = os.path.join(lang_folder, 'ppts')
    
    create_directory(images_folder)

    Application = win32com.client.Dispatch("PowerPoint.Application")
    Application.Visible = True

    # open the transcription in the writable mode
    try:
        transcription = open(os.path.join(lang_folder, 'transcription.txt'), 'a')
    except IOError:
        transcription = open(os.path.join(lang_folder, 'transcription.txt'), 'w')

    folder_for_ppt = lang_folder


    
    for each_ppt in os.listdir(folder_for_ppt):
        # each_ppt not in con_set and
        if  (each_ppt.endswith('ppt') or each_ppt.endswith('pptx')):

            con_slide_file.write(each_ppt+ '\n')
            print("working for = ",each_ppt)
            # create an object for the powerpoint file
            try:
                presentation_object = Application.Presentations.Open(os.path.join(folder_for_ppt, each_ppt))
            except Exception as e:
                # corrupt slide
                print(each_ppt, 'could not open ',e)
                # input('waiting')
                # os.remove(os.path.join(slides_folder, each_ppt))
                continue

            # provide the heading for each slide
            trans = ["SlideName - " + each_ppt]
            transcription.write(trans[0] + '\n')
            # print("Working on " + trans[0])

            # keeps a count for the images
            count = 1

            # for each opened slide in the ppt
            # print('working for = ', each_ppt)

            for i_slide, each_slide_object in enumerate(presentation_object.Slides):
                trans = []
                # bound = []
                # bound.append("Slide " + str(count))

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
                    if each_shape.HasTextFrame and each_shape.TextFrame.HasText:
                        elems = each_shape.TextFrame.TextRange.Lines()
                        for elem in elems:
                            # elem.Text is the complete string
                            if elem.Text not in ("\r", "\n", " ", u"\u000D", u"\u000A"): # , u"\u000B", u"\u0009"
                                was_anything_found = True
                                print(elem.Text)
                                result = charwise_hex_string(elem.Text)
                                trans.append(str(int(elem.BoundLeft)) + ' ' + str(int(elem.BoundTop)) + ' ' + str(
                                        int(elem.BoundWidth)) + ' ' + str(int(elem.BoundHeight)) + ' ' + result)
                    else:
                        # print('adding for delete a single image')
                        to_be_deleted_shape.append(each_shape)
                    # input('xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxjust processed an independent')
                # find out the place where that image is getting the call.
                    # if each_shape.TextFrame.HasText:
                    #     print(i_shape,' = ', 'HAS TEXT CALL')
                    #     try:
                    #         elems = each_shape.TextFrame.TextRange.Lines()
                    #         # wasTextAdded = False

                    #         print('length of elements = ',len(elems))
                    #         for elem in elems:
                    #             # elem.Text is the complete string
                    #             if elem.Text not in ("\r", "\n", " ", u"\u000D", u"\u000A"):
                    #                 print(elem.Text)
                    #                 was_anything_found = True
                    #                 # wasTextAdded = True
                    #                 result = charwise_hex_string(elem.Text)
                    #                 trans.append(str(int(elem.BoundLeft)) + ' ' + str(int(elem.BoundTop)) + ' ' + str(
                    #                     int(elem.BoundWidth)) + ' ' + str(int(elem.BoundHeight)) + ' ' + result)
                    #             # if not wasTextAdded:
                    #             #     each_slide_object.Shapes[i_shape].Delete()
                    #     except:
                    #         print(i_shape,' = ', 'exception')
                    #         try:
                    #             smart = each_shape.GroupItems

                    #             for i in range(smart.Count):
                    #                 elem = smart[i].TextFrame.TextRange.Lines()
                    #                 for s in elem:
                    #                     was_anything_found = True
                    #                     result = charwise_hex_string(elem.Text)
                    #                     trans.append(
                    #                         str(int(elem.BoundLeft)) + ' ' + str(int(elem.BoundTop)) + ' ' + str(
                    #                             int(elem.BoundWidth)) + ' ' + str(int(elem.BoundHeight)) + ' ' + result)
                    #         except:
                    #             print(i_shape,'has text excep excep')
                    #             # try:
                    #             #     each_shape.Delete()
                    #             # except:
                    #             #     print()
                    # else:
                    #     print(i_shape,' = ', 'NOT HAVE TEXT CALL')
                    #     try:
                    #         each_shape.Delete()
                    #     except:
                    #         print()

                        # add some other image here please.

                        # try:
                        #     smart = each_shape.GroupItems

                        #     # if smart.Count == 0:
                        #     #         each_slide_object.Shapes[i_shape].Delete()
                                    
                        #     for i in range(smart.Count):
                        #         elem = smart[i].TextFrame.TextRange.Lines()
                        #         for s in elem:
                        #             was_anything_found = True
                        #             result = charwise_hex_string(elem.Text)
                        #             trans.append(str(int(elem.BoundLeft)) + ' ' + str(int(elem.BoundTop)) + ' ' + str(
                        #                 int(elem.BoundWidth)) + ' ' + str(int(elem.BoundHeight)) + ' ' + result)
                        # except Exception as e:
                        #     print(i_shape,' = no text excep')
                            # try:
                            #     each_shape.Delete()
                            # except:
                            #     print()
                
                # else for the shapes for loop.
                else:
                    # Everything good store the slide as image
                    name = each_ppt + str(count - 1) + '.jpg'
                    
                    if was_anything_found:

                        try:
                            [temp.Delete() for temp in to_be_deleted_shape]
                        except:
                            print('exception during delete shape')

                        print('saving ======= ', name)
                        each_slide_object.export(os.path.join(images_folder, name), 'JPG')
                        transcription.write('\n'.join(trans) + '\n')

            presentation_object.Close()

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
    i_draw_bb.main(CURR_LANG)
    
