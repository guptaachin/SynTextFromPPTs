# -*- coding: utf-8 -*-
################### Arguments ################################################
# https://support.microsoft.com/en-us/help/827745/how-to-change-the-export-resolution-of-a-powerpoint-slide
# 72
###############################################################################

import win32com.client, sys
import os
import argparse


# slides folder contains all the slides
# images folder is the place where all the images will be written down
images_folder = os.path.join(os.getcwd(), 'images')
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


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("language", help="lang_ja,lang_ko,lang_es")
    args = parser.parse_args()
    CURR_LANG = args.language

    data_folder = os.path.join(os.path.join(os.getcwd(), 'data'), CURR_LANG)

    con_slides_file_path = os.path.join(data_folder, 'considered.txt')
    con_set = populate_links_have(con_slides_file_path)
    con_slide_file = open(con_slides_file_path, 'a')

    images_folder = os.path.join(os.path.join(os.path.join(os.getcwd(), 'data'), CURR_LANG), 'images')

    create_directory(images_folder)

    Application = win32com.client.Dispatch("PowerPoint.Application")
    Application.Visible = True

    # open the transcription in the writable mode
    try:
        transcription = open(os.path.join(data_folder, 'transcription.txt'), 'a')
    except IOError:
        transcription = open(os.path.join(data_folder, 'transcription.txt'), 'w')

    for each_ppt in os.listdir(data_folder):
        if each_ppt not in con_set and (each_ppt.endswith('ppt') or each_ppt.endswith('pptx')):
            con_slide_file.write(each_ppt+ '\n')
            # create an object for the powerpoint file
            try:
                presentation_object = Application.Presentations.Open(os.path.join(data_folder, each_ppt))
            except:
                # corrupt slide
                # print('could not open ',each_ppt)
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

            for each_slide_object in presentation_object.Slides:
                trans = []
                # bound = []
                # bound.append("Slide " + str(count))

                print('count ------------------------------------------------------- ', count)
                print('length of shapes - ', len(each_slide_object.Shapes))
                # input()


                trans.append("Slide " + str(count))
                count += 1
                was_anything_found = False
                for each_shape in each_slide_object.Shapes:

                    print('each_shape.TextFrame.HasText = ', each_shape.TextFrame.HasText)

                    if each_shape.TextFrame.HasText:
                        try:
                            elems = each_shape.TextFrame.TextRange.Lines()
                            for elem in elems:

                                # elem.Text is the complete string
                                if elem.Text not in ("\r", "\n", " ", u"\u000D", u"\u000A"):
                                    was_anything_found = True
                                    result = charwise_hex_string(elem.Text)
                                    trans.append(str(int(elem.BoundLeft)) + ' ' + str(int(elem.BoundTop)) + ' ' + str(
                                        int(elem.BoundWidth)) + ' ' + str(int(elem.BoundHeight)) + ' ' + result)
                        except:
                            try:
                                smart = each_shape.GroupItems
                                for i in range(smart.Count):
                                    elem = smart[i].TextFrame.TextRange.Lines()
                                    for s in elem:
                                        was_anything_found = True
                                        result = charwise_hex_string(elem.Text)
                                        trans.append(
                                            str(int(elem.BoundLeft)) + ' ' + str(int(elem.BoundTop)) + ' ' + str(
                                                int(elem.BoundWidth)) + ' ' + str(int(elem.BoundHeight)) + ' ' + result)
                            except:
                                print('')
                    else:
                        try:
                            smart = each_shape.GroupItems
                            for i in range(smart.Count):
                                elem = smart[i].TextFrame.TextRange.Lines()
                                for s in elem:
                                    was_anything_found = True
                                    result = charwise_hex_string(elem.Text)
                                    trans.append(str(int(elem.BoundLeft)) + ' ' + str(int(elem.BoundTop)) + ' ' + str(
                                        int(elem.BoundWidth)) + ' ' + str(int(elem.BoundHeight)) + ' ' + result)
                        except Exception as e:
                            print('Exception--------------', e)
                else:
                    # Everything good store the slide as image
                    name = each_ppt + str(count - 1) + '.jpg'
                    print('saving ======= ', name)
                    if was_anything_found:
                        each_slide_object.export(os.path.join(images_folder, name), 'JPG')
                        transcription.write('\n'.join(trans) + '\n')
                    # trans = []

            presentation_object.Close()

    Application.Quit()
    transcription.close()
    con_slide_file.close()

def create_directory(directory):
    if not os.path.exists(directory):
        os.makedirs(directory)


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
