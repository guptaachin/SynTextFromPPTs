# -*- coding: utf-8 -*-
################### Arguments ################################################
# https://support.microsoft.com/en-us/help/827745/how-to-change-the-export-resolution-of-a-powerpoint-slide
# 72 as Decimal
###############################################################################

import win32com.client, sys
import os
import argparse
import i_utilities_ifpeb


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("language", help="lang_ja,lang_ko,lang_es, other")
    args = parser.parse_args()
    CURR_LANG = args.language

    data_folder = os.path.join(os.getcwd(), 'data')
    lang_folder, images_folder, image_pool_folder, ppt_folder = i_utilities_ifpeb.init_folder_hierarchy(data_folder,
                                                                                                        CURR_LANG)

    BATCH_COUNTER = -1
    transcription_cl, transcription_wl, transcription_ll = None, None, None
    folder_for_ppt = ppt_folder
    con_set = i_utilities_ifpeb.populate_links_have(lang_folder)

    # takes care of the condition when the process and stopped. Essential since MS-PPT crashes sometimes.
    is_first_call = True

    # this is the languages call
    lst_files = []
    paths = ["D:/lang_ja", "D:/lang_ko", "D:/lang_es"]
    for i in paths:
        for f in os.listdir(i):
            lst_files.append(i+"/"+f)

    for each_ppt in lst_files:

        if (each_ppt.endswith('ppt') or each_ppt.endswith('pptx')):

            if (each_ppt in con_set):
                print('continuing - ', each_ppt)
                continue
                # launching the Microsift powerpoint
            Application = win32com.client.Dispatch("PowerPoint.Application")
            Application.Visible = True
            # ------------------------------------
            print('PPTS processed = ', len(con_set))

            # implement the batching to save intermediate results
            if len(con_set) % i_utilities_ifpeb.BATCH == 0 or is_first_call:
                is_first_call = False
                if transcription_cl:
                    transcription_cl.close()
                if transcription_wl:
                    transcription_wl.close()
                if transcription_ll:
                    transcription_ll.close()
                BATCH_COUNTER = int(len(con_set) / i_utilities_ifpeb.BATCH)
                filename_cl = os.path.join(lang_folder, 'transcription_cl_' + str(BATCH_COUNTER) + '.txt')
                filename_ll = os.path.join(lang_folder, 'transcription_ll_' + str(BATCH_COUNTER) + '.txt')
                filename_wl = os.path.join(lang_folder, 'transcription_wl_' + str(BATCH_COUNTER) + '.txt')
                try:
                    transcription_cl = open(filename_cl, 'a')
                    transcription_wl = open(filename_wl, 'a')
                    transcription_ll = open(filename_ll, 'a')
                except IOError:
                    transcription_cl = open(filename_cl, 'w')
                    transcription_wl = open(filename_wl, 'w')
                    transcription_ll = open(filename_ll, 'w')

                # print("file to be used = ",filename)
            # ---------------------------------------------------------------
            # create an object for the powerpoint file
            try:
                presentation_object = Application.Presentations.Open(each_ppt)
            except Exception as e:
                # corrupt slide
                print(each_ppt, 'could not open ', e)
                continue

            print("working for = ", each_ppt)
            con_set.add(each_ppt)

            # open up a section in transcription for the current slide
            trans_cl = ["SlideName - " + each_ppt]
            trans_wl = ["SlideName - " + each_ppt]
            trans_ll = ["SlideName - " + each_ppt]
            transcription_cl.write(trans_cl[0] + '\n')
            transcription_wl.write(trans_wl[0] + '\n')
            transcription_ll.write(trans_ll[0] + '\n')
            # ----------------------------------------------------------
            try:
                for sl_index, each_slide_object in enumerate(presentation_object.Slides):
                    process_this_slide(0, sl_index, each_slide_object, con_set, trans_cl, transcription_cl,
                                       presentation_object,
                                       BATCH_COUNTER, image_pool_folder, each_ppt, images_folder)
                    break
                    # process_this_slide(1, sl_index, each_slide_object, con_set, trans_wl, transcription_wl,
                    #                    presentation_object,
                    #                    BATCH_COUNTER, image_pool_folder, each_ppt, images_folder)
                    # process_this_slide(2, sl_index, each_slide_object, con_set, trans_ll, transcription_ll,
                    #                    presentation_object,
                    #                    BATCH_COUNTER, image_pool_folder, each_ppt, images_folder)
                # call this method with multiple threads.

                try:
                    presentation_object.Close()
                except Exception as e:
                    print('problem with closing the file', e)
            except:
                continue

    Application.Quit()
    transcription_cl.close()
    transcription_wl.close()
    transcription_ll.close()


def process_this_slide(level, sl_index, each_slide_object, con_set, trans, transcription, presentation_object,
                       BATCH_COUNTER, image_pool_folder, each_ppt, images_folder):
    print('============================ slide no ================================ ppt - ', len(con_set), 'slide = ',
          str(sl_index + 1), '/', len(presentation_object.Slides))

    # Divide the groups of all the slides.
    print('BEFORE number of shapes in the current slide = ', len(each_slide_object.Shapes))
    in_group_limit_satisfied = i_utilities_ifpeb.ungroup_all_shapes(each_slide_object)
    print('REVISED number of shapes in the current slide = ', len(each_slide_object.Shapes))

    # if (not in_group_limit_satisfied):
    #     print("SKIPPING this slide.")
    #     return
    # -----------------------------------------

    # initilizaions for the slide processing.
    trans = []
    trans.append("Slide " + str(sl_index))
    was_anything_found = False
    to_be_processed_shapes = []

    import time
    st_time = time.time()
    print('Starting to loop through ungrouped shapes on the shapes')
    # finally process the slide. Extract the text that is in the slides.
    was_anything_found = True

    for i in range(len(each_slide_object.Shapes)):
        try:
            each_shape = each_slide_object.Shapes[i]
            # if each_shape.HasTextFrame and each_shape.TextFrame.HasText and not each_shape.HasSmartArt:
            to_be_processed_shapes.append(each_shape)

        except :
            continue
            # for other shapes. keep on singing.
    else:  # for loop else.
        name = each_ppt.split("/")[2] + "_" + str(sl_index) + "_" + str(BATCH_COUNTER) + '.jpg'
        if was_anything_found:
            try:
                deleteThese(to_be_processed_shapes)
                # print(os.path.join(images_folder, "test"))
                each_slide_object.export(os.path.join(images_folder, name), 'JPG')
                transcription.write('\n'.join(trans) + '\n')
                print('    SAVING ======= ', name)
            except Exception as e:
                print('error during export = ',e)

    print('Done looping through the shapes.', time.time() - st_time)


def deleteThese(shapes):
    for each_shape in shapes:
        try:
            each_shape.Delete()
        except:
            print('exception during delete shape')

if __name__ == '__main__':
    main()