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
    parser.add_argument("language", help="lang_ja,lang_ko,lang_es")
    args = parser.parse_args()
    CURR_LANG = args.language

    data_folder = os.path.join(os.getcwd(), 'data')
    lang_folder, images_folder, image_pool_folder, ppt_folder = i_utilities_ifpeb.init_folder_hierarchy(data_folder, CURR_LANG)

    BATCH_COUNTER = -1
    transcription = None
    folder_for_ppt = ppt_folder
    con_set = i_utilities_ifpeb.populate_links_have(lang_folder)

    # takes care of the condition when the process and stopped. Essential since MS-PPT crashes sometimes.
    is_first_call = True

    for each_ppt in os.listdir(folder_for_ppt):
            
        if  (each_ppt.endswith('ppt') or each_ppt.endswith('pptx')):

            if(each_ppt in con_set):
                print('continuing - ',each_ppt)
                continue  
            # launching the Microsift powerpoint
            Application = win32com.client.Dispatch("PowerPoint.Application")
            Application.Visible = True
            # ------------------------------------
            print('PPTS processed = ', len(con_set))

            # implement the batching to save intermediate results
            if len(con_set) % i_utilities_ifpeb.BATCH == 0 or is_first_call:
                is_first_call = False
                if transcription:
                    transcription.close()
                BATCH_COUNTER  = int(len(con_set) / i_utilities_ifpeb.BATCH)
                filename = os.path.join(lang_folder, 'transcription_'+str(BATCH_COUNTER)+'.txt')
                try:
                    transcription = open(filename, 'a')
                except IOError:
                    transcription = open(filename, 'w')
                print("file to be used = ",filename)
            # ---------------------------------------------------------------
            # create an object for the powerpoint file
            try:
                presentation_object = Application.Presentations.Open(os.path.join(folder_for_ppt, each_ppt))
            except Exception as e:
                # corrupt slide
                print(each_ppt, 'could not open ',e)
                continue

            print("working for = ",each_ppt)
            con_set.add(each_ppt)

            # open up a section in transcription for the current slide
            trans = ["SlideName - " + each_ppt]
            transcription.write(trans[0] + '\n')
            #----------------------------------------------------------
            try :
                for sl_index, each_slide_object in enumerate(presentation_object.Slides):
                    process_this_slide(sl_index, each_slide_object, con_set, trans, transcription, presentation_object, 
                    BATCH_COUNTER, image_pool_folder, each_ppt, images_folder)
                # call this method with multiple threads.

                try:
                    presentation_object.Close()
                except Exception as e:
                    print('problem with closing the file',e)
            except:
                continue

    Application.Quit()
    transcription.close()


def process_this_slide(sl_index, each_slide_object, con_set, trans, transcription, presentation_object, BATCH_COUNTER, image_pool_folder, each_ppt, images_folder):
    print('============================ slide no ================================ ppt - ',len(con_set),'slide = ',str(sl_index+1),'/', len(presentation_object.Slides))
    
    # Divide the groups of all the slides.
    print('BEFORE number of shapes in the current slide = ', len(each_slide_object.Shapes))
    in_group_limit_satisfied = i_utilities_ifpeb.ungroup_all_shapes(each_slide_object)
    print('REVISED number of shapes in the current slide = ', len(each_slide_object.Shapes))
    
    if(not in_group_limit_satisfied):
        print("SKIPPING this slide.")
        return
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
    for i in range(len(each_slide_object.Shapes)):
        try:
            each_shape = each_slide_object.Shapes[i]
            if each_shape.HasTextFrame and each_shape.TextFrame.HasText and not each_shape.HasSmartArt:
                elems = each_shape.TextFrame.TextRange.Lines()
                was_anything_found = i_utilities_ifpeb.save_results_for(elems, trans)
            else: # if has text loop
                to_be_processed_shapes.append(each_shape)
        except :
            continue
            # for other shapes. keep on singing.
    else: # for loop else.
        # Everything good store the slide as image
        name = each_ppt +"_"+ str(sl_index) + "_" + str(BATCH_COUNTER) + '.jpg'
        if was_anything_found:
            try:
                i_utilities_ifpeb.process_these_shapes(to_be_processed_shapes, each_slide_object, image_pool_folder)
            except:
                print('exception during delete shape')
            print('    SAVING ======= ', name)
            try:
                each_slide_object.export(os.path.join(images_folder, name), 'JPG')
                transcription.write('\n'.join(trans) + '\n')
            except:
                print('error during export')

    print('Done looping through the shapes.', time.time() - st_time)



if __name__=='__main__':
    main()