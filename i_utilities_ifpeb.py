import os
import random

BATCH = 100
THRESH_HOLD_GP = 35

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

def save_results_for(elems, trans):
    was_anything_found = False
    for elem in elems:
        if elem.Text not in ("\r", "\n", " ", u"\u000D", u"\u000A"): # , u"\u000B", u"\u0009"
            # skip if only spaces are there
            if(elem.Text.isspace()):
                continue
            # makes it eligible for the slide to be recorded as a sample
            was_anything_found = True
            text_in_unicode = charwise_hex_string(elem.Text)
            # See the code in the end if you are trying to extract.
            trans.append(str(int(elem.BoundLeft)) + ' ' + str(int(elem.BoundTop)) + ' ' 
            + str(int(elem.BoundWidth)) + ' ' + str(int(elem.BoundHeight)) + ' ' + text_in_unicode)
    return was_anything_found

def init_folder_hierarchy(data_folder, CURR_LANG):
    lang_folder = os.path.join(data_folder, CURR_LANG)
    images_folder = os.path.join(lang_folder, 'images') # data/lang_ja/images
    image_pool_folder = os.path.join(data_folder, 'image_pool_2')
    ppt_folder = os.path.join(lang_folder, 'ppts') # ppts change

    if CURR_LANG == 'lang_ko' or CURR_LANG == 'lang_es':
        ppt_folder = 'D:/'+CURR_LANG

    create_directory(images_folder)
    return lang_folder, images_folder, image_pool_folder, ppt_folder

def create_directory(directory):
    if not os.path.exists(directory):
        os.makedirs(directory)

def ungroup_all_shapes(each_slide_object):
    print('ungrouping now.')
    print('if the terminal holds up. Please check the MS PPT Window.')
    previous_shapes = len(each_slide_object.Shapes)
    if previous_shapes > THRESH_HOLD_GP:
        print('ungrouping prev cutt off')
        return False
    for each_shape in each_slide_object.Shapes:
        try:
            each_shape.Ungroup()
        except:
            continue
        current_len = len(each_slide_object.Shapes)
        if current_len > THRESH_HOLD_GP:
            print('ungrouping cut- off')
            return False
    print('ungrouping complete')
    return True

def populate_links_have(lang_folder):
    _set = set()
    for each_file in os.listdir(lang_folder):
        if'transcription_' in each_file:
            filename = os.path.join(lang_folder, each_file)
            lnk_file = open(filename, 'r')
            present_lines = lnk_file.readlines()
            for each_line in present_lines:
                curr_entry = each_line.rstrip()
                if "SlideName" in curr_entry :
                    split_data = curr_entry.split(" - ")[1].strip()
                    _set.add(split_data)
            lnk_file.close()
    return _set

def process_these_shapes(to_be_processed_shapes, each_slide_object, image_pool_folder):
    print('    processing shapes, replacing images and deleting the rest')
    images_placed_set = set()
    for each_shape in to_be_processed_shapes:

        if(each_shape.Width < 10 or each_shape.Height < 10):
            delete_this_shape(each_shape)
            continue
            
        if((each_shape.Left, each_shape.Top, each_shape.Width, each_shape.Height) not in images_placed_set):
            ran_image = random.choice(os.listdir(image_pool_folder))
            sh_obj = each_slide_object.Shapes.AddPicture(os.path.join(image_pool_folder, ran_image), True, False, # change 
                                            each_shape.Left, each_shape.Top, each_shape.Width, each_shape.Height)
            while sh_obj.ZOrderPosition > 1:
                sh_obj.ZOrder(3)

            images_placed_set.add((each_shape.Left, each_shape.Top, each_shape.Width, each_shape.Height))
        
        # the shape itself is delected after the images has already been added to the to their positions.
        delete_this_shape(each_shape)
    
    print('    Done processing shapes.')

def delete_this_shape(temp):
    try:
        temp.Delete() 
    except:
        print('exception during delete shape')