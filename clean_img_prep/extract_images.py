import os
import shutil


data_path = os.path.join(os.getcwd(), 'filtered_flickr_data')
target_path = os.path.join(os.getcwd(), 'image_pool')



list_of_directories = os.listdir(data_path)

for each_dir in list_of_directories:
    list_of_images = os.listdir(os.path.join(data_path, each_dir))
    for each_image in list_of_images:
        shutil.copy(os.path.join(os.path.join(data_path, each_dir), each_image), target_path)
        print('done copying = ',each_image)
        
