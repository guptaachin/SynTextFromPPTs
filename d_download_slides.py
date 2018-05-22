# -*- coding: utf-8 -*-
import web_interactions
from web_interactions import *
import os
######################Arguments ##############################################
# the keywords should be pointed to by the file

base_data_folder = os.path.join(os.getcwd(), 'data')
keywords_file_path = os.path.join(base_data_folder, 'new_words.txt')

links_not_down_path = os.path.join(base_data_folder, 'links_not_down.txt')
links_down_path = os.path.join(base_data_folder, 'links_down.txt')
lang_file_path = os.path.join(base_data_folder, 'lang.txt')

lang_mapping = {'ja':'Japanese', 'ko':'Korean', 'es': 'Spanish'}


##############################################################################

def main():
    # Build a service object for interacting with the API. Visit
    # the Google APIs Console <http://code.google.com/apis/console>
    # to get an API key for your own application.

    count = 0
    # code to prevent repetitive downloads
    dir_counter, file_counter = 0, 0
    directories = get_list_of_directories(base_data_folder)
    best_dir_index = get_eligible_directory(directories)
    dir_counter = best_dir_index
    create_directory(os.path.join(base_data_folder, str(dir_counter)))

    unique = {}

    downloaded_set = populate_downloaded_set()

    gac = web_interactions.Google_Api()

    lang_file = open(lang_file_path, 'r')
    words_file = open(keywords_file_path, 'r')
    links_d_store = open(links_down_path, 'a')
    links_nd_store = open(links_not_down_path, 'a')

    l_lines = lang_file.readlines()
    w_lines = words_file.readlines()

    n_calls = 0

    for lang in l_lines:
        count = 30000
        for word in w_lines:

            word = word.rstrip().strip()
            lang = lang.rstrip().strip()
            # make a google api call
            links_list = gac.get_rest_object(word, lang)
            word = "+".join(word.split(' '))
            _ln = lang.split('_')[1]
            n_calls += 1

            print('calls so far = ',n_calls)

            url = None

            for link in links_list:

                # deciding the final directory
                curr_path = os.path.join(base_data_folder, str(dir_counter))

                files_in_this_folder = get_number_files(curr_path)
                file_counter = files_in_this_folder
                if(files_in_this_folder > 9999):
                    dir_counter += 1
                    file_counter = 0

                    curr_path = os.path.join(base_data_folder, str(dir_counter))
                    create_directory(curr_path)

                files_present = os.listdir(curr_path)

                if link not in unique:
                    file_counter += 1
                    file_path = os.path.join(curr_path, 'sl_') + _ln + '_' + word + '_' + str(count) + '.ppt'
                    file_name = 'sl_' + _ln + '_' + word + '_' + str(count) + '.ppt'
                    count += 1
                    file_counter += 1
                    # the link might be different
                    unique[link] = 1

                    # assuming that the order in which the links are returned will always be the same

                    if file_name in files_present:
                        write_n_flush(links_d_store, link)
                        continue

                    # url = gac.download(link, file_path)
                    #
                    # # for a successfully downloaded file url is None
                    # if url is not None:
                    #     file_counter -= 1
                    #     write_n_flush(links_nd_store, url)

            if url is None:
                write_n_flush(links_d_store, link)

    lang_file.close()
    words_file.close()
    links_d_store.close()
    links_nd_store.close()

def populate_downloaded_set():
    ds = set()

    with open(links_down_path, 'r') as f:
        for line in f:
            line = line.rstrip()
            ds.add(line)

    return ds

def write_n_flush(links_d_store, link):
    links_d_store.write(link + '\n')
    links_d_store.flush()
    os.fsync(links_d_store.fileno())

def create_directory(directory):
    if not os.path.exists(directory):
        os.makedirs(directory)

def get_list_of_directories(mypath):
    onlydirs = [f for f in os.listdir(mypath) if os.path.isdir(os.path.join(mypath, f))]
    return onlydirs

def get_number_files(mypath):
    list_of_files = os.listdir(mypath)
    return len(list_of_files)

def get_eligible_directory(directories):
    best_dir = -1
    for each_dir in directories:
        mypath = os.path.join(base_data_folder, each_dir)
        number_files = get_number_files(mypath)
        if(number_files < 9999):
            best_dir = int(each_dir)
            break

    if(best_dir == -1):
        best_dir = int(each_dir)
        best_dir += 1

    return best_dir


if __name__ == '__main__':
    # passing the values since i do not want to download
    # pass
    main()