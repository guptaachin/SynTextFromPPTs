# -*- coding: utf-8 -*-
import web_interactions
from web_interactions import *
from get_only_links import SEPARATOR
import os
######################Arguments ##############################################
# the keywords should be pointed to by the file

base_data_folder = os.path.join(os.getcwd(), 'data')
links_not_down_path = os.path.join(base_data_folder, 'links_not_down.txt')
links_down_path = os.path.join(base_data_folder, 'links.txt')
CURR_DIR = 1
language_list = ['lang_ja',
'lang_ko',
'lang_es']
##############################################################################

def main():
    # Build a service object for interacting with the API. Visit
    # the Google APIs Console4 <http://code.google.com/apis/console>
    # to get an API key for your own application.

    # code to prevent repetitive downloads

    links_d_store = open(links_down_path, 'r')
    links_nd_store = open(links_not_down_path, 'w')

    l_lines = links_d_store.readlines()
    for el in language_list:
        create_directory(os.path.join(base_data_folder, str(el)))

    count = 3000
    for lnk in l_lines:
        link_arr = lnk.rstrip().split(SEPARATOR)
        link = link_arr[2]
        curr_path = os.path.join(base_data_folder, link_arr[0])
        file_path = os.path.join(curr_path, 'sl_') + link_arr[0]+'_'+link_arr[1] + '_' + str(count) + '_'+'.ppt'
        url = web_interactions.Google_Api.download(link, file_path)

        if url is not None:
            write_n_flush(links_nd_store, url)
        count += 1
    links_d_store.close()
    links_nd_store.close()


def write_n_flush(links_d_store, link):
    links_d_store.write(link + '\n')
    links_d_store.flush()
    os.fsync(links_d_store.fileno())


def create_directory(directory):
    if not os.path.exists(directory):
        os.makedirs(directory)


if __name__ == '__main__':
    # passing the values since i do not want to download
    # pass
    main()