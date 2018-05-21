# -*- coding: utf-8 -*-
import web_interactions
from web_interactions import *
from get_only_links import SEPARATOR
import os
import argparse
base_data_folder = os.path.join(os.getcwd(), 'data')


def main():

    parser = argparse.ArgumentParser()
    parser.add_argument("language", help="lang_ja,lang_ko,lang_es")
    args = parser.parse_args()
    CURR_LANG = args.language

    links_not_down_path = os.path.join(base_data_folder, 'links_not_down_'+CURR_LANG+'.txt')
    links_already_dowloaded_path = os.path.join(base_data_folder, 'links_downloaded_'+CURR_LANG+'.txt')
    links_down_path = os.path.join(base_data_folder, 'links_'+CURR_LANG+'.txt')

    links_have = populate_links_have(links_already_dowloaded_path)

    links_to_down_store = open(links_down_path, 'r')
    links_nd_store = open(links_not_down_path, 'a')
    links_already_file = open(links_already_dowloaded_path, 'a')

    l_lines = links_to_down_store.readlines()

    my_path = os.path.join(base_data_folder, CURR_LANG)

    create_directory(my_path)

    number_of_files = get_number_files(my_path)

    count = number_of_files + 3000
    for lnk in l_lines:
        link_arr = lnk.rstrip().split(SEPARATOR)
        link = link_arr[2]
        curr_path = os.path.join(base_data_folder, link_arr[0])
        file_path = os.path.join(curr_path, 'sl_') + link_arr[0]+'_'+link_arr[1] + '_' + str(count) + '_'+'.ppt'

        if(link in links_have):
            print('skipping')
            continue

        # download the file from this link
        url = web_interactions.Google_Api.download(link, file_path)

        if url is not None:
            if(url is not 'None'):
                write_n_flush(links_nd_store, url)
        else:
            write_n_flush(links_already_file, link)

        count += 1
    links_to_down_store.close()
    links_nd_store.close()
    links_already_file.close()


def write_n_flush(links_d_store, link):
    links_d_store.write(link + '\n')
    links_d_store.flush()
    os.fsync(links_d_store.fileno())


def create_directory(directory):
    if not os.path.exists(directory):
        os.makedirs(directory)


def get_number_files(mypath):
    list_of_files = os.listdir(mypath)
    return len(list_of_files)


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


if __name__ == '__main__':
    # passing the values since i do not want to download
    # pass
    main()