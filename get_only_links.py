# -*- coding: utf-8 -*-
import web_interactions
from web_interactions import *
import os
import argparse

base_data_folder = os.path.join(os.getcwd(), 'data')
keywords_file_path = os.path.join(base_data_folder, 'new_words.txt')

SEPARATOR = '__SEPARATOR__'

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("language", help="lang_ja,lang_ko,lang_es")
    args = parser.parse_args()
    CURR_LANG = args.language

    links_path = os.path.join(base_data_folder, 'links_'+CURR_LANG+'.txt')
    links_have, words_have = populate_links_have(links_path)

    gac = web_interactions.Google_Api()

    # lang_file = open(lang_file_path, 'r')
    words_file = open(keywords_file_path, 'r')
    links_d_store = open(links_path, 'a')

    l_lines = [CURR_LANG]
    w_lines = words_file.readlines()

    for lang_index in range(0, len(l_lines)):
        for word_index in range(0, len(w_lines)):

            word = w_lines[word_index].rstrip().strip()
            lang = l_lines[lang_index].rstrip().strip()

            if (word in words_have):
                continue

            print('new word encountered = ',word,' for lang = ',lang)

            # make a google api call
            links_list, call_successful = gac.get_rest_object(word, lang)

            # record all the cases when the words did not return any result.
            if(call_successful and len(links_list) == 0):
                print('writing for bogus word - ','None '+word)
                links_list.append('None '+word) # will append 'None fallacious'

            for link in links_list:
                print('writing ',link)
                if link not in links_have:
                    links_have.add(link)
                    string_to_store = lang+SEPARATOR+word+SEPARATOR+link
                    write_n_flush(links_d_store, string_to_store)

    words_file.close()
    links_d_store.close()

def write_n_flush(links_d_store, link):
    links_d_store.write(link + '\n')
    links_d_store.flush()
    os.fsync(links_d_store.fileno())


def populate_links_have(links_path):
    s = set()
    w = set()
    lnk_file = open(links_path, 'r')
    present_lines = lnk_file.readlines()
    for each_line in present_lines:
        curr_entry = each_line.rstrip()
        cur_sep = curr_entry.split(SEPARATOR)
        s.add(cur_sep[2])
        w.add(cur_sep[1])
    return s,w

if __name__ == '__main__':
    main()