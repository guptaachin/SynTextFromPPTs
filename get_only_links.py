# -*- coding: utf-8 -*-
import web_interactions
from web_interactions import *
import os

base_data_folder = os.path.join(os.getcwd(), 'data')
keywords_file_path = os.path.join(base_data_folder, 'new_words.txt')
lang_file_path = os.path.join(base_data_folder, 'lang.txt')

links_path = os.path.join(base_data_folder, 'links.txt')
SEPARATOR = '__SEPARATOR__'

def main():
    # Build a service object for interacting with the API. Visit
    # the Google APIs Console <http://code.google.com/apis/console>
    # to get an API key for your own application.

    links_have = populate_links_have()

    last_google_api = get_last_apicall_l_w()

    gac = web_interactions.Google_Api()
    lang_file = open(lang_file_path, 'r')
    words_file = open(keywords_file_path, 'r')
    links_d_store = open(links_path, 'a')

    l_lines = lang_file.readlines()
    w_lines = words_file.readlines()

    _lan = last_google_api[0]
    _wrd = last_google_api[1]

    for lang_index in range(l_lines.index(_lan+'\n'), len(l_lines)):
        for word_index in range(w_lines.index(_wrd+'\n'), len(w_lines)):

            word = w_lines[word_index].rstrip().strip()
            lang = l_lines[lang_index].rstrip().strip()

            # make a google api call
            links_list = gac.get_rest_object(word, lang)

            for link in links_list:
                if link not in links_have:
                    links_have.add(link)
                    string_to_store = lang+SEPARATOR+word+SEPARATOR+link
                    write_n_flush(links_d_store, string_to_store)

    lang_file.close()
    words_file.close()
    links_d_store.close()


def get_last_apicall_l_w():
    lang_file = open(links_path, 'r')
    present_lines = lang_file.readlines()
    last_entry = present_lines[-1].rstrip()
    last_sep = last_entry.split(SEPARATOR)
    return last_sep[0:2]


def write_n_flush(links_d_store, link):
    links_d_store.write(link + '\n')
    links_d_store.flush()
    os.fsync(links_d_store.fileno())


def populate_links_have():
    s = set()
    lnk_file = open(links_path, 'r')
    present_lines = lnk_file.readlines()
    for each_line in present_lines:
        curr_entry = each_line.rstrip()
        cur_sep = curr_entry.split(SEPARATOR)
        s.add(cur_sep[2])
    return s

if __name__ == '__main__':
    main()