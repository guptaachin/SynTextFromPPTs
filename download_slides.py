# -*- coding: utf-8 -*-
import web_interactions
from web_interactions import *
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
    global count

    unique = {}

    files_present = os.listdir(base_data_folder)
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

            for link in links_list:
                if link not in unique:
                    file_path = os.path.join(store_location, 'sl_') + _ln + '_' + word + '_' + str(count) + '.ppt'
                    file_name = 'sl_' + _ln + '_' + word + '_' + str(count) + '.ppt'
                    count += 1
                    # the link might be different
                    unique[link] = 1
                    # not doing anything with this for now

                    if file_name in files_present:
                        print('skipping,',file_name)
                        continue

                    url = gac.download(link, file_path)

                    if url is not None:
                        print('writing ***************************')
                        print(url)
                        links_nd_store.write(url+'\n')
                        links_nd_store.flush()
                        os.fsync(links_nd_store.fileno())

                links_d_store.write(link+'\n')
                links_d_store.flush()
                os.fsync(links_d_store.fileno())


    lang_file.close()
    words_file.close()
    links_d_store.close()
    links_nd_store.close()

if __name__ == '__main__':
    # passing the values since i do not want to download
    # pass
    main()