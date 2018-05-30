# imports
import googleapiclient.discovery
from urllib.request import urlopen
import os
import urllib
import socket
socket.setdefaulttimeout(30)

# imports
SEPARATOR = '__SEPARATOR__'
apikey = ""
search_eng_id = ''

class Google_Api:

    def __init__(self):
        try:
            self.service = googleapiclient.discovery.build("customsearch", "v1", developerKey=apikey)
        except Exception as e:
            print('exception Google_api, constructor - ',e)

    def get_rest_object(self, word, language):
        print(language, word)

        ques_list = ['filetype:ppt ','filetype:pptx ']
        local_list = []

        for q in ques_list:
            question = q+word
            try:    
                res = self.service.cse().list(
                q=question,
                cx=search_eng_id,
                lr=language
                ).execute()
            except Exception as e:
                print("GPI Error = ", e)
                return [], False

            local_list.append(self.get_links(res))

        return local_list[0]+local_list[1], True

    def get_links(self, res):
        items = res.get('items', [])
        links_list = []
        for item in items:
            links_list.append(item['link'])

        return links_list

        
    @staticmethod
    def download(url, file_name):
        try:
            # u = urlopen(url)
            print('trying to dl ',url)
            urllib.request.urlretrieve(url, file_name)
            print('Downloaded = ', file_name)
        except Exception as e:
            # do not download this ppt.
            print('Not downloading = ', url)
            print('error = ',e)
            return url + SEPARATOR +str(e)
        return None
