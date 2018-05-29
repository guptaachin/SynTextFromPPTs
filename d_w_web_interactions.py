# imports
import googleapiclient.discovery
from urllib.request import urlopen
import os
import urllib
import socket
socket.setdefaulttimeout(30)

# imports
SEPARATOR = '__SEPARATOR__'
# guptaachin01
#apikey = "AIzaSyCmyluZAx9OEqjwi2zKmBfvGnu4KrTEstQ"
#search_eng_id = '009893106914731719450:so9uldhqkus'

# achingupta3000
#apikey = "AIzaSyCcX5bGHUDcfus-VBu0x0TnUDaXB1SbEa8"
#search_eng_id = "016135740530810151881:rfvtqvhszgm"

# achingupta1756
#apikey = "AIzaSyAqK5IMRg8lM33laLR_8i2wtf9ooe-ivTU"
#search_eng_id = "000402284428032715281:9tf38txb44o"
#
# achingupta1757
#apikey = "AIzaSyAB06oSYNb212mEvauGHkieqkXV_wxR3i0"
#search_eng_id = "013024238142372970011:tcdycx5c9gy"
#
# achingupta1758
apikey = "AIzaSyBmQ4govoQdiAdYFepvsLOgI8RQK7Df5zI"
search_eng_id = "003684733956413315411:4-am4d9jjty"
#
# achingupta1759
#apikey = "AIzaSyDmR8JqGz2MzU-9XOqIdOzIU-ig6ScEELg"
#search_eng_id = "013649291752998755986:nluep9yq7ks"

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