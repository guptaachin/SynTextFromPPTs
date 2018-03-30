# imports
import googleapiclient.discovery
from urllib.request import urlopen
import os
from pprint import pprint
# imports
store_location = os.getcwd()+'/data'

apikey = "AIzaSyCcX5bGHUDcfus-VBu0x0TnUDaXB1SbEa8" # this is the apikey
search_eng_id = '016135740530810151881:rfvtqvhszgm' #Please read README to get api key

# is a singleton class
class Google_Api:

    def __init__(self):
        try:
            self.service = googleapiclient.discovery.build("customsearch", "v1", developerKey=apikey)
        except Exception as e:
            print('exception Google_api, constructor - ',e)

    def get_rest_object(self, word, language):
        question = 'filetype:ppt '+word

        try:
            res = self.service.cse().list(
              q=question,
              cx=search_eng_id,
              lr=language
            ).execute()
        except Exception as e:
            print("ERROR OCCURRED for lang - ", language, 'word ',word,' error = ', e)
            return

        return self.get_links(res)

    def get_links(self, res):
        items = res.get('items', [])
        links_list = []
        for item in items:
            links_list.append(item['link'])

        return links_list

    def download(self, url, file_name):
        try:
            u = urlopen(url)
        except Exception as e:
            # do not download this ppt.
            print('Not downloading = ', url)
            print('error = ',e)
            return url + '__SEP__' +str(e)
        f = open(file_name, 'wb')
        print(file_name)
        meta = u.getheaders()
        value = ''
        for item in meta:
            if item[0] == 'Content-Length':
                value = int(item[1])
        if not value:
            return
        file_size = value
        print("Downloading: %s Bytes: %s" % (file_name, file_size))

        file_size_dl = 0
        block_sz = 8192
        while True:
            buffer = u.read(block_sz)
            if not buffer:
                break

            file_size_dl += len(buffer)
            f.write(buffer)
            status = r"%10d  [%3.2f%%]" % (file_size_dl, file_size_dl * 100. / file_size)
            status = status + chr(8) * (len(status) + 1)
            # print status,

        f.close()
        print("Done")

        return None