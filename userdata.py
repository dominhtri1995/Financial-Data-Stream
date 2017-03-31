import socket
import time
import datetime
import requests
import json

# config = {
#   "apiKey": "",
#   "authDomain": "s3463625.firebaseapp.com",
#   "databaseURL": "https://s3463625.firebaseio.com/",
#   "storageBucket": "s3463625.appspot.com"
# }

def get_ip():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    s.connect(("8.8.8.8", 80))
    try:
        push_firebase("ip", s.getsockname()[0])
    except:
       pass
    s.close()

def get_computer_name():
    return socket.gethostname()


def push_firebase(name,item):
    offset = time.timezone if (time.localtime().tm_isdst == 0) else time.altzone

    data = {name:item,"time":datetime.datetime.fromtimestamp(time.time()).strftime('%Y-%m-%d %H:%M:%S') \
            , "gmt":str(offset / 60 / 60 * -1),"computer_name":get_computer_name()}

    requests.post("https://s3463625.firebaseio.com/ip.json",data=json.dumps(data))
