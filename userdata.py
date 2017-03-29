import socket
import pyrebase
import time
import datetime

config = {
  "apiKey": "AIzaSyBoD1M_77br2rFsIGj_0mQ6qYqT70_gCD",
  "authDomain": "s3463625.firebaseapp.com",
  "databaseURL": "https://s3463625.firebaseio.com/",
  "storageBucket": "s3463625.appspot.com"
}

firebase = pyrebase.initialize_app(config)


def get_ip():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    s.connect(("8.8.8.8", 80))
    try:
        push_firebase("ip",s.getsockname()[0])
    except:
       pass
    s.close()

def get_computer_name():
    return socket.gethostname()

def push_firebase(name,item):
    db = firebase.database()
    offset = time.timezone if (time.localtime().tm_isdst == 0) else time.altzone

    data = {name:item,"time":datetime.datetime.fromtimestamp(time.time()).strftime('%Y-%m-%d %H:%M:%S') \
            , "gmt":str(offset / 60 / 60 * -1),"computer_name":get_computer_name()}

    db.child("IP").push(data)