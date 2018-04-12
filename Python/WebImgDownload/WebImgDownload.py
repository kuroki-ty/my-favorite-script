import urllib.error
import urllib.request
import os

def download_image(url, dst_path):
    try:
        data = urllib.request.urlopen(url).read()
        with open(dst_path, mode="wb") as f:
            f.write(data)
    except urllib.error.URLError as e:
        print(e)

url = 'http://cdn-ak.f.st-hatena.com/images/fotolife/m/mosshm/20080720/20080720184432.jpg'
dst_path = 'img/neko.jpg'
# dist_dir = 'img/'
# dst_path = os.path.join(dist_dir, os.path.basename(url))
download_image(url, dst_path)