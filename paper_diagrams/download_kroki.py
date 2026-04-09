import os
import zlib
import base64
import urllib.request
import glob

def get_kroki_url(text):
    compressed = zlib.compress(text.encode('utf-8'), 9)
    encoded = base64.urlsafe_b64encode(compressed).decode('ascii')
    return f"https://kroki.io/mermaid/png/{encoded}"

def download_diagrams():
    mmd_files = glob.glob("*.mmd")
    for f in mmd_files:
        png_name = f.replace('.mmd', '.png')
        with open(f, 'r', encoding='utf-8') as file:
            content = file.read()
        
        url = get_kroki_url(content)
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        try:
            res = urllib.request.urlopen(req)
            if res.getcode() == 200:
                with open(png_name, 'wb') as img_file:
                    img_file.write(res.read())
                print(f"Successfully downloaded {png_name}")
            else:
                print(f"Failed {png_name}: {res.getcode()}")
        except Exception as e:
            print(f"Error {png_name}: {e}")

if __name__ == "__main__":
    download_diagrams()
