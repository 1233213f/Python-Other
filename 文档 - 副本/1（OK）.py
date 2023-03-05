from sys import argv
import os
import urllib.request
import mimetypes

# default constants
# 输出文件夹，下载的图片将会保存在这个文件夹
outputFolder = 'E:/Github/'
os.makedirs(outputFolder, exist_ok=True)
# 输入文件夹，图片的url链接保存在url.txt中
inputFile = 'E:/Github/urls.txt'

# read file from argument or keeping default
if len(argv) == 2:
    inputFile = argv[1]

# tests if provided file is a text file1
if mimetypes.guess_type(inputFile)[0] != 'text/plain':
    print ('File is not a valid text file')
    exit()

print ('Running Python Downloader...')

try:
    file = open(inputFile)
    # looping over the file: line by line
    with file as f:
        for line in f:
            # taking everything after the last / as filename
            # outputFolder has to exist beforehand
            filename = outputFolder + line.rsplit('/', 1)[-1].strip()
            print ('\nTrying to download: ' + line.strip())
            try:
                # downloading file, continues with error message when not found
                urllib.request.urlretrieve(line, filename)
                print ('Saved as ' + filename)
            except IOError:
                print ('Could not find file')
                pass
            except Exception:
                continue
except IOError:
    print ('Could not open file: ' + inputFile)