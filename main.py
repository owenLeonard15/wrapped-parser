from posixpath import dirname
import cv2
import pytesseract
import xlsxwriter
from tqdm import tqdm 
# import os

# configure xlsxwriter and pytesseract
workbook = xlsxwriter.Workbook('genzdesigns_wrapped.xlsx')
worksheet = workbook.add_worksheet()
pytesseract.pytesseract.tesseract_cmd = r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'

# the array to store all results in
results = []

# rename image files --- only once
# dirname = './wrapped_images copy'
# for i, filename in enumerate(os.listdir(dirname)):
#     os.rename(dirname + '/' + filename, dirname + '/' + str(i) + '.jpg')

# process every image, use progress bar (tqdm) for live tracking
for i in tqdm(range(870)):
    img = cv2.imread('./wrapped_images copy/' + str(i) + '.jpg')

    # convert to grayscale
    grayImage = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # convert to black and white
    (thresh, blackAndWhiteImage) = cv2.threshold(grayImage, 127, 255, cv2.THRESH_BINARY)
    inverted = cv2.bitwise_not(blackAndWhiteImage)

    # extract text from image
    data = pytesseract.image_to_string(inverted)

    # separate values into array separated by #
    # the first artist should come after the first hash, and last item will include everything from the last song onwards
    # any hashes within artist name or song name will break this 
    arr = data.split("#")
    length = len(arr)
    artists = [] 
    songs = []
    for i in range(1,length-1):
        if i % 2 == 1:
            artists.append(arr[i][2:])
        else:
            songs.append(arr[i][2:])
            
    # clean data
    for i in range(len(songs)):
        songs[i] = songs[i].split("\n")[0]
        songs[i] = songs[i].split('...')[0]
        artists[i] = artists[i].split('...')[0]


    # song 5
    song5 = arr[len(arr)-1].split('\n')[0][2:].split('...')[0]
    songs.append(song5)

    # time and genre
    time, genre = '',''
    endarr = arr[len(arr)-1].split('\n')
    for i in range (len(endarr)):
        if len(endarr[i]) > 0:
            if endarr[i][0].isdigit() and (endarr[i][2].isdigit() or endarr[i][3].isdigit()):
                time = endarr[i].split(' ')[0]
                genre = endarr[i][len(time)+1:]

    # put it all together
    everything = artists + songs
    if len(time) > 0:
        everything.append(time)
    if len(genre) > 0: 
        everything.append(genre)

    # if length < 12, then the columns were read vertically and we need to take the alternate route
    if len(everything) == 12:
        results.append(everything)
    else:
        arr = data.split("#")
        length = len(arr)
        artists = [] 
        songs = []
        for i in range(1,length):
            if i < 6:
                artists.append(arr[i])
            else:
                songs.append(arr[i])
        
        # clean data
        for i in range(len(songs)):
            songs[i] = songs[i].split("\n")[0]
            songs[i] = songs[i].split('...')[0]
            songs[i] = songs[i][2:]
        for i in range(len(artists)):
            artists[i] = artists[i].split("\n")[0]
            artists[i] = artists[i].split('...')[0]
            artists[i] = artists[i][2:]

        # time 
        time = ''
        if(len(arr) > 4):
            timeArr = arr[5].split('\n')
            for i in range (len(timeArr)):
                if len(timeArr[i]) > 0:
                    if len(timeArr[i]) > 3 and timeArr[i][0].isdigit() and (timeArr[i][2].isdigit() or timeArr[i][3].isdigit()):
                        time = timeArr[i].split(' ')[0]
        
        # genre
        genreArr = arr[len(arr)-1].split('\n')
        genre = genreArr[len(genreArr)-2]

        # put it all back together
        everything = artists + songs
        everything.append(time)
        everything.append(genre)

# add to spreadsheet
col = 0
for row, data in enumerate(results):
    worksheet.write_row(row, col, data)

workbook.close()