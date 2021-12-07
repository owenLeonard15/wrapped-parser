# spotifywrapped-parser

A script used to convert hundreds of spotify wrapped summary screenshots taken from instagram stories into data in an excel spreadsheet for further analysis. 

The images are first optimized for contrast and readability by processing them from color into grayscale, then black and white, and lastly they are inverted. The text in the processed images is then parsed for the desired values and read into an array using pytesseract. Finally, these arrays are written to an excel spreadsheet to be sent to the client. 