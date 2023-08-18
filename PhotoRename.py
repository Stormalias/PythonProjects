import os
from PIL import Image

filePath = "C:/Users/User/Pictures/Test/"
count = 0

def renamer():
    for file in os.listdir(filePath):
        newName = str(count) + "Photo.jpg"
        oldName = filePath + file
        newName = filePath + newName
        os.rename(oldName, newName)
        count += 1


with Image.open("C:/Users/The Weenus Machine/Pictures/4669fd1b60d9592b3191d226918cc766.png") as im:
    im.show()

