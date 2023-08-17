import os

filePath = "C:/Users/User/Pictures/Test/"
count = 0

for file in os.listdir(filePath):
    newName = str(count) + "Photo.jpg"
    oldName = filePath + file
    newName = filePath + newName
    os.rename(oldName, newName)
    count += 1

