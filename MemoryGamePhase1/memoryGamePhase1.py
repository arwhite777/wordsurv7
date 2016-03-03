from cImage import *
import random

def createImageList(howMany):
    imageList=[]
    for i in range(1,howMany+1,1):
        imageList.append("pic"+str(i)+".gif")
    return imageList

def createImageBoard(howBig):
    imageBoard=[]
    for lst in range(howBig):
        imageBoard.append([""]*howBig)
    return imageBoard
    
def createIndicesForPairs(howBig):
    indicesPairs=[]
    for row in range(howBig):
        for col in range(howBig):
            indicesPairs.append([row,col])
    return indicesPairs

def allocatePics(indicesPairs,imageList,imageBoard):
    #Allocate picture positions            
    while len(indicesPairs)>0:
        #pick a picture
        imName=imageList.pop(random.randrange(0,len(imageList)))
        #put the pick in two random locations
        pair=indicesPairs.pop(random.randrange(len(indicesPairs)))
        imageBoard[pair[0]][pair[1]]=imName
        pair=indicesPairs.pop(random.randrange(len(indicesPairs)))
        imageBoard[pair[0]][pair[1]]=imName

def displayPics(howBig,imageBoard,imageWindow,width,height):
    for col in range(0,howBig):
        for row in range(0,howBig):
            imName=(imageBoard[row][col])
            myIm=FileImage(imName)
            myIm.setPosition(col*width,row*height)
            myIm.draw(imageWindow)

def memoryGame(howBig):
    height=90
    width=90   
    imageList=createImageList(31)
    print("IMAGE LIST =",imageList)
    imageBoard=createImageBoard(howBig)
    print("IMAGE BOARD =",imageBoard)
    indicesPairs=createIndicesForPairs(howBig)
    print("INDEX PAIRS =",indicesPairs)
    allocatePics(indicesPairs,imageList,imageBoard)
    print("IMAGE BOARD =",imageBoard)
    for i in range(len(imageBoard)):
        print("Row",i,imageBoard[i])
    imageWindow = ImageWin("Memory Game", width*howBig, height*howBig) 
    displayPics(howBig,imageBoard,imageWindow,width,height)
    imageWindow.exitOnClick()

memoryGame(4)
    
