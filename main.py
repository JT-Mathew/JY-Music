from pptx import *
from tkinter import *
from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT, MSO_VERTICAL_ANCHOR
from pptx.util import Cm, Pt
from JYPop import Application
import pandas as pd

#adds verse to slide
def addPara(textBoxText, para):
    splitPara = para.splitlines()
    lineCount = len(splitPara)
    counter = lineCount
    newPara = []
    fontSize = Lines5
    characterLimit = Limit5
    
    if len(splitPara) == 6:
        fontSize = Lines6
        characterLimit = Limit6
    elif len(splitPara) > 6:
        fontSize = Lines7
        characterLimit = Limit7
    
    if counter < 6:
        counter = lineCount
        counter = checkLineCount(splitPara, counter, Limit5)

        if counter >= 6:
            counter = 6
        
    if counter == 6:
        counter = lineCount
        counter = checkLineCount(splitPara, counter, Limit6)
        
        if counter > 6:
            counter = 7
        else:
            fontSize = Lines6
            characterLimit = Limit6

    if counter > 6:
        counter = lineCount
        fontSize = Lines7
        characterLimit = Limit7

    for line in splitPara:
        while len(line) > characterLimit:
            #lineCount = lineCount + 1
            index = line[:characterLimit].rindex(' ')
            line = line[:index] + '\n' + line[index:]
            break
        newPara.append(line)
    
    splitPara = newPara

    count = 0

    for line in splitPara:
        addLine(textBoxText, line, count, fontSize)
        count = 1

#adds a line to slide
def addLine(textBoxText, line, count, fontSize):
    if count == 0:
        textBoxPara = textBoxText.paragraphs[0]
    else:
        textBoxPara = textBoxText.add_paragraph()
    textBoxPara.space_after = Pt(10)
    textBoxPara.line_spacing = 0.8
    textBoxPara.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
    textBoxPara.font.name = 'Calibri (Body)'
    textBoxPara.font.size = Pt(fontSize)
    textBoxPara.text = line

#checks the line count
def checkLineCount(splitPara, count, limit):
    check = 0
    for line in splitPara:
        if len(line) > limit:
            if check == 0:
                count = count + 1
                check = 1
            else:
                check = 0
    return count

#to determine font size based on number of lines
Lines5 = 60
Lines6 = 56
Lines7 = 51

#character limit based on font size
Limit5 = 30
Limit6 = 33
Limit7 = 38

#textbox dimensions and location
left = Cm(2.42)
top = Cm(2)
width = Cm(29.01)
height = Cm(14.5)

titleLeft = Cm(2.27)
titleTop = Cm(4.48)
titleWidth = Cm(29.21)
titleHeight = Cm(9.68)


df = pd.read_csv("extra/database.csv")
try: 
    url = f'https://docs.google.com/spreadsheets/d/1P3Qu1EQLgcQYWSZQwjY5OWmEnnJMvSSgLkasa6rMC6E/gviz/tq?tqx=out:csv'
    df = pd.read_csv(url)
except:
    df = pd.read_csv("extra/database.csv")

allSongs = df.values.tolist()
fullSongList = df['Song'].tolist()

x = 0
for song in allSongs:
    allSongs[x] = [x for x in song if str(x) != 'nan']
    x = x + 1

window = Tk()
window.title("JY Australia Music Slides")
Application.getWindow(window)
Application.saveSongList(fullSongList)
app = Application(master=window)
app.mainloop()

chosenSongs = Application.song_List
savePath = Application.filepath

#allSongs: database
#fullSongList: list of all song names
#chosenSongs: list of selected songs
#songIndex: index of all the selected songs

pr1 = Presentation("extra/MusicSlidesTemplate.pptx")

slide1_register = pr1.slide_layouts[6]

startSlide = pr1.slides.add_slide(slide1_register)
startTextBox = startSlide.shapes.add_textbox(titleLeft, titleTop, titleWidth, titleHeight)
startBoxText = startTextBox.text_frame
startBoxText.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
startText = startBoxText.paragraphs[0]

startText.font.name = 'Calibri Light (Headings)'
startText.font.size = Pt(80)
startText.line_spacing = 0.9
startText.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER

presentationName = Application.presentationName
tempName = presentationName

while len(presentationName) > 25:
    index = presentationName[:25].rindex(' ')
    tempName = presentationName[:index] + '\n' + presentationName[index:]
    break

startText.text = tempName


for song in chosenSongs:
    titleSlide = pr1.slides.add_slide(slide1_register)
    titleTextBox = titleSlide.shapes.add_textbox(titleLeft, titleTop, titleWidth, titleHeight)
    titleBoxText = titleTextBox.text_frame
    titleBoxText.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
    titleText = titleBoxText.paragraphs[0]

    titleText.font.name = 'Calibri Light (Headings)'
    titleText.font.size = Pt(80)
    titleText.line_spacing = 0.9
    titleText.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
    
    songName = song

    while len(song) > 25:
        #lineCount = lineCount + 1
        index = song[:25].rindex(' ')
        songName = song[:index] + '\n' + song[index:]
        break

    titleText.text = songName

    songIndex = fullSongList.index(song)
    lyrics = allSongs[songIndex][1:]

    for verse in lyrics:
        slide = pr1.slides.add_slide(slide1_register)
        textBox = slide.shapes.add_textbox(left, top, width, height)
        textBoxText = textBox.text_frame
        addPara(textBoxText, verse)

pr1.save(savePath)

