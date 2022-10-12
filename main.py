from pptx import *
from tkinter import *
from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT, MSO_VERTICAL_ANCHOR
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.dml.color import RGBColor
from pptx.util import Cm, Pt
from JYPop import Application
import pandas as pd
import math
import os.path

#Merina gave some good feedback, this is her credit.

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

    if fontSize == Lines5:
        textBoxPara.line_spacing = 0.9
    elif fontSize == Lines6:
        textBoxPara.line_spacing = 0.8
    else:
        textBoxPara.line_spacing = 0.7

    textBoxPara.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
    textBoxPara.font.name = 'Calibri (Body)'
    textBoxPara.font.size = Pt(fontSize)
    textBoxPara.text = line

#checks the line count
def checkLineCount(splitPara, count, limit):
    for line in splitPara:
        if len(line) > limit:
            count = count + 0.7
    count = round(count)
    return count

#adds a jy icon link
def add_hyper_jy(pres, slide, link):
    pic = slide.shapes.add_picture(jy_icon_path, img_left, img_top, img_height)
    add_link(pres, pic, link)

#adds a link to a shape
def add_link(pres, shape, link):
    click = shape.click_action
    click.target_slide = pres.slides[link]
    click.action

def add_Text_Frame(slide, left, top, width, height, vAlign):
    textBox = slide.shapes.add_textbox(left, top, width, height)
    textFrame = textBox.text_frame
    textFrame.vertical_anchor = vAlign
    
    return textFrame

def add_Heading(textFrame, fontName, fontSize, lineSpacing, hAlign, text):
    textF_text = textFrame.paragraphs[0]
    textF_text.font.name = fontName
    textF_text.font.size = Pt(fontSize)
    textF_text.line_spacing = lineSpacing
    textF_text.alignment = hAlign

    textF_text.text = text


#to determine font size based on number of lines
Lines5 = 60
Lines6 = 56
Lines7 = 51

#character limit based on font size
Limit5 = 30
Limit6 = 33
Limit7 = 38

#textbox dimensions and location


headerLeft = Cm(4.23)
headerTop = Cm(3.12)
headerWidth = Cm(25.4)
headerHeight = Cm(6.63)

subsongTitles_left = Cm(4.23)
subsongTitles_top = Cm(14.97)
subsongTitles_width = Cm(25.4)
subsongTitles_height = Cm(1.25)

jy_icon_path = os.path.join("extra", "JY-Icon-White.png")

img_title_left = Cm(15.61)
img_title_top = Cm(11.96)
title_img_height = Cm(2.64)

img_left = Cm(16.23)
img_top = Cm(17.37)
img_height = Cm(1.4)

indexBoxLeft = [Cm(2.48), Cm(17.08)]
indexBoxTop = [Cm(3.55), Cm(4.77), Cm(6), Cm(7.23), Cm(8.45), Cm(9.68), Cm(10.9), Cm(12.13), Cm(13.36), Cm(14.58)]
indexBoxWidth = Cm(14.3)
indexBoxHeight = Cm(1.04)

indexHeading_fontName = 'Calibri Light (Headings)'
indexHeading_fontSize = 48
indexHeading_lineSpacing = 0.9
indexHeading_HAlign = PP_PARAGRAPH_ALIGNMENT.CENTER
indexHeading_VAlign = MSO_VERTICAL_ANCHOR.MIDDLE
indexHeading_text = "Index"
indexHeading_left = Cm(13.75)
indexHeading_top = Cm(0.54)
indexHeading_width = Cm(6.22)
indexHeading_height = Cm(2.13)

songTitles_fontName = 'Calibri Light (Headings)'
songTitles_fontSize = 80
songTitles_lineSpacing = 0.9
songTitles_HAlign = PP_PARAGRAPH_ALIGNMENT.CENTER
songTitles_VAlign = MSO_VERTICAL_ANCHOR.MIDDLE
songTitles_left = Cm(2.27)
songTitles_top = Cm(4.48)
songTitles_width = Cm(29.21)
songTitles_height = Cm(9.68)

verse_fontName = 'Calibri (Body)'
verse_HAlign = PP_PARAGRAPH_ALIGNMENT.CENTER
verse_VAlign = MSO_VERTICAL_ANCHOR.MIDDLE
verse_left = Cm(2.42)
verse_top = Cm(2)
verse_width = Cm(29.01)
verse_height = Cm(14.5)

df = pd.read_csv(os.path.join("extra", "database.csv"))
try: 
    url = f'https://docs.google.com/spreadsheets/d/1P3Qu1EQLgcQYWSZQwjY5OWmEnnJMvSSgLkasa6rMC6E/gviz/tq?tqx=out:csv'
    df = pd.read_csv(url)
except:
    df = pd.read_csv(os.path.join("extra", "database.csv"))

allSongs = df.values.tolist()
fullSongList = df['Song'].tolist()

x = 0
for song in allSongs:
    allSongs[x] = [x for x in song if str(x) != 'nan']
    x = x + 1

window = Tk()
window.title("JY Music Slides Generator")
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

pr1 = Presentation(os.path.join("extra", "MusicSlidesTemplate.pptx"))

slide1_register = pr1.slide_layouts[6]

startSlide = pr1.slides.add_slide(slide1_register)
startTextBox = startSlide.shapes.add_textbox(headerLeft, headerTop, headerWidth, headerHeight)
startBoxText = startTextBox.text_frame
startBoxText.vertical_anchor = MSO_VERTICAL_ANCHOR.BOTTOM
startText = startBoxText.paragraphs[0]

startText.font.name = 'Calibri Light (Headings)'
startText.font.size = Pt(72)
startText.line_spacing = 0.9
startText.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER

presentationName = Application.presentationName
tempName = presentationName

while len(presentationName) > 25:
    index = presentationName[:25].rindex(' ')
    tempName = presentationName[:index] + '\n' + presentationName[index:]
    break

startText.text = tempName

pic = startSlide.shapes.add_picture(jy_icon_path, img_title_left, img_title_top, title_img_height)

subTextBox = startSlide.shapes.add_textbox(subsongTitles_left, subsongTitles_top, subsongTitles_width, subsongTitles_height)
subBoxText = subTextBox.text_frame
subBoxText.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
subText = subBoxText.paragraphs[0]

subText.font.name = 'Calibri (Body)'
subText.font.size = Pt(24)
subText.line_spacing = 0.9
subText.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER

subText.text = "Jesus Youth Australia"

indexPages = math.ceil(len(chosenSongs)/20)
indexSlideIndex = []
presIndex = []

for x in range(indexPages):
    indexSlide = pr1.slides.add_slide(slide1_register)
    indexSlideIndex.append(indexSlide)

    indexTitleTF = add_Text_Frame(indexSlide, indexHeading_left, indexHeading_top, indexHeading_width, indexHeading_height, indexHeading_VAlign)
    add_Heading(indexTitleTF, indexHeading_fontName, indexHeading_fontSize, indexHeading_lineSpacing, indexHeading_HAlign, indexHeading_text)
    add_hyper_jy(pr1, indexSlide, 0)


for song in chosenSongs:
    titleSlide = pr1.slides.add_slide(slide1_register)
    
    hyperIndex = math.floor(chosenSongs.index(song)/20) + 1
    songName = song
    while len(song) > 25:
        #lineCount = lineCount + 1
        index = song[:25].rindex(' ')
        songName = song[:index] + '\n' + song[index:]
        break
    songIndex = fullSongList.index(song)
    lyrics = allSongs[songIndex][1:]

    songTitleTF = add_Text_Frame(titleSlide, songTitles_left, songTitles_top, songTitles_width, songTitles_height, songTitles_VAlign)
    add_Heading(songTitleTF, songTitles_fontName, songTitles_fontSize, songTitles_lineSpacing, songTitles_HAlign, songName)    
    add_hyper_jy(pr1, titleSlide, hyperIndex)

    presIndex.append(pr1.slides.index(titleSlide))
    for verse in lyrics:
        verseSlide = pr1.slides.add_slide(slide1_register)
        
        verseTf = add_Text_Frame(verseSlide, verse_left, verse_top, verse_width, verse_height, verse_VAlign)
        addPara(verseTf, verse)
        add_hyper_jy(pr1, verseSlide, hyperIndex)

indexIndex = 0
for x in indexSlideIndex:
    for y in indexBoxLeft:
        for z in indexBoxTop:
            indexRect = x.shapes.add_shape(MSO_SHAPE_TYPE.AUTO_SHAPE, y, z, indexBoxWidth, indexBoxHeight)
            
            indexRect.fill.background()
            indexRect.line.fill.background()
            
            rectText = indexRect.text_frame
            rectText.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
            indexText = rectText.paragraphs[0]            

            indexText.font.name = 'Calibri (Body)'
            indexText.font.size = Pt(24)
            indexText.alignment = PP_PARAGRAPH_ALIGNMENT.LEFT

            try:
                add_link(pr1, indexRect, presIndex[indexIndex])

                if indexIndex <9:
                    indexText.text = str(indexIndex + 1) + '.    ' + chosenSongs[indexIndex]
                elif indexIndex <99:
                    indexText.text = str(indexIndex + 1) + '.  ' + chosenSongs[indexIndex]
                else:
                    indexText.text = str(indexIndex + 1) + '.' + chosenSongs[indexIndex]
            except:
                pass

            indexIndex = indexIndex + 1


if Application.save == 1:
    pr1.save(savePath)

