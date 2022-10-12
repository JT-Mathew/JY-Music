from pptx import *
from tkinter import *
from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT, MSO_VERTICAL_ANCHOR
from pptx.enum.shapes import MSO_SHAPE_TYPE
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
    fontSize = verse_fontSize_lines5
    characterLimit = verse_charLimit_lines5
    
    if len(splitPara) == 6:
        fontSize = verse_fontSize_lines6
        characterLimit = verse_charLimit_lines6
    elif len(splitPara) > 6:
        fontSize = verse_fontSize_lines7
        characterLimit = verse_charLimit_lines7
    
    if counter < 6:
        counter = lineCount
        counter = checkLineCount(splitPara, counter, verse_charLimit_lines5)

        if counter >= 6:
            counter = 6
        
    if counter == 6:
        counter = lineCount
        counter = checkLineCount(splitPara, counter, verse_charLimit_lines6)
        
        if counter > 6:
            counter = 7
        else:
            fontSize = verse_fontSize_lines6
            characterLimit = verse_charLimit_lines6

    if counter > 6:
        counter = lineCount
        fontSize = verse_fontSize_lines7
        characterLimit = verse_charLimit_lines7

    for line in splitPara:
        line = process_Limit(line, characterLimit)
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

    if fontSize == verse_fontSize_lines5:
        textBoxPara.line_spacing = 0.9
    elif fontSize == verse_fontSize_lines6:
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
    pic = slide.shapes.add_picture(jy_icon_path, norm_JYIcon_left, norm_JYIcon_top, norm_JYIcon_height)
    add_link(pres, pic, link)

#adds a link to a shape
def add_link(pres, shape, link):
    click = shape.click_action
    click.target_slide = pres.slides[link]
    click.action

#adds a textFrame
def add_Text_Frame(slide, left, top, width, height, vAlign):
    textBox = slide.shapes.add_textbox(left, top, width, height)
    textFrame = textBox.text_frame
    textFrame.vertical_anchor = vAlign
    
    return textFrame

#adds a heading to a textFrame
def add_Heading(textFrame, fontName, fontSize, lineSpacing, hAlign, text):
    textF_text = textFrame.paragraphs[0]
    textF_text.font.name = fontName
    textF_text.font.size = Pt(fontSize)
    textF_text.line_spacing = lineSpacing
    textF_text.alignment = hAlign

    textF_text.text = text



#processes a string through a characterlimit
def process_Limit(line, limit):
    tempName = line
    while len(line) > limit:
        index = line[:limit].rindex(' ')
        tempName = line[:index] + '\n' + line[index:]
        break
    return tempName

#removes nan
def clean_Database(df):
    fullDatabase = df.values.tolist()
    x = 0
    for song in fullDatabase:
        fullDatabase[x] = [x for x in song if str(x) != 'nan']
        x = x + 1
    return fullDatabase

#builds the Title Slide
def build_Title_Slide(pres, slide_register):
    titleSlide = pres.slides.add_slide(slide_register)
    startBoxText = add_Text_Frame(titleSlide, title_left, title_top, title_width, title_height, title_VAlign)
    add_Heading(startBoxText, title_fontName, title_fontSize, title_lineSpacing, title_HAlign, presentationName)

    subBoxText = add_Text_Frame(titleSlide, subTitle_left, subTitle_top, subTitle_width, subTitle_height, subTitle_VAlign)
    add_Heading(subBoxText, subTitle_fontName, subTitle_fontSize, subTitle_lineSpacing, subTitle_HAlign, subTitle_text)

    pic = titleSlide.shapes.add_picture(jy_icon_path, title_JYIcon_left, title_JYIcon_top, title_JYIcon_height)

#builds the Index Slides
def build_Index_Slide(pres, slide_register, indexSlideIndex):
    indexSlide = pres.slides.add_slide(slide_register)
    indexSlideIndex.append(indexSlide)
    indexTitleTF = add_Text_Frame(indexSlide, indexHeading_left, indexHeading_top, indexHeading_width, indexHeading_height, indexHeading_VAlign)
    add_Heading(indexTitleTF, indexHeading_fontName, indexHeading_fontSize, indexHeading_lineSpacing, indexHeading_HAlign, indexHeading_text)
    add_hyper_jy(pres, indexSlide, 0)

#builds the SongTitle Slides
def build_Song_Title_Slide(pres, slide_register, hyperIndex, presIndex):
    titleSlide = pres.slides.add_slide(slide_register)
    songTitleTF = add_Text_Frame(titleSlide, songTitles_left, songTitles_top, songTitles_width, songTitles_height, songTitles_VAlign)
    add_Heading(songTitleTF, songTitles_fontName, songTitles_fontSize, songTitles_lineSpacing, songTitles_HAlign, songName)    
    add_hyper_jy(pres, titleSlide, hyperIndex)
    presIndex.append(pres.slides.index(titleSlide))

#builds the SongVerse Slides
def build_Song_Verse_Slide(pres, slide_register, hyperIndex):
    verseSlide = pres.slides.add_slide(slide_register)  
    verseTf = add_Text_Frame(verseSlide, verse_left, verse_top, verse_width, verse_height, verse_VAlign)
    addPara(verseTf, verse)
    add_hyper_jy(pres, verseSlide, hyperIndex)

#os paths
jy_icon_path = os.path.join("extra", "JY-Icon-White.png")
database_path = os.path.join("extra", "database.csv")
presentation_path = os.path.join("extra", "MusicSlidesTemplate.pptx")

#TitleSlide title constants
title_fontName = 'Calibri Light (Headings)'
title_fontSize = 72
title_lineSpacing = 0.9
title_HAlign = PP_PARAGRAPH_ALIGNMENT.CENTER
title_VAlign = MSO_VERTICAL_ANCHOR.BOTTOM
title_left = Cm(4.23)
title_top = Cm(3.12)
title_width = Cm(25.4)
title_height = Cm(6.63)
title_limit = 25

#TitleSlide JY Icon constants
title_JYIcon_left = Cm(15.61)
title_JYIcon_top = Cm(11.96)
title_JYIcon_height = Cm(2.64)

#TitleSlide subtitle constants
subTitle_fontName = 'Calibri (Body)'
subTitle_fontSize = 24
subTitle_lineSpacing = 0.9
subTitle_HAlign = PP_PARAGRAPH_ALIGNMENT.CENTER
subTitle_VAlign = MSO_VERTICAL_ANCHOR.MIDDLE
subTitle_left = Cm(4.23)
subTitle_top = Cm(14.97)
subTitle_width = Cm(25.4)
subTitle_height = Cm(1.25)
subTitle_text = "Jesus Youth Australia"

#normalSlide JY Icon constants
norm_JYIcon_left = Cm(16.23)
norm_JYIcon_top = Cm(17.37)
norm_JYIcon_height = Cm(1.4)

#indexSlide title constants
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

#indexSlide songName constants
indexBoxLeft = [Cm(2.48), Cm(17.08)]
indexBoxTop = [Cm(3.55), Cm(4.77), Cm(6), Cm(7.23), Cm(8.45), Cm(9.68), Cm(10.9), Cm(12.13), Cm(13.36), Cm(14.58)]
indexBoxWidth = Cm(14.3)
indexBoxHeight = Cm(1.04)

#songNameSlide constants
songTitles_fontName = 'Calibri Light (Headings)'
songTitles_fontSize = 80
songTitles_lineSpacing = 0.9
songTitles_HAlign = PP_PARAGRAPH_ALIGNMENT.CENTER
songTitles_VAlign = MSO_VERTICAL_ANCHOR.MIDDLE
songTitles_left = Cm(2.27)
songTitles_top = Cm(4.48)
songTitles_width = Cm(29.21)
songTitles_height = Cm(9.68)
songTitles_limit = 25

#songVerseSlide constants
verse_fontName = 'Calibri (Body)'
verse_HAlign = PP_PARAGRAPH_ALIGNMENT.CENTER
verse_VAlign = MSO_VERTICAL_ANCHOR.MIDDLE
verse_left = Cm(2.42)
verse_top = Cm(2)
verse_width = Cm(29.01)
verse_height = Cm(14.5)
#fontSize based on line count
verse_fontSize_lines5 = 60
verse_fontSize_lines6 = 56
verse_fontSize_lines7 = 51
#characterLimit based on line count
verse_charLimit_lines5 = 30
verse_charLimit_lines6 = 33
verse_charLimit_lines7 = 38

#pull from database
df = pd.read_csv(database_path)
try: 
    url = f'https://docs.google.com/spreadsheets/d/1P3Qu1EQLgcQYWSZQwjY5OWmEnnJMvSSgLkasa6rMC6E/gviz/tq?tqx=out:csv'
    df = pd.read_csv(url)
except:
    df = pd.read_csv(os.path.join("extra", "database.csv"))

#variables
allSongs = clean_Database(df)
fullSongList = df['Song'].tolist()
indexSlideIndex = []
presIndex = []

#PopUp Window
window = Tk()
window.title("JY Music Slides Generator")
Application.getWindow(window)
Application.saveSongList(fullSongList)
app = Application(master=window)
app.mainloop()

#save user inputs
chosenSongs = Application.song_List
savePath = Application.filepath
presentationName = process_Limit(Application.presentationName, title_limit)

#calculate number of Index pages required
indexPages = math.ceil(len(chosenSongs)/20)

#BUILD PPT
pr1 = Presentation(presentation_path)
slide1_register = pr1.slide_layouts[6]

#Title Slide
build_Title_Slide(pr1, slide1_register)

#Index Slide
for x in range(indexPages):
    build_Index_Slide(pr1, slide1_register, indexSlideIndex)

for song in chosenSongs:
    hyperIndex = math.floor(chosenSongs.index(song)/20) + 1
    songName = process_Limit(song, songTitles_limit)
    songIndex = fullSongList.index(song)
    lyrics = allSongs[songIndex][1:]

    build_Song_Title_Slide(pr1, slide1_register, hyperIndex, presIndex)

    for verse in lyrics:
        build_Song_Verse_Slide(pr1, slide1_register, hyperIndex)

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

