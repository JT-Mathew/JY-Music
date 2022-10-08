from pptx import *
from tkinter import *
from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT
from pptx.util import Cm, Pt

def addPara(textBoxText, para):
    splitPara = para.splitlines()
    
    fontSize = Lines5
    if len(splitPara) == 6:
        fontSize = Lines6
    elif len(splitPara) > 6:
        fontSize = Lines7
        
    count = 0
    for line in splitPara:
        addLine(textBoxText, line, count, fontSize)
        count = 1

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

Lines5 = 60
Lines6 = 56
Lines7 = 51

left = Cm(2.42)
top = Cm(2)
width = Cm(29.01)
height = Cm(14.5)

song = []
song.append("""Bless the Lord
Oh my soul, oh my soul
Worship his holy name
Sing like never before
Oh my soul
I’ll worship your holy name""")

song.append("""The sun comes up, it’s a new day dawning
It’s time to sing your song again
Whatever may pass and whatever lies before me
Let me be singing when the evening comes
""")

pr1 = Presentation("MusicSlidesTemplate.pptx")

slide1_register = pr1.slide_layouts[6]

for verse in song:
    slide = pr1.slides.add_slide(slide1_register)
    textBox = slide.shapes.add_textbox(left, top, width, height)
    textBoxText = textBox.text_frame
    addPara(textBoxText, verse)

pr1.save('testPower.pptx')

