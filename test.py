from tkinter import *
from tkinter.filedialog import asksaveasfilename
from pptx import *

def savePowerPoint():
    filepath = asksaveasfilename(defaultextension=".pptx", filetypes=[("ppt files", "*.pptx"), ("All Files", "*.*")])
    if not filepath:
        return
    saveFileName(filepath)

def saveFileName(fileName):
    prs.save(fileName)

window = Tk()

prs = Presentation("MusicSlidesTemplate.pptx")

saveBtn = Button(window, text="save", command=savePowerPoint)
saveBtn.pack()

fileName = "Slides.pptx"

window.mainloop()


