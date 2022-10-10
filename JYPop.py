from tkinter.filedialog import asksaveasfilename
from pptx import Presentation 
from tkinter import *

class Application(Frame):
    full_Song_List = []
    song_List = []
    window = 1
    save = 0
    filepath = ""

    #Init
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.createWindow()    

    #Create main window
    def getWindow(window):
        Application.window = window

    def createWindow(self):
        # create a Tk root window

        w = 500 # width for the Tk window
        h = 500 # height for the Tk window

        # get screen width and height
        ws = Application.window.winfo_screenwidth() # width of the screen, to determine positioning of window
        hs = Application.window.winfo_screenheight() # height of the screen, to determine positioning of window

        # calculate x and y coordinates for middle of the screen
        x = (ws/2) - (w/2)
        y = (hs/2) - (h/2)

        # set the dimensions of the screen, w and h
        # and where it is placed, x and y
        Application.window.geometry('%dx%d+%d+%d' % (w, h, x, y))
        Application.window.rowconfigure(0, minsize=450, weight=1)
        Application.window.columnconfigure(0, minsize=250, weight=1)
        Application.window.resizable(False, False)

        frame_top = Frame(Application.window, bd=2)
        frame_left = Frame(frame_top, bd=2)
        frame_right = Frame(frame_top, bd=2)
        frame_bottom = Frame(Application.window, bd=2)

        #left column
        self.search_var = StringVar()
        self.search_var.trace("w", self.updateList)

        frame_left_top = Frame(frame_left, bd=2)

        self.entryLabel = Label(frame_left_top, text="    Filter:")
        self.entry = Entry(frame_left_top, textvariable=self.search_var, width=13)
        self.songListBoxFrom = Listbox(frame_left, width=25, height=15, selectmode = "single")
        self.addSongBtn = Button(frame_left, text="Add Song", command=self.addSong)

        self.entryLabel.grid(row=0, column=0, padx=10, pady=3)
        self.entry.grid(row=0, column=1, padx=10, pady=3)
        self.songListBoxFrom.grid(row=1, column=0, padx=10, pady=3)
        self.addSongBtn.grid(row=2, column=0, padx=10, pady=3)

        #right column
        self.clearBtn = Button(frame_right, text="Clear List", command=self.clearList)
        self.songListBoxTo = Listbox(frame_right, width=25, height=15, selectmode = "single")
        self.removeSongBtn = Button(frame_right, text="Remove Song", command=self.removeSong)

        self.clearBtn.grid(row=0, column=0, padx=10, pady=8)
        self.songListBoxTo.grid(row=1, column=0, padx=10, pady=3)
        self.removeSongBtn.grid(row=2, column=0, padx=10, pady=3)

        #bottom column
        self.saveBtn = Button(frame_bottom, text="Create Slides", command=self.getSongList)

        self.saveBtn.grid(row=0, column=0, padx=10, pady=8)

        frame_top.grid(row=0, sticky="ns")
        frame_left_top.grid(row=0, column=0, sticky="w")
        frame_left.grid(row=0, column=0, sticky="ns")
        frame_right.grid(row=0, column=1, sticky="ns")
        frame_bottom.grid(row=1, sticky="ns")
        
        self.updateList()

    def updateList(self, *args):
        search_term = self.search_var.get()

        # Just a generic list to populate the listbox

        self.songListBoxFrom.delete(0, END)

        for item in Application.full_Song_List:
                if search_term.lower() in item.lower():
                    self.songListBoxFrom.insert(END, item)

    def addSong(self):
        self.songListBoxTo.insert(END, self.songListBoxFrom.get(ANCHOR))

    def removeSong(self):
        self.songListBoxTo.delete(ANCHOR)

    def clearList(self):
        self.songListBoxTo.delete(0, END)

    def saveSongList(songList):
        Application.full_Song_List = songList

    def getSongList(self):
        for item in self.songListBoxTo.get(0, END):
            Application.song_List.append(item)

        Application.saveFile()
        if Application.save == 1:
            Application.window.destroy()

    def saveFile():
        filepath = asksaveasfilename(defaultextension=".pptx", filetypes=[("ppt Files", "*.ppt"), ("All Files", "*.*")])
        if not filepath:
            return
        Application.save = 1
        Application.filepath = filepath



