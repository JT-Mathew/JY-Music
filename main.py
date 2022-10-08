from pptx import Presentation 
from tkinter import *

class Application(Frame):
    
    #Init
    def __init__(self, master=None):
        Frame.__init__(self, master)

        self.createWindow()
    

    #Create main window
    def createWindow(self):
        # create a Tk root window

        w = 500 # width for the Tk window
        h = 500 # height for the Tk window

        # get screen width and height
        ws = window.winfo_screenwidth() # width of the screen, to determine positioning of window
        hs = window.winfo_screenheight() # height of the screen, to determine positioning of window

        # calculate x and y coordinates for middle of the screen
        x = (ws/2) - (w/2)
        y = (hs/2) - (h/2)

        # set the dimensions of the screen, w and h
        # and where it is placed, x and y
        window.geometry('%dx%d+%d+%d' % (w, h, x, y))
        window.rowconfigure(0, minsize=450, weight=1)
        window.columnconfigure(0, minsize=250, weight=1)
        window.resizable(False, False)

        frame_left = Frame(window, bd=2)
        frame_right = Frame(window, bd=2)
        frame_bottom = Frame(window, bd=2)

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

        frame_left_top.grid(row=0, column=0, sticky="w")
        frame_left.grid(row=0, column=0, sticky="ns")
        frame_right.grid(row=0, column=1, sticky="ns")
        frame_bottom.grid(row=1, sticky="ns")

        #bottom column

        self.updateList()

    def updateList(self, *args):
        search_term = self.search_var.get()

        # Just a generic list to populate the listbox
        songList = ["Lord I need you", "10000 reasons", "Above All", "Days Gone By", "Holy Holy", "Hallelujiah", "Dreamer", "Amazing Grace", "Broken Vessels", "Oceans", "My Lighthouse", "New wine", "Still", "Reckless Love", "So Will I", "You Say"]

        self.songListBoxFrom.delete(0, END)

        for item in songList:
                if search_term.lower() in item.lower():
                    self.songListBoxFrom.insert(END, item)

    def addSong(self):
        self.songListBoxTo.insert(END, self.songListBoxFrom.get(ANCHOR))

    def removeSong(self):
        self.songListBoxTo.delete(ANCHOR)

    def clearList(self):
        self.songListBoxTo.delete(0, END)

    def getSongList():
        songList = ["Lord I need you", "10000 reasons", "Above All", "Days Gone By", "Holy Holy", "Hallelujiah", "Dreamer", "Amazing Grace", "Broken Vessels", "Oceans", "My Lighthouse", "New wine", "Still", "Reckless Love", "So Will I", "You Say"]
        return songList

window = Tk()
window.title("JY Australia Music Slides")
app = Application(master=window)
app.mainloop()

