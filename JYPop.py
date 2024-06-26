from tkinter.filedialog import asksaveasfilename
from tkinter import *
import os.path

class Application(Frame):
    full_Song_List = []
    song_List = []
    list2 = []
    window = 1
    save = 0
    filepath = ""
    presentationName = ""
    jy_image_path = os.path.join("extra", "JY-Icon-White.png")
    mode = "Classic"
    onlineMode = 0
    onlineText = "Connected to Online Songbank"
    offlineText = "Offline - Using Local Songbank"

    #Init
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.createWindow()    

    #Create main window
    def getWindow(window):
        Application.window = window

    def createWindow(self):
        # create a Tk root window

        w = 550 # width for the Tk window
        h = 570 # height for the Tk window

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
        frame_middle_top = Frame(Application.window, bd=2)
        frame_middle = Frame(Application.window, bd=2)
        frame_bottom = Frame(Application.window, bd=2)

        #img = ImageTk.PhotoImage(Image.open(Application.jy_image_path))

        #left column
        self.search_var = StringVar()
        self.search_var.trace("w", self.updateList)

        frame_left_top = Frame(frame_left, bd=2)

        self.entryLabel = Label(frame_left_top, text="Filter:")
        self.entry = Entry(frame_left_top, textvariable=self.search_var, width=13)
        self.songListBoxFrom = Listbox(frame_left, width=25, height=15, selectmode = "multiple")
        self.addSongBtn = Button(frame_left, text="Add Song", command=self.addSong)
        self.addAllBtn = Button(frame_left, text="Add All", command=self.addAll)

        self.entryLabel.grid(row=0, column=0, padx=10, pady=3)
        self.entry.grid(row=0, column=1, padx=10, pady=3)
        self.songListBoxFrom.grid(row=1, column=0, padx=10, pady=3)
        self.addSongBtn.grid(row=2, column=0, padx=10, pady=3)
        self.addAllBtn.grid(row=3, column=0, padx=10, pady=3)

        #right column
        self.search_var2 = StringVar()
        self.search_var2.trace("w", self.updateList2)

        frame_right_top = Frame(frame_right, bd=2)

        self.entryLabel2 = Label(frame_right_top, text="Filter:")
        self.entry2 = Entry(frame_right_top, textvariable=self.search_var2, width=13)
        self.songListBoxTo = Listbox(frame_right, width=25, height=15, selectmode = "single")
        self.removeSongBtn = Button(frame_right, text="Remove Song", command=self.removeSong)
        self.removeBtn = Button(frame_right, text="Remove All", command=self.removeList)

        self.entryLabel2.grid(row=0, column=0, padx=10, pady=3)
        self.entry2.grid(row=0, column=1, padx=10, pady=3)
        self.songListBoxTo.grid(row=1, column=0, padx=10, pady=3)
        self.removeSongBtn.grid(row=2, column=0, padx=10, pady=3)
        self.removeBtn.grid(row=3, column=0, padx=10, pady=3)

        
        modes = [(1,"Black"), (2,"Dark"), (3,"Classic"), (4,"Light"), (5,"White")]
        
        #middle row
        if self.onlineMode == 0:
            onlineStatusText = self.offlineText
            colour = "red"
        elif self.onlineMode == 1:
            onlineStatusText = self.onlineText
            colour = "green"

        self.modeLbl = Label(frame_middle_top, text=onlineStatusText, fg=colour)
        self.modeLbl.grid(row=0, column=0, padx=10, pady=3)

        #middle bottom row
        self.v = IntVar()
        self.v.set(3)

        self.modeLbl = Label(frame_middle, text="Mode: ")
        self.blackButton = Radiobutton(frame_middle, text="Black", variable=self.v, command=self.saveMode, value=1)
        self.darkButton = Radiobutton(frame_middle, text="Dark", variable=self.v, command=self.saveMode, value=2)
        self.oldButton = Radiobutton(frame_middle, text="Classic", variable=self.v, command=self.saveMode, value=3)
        self.lightButton = Radiobutton(frame_middle, text="Light", variable=self.v, command=self.saveMode, value=4)
        self.whiteButton = Radiobutton(frame_middle, text="White", variable=self.v, command=self.saveMode, value=5)
        
        self.modeLbl.grid(row=0, column=0, padx=10, pady=3)
        self.blackButton.grid(row=0, column=1, padx=10, pady=3)
        self.darkButton.grid(row=0, column=2, padx=10, pady=3)
        self.oldButton.grid(row=0, column=3, padx=10, pady=3)
        self.lightButton.grid(row=0, column=4, padx=10, pady=3)
        self.whiteButton.grid(row=0, column=5, padx=10, pady=3)

        #bottom row
        self.saveLbl = Label(frame_bottom, text="Title: ")
        self.saveEntry = Entry(frame_bottom, width=13)
        self.saveEntry.insert(END, 'JY Music')
        self.saveBtn = Button(frame_bottom, text="Save Slides", command=self.getSongList)

        self.saveLbl.grid(row=1, column=1, padx=10, pady=3)
        self.saveEntry.grid(row=1, column=2, padx=10, pady=3)
        self.saveBtn.grid(row=1, column=3, padx=10, pady=3)

        #frames
        frame_top.grid(row=0, sticky="n")
        frame_left_top.grid(row=0, column=0, sticky="w")
        frame_right_top.grid(row=0, column=0, sticky="w")
        frame_left.grid(row=0, column=0, sticky="n")
        frame_right.grid(row=0, column=1, sticky="n")
        frame_middle_top.grid(row=1, sticky="s")
        frame_middle.grid(row=2, sticky="s")
        frame_bottom.grid(row=3, sticky="s")

        self.updateList()
        self.updateList2()

    def saveMode(self):
        if self.v.get() == 1:
            Application.mode = "Black"
        elif self.v.get() == 2:
            Application.mode = "Dark"
        elif self.v.get() == 3:
            Application.mode = "Classic"
        elif self.v.get() == 4:
            Application.mode = "Light"
        elif self.v.get() == 5:
            Application.mode = "White"

    def onlineStatus(online):
        Application.onlineMode = online

    def updateList(self, *args):
        search_term = self.search_var.get()

        # Just a generic list to populate the listbox

        self.songListBoxFrom.delete(0, END)

        for item in Application.full_Song_List:
                if search_term.lower() in item.lower():
                    self.songListBoxFrom.insert(END, item)
    
    def updateList2(self, *args):
        search_term2 = self.search_var2.get()

        # Just a generic list to populate the listbox
        for item in self.songListBoxTo.get(0, END):
            if item in Application.list2:
                pass
            else:
                Application.list2.append(item)

        self.songListBoxTo.delete(0, END)

        for item in Application.list2:
                if search_term2.lower() in item.lower():
                    self.songListBoxTo.insert(END, item)

    def addSong(self):
        for x in self.songListBoxFrom.curselection():
            self.songListBoxTo.insert(END, self.songListBoxFrom.get(x))
        self.songListBoxFrom.selection_clear(0, END)
        #self.songListBoxTo.insert(END, self.songListBoxFrom.get(ANCHOR))

    def addAll(self):
        self.removeList()
        for item in Application.full_Song_List:
            self.songListBoxTo.insert(END, item)

    def removeSong(self):
        self.songListBoxTo.delete(ANCHOR)

    def removeList(self):
        self.songListBoxTo.delete(0, END)

    def saveSongList(songList):
        Application.full_Song_List = songList

    def getSongList(self):
        for item in self.songListBoxTo.get(0, END):
            Application.song_List.append(item)

        Application.saveFile()
        if Application.save == 1:
            Application.presentationName = self.saveEntry.get()
            Application.window.destroy()

    def saveFile():
        filepath = asksaveasfilename(defaultextension=".pptx", filetypes=[("ppt Files", "*.pptx"), ("All Files", "*.*")])
        if not filepath:
            return
        Application.save = 1
        Application.filepath = filepath



