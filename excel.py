import pandas as pd
from JYPop import Application
from tkinter import *

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
list = Application.song_List
savePath = Application.filepath

indexes = []
for x in list:
    indexes.append(fullSongList.index(x))





#index = songList.index("Above All")
#print(songs[index][1:])