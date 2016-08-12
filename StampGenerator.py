import tkinter as tk
from tkinter import *
import StampGeneratorGUI
import transmittalgeneratorGUI
from PIL import Image, ImageTk


entryArray=[]
inputArray=[]
    
def addEntry():
     
        nextRow= len(entryArray) 
        
        if len(entryArray) >=9:
                return
        else: 
            if nextRow >= 4:
                newRow = len(entryArray)%4 + 1    
                nextEntry =  Entry(entryContainer, width=5)
                nextLabel = Label(entryContainer, text='SD#')
                nextLabel.grid(row=newRow, column=3)
                nextEntry.grid(row=newRow, column=4, padx=5, pady=5)
                entryArray.append(nextEntry)          
                
def entryToArray():
    for entry in range(len(entryArray)):
        val = str(entryArray[entry].get()).lower()
        inputArray.append(val)

master = tk.Tk()
master.wm_iconbitmap(r'./Templates/ICON.ico')
master.title('Shop Drawing Stamp and Transmittal Generator, by William Bartos')
master.geometry('600x480')
master.resizable(width=FALSE, height=FALSE)

backgroundImg =Image.open(r'./Templates/ICON.gif')
photo = ImageTk.PhotoImage(backgroundImg)
label=Label(image=photo)
label.image= photo
label.place(in_=master)

ml1=Label(master, text= 'Shop Drawing Stamp and Transmittal Generator' + '\n' + 'v1.1')
ml1.pack(padx = 10, pady=40)

entryContainer = Frame(height = 110, width = 150)
entryContainer.pack(padx=100, pady=0, fill=BOTH)
entryContainer.grid_rowconfigure(0, weight=1)
entryContainer.grid_rowconfigure(4, weight=1)
entryContainer.grid_columnconfigure(0, weight=1)
entryContainer.grid_columnconfigure(6, weight=1)

el1=Label(entryContainer, text= 'SD#')
el2=Label(entryContainer, text= 'SD#')
el3=Label(entryContainer, text= 'SD#')
el4=Label(entryContainer, text= 'SD#')

el1.grid(row=1, column=1)
el2.grid(row=2, column=1)
el3.grid(row=3, column=1)
el4.grid(row=4, column=1)

e1 = Entry(entryContainer, width=5)
e2 = Entry(entryContainer, width=5)
e3 = Entry(entryContainer, width=5)
e4 = Entry(entryContainer, width=5)

e1.grid(row=1, column=2, padx=5, pady=5)
e2.grid(row=2, column=2, padx=5, pady=5)
e3.grid(row=3, column=2, padx=5, pady=5)
e4.grid(row=4, column=2, padx=5, pady=5)

entryArray.extend((e1,e2,e3, e4))

entryButtonContainer=Frame(height=50, width=200)
entryButtonContainer.pack(padx=200, pady=10, fill=BOTH)

bl1= Button(entryButtonContainer, text='Add More Entries', command=addEntry)
bl2= Button(entryButtonContainer, text='Apply', command=entryToArray)
bl1.pack()
bl2.pack(pady=10)

buttonContainer = Frame(height = 100, width = 150)
buttonContainer.pack(side=BOTTOM, padx=10, pady=0, fill=BOTH)

b1 = tk.Button(buttonContainer, text='Generate Stamps', command= lambda: StampGeneratorGUI.stampWriter(inputArray),height = 2, width = 3, padx=50)
b2 = tk.Button(buttonContainer, text='Generate Transmittal', command= lambda: transmittalgeneratorGUI.transmittalWriter(inputArray), height = 2, width = 3, padx=50)

b1.place(in_=buttonContainer, relx=.25, rely=.5, anchor='center')
b2.place(in_=buttonContainer, relx=.75, rely=.5, anchor='center')

tk.mainloop()







