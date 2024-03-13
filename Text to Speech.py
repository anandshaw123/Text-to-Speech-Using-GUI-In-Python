from tkinter import *
from PIL import ImageTk,Image
from tkinter import filedialog
from tkinter.ttk import Combobox
import pyttsx3
import os
import win32com.client

obj = Tk()






obj.title('Text To Speech')
obj.geometry('1000x590+200+200')
obj.resizable(False,False) # You Did not Resize the window
# icon add
obj.iconbitmap('adobe_Ai.ico')
obj.configure(bg='yellow')



engine = pyttsx3.init()

def speaknow():
    text=text_area.get(1.0, END)
    gender = gender_combobox.get()
    speed = speed_combobox.get()
    voices = engine.getProperty('voices')

    def setvoice(): 
        if(gender =='Male'):
            engine.setProperty('voice',voices[0].id)
            engine.say(text)
            engine.runAndWait()
        else:
            engine.setProperty('voice',voices[1].id)
            engine.say(text)
            engine.runAndWait()    

    if(text):
        if(speed == "Fast"):
            engine.setProperty('rate',250)
            setvoice()
        elif (speed == 'Normal'):
            engine.setProperty('rate',150)    
            setvoice()
        else:
            engine.setProperty('rate',60)
            setvoice()

def download():

    text=text_area.get(1.0, END)
    gender = gender_combobox.get()
    speed = speed_combobox.get()
    voices = engine.getProperty('voices')

    def setvoice(): 
        if(gender =='Male'):
            engine.setProperty('voice',voices[0].id)
            path = filedialog.askdirectory()
            os.chdir(path)
            engine.save_to_file(text, 'text.mp3')
            engine.runAndWait()
        else:
            engine.setProperty('voice',voices[1].id)
            path = filedialog.askdirectory()
            os.chdir(path)
            engine.save_to_file(text,'text.mp3')
            engine.runAndWait()    

    if(text):
        if(speed == "Fast"):
            engine.setProperty('rate',250)
            setvoice()
        elif (speed == 'Normal'):
            engine.setProperty('rate',150)    
            setvoice()
        else:
            engine.setProperty('rate',60)
            setvoice()



# Lets add White Border in page
top_frame = Frame(obj,bg="white",width=1000,height=100)
top_frame.place(x=0,y=0)

# 33444443434455545554%$%$%%%%%$%%$%%%%%%%$%%$%$%$%$%$%$%$%$%$%$%$%%$%$%$%%$


#Typing Box
text_area = Text(obj,font="Robote 20",bg="white",relief=GROOVE,wrap=WORD)
text_area.place(x=10,y=150,width=400,height=230)

#Name Added Voice and Speed
Label(obj,text="Voice",font="arial 15 bold",bg="#305065",fg="white").place(x=580,y=160) 

Label(obj,text="Speed",font="arial 15 bold",bg="#305065",fg="white").place(x=760,y=160)


#Gender Box
gender_combobox =Combobox(obj,values=['Male','Female'],font="arial 14",state='r',width=10)
gender_combobox.place(x=550,y=200)
gender_combobox.set('Male')

#Speed Box
speed_combobox =Combobox(obj,values=['Fast','Slow','Normal'],font="arial 14",state='r',width=10)
speed_combobox.place(x=730,y=200)
speed_combobox.set('Normal')


imageicon=PhotoImage(file='Oxy.png')
b1 = Button(obj,text="SPEAK",compound=LEFT,image=imageicon,fg="red",bg = "black",width=150,height=90,font='arial 14 bold',command=speaknow)
b1.place(x=650,y=290)# adjust speak button



obj.mainloop()