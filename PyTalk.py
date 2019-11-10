# pyinstaller --onefile --windowed filename.py

import tkinter as tk
from tkinter.filedialog import askopenfile 
from random             import randint
from time               import sleep
from win32com.client    import Dispatch
  
root = tk.Tk() 
root.geometry('830x380')
root.resizable(0, 0)
root.title("Words 2 Sounds for RHDHVers - HCM/VN - v0.3") 

speakout = Dispatch("SAPI.SpVoice") 
wordlist = []

fontsize = 30
fonttype = 'Courier'
wordread = ''
tword = 0
wword = 0

btn_open = tk.Button(root, text = " Open  ", font=(fonttype, fontsize), command = lambda: open_file()).grid(row = 0, column = 0, padx = 10, pady = 10)
btn_play = tk.Button(root, text = " Play  ", font=(fonttype, fontsize), command = lambda: play_file()).grid(row = 1, column = 0, padx = 10, pady = 10)
btn_chck = tk.Button(root, text = " Chck  ", font=(fonttype, fontsize), command = lambda: chck_word()).grid(row = 2, column = 0, padx = 10, pady = 10)

tk.Label(root, text = " Number of words  ", font=(fonttype, fontsize)).grid(row = 0, column = 1)
ntime = tk.Text(root, font=(fonttype, fontsize), height = 1, width = 6)
ntime.grid(row = 0, column = 2)

tk.Label(root, text = " Delay time (sec) ", font=(fonttype, fontsize)).grid(row = 1, column = 1)
dtime = tk.Text(root, font=(fonttype, fontsize), height = 1, width = 6)
dtime.grid(row = 1, column = 2)

check = tk.Text(root, font=(fonttype, fontsize), height = 1, width = 24)
check.grid(row = 2, column = 1, columnspan = 2)


def open_file(): 
    global wordlist
    file = askopenfile(mode ='r') 
    if file is not None: 
        wordlist = []        
        fname    = file.name
        ind      = fname.rfind('/')
        fname    = fname[ind+1:]
        fname    = 'File name: ' + fname + ' ' * (20-len(fname))
        tk.Label(root, text = fname, font=(fonttype, fontsize), anchor = 'w').grid(row = 4, column = 0, columnspan = 3)
        
        for line in file:
            wordlist.append(line.strip()) 
            
            
def play_file():
    global wordread
    global tword
    
    check.delete("1.0", "end-1c")
        
    times = int(ntime.get("1.0","end-1c"))
    delay = float(dtime.get("1.0","end-1c"))
    nword = len(wordlist)-1 
    
    if (times == 1): rep = 3
    else: rep = 1
    
    for i in range(times):
        value = randint(0, nword)
        
        for j in range(rep):
            speakout.Speak(wordlist[value])
            sleep(delay)
            
    wordread = wordlist[value]
        

def chck_word():
    global wordread
    global tword
    global wword
    
    times = int(ntime.get("1.0","end-1c"))
    
    if (times == 1):
        tword += 1
    
        word_in = check.get("1.0","end-1c")
        
        if (word_in.lower() == wordread.lower()):
            expression = 'Correct!!! ' + wordread
        else:
            expression = 'Wrong!!! ' + wordread
            wword     += 1
            
        check.delete("1.0", "end-1c"  )
        check.insert("1.0", expression)
        
        expression = " Correct percentage = " + "{:.1f}".format((1 - wword/tword)*100) + "% "
        tk.Label(root, text = expression, font=(fonttype, fontsize), anchor = 'w').grid(row = 5, column = 0, columnspan = 3)
    else:
        check.delete("1.0", "end-1c"  )
        check.insert("1.0", " Not in check mode.")
  
    
root.mainloop() 