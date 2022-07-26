
"""
Tkinter GUI Module
Created on Tue Jul 26 20:05:39 2022

@author: Yew Choon Min
"""

import tkinter as tk

def show_window():
    """
        Create a window to prompt user for required infos

        Returns
        -------
        None.

    """
    window = tk.Tk()
    window.geometry("420x200")
    window.title("ERP Files Generator")
    window.config(bg="#c1ccde")
    #Defining style of window
    
    label_partnum = tk.Label(text = "Part Number eg: GI130-1234", bg="#c1ccde")
    entry_partnum = tk.Entry(width = 50)
    label_partdesc = tk.Label(text = "Part Description", bg="#c1ccde")
    entry_partdesc = tk.Entry(width = 50)
    label_ECO = tk.Label(text = "ECO Number (12 digits)", bg="#c1ccde")
    entry_ECO= tk.Entry(width = 50)
    finish_button = tk.Button(text="Generate Files", width=25, height=5,bg="white")
    #Display description textbox to prompt for user input
    
    label_partnum.pack()
    entry_partnum.pack()
    label_partdesc.pack()
    entry_partdesc.pack()
    label_ECO.pack()
    entry_ECO.pack()
    finish_button.pack()
    
    partnum = entry_partnum.get()
    partdesc = entry_partdesc.get()
    ECO = entry_ECO.get()
    #Register part num, part description, ECO Number into variable
    
    window.mainloop() #Place window on computer screen
    
    return partnum, partdesc, ECO

show_window()