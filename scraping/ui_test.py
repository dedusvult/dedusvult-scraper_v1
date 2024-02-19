from tkinter import *
import webbrowser

root = Tk()
link = Label(root, text="http://stackoverflow.com", fg="blue", cursor="hand2")
link.pack()
link.bind("<Button-1>", lambda event: webbrowser.open(link.cget("text")))
root.mainloop()