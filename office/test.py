from tkinter import *

class Application(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.pack()

if __name__=='__main__':
    app = Application()
    # to do
    app.mainloop()