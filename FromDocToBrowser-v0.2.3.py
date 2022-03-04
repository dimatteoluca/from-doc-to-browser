import textract, webbrowser, keyboard, time, threading
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from PIL import Image, ImageTk

webbrowser.register("Chrome", None, webbrowser.BackgroundBrowser("C:\Program Files\Google\Chrome\Application\chrome.exe"))
webbrowser.register("Edge", None, webbrowser.BackgroundBrowser("C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"))
webbrowser.register("Firefox", None, webbrowser.BackgroundBrowser("C:\Program Files\Mozilla Firefox\Firefox.exe"))

file = ""
browser = "Edge"
webpage = "https://clickdimensions.com/form/default.html"

start_ctrl = False
viola_ctrl = False

celeste = "#0ee0ed"
azzurro = "#0ca3ea"
blu = "#0f388e"
rosa = "#d372a9" #"#bd74ad"
viola = "#7c6bac"
grigio = "#e0e0e0"
grigio_scuro = "#858585"

def browse():

    global file
    global start_ctrl
    global viola_ctrl

    file = filedialog.askopenfilename()
    file = str(file)
    file = file.replace("/","\\")
    #print(file)

    start_ctrl = True
    start_btn["state"] = NORMAL
    start_btn["bg"] = viola
    start_btn["cursor"] = "hand2"
    viola_ctrl = True

def real_start():

    start_btn["state"] = DISABLED
    start_btn["bg"] = grigio
    start_btn["cursor"] = "arrow"

    global file
    global viola_ctrl

    try:
        #print("----------START(1)----------")

        if file.lower().endswith(".docx") or file.lower().endswith(".odt"):

            text = textract.process(r'%s' %file)
            #print("----------START(2)----------")
            text = text.decode('utf-8')
            text_lines = text.split('\n')

            # Remove all the empty lines
            try:
                while True:
                    text_lines.remove('')
            except ValueError:
                pass

            #for line in range(len(text_lines)):    
            #    print(text_lines[line])

            webbrowser.get(browser).open_new_tab(webpage)
            for line in range(len(text_lines)):
                time.sleep(0.75)
                keyboard.press_and_release('tab')
                keyboard.write(text_lines[line])

            #print("----------START(end)----------")

        else:
            if file == "":
                print("ATTENZIONE: nessun file selezionato")
                messagebox.showerror(title="Attenzione", message="Nessun file selezionato.")
            else:
                print("ATTENZIONE: formato del file non supportato")
                messagebox.showerror(title="Attenzione", message="Il formato del file selezionato non è supportato.")

    except textract.exceptions.MissingFileError:
        print("ATTENZIONE: file non trovato")
        messagebox.showerror(title="Attenzione", message="File non trovato.")

    except KeyError:
        print("ATTENZIONE: il formato del file non è corretto\nSalvare nuovamente il file come .DOCX o .ODT tramite il comando 'Salva con nome...'")
        messagebox.showerror(title="Attenzione", message="Il formato del file non è corretto.\nSalvare nuovamente il file come .DOCX o .ODT tramite il comando 'Salva con nome...'.")

    except textract.exceptions.ShellError:
        print("ATTENZIONE: selezionare un file .DOCX o .ODT")
        messagebox.showerror(title="Attenzione", message="Selezionare un file .DOCX o .ODT.")

    start_btn["state"] = NORMAL
    start_btn["bg"] = grigio_scuro
    start_btn["cursor"] = "hand2"
    viola_ctrl = False


def start():

    thread = threading.Thread(target=real_start, daemon=True)
    thread.start()

def ok(newWindow, br_clicked, wp_entry):

    global browser
    global webpage

    browser = br_clicked.get()
    if (wp_entry != ""):
        try:
            webpage = str(wp_entry.get())
        except:
            print("---------ok() exception---------")

    newWindow.destroy()

    browse_btn["state"] = NORMAL
    browse_btn["bg"] = rosa
    browse_btn["cursor"] = "hand2"

    if start_ctrl==True:
        start_btn["state"] = NORMAL
        start_btn["bg"] = viola
        start_btn["cursor"] = "hand2"
        if viola_ctrl==True:
            start_btn["bg"] = viola
        else:
            start_btn["bg"] = grigio_scuro

    settings_btn["state"] = NORMAL
    settings_btn["bg"] = azzurro
    settings_btn["cursor"] = "hand2"

def myDestroy(newWindow):

    newWindow.destroy()

    browse_btn["state"] = NORMAL
    browse_btn["bg"] = rosa
    browse_btn["cursor"] = "hand2"

    if start_ctrl==True:
        start_btn["state"] = NORMAL
        start_btn["cursor"] = "hand2"
        if viola_ctrl==True:
            start_btn["bg"] = viola
        else:
            start_btn["bg"] = grigio_scuro

    settings_btn["state"] = NORMAL
    settings_btn["bg"] = azzurro
    settings_btn["cursor"] = "hand2"

def disable_event():
    pass

def settings():

    #Buttons reset
    browse_btn["state"] = DISABLED
    browse_btn["bg"] = grigio
    browse_btn["cursor"] = "arrow"

    start_btn["state"] = DISABLED
    start_btn["bg"] = grigio
    start_btn["cursor"] = "arrow"

    settings_btn["state"] = DISABLED
    settings_btn["bg"] = grigio
    settings_btn["cursor"] = "arrow"

    #New window
    newWindow = Toplevel(root)
    newWindow.attributes('-topmost', True) 
    newWindow.resizable(0, 0)
    newWindow.iconbitmap("./fdtb.ico")
    newWindow.title("Settings")
    newWindow.protocol("WM_DELETE_WINDOW", disable_event)
    newWindow.geometry("+1325+430")

    nw_canvas = Canvas(newWindow, height=190, width=350)
    nw_canvas.pack()
    nw_canvas.grid(rowspan=5, columnspan=1)

    #browser
    browser_label = Label(newWindow, text="Browser:")
    browser_label.grid(row=0, column=0)#, sticky=E)
    br_options = ["Chrome", "Edge", "Firefox"]
    br_clicked = StringVar()
    if browser=="Chrome":
        br_clicked.set(br_options[0])
    elif browser=="Edge":
        br_clicked.set(br_options[1])
    else:
        br_clicked.set(br_options[2])
    br_drop = OptionMenu(newWindow, br_clicked, *br_options)
    br_drop.grid(row=1, column=0)
    #webpage url
    webpage_label = Label(newWindow, text="Webpage URL:")
    webpage_label.grid(row=2, column=0, pady=(15,0))#, columnspan=2)#, sticky=E)
    wp_entry = Entry(newWindow, width=55, justify="center")
    wp_entry.insert(0, webpage)
    wp_entry.grid(row=3, column=0)#, columnspan=2)
    #parent widget for the buttons
    buttonsFrame = Frame(newWindow)
    buttonsFrame.grid(row=4, column=0, rowspan=1, pady=(25,5))#, columnspan=2)
        #ok
    ok_text = StringVar()
    ok_btn = Button(buttonsFrame, textvariable=ok_text, height=1, width=10, command=lambda:ok(newWindow, br_clicked, wp_entry), cursor="hand2")
    ok_text.set("Ok")
    ok_btn.grid(row=0, column=0)
        #cancel
    cancel_text = StringVar()
    cancel_btn = Button(buttonsFrame, textvariable=cancel_text, height=1, width=10, command=lambda:myDestroy(newWindow), cursor="hand2")
    cancel_text.set("Cancel")
    cancel_btn.grid(row=0, column=1)


root = Tk()

root.title("FromDocToBrowser")                  # window title
root.iconbitmap("./fdtb.ico")                   # window icon
root.attributes('-topmost', True)               # the window is always on top
root.resizable(0, 0)                            # the window is not resizable
root.geometry("+1300+400")

canvas = Canvas(root, width=400)#, height=250)    # needed for the layout
#canvas.pack()                                  # canvas size depends on the size of the elements in it
canvas.grid(rowspan=4, columnspan=1)            # grid layout (2x1)

logo = Image.open("fdtb.png").resize([200,200])
logo = ImageTk.PhotoImage(logo)             # converts the Pillow Image into a Tkinter Image
logo_label = Label(image=logo)       
logo_label.image = logo
logo_label.grid(row=0, column=0, pady=(30,0))

btns_frame = Frame(root)
btns_frame.grid(row=1, column=0, columnspan=2, rowspan=2, pady=(30,40))

browse_text = StringVar()
browse_btn = Button(btns_frame, textvariable=browse_text, font=("Raleway", "11","bold"), bg=rosa, fg="white", height=2, width=12, command=lambda:browse(), cursor="hand2")
browse_text.set("Browse")
browse_btn.grid(row=0, column=0)

start_text = StringVar()
start_btn = Button(btns_frame, textvariable=start_text, font=("Raleway", "11","bold"), bg=grigio, fg="white", height=2, width=12, command=lambda:start(), state=DISABLED)
start_text.set("Start")
start_btn.grid(row=0, column=1)

settings_text = StringVar()
settings_btn = Button(btns_frame, textvariable=settings_text, font=("Raleway", "11","bold"), bg=azzurro, fg="white", height=2, width=12, command=lambda:settings(), cursor="hand2")
settings_text.set("Settings")
settings_btn.grid(row=1, column=0)

quit_text = StringVar()
quit_btn = Button(btns_frame, textvariable=quit_text, font=("Raleway", "11","bold"), bg=blu, fg="white", height=2, width=12, command=root.quit, cursor="hand2")
quit_text.set("Quit")
quit_btn.grid(row=1, column=1)

info_label = Label(root, text="Luca Di Matteo", fg="grey")
info_label.grid(row=3, column=0)

root.mainloop()

# .DOCX & .ODT
# funzionano senza problemi

# fake.DOCX
# KeyError: "There is no item named 'word/document.xml' in the archive"
# --> chiedere di salvare nuovamente il file in modo corretto ('Salva con nome...' .docx .odt)

# .DOC & fake.DOC
# textract.exceptions.ShellError: The command `antiword C:\Users\lucad\Downloads\FromDocToBrowser\test-doc.doc` failed with exit code 127
# --> installare 'antiword' tramite pip
# --> textract.exceptions.ShellError: The command `antiword C:\Users\lucad\Downloads\FromDocToBrowser\test-fintodoc.doc` failed with exit code 1
# --> installare pywin32 tramite pip e modificare il path con la funzione GetShortPathName
# --> FileNotFoundError: [WinError 2] Impossibile trovare il file specificato
# --> RIP