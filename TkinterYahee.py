from tkinter import filedialog
import tkinter as tk
import yaheeFunctions as m
import pandas as pd

hauptFenster = tk.Tk()
hauptFenster.geometry("800x500")
hauptFenster.winfo_toplevel().title('YAHEE T1')
instructionLabel = tk.Label(hauptFenster, text="Geben sie die Daten ein!")


def changeDirectory():
    hauptFenster.directory = filedialog.askdirectory()  # Instantly runs the filedialog
    print("gedrückt")
    print(hauptFenster.directory)


def deleteInput():
    Sdus.delete(0, 'end')
    Schiff.delete(0, 'end')
    BL.delete(0, 'end')
    BLDATUM.delete(0, 'end')
    Incoterm.delete(0, 'end')
    Transportpreis.delete(0, 'end')


def createFile(Sdus, Schiff, BL, BLDATUM,Incoterm, Transportpreis,
               Inlandspreis, Packliste):
    Sendungsnr = Sdus.get()
    Schiff1 = Schiff.get()
    BL1 = BL.get()
    BLDatum = BLDATUM.get()
    Incoterm1 = "FOB " + Incoterm.get()
    Transportpreis1 = Transportpreis.get() + " EUR"
    Inlandpreis1 = str(Inlandspreis) + " EUR"
    writer = pd.ExcelWriter(hauptFenster.directory + '/' + Sendungsnr + ".xlsx", engine='xlsxwriter')  # IMPORTANT!!!

    T1, verzollung, Rechnungsnr, Rechnungsdatum,Containernr = m.createT1(Packliste)# add rechnungsnummer Containernr. and rechnungsdatum as return
    m.createWorkbook(T1, verzollung, Sendungsnr, Schiff1, BL1, BLDatum, Rechnungsnr, Rechnungsdatum, Containernr,
                     Incoterm1, Transportpreis1, Inlandpreis1, writer)
    print('done')
    print("Created: " + Sendungsnr + ".xlsx")
    deleteInput()
    return


# per button auswahl
def open_file_dialog():
    global file_path
    file_path = filedialog.askopenfilename(initialdir='D://file')
    print("Ausgewählte Datei:", file_path)


# per dnd auswahl
# def drop(event):
# listb.insert("end",event.data)


# listb = tk.Listbox(hauptFenster, selectmode=tk.SINGLE)
# listb.drop_target_register(dnd.DND_FILES)
# listb.bind("<<Drop>>",drop)
# listb.grid(row=11, column=1)


Sdus = tk.Entry(hauptFenster)
Sdus.grid(row=0, column=1)
tk.Label(hauptFenster, text='SDUS').grid(row=0)

Schiff = tk.Entry(hauptFenster)
Schiff.grid(row=1, column=1)
tk.Label(hauptFenster, text='Schiff').grid(row=1)

BL = tk.Entry(hauptFenster)
BL.grid(row=2, column=1)
tk.Label(hauptFenster, text='BL').grid(row=2)

BLDATUM = tk.Entry(hauptFenster)
BLDATUM.grid(row=3, column=1)
tk.Label(hauptFenster, text='BLDatum').grid(row=3)



Incoterm = tk.Entry(hauptFenster)
Incoterm.grid(row=7, column=1)
tk.Label(hauptFenster, text='Incoterm').grid(row=7)

Transportpreis = tk.Entry(hauptFenster)
Transportpreis.grid(row=8, column=1)
tk.Label(hauptFenster, text='Transportpreis').grid(row=8)

Inlandspreis = 795

SdusButton = tk.Button(text='Create',
                       command=lambda: createFile(Sdus, Schiff, BL, BLDATUM,
                                                  Incoterm, Transportpreis, Inlandspreis, file_path))
SdusButton.grid(row=12, column=1)

PacklistWahlButton = tk.Button(text='wähle Packliste',
                               command=lambda: open_file_dialog())
SpeicherOrt = tk.Button(text='Wähle Zielordner', command=lambda: changeDirectory())
SpeicherOrt.grid(row=17, column=1)
PacklistWahlButton.grid(row=10, column=1)
hauptFenster.update()
hauptFenster.mainloop()
