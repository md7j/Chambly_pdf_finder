import urllib.request, urllib.error, urllib.parse
import re
import wget
import os
from tkinter import *
import tkinter.ttk
from tkinter import scrolledtext 
import PyPDF2
import subprocess

window=Tk()
liste = Listbox(window, font=("Helvetica", 10), bd=3, width=50, height=18)
liste.place(x=150, y=450)


def charger_docs(event):
    res = []

    charger_docs_progress = tkinter.ttk.Progressbar(window, length=500, mode="indeterminate", value=0)
    charger_docs_progress.place(x=150, y=180)
    charger_docs_btn.destroy()
    window.update()
    charger_docs_progress.start(25)
    for year in range(2007, 2008):
        year_content = []
        try:
            response = urllib.request.urlopen(url = 'http://www.ville-chambly.fr/Vie-municipale/Le-conseil-municipal/Les-seances-du-Conseil-Municipal/{}'.format(year))
            webContent = response.read().decode("utf-8")
            links = re.findall('href=\"([^"]+[.](pdf|PDF))"', webContent)
            for link in links:
                year_content.append(link[0])
                if "http://" not in year_content[-1]:
                    year_content[-1] = "http://www.ville-chambly.fr/" + year_content[-1]
                window.update()
            res.append(year_content)
        except:
            pass
    os.mkdir('{}/Documents'.format(os.getcwd()))
    for idx, year_content in enumerate(res):
        os.mkdir('{}/Documents/{}'.format(os.getcwd(), idx + 2007))
        for link in year_content:
            wget.download(link, '{}/Documents/{}/{}'.format(os.getcwd(), idx + 2007, re.findall(r'([^/]+(pdf|PDF))$', link)[0][0]))
            window.update()
    charger_docs_progress.destroy()
    fini=Label(window, text="Terminé", fg='green', font=("Helvetica", 12))
    fini.place(x=150, y=180)

def chercher(event): 
    liste.delete(0, liste.size() - 1)
    list_entries = 1
    chercher_progress = tkinter.ttk.Progressbar(window, length=360, mode="indeterminate", value=0)
    chercher_progress.place(x=150, y=750)
    window.update()
    chercher_progress.start(25)

    for search_year, year in zip(checkbuttons_search, years):
        if search_year.get() == 1:
            files = os.listdir("{}/Documents/{}/".format(os.getcwd(), year))
            for file in files:
                pdfRead = PyPDF2.PdfFileReader("{}/Documents/{}/{}".format(os.getcwd(), year, file))
                nbPages = pdfRead.getNumPages()
                for i in range(nbPages):
                    page = pdfRead.getPage(i)
                    pageContent = page.extractText()
                    valid1 = not box_recherche1.get() or (champ_recherche1.get() != "" and re.search(champ_recherche1.get(), pageContent, re.IGNORECASE))
                    valid2 = not box_recherche2.get() or (champ_recherche2.get() != "" and re.search(champ_recherche2.get(), pageContent, re.IGNORECASE))
                    valid3 = not box_recherche3.get() or (champ_recherche3.get() != "" and re.search(champ_recherche3.get(), pageContent, re.IGNORECASE))
                    if valid1 and valid2 and valid3:
                        liste.insert(list_entries, "{}:{}".format(year + "/" + file, i + 1))
                        window.update()
                        list_entries = list_entries + 1
                    window.update()
    chercher_progress.destroy()
    window.update()
    fini=Label(window, text="Terminé", fg='green', font=("Helvetica", 12))
    fini.place(x=150, y=765)

def view_page(event):
    query = liste.get(liste.curselection()[0]).split(":")
    path_to_acrobat = os.path.abspath('C:\\Program Files (x86)\\Adobe\\Acrobat Reader DC\\Reader\\AcroRd32.exe')
    process = subprocess.Popen([path_to_acrobat, '/A', 'page=' + query[1] + '&search=' + champ_recherche1.get(), "{}/Documents/{}".format(os.getcwd(), query[0])], shell=False, stdout=subprocess.PIPE)
    window.update()


titre=Label(window, text="DOCUMENTS MUNICIPAUX CHAMBLY", fg='red', font=("Helvetica", 16))
titre.place(x=220, y=50)
documents=Label(window, text="Documents", font=("Helvetica", 14))
documents.place(x=100, y=120)
if ("Documents" in os.listdir(os.getcwd())):
    docs_chargés=Label(window, text="Documents déja chargés", fg='green', font=("Helvetica", 12))
    docs_chargés.place(x=150, y=180)
else:
    charger_docs_btn=Button(window, text="Charger documents", font=("Helvetica", 12))
    charger_docs_btn.place(x=150, y=180)
    charger_docs_btn.bind('<Button-1>', charger_docs)
recherche=Label(window, text="Recherche", font=("Helvetica", 14))
recherche.place(x=100, y=250)

champ_recherche1=Entry(window, font=("Helvetica", 12), bd=3, width=28)
champ_recherche1.place(x=150, y=305)
box_recherche1 = BooleanVar()
Checkbutton(window, text="Inclure", var=box_recherche1).place(x=410, y=305)

champ_recherche2=Entry(window, font=("Helvetica", 12), bd=3, width=28)
champ_recherche2.place(x=150, y=330)
box_recherche2 = BooleanVar()
Checkbutton(window, text="Inclure", var=box_recherche2).place(x=410, y=330)

champ_recherche3=Entry(window, font=("Helvetica", 12), bd=3, width=28)
champ_recherche3.place(x=150, y=355)
box_recherche3 = BooleanVar()
Checkbutton(window, text="Inclure", var=box_recherche3).place(x=410, y=355)

champ_recherche_btn=Button(window, text="Chercher", font=("Helvetica", 12))
champ_recherche_btn.place(x=530, y=320)
champ_recherche_btn.bind('<Button-1>', chercher)

voir_btn=Button(window, text="Voir", font=("Helvetica", 12))
voir_btn.place(x=630, y=320)
voir_btn.bind('<Button-1>', view_page)

years_frame = Frame(window, height=40, width=50)
years_frame.place(x=110, y=380)
years = []
try:
    years = os.listdir("{}/Documents".format(os.getcwd()))
except:
    pass
checkbuttons_search = []
search_years = []
for idx, year in enumerate(years):
    checkbuttons_search.append(IntVar())
    Checkbutton(years_frame, text=year, var=checkbuttons_search[-1]).grid(column=idx%7, row=int(idx/7),ipadx=15)
window.title('DOCUMENTS MUNICIPAUX CHAMBLY')
window.geometry("800x800")
window.mainloop()
