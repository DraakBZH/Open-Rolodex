from lxml import etree, objectify
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter.messagebox import *
import tkinter.font as tkFont
import os
import sys
import configparser  # pour la lecture du fichier ini
from PIL import Image, ImageTk
import operator
import unicodedata as uni
import shutil
import webbrowser

# ----------------------------------------------------------------------------
# déclaration
listimage = []
listtemp = []
datacat = {}
datalist = []
dataitem = []
ontop = True
resolution = (25, 25)
bgcolor = "#36393E"
textcolor = "#E5E5E5"
rougecolor = "#FF2B47"
bleucolor = "#666BFF"
datafile = 'data.xml'
dtdfile = 'data.dtd'
contact = {
    'Nom': '',
    'Adresse': '',
    'Code_Cli_OR': '',
    'Code_Cli_OT': '',
    'Tel': '',
    'Tel_SAV': '',
    'Fax': '',
    'Nom_commercial': '',
    'Tel_commercial': '',
    'Email': '',
    'Web_site': '',
    'ID_OR': '',
    'MDP_OR': '',
    'ID_OT': '',
    'MDP_OT': '',
    'Cmde_SAV': '',
    'PEC': '',
    'PEC_Derogation': '',
    'PEC_facturation': '',
    'PEC_annulation': '',
    'Annotations': '',
    'N_Finess': '',
    'N_Siret': '',
    'N_Intracomm': '',
    'Logo': ''}


# Fonction et méthode, bref, définition
def alwaysontop():
    global ontop
    ontop = not ontop
    root.wm_attributes("-topmost", ontop)
    win_contact.wm_attributes("-topmost", ontop)


def mycanva(event):
    canvas.configure(scrollregion=canvas.bbox("all"), width=200, height=500)


def alert(text):
    showinfo("Erreur", text)


def findicon():
    win_contact.filename = filedialog.askopenfilename(
        initialdir="/", title="Select file",
        filetypes=(("image files", "*.jpg;*png;*gif"), ("all files", "*.*")))
    win_contact.nametowidget("!!Logo").delete(0, END)
    win_contact.nametowidget("!!Logo").insert(END, win_contact.filename)


def normaliser(string):
    result = uni.normalize('NFD', string).encode('ascii', 'ignore')
    result = result.decode('ascii').lower()
    return result


def prefenetre(event):
    fenetre(event.widget.cget('textvariable'))


def ajoutcontact():
    addmodifcontact("0")


def modifiercontact():
    addmodifcontact(win_contact.nametowidget("!BModif").cget('text'))


def supprcontact():
    global dataroot
    global datafile
    global dtdfile
    removeCONTACT(
        win_contact.nametowidget("!BSuppr").cget('textvariable'), dataroot)
    titre = 'Suppression'
    message = 'Êtes vous sur de vouloir supprimer le contact ?'
    if askokcancel(titre, message):
        saveXML(dataroot, datafile, dtdfile)
        deleteframe()
        on_closing_contact()


def copy(info):
    if info == 'Adresse' or info == 'Annotations':
        content = win_contact.nametowidget("!!%s" % info).get(1.0, END)
    else:
        content = win_contact.nametowidget("!!%s" % info).get()
    root.clipboard_clear()
    root.clipboard_append(content)


def openurl(url):
    webbrowser.open_new(win_contact.nametowidget("!!%s" % url).get())


def getcontact(ident):
    global contact
    global dataroot
    tempcontact = contact
    for cat in dataroot:
        for con in cat:
            if con.get("data-id") == ident:
                for elem in con:
                    tempcontact[elem.tag] = elem.text
                return tempcontact
    return False


def addmodifcontact(ident):
    global contact
    global popupMenu
    if ident != '0':
        modifcontact = getcontact(ident)
        modif = True
    else:
        modif = False

    starty = 100
    win_contact.withdraw()
    win_contact.deiconify()
    win_contact.nametowidget("!icon").config(image=logo_icon)
    if modif:
        win_contact.nametowidget("!contact").config(text='Modif Contact')
        win_contact.nametowidget("!BSave").config(text='modif')
        win_contact.nametowidget("!BSave").config(textvariable=ident)
        win_contact.nametowidget("!BSuppr").config(text='suppr')
        win_contact.nametowidget("!BSuppr").config(textvariable=ident)
        win_contact.nametowidget("!BSuppr").place(x=160, y=55)
    else:
        win_contact.nametowidget("!contact").config(text='Nouveau Contact')
        win_contact.nametowidget("!Categorie").place(x=5, y=55)
        tkvar = StringVar(root, name='!!Categorie')
        choices = {'Fournisseurs', 'Mutuelles', 'Magasins'}
        tkvar.set('Aucune')
        popupMenu = OptionMenu(win_contact, tkvar, *choices)
        popupMenu.place(x=110, y=55)
        win_contact.nametowidget("!BSave").config(text='new')
        win_contact.nametowidget("!BSave").config(textvariable="new")
        win_contact.nametowidget("!BSuppr").place_forget()

    win_contact.nametowidget("!BSave").place(x=210, y=55)
    win_contact.nametowidget("!BModif").place_forget()

    for key, value in contact.items():
        if key == 'Adresse':
            win_contact.nametowidget("!!%s" % key).config(
                state="normal", background=textcolor, foreground=bgcolor,
                selectbackground=bgcolor, selectforeground=textcolor)
            win_contact.nametowidget("!%s" % key).place(x=5, y=starty)
            win_contact.nametowidget("!!%s" % key).place(x=110, y=starty)
            win_contact.nametowidget("!!%s" % key).delete(1.0, END)
            adryscrollbar.place(in_=win_contact.nametowidget(
                "!!%s" % key), relx=1.0, relheight=1.0, bordermode="outside")
            if modif and modifcontact[key] is not None:
                win_contact.nametowidget(
                    "!!%s" % key).insert(
                    END, modifcontact[key].replace('\\n', '\n'))
            starty += 71
        elif key == 'Annotations':
            win_contact.nametowidget("!!%s" % key).config(
                state="normal", background=textcolor, foreground=bgcolor,
                selectbackground=bgcolor, selectforeground=textcolor)
            win_contact.nametowidget("!%s" % key).place(x=5, y=starty)
            win_contact.nametowidget("!!%s" % key).place(x=110, y=starty)

            win_contact.nametowidget("!!%s" % key).delete(1.0, END)
            anotyscrollbar.place(in_=win_contact.nametowidget(
                "!!%s" % key), relx=1.0, relheight=1.0, bordermode="outside")
            if modif and modifcontact[key] is not None:
                win_contact.nametowidget(
                    "!!%s" % key).insert(
                    END, modifcontact[key].replace('\\n', '\n'))
            starty += 86
        else:
            win_contact.nametowidget("!!%s" % key).config(
                state="normal", background=textcolor, foreground=bgcolor,
                selectbackground=bgcolor, selectforeground=textcolor)
            win_contact.nametowidget("!%s" % key).place(x=5, y=starty)
            win_contact.nametowidget("!!%s" % key).place(x=110, y=starty)

            win_contact.nametowidget("!!%s" % key).delete(0, END)
            if modif:
                if modifcontact[key] is None:
                    modifcontact[key] = ""
                win_contact.nametowidget(
                    "!!%s" % key).insert(END, modifcontact[key])
            if key == "Logo":
                win_contact.nametowidget("!BLogo").place(x=205, y=starty)
            starty += 21


def fenetre(ident):
    global dataroot
    global Nom
    global listimage
    global listtemp
    global popupMenu
    listtemp = []
    starty = 55
    win_contact.withdraw()
    win_contact.deiconify()
    ident = str(ident)
    popupMenu.place_forget()
    win_contact.nametowidget("!BModif").config(text=ident)
    win_contact.nametowidget("!BModif").place(x=210, y=10)
    win_contact.nametowidget("!Categorie").place_forget()
    win_contact.nametowidget("!BSave").place_forget()
    win_contact.nametowidget("!BSuppr").place_forget()
    win_contact.nametowidget("!contact").config(text='Contact')
    for cat in dataroot:
        for contact in cat:
            if contact.get('data-id') == ident:
                for info in contact:
                    if info.tag == "Logo":
                        try:
                            con_icon = ImageTk.PhotoImage(Image.open(
                                info.text).resize((45, 45), Image.ANTIALIAS))
                        except Exception as e:
                            con_icon = ImageTk.PhotoImage(Image.open(
                                "./img/contact.png").resize(
                                (45, 45), Image.ANTIALIAS))
                        listtemp.append(con_icon)  # image dans liste!
                        win_contact.nametowidget(
                            "!icon").config(image=con_icon)
                        win_contact.nametowidget(
                            "!%s" % info.tag).place_forget()
                        win_contact.nametowidget(
                            "!!%s" % info.tag).place_forget()
                        win_contact.nametowidget(
                            "!BLogo").place_forget()
                    elif info.text is not None and info.tag == "Adresse":
                        win_contact.nametowidget(
                            "!%s" % info.tag).place(x=5, y=starty)
                        win_contact.nametowidget(
                            "!!%s" % info.tag).place(x=110, y=starty)
                        win_contact.nametowidget(
                            "!!%s" % info.tag).config(
                            state="normal", background=bgcolor,
                            foreground=textcolor, selectbackground=textcolor,
                            selectforeground=bgcolor)
                        win_contact.nametowidget(
                            "!!%s" % info.tag).delete(1.0, END)
                        temp = info.text.replace('\\n', '\n')
                        win_contact.nametowidget(
                            "!!%s" % info.tag).insert(END, temp)
                        win_contact.nametowidget(
                            "!!%s" % info.tag).config(state="disabled")
                        adryscrollbar.place(
                            in_=win_contact.nametowidget("!!%s" % info.tag),
                            relx=1.0, relheight=1.0, bordermode="outside")
                        win_contact.nametowidget(
                            "!B%s" % info.tag).place(x=255, y=starty)
                        starty += 71
                    elif info.text is not None and info.tag == "Annotations":
                        win_contact.nametowidget(
                            "!%s" % info.tag).place(x=5, y=starty)
                        win_contact.nametowidget(
                            "!!%s" % info.tag).place(x=110, y=starty)
                        win_contact.nametowidget(
                            "!!%s" % info.tag).config(
                            state="normal", background=bgcolor,
                            foreground=textcolor, selectbackground=textcolor,
                            selectforeground=bgcolor)
                        win_contact.nametowidget(
                            "!!%s" % info.tag).delete(1.0, END)
                        temp = info.text.replace('\\n', '\n')
                        win_contact.nametowidget(
                            "!!%s" % info.tag).insert(END, temp)
                        win_contact.nametowidget(
                            "!!%s" % info.tag).config(state="disabled")
                        anotyscrollbar.place(
                            in_=win_contact.nametowidget("!!%s" % info.tag),
                            relx=1.0, relheight=1.0, bordermode="outside")
                        win_contact.nametowidget(
                            "!B%s" % info.tag).place(x=255, y=starty)
                        starty += 86
                    elif info.text is not None:
                        win_contact.nametowidget(
                            "!%s" % info.tag).place(x=5, y=starty)
                        win_contact.nametowidget(
                            "!!%s" % info.tag).place(x=110, y=starty)
                        win_contact.nametowidget(
                            "!!%s" % info.tag).config(
                            state="normal")
                        win_contact.nametowidget(
                            "!!%s" % info.tag).delete(0, END)
                        win_contact.nametowidget(
                            "!!%s" % info.tag).insert(0, info.text)
                        win_contact.nametowidget(
                            "!!%s" % info.tag).config(state="readonly")
                        win_contact.nametowidget(
                            "!B%s" % info.tag).place(x=255, y=starty)
                        starty += 21
                    else:
                        win_contact.nametowidget(
                            "!%s" % info.tag).place_forget()
                        win_contact.nametowidget(
                            "!!%s" % info.tag).place_forget()
                        win_contact.nametowidget(
                            "!B%s" % info.tag).place_forget()


def importXML(xml, dtd):
    if os.path.exists(xml):
        try:
            dtd = etree.DTD(open(dtd, 'rb'))
            tree = objectify.parse(open(xml, 'rb'))
        except Exception as e:
            showinfo(
                "Erreur", "Le fichier %s est corrompu. "
                "Corrigez le fichier ou supprimez le. "
                "L'erreur est la suivante: %s" % (xml, e))
            sys.exit(0)

        if not dtd.validate(tree):
            showinfo(
                "Erreur", "Le fichier %s n'est pas conforme."
                "Corrigez le fichier ou supprimez le. L'erreur est "
                "la suivant: %s" % (xml, dtd.error_log.filter_from_errors()))
            sys.exit(0)
        parser = etree.XMLParser(remove_blank_text=True)
        returndata = etree.parse(xml, parser)
        return returndata.getroot()

    else:
        root = etree.Element("Annuaire")
        etree.SubElement(root, "Fournisseurs")
        etree.SubElement(root, "Mutuelles")
        etree.SubElement(root, "Magasins")
        tree = etree.ElementTree(root)
        tree.write(xml, pretty_print=True)
        return tree.getroot()


def saveXML(data, xml, dtd):
    et = etree.ElementTree(data)
    et.write(xml, pretty_print=True)


def findlastidcontact(data):
    liste = []
    for child in data:
        for children in child:
            liste.append(children.attrib['data-id'])
    if len(liste) == 0:
        result = 'Con-00000'
    else:
        liste.sort()
        index = liste[-1].split("-")[-1]
        index = int(index) + 1
        result = 'Con-%05d' % index

    return result


def verifAddCONTACT():
    global contact
    global popupMenu
    global dataroot
    valid = True
    Ncontact = contact
    if win_contact.nametowidget("!BSave").cget('text') == "modif":
        modif = True
    else:
        modif = False

    for key, value in Ncontact.items():

        if key == 'Adresse':
            Ncontact[key] = win_contact.nametowidget("!!Adresse").get(1.0, END)
            Ncontact[key] = Ncontact[key].replace('\n', '\\n')
            if Ncontact[key] == '\n' or Ncontact[key] == '\\n':
                Ncontact[key] = ''
        elif key == 'Annotations':
            Ncontact[key] = win_contact.nametowidget(
                "!!Annotations").get(1.0, END)
            Ncontact[key] = Ncontact[key].replace('\n', '\\n')
            if Ncontact[key] == '\n' or Ncontact[key] == '\\n':
                Ncontact[key] = ''
        else:
            Ncontact[key] = win_contact.nametowidget("!!%s" % key).get()

    if Ncontact['Nom'] == '' or Ncontact['Nom'] is None:
        alert("Vous devez obligatoirement entrer un Nom")
        valid = False
    if Ncontact['Logo'] == '' or Ncontact['Nom'] is None:
        Ncontact['Logo'] = './img/contact.png'
    else:
        try:
            with open(Ncontact['Logo']):
                pass
        except Exception as e:
            alert("Le fichier choisi en logo est introuvable")
            valid = False
    if tkvar.get() == 'Aucune' and not modif:
        alert("Veuillez selectionner une Catégorie")
        valid = False

    if valid:
        if modif:
            alterCONTACT(
                win_contact.nametowidget(
                    "!BSave").cget('textvariable'), Ncontact, dataroot)
        else:
            addCONTACT(Ncontact, tkvar.get(), dataroot)

        alert("Enregistrement Valider")
        on_closing_contact()


def addCONTACT(newcontact, categorie, data):
    global datafile
    global dtdfile
    for child in data:
        if child.tag == categorie:
            con_id = findlastidcontact(data)
            contact = etree.SubElement(child, "Contact")
            contact.attrib['data-id'] = con_id
            for key, value in newcontact.items():
                if key == 'Logo':
                    ext = os.path.splitext(newcontact['Logo'])
                    shutil.copyfile(
                        newcontact['Logo'],
                        './img/contact/%s%s' % (con_id, ext[-1]))
                    value = './img/contact/%s%s' % (con_id, ext[-1])

                etree.SubElement(contact, key).text = str(value)
    saveXML(data, datafile, dtdfile)
    deleteframe()


def alterCONTACT(ident, modification, data):
    global datafile
    global dtdfile
    for child in data:
        for children in child:
            if children.attrib['data-id'] == ident:
                for content in children:
                    content.text = str(modification[content.tag])
                    if (content.tag == 'Adresse' or content.tag == 'Annotations'):
                        content.text = str(modification[content.tag])[:-2]
                    if content.tag == 'Logo' and "./img/contact" not in modification['Logo']:
                        ext = os.path.splitext(modification['Logo'])
                        shutil.copyfile(
                            modification['Logo'],
                            './img/contact/%s%s' % (ident, ext[-1]))
                        content.text = './img/contact/%s%s' % (ident, ext[-1])
                        modification['Logo'] = content.text
    saveXML(data, datafile, dtdfile)
    deleteframe()


def removeCONTACT(ident, data):
    for child in data:
        for children in child:
            if children.attrib['data-id'] == ident:
                children.getparent().remove(children)


def on_closing():
    global tailleposition
    global tailleposition_contact
    # enregistrement cfg
    cfg_save = configparser.ConfigParser()
    cfg_save.read('config.cfg')
    cfg_save.set('Fenetre', 'tailleposition', root.winfo_geometry())
    cfg_save.set(
        'Fenetre_Contact', 'tailleposition_contact', win_contact.geometry())
    cfg_save.write(open('config.cfg', 'w'))
    root.destroy()


def on_closing_contact():
    global tailleposition_contact
    tailleposition_contact = win_contact.winfo_geometry()
    win_contact.withdraw()


def press(event):
    result = saisi.get()
    result = normaliser(result)
    breset.place(x=162, y=36)
    temp = frame.winfo_children()
    for elem in temp:
        if "separator" not in str(elem):
            if "con-" in str(elem):
                temp2 = normaliser(elem.cget('text'))
                if result in temp2:
                    elem.grid()
                else:
                    elem.grid_remove()


def reset():
    saisi.delete(0, END)
    breset.place_forget()
    for child in frame.children:
        try:
            frame.nametowidget(child).grid()
        except Exception as e:
            pass


def deleteframe():
    temp = frame.winfo_children()
    for elem in temp:
        elem.destroy()
    generatedata()
    generatemainframe()


def generatemainframe():
    global dataroot
    global datalist
    global dataitem
    global datacat
    r = 2

    for key, value in datacat.items():
        Label(
            frame, text=key, borderwidth=1, name="!!" + key.lower(),
            font=Categorie1).grid(row=r, column=0, sticky='w', columnspan=3)
        r = r + 1
        ttk.Separator(frame, orient=HORIZONTAL).grid(
            row=r, column=0, sticky='ew', columnspan=3, pady=3)
        for idatalist in datacat[key]:
            r = r + 1
            try:
                im = ImageTk.PhotoImage(
                    Image.open(idatalist[2]).resize(
                        resolution, Image.ANTIALIAS))
            except Exception as e:
                im = ImageTk.PhotoImage(
                    Image.open("./img/contact.png").resize(
                        resolution, Image.ANTIALIAS))
            listimage.append(im)
            # c'est naze, mais les photos doivent etre
            # dans une liste sinon ça ne marche pas...
            findname = normaliser(idatalist[0])
            icone = Label(
                frame, text=idatalist[1].decode("utf-8"),
                borderwidth=1, font=Liste1,
                textvariable=idatalist[0], image=listimage[-1],
                name="!!" + findname + "_img")

            icone.grid(row=r, column=0, sticky='w')
            icone.bind("<Button-1>", prefenetre)
            name = Label(
                frame, text=idatalist[1].decode("utf-8"), borderwidth=1,
                font=Liste1, textvariable=idatalist[0],
                name="!!" + findname + "_lab")
            name.grid(row=r, column=1, sticky='w')
            name.bind("<Button-1>", prefenetre)

        r = r + 1


# ----------------------------------------------------------------------------
# Fichier CFG - ouverture ou création
cfg = configparser.ConfigParser()
if os.path.exists("config.cfg"):
    cfg.read('config.cfg')
    tailleposition = cfg.get('Fenetre', 'tailleposition')
    tailleposition_contact = cfg.get(
        'Fenetre_Contact', 'tailleposition_contact')
else:
    # construction du fichier de configuration
    S1 = 'Fenetre'
    cfg.add_section(S1)
    cfg.set(S1, 'tailleposition', "210x630+120+120")
    tailleposition = "210x630+120+120"

    S2 = 'Fenetre_Contact'
    cfg.add_section(S2)
    cfg.set(S2, 'tailleposition_contact', "210x350+340+250")
    tailleposition_contact = "210x350+340+250"

    cfg.write(open('config.cfg', 'w'))


def generatedata():
    global dataroot
    global datalist
    global dataitem
    global datacat
    # ----------------------------------------------------------------------------
    # Importation des données XML
    dataroot = importXML(datafile, dtdfile)

    # generation de la liste d'element pour la fenetre main
    for categorie in dataroot:

        for element in categorie:
            element_name = element.find('Nom').text
            element_name = element_name.encode('utf-8')
            element_id = element.get('data-id')
            element_logo = element.find('Logo').text
            dataitem.append(element_id)
            dataitem.append(element_name)
            dataitem.append(element_logo)
            datalist.append(dataitem)
            dataitem = []

        datalist = sorted(datalist, key=operator.itemgetter(1))
        datacat[categorie.tag] = datalist
        datalist = []


generatedata()
# ----------------------------------------------------------------------------
# propriété fenetres principale
root = Tk()
root.tk_setPalette(background=bgcolor, foreground=textcolor)
root.wm_attributes("-topmost", 1)
root.minsize(width=235, height=575)
root.geometry(tailleposition)
root.title("Open Rolodex")
root.iconbitmap(default='./img/repository.ico') 

myframe = Frame(root, relief=FLAT, width=100, height=300, bd=1)
myframe.place(x=5, y=60)

# gestion de la frame dans un canva dans une frame pour une scrollbar
canvas = Canvas(myframe)
frame = Frame(canvas)
myscrollbar = Scrollbar(myframe, orient="vertical", command=canvas.yview)
canvas.configure(yscrollcommand=myscrollbar.set)
myscrollbar.pack(side="right", fill="y")
canvas.pack(side="left")
canvas.create_window((0, 0), window=frame, anchor='nw')
frame.bind("<Configure>", mycanva)

# ----------------------------------------------------------------------------
# fonts et titre
Titre1 = tkFont.Font(family='Helvetica', size=18, weight='bold')
Categorie1 = tkFont.Font(family='Helvetica', size=12, weight='bold')
Liste1 = tkFont.Font(family='Helvetica', size=11, weight='normal')

# ----------------------------------------------------------------------------
# propriété de la fenetre contact
win_contact = Toplevel(root)
win_contact.title("Contact")
win_contact.wm_attributes("-topmost", 1)
win_contact.minsize(width=280, height=750)
win_contact.geometry(tailleposition_contact)
win_contact.withdraw()

# habillage fenetre contact
Label(
    win_contact, text='Contact', borderwidth=1,
    font=Titre1, name='!contact').place(x=65, y=12)

logo_icon = ImageTk.PhotoImage(
    Image.open("./img/contact.png").resize((50, 50), Image.ANTIALIAS))
Label(
    win_contact, text='X', image=logo_icon, width=50, height=50,
    name="!icon").place(x=10, y=5)

logo_save = ImageTk.PhotoImage(
    Image.open("./img/save.png").resize((35, 35), Image.ANTIALIAS))
Button(
    win_contact, name='!BSave', image=logo_save,
    width=35, height=35, command=verifAddCONTACT)

logo_modif = ImageTk.PhotoImage(
    Image.open("./img/modif.png").resize((35, 35), Image.ANTIALIAS))
Button(
    win_contact, name='!BModif', image=logo_modif,
    width=35, height=35, command=modifiercontact)

logo_suppr = ImageTk.PhotoImage(
    Image.open("./img/trash.png").resize((35, 35), Image.ANTIALIAS))
Button(
    win_contact, name='!BSuppr', image=logo_suppr,
    width=35, height=35, command=supprcontact)

Label(win_contact, name='!Categorie', text='Catégorie :')
Label(win_contact, name='!Nom', text='Nom :')
Label(win_contact, name='!Adresse', text='Adresse :')
Label(
    win_contact, name='!Code_Cli_OR', text='Code Cli OR :',
    foreground=rougecolor)
Label(
    win_contact, name='!Code_Cli_OT', text='Code Cli OT :',
    foreground=bleucolor)
Label(win_contact, name='!Tel', text='Tel :')
Label(win_contact, name='!Tel_SAV', text='Tel_SAV :')
Label(win_contact, name='!Fax', text='Fax :')
Label(win_contact, name='!Nom_commercial', text='Nom commercial :')
Label(win_contact, name='!Tel_commercial', text='Tel commercial :')
Label(win_contact, name='!Email', text='Email :')
Label(win_contact, name='!Web_site', text='Web site :')
Label(win_contact, name='!ID_OR', text='ID OR :', foreground=rougecolor)
Label(win_contact, name='!MDP_OR', text='MDP OR :', foreground=rougecolor)
Label(win_contact, name='!ID_OT', text='ID OT :', foreground=bleucolor)
Label(win_contact, name='!MDP_OT', text='MDP OT :', foreground=bleucolor)
Label(win_contact, name='!Cmde_SAV', text='Cmde SAV :')
Label(win_contact, name='!PEC', text='PEC :')
Label(win_contact, name='!PEC_Derogation', text='PEC Derogation :')
Label(win_contact, name='!PEC_facturation', text='PEC Facturation :')
Label(win_contact, name='!PEC_annulation', text='PEC Annulation :')
Label(win_contact, name='!Annotations', text='Annotations :')
Label(win_contact, name='!N_Finess', text='N° Finess :')
Label(win_contact, name='!N_Siret', text='N° Siret :')
Label(win_contact, name='!N_Intracomm', text='N° Intracomm :')
Label(win_contact, name='!Logo', text='Logo :')

tkvar = StringVar(root, name='!!Categorie')
choices = {'Fournisseurs', 'Mutuelles', 'Magasins'}
tkvar.set('Aucune')  # set the default option
popupMenu = OptionMenu(win_contact, tkvar, *choices)

Entry(
    win_contact, name='!!Nom', state='readonly', width='23',
    insertbackground=textcolor)
Text(win_contact, name='!!Adresse', state='disabled', width='15', height='4')
Entry(
    win_contact, name='!!Code_Cli_OR', state='readonly', width='23',
    foreground=rougecolor, insertbackground=textcolor)
Entry(
    win_contact, name='!!Code_Cli_OT', state='readonly', width='23',
    foreground=bleucolor)
Entry(win_contact, name='!!Tel', state='readonly', width='23')
Entry(win_contact, name='!!Tel_SAV', state='readonly', width='23')
Entry(win_contact, name='!!Fax', state='readonly', width='23')
Entry(win_contact, name='!!Nom_commercial', state='readonly', width='23')
Entry(win_contact, name='!!Tel_commercial', state='readonly', width='23')
Entry(win_contact, name='!!Email', state='readonly', width='23')
Entry(win_contact, name='!!Web_site', state='readonly', width='23')
Entry(
    win_contact, name='!!ID_OR', state='readonly', width='23',
    foreground=rougecolor)
Entry(
    win_contact, name='!!MDP_OR', state='readonly', width='23',
    foreground=rougecolor)
Entry(
    win_contact, name='!!ID_OT', state='readonly', width='23',
    foreground=bleucolor)
Entry(
    win_contact, name='!!MDP_OT', state='readonly', width='23',
    foreground=bleucolor)
Entry(win_contact, name='!!Cmde_SAV', state='readonly', width='23')
Entry(win_contact, name='!!PEC', state='readonly', width='23')
Entry(win_contact, name='!!PEC_Derogation', state='readonly', width='23')
Entry(win_contact, name='!!PEC_facturation', state='readonly', width='23')
Entry(win_contact, name='!!PEC_annulation', state='readonly', width='23')
Text(
    win_contact, name='!!Annotations',
    state='disabled', width='15', height='5')
Entry(win_contact, name='!!N_Finess', state='readonly', width='23')
Entry(win_contact, name='!!N_Siret', state='readonly', width='23')
Entry(win_contact, name='!!N_Intracomm', state='readonly', width='23')
Entry(win_contact, name='!!Logo', state='readonly', width='15')

logo_copy = ImageTk.PhotoImage(
    Image.open("./img/copy.png").resize((15, 15), Image.ANTIALIAS))

logo_internet = ImageTk.PhotoImage(
    Image.open("./img/internet.png").resize((15, 15), Image.ANTIALIAS))

Button(
    win_contact, name='!BNom', image=logo_copy, width=15, height=15,
    command=lambda: copy('Nom'))
Button(
    win_contact, name='!BAdresse', image=logo_copy, width=15, height=15,
    command=lambda: copy('Adresse'))
Button(
    win_contact, name='!BCode_Cli_OR', image=logo_copy, width=15, height=15,
    command=lambda: copy('Code_Cli_OR'))
Button(
    win_contact, name='!BCode_Cli_OT', image=logo_copy, width=15, height=15,
    command=lambda: copy('Code_Cli_OT'))
Button(
    win_contact, name='!BTel', image=logo_copy, width=15, height=15,
    command=lambda: copy('Tel'))
Button(
    win_contact, name='!BTel_SAV', image=logo_copy, width=15, height=15,
    command=lambda: copy('Tel_SAV'))
Button(
    win_contact, name='!BFax', image=logo_copy, width=15, height=15,
    command=lambda: copy('Fax'))
Button(
    win_contact, name='!BNom_commercial', image=logo_copy, width=15, height=15,
    command=lambda: copy('Nom_commercial'))
Button(
    win_contact, name='!BTel_commercial', image=logo_copy, width=15, height=15,
    command=lambda: copy('Tel_commercial'))
Button(
    win_contact, name='!BEmail', image=logo_copy, width=15, height=15,
    command=lambda: copy('Email'))
Button(
    win_contact, name='!BWeb_site', image=logo_internet, width=15, height=15,
    command=lambda: openurl('Web_site'))
Button(
    win_contact, name='!BID_OR', image=logo_copy, width=15, height=15,
    command=lambda: copy('ID_OR'))
Button(
    win_contact, name='!BMDP_OR', image=logo_copy, width=15, height=15,
    command=lambda: copy('MDP_OR'))
Button(
    win_contact, name='!BID_OT', image=logo_copy, width=15, height=15,
    command=lambda: copy('ID_OT'))
Button(
    win_contact, name='!BMDP_OT', image=logo_copy, width=15, height=15,
    command=lambda: copy('MDP_OT'))
Button(
    win_contact, name='!BCmde_SAV', image=logo_copy, width=15, height=15,
    command=lambda: copy('Cmde_SAV'))
Button(
    win_contact, name='!BPEC', image=logo_copy, width=15, height=15,
    command=lambda: copy('PEC'))
Button(
    win_contact, name='!BPEC_Derogation', image=logo_copy, width=15, height=15,
    command=lambda: copy('PEC_Derogation'))
Button(
    win_contact, name='!BPEC_facturation', image=logo_copy, width=15,
    height=15, command=lambda: copy('PEC_facturation'))
Button(
    win_contact, name='!BPEC_annulation', image=logo_copy, width=15, height=15,
    command=lambda: copy('PEC_annulation'))
Button(
    win_contact, name='!BAnnotations', image=logo_copy, width=15, height=15,
    command=lambda: copy('Annotations'))
Button(
    win_contact, name='!BN_Finess', image=logo_copy, width=15, height=15,
    command=lambda: copy('N_Finess'))
Button(
    win_contact, name='!BN_Siret', image=logo_copy, width=15, height=15,
    command=lambda: copy('N_Siret'))
Button(
    win_contact, name='!BN_Intracomm', image=logo_copy, width=15, height=15,
    command=lambda: copy('N_Intracomm'))

logo_butt = ImageTk.PhotoImage(
    Image.open("./img/photo.png").resize((18, 18), Image.ANTIALIAS))
Button(
    win_contact, name='!BLogo',
    image=logo_butt, width=42, height=14, command=findicon)

adryscrollbar = Scrollbar(
    win_contact, orient=VERTICAL,
    command=win_contact.nametowidget("!!Adresse").yview)
win_contact.nametowidget("!!Adresse")['yscroll'] = adryscrollbar.set
anotyscrollbar = Scrollbar(
    win_contact, orient=VERTICAL,
    command=win_contact.nametowidget("!!Annotations").yview)
win_contact.nametowidget("!!Annotations")['yscroll'] = anotyscrollbar.set

for child in win_contact.children:
    if "!!" in child and '!!Categorie' not in child:
        if '!!Annotations' not in child and '!!Adresse' not in child:
            win_contact.nametowidget(child).config(
                state="readonly", readonlybackground=bgcolor,
                selectbackground=textcolor, selectforeground=bgcolor)

# ----------------------------------------------------------------------------
# Définition du menu

menubar = Menu(root)

menu1 = Menu(menubar, tearoff=0)
menu1.add_command(label="Ajouter", command=ajoutcontact)
menu1.add_separator()
menu1.add_command(label="Quitter", command=on_closing)
menubar.add_cascade(label="Fichier", menu=menu1)

menu2 = Menu(menubar, tearoff=0)
menu2.add_checkbutton(
    label="Pas Toujours Visible", onvalue=True,
    offvalue=False, variable=ontop, command=alwaysontop)
menubar.add_cascade(label="Vue", menu=menu2)

root.config(menu=menubar)

# ----------------------------------------------------------------------------
# Habillage et generation de la fenetre main

# titre annuaire fixe
Label(
    root, text='Annuaire', borderwidth=1,
    font=Titre1, name='!!annuaire').place(x=5, y=5)

# champ de recherche
saisi = Entry(root, width=20)
saisi.place(x=35, y=37)
saisi.bind('<KeyRelease>', press)

croix = ImageTk.PhotoImage(
    Image.open("./img/close.png").resize((15, 15), Image.ANTIALIAS))
breset = Button(
    root, text='X', command=reset, image=croix,
    width=13, height=13, name="!!breset")
breset.place(x=162, y=36)
breset.place_forget()

loupe = ImageTk.PhotoImage(
    Image.open("./img/find.png").resize((20, 20), Image.ANTIALIAS))
loupelabel = Label(root, image=loupe, name="!!loupe")
loupelabel.place(x=5, y=35)

generatemainframe()

root.update()

win_contact.protocol("WM_DELETE_WINDOW", on_closing_contact)
root.protocol("WM_DELETE_WINDOW", on_closing)  # fermeture de la fenetre
root.mainloop()
