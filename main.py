import os
from customtkinter import *
from docx2pdf import convert
from docxtpl import DocxTemplate
from time import strftime, localtime

# functions
def center_window(root, width, height):
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    
    x = (screen_width // 2) - (width // 2)
    y = (screen_height // 2) - (height // 2)

    root.geometry(f"{width}x{height}+{x}+{y}")

def main():
    root = CTk()
    root.title("Attestation de poursuite d'étude")
    center_window(root, 400, 600)


    def choix_de_filiere(filiere):
        if filiere == "STPI" and ("Troisième année" in niveaus):
            niveaus.remove("Troisième année")
        elif "Troisième année" not in niveaus:
            niveaus.append("Troisième année")
        niveau_menu.configure(values=niveaus)

    def choix_de_niveau(niveau):
        print(f"Niveau selectioné: {niveau}")
    
    def clear():
        nom_entry.delete(0, 'end')
        prenom_entry.delete(0, 'end')
        filiere_menu.set("STPI")
        niveau_menu.set("Première année")
        naiss_entry.delete(0, 'end')
        massar_entry.delete(0, 'end')

    def valider_clicked():
        doc = DocxTemplate("docs/attestation.docx")

        nom = nom_entry.get()
        prenom = prenom_entry.get()
        filiere = filiere_menu.get()
        niveau = niveau_menu.get()
        naiss = naiss_entry.get()
        massar = massar_entry.get()

        doc.render({"nom_prenom":f"{nom.upper()} {prenom.upper()}",
                    "date_naiss":f"{naiss}",
                    "massar":f"{massar.upper()}",
                    "niveau":f"{niveau}",
                    "filiere":f"{filiere}",
                    "date":f"{strftime("%d/%m/%Y",localtime())}"})
        
        os.mkdir(f"docs/{prenom} {nom} {filiere}")
        doc.save(f"docs/{prenom} {nom} {filiere}/Attestation de scolarité.docx")
        convert(f"docs/{prenom} {nom} {filiere}/Attestation de scolarité.docx",
                f"docs/{prenom} {nom} {filiere}/Attestation de scolarité.pdf")
        clear()

    # !
    # !
    # !
    # !
    # ! --> You should verify if the user didn't fill all the entries. If yes, show him a message error
    # ! -->     and let him continue entring his infomations until all informations is entred.
    # !
    # !
    # ! --> You can use this as a seperated class/funtion on your massar project, but instead of entring
    # ! -->     the student informations to give him a attestation, it's enough to take just the CIN(carte nationale)
    # ! -->     or CNE(code nationale d'étudiant) and search for the student on the data base and give him 
    # ! -->     the attestation.
    # !
    # !
    # !
    # !

    # frames
    top_frame = CTkFrame(root, fg_color='transparent', border_width=2, border_color='#008FE0')
    buttom_frame = CTkFrame(root, fg_color='transparent', border_width=2, border_color='#008FE0')
    top_frame.pack(fill='x', padx=15, pady=(15,0))
    buttom_frame.pack(expand=True, fill='both', padx=15, pady=15)

    for i in range(8):
        buttom_frame.grid_rowconfigure(i, weight=1)
    for i in range(2):
        buttom_frame.grid_columnconfigure(i, weight=1)
    
    # buttons, labels,. ..
    title_label = CTkLabel(top_frame, text="Attestation de poursuite d'étude", font=("Helvitica", 16, "bold"))
    title_label.pack(padx=5, pady=5)

    filieres = ["STPI", "G.Civil", "DSCC", "G.Indus", "SICC", "G.Info", "G.Elect", "MGSI","ITIRC", "GSEIR"]
    filiere_menu = CTkOptionMenu(buttom_frame, values=filieres, command=choix_de_filiere)
    niveaus = ["Première année", "Deuxième année"]
    niveau_menu = CTkOptionMenu(buttom_frame, values=niveaus, command=choix_de_niveau)

    nom_label = CTkLabel(buttom_frame, text="Nom:", font=("Helvitica", 14))
    nom_entry = CTkEntry(buttom_frame)
    prenom_label = CTkLabel(buttom_frame, text="Prénom:", font=("Helvitica", 14))
    prenom_entry = CTkEntry(buttom_frame)
    filiere_label = CTkLabel(buttom_frame, text="Filière:", font=("Helvitica", 14))
    niveau_label = CTkLabel(buttom_frame, text="Niveau:", font=("Helvitica", 14))
    naiss_label = CTkLabel(buttom_frame, text="Date de Naissance:\n(jj/mm/aaaa)")
    naiss_entry = CTkEntry(buttom_frame)
    massar_label = CTkLabel(buttom_frame, text="Code Massar:")
    massar_entry = CTkEntry(buttom_frame)
    valider_btn = CTkButton(buttom_frame, text="Valider", command=valider_clicked)
    
    nom_label.grid(row=0, column=0, pady=(20,0))
    nom_entry.grid(row=0, column=1, pady=(20,0))
    prenom_label.grid(row=1, column=0)
    prenom_entry.grid(row=1, column=1)
    filiere_label.grid(row=2, column=0)
    filiere_menu.grid(row=2, column=1)
    niveau_label.grid(row=3, column=0)
    niveau_menu.grid(row=3, column=1)
    naiss_label.grid(row=4, column=0)
    naiss_entry.grid(row=4, column=1)
    massar_label.grid(row=5, column=0)
    massar_entry.grid(row=5, column=1)
    valider_btn.grid(row=6, column=0, columnspan=2)

    # run
    root.mainloop()

if __name__ == "__main__":
    main()