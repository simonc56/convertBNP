#!/usr/bin/env python
# -*- coding:utf-8 -*-
#
# auteur             twolaw_at_free_dot_fr
# nom                convertBNP_4col.py
# description        Lit les relevés bancaires de la BNP en PDF dans le répertoire courant pour en générer des CSV
#                    Nécessite le fichier pdftotext.exe en version 3.03 issu de l'archive xpdf (gratuit, GPL2)
# ------------------
# 10-nov-2013 v1     pour python3
# 26-nov-2014 v1.1   ajout de la version "_4col" qui sépare crédit et débits en 2 colonnes distinctes
# 28-jan-2015 v1.2   correction du bug "Mixing iteration and read methods would lose data"
# ------------------
# chaque opération bancaire contient 4 éléments :
#
#   - date : string de type 'JJ/MM/AAAA'
#   - description de l'opération : string
#   - valeur de débit : string
#   - valeur de crédit : string
#
# fichiers PDF de la forme RCHQ_101_300040012300001234567_20131026_2153.PDF

PREFIXE_COMTE = "RCHQ_101_300040012300001234567_"
PREFIXE_COMPTE = "RLV_CHQ_300040181600000906809_"

CSV_SEP        = ";"

import os, subprocess, shutil
import re
pattern = re.compile('(\W)')

if os.name == 'nt':
    PDFTOTEXT = 'pdftotext.exe'
    PREFIXE_CSV    = "Relevé BNP "
else: 
    PDFTOTEXT = 'pdftotext'
    PREFIXE_CSV    = "Relevé_BNP_"

class uneOperation:
    """Une opération bancaire = une date, un descriptif,
    une valeur de débit, une valeur de crédit et un interrupteur de validité"""

    def __init__(self, date="", desc="", value = "", debit="", credit=""):
        self.date   = date
        self.desc   = desc
        self.value  = value
        self.debit  = debit
        self.credit = credit
        self.valide = True
        if not len(self.date) >= 10 or int(self.date[:2]) > 31 or int(self.date[3:4]) > 12:
            self.valide = False
    def estRemplie(self):
        ## return len(self.date) >=10 and len(desc) > 0 and len(value) > 0 \
        ##    and (len(credit) > 0 or len(debit) > 0) 
        return len(self.date) >=10 and (len(self.credit) > 0 or len(self.debit) > 0) 

class UnReleve:
    """Un relevé de compte est une liste d'opérations bancaires
    sur une durée définie"""
    def __init__(self, nom="inconnu"):
        self.nom = nom
        self.liste = []

    def ajoute(self, Ope):
        """Ajoute une opération à la fin de la liste du relevé bancaire"""
        self.liste.append(Ope)

    def ajoute_from_TXT(self, fichier_txt, annee, mois):
        """Parse un fichier TXT pour en extraire les
        opérations bancaires et les mettre dans le relevé"""
        print('[txt->   ] Lecture    : '+fichier_txt)

        with open(fichier_txt) as file:
            # ignore les lignes avec les coordonnées et le blabla
            for ligne in file:
                monnaie = re.search('Monnaie du compte\s*: (\w*)', ligne)
                if monnaie:
                    monnaie = monnaie.group(1)
                    break
            
            Ope = uneOperation()
            date = ""
            operation = []
            num = -1
            Table = faux
            vide = 0

            for ligne in file:
                num = num+1
                if len(ligne) < 2:           # ligne vide, trait du tableau
                    vide = vide + 1
                    if vide > 2:
                        Table = Faux
                        if Ope.estRemplie():         # on ajoute la précédente 
                            self.ajoute(Ope)          # opération si elle est valide     
                            Ope = uneOperation()
                            date = ""
                            operation = []
                if Table==Faux
            
                print('{}({}): {}'.format(num, len(ligne), ligne))
                
                date_ou_pas = ligne[:12].split()  # premier caractères de la ligne (date?)
                if 1 == len(date_ou_pas):
                    date_ou_pas = pattern.split(date_ou_pas[0])
                                                
                dernier = ligne[-14::].split()    # derniers caractètres (valeur?)
                if 1 == len(dernier):
                    dernier = pattern.split(dernier[0])

                if estDate(date_ou_pas):          # est-ce une date
                    date_valeur = ligne[132:142].split() # il y a aussi une date valeur
                    if 1 == len(date_valeur):
                        date_valeur = pattern.split(date_valeur[0])
                    if Ope.estRemplie():          # on ajoute la précédente 
                        l_operation = ' '.join(operation)
                        Ope.desc = l_operation
                        self.ajoute(Ope)          # opération si elle est valide
                        Ope = uneOperation()
                    operation = []
                    date = date_ou_pas
                    
                operation.extend(ligne[12:64].split())   
    
                if date : # si on a deja trouvé une date
                    la_date     = list2date(date, annee, mois)
                    Ope.date = la_date
                if estArgent(dernier):
                    la_valeur   = list2valeur(dernier) 
                    if len(ligne) < 180:
                        Ope.credit = la_valeur;
                    else:
                        Ope.debit = la_valeur;
            if Ope.estRemplie():         # on ajoute la précédente 
                self.ajoute(Ope)          # opération si elle est valide     

    def genere_CSV(self, filename=""):
        """crée un fichier CSV qui contiendra les opérations du relevé
        si ce CSV n'existe pas deja"""
        if filename == "":
            filename = self.nom
        filename = filename + ".csv"
        if not filename in deja_en_csv:
            print('[   ->csv] Export     : '+filename)
            with open(filename, "w") as file:
                file.write("Date"+CSV_SEP+"Opération"+CSV_SEP+"Débit"+CSV_SEP+"Crédit\n")
                for Ope in self.liste:
                    file.write(Ope.date+CSV_SEP+Ope.desc+CSV_SEP+Ope.debit+CSV_SEP+Ope.credit+"\n")
                file.close()

def extraction_PDF(pdf_file, deja_en_txt, temp):
    """Lit un relevé PDF et le convertit en fichier TXT du même nom
    s'il n'existe pas deja"""
    txt_file = pdf_file[:-3]+"txt"
    if not txt_file in deja_en_txt:
        print('[pdf->txt] Conversion : '+pdf_file)
        subprocess.call([PDFTOTEXT, '-layout', pdf_file, txt_file])
        temp.append(txt_file)

def estDate(liste):
    """ Attend un format ['JJ', '.' 'MM']"""
    if len(liste) != 3:
        return False
    if len(liste[0]) ==2 and liste[1] == '.' and len(liste[2]) == 2:
        return True
    return False

def estArgent(liste):
    """ Attend un format ['[0-9]*', ',', '[0-9][0-9]'] """
    if len(liste) < 3:
        return False
    if liste[-2] == ',':
        return True
    return False

def list2date(liste, annee, mois):
    """renvoie un string"""
    if mois == '01' and liste[2] == '12':
        return liste[0]+'/'+liste[2]+'/'+str(int(annee)-1)
    else:
        return liste[0]+'/'+liste[2]+'/'+annee

def list2valeur(liste):
    """renvoie un string"""
    liste_ok = [x for x in liste if x != '.']
    return "".join(liste_ok)

def filtrer(liste, filetype):
    """Renvoie les fichiers qui correspondent à l'estension donnée en paramètre"""
    files = [fich for fich in liste if str.lower(fich[-3::])==filetype]
    return files

def mois_dispos(liste):
    """Renvoie une liste des relevés disponibles de la forme
    [['2012', '10', '11', '12']['2013', '01', '02', '03', '04']]"""
    liste_tout = []
    les_annees = []
    for releve in liste:
        if releve[:len(PREFIXE_COMPTE)] == PREFIXE_COMPTE:
            annee = releve[len(PREFIXE_COMPTE):len(PREFIXE_COMPTE)+4]
            mois  = releve[len(PREFIXE_COMPTE)+4:len(PREFIXE_COMPTE)+6]
            if not annee in les_annees:
                les_annees.append(annee)
                liste_annee = [annee, mois]
                liste_tout.append(liste_annee)
            else:
                liste_tout[les_annees.index(annee)].append(mois)
    return liste_tout

# fonction inutilisée
def est_dispo(annee, mois, liste):
    """Verifie si le relevé de ce mois/année est disponible
    dans la liste donnée"""
    for annee_de_liste in liste:
        if annee == annee_de_liste[0]:
            if mois in annee_de_liste:
                return True
    return False

def affiche(liste):
    """Affiche à l'écran les mois dont les relevés sont disponibles"""
    print("Relevés disponibles:")
    for annee in liste:
        ligne_12 = ['  ']*12
        for i in annee:
            if len(i) == 2:
                ligne_12[int(i)-1] = i
        ligne = annee[0]+': '+' '.join(ligne_12)
        print(ligne)
    print("")

# On demarre ici

print('\n******************************************************')
print('*   Convertisseur de relevés bancaires BNP Paribas   *')
print('********************  PDF -> CSV  ********************\n')
chemin=os.getcwd()
fichiers = os.listdir(chemin)

if shutil.which(PDFTOTEXT) is None:
    print("Fichier {} absent !".format(PDFTOTEXT))
    input("Bye bye :(")
    exit()
mes_pdfs = filtrer(fichiers, 'pdf')
deja_en_txt = filtrer(fichiers, 'txt')
deja_en_csv = filtrer(fichiers, 'csv')

mes_mois_disponibles = mois_dispos(mes_pdfs)
mes_mois_deja_en_txt = mois_dispos(deja_en_txt)

if len(mes_mois_disponibles) == 0:
    print("Il n'y a pas de relevés de compte en PDF dans ce répertoire")
    print("correspondant au préfixe "+PREFIXE_COMPTE)
    print("\nIl faut placer les fichiers convertBNP.py et pdftotext.exe")
    print("à côté des fichiers de relevé de compte en PDF et adapter")
    print("la ligne 18 (PREFIXE_COMPTE = XXXXX) du fichier convertBNP.py")
    print("pour la faire correspondre à votre numéro de compte.\n")
    input("Bye bye :(")
    exit()
affiche(mes_mois_disponibles)

touch = 0
temp_list = []

# on convertit tous les nouveaux relevés PDF en TXT sauf si CSV deja dispo
for releve in mes_pdfs:
    if releve[:len(PREFIXE_COMPTE)] == PREFIXE_COMPTE:
        annee = releve[len(PREFIXE_COMPTE):len(PREFIXE_COMPTE)+4]
        mois  = releve[len(PREFIXE_COMPTE)+4:len(PREFIXE_COMPTE)+6]
        csv = PREFIXE_CSV+annee+'-'+mois+".csv"
        if not csv in deja_en_csv:
            touch = touch + 1
            extraction_PDF(releve, deja_en_txt, temp_list)
if touch != 0:
    print("")

# on remet à jour la liste de TXT
fichiers = os.listdir(chemin)
deja_en_txt = filtrer(fichiers, 'txt')
mes_mois_deja_en_txt = mois_dispos(deja_en_txt)

# on convertit tous les nouveaux TXT en CSV
for txt in deja_en_txt:
    if txt[:len(PREFIXE_COMPTE)] == PREFIXE_COMPTE:
        annee = txt[len(PREFIXE_COMPTE):len(PREFIXE_COMPTE)+4]
        mois  = txt[len(PREFIXE_COMPTE)+4:len(PREFIXE_COMPTE)+6]
        csv = PREFIXE_CSV+annee+'-'+mois+".csv"
        if not csv in deja_en_csv:
            releve = UnReleve()
            releve.ajoute_from_TXT(txt, annee, mois)
            releve.genere_CSV(PREFIXE_CSV+annee+'-'+mois)

# on efface les fichiers TXT
if len(temp_list) :
    print"[txt-> x ] Nettoyage\n")
    for txt in temp_list:
        os.remove(txt)

if touch == 0:
    input("Pas de nouveau relevé. Bye bye.")
else:
    print(str(touch)+" relevés de comptes convertis.")
    input("Terminé. Bye bye.")

# EOF
