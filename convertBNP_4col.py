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

PREFIXE_COMPTE = "RCHQ_101_300040012300001234567_"
PREFIXE_COMPTE = "300040181600000906809"

CSV_SEP        = ";"
deja_en_csv    = ""
deja_en_xlsx   = ""

import pdb

import argparse, os, re, subprocess, shutil, sys 
import xlsxwriter

pattern = re.compile('(\W)')

from locale import *
setlocale(LC_NUMERIC, '')
from datetime import datetime as dt

if os.name == 'nt':
    PDFTOTEXT = 'pdftotext.exe'
    PREFIXE_CSV    = "Relevé BNP "
else: 
    PDFTOTEXT = 'pdftotext'
    PREFIXE_CSV    = "Relevé_BNP_"

class uneOperation:
    """Une opération bancaire = une date, un descriptif,
    une valeur de débit, une valeur de crédit et un interrupteur de validité"""

    def __init__(self, date="", desc="", value = "", debit=0.0, credit=0.0):
        self.date   = date
        self.date_valeur = ""
        self.date_oper = ""
        self.desc   = desc
        self.value  = value
        self.debit  = debit
        self.credit = credit
        self.valide = True
        if not len(self.date) >= 10 or int(self.date[:2]) > 31 or int(self.date[3:4]) > 12:
            self.valide = False

    def estRemplie(self, operation=[]):
        ## return len(self.date) >=10 and len(desc) > 0 and len(value) > 0 \
        ##    and (len(credit) > 0 or len(debit) > 0)
        resu = len(self.date) >=10 and (self.credit > 0.0 or self.debit > 0.0) 
        if resu:
            # this one is OK, fill its description
            if len(operation):
                if 0 == len(self.desc):
                    self.desc = ' '.join(operation)
                    for num, date in enumerate(operation):
                        try:
                            date_oper = dt.strptime(date, '%d/%m/%y').strftime('%d/%m/%Y')
                            self.date_oper = date_oper
                            break;
                        except ValueError as e:
                            continue                       
        return resu

class UnReleve:
    """Un relevé de compte est une liste d'opérations bancaires
    sur une durée définie"""
    def __init__(self, nom="inconnu"):
        self.nom = nom
        self.liste = []

    def ajoute(self, Ope):
        """Ajoute une opération à la fin de la liste du relevé bancaire"""
        self.liste.append(Ope)

    def ajoute_from_TXT(self, fichier_txt, annee, mois, verbosity=False):
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
            # à présent, en-tête et SOLDE  / Date /valeur        
            for ligne in file:
                if re.search('Date\s*Nature\s*des\*', ligne, re.IGNORECASE): continue
                if re.search('Solde\s*', ligne, re.IGNORECASE): break
            
            operation = ligne.split()
            for num, date in enumerate(operation):
                try:
                    basedate = dt.strptime(date, '%d.%m.%Y').strftime('%d/%m/%Y')
                    break;
                except ValueError as e:
                    continue
            
            # montant ?
            la_valeur = atof(''.join(operation[num+1:]))
            ligne = ' '.join(operation[:num+1])
            
            # dans quel sens ?        
            if re.match('crediteur', operation[1], re.IGNORECASE):
                Ope = uneOperation(basedate, ligne, "", 0.0, la_valeur)
                solde_init = la_valeur
            elif re.match('debiteur', operation[1], re.IGNORECASE):
                Ope = uneOperation(basedate, ligne, "", la_valeur, 0.0)
                solde_init = -la_valeur
            else:
                raise ValueError(ligne+"ne peut pas être interprétée")

            # crée une entrée avec le solde initial
            self.ajoute(Ope)    
            if verbosity:    
                print('{}({}): {}'.format(num, len(ligne), ligne))
                print('Solde initial: {}  au {}'.format(solde_init, basedate))
                print('Date:{} -- desc:{} -- debit {} -- credit {}'.format(Ope.date, Ope.desc,  
                                                                           Ope.debit, Ope.credit))      
            Ope = uneOperation()
            date = ""
            operation = []
            la_date = ""
            num = -1
            Table = False
            vide = 0
            somme_cred = 0.0 # To check sum of cred
            somme_deb = 0.0  # To check sum of deb
            pdb.set_trace()
            for ligne in file:
                num = num+1
                if len(ligne) < 2:           # ligne vide, trait du tableau
                    vide = vide + 1
                    if vide > 2:
                        Table = False
                        if len(operation) > 0:
                            if Ope.estRemplie(operation): # on ajoute la précédente 
                                self.ajoute(Ope)          # opération si elle est valide     
                                if verbosity:    
                                    print('Date:{} -- desc:{} -- debit {} -- credit {}'.format(Ope.date, Ope.desc,  
                                                                                               Ope.debit, Ope.credit))      
                                Ope = uneOperation()
                                date = ""
                                la_date = ""
                                operation = []
                    continue
                
                # this line ends the table
                if re.match('.*?total des montants\s', ligne, re.IGNORECASE):
                    if verbosity:    
                        print('{}({}): {}'.format(num, len(ligne), ligne))
                    break;
                if re.match('.*?total des operations\s', ligne, re.IGNORECASE):
                    if verbosity:    
                        print('{}({}): {}'.format(num, len(ligne), ligne))
                    break;

                if Table == False:
                    pass
                
                if verbosity:    
                    print('{}({}): {}'.format(num, len(ligne), ligne))
                
                date_ou_pas = ligne[:12].split()  # premier caractères de la ligne (date?)
                if 1 == len(date_ou_pas):
                    date_ou_pas = pattern.split(date_ou_pas[0])

                # si une ligne se termine par un montant, il faut l'extraire pour qu'il reste la
                # date valeur
                dernier = ligne[-22::].split()    # derniers caractères (valeur?)
                if 1 == len(dernier):
                    dernier = pattern.split(dernier[0])

                if estArgent(dernier):
                    # si l'operation précédente est complète, on la sauve
                    if Ope.estRemplie(operation): 
                        self.ajoute(Ope)          # opération si elle est valide
                        if verbosity:                                                       
                            print('Date:{} -- desc:{} -- debit {} -- credit {}'.format(Ope.date, Ope.desc,  
                                                                                       Ope.debit, Ope.credit))
                        Ope = uneOperation()
                        operation = []                # we are on a new op
                    la_valeur   = list2valeur(dernier)
                    try:
                        if len(ligne) < 180:
                            Ope.debit = atof(la_valeur)
                            somme_deb += Ope.debit
                        else:
                            Ope.credit = atof(la_valeur)
                            somme_cred += Ope.credit
                    except ValueError as e:
                        print('Failed to convert {} to a float: {}'.format(la_valeur, e))
                    previous_ligne = ligne    
                    ligne = ligne[:145] # truncate the money amount
                    ligne = ligne.rstrip()
                
                if estDate(date_ou_pas):          # est-ce une date
                    date_valeur = ligne[85:90].split() # il y a aussi une date valeur
                    if 1 == len(date_valeur):
                        date_valeur = pattern.split(date_valeur[0])
                    if Ope.estRemplie(operation):          # on ajoute la précédente 
                        self.ajoute(Ope)          # opération si elle est valide
                        if verbosity:    
                            print('Date:{} -- desc:{} -- debit {} -- credit {}'.format(Ope.date, Ope.desc,  
                                                                                       Ope.debit, Ope.credit))      
                        Ope = uneOperation()

                    operation = []                # we are on a new op
                    la_date =  ''
                    date = date_ou_pas
                    
                operation.extend(ligne[12:64].split())   
    
                if date : # si on a deja trouvé une date
                    la_date     = list2date(date, annee, mois)
                    date = ""
                    if (len(date_valeur) < 3):
                        # there are left and right pages
                        date_valeur = ligne[109:114].split() # il y a aussi une date valeur
                        if 1 == len(date_valeur):
                            date_valeur = pattern.split(date_valeur[0])
                            
                    if (len(date_valeur) < 3):
                        print('line 223');
                        print(ligne)
                        print(ligne[85:91])
                        print(ligne[109:114])
                        print(date_valeur);

                    la_date_valeur  = list2date(date_valeur, annee, mois)
                    Ope.date = la_date
                    Ope.date_valeur = la_date_valeur
                
            # end of main table                
            if Ope.estRemplie(operation):         # on ajoute la précédente 
                if verbosity:    
                    print('Date:{} -- desc:{} -- debit {} -- credit {}'.format(Ope.date, Ope.desc,  
                                                                               Ope.debit, Ope.credit))      
                self.ajoute(Ope)          # opération si elle est valide     
        
            if verbosity:    
                print('Exited main loop')
                print('{}({}): {}'.format(num, len(ligne), ligne))

            operation = ligne.split(); 
            start = 3
            count = 4
            for elem in operation[count:]:
                if ',' in elem:
                    break
                count = count + 1
            le_debit = ''.join(operation[start:count+1])
            start = count + 1
            count = start + 1
            for elem in operation[count:]:
                if ',' in elem:
                   break
                count = count + 1
            le_credit = ''.join(operation[start:count+1])
            try:
                le_debit = atof(le_debit)
            except ValueError as e:
                print('Failed to convert {} to a float: {}'.format(le_debit, e))
            try:
                le_credit = atof(le_credit)
            except ValueError as e:
                print('Failed to convert {} to a float: {}'.format(le_debit, e))
            
            # here, we have "solde .. au
            for ligne in file:
                if re.search('Solde\s*', ligne, re.IGNORECASE): 
                    break
            
            operation = ligne.split()
            for num, date in enumerate(operation):
                try:
                    basedate = dt.strptime(date, '%d.%m.%Y').strftime('%d/%m/%Y')
                    break;
                except ValueError as e:
                    continue
            
            # montant ?
            try:
                la_valeur = atof(''.join(operation[num+1:]))
                ligne = ' '.join(operation[:num+1])
            except ValueError as e:
                print(operation)
                pdb.set_trace()

            # dans quel sens ?        
            if re.match('crediteur', operation[1], re.IGNORECASE):
                Ope = uneOperation(basedate, ligne, "", 0.0, la_valeur)
                solde_final = la_valeur
            elif re.match('debiteur', operation[1], re.IGNORECASE):
                Ope = uneOperation(basedate, ligne, "", la_valeur, 0.0)
                solde_final = -la_valeur
            else:
                raise ValueError(ligne+" ne peut pas être interprétée")
            # check that solde_deb = le_debit;
            if abs(somme_deb  - le_debit) > .01:
                if verbosity: 
                    print("La somme des débits {} n'est pas égale au débit totat {}".format(somme_deb, le_debit))         
                else:
                    raise ValueError('La somme des débits {} n''est pas égale au débit total {}'.format(somme_deb, le_debit))

            # check that solde_cred = le_credit;
            if abs(somme_cred  - le_credit) > .01:
                if verbosity: 
                    print("La somme des crédits {} n'est pas égale au crédit total {}".format(somme_cred, le_credit))               
                else:                    
                    raise ValueError('La somme des crédits {} n''est pas égale au crédit totat {}'.format(somme_cred, le_credit))                    
            # check that solde_init - le_credit + le_debit == solde_final
            mouvements = solde_init - le_debit + le_credit
            if abs(solde_final - mouvements) > .01:
                if verbosity: 
                    print("La somme des mouvements {} n'arrive pas au solde final {}".format(mouvements, solde_final))      
                else:      
                    raise ValueError('La somme des mouvements n''arrive pas au solde final {}'.format(mouvement, solde_final))
            
            # duplicate the current operation
            OpeTot = uneOperation(basedate, "TOTAL DES MONTANTS", "", le_debit, le_credit)
            OpeTot.desc = "TOTAL DES MONTANTS"
            OpeTot.debit = le_debit
            OpeTot.credit = le_credit
            # dump it
            self.ajoute(OpeTot)
            
            # crée une entrée avec le solde final   
            self.ajoute(Ope)    
            if verbosity:    
                print('{}({}): {}'.format(num, len(ligne), ligne))
                print('Solde final: {}  au {}'.format(solde_final, basedate))
                

    def genere_CSV(self, filename=""):
        """crée un fichier CSV qui contiendra les opérations du relevé
        si ce CSV n'existe pas deja"""
        if filename == "":
            filename = self.nom
        filename_csv = filename + ".csv"
        if not filename_csv in deja_en_csv:
            print('[   ->csv] Export     : '+filename_csv)
            with open(filename_csv, "w") as file:
                file.write("Date"+CSV_SEP+"Date_Valeur"+CSV_SEP+"Date_Oper"+CSV_SEP+"Débit"+CSV_SEP+"Crédit"+CSV_SEP+"Opération\n")
                for Ope in self.liste:
                    file.write(Ope.date+CSV_SEP+Ope.date_valeur+CSV_SEP+Ope.date_oper+CSV_SEP+str(Ope.debit)+CSV_SEP+str(Ope.credit)+CSV_SEP+Ope.desc+"\n")
                file.close()
        filename_xlsx = filename + ".xlsx"
        if not filename_xlsx in deja_en_xlsx:
            print('[   ->xlsx] Export     : '+filename_xlsx)
            pdb.set_trace()
            workbook = xlsxwriter.Workbook(filename_xlsx)
            worksheet = workbook.add_worksheet()
            worksheet.set_column(0, 2, 12)
            worksheet.set_column(3, 4, 10)
            worksheet.set_column(5, 5, 64)
            cell_format = workbook.add_format({'bold': True})
            worksheet.write_row('A1', ["Date", "Date_Valeur", "Date_Oper", "Débit", "Crédit", "Opération"], cell_format);
            row = 1;
            for Ope in self.liste[:-2]:
                worksheet.write_row(row, 0, [Ope.date, Ope.date_valeur, Ope.date_oper, Ope.debit, Ope.credit, Ope.desc])
                row = row + 1
            # generate a control formula
            Ope = self.liste[-2]
            worksheet.write(row, 0, Ope.date)
            worksheet.write(row, 5, 'Somme de contrôle')
            # EXELL formula are stored in english but displayed in locale
            worksheet.write_formula(row, 3, '=SUM(D3:D'+str(row)+')')
            worksheet.write_formula(row, 4, '=SUM(E3:E'+str(row)+')')
            row = row + 1
            for Ope in self.liste[-2:]:
                worksheet.write_row(row, 0, [Ope.date, Ope.date_valeur, Ope.date_oper, Ope.debit, Ope.credit, Ope.desc])
                row = row + 1
            workbook.close()
            
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
    if (len(liste) < 3):
        print ('line 297')
        print (liste)
        
    if mois == '01' and liste[2] == '12':
        return liste[0]+'/'+liste[2]+'/'+str(int(annee)-1)
    else:
        return liste[0]+'/'+liste[2]+'/'+annee

def list2valeur(liste):
    """renvoie un string"""
    liste_ok = [x for x in liste if x != '.']
    return "".join(liste_ok)

def filtrer(liste, filetype):
    """Renvoie les fichiers qui correspondent à l'extension donnée en paramètre"""
    files = [fich for fich in liste if fich.split('.')[-1].lower() == filetype]
    return files

def mois_dispos(liste):
    """Renvoie une liste des relevés disponibles de la forme
    [['2012', '10', '11', '12']['2013', '01', '02', '03', '04']]"""
    liste_tout = []
    les_annees = []

    for releve in liste:
        operation = releve.split('.') # strip the extension
        operation = operation[0].split('_') # get chunks
        if "FRAIS" in operation[0]:
            continue
        for num, val in enumerate(operation[1:]):
            if PREFIXE_COMPTE not in val:
                continue
            # the next element contains the date 
            annee = operation[num+2][0:4]
            mois  = operation[num+2][4:6]
            if not annee in les_annees:
                les_annees.append(annee)
                liste_annee = [annee, mois]
                liste_tout.append(liste_annee)
            else:
                liste_tout[les_annees.index(annee)].append(mois)
    liste_tout.sort()            
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
def main(*args, **kwargs):
    print('\n******************************************************')
    print('*   Convertisseur de relevés bancaires BNP Paribas   *')
    print('********************  PDF -> CSV  ********************\n')

    parser = argparse.ArgumentParser()
    parser.add_argument("--verbosity", help="increase output verbosity")
    parser.add_argument("--prefixe", help="prefixe des fichiers à traiter")
    myargs = parser.parse_args()

    # if myargs.prefixe:
    #     PREFIXE_COMPTE = myargs.prefixe
    # else:
    #     PREFIXE_COMPTE = CSV_SEP
        
    chemin = os.getcwd()
    fichiers = os.listdir(chemin)
    
    if shutil.which(PDFTOTEXT) is None:
        print("Fichier {} absent !".format(PDFTOTEXT))
        input("Bye bye :(")
        exit()

    mes_pdfs = filtrer(fichiers, 'pdf')
    deja_en_txt = filtrer(fichiers, 'txt')
    deja_en_csv = filtrer(fichiers, 'csv')
    deja_en_xlsx = filtrer(fichiers, 'xlsx')

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
            datefichier = releve.split('_')[-2]
            annee = datefichier[:-4]
            mois  = datefichier[-4:-2]
            csv = PREFIXE_CSV+annee+'-'+mois+".csv"
            xlsx= PREFIXE_CSV+annee+'-'+mois+".xlsx"
            if not csv in deja_en_csv:
                touch = touch + 1
                extraction_PDF(releve, deja_en_txt, temp_list)
            elif not xlsx in deja_en_xlsx:
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
            datefichier = txt.split('_')[-2]
            annee = datefichier[:-4]
            mois  = datefichier[-4:-2]
            csv = PREFIXE_CSV+annee+'-'+mois+".csv"
            xlsx= PREFIXE_CSV+annee+'-'+mois+".xlsx"
            if not csv in deja_en_csv:
                releve = UnReleve()
                releve.ajoute_from_TXT(txt, annee, mois, myargs.verbosity)
                releve.genere_CSV(PREFIXE_CSV+annee+'-'+mois)
            elif not xlsx in deja_en_xlsx:    
                releve = UnReleve()
                releve.ajoute_from_TXT(txt, annee, mois, myargs.verbosity)
                releve.genere_CSV(PREFIXE_CSV+annee+'-'+mois)
                
    # on efface les fichiers TXT
    if len(temp_list) :
        print("[txt-> x ] Nettoyage\n")
        for txt in temp_list:
            os.remove(txt)

    if touch == 0:
        input("Pas de nouveau relevé. Bye bye.")
    else:
        print(str(touch)+" relevés de comptes convertis.")
        input("Terminé. Bye bye.")
        
    # EOF

    return 0

if __name__== "__main__": 
    sys.exit(main(sys.argv[1:]))
