convertBNP
==========

Convertisseur de relevés de comptes BNP du format PDF vers le format CSV

Ce script python convertBNP.py lit les relevés bancaires fournis par la
banque BNP Paribas en PDF dans le répertoire courant pour en générer des
fichiers au format CSV (fichiers textes avec valeur séparés par des points-virgules)
Les fichiers CSV sont ouvrables par excel ou n'importe quel autre tableur.

Le script utilise le fichier pdftotext.exe  en version 3.03 issu de l'archive xpdf
(xpdf est opensource et gratuit, sous licence GPL2).
Il ne fonctionne pas avec la version 3.04.

Script créé le 10/11/2013 pour python3 sur windows

A la demande d'un utilisateur, j'ai créé une variante qui sépare les débits
et crédits en 2 colonnes distinctes : convertBNP_4col.py


Installation
------------

1. Installer Python 3.x.x
2. Extraire pdftotext.exe et convertBNP.py dans le répertoire des relevés de compte PDF
3. Ouvrir le fichier convertBNP.py avec le Bloc-Notes
   Modifier la ligne contenant :
   PREFIXE_COMPTE = "RCHQ_101_300040012300001234567_"
   en y mettant votre numéro de compte (voir le nom de vos fichiers PDF)
   cette ligne sert à identifier les fichiers du compte bancaire à convertir.
4. Lancer le script en double-cliquant sur le fichier convertBNP.py
   ou avec la ligne de commande suivante (il faut être dans le dossier) :
   "python convertBNP.py"


Script en action (exemple)
--------------------------

    ******************************************************
    *   Convertisseur de relevés bancaires BNP Paribas   *
    ********************  PDF -> CSV  ********************

    Relevés disponibles:
    2011:             05 06 07 08 09 10 11 12
    2012: 01 02 03 04 05 06 07 08 09 10 11 12
    2013: 01 02 03 04 05 06 07 08 09 10 11 12
    2014: 01 02 03 04 05 06 07 08 09 10 11 12

    [pdf->txt] Conversion : RCHQ_101_300040012300001234567_20140926_2226.pdf
    [pdf->txt] Conversion : RCHQ_101_300040012300001234567_20141026_2239.pdf
    [pdf->txt] Conversion : RCHQ_101_300040012300001234567_20141126_2218.pdf
    [pdf->txt] Conversion : RCHQ_101_300040012300001234567_20141226_2224.pdf

    [txt->   ] Lecture    : RCHQ_101_300040012300001234567_20140926_2226.txt
    [   ->csv] Export     : Relevé BNP 2014-09.csv
    [txt->   ] Lecture    : RCHQ_101_300040012300001234567_20141026_2239.txt
    [   ->csv] Export     : Relevé BNP 2014-10.csv
    [txt->   ] Lecture    : RCHQ_101_300040012300001234567_20141126_2218.txt
    [   ->csv] Export     : Relevé BNP 2014-11.csv
    [txt->   ] Lecture    : RCHQ_101_300040012300001234567_20141226_2224.txt
    [   ->csv] Export     : Relevé BNP 2014-12.csv
    [txt-> x ] Nettoyage

    4 relevés de comptes convertis.
    Terminé.
 