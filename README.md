# convertBNP

## Description

Convertisseur de relevés de comptes BNP du format PDF vers le formats
CSV et XLSX.

Ce script Python lit les relevés bancaires fournis par la banque BNP
Paribas en PDF pour en générer des fichiers tableurs. Le format CSV
(fichiers textes avec valeur séparés par des points-virgules) est
reconnu par excel ou n'importe quel autre tableur. Les format XLSX ne
requiert pas de conversion des données.

Le script utilise le fichier pdftotext.exe en version 3.03 issu de
l'archive xpdf (xpdf est opensource et gratuit, sous licence GPL2).
Il ne fonctionne pas avec la version 3.04.

## Versions

Il y a trois variantes :

* *convertBNP* : script original créé le 10/11/2013 pour python3 sur
   Windows, suppose une mise en forme bien précise, utilisée autour de
   2013.
* *convertBNP_4col.py* : créé à la demande d'un utilisateur, sépare les
  débits et crédits en 2 colonnes distinctes. Mêmes restrictions de mise
  en forme.
* *convertBNP_5col.py* : ajoute une date des opérations, si elle est
  différente de la date du mouvement. Compatible avec les mises en
  formes utilisées de 2012 à 2018 (et au-delà ...).
* les fichiers de sortie sont de la forme *'Relevé BNP YYYY-MM'* sous
  Windows et *'Relevé_BNP_YYYY-MM'* ailleurs.

## Installation dans un environnement virtuel

Cette méthode est particulièrement utile quand il y a plusieurs
versions de Python installées, ou que l'on ne peut/veut pas modifier
les modules installés au niveau du système.

1. Pré-requis : python3 et un répertoire de travail généré à partir
des sources GitHub.

2. Créer un environnement virtuel :

         2.1 en ligne de commande, à partir de ce répertoire :
   
                $ python3 -mvenv .
           
         2.2 en mode graphique : créer un environnement virtuel, choisir comme répertoire de 
         destination celui créé à partir de GitHub

3. activer cet environnement et installer le module python "XlsxWriter" via pip :

           $ source bin/activate
           $ pip3 install XslxWriter

4. le script comprend à présent 3 arguments :

      - --verbosity : 1 pour afficher chaque ligne analysée, 2 pour
              entrer en mode pas-à-pas à certains endroits critiques
              du traitement

      - --prefixe : une sous-chaine qui se trouve dans les noms des fichiers
   pdf à analyser. Le préfixe peut aussi être lu via un fichier
   "prefixe_compte.txt" dans le même répertoire que le script python,
   dont le contenu, en une seule ligne, correspond à la partie fixe
   des relevés :

                        RCHQ_101_300040012300001234567

      - --dir : un spécificateur relatif ou absolu du répertoire
    contenant les fichiers à transformer. Les fichiers générés seront
    stockés dans le même répertoire.

## Installation (méthode originale)
1. Installer Python 3.x.x
2. Extraire pdftotext.exe et convertBNP.py dans le répertoire des relevés de compte PDF.

3. créer un fichier "prefixe_compte.txt" avec le Bloc-Notes, en y
   mettant votre numéro de compte (voir le nom de vos fichiers PDF).
   cette ligne sert à identifier les fichiers du compte bancaire à
   convertir. Exemple de contenu en une seule ligne :

            RCHQ_101_300040012300001234567

4. Lancer le script en double-cliquant sur le fichier convertBNP.py
   ou avec la ligne de commande suivante (il faut être dans le dossier) :
   "python convertBNP.py"

## Changements majeurs
- au fil du temps (2012 -- 2018), il y a eu des changements de mise en
  forme dans les fichiers pdf.  Ce script identifie à présent la
  pagination pour aller chercher les bons champs aux bons endroits.
  
- les fichiers sont exportés au format csv et xlsx (Exell / Libreoffice).

- Prend en compte les "locales" pour déterminer le type de séparateur
  décimal.

- Génère une colonne complémentaire avec la date du mouvement. Par
  exemple, pour un retrait par carte, on peut avoir une date pour le
  retrait (un samedi), une date pour l'opération (le lundi suivant) et
  une date valeur (en fin du mois).

- La somme des mouvements de type débit et crédit est comparée aux
  données en fin de tableau. En cas de différence, cela est considéré
  comme une condition d'erreur et il n'y a pas de fichier de sortie
  généré.

## Utilisation

(en ligne de commande, environnement "convertBNP")

          $ (convertBNP) python3  convertBNP_5col.py --verbosity 1 --dir "~/Document/Comptes/BNP"

(dans emacs, avec gud  et deboguage : cliquer sur "Python", "debugger")

             Run pdb (like this): python3 -mpdb convertBNP_5col.py --dir "../BNP FR"

## Script en action (exemple)

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
 
