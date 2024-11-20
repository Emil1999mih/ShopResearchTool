Overview

Program-ul "Shop Research Tool" ofera utilizatorilor o metoda simpla si eficienta de a descarca si analiza produsele unui magazin construit pe platforma Shopify.
Acesta extrage informatiile produsului si le salveaza intr-un format Excel si face cateva analize de pret simple. 

Facilitati

Webscraping folosind requests din magazine Shopify
Salvare in Excel in mai multe foi de calcul 
Analiza a distribuiri de pret ( cate produse sunt anumite categorii de pret)
Vizualizare grafica a distribuirii de pret
interfata simpla si eficienta

Conditii de instalare:

Python 3.0 sau mai nou 
Librariile: - openpyxl
            - requests
            - tkinter

Instalare:

1. Se cloneaza local acest repository
2. Se instaleaza librariile mentionate mai sus

Utilizare:

1. Rulati fisierul main.py
2. Introduceti adresa URL (https://magazinultau.ro)
3. Alegeti locatia si numele  dorit pentru salvarea fisierului
4. Apasati "Save Data to Excel"

Structura fisierelor:

main.py: este punctul de pornire al aplicatiei si verifica daca pachetele necesare au fost instalate 
Interface.py: este implementarea GUI
skimmer.py: modulul ce realizeaza scraping-ul website-ului
excelbuilder.py: modulul ce genereaza fisierul excel si analiza datelor

Metode prin care se gestioneaza erorilor:

  1. In main.py programul verifica daca pachetele  necesare rularii  sunt instalate
  2. In main.py se verifica deasemenea daca sunt fisiere lipsa
  3. Interfata GUI produce mesajul erorilor  daca este cazul

Limitarile programului:

Programul este limitat la website-urile ce stocheaza informatiile despre produse in products.json ( https://magazinultau.com/products.json)


Thanks!
