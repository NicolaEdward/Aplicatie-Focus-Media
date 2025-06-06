# Aplicatie Focus Media

Aplicatia foloseste in mod implicit o baza de date MySQL. Conectarea se face
citind variabilele de mediu:

```
MYSQL_HOST      - adresa serverului MySQL
MYSQL_PORT      - portul (implicit 3306)
MYSQL_USER      - utilizatorul
MYSQL_PASSWORD  - parola
MYSQL_DATABASE  - baza de date
```

Daca `MYSQL_HOST` nu este definit, aplicatia revine la fisierul local
`locatii.db` cu SQLite (util pentru teste sau dezvoltare fara un server
dedicat).

Dependintele necesare se instaleaza cu:

```
pip install -r requirements.txt
```


## Migrarea bazei de date SQLite la MySQL

Dupa configurarea variabilelor de mediu pentru MySQL, executa:

```bash
python migrate_to_mysql.py
```

Scriptul va copia in MySQL toate tabelele si datele din fisierul `locatii.db`.

## Autentificare

La prima rulare este creat automat contul `admin` cu parola `admin`. După
autentificare, administratorul poate adăuga alte conturi din fereastra de
administrare a utilizatorilor. Vânzătorii pot adăuga clienți și pot închiria
locații, dar nu pot modifica sau șterge locațiile existente.

Funcția "Raport Vânzători" generează un Excel cu totalul contractelor pe lună
pentru fiecare utilizator.


