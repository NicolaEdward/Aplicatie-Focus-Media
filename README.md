# Aplicatie Focus Media

Aplicatia foloseste in mod implicit o baza de date MySQL. Conectarea se face
citind variabilele de mediu:

```
MYSQL_HOST      - adresa serverului MySQL (poate include portul, ex. `host:3306`)
MYSQL_PORT      - portul serverului MySQL
MYSQL_USER      - utilizatorul
MYSQL_PASSWORD  - parola
MYSQL_DATABASE  - baza de date
```

Poti crea un fisier `.env` in radacina proiectului cu aceste variabile. Un
exemplu este disponibil in fisierul `.env.example`:

```
MYSQL_HOST=127.0.0.1
MYSQL_PORT=3306
MYSQL_USER=utilizator
MYSQL_PASSWORD=parola
MYSQL_DATABASE=aplicatie_vanzari
```

Daca providerul iti ofera adresa impreuna cu portul (ex. `example.com:1234`),
poti pune aceasta valoare direct in `MYSQL_HOST` si lasa `MYSQL_PORT` necompletat.

Pentru un server MySQL gazduit la distanta seteaza `MYSQL_HOST` la adresa
respectiva si completeaza `MYSQL_USER`, `MYSQL_PASSWORD` si `MYSQL_DATABASE`
cu datele oferite de providerul tau. Aplicatia va folosi aceste informatii la
fiecare pornire.


Aplicatia necesita un server MySQL configurat cu variabilele de mediu de mai sus.
Fisierul `locatii.db` este folosit doar pentru teste sau pentru migrarea
initiala catre MySQL.

Dependintele necesare se instaleaza cu:

```
pip install -r requirements.txt
```


## Migrarea bazei de date SQLite la MySQL

Dupa ce ai completat fisierul `.env` cu datele serverului MySQL, executa:

```bash
python migrate_to_mysql.py
```

Scriptul va copia in MySQL toate tabelele si datele din fisierul `locatii.db`.
Daca tabelele din MySQL contin deja date, acestea vor fi sterse inainte de
import pentru a evita erorile legate de chei primare duplicate.

In cazul unei baze de date gazduite de Aiven, dupa rularea acestui script
toate datele din `locatii.db` vor fi transferate in serviciul online si
aplicatia va folosi exclusiv conexiunea configurata in `.env`.

In cazul in care apare mesajul de eroare "Access denied for user", verifica
fișierul `.env` sau variabilele de mediu folosite la conectare. Parola sau
utilizatorul MySQL trebuie să corespundă setărilor serverului. Poți porni de la
exemplul din `.env.example` și să îl adaptezi pentru sistemul tău.

## Autentificare

La prima rulare este creat automat contul `admin` cu parola `admin`. Parolele
sunt salvate folosind un hash PBKDF2 cu sare aleatorie pentru o securitate
suplimentară. După autentificare, administratorul poate adăuga alte conturi din
fereastra de administrare a utilizatorilor. Vânzătorii pot adăuga clienți și pot
închiria locații, dar nu pot modifica sau șterge locațiile existente.

Funcția "Raport Vânzători" generează un Excel cu totalul contractelor pe lună
pentru fiecare utilizator.

## Optimizare prin cache

La pornire aplicația încarcă toate locațiile în memorie pentru a naviga mai
rapid prin listă. După fiecare operație care modifică baza de date, cache-ul se
reînnoiește automat astfel încât informațiile afișate să fie actualizate.
Toate instanțele aplicației verifică periodic o valoare "version" din baza de
date, iar atunci când aceasta se modifică cache-ul se reîncarcă automat astfel
încât modificările realizate pe alt calculator devin vizibile aproape în timp
real.


