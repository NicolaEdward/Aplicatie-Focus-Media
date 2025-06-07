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

Poti crea un fisier `.env` in radacina proiectului cu aceste variabile. Un
exemplu este disponibil in fisierul `.env.example`:

```
MYSQL_HOST=127.0.0.1
MYSQL_PORT=3306
MYSQL_USER=root
MYSQL_PASSWORD=Root2001
MYSQL_DATABASE=aplicatie_vanzari
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


