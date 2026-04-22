# Gestione Certificati DPI

Applicazione desktop in Python per gestire revisioni DPI, generare certificati Word/PDF, stampare in duplice copia e mantenere uno storico locale delle revisioni.

## Funzioni principali

- archivio locale delle revisioni con ricerca per cliente, lotto e seriale
- modifica di una revisione gia salvata
- anteprima certificato dallo storico
- eliminazione revisione e file collegati
- login iniziale con utenti locali
- gestione utenti con password non salvate in chiaro
- compilazione automatica della norma da codice SAP Irudek
- ricerca prodotto Irudek per nome
- aggiornamento manuale del catalogo Irudek
- esportazione Word e PDF
- stampa 2 copie

## Struttura progetto

- `app.py`: applicazione principale
- `avvia_programma.bat`: avvio rapido
- `avvia_programma.vbs`: avvio senza finestra CMD
- `irudek_norme.json`: catalogo locale SAP -> norma
- `CERTIFICATO DI REVISIONE - ... .docx`: modelli certificato

## Requisiti

- Windows
- Python 3.11+ consigliato
- Microsoft Word installato per esportazione PDF

## Dipendenze Python

```powershell
pip install -r requirements.txt
```

## Avvio locale

Da Esplora File:

- doppio clic su `avvia_programma.bat`

Oppure da terminale:

```powershell
cd "C:\Users\banco\Desktop\CERTIFICATI"
python app.py
```

## Login iniziale

Utente amministratore predefinito:

- utente: `admin`
- password: `admin`

Dopo il primo accesso puoi creare altri utenti dall'applicazione.

## Note importanti

- i dati di lavoro locali come `storico_revisioni.db`, `users.json`, `settings.json` e la cartella `output` non vanno pubblicati su GitHub
- l'export PDF usa Microsoft Word tramite automazione locale di Windows
- l'aggiornamento catalogo Irudek e manuale

## Pubblicazione su GitHub

I passaggi pronti sono descritti in `PUBLISH_ON_GITHUB.md`.
