# Gestione Certificati DPI

Applicazione desktop in Python per gestire revisioni DPI, generare certificati Word/PDF, stampare in duplice copia e mantenere uno storico locale delle revisioni.

Ora il progetto include anche una versione web locale, da aprire nel browser sul PC Windows.

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
- versione web locale con login, storico, modifica ed export file

## Struttura progetto

- `app.py`: applicazione principale
- `webapp.py`: applicazione web locale Flask
- `avvia_programma.bat`: avvio rapido
- `avvia_programma.vbs`: avvio senza finestra CMD
- `avvia_webapp.bat`: avvio rapido versione web
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

## Avvio web locale

Da Esplora File:

- doppio clic su `avvia_webapp.bat`

Poi apri il browser su:

```text
http://127.0.0.1:5000
```

Oppure da terminale:

```powershell
cd "C:\Users\banco\Desktop\CERTIFICATI"
python app.py
```

Per la versione web:

```powershell
cd "C:\Users\banco\Desktop\CERTIFICATI"
python webapp.py
```

## Pubblicazione online

Il progetto ora e' preparato anche per essere pubblicato come web app pubblica.

File pronti per il deploy:

- `webapp.py`: backend Flask
- `.replit`: configurazione run/deploy per Replit
- `railway.json`: healthcheck e restart policy per Railway
- `Dockerfile`: ambiente server con LibreOffice per export PDF
- `render.yaml`: configurazione Render con web service e disco persistente
- `.dockerignore`: esclusione dei file locali non da pubblicare

### Railway

Per usare Railway puoi partire direttamente dalla repository GitHub pubblica.

Documentazione ufficiale Railway usata per questa configurazione:

- Deploy Flask: https://docs.railway.com/guides/flask
- Config as code: https://docs.railway.com/config-as-code/reference
- Start command: https://docs.railway.com/deployments/start-command
- Volumi persistenti: https://docs.railway.com/guides/volumes
- Domini pubblici e custom: https://docs.railway.com/networking/domains/working-with-domains

Flusso consigliato:

1. crea un nuovo progetto su Railway
2. scegli `Deploy from GitHub repo`
3. collega `Gmar125286/certificati-dpi`
4. lascia che Railway usi il `Dockerfile` presente
5. dopo il primo deploy aggiungi un `Volume` al servizio
6. monta il volume, per esempio su `/data`
7. imposta la variabile `DATA_DIR=/data` se Railway non la compila da solo
8. nella sezione networking genera il dominio pubblico `*.up.railway.app`

Alla fine ottieni un link pubblico tipo:

```text
https://gestione-certificati-dpi.up.railway.app
```

Poi puoi aggiungere anche un dominio personalizzato dal pannello Domains.

### Replit

Per usare Replit puoi importare direttamente la repository pubblica:

```text
https://replit.com/github.com/Gmar125286/certificati-dpi
```

Documentazione ufficiale Replit usata per questa configurazione:

- Import da GitHub: https://docs.replit.com/getting-started/quickstarts/import-from-github
- Configurazione `.replit`: https://docs.replit.com/replit-app/configuration
- Deployments: https://docs.replit.com/cloud-services/deployments/about-deployments
- Custom domains: https://docs.replit.com/cloud-services/deployments/custom-domains

Flusso consigliato:

1. importa la repo su Replit
2. verifica che il progetto parta con `Run`
3. apri `Deployments`
4. crea un deployment `Autoscale` oppure `Reserved VM`
5. usa il comando di deploy gia pronto nel file `.replit`
6. al termine ottieni un link pubblico tipo `https://nome-app.replit.app`

Poi, se vuoi, puoi collegare anche un dominio personalizzato dal pannello Replit.

### Render

La strada piu semplice per ottenere un link pubblico e':

1. aprire la repo GitHub pubblica
2. creare un nuovo servizio su Render collegando la repo
3. usare il `render.yaml` gia presente
4. lasciare il mount del disco su `/opt/app/data`
5. attendere il deploy

Al termine Render assegna un URL pubblico tipo:

```text
https://gestione-certificati-dpi.onrender.com
```

Poi puoi associare anche un dominio personalizzato dal pannello Render.

[![Deploy to Render](https://render.com/images/deploy-to-render-button.svg)](https://render.com/deploy?repo=https://github.com/Gmar125286/certificati-dpi)

### Note per la versione online

- i dati applicativi (`storico_revisioni.db`, `users.json`, `settings.json`, output generati) vengono salvati nella cartella dati persistente del server
- la stampa diretta di 2 copie resta pensata per ambiente Windows locale
- l'export PDF lato server usa LibreOffice invece di Microsoft Word

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
