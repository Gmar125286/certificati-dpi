# Pubblicazione su GitHub

Account GitHub rilevato in questo ambiente:

- `Gmar125286`

## Stato attuale

La cartella e pronta per essere pubblicata, ma da questo ambiente non posso creare direttamente il repository remoto perché:

- `git` non e installato sul PC
- `gh` non e installato sul PC
- gli strumenti GitHub disponibili qui non espongono la creazione di un nuovo repository remoto

## Nome repository consigliato

- `gestione-certificati-dpi`

## Cosa pubblicare

Pubblica questi file:

- `app.py`
- `avvia_programma.bat`
- `avvia_programma.vbs`
- `requirements.txt`
- `README.md`
- `irudek_norme.json`
- i modelli `CERTIFICATO DI REVISIONE - ... .docx`
- `gestione_certificati_dpi.ico`

Non pubblicare:

- `output/`
- `storico_revisioni.db`
- `users.json`
- `settings.json`

## Procedura rapida dal browser

1. Vai su GitHub con l'account `Gmar125286`.
2. Crea un nuovo repository chiamato `gestione-certificati-dpi`.
3. Non aggiungere file automatici se vuoi caricare questa cartella cosi com'e.
4. Carica i file del progetto, rispettando il `.gitignore`.

## Procedura con Git, quando disponibile

Se in futuro installi Git sul PC, dalla cartella del progetto esegui:

```powershell
cd "C:\Users\banco\Desktop\CERTIFICATI"
git init
git add .
git commit -m "Prima pubblicazione applicativo"
git branch -M main
git remote add origin https://github.com/Gmar125286/gestione-certificati-dpi.git
git push -u origin main
```

## Pubblicazione online dell'app

Questo progetto e un'app desktop Windows/Tkinter, quindi GitHub serve per:

- versionare il codice
- condividerlo
- scaricare aggiornamenti

Non rende automaticamente l'app "web". Per averla davvero online su internet servirebbe un secondo progetto:

- versione web con backend/database
- oppure packaging desktop con installer da distribuire tramite Releases GitHub
