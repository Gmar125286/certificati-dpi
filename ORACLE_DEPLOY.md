# Deploy su Oracle Cloud Always Free

Questa guida serve a pubblicare la web app su una VM Oracle Cloud Always Free con un link pubblico reale.

## 1. Crea la VM

Nel pannello Oracle Cloud:

1. crea una Compute Instance `Always Free`
2. scegli `Ubuntu` come sistema operativo
3. assegna una `Public IP`
4. salva la chiave SSH privata sul tuo PC

Fonti ufficiali:

- Always Free resources: https://docs.oracle.com/en-us/iaas/Content/FreeTier/freetier_topic-Always_Free_Resources.htm

## 2. Apri le porte in Oracle Cloud

Nel VCN/Subnet della VM aggiungi ingress rules TCP:

- `22` per SSH
- `80` per HTTP
- `443` per HTTPS

Le guide Oracle indicano esplicitamente che per pubblicare un web server accessibile da Internet bisogna aprire le porte del listener, per esempio `80/TCP`.

Riferimenti:

- https://docs.oracle.com/en/learn/publish-webserver-using-oci/
- https://docs.oracle.com/cd/E97706_01/pdf/gsg/OCI_Getting_Started.pdf

## 3. Collegati in SSH

Da Windows PowerShell:

```powershell
ssh -i C:\percorso\chiave.pem ubuntu@IP_PUBBLICO_VM
```

## 4. Clona il progetto

Sulla VM:

```bash
git clone https://github.com/Gmar125286/certificati-dpi.git
cd certificati-dpi
chmod +x oracle_setup.sh
./oracle_setup.sh
```

## 5. Verifica che l'app sia online

Apri nel browser:

```text
http://IP_PUBBLICO_VM
```

Se tutto e' corretto, vedrai la pagina di login della web app.

## 6. Dove vengono salvati i dati

La configurazione Oracle usa:

- `docker-compose.oracle.yml`
- cartella persistente `./oracle-data`

Dentro `oracle-data` resteranno:

- `storico_revisioni.db`
- `users.json`
- `settings.json`
- file DOCX/PDF generati

## 7. Comandi utili

Stato container:

```bash
cd ~/certificati-dpi
sudo docker compose -f docker-compose.oracle.yml ps
```

Log applicazione:

```bash
cd ~/certificati-dpi
sudo docker compose -f docker-compose.oracle.yml logs -f
```

Riavvio:

```bash
cd ~/certificati-dpi
sudo docker compose -f docker-compose.oracle.yml up -d --build
```

Stop:

```bash
cd ~/certificati-dpi
sudo docker compose -f docker-compose.oracle.yml down
```

## 8. Dominio personalizzato

Se hai un dominio:

1. crea un record `A`
2. punta il record all'IP pubblico della VM Oracle
3. quando vuoi, possiamo aggiungere anche HTTPS con Nginx o Caddy

## 9. Limite importante

Le istanze Always Free possono essere reclamate se restano considerate inattive da Oracle.
La documentazione Oracle segnala inoltre che a volte puo' esserci `out of host capacity` nella creazione delle risorse gratuite.
