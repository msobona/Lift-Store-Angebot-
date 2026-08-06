# Repository auf GitHub anlegen (einmalig)

Der Cloud-Agent darf unter deinem Account **kein neues Repo** erstellen  
(nur Zugriff auf `WAMAS_WebApp_4.0`). Deshalb einmalig bei dir lokal:

## Variante A – Script (empfohlen)

1. Diesen Ordner lokal speichern (oder ZIP entpacken).
2. Terminal im Projektordner öffnen.
3. Ausführen:

```bash
gh auth login
chmod +x scripts/publish_to_github.sh
# optional: OWNER=msobona VISIBILITY=private ./scripts/publish_to_github.sh
./scripts/publish_to_github.sh
```

Es entsteht das Repo **`WAMAS-Lift-Store-Angebot`**  
(Anzeigename auf GitHub kannst du danach auf **„WAMAS Lift & Store Angebot“** setzen).

## Variante B – GitHub Website + Bundle

1. Auf GitHub: **New repository** → Name `WAMAS-Lift-Store-Angebot` → **Private** → **ohne** README anlegen.
2. Lokal:

```bash
git clone WAMAS-Lift-Store-Angebot.bundle WAMAS-Lift-Store-Angebot
cd WAMAS-Lift-Store-Angebot
git remote add origin https://github.com/DEIN_USER/WAMAS-Lift-Store-Angebot.git
git push -u origin master
```

## Anzeigename

GitHub → Settings → General → Repository name bleibt `WAMAS-Lift-Store-Angebot`,  
unter Description z. B. **WAMAS Lift & Store Angebot**.
