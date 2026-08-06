# Windows: Repo anlegen (ohne Linux-Befehle)

Du bist in **CMD** unter Windows. Dort funktionieren `gh`, `chmod` und `./script.sh` nicht von alleine.

## Variante 1 – Doppelklick / CMD (einfachste)

1. ZIP entpacken, z. B. nach  
   `C:\Users\msobona\WAMAS-Lift-Store-Angebot`
2. **CMD öffnen** und in den Ordner wechseln:

```bat
cd /d C:\Users\msobona\WAMAS-Lift-Store-Angebot
```

3. Im Browser leeres Repo anlegen: https://github.com/new  
   - Name: `WAMAS-Lift-Store-Angebot`  
   - **Private**  
   - **kein** README, keine License, kein .gitignore  
4. Script starten:

```bat
scripts\publish_to_github.bat msobona
```

Wenn nach dem Push nach Benutzer/Passwort gefragt wird:  
GitHub-Passwort = **Personal Access Token** (nicht das Login-Passwort).  
Token: GitHub → Settings → Developer settings → Personal access tokens.

## Variante 2 – Nur Git (ohne Script)

Nach dem Anlegen des leeren Repos im Browser:

```bat
cd /d C:\Users\msobona\WAMAS-Lift-Store-Angebot
git remote remove origin
git remote add origin https://github.com/msobona/WAMAS-Lift-Store-Angebot.git
git branch -M master
git push -u origin master
```

## Optional: GitHub CLI unter Windows

Nur wenn du `gh` nutzen willst:

1. https://cli.github.com/ installieren  
2. Terminal **neu** öffnen  
3. `gh auth login`  
4. dann wieder `scripts\publish_to_github.bat msobona`

## Nicht in CMD verwenden

| Linux | Windows CMD |
|-------|-------------|
| `chmod +x …` | nicht nötig |
| `./scripts/….sh` | `scripts\publish_to_github.bat` |
| `OWNER=msobona ./…` | `scripts\publish_to_github.bat msobona` |
