# FoldNote

Kör lokalt:

```bash
pip install -r requirements.txt
python app.py
```

Öppna på datorn: `http://localhost:5000`

För telefon på samma Wi‑Fi: `http://DIN-DATORS-IP:5000`

Servern körs på `0.0.0.0`. Data sparas automatiskt i `data/foldnote.xlsx`.

## Excel-översikt
Excel-filen innehåller bladen `Notes`, `Segments`, `Versions` och `Dashboard`. Dashboard byggs automatiskt om när data sparas/exporteras och visar anteckningar, originalstycken och alternativa versioner i läsbar struktur.

## Markera text direkt
I varje originalstycke kan du markera valfri del av texten med mus eller finger. En åtgärdsrad visas då med `Skapa alternativ version`. Den markerade texten kopieras automatiskt till en ny alternativ version, där du kan skriva om just den valda delen.


## FoldNote v3 – dokumentbaserad editor
- Skriv hela anteckningen som ett sammanhängande dokument.
- Markera valfri text direkt i dokumentet.
- Skapa alternativa formuleringar i sidopanelen.
- Behåll originalet och jämför alternativa versioner.
- Knappen `Ersätt markerad text` ersätter källtexten i dokumentet.
- Ångra/gör om finns i editorn.
- På smal/fälld skärm staplas panelerna vertikalt; på bred/vikbar skärm visas dokument och alternativ sida vid sida.
