# CSN-konverterare

Statisk GitHub Pages-sida som läser Excel-mallen lokalt i webbläsaren och skapar XML enligt strukturen i CSN-exempelfilerna.

## Viktigt
- Lägg inte upp riktiga rapportfiler eller personnummer i GitHub-repot.
- Sidan kör konverteringen i webbläsaren. Filen skickas inte till någon server.
- Nuvarande version använder SheetJS via CDN för att läsa `.xlsx`.

## Publicera på GitHub Pages
1. Skapa ett repo, till exempel `csn-konverterare`.
2. Lägg in `index.html`, `style.css`, `script.js` och `README.md`.
3. Gå till Settings → Pages.
4. Välj branch `main` och root folder.
5. Öppna länken GitHub Pages ger dig.
