# Frysker Vertaler — Installatie

Voegt een knop **"Oersette nei Frysk"** toe aan Outlook waarmee je e-mailtekst vertaalt naar it Frysk via [frisian.eu](https://frisian.eu).

Werkt op Outlook Mac, Windows en het web. Geen server nodig.

---

## Stap 1 — Bestanden op GitHub Pages zetten

1. Maak een gratis account op [github.com](https://github.com) (als je die nog niet hebt)
2. Maak een nieuwe repository aan, bijv. `frysker`
3. Upload de twee bestanden uit de map `src/`:
   - `commands.html`
   - `commands.js`
4. Ga naar **Settings → Pages** → stel in: **Branch: main**, map: `/ (root)`
5. GitHub geeft je een URL, bijv. `https://caspar.github.io/frysker`

---

## Stap 2 — manifest.xml aanpassen

Open `manifest.xml` en vervang **alle** voorkomens van `JOUW_GITHUB_URL` door je GitHub Pages URL (zonder `https://`):

```
JOUW_GITHUB_URL  →  caspar.github.io/frysker
```

Er zijn 6 plekken om aan te passen. Sla op.

> Je hoeft de icoontjes niet te hebben — Outlook toont een standaard icoontje als de afbeeldingen niet bestaan.

---

## Stap 3 — Add-in laden in Outlook

### Voor jezelf (sideloaden)

**Mac:**
1. Open Outlook → nieuw bericht
2. Klik op de drie puntjes (`...`) in de werkbalk van het bericht
3. Kies **Invoegtoepassingen ophalen**
4. Klik op **Mijn invoegtoepassingen** → **Aangepaste invoegtoepassing toevoegen** → **Toevoegen vanuit bestand**
5. Selecteer `manifest.xml`

**Web (outlook.com / OWA):**
1. Open een nieuw bericht
2. Klik op de drie puntjes (`...`) → **Invoegtoepassingen ophalen**
3. **Mijn invoegtoepassingen** → **Aangepaste invoegtoepassing toevoegen** → **Toevoegen vanuit bestand**
4. Selecteer `manifest.xml`

### Voor meerdere gebruikers (via M365 Admin Center)

1. Ga naar [admin.microsoft.com](https://admin.microsoft.com)
2. **Instellingen → Geïntegreerde apps → App toevoegen**
3. Kies **Office-invoegtoepassing uploaden** en upload `manifest.xml`
4. Wijs toe aan gebruikers of de hele organisatie
5. Gebruikers zien de knop automatisch verschijnen — niets te doen

---

## Gebruik

1. Open een nieuw e-mailbericht en schrijf je tekst in het **Nederlands**
2. Klik op **"Oersette nei Frysk"** in het lint
3. De tekst wordt vervangen door de Fryske vertaling

---

## Bestandsstructuur

```
outlook_frysk/
  manifest.xml        ← Pas JOUW_GITHUB_URL aan (stap 2)
  src/
    commands.html     ← Upload naar GitHub Pages
    commands.js       ← Upload naar GitHub Pages
```
