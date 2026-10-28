# Statistik Brug Aarhus – Book Naturen

Robot, der henter bookinger og faciliteter (shelterpladser, bålhytter, bålpladser m.m.) fra Book Naturen-API'et og lægger dem som Excel-filer i SharePoint, hvor de bruges som datakilde i Power BI.

## Hvad robotten gør

Robotten kører dagligt via OpenOrchestrator og laver to ting:

1. **Faciliteter** – henter den komplette liste over steder (inkl. slettede) og overskriver `faciliteter.xlsx`.
2. **Bookinger** – downloader `shelterbookinger.xlsx` fra SharePoint, henter bookinger fra et halvt år tilbage til to år frem, opdaterer eksisterende bookinger (fx nye aflysninger), tilføjer nye og uploader filen igen.

Bookinger ældre end et halvt år hentes ikke igen, så API'et ikke belastes unødigt.

## Første kørsel

Den daglige kørsel forudsætter, at `shelterbookinger.xlsx` allerede findes i SharePoint. Første gang (eller hvis kolonnerne ændres) skal første kørsel-scriptet bruges, som henter alle bookinger fra 1. januar 2025 og frem.

## Output

Filerne lægges i `Shared Documents/Data/BookNaturen` på det SharePoint-site, der er angivet i konstanten `PowerBiSharepoint`.

| Fil | Ark | Tabel | Indhold |
|---|---|---|---|
| `shelterbookinger.xlsx` | Bookinger | Bookinger | Én række pr. booking, inkl. pladsoplysninger |
| `faciliteter.xlsx` | Sheltere | tblSheltere | Én række pr. facilitet |

I Power BI kobles de to tabeller på `Plads-ID` (én-til-mange fra faciliteter til bookinger).

## Opsætning i OpenOrchestrator

**Konstanter**

| Navn | Indhold |
|---|---|
| `brugaarhusapiurl` | Basis-URL til Book Naturen-API'et |
| `PowerBiSharepoint` | URL til SharePoint-sitet |

**Credentials**

| Navn | Brugernavn | Password |
|---|---|---|
| `brugaarhusapikey` | API-brugernavn | API-nøgle |
| `SharePointAPI` | Tenant | Client ID |
| `SharePointCert` | Thumbprint | Sti til certifikat |
| `brugaarhushashnoegle` | – | Hemmelig nøgle til pseudonymisering |

Hash-nøglen kan alternativt sættes som miljøvariablen `BrugAarhusHashNoegle`, som så bruges i stedet.

## Persondata

Navne og e-mails gemmes ikke i Excel-filerne (styres af `INKLUDER_PERSONDATA`). I stedet beregnes:

- **Person-ID** – et fast, anonymt tal ud fra e-mailen (HMAC-SHA256 med den hemmelige nøgle), så man kan se, om samme person booker flere gange, uden at kunne se hvem.
- **Aarhus-mail** – Ja/Nej for, om der er booket med en `@aarhus.dk`-adresse.

Person-ID er pseudonymiserede personoplysninger og skal behandles derefter.

> **Vigtigt:** Hash-nøglen må aldrig ændres, da alle Person-ID'er så skifter. Den må heller ikke deles eller committes, da Person-ID'er ellers kan kobles tilbage til e-mailadresser.

## Værd at vide om data

- API'et bruger 1. januar 1900 som "ingen dato". Disse datoer bliver til tomme celler.
- En booking er aflyst, hvis den har en aflysningsdato. En booking uden bekræftelsesdato er ikke bekræftet.
- Feltet `Nights` i API'et indeholder reelt antal **timer**. Sluttidspunktet beregnes som starttidspunkt + antal timer.
- Pladsbeskrivelser renses for HTML, og usynlige kontroltegn fjernes, da Excel ikke tillader dem.
- Filen uploades i bidder via en midlertidig fil for at undgå en låsefejl i Office365-pakken på Windows.

## Afhængigheder

- OpenOrchestrator 3.0.0
- Pillow 12.3.0
- requests 2.34.2
- pandas 3.0.6
- openpyxl 3.1.5
- Office365-REST-Python-Client 3.2.0

Kræver Python 3.11 eller nyere.
