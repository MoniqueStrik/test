# Firap macro's Gemeente Best: projectcontext

Dit bestand is bedoeld om in Visual Studio Code verder te werken (ook goed te lezen door Claude Code). Het beschrijft wat er is gebouwd, hoe het werkt, welke afspraken gelden en wat nog open staat.

Gebruiker: financieel adviseur bij de Gemeente Best. Antwoorden in het Nederlands. Gebruik in antwoorden geen streepjes als leesteken (geen " - " of lange streepjes), maar punt of komma. Wees kritisch en praat niet standaard mee.

## 1. Doel

Een Excel werkmap (`.xlsm`) voor de financiële rapportages (firap) met een dashboard en drie knoppen:

1. Inleesbestand voor het administratiepakket (CODA) maken of bijwerken.
2. Verschillenanalyses en de ontwikkeling van het begrotingssaldo maken of bijwerken.
3. Overgang naar een nieuwe firap (1e, 2e of 3e firap plus jaar).

## 2. Bestanden in de repo

| Bestand | Rol |
|---|---|
| `Format_nieuwe_financiele_rapportage.xlsm` | **Laatste en leidende werkmap** (door de gebruiker in Excel opgeslagen en daarna door mij bijgewerkt). Hierin verder werken. |
| `modInleesFirap.bas` | VBA module: inleesbestand (tabblad tbv adm pakket) en controle (tabblad controle inlees). |
| `modVerschillen.bas` | VBA module: verschillenanalyses, hulpblad en sectie D van Ontwikkeling begrotingssaldo. |
| `modDashboard.bas` | VBA module: overgang naar nieuwe firap, mutaties wissen. |
| `W16_2026_3e_firap_2026.xlsm` | Verouderde tussenversie. Niet meer gebruiken. |
| `Format_nieuw_begrotingswijzigingen_*.xlsm` | Andere, niet gerelateerde bestanden (begrotingswijzigingen). Niet aanraken. |

De `.bas` bestanden zijn de bron van waarheid voor de code. De VBA in het `.xlsm` is een kopie daarvan.

Branch: `claude/relaxed-wozniak-7t2am2`. Geen pull request gemaakt.

## 3. Tabbladen van de werkmap

| Tabblad | Zichtbaar | Inhoud |
|---|---|---|
| `Dashboard` | ja, vooraan, lichtgroen | Invoer, drie knoppen, status, meldingen. |
| `1e firap 2027` (naam wisselt) | ja | Bronblad. Kolom A grootboek of categorie, C omschrijving, D I/S, daarna per jaar een paar N/B plus bedrag (E/F, G/H, I/J, K/L, M/N). P rubriek investeringen, Q raadsproof omschrijving, R doorgevulde R-/O-code. Rij 10 bevat de jaartallen. |
| `Ontwikkeling begrotingssaldo` | ja | Secties A, B en D. D6 en D20 zijn invoercellen (via dashboard bij een 1e firap). |
| `verschillenanalyses firap` | ja | Presentatieblad met vier secties en het controleblok. Alleen tekst en bedragen als formules naar het bronblad. |
| `tbv adm pakket` | ja | Het CODA inleesbestand (was `tbv Coda`). Kolommen A t/m E plus een bedragkolom per jaar. |
| `controle inlees` | ja (wordt door de macro gemaakt) | Controle met formules naar bron en inleesbestand. B3 is de statusregel. Hulpkolommen M t/m O. |
| `taakvelden grootbnr en cat` | ja | Niet door de macro's gebruikt. |
| `hulp verschillenanalyse` | verborgen | Ankers (rijnummers per sectie) en controlesommen. |
| `omschrijvingen` | verborgen (wordt zichtbaar bij concepten) | Tabel sleutel naar omschrijving (max 27 tekens) voor kolom E van het inleesbestand. |

## 4. Dashboard: celkaart

| Cel | Betekenis |
|---|---|
| C4 | Rapportage: `1e firap`, `2e firap` of `3e firap` (keuzelijst). |
| C5 | Begrotingsjaar (vier cijfers). |
| C6 | Begrotingswijzigingsnummer, bijvoorbeeld `W03/27`. |
| C7 | Formule: bronbladnaam, bijvoorbeeld `1e firap 2027`. |
| C8 | Begrotingssaldo. Verplicht bij een echte overgang naar een 1e firap, komt in Ontwikkeling D6. |
| C9 | Saldo onvoorzien. Verplicht bij een 1e firap, komt in Ontwikkeling D20 en in het bronblad bij "Budget voor onvoorziene uitgaven na vaststelling". |
| C10 | Datum vergadering raad. Komt onderaan het bronblad in de regel `Best, <datum>`. |
| C15, C16, C17 | Status inleesbestand, verschillenanalyse en overgang. D15 t/m D17 het tijdstip van de laatste actie. |
| B29 en lager | Meldingen (bronfouten, concepten). |

Alle macro's lezen de bronbladnaam uit C4 en C5. Staat die niet in de werkmap, dan stoppen ze met een melding.

## 5. Macro's

### modInleesFirap

* `MaakOfWerkInleesbestandBij`: knop. Bouwt `tbv adm pakket` en `controle inlees`.
* `InleesbestandStil`: zonder meldingsvensters (voor tests, resultaat in `gResultaat`).
* `MaakKnopTabblad`: maakt in een werkmap zonder dashboard het tabblad `Dashboard` met invoercellen en drie knoppen.

Regels (uit de skill "inlees financiële rapportages", versie 3):

* Eén inleesregel per bronregel met een bedrag. Teken uitsluitend uit de N/B kolom links van het bedrag: B wordt negatief, N blijft positief. Afgerond op hele euro's, half omhoog.
* Delen en de volgorde: deel 1 (exploitatie), saldoregel 611004 met categorie 43812, dan 3a investeringen, 3b-1 reserves, 3b-2 voorzieningen. Deel 2 (onvoorzien) levert geen inleesregels.
* Kolom A: grootboek (deel 1), `+` voor investeringen, reserves en voorzieningen, balansnummers (7 cijfers die met 0 of 1 beginnen) als tekst. Kolom B: 5 cijferige categorie uit de bedragregel zelf (begint met 4 of 8). Kolom C: verbijzondering (investeringsnummer, R- of O-nummer, of deel na het streepje bij `47500 - G0901`).
* Saldoregel 611004 per jaar is het negatief van de som van deel 1. Sluitrekening 2999999 per jaar is het negatief van alle regels erboven.
* Omschrijving: `W03-27 ` (wijzigingsnummer met streepje) plus een samenvatting. Totaal maximaal 35 tekens. Bij OHV altijd het O-nummer, bij reserves het R-nummer.
* Bronfouten stoppen de macro en leveren een lijst op het dashboard: categorie niet 5 cijfers met 4 of 8 vooraan, bedrag zonder N/B, negatief bedrag in de bron, ontbrekend anker of saldoregel.
* `controle inlees` bevat formules naar het bronblad en naar `tbv adm pakket`. Status in B3 begint met `OK` of `LET OP`.
* Omschrijvingen komen uit het verborgen tabblad `omschrijvingen`. Onbekende regels krijgen een concept (gele regel) dat de gebruiker moet nalopen. Bij concepten maakt de macro het tabblad zichtbaar.

### modVerschillen

* `MaakOfWerkVerschillenanalyseBij`: knop. Vraagt Ja (bijwerken) of Nee (volledig opnieuw opbouwen).
* `VerschillenBijwerkenStil` en `VerschillenVolledigStil`: stille varianten.

Regels (uit de skill "verschillenanalyses 3e firap volledig v2"):

* Exploitatie in drie blokken: incidenteel, structureel, overhevelingen en aframingen naar het volgende jaar (herkend aan de verwijzing naar de lijst "Overheveling of aframen"). Mutaties onder de grens (€ 50.000, hulpblad B4) in één regel "Diverse verschillen". Aflopend gesorteerd. Tekenconventie exploitatie: nadeel positief.
* Investeringen per rubriek uit kolom P (door de gebruiker ingevuld, nooit zelf verzinnen). Is P leeg dan losse regels. Is P deels leeg dan stoppen met een melding.
* Reserves en voorzieningen per mutatie, B is toename (positief), N is afname.
* Controleblok met vijf LET OP regels, leeg als het aansluit. Totaal en subtotaal labels en kleuren volgen het teken.
* Ankers op het bronblad altijd op tekst zoeken, nooit op vast rijnummer.
* Bijwerken: verdwenen bronregels opruimen, nieuwe regels in het juiste blok invoegen, subtotalen en totalen herstellen. Onverbonden paren met gelijk bedrag en tegengesteld teken gelden als budgetneutraal en krijgen geen regel.
* Bij volledig opbouwen worden bestaande omschrijvingen hergebruikt, daarna kolom Q, daarna een concept.
* Sectie D van `Ontwikkeling begrotingssaldo` wordt altijd opnieuw gekoppeld aan het bronblad. Labels en kleuren volgen het teken.
* Alle formules worden met een hulpfunctie `SetF` geschreven die een tekstopmaak eerst op Standaard zet (anders komt de formule als tekst in de cel).

### modDashboard

* `OvergangNaarNieuweFirap`: knop. Hernoemt het bronblad naar de keuze op het dashboard, zet de jaartallen in rij 10 (eerste jaar plus vier), zet het wijzigingsnummer, het zaaknummer (BL + laatste twee cijfers van het jaar + `-xxxxx`) en de datum vergadering raad, past kopteksten aan, schrijft D6, D20 en F335 bij een 1e firap. Vraagt apart of alle mutaties gewist moeten worden.
* Is de gekozen naam gelijk aan het bestaande bronblad dan is het alleen een bijwerking van wijzigingsnummer, datum en kopteksten. Er wordt dan niets gewist en C8 en C9 zijn niet verplicht.
* `OvergangStil` (wissen) en `OvergangStilBehoud` (niet wissen): stille varianten.
* Wissen laat kopregels, saldoregels en formules staan en wist per mutatieregel D t/m N, P, Q en bij gewone mutatieregels ook A en C, plus doorlopende toelichtingen naast een mutatieregel. Daarna wordt de analyse leeg opnieuw opgebouwd en worden de inleesregels op `tbv adm pakket` gewist.

## 6. Afspraken en valkuilen

* Het bronblad is vastgesteld materiaal. De inlees en analyse macro's wijzigen het nooit. Alleen Overgang past naam, jaartallen, wijzigingsnummer, zaaknummer, datum, kopteksten en bij 1e firap F335 aan, en wist op verzoek mutaties.
* `.bas` bestanden: alleen ASCII, regeleinden CRLF, codering Windows-1252. Euro en trema met `ChrW(8364)` en `ChrW(235)`.
* Namen van variabelen mogen niet botsen met functies of sleutelwoorden: een array `rGB` botst met `RGB()`, `oN` botst met het woord `On`. Dat gaf in Excel en LibreOffice compilatiefouten.
* Geen `Scripting.Dictionary` gebruiken (werkt niet op Mac). Collections wel.
* Excel zet bij opslaan de hoofdletters van variabelen gelijk (bijvoorbeeld `Txt` wordt `txt`). Dat is onschuldig.
* Controles van tabblad `Dashboard` C15 verwijzen pas naar `controle inlees` nadat dat blad bestaat. Een directe verwijzing naar een nog niet bestaand blad gaf bij openen de melding over een gekoppelde werkmap.
* Excel meldde onleesbare inhoud over `calcChain.xml` nadat formules waren gewist. Oplossing: dat onderdeel uit het bestand halen, Excel bouwt het zelf opnieuw op.
* Een formule schrijven in een cel met opmaak Tekst levert tekst op. Daarom `SetF`.

## 7. Werken in Visual Studio Code

1. Open de map van de repo in VS Code. Bewerk de `.bas` bestanden. Handig: een extensie met VBA syntax highlighting.
2. Controleer na elke wijziging de regeleinden (CRLF) en dat er geen tekens buiten ASCII in staan.
3. Zet de code in Excel: Alt+F11, rechtsklik op de oude module, Verwijderen (zonder exporteren), daarna Bestand, Bestand importeren en het `.bas` bestand kiezen. Sla op als `.xlsm`.
4. Test in Excel met de knoppen. De code is in LibreOffice (VBA modus) getest, nooit in Excel door mij.
5. Commit alleen op de branch `claude/relaxed-wozniak-7t2am2`.

Nieuwe macro's koppelen aan een knop: tabblad Ontwikkelaars aanzetten, Invoegen, Knop (formulierbesturingselement), macro kiezen. Of `MaakKnopTabblad` draaien in een andere werkmap.

## 8. Testen

Een snelle handmatige controle in Excel na elke wijziging:

1. Knop Inleesbestand: `controle inlees` B3 begint met `OK`. Op `tbv adm pakket` staan de jaren vanaf het jaar van de rapportage.
2. Knop Verschillenanalyse, eerst Ja en dan Nee: alle vijf controleregels leeg.
3. Voeg in het bronblad een mutatie toe (grootboekkop, categorie 4 of 8, I of S, N of B, bedrag) en draai beide knoppen: de regel komt op de goede plek en alles sluit aan.
4. Overgang met en zonder wissen op een kopie van het bestand.

## 9. Open punten en ideeën

* Alles is niet in Excel getest. Eerst een rondje door de knoppen in Excel.
* Aanname: balansnummers van 7 cijfers beginnen met 0 of 1, de rest is een investeringsnummer (constante `BALANS_EERSTE_CIJFER` in `modInleesFirap`).
* Omschrijvingen van het inleesbestand zijn een samenvatting per bronregel. Een macro kan niet samenvatten: de tabel `omschrijvingen` bevat mijn voorstellen voor de eerste firap, nieuwe regels zijn concepten.
* Bij een 2e of 3e firap is het begrotingssaldo (C8) niet verplicht. Nog te bepalen hoe D6 en D20 daar gevuld moeten worden.
* Overgang laat de vaste regels van de vorige rapportage in `Ontwikkeling begrotingssaldo` (A10:D14) staan tenzij mutaties worden gewist. Het zaaknummer (BL..) en de datum "Best, ..." worden bij overgang vervangen, het peildatumdeel in de kopregel van Ontwikkeling is bewust weggehaald.
* Mogelijke uitbreiding: wisfunctie ook voor de 2e en 3e firap met eigen regels, en een export van de meldingen naar een tekstbestand.
