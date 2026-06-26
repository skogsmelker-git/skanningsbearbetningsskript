### Guide/info för inskanningsarbetet
Detta dokument beskriver ett Python skript som tillsammans OCRmyPDF effektiviserar flera delar av inskanningsarbetet. Det innehåller både dokumentation på hur vi gjorde och tänkte för att skapa skriptet samt en guide på hur vi tänker använda det i processen. 
Vi börjar med guiden eftersom det är delen vi oftast kommer återkomma till, därefter finns dokumentation och sådant. 
Bulletpoints över vad skriptet gör:
- Delar automatiskt intygen så att varje intyg blir en egen PDF
- Lägger in filnamn, volymnamn och personnummer i ett exceldokument.
- Extraherar personnummer ur innehållet i PDF:en, detta fungerar med alla olika format vi har hittat samt T nummer
o Detta gör den genom att kolla alla personnummer liknande sifferkombinationer som förekommer i varje bevis och väljer den som förekommer flest gånger, därmed blir det rätt även om OCR-skanningen inte träffade rätt på alla
o Om den inte hittar en 10-siffrig sifferkombination letar den efter andra format vi hittat, tex yyyy-mm-dd, yyyymmdd, yyyy-mmdd med mera.
- Rapporterar fel till en log fil
- Skippar hela volymen om någonting är fel med någon PDF däri, rapporterar sedan detta till en separat log fil

### Användning
För att göra OCR mha OCRmyPDF skriver man:
exec ocrmypdf -l swe --deskew --output-type pdf <pdf>
för att köra igenom hela mappar gör man:

```find . -name '*.pdf' -printf '%p\n' -exec ocrmypdf -l swe --output-type pdf '{}' '{}' \;```

För att använda bearbetningsskriptet gör man i nuläget följande i WSL:

```Python3 bearbetningsskript.py input/ output/```

men vi har tänkt att skapa executable filer så att man inte behöver skriva skriptet varje gång. Detta minskar också risken att något går snett

### Workflow
1. Ta ut alla bevisen ur fasciklarna, låt bilagor ligga kvar.
2. Skanna filerna till USB så att de ligger i mappar med samma namn som volymen.
3. För över mapparna från USB till "Obearbetat".
4. Kör OCRmyPDF kommandot.
5. Kör python skriptet.
6. Spara till "bearbetat"
8. Kolla så att allt ser bra ut och lägg sedan in i "Färdigt"
Klart! 

### Dokumentation
Här beskriver vi hur vi har tänkt
Installation
Det krävs några olika verktyg för att få detta att fungera, dessa är:
- WSL (Windows Subsystem for Linux)
- OCRmyPDF
- Python3

#### Installera WSL
För att installera WSL gör vi följande. Vi har valt Debian för att göra det så enkelt som möjligt. 

```Wsl –install```

sen startar man om datorn.
Därefter kör man:

```wsl.exe -–install Debian```

sedan får man göra en användare på Linux, tips är att använda samma uppgifter som användaren på datorn för att göra det enkelt. Klart!
Förberedelser i WSL
Efter att ha installerat och öppnat WSL behöver vi installera OCRmyPDF samt tillägget för svenska. Först gör vi en systemuppgradering med:

```sudo apt update && apt upgrade```

Sedan installerar vi OCRmyPDF och tillägget för svenska:
Nu ska vi ladda ner Python och lite annat som krävs för skriptet

```sudo apt install python3 pip python3-pypdf python3-tqdm python3-openpyxl```	

nu är de programmen vi behöver färdiga för användning. Vi behöver dock göra en liten ändring för att inte hela WSL lagringsutrymmet ska fyllas av temporära filer. För att fixa det gör vi ett bind mount från Windows filsystem över WSL mappen där OCRmyPDF lägger temporära filer. Så vi går in i WSL och redigerar följande fil:

```sudo nano /etc/fstab```

Där skriver vi 

```/mnt/c/Users/<windows-användare>/SKANNINGSMAPPEN/OCRmyPDF-tmp /tmp none bind 0 0```

Där <windows-användare> är användarnamnet på datorn, tex melste. Därefter måste man tillämpa de nya inställningarna genom att:

```sudo systemctl daemon-reload```

```sudo mount -a```

klart!
#### Genomgång av skriptet
Skriptet går att finna här.
#### Problem och förbättringar
- Vet inte om man kan komma åt filareorna genom WSL eftersom det är Microsofts egna nätverksdiskar…
- Det kan finnas fler variationer på personnummer eller bevisformat vi inte tänkt på, men dessa borde vi hitta i log-filerna
- Vi måste komma ihåg att konvertera allt till pdf/a2u
- Hittar den personnumer i formatet xxxxxx-Rxxx? Det finns ett exempel på ett sådant i volym 2016 860301-860829.
- Det skulle vara bra att införa någon ”alert” så att den varnar om ett bevis består av många sidor (kke 10+) så att man enkelt kan kolla om det blev något fel
- Vi kan göra ett valideringsskript som bara kollar om det förekommer flera personnummer i samma dokument, kan också rapportera om filen är över 10 sidor eller något sånt.


#### Referenser
- How to install Linux on Windows with WSL, Microsoft. (Länk)
- Cookbook – Basic examples, OCRmyPDF. (Länk)
- Installing additional language packs, OCRmyPDF. (Länk)
