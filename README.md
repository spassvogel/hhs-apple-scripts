# hhs-apple-scripts

Verzameling scripts voor HHS docenten
## mail-students

Het verzenden van individuele mailtjes naar studenten klinkt als nogal veel werk dus ik heb een scriptje gemaakt wat het een en ander zou moeten vergemakkelijken. 

### Allereerst een paar pre-requisites

- het werkt alleen op Macsje moet Outlook hebben geïnstalleerd
- je moet al je assessments/beoordelingen die je wilt verzenden in een map op je computer hebben staan in pdf formaat. 
- de naam van het bestand moet het studentnummer zijn. niet meer, niet minder. Dus  20014505.pdf  bijvoorbeeld. De reden dat dit zo precies is is omdat het script dat gebruikt voor de addressering.
 
Het script vind je hier: [mail-students.scp](https://github.com/spassvogel/hhs-apple-scripts/blob/main/mail-students.scpt?raw=true)

ik zal eventuele fixes of verbeteringen hier ook doorvoeren

### Hoe werkt het?

- Je download het script, als je hem dan opent dan opent in  Script Editor. 
- Dan klik je op de Play knop.Als eerste geef je de mail subject op. Dan de body text (gewoon de broodtext van de mail dus). Je moet dat wel algemeen houden, is  niet uniek per student oid.
- Dan selecteer je een map waar de .pdf bestanden staan. Daarna zal ie Outlook openen, en een X aantal mailtjes klaarzetten naar de studenten (gebruikmakend van studentnr@student.hhs.nl) met de pdf als bijlage
- Je moet dan nog zelf alle mails verzenden. Eventueel kun je de mailtjes aanpassen met een gepersonaliseerde boodschap ofzo.
