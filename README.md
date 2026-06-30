# CSN-konverterare

Statiskt webbläsarverktyg som konverterar INSU:s CSN Excel-mall till XML-formatet `studeranderapport`.

## Version 4

- Visar hittade rader, exporterade rader, ignorerade rader och varningar.
- Har debug-panel rad för rad.
- Maskerar personnummer i gränssnittet.
- Exporterar bara rader där obligatoriska fält finns.
- Skapar XML lokalt i webbläsaren.

## Viktigt

Lägg aldrig upp riktiga CSN-filer, personnummer eller testdata med personuppgifter i repot.


## v5

Ändrat XML-nedladdningen så filen sparas som vanlig text/UTF-8-byteformat, men med samma XML-deklaration som de befintliga CSN-exempelfilerna (`encoding="utf-16"`). Detta matchar filerna som CSN redan accepterar bättre än äkta UTF-16LE med BOM.


## Version 6

- Fix för Excel-datum som kunde exporteras som YYYY-DD-MM i stället för YYYY-MM-DD.
