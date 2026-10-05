# Arter per naturvärdesbiotop

ArcGIS Pro-verktyg som sammanställer vilka arter som förekommer i varje naturvärdesbiotop
och hur många av dem som är invasiva, skyddade, rödlistade eller övriga värdearter. Resultatet
är en tabell med en rad per biotop, och valfritt en Excel-fil:

| Biotop (Kart-ID) | Artlista | Varav invasiva (lista) | Varav invasiva (antal) | Varav skyddade (lista) | Varav skyddade (antal) | Varav rödlistade (lista) | Varav rödlistade (antal) | Varav övriga värdearter (lista) | Varav övriga värdearter (antal) |
|---|---|---|---|---|---|---|---|---|---|

Övriga värdearter är arter som varken är invasiva, skyddade eller rödlistade.

## Krav

ArcGIS Pro 3.x, Basic-licens räcker. Inga paket utöver det som följer med Pro.

## Lägg till i ArcGIS Pro

Catalog-fönstret, högerklicka Toolboxes, Add Toolbox, välj `ArterPerNaturvardesbiotop.pyt`.

## Indata

- Ett punktlager med artförekomster med ett artnamnsfält och heltalsfält (1/0) för rödlistad,
  skyddad och invasiv. Standardnamnen följer SIS-mallen för NVI: `taxon_svensktNamn`,
  `A_rodlist`, `A_skyddad`, `A_IV`, samt källfältet `kalla`.
- Ett polygonlager med biotoper och ett ID-fält, normalt `A_kartID`. Polygoner med samma ID
  räknas som samma biotop.

Urval och definitionsfrågor på båda lagren respekteras.

## Parametrar

| Parameter | Förklaring |
|---|---|
| Artförekomster | Punktlager med arter. |
| Naturvärdesbiotoper | Polygonlager med biotoper. |
| Biotopens ID-fält | Fält som identifierar biotopen, standard `A_kartID`, annars `EkoID`. |
| Artnamnsfält | Standard `taxon_svensktNamn`. |
| Utdatatabell | Tabell i en fil-geodatabas. |
| Excel-fil | Valfri .xlsx med rubrikerna som kolumnnamn. |
| Rödlistad, Skyddad, Invasiv | Flaggfält (1/0). Lämnas ett tomt blir motsvarande kolumner tomma och arterna räknas som övriga värdearter. |
| Källfält, Värde för egna fynd | Arter som i en biotop bara är kända från andra källor än egen inventering får "(ADB)" efter namnet. Standard `kalla` = `Fältinventering`. |
| Kontrolltabell | Valfri tabell med en rad per biotop och art, med antal fynd och källa. |
| Sökavstånd | Räkna även punkter inom ett avstånd från biotopen, till exempel 5 m för GPS-fel. |
| Ta med biotoper utan arter | Standard ja. |

## Bra att veta

- Artnamnen används som de står, bortsett från mellanslag i början och slutet. Stavningsvarianter
  ger separata arter, så rensa namnen först.
- En punkt på en gemensam gräns mellan två biotoper räknas i båda. En punkt räknas bara en gång
  per biotop även om biotopen består av flera polygoner.
- En art som är både skyddad och rödlistad står i båda kolumnerna.
- Listorna sorteras i svensk bokstavsordning och biotoperna numeriskt när ID:t är ett tal.
