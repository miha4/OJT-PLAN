# OJT Planner VBA

## Kako uporabiti
1. Excel datoteko `OJT PLANER.xlsx` shrani kot `OJT PLANER.xlsm`.
2. Odpri VBA editor (ALT+F11) in uvozi modul `OJTPlanner.bas`.
3. Na listu `Nastavitve` izpolni stolpce za skupine in poti.
4. Za kopijo plana zaženi `Build_OJT_Plan`.
5. Za predlaganje+dodelitve zaženi `Planiraj_OJT`. Makro skupine bere po vrstnem redu stolpcev na listu `Nastavitve` in dodelitve izvaja samo za skupine, kjer je `PLANIRAJ = DA`; skupine z `NE` samo prikaže v kopiji plana in jih pri dodeljevanju preskoči.

## Robni testi (priporočeno)
- Prazna pot do trackerja (C32): pričakuj napako.
- Neobstoječ list skupine v trackerju: pričakuj napako Workbooks/Worksheet.
- Kandidat brez `Xs`: brez popupa.
- Kandidat `Xs`, oba inštruktorja brez X1/X2/X3: brez popupa, v zaključnem opozorilu se izpiše diagnostika `brez inštruktorjev`.
- Kandidat `Xs`, ki nima vrstice ur v območju kumulativnih ur: makro ga ne preskoči več, ampak nadaljuje z 0 začetnimi urami in to zabeleži v diagnostiki.
- Faza 1->2 na meji: zaradi +1 izmene rezerve ostane še v fazi 1.
- Vnos 0 v popupu: preskoči dodelitev.
- Prazen vnos izmene: preskoči dodelitev.

## Opombe
- Makro ne piše v `OJTracker`, samo bere.
- `Build_OJT_Plan` prikaže vse veljavno nastavljene skupine, tudi če imajo `PLANIRAJ = NE`.
- `Planiraj_OJT` najprej osveži prikaz vseh skupin, nato planira samo aktivne skupine (`PLANIRAJ = DA`) po zaporedju stolpcev v nastavitvah.
- Če aktivna skupina nima nobenega možnega predloga, zaključno okno pokaže diagnostiko po skupinah: koliko `Xs` je našel, koliko kandidatov nima vrstice ur in koliko jih nima prostega inštruktorja.
- Vpisuje v `OJT Plan`, ob vsakem zagonu se list počisti.
- Če list uporablja oznake `KANDIDAT #...`, makro planira samo te vrstice; `Xs` pri inštruktorju oziroma osebi, ki se šola v drugi skupini, se v tej skupini ne obravnava kot kandidat.
- Seznam možnih inštruktorjev je dedupliciran po ID-ju, zato se isti mentor ne ponudi dvakrat, tudi če je prisoten v splošnem seznamu in še v primarni/sekundarni vrstici kandidata.
- Mentor, ki je v istem zagonu planiranja že dodeljen na isti datum, se pri naslednjih kandidatih na ta datum ne ponudi več.
- Vnos `B` razveljavi zadnjo dodelitev: vrne prvotne izmene v celicah kandidata in inštruktorja, odstrani komentarje makra, odšteje dodane ure in pobriše vpis ur v panelu.

- Pri izpisu `PREDVIDENE URE` makro uporablja dejansko kumulativno vrstico ur (`URE (kopija)` oziroma nastavljeno območje `ZAČETNA/KONČNA VRSTICA UR`). Vrstico ur prepozna tudi po oznaki `URE <ime kandidata>`, zato navadne urniške vrstice kandidata ne zamenja več za kumulativne ure. Če je celica na izbrani datum prazna ali je na začetku novega mesečnega bloka relativno `0`, uporabi zadnjo znano kumulativno vrednost levo od tega datuma, tudi če je ta pred nastavljenim začetkom trenutnega bloka.
- Komentarji se dodajo kot threaded comments.
