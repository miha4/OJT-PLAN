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

- Pri izpisu `PREDVIDENE URE` makro uporablja dejansko kumulativno vrstico ur (`URE (kopija)` oziroma nastavljeno območje `ZAČETNA/KONČNA VRSTICA UR`). Vrstico zna najti po ID-ju kandidata ali po imenu, tudi če je v OJTrackerju označena kot `URE Ime Priimek`. Če je celica na izbrani datum prazna ali je na začetku planiranja relativno `0`, uporabi zadnjo znano kumulativno vrednost levo od tega datuma.
- Komentarji se dodajo kot threaded comments.
