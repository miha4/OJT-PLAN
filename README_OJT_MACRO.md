# OJT Planner VBA

## Kako uporabiti
1. Excel datoteko `OJT PLANER.xlsx` shrani kot `OJT PLANER.xlsm`.
2. Odpri VBA editor (ALT+F11) in uvozi modul `OJTPlanner.bas`.
3. Na listu `Nastavitve` izpolni stolpce za skupine in poti.
4. Za kopijo plana zaženi `Build_OJT_Plan`.
5. Za predlaganje+dodelitve zaženi `Planiraj_OJT`.

## Robni testi (priporočeno)
- Prazna pot do trackerja (C32): pričakuj napako.
- Neobstoječ list skupine v trackerju: pričakuj napako Workbooks/Worksheet.
- Kandidat brez `Xs`: brez popupa.
- Kandidat `Xs`, oba inštruktorja brez X1/X2/X3: brez popupa.
- Faza 1->2 na meji: zaradi +1 izmene rezerve ostane še v fazi 1.
- Vnos 0 v popupu: preskoči dodelitev.
- Prazen vnos izmene: preskoči dodelitev.

## Opombe
- Makro ne piše v `OJTracker`, samo bere.
- Vpisuje v `OJT Plan`, ob vsakem zagonu se list počisti.
- Komentarji se dodajo kot threaded comments.
