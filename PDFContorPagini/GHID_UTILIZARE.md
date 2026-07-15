# Ghid de utilizare — FișiereContorPagini

Aplicație pentru numărat automat paginile din fișiere PDF și Word dintr-un folder (inclusiv subfoldere), utilă pentru facturare/evidență printuri și copii.

## Pornire

Nu necesită instalare. Deschide direct `FisiereContorPagini.exe` (din `bin\Release` după un build, sau din arhiva descărcată de pe GitHub).

Pentru numărarea fișierelor Word (.doc, .docx, .docm, .dotx, .dot) este necesar ca **Microsoft Word să fie instalat** pe calculator. Pentru PDF nu e nevoie de niciun program suplimentar.

## Pași de utilizare

1. **Bifează tipurile de fișiere** pe care vrei să le numeri: `.pdf`, `.docx`, `.doc`, `.docm`, `.dotx`, `.dot`. Poți bifa oricâte simultan.

2. Apasă **"Selectează locația"** și alege folderul care conține fișierele. Aplicația caută automat și în toate subfolderele.

3. Scanarea pornește imediat. În timp ce rulează, vezi live numărul de fișiere procesate și totalul de pagini de până acum, plus o bară de progres cu procent. Dacă durează prea mult sau ai ales folderul greșit, apasă **"Anulează scanarea"**.

4. La final apare un mesaj cu totalul de pagini, iar eticheta din stânga se actualizează.

5. Apasă **"Lista fișierelor scanate"** pentru a vedea fiecare fișier în parte, cu numărul lui de pagini. Poți da click pe antetul unei coloane (Pagini, Fișier, Status) ca să sortezi lista. Fișierele care nu au putut fi citite (corupte, protejate cu parolă etc.) apar și ele, cu mesajul exact de eroare în coloana Status — nu sunt ascunse.

6. Din aceeași fereastră, butonul **"Export CSV"** salvează rezultatul într-un fișier `.csv` pe care îl poți deschide direct în Excel.

7. Butonul **"Ajutor"** din colțul din dreapta sus deschide oricând acest ghid direct din aplicație.

## Întrebări frecvente

**De ce nu apare niciun fișier Word?** Verifică dacă Microsoft Word e instalat pe acest calculator — fără el, aplicația nu poate deschide fișierele Word (PDF-urile funcționează independent de Word).

**Un fișier apare cu eroare în listă — ce înseamnă?** Fișierul e corupt, protejat cu parolă, sau într-un format neașteptat. Mesajul exact de eroare apare în coloana Status din "Lista fișierelor scanate".

**Aplicația pare blocată în timpul scanării?** Pentru foldere foarte mari cu multe fișiere Word, scanarea poate dura — Word Automation procesează fișierele Word unul câte unul (nu se pot paraleliza, spre deosebire de PDF-uri, care sunt procesate simultan pe mai multe nuclee). Poți oricând apăsa "Anulează scanarea".
