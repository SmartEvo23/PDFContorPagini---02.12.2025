# FișiereContorPagini

Aplicație desktop Windows (WinForms, .NET Framework 4.7.2) care numără automat paginile din fișiere PDF și Word dintr-un folder, inclusiv subfoldere — utilă pentru facturare sau evidență printuri/copii.

## Funcționalități

- Numără paginile din `.pdf`, `.docx`, `.doc`, `.docm`, `.dotx`, `.dot`.
- Scanare recursivă (caută și în subfoldere), rulată asincron — interfața rămâne funcțională în timpul scanării, cu buton de anulare.
- Progres live: fișiere procesate + total pagini de până acum.
- PDF-urile sunt procesate în paralel, pe mai multe nuclee, pentru viteză.
- Listă sortabilă a fișierelor scanate, cu pagini și status (inclusiv fișierele care au eșuat la citire, cu mesajul de eroare).
- Export CSV al rezultatelor.
- Calculator de cost (preț per pagină → total de plată).
- Rulează portabil, fără instalare.

## Cerințe

- Windows cu .NET Framework 4.7.2 (prezent implicit pe majoritatea instalărilor Windows moderne).
- Microsoft Word instalat, doar dacă vrei să numeri și fișiere Word (nu e necesar pentru PDF).

## Utilizare

Vezi [GHID_UTILIZARE.md](PDFContorPagini/GHID_UTILIZARE.md) pentru instrucțiuni pas cu pas, sau apasă butonul **"Ajutor"** din aplicație.

## Build din sursă

1. Deschide `FisiereContorPagini.slnx` în Visual Studio.
2. Lasă Visual Studio să restaureze pachetele NuGet (automat la build).
3. **Build → Rebuild Solution** (configurația Release recomandată pentru distribuire).
4. Executabilul rezultă în `PDFContorPagini/bin/Release/FisiereContorPagini.exe`, împreună cu DLL-urile de care are nevoie — poți muta tot folderul oriunde, fără instalator.

## Structură proiect

Codul sursă e în folderul [`PDFContorPagini/`](PDFContorPagini/).
