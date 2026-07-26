# Checklist di implementazione del design system

Questa checklist chiude la Fase 1 e prepara la Fase 2/3.

## Completato in questo pacchetto

- [x] principi e identità;
- [x] palette chiara e scura;
- [x] token macchina in JSON;
- [x] tipografia;
- [x] griglia e spaziatura;
- [x] geometria dei componenti;
- [x] stati interattivi;
- [x] inventario componenti;
- [x] layout responsive;
- [x] accessibilità;
- [x] iconografia;
- [x] normalizzazione terminazioni di riga;
- [x] validatore automatico;
- [x] tavola visiva dei token.

## Criteri di uscita dalla Fase 1

- `python tools/validate_design_system.py` termina con codice 0;
- tutti i file JSON sono validi;
- i rapporti di contrasto dichiarati superano i limiti;
- il mockup approvato è conservato in `docs/assets/`;
- nessuna modifica è stata fatta alla GUI Tkinter;
- il commit è effettuato su `feature/pyside6-ui`.

## Prossimo obiettivo

Separare progressivamente il core dalla GUI, iniziando dall'inventario degli import e delle dipendenze specifiche di Tkinter/Windows, senza modificare il comportamento della baseline.
