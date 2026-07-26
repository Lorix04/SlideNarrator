# Rapporto test — SlideNarrator Adaptive 1.0.0

Data: 26 luglio 2026  
Branch di destinazione: `feature/pyside6-ui`  
Ambito: documentazione e token del design system; nessuna modifica alla GUI Tkinter o al motore applicativo.

## Risultato

**SUPERATO**

## Controlli eseguiti

- 7 documenti obbligatori presenti;
- JSON dei token valido;
- 16 coppie di contrasto verificate;
- 6 unit test superati;
- breakpoint strettamente crescenti;
- bersaglio minimo applicativo di 36 × 36;
- altezza controlli minima di 40;
- asset ufficiali rilevati;
- mockup approvato conservato come riferimento;
- estrazione e test del pacchetto overlay completati;
- estrazione e test del pacchetto completo completati;
- nessun file `.pyc` o `__pycache__` incluso;
- nessun file Python applicativo della baseline modificato.

## Rapporti di contrasto

| Tema | Coppia | Rapporto | Limite |
|---|---|---:|---:|
| Chiaro | testo principale / superficie | 16,29:1 | 4,5:1 |
| Chiaro | testo secondario / superficie | 6,03:1 | 4,5:1 |
| Chiaro | testo su primario | 9,67:1 | 4,5:1 |
| Chiaro | testo scuro su accento | 5,83:1 | 4,5:1 |
| Chiaro | successo | 4,64:1 | 4,5:1 |
| Chiaro | avviso | 4,96:1 | 4,5:1 |
| Chiaro | errore | 5,76:1 | 4,5:1 |
| Chiaro | informazione | 5,49:1 | 4,5:1 |
| Scuro | testo principale / superficie | 15,12:1 | 4,5:1 |
| Scuro | testo secondario / superficie | 9,43:1 | 4,5:1 |
| Scuro | testo su primario | 8,07:1 | 4,5:1 |
| Scuro | testo su accento | 8,28:1 | 4,5:1 |
| Scuro | successo | 7,17:1 | 4,5:1 |
| Scuro | avviso | 7,93:1 | 4,5:1 |
| Scuro | errore | 6,58:1 | 4,5:1 |
| Scuro | informazione | 8,08:1 | 4,5:1 |

## Comandi di riproduzione

```powershell
python .\tools\validate_design_system.py
python -m unittest tests.test_design_system -v
```

## Nota

Questi test verificano la specifica e i token. I test visivi della UI QML inizieranno dopo la creazione dello scheletro PySide6/Qt Quick.
