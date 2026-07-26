# Token SlideNarrator Adaptive

`slidenarrator.tokens.json` è la fonte macchina del design system.

Regole:

- i nomi dei token restano in inglese per l'uso nel codice;
- i valori geometrici sono pixel logici Qt;
- non duplicare colori direttamente nei file QML;
- i temi chiaro e scuro condividono gli stessi ruoli semantici;
- ogni modifica deve superare `python tools/validate_design_system.py`.

Nella futura implementazione, questi token verranno esposti a QML tramite un singleton `Theme.qml` oppure un oggetto Python registrato come singleton. La scelta definitiva avverrà nella fase di scheletro PySide6/QML.
