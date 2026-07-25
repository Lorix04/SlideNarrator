# SlideNarrator 2.9.0 — baseline stabile Tkinter

Questo documento congela la versione Tkinter utilizzata come riferimento prima della migrazione a PySide6 e Qt Quick/QML.

## Identità della baseline

- **Versione applicazione:** 2.9.0
- **Interfaccia:** Tkinter / ttkbootstrap
- **Tag Git previsto:** `v2.9.0-tkinter`
- **Branch della nuova UI:** `feature/pyside6-ui`
- **Pacchetto sorgente verificato:** `SlideNarrator_repository_SETTINGS_DOWNLOADS_TESTED.zip`
- **SHA-256 del pacchetto sorgente:** `17366c0cb6f87bb2beaad466b8e826fdc636d007c055b2421da055adac0a74eb`

Il commit esatto della baseline deve essere recuperabile con:

```bash
git rev-parse v2.9.0-tkinter
```

## Regole di conservazione

1. Il tag `v2.9.0-tkinter` non deve essere spostato o ricreato su un commit differente.
2. Le correzioni urgenti alla versione Tkinter devono partire da un branch dedicato basato sul tag.
3. Lo sviluppo PySide6/QML deve avvenire nel branch `feature/pyside6-ui`.
4. La GUI Tkinter non deve essere eliminata prima della parità funzionale della nuova interfaccia.
5. Ogni modifica grafica deve produrre screenshot di prova.
6. Ogni pacchetto deve essere verificato dopo una nuova estrazione.

## Entry point principali

- `slide_narrator_gui.py` — interfaccia Tkinter moderna.
- `slide_narrator_gui_legacy.py` — callback e logica GUI consolidata.
- `slide_narrator.py` — pipeline principale e CLI.
- `slide_narrator_batch.py` — elaborazione batch.
- `video_export.py` — esportazione video e sottotitoli.
- `voice_library.py`, `voice_clone.py`, `voice_manager.py` — gestione delle voci.

## Dipendenze e strumenti

- Dipendenze base: `requirements.txt`.
- Voci locali opzionali: `requirements-clone.txt`.
- Coqui XTTS opzionale: `requirements-xtts.txt`.
- Build Windows: `requirements-build.txt`.
- Strumenti esterni rilevati o installabili: FFmpeg, LibreOffice e Microsoft PowerPoint.

## Verifica locale

Eseguire dalla radice della repository:

```bash
python tools/verify_tkinter_baseline.py
```

La verifica controlla i file richiesti e confronta gli hash con `baseline/TKINTER_BASELINE_SHA256SUMS.txt`.

## Ripristino della baseline

Per ispezionare il codice senza modificare branch esistenti:

```bash
git switch --detach v2.9.0-tkinter
```

Per creare un branch di manutenzione:

```bash
git switch -c hotfix/tkinter-2.9.x v2.9.0-tkinter
```
