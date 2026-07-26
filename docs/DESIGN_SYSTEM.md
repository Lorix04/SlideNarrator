# SlideNarrator Adaptive — Design System

Versione: **1.0.0**  
Stato: **approvato per il prototipo PySide6/QML**  
Ambito: Windows, macOS e Linux

## 1. Scopo

SlideNarrator Adaptive traduce il mockup approvato in regole tecniche riutilizzabili. Il design system deve rendere l'interfaccia:

- professionale e leggibile per sessioni lunghe;
- riconoscibile grazie alla palette blu petrolio e ciano;
- coerente su Windows, macOS e Linux;
- adattiva a finestre ridimensionabili e schermi ad alta densità;
- completamente utilizzabile con tastiera e tecnologie assistive;
- implementabile con PySide6, Qt Quick e QML senza copiare uno stile nativo.

## 2. Principi

### Chiarezza prima della decorazione

Ogni elemento deve spiegare una funzione, uno stato o una gerarchia. Ombre, gradienti e animazioni non devono mai competere con il contenuto.

### Una sola finestra

Le pagine principali si aprono nell'area contenuti della finestra. Dialoghi separati sono riservati a conferme, autorizzazioni, selezioni di file e operazioni realmente modali.

### Coerenza multipiattaforma

Il contenuto dell'app conserva la stessa identità su tutti i sistemi. La barra del titolo resta nativa quando possibile. Non si personalizzano gli stili nativi Windows o macOS.

### Accessibilità integrata

Focus, contrasto, dimensioni dei bersagli, nomi accessibili e ordine di navigazione fanno parte del componente, non sono correzioni successive.

### Progressione graduale

Il futuro stile QML personalizzato userà **Basic** come fallback per i controlli non ancora implementati. I componenti applicativi vengono aggiunti in ordine di priorità.

## 3. Unità e coordinate

Tutte le misure sono espresse in pixel logici indipendenti dal dispositivo, equivalenti alle coordinate usate da Qt Quick. Le risorse raster devono includere varianti ad alta risoluzione oppure essere sostituite da SVG quando appropriato.

## 4. Griglia e spaziatura

La griglia base è **4**. La maggior parte delle distanze usa multipli di **8**.

| Token | Valore | Uso |
|---|---:|---|
| `space.1` | 4 | separazioni minime |
| `space.2` | 8 | distanza icona-testo |
| `space.3` | 12 | padding compatto |
| `space.4` | 16 | distanza standard |
| `space.5` | 20 | gap tra blocchi medi |
| `space.6` | 24 | margine pagina largo |
| `space.8` | 32 | sezioni principali |
| `space.10` | 40 | separazioni ampie |
| `space.12` | 48 | hero e stati vuoti |
| `space.16` | 64 | separazioni eccezionali |

Non usare valori casuali se esiste un token equivalente.

## 5. Geometria

| Elemento | Valore |
|---|---:|
| Sidebar espansa | 228 |
| Sidebar compatta | 72 |
| Barra di stato | 30 |
| Pulsanti e campi | 40 |
| Pulsante icona | 36 × 36 |
| Bersaglio minimo applicativo | 36 × 36 |
| Icona navigazione | 20 |
| Icona tecnologia/file | 24 |
| Raggio pulsante | 8 |
| Raggio card | 12 |
| Raggio dialogo | 16 |
| Bordo standard | 1 |
| Anello focus | 2 |

La finestra raccomandata è **1180 × 720**. Il minimo supportato è **800 × 640**; a questa dimensione il contenuto usa una sola colonna e scorrimento verticale.

## 6. Palette chiara

| Ruolo | Colore |
|---|---|
| Sfondo | `#F4F7FA` |
| Superficie | `#FFFFFF` |
| Superficie alternativa | `#FBFCFD` |
| Sidebar | `#082F3D` |
| Primario | `#0E4B5A` |
| Accento | `#18B6D2` |
| Accento hover | `#0EA3BF` |
| Testo principale | `#17212B` |
| Testo secondario | `#536574` |
| Bordo | `#D4DEE6` |
| Successo | `#147D4B` |
| Avviso | `#9A5B00` |
| Errore | `#B42318` |
| Informazione | `#0D6B83` |

Il testo sopra il ciano usa `#082F3D` o `#17212B`, non bianco: il bianco sul ciano principale non raggiunge il contrasto stabilito.

## 7. Palette scura

| Ruolo | Colore |
|---|---|
| Sfondo | `#0A141B` |
| Superficie | `#10222C` |
| Superficie alternativa | `#132A35` |
| Sidebar | `#071F29` |
| Primario | `#24BBD2` |
| Accento | `#36C6DA` |
| Testo principale | `#F2F7FA` |
| Testo secondario | `#B9C7D0` |
| Bordo | `#29404C` |
| Successo | `#69D49A` |
| Avviso | `#F2BB5E` |
| Errore | `#FF8A80` |
| Informazione | `#78D5E4` |

Il tema scuro non è un'inversione automatica. Ombre, bordi e superfici hanno token dedicati.

## 8. Tipografia

Famiglia preferita, in ordine di disponibilità:

1. Inter;
2. Segoe UI Variable / Segoe UI;
3. SF Pro Text;
4. Noto Sans;
5. sans-serif di sistema.

Non distribuire file di font nel repository. La selezione avviene tramite il font installato sul sistema.

| Ruolo | Dimensione | Peso | Interlinea |
|---|---:|---:|---:|
| Titolo pagina | 28 | 600 | 36 |
| Titolo sezione | 20 | 600 | 28 |
| Titolo card | 17 | 600 | 24 |
| Corpo | 14 | 400 | 20 |
| Corpo secondario | 13 | 400 | 18 |
| Didascalia | 12 | 400 | 16 |
| Badge | 12 | 600 | 16 |

Il testo non deve essere trasformato in immagini. Le etichette devono poter crescere senza sovrapporsi.

## 9. Superfici ed elevazione

- Sfondo: nessuna ombra.
- Card: bordo da 1 px e ombra molto leggera.
- Menu, popup e dialoghi: ombra più evidente, sempre morbida.
- Massimo tre livelli visivi: sfondo, superficie, overlay.
- Nessun blur pesante, glow, neon o effetto vetro liquido.

## 10. Iconografia

Esistono due famiglie:

1. **Navigazione e comandi**: lineari, monocromatiche, 20 px, stesso spessore.
2. **File e tecnologie**: semplici, colorate, 24 px, usate solo quando il colore aiuta il riconoscimento.

L'icona ufficiale SlideNarrator non deve essere ridisegnata. Usare gli asset esistenti in `assets/`.

## 11. Stati interattivi

Ogni controllo supporta almeno:

- normale;
- hover;
- premuto;
- focus tastiera;
- disabilitato;
- caricamento, quando applicabile;
- errore, quando applicabile.

L'hover non sostituisce il focus. Il focus usa un anello esterno da 2 px, separato dal bordo del componente.

## 12. Pulsanti

### Primario

Una sola azione primaria per area. Fondo primario o ciano; il colore del testo dipende dal contrasto del fondo.

### Secondario

Sfondo trasparente o superficie, bordo visibile, testo primario.

### Distruttivo

Usare solo per azioni irreversibili o potenzialmente dannose. Richiede conferma quando la perdita non è recuperabile.

### Icona

Il simbolo può essere 16–20 px, ma il contenitore interattivo resta almeno 36 × 36.

## 13. Card

- raggio 12;
- bordo 1;
- padding 16 o 20;
- titolo sempre testuale;
- l'intera card è cliccabile solo quando rappresenta una singola destinazione;
- non combinare card cliccabile, freccia e pulsante per la stessa azione;
- evitare card annidate.

## 14. Stato e feedback

Ogni stato usa colore, testo e quando utile un'icona.

Esempi preferiti:

- `Pronto`;
- `Installato`;
- `Connessione richiesta`;
- `Non rilevato`;
- `In elaborazione`;
- `Completato`;
- `Errore`.

Le attività lunghe mostrano progresso, passaggio corrente, possibilità di annullamento e risultato finale.

## 15. Movimento

| Token | Durata | Uso |
|---|---:|---|
| Rapido | 90 ms | hover, pressione |
| Standard | 150 ms | transizioni di stato |
| Lento | 220 ms | apertura pannelli e cambi layout |

Usare curve morbide e brevi. Disabilitare le animazioni non essenziali quando è attiva la modalità di movimento ridotto.

## 16. Contenuto e microcopy

- Italiano corretto e diretto.
- Verbi all'infinito o imperativo coerente.
- Titoli brevi, descrizioni utili.
- Evitare gergo tecnico quando non necessario.
- Gli errori devono spiegare cosa è successo e come proseguire.
- Non usare `Disponibile` quando il significato reale è `Pronto`, `Supportato` o `Connessione richiesta`.

## 17. Implementazione QML

La prima implementazione userà Qt Quick Controls con uno stile personalizzato basato su **Basic**. Lo stile va selezionato prima del caricamento dei file QML che importano Qt Quick Controls.

Ordine iniziale dei componenti:

1. token e tema;
2. pulsanti e pulsanti icona;
3. card e badge;
4. sidebar;
5. righe progetto/tecnologia;
6. barra di avanzamento;
7. campi e selettori;
8. dialoghi e notifiche.

## 18. Governance

Una modifica al design system richiede:

1. aggiornamento dei token;
2. aggiornamento della documentazione;
3. esecuzione di `python tools/validate_design_system.py`;
4. aggiornamento della tavola visiva;
5. screenshot prima/dopo quando esiste una UI;
6. commit separato e descrittivo.
