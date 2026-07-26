# Inventario componenti UI

Questo documento definisce i componenti QML da implementare e il loro contratto minimo. I nomi sono proposte stabili per la futura cartella `ui/components/`.

## Componenti fondamentali

### `AppButton.qml`

Varianti: `primary`, `secondary`, `quiet`, `danger`.

Proprietà minime:

- `text`;
- `iconName`;
- `variant`;
- `busy`;
- `enabled`;
- `accessibleDescription`.

Regole:

- altezza 40;
- padding orizzontale 16;
- raggio 8;
- focus esterno da 2;
- il testo non viene nascosto nello stato busy;
- Enter e Spazio attivano il pulsante.

### `AppIconButton.qml`

- contenitore 36 × 36;
- icona 18–20;
- tooltip obbligatorio;
- nome accessibile obbligatorio;
- forma quadrata arrotondata o circolare secondo il contesto.

### `AppCard.qml`

Proprietà:

- `title`;
- `subtitle`;
- `clickable`;
- `selected`;
- slot per contenuto e azioni.

Una card cliccabile deve avere un'unica azione primaria. Se contiene più comandi, la superficie non è cliccabile.

### `AppBadge.qml`

Varianti: `neutral`, `info`, `success`, `warning`, `error`.

- testo obbligatorio;
- icona facoltativa;
- altezza minima 24;
- raggio pillola;
- nessun significato affidato esclusivamente al colore.

### `AppProgressBar.qml`

- valore 0–1;
- stato determinato o indeterminato;
- etichetta percentuale facoltativa;
- descrizione accessibile del progresso;
- non usare animazioni continue aggressive.

## Navigazione

### `AppSidebar.qml`

Stati: `expanded`, `compact`, `overlay`.

- larghezza 228 in modalità espansa;
- larghezza 72 in modalità compatta;
- Impostazioni ancorata in basso;
- ordine tastiera uguale all'ordine visivo;
- il logo non è un pulsante salvo funzione esplicita.

### `SidebarItem.qml`

- altezza minima 44;
- icona 20;
- indicatore attivo verticale da 3;
- testo bianco nello stato attivo;
- tooltip in modalità compatta;
- stato attivo esposto alle tecnologie assistive.

### `PageHeader.qml`

Contiene titolo, sottotitolo e azioni. A larghezze ridotte le azioni non essenziali confluiscono nel menu overflow.

## Contenuti Home

### `PrimaryActionCard.qml`

- illustrazione compatta;
- titolo;
- descrizione;
- una sola azione;
- altezza adattiva;
- l'illustrazione non supera il 35% della larghezza utile.

### `RecentProjectRow.qml`

- icona tipo file;
- nome progetto;
- trasformazione;
- data;
- badge stato;
- menu contestuale 36 × 36.

Troncamento consentito solo sul nome del progetto, con tooltip del valore completo.

### `CurrentActivityCard.qml`

- titolo progetto;
- fase;
- percentuale;
- barra di avanzamento;
- conteggio slide;
- tempo residuo stimato;
- azione dettagli.

Il tempo residuo deve essere marcato come stima.

### `TechnologyRow.qml`

- pittogramma 24;
- nome;
- descrizione breve;
- stato testuale;
- eventuale azione.

Stati consigliati: `Pronto`, `Installato`, `Connessione richiesta`, `Non rilevato`, `Errore`.

### `VoiceCard.qml`

- avatar astratto;
- nome voce;
- lingua, genere e stile;
- motore;
- anteprima;
- apertura libreria.

La forma d'onda compare soltanto durante la riproduzione.

## Input e configurazione

### `AppTextField.qml`

- label esterna persistente;
- testo di supporto;
- errore testuale;
- pulsante cancella solo quando utile;
- nome e descrizione accessibili;
- non usare il placeholder come unica etichetta.

### `AppComboBox.qml`

- altezza 40;
- selezione da tastiera;
- ricerca per elenchi lunghi;
- testo selezionato sempre leggibile;
- popup entro i bordi della finestra.

### `AppSwitch.qml`

- etichetta cliccabile;
- stato testuale quando l'effetto non è immediato;
- focus visibile;
- semantica `checked` esposta.

### `AppSlider.qml`

- valore corrente visibile;
- controllo da frecce;
- passo definito;
- minimo e massimo accessibili;
- campo numerico affiancato per regolazioni precise quando necessario.

### `FilePicker.qml`

- campo percorso;
- pulsante Sfoglia;
- drag-and-drop;
- tipi accettati indicati;
- errore leggibile;
- percorso lungo copiabile.

## Feedback e overlay

### `AppDialog.qml`

- titolo;
- descrizione;
- pulsante predefinito;
- pulsante Annulla;
- focus intrappolato correttamente;
- Escape chiude solo quando sicuro;
- al termine il focus torna al comando che ha aperto il dialogo.

### `AppToast.qml`

Usare per feedback non bloccanti. Non usare come unico luogo per errori critici o informazioni che devono essere consultate più tardi.

### `EmptyState.qml`

- illustrazione discreta;
- titolo;
- spiegazione;
- azione utile;
- nessun tono promozionale.

### `ErrorState.qml`

- titolo umano;
- causa sintetica;
- passaggio successivo;
- dettagli tecnici espandibili e copiabili.

## Stati obbligatori per la pagina di catalogo

Ogni componente deve essere mostrato in:

- tema chiaro e scuro;
- normale;
- hover;
- premuto;
- focus;
- disabilitato;
- errore o busy, se applicabile.
