# Layout responsive

## Obiettivo

SlideNarrator è un'app desktop ridimensionabile. Il responsive non consiste nel ridurre indefinitamente tutti gli elementi: quando lo spazio non basta, la struttura cambia.

Le soglie sono espresse in pixel logici Qt, non in pixel fisici del monitor.

## Soglie

| Modalità | Larghezza finestra | Sidebar | Home |
|---|---:|---:|---|
| Wide | ≥ 1480 | 228, espansa | due colonne 60/40 |
| Standard | 1180–1479 | 220–228, espansa | due colonne 56/44 |
| Compact | 960–1179 | 72, icone | azioni su due colonne; contenuti principali impilati |
| Narrow | 800–959 | 64–72 o overlay | una colonna con scorrimento |

Dimensione minima supportata: **800 × 640**. Dimensione raccomandata: **1180 × 720**.

## Wide

- margine pagina 24;
- gap colonne 18–20;
- colonna progetti 60%;
- colonna attività/strumenti 40%;
- la card voce occupa la colonna sinistra;
- nessuna scrollbar orizzontale.

## Standard

- margine pagina 20;
- gap 16–18;
- le descrizioni secondarie possono andare su due righe;
- le card superiori mantengono due colonne;
- il menu overflow raccoglie le azioni meno frequenti dell'header.

## Compact

- sidebar a sole icone con tooltip;
- margine pagina 16;
- le due azioni principali possono restare affiancate se ciascuna mantiene almeno 340 di larghezza;
- Progetti recenti, Attività corrente, Strumenti e Voce preferita diventano una sequenza verticale;
- le righe recenti nascondono la data secondaria prima di comprimere il nome.

## Narrow

- una sola colonna;
- sidebar compatta oppure pannello overlay richiamabile dall'header;
- illustrazioni delle card azione ridotte o nascoste;
- tutte le azioni restano testuali;
- contenuto in `Flickable` verticale;
- la barra di stato mostra solo stato principale e progresso essenziale.

## Regole di priorità

Quando manca spazio, procedere in questo ordine:

1. ridurre margini e gap entro i token previsti;
2. comprimere la sidebar;
3. spostare azioni secondarie nel menu overflow;
4. riorganizzare due colonne in una;
5. ridurre o nascondere illustrazioni decorative;
6. consentire più righe ai testi;
7. introdurre scorrimento verticale.

Non fare:

- ridurre il corpo testo sotto 12;
- ridurre i controlli sotto 36 × 36;
- tagliare etichette di azioni senza tooltip;
- creare scorrimento orizzontale della pagina;
- sovrapporre controlli.

## Implementazione QML prevista

- `RowLayout` e `ColumnLayout` per la struttura;
- `Layout.minimumWidth`, `Layout.preferredWidth` e `Layout.fillWidth`;
- `LayoutItemProxy` quando un componente deve cambiare gerarchia;
- `Flickable` per la modalità a colonna singola;
- `Loader` per illustrazioni e contenuti non necessari nelle modalità compatte;
- proprietà centralizzata `layoutMode` calcolata dalla larghezza utile.

## Matrice di prova

| Risoluzione | Scala | Risultato atteso |
|---|---:|---|
| 1920 × 1200 | 100% | Wide |
| 1600 × 900 | 100% | Wide |
| 1366 × 768 | 100% | Standard |
| 1180 × 720 | 100% | Standard minimo raccomandato |
| 1100 × 700 | 100% | Compact |
| 960 × 640 | 100% | Compact limite |
| 800 × 640 | 100% | Narrow minimo |
| 1366 × 768 | 150% simulato | nessun taglio o sovrapposizione |
| 1920 × 1080 | 200% simulato | navigazione e contenuti ancora raggiungibili |

Ogni pagina futura deve produrre screenshot almeno nelle modalità Wide, Standard, Compact e tema scuro.
