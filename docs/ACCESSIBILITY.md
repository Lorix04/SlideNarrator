# Accessibilità

Obiettivo minimo interno: **WCAG 2.2 livello AA**, applicato ai principi trasferibili a un'applicazione desktop. Alcune scelte, come l'anello di focus, seguono volontariamente indicazioni più forti.

## Tastiera

- Tutte le funzioni principali devono essere raggiungibili senza mouse.
- L'ordine Tab coincide con l'ordine visivo e logico.
- Enter e Spazio attivano pulsanti e controlli.
- Escape chiude menu e dialoghi quando non causa perdita di dati.
- Le frecce gestiscono menu, liste, slider e selettori secondo le convenzioni della piattaforma.
- Nessun focus intrappolato fuori dai dialoghi modali.
- Dopo la chiusura di un overlay, il focus torna al comando di origine.

## Indicatore di focus

- Sempre visibile quando il controllo riceve focus da tastiera.
- Anello esterno di 2 px.
- Separazione di 2 px dal bordo quando possibile.
- Contrasto di almeno 3:1 rispetto allo stato non focalizzato.
- Mai rimuovere il focus solo per ragioni estetiche.

## Dimensioni interattive

Il design system usa un bersaglio applicativo minimo di **36 × 36**, superiore al riferimento minimo WCAG di 24 × 24.

- Pulsanti testuali: altezza 40.
- Pulsanti icona: 36 × 36.
- Elementi sidebar: altezza minima 44.
- Menu contestuali: contenitore 36 × 36 anche con icona piccola.
- Le righe cliccabili devono avere almeno 44 di altezza.

## Contrasto

- Testo normale: almeno 4,5:1.
- Testo grande: almeno 3:1.
- Controlli, indicatori, focus e grafica necessaria: almeno 3:1 contro colori adiacenti.
- Gli stati non dipendono soltanto dal colore.
- Il ciano principale usa testo scuro; il bianco non è consentito sul ciano `#18B6D2` per etichette normali.

I rapporti definiti nei token vengono controllati da `tools/validate_design_system.py`.

## Qt Accessible

Per controlli personalizzati e contenuti informativi usare le proprietà QML:

- `Accessible.name`;
- `Accessible.description`;
- `Accessible.role`;
- `Accessible.focusable`;
- `Accessible.checked`;
- `Accessible.onPressAction` o azione equivalente.

Ogni azione invocata da una tecnologia assistiva deve produrre lo stesso risultato del clic o della pressione da tastiera.

## Testi e nomi

- Il nome accessibile del controllo deve contenere la sua etichetta visibile.
- Le icone senza testo hanno tooltip e nome accessibile.
- I campi usano una label persistente; il placeholder non è l'unica etichetta.
- Gli errori sono associati al campo e letti dopo la validazione.
- Le percentuali includono il contesto, ad esempio `Generazione audio, 68 percento`.

## Audio e movimento

- Nessun audio parte automaticamente durante la navigazione.
- L'anteprima vocale ha Play, Pausa e Stop.
- Le animazioni non essenziali si disattivano con l'opzione movimento ridotto.
- Nessun lampeggio rapido o elemento pulsante continuo.

## Scaling e alta densità

Qt usa coordinate indipendenti dal dispositivo, ma il layout deve essere verificato con fattori di scala reali o simulati:

```text
QT_SCALE_FACTOR=1.25
QT_SCALE_FACTOR=1.5
QT_SCALE_FACTOR=2
```

Le risorse raster devono essere nitide a 100%, 150% e 200%. Preferire SVG per icone lineari.

## Screen reader

Matrice prevista:

- Windows: Narrator e, quando possibile, NVDA;
- macOS: VoiceOver;
- Linux: Orca su una distribuzione supportata.

Verifiche minime:

1. titolo finestra e pagina annunciati;
2. sidebar annunciata come navigazione;
3. stato selezionato comunicato;
4. pulsanti nominati correttamente;
5. progresso annunciato senza ripetizioni continue;
6. errori e completamenti comunicati;
7. dialoghi con titolo, descrizione e azione predefinita.

## Checklist per ogni componente

- [ ] raggiungibile da tastiera;
- [ ] focus visibile;
- [ ] nome accessibile;
- [ ] ruolo corretto;
- [ ] stato esposto;
- [ ] contrasto verificato;
- [ ] bersaglio sufficiente;
- [ ] testo ridimensionabile;
- [ ] tema chiaro e scuro;
- [ ] comportamento con movimento ridotto.
