# Sicurezza e distribuzione dei binari

## Build pubbliche

Le release Windows devono essere create da un ambiente pulito usando `Build_EXE_SlideNarrator.bat`. La build pubblica predefinita usa PyInstaller in modalità **onedir**, disabilita UPX, aggiunge metadati di versione e richiede il normale livello `asInvoker`.

Non pubblicare l'eseguibile sperimentale `onefile`: l'autoestrazione e il bootloader condiviso possono aumentare i rilevamenti euristici.

## Firma

Prima di una release stabile, firma `SlideNarrator.exe` e l'installer con un certificato Authenticode attendibile e con la stessa identità di editore a ogni versione. Il file `Firma_RELEASE_SlideNarrator.bat` usa il certificato presente nell'archivio Windows indicato dalla variabile `SLIDENARRATOR_CERT_SUBJECT`. Un certificato autofirmato è adatto solo ai test interni.

## Verifica

- Esegui `Verifica_RELEASE_SlideNarrator.bat`.
- Conserva e pubblica gli hash SHA-256.
- Testa la release su Windows 10 e 11 puliti, senza Python.
- Invia ai produttori antivirus soltanto il binario finale, già firmato e corrispondente alla release.
- Non chiedere agli utenti di disattivare l'antivirus o aggiungere esclusioni.

## Segnalazione falsi positivi

Per Microsoft Defender usa il portale Microsoft Security Intelligence e indica che il file è stato rilevato erroneamente. Includi il link al tag GitHub, l'hash SHA-256, il comando di build, la versione di PyInstaller e la firma digitale.

## Vulnerabilità

Non pubblicare dettagli sensibili in una issue pubblica. Contatta il manutentore della repository con una descrizione riproducibile, versione interessata e impatto stimato.
