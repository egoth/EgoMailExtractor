# EgoMailExtractor Community 1.40.2

## Novità v1.40

- I pacchetti di esportazione dei profili includono anche l'icona del profilo; l'importazione la ripristina.
- **File → Apri → Cartella workspace** apre la cartella dati dell'applicazione.
- Nella modifica email è visibile il selettore **Protocollo**, attualmente con la sola scelta **IMAP**.
- La Pro aggiunge **Profilo → Tabella totali…** per configurare righe, colonne, raggruppamenti anno/mese e somme dei campi numerici nel foglio Excel `Tabella totali`.
- Nella Community le funzioni Tabella totali e Wizard sono visibili ma rimandano alla versione Pro.
- Splash screen e Informazioni su non mostrano il nome dell'autore.


EgoMailExtractor Community è l'edizione open source di EgoMailExtractor, distribuita con licenza Apache License 2.0.

## Limiti dell'edizione Community

- una sola email configurata/attiva;
- un solo profilo di estrazione;
- nessun Wizard visuale.

Non sono limitati il numero di campi, tipi mail, regole, messaggi elaborati o esportazioni. Le regole possono essere create, modificate ed eliminate dalla schermata completa.

## Build Windows

Eseguire `BUILD_COMPLETO.bat`. Lo script crea l'ambiente virtuale, installa le dipendenze, esegue i controlli e genera la distribuzione tramite PyInstaller.

## Licenza

Vedi `LICENSE.txt` (Apache License 2.0) e `NOTICE`.
