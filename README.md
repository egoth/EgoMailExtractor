# EgoMailExtractor Community 1.47.5

## Novità v1.45.2

Novità principali
- Pipeline persistente in tre fasi: Lettura, Estrazione e Consolidamento.
- La Lettura applica già i criteri dei Tipi mail valutabili dagli header
  (mittente, oggetto, data, destinatario, Message-ID) e salva su disco la
  sequenza ordinata delle sole mail candidate.
- Se la Lettura è completa e l'esecuzione viene interrotta, alla ripresa il
  file JSON viene riutilizzato senza ripetere la scansione.
- L'Estrazione non modifica direttamente l'Excel: scrive un journal JSON
  atomico delle operazioni da consolidare, compresa l'indicazione di
  riconciliazione e le relative chiavi. Il journal viene salvato a ogni
  operazione estratta.
- Consolidamento transazionale: Excel originale preservato, copia temporanea,
  applicazione in memoria delle modifiche, un solo salvataggio finale e
  sostituzione del file definitivo solo a completamento riuscito.
- In caso di interruzione del Consolidamento il temporaneo viene eliminato e
  alla ripresa si ricomincia dall'originale usando lo stesso journal.
- Tre barre di avanzamento: Lettura, Estrazione, Consolidamento.
- La home può mostrare fino a 6 valori scelti tra formule finali, formule
  riepilogative e campi calcolati.
- Le formule riepilogative possono eseguire Somma/Conta condizionata da un
  campo data nell'ultimo mese o anno, scegliendo tra:
    * finestra mobile (30 / 365 giorni);
    * periodo di calendario corrente (mese / anno corrente).


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
