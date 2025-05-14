Milestone-CameraReport-Extended

🎯 Report telecamere Milestone ad uso investigativo

Generatore di report per Milestone XProtect, pensato per sistemi di videosorveglianza investigativa della Polizia Giudiziaria. Raccoglie metadati su Procura, Reparto, Procedimento, Stato, Connessioni e Snapshot per ogni telecamera.

Descrizione

Questo script PowerShell è sviluppato per generare un report investigativo strutturato delle telecamere registrate su un server Milestone XProtect, destinato all'uso in contesto di Polizia Giudiziaria. Il suo scopo è offrire una fotografia dettagliata dello stato operativo e della configurazione investigativa di ogni singola telecamera.

Ogni telecamera è assegnata ad una Procura, un Reparto investigativo, ed è legata ad un Procedimento Penale e a un Registro identificativo. Lo script consente di:

Estrarre e interpretare i metadati presenti nei campi ShortName e Description

Salvare uno snapshot dell'ultima immagine registrata per ciascuna telecamera

Generare un file Excel con tutte le informazioni raccolte, comprensivo di immagini

Verificare se la telecamera è attiva, registrante, e se è stata esportata per le autorità

⚠️ I dati devono essere inseriti manualmente dall'operatore all'interno dei campi ShortName e Description tramite Milestone Management Client al momento della configurazione delle telecamere. Questo garantisce che i metadati siano sempre accessibili e utilizzabili dal sistema di reportistica.

Campi gestiti e significato

Da ShortName (assegnazione investigativa):

P: Procura di riferimento

PP: Procedimento Penale

R: Registro identificativo

S: Stato operativo (es. ATTIVA, GUASTA, DISATTIVATA)

Da Description (supporto operativo):

REPARTO: Reparto investigativo assegnato

DESCRIZIONE: Descrizione del sito o contesto investigativo

TRASMISSIONE: Informazioni tecniche per la connessione della telecamera (es. modello router UMTS, numero SIM, LAN o WiFi)

RESET: Modello e numero telefonico dell'apparato GSM utilizzato per riavvio remoto in caso di blocco

ESPORTATO: Stato di esportazione immagini per PG (es. SI/NO)

NOTE: Ulteriori informazioni di campo

Funzionalità dello script

Connessione automatica al server Milestone con finestra di autenticazione

Parsing preciso dei metadati tramite regex

Salvataggio snapshot con nome basato su CameraId

Esportazione in file Excel con:

Tabelle complete

Colonna immagini snapshot

Colonna "ValiditaMetadati" per facilitare l'analisi

Requisiti

PowerShell 5.1 o successivo

Accesso al server Milestone XProtect con permessi di lettura

Moduli PowerShell:

Install-Module -Name MilestonePSTools -Scope CurrentUser -Force
Install-Module -Name ImportExcel -Scope CurrentUser -Force

Se già installati, puoi aggiornarli con:

Update-Module MilestonePSTools
Update-Module ImportExcel

Output

Cartella Snapshots_YYYYMMDD_HHmmss con immagini JPG

File Excel NOMESERVER_ReportTelecamere_YYYYMMDD_HHmmss.xlsx

Utilizzo

Apri PowerShell come amministratore

Esegui lo script Milestone-CameraReport-Extended.ps1

Autenticati al server Milestone

Attendi il completamento del report e apri il file Excel generato

Finalità investigativa

Questo strumento è pensato per l'impiego all'interno di infrastrutture investigative temporanee o permanenti, come quelle usate in attività di videosorveglianza ad uso della Polizia Giudiziaria. Rende possibile:

Monitorare lo stato e l'efficienza delle telecamere

Documentare le assegnazioni procedurali

Fornire in modo immediato i dati necessari per supporto tecnico e legale

Automatizzare la verifica di attività su larga scala

Contatti e assistenza

In caso di anomalie, il campo TRASMISSIONE fornisce modello e contatto del router (es. UMTS), mentre il campo RESET riporta il dispositivo GSM assegnato e il numero per riavvii remoti. Queste informazioni supportano interventi rapidi su apparati bloccati o non raggiungibili.

Per richieste di supporto è possibile integrare lo script con funzioni di notifica o schedulazione automatica.

Creato da: Roby
