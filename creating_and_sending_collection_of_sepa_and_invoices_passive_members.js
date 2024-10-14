// Hilfsfunktion zum Abrufen einer Datei nach Namen
/**
 * Sucht und gibt eine Datei innerhalb eines Ordners anhand ihres Namens zurück.
 *
 * @param {Folder} folder - Der Google Drive-Ordner, in dem gesucht werden soll.
 * @param {string} fileName - Der Name der gesuchten Datei.
 * @return {File|null} - Die gefundene Datei oder null, wenn keine Datei gefunden wurde.
 */
function getFileByName(folder, fileName) {
    // Ruft alle Dateien im angegebenen Ordner mit dem angegebenen Namen ab
    const files = folder.getFilesByName(fileName);
    
    // Gibt die erste gefundene Datei zurück oder null, falls keine Datei vorhanden ist
    return files.hasNext() ? files.next() : null;
  }
  
  
  // Funktion zum Schreiben von Text in eine spezifische Zelle im "Mission_Control" Blatt
  /**
   * Schreibt einen gegebenen Text in eine bestimmte Zelle im "Mission_Control" Blatt.
   *
   * @param {string} zellenAdresse - Die Adresse der Zielzelle (z.B. "A1").
   * @param {string} text - Der Text, der in die Zelle geschrieben werden soll.
   */
  function schreibeTextInZelleInMissionControl(zellenAdresse, text) {
    // Zugriff auf das aktuell geöffnete Spreadsheet
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Zugriff auf das Blatt mit dem Namen "Mission_Control"
    var sheet = spreadsheet.getSheetByName("Mission_Control");
    
    // Löscht den aktuellen Inhalt der angegebenen Zelle
    sheet.getRange(zellenAdresse).clearContent();
    
    // Setzt den neuen Text in die angegebene Zelle
    sheet.getRange(zellenAdresse).setValue(text);
    
    // Erzwingt das sofortige Anwenden aller ausstehenden Änderungen
    SpreadsheetApp.flush();
  }
  
  
  // Funktion zum Löschen des Inhalts einer spezifischen Zelle im "Mission_Control" Blatt
  /**
   * Löscht den Inhalt einer bestimmten Zelle im "Mission_Control" Blatt.
   *
   * @param {string} zellenAdresse - Die Adresse der Zielzelle (z.B. "A1").
   */
  function loescheZellInhaltInMissionControl(zellenAdresse) {
    // Zugriff auf das aktuell geöffnete Spreadsheet
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Zugriff auf das Blatt mit dem Namen "Mission_Control"
    var sheet = spreadsheet.getSheetByName("Mission_Control");
    
    // Löscht den Inhalt der angegebenen Zelle
    sheet.getRange(zellenAdresse).clearContent();
    
    // Erzwingt das sofortige Anwenden aller ausstehenden Änderungen
    SpreadsheetApp.flush();
  }
  
  
  /**
   * Lädt die Konfigurationswerte aus dem "Config" Blatt.
   *
   * Diese Funktion liest alle Schlüssel-Wert-Paare aus dem "Config" Blatt und erstellt daraus ein Konfigurationsobjekt.
   *
   * @return {Object} Ein Objekt mit den Konfigurationsparametern.
   * @throws {Error} Wenn das "Config" Blatt nicht gefunden wird.
   */
  function getConfig() {
    // Zugriff auf das "Config" Blatt im aktuellen Spreadsheet
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
    
    // Überprüft, ob das "Config" Blatt existiert
    if (!sheet) {
      throw new Error('Konfigurationsblatt "Config" nicht gefunden.');
    }
    
    // Liest alle Daten im "Config" Blatt
    const data = sheet.getDataRange().getValues();
    const config = {};
    
    // Durchläuft alle Zeilen des Blatts, beginnend bei der zweiten Zeile (Index 1), um die Header zu überspringen
    for (let i = 1; i < data.length; i++) {
      const key = data[i][0];   // Der Schlüssel befindet sich in der ersten Spalte
      const value = data[i][1]; // Der Wert befindet sich in der zweiten Spalte
      
      // Fügt das Schlüssel-Wert-Paar zum Konfigurationsobjekt hinzu, falls der Schlüssel nicht leer ist
      if (key) {
        config[key] = value;
      }
    }
    
    return config;
  }
  
  
  /**
   * Testet die getConfig-Funktion und gibt alle Konfigurationswerte in den Logs aus.
   *
   * Diese Funktion dient zur Überprüfung, ob die Konfiguration korrekt geladen wird.
   */
  function testGetConfig() {
    try {
      // Ruft die Konfigurationswerte ab
      const config = getConfig();
      
      // Konvertiert bestimmte Parameter in die benötigten Typen
      config.IS_TEST_MODE = config.TestMode.toString().toLowerCase() === 'true';
      config.TEST_EMAIL = config.TestEmail;
      
      // Loggt alle Konfigurationswerte
      Logger.log('Konfigurationswerte:');
      for (const key in config) {
        if (config.hasOwnProperty(key)) {
          Logger.log(`${key}: ${config[key]}`);
        }
      }
    } catch (error) {
      // Loggt einen Fehler, falls das Laden der Konfiguration fehlschlägt
      Logger.log(`Fehler beim Laden der Konfiguration: ${error.message}`);
    }
  }
  
  
  // Funktion zum Versenden einer E-Mail mit Anhang
  /**
   * Sendet eine E-Mail mit einem Anhang an einen Empfänger.
   *
   * @param {string} recipient - Die E-Mail-Adresse des Empfängers.
   * @param {string} subject - Der Betreff der E-Mail.
   * @param {string} bodyHtml - Der HTML-Inhalt der E-Mail.
   * @param {Blob} attachment - Der Anhang der E-Mail als Blob.
   */
  function sendEmailWithAttachment(recipient, subject, bodyHtml, attachment) {
    MailApp.sendEmail({
      to: recipient,
      subject: subject,
      htmlBody: bodyHtml,
      attachments: [attachment]
    });
  }
  
  
  /**
   * Lädt das E-Mail-Template aus einer JSON-Datei in Google Drive.
   *
   * Diese Funktion liest eine JSON-Datei, die verschiedene E-Mail-Templates enthält, und gibt das gewünschte Template zurück.
   *
   * @param {string} templateName - Der Name des Templates (z.B. "invoice" oder "sepa").
   * @param {Object} CONFIG - Das Konfigurationsobjekt, das die ID der JSON-Datei enthält.
   * @return {Object} Das Template-Objekt mit "subject" und "bodyHtml".
   * @throws {Error} Wenn das angeforderte Template nicht gefunden wird.
   */
  function loadEmailTemplate(templateName, CONFIG) {
    // Ruft die ID der JSON-Datei mit den E-Mail-Templates aus der Konfiguration ab
    const templateFileId = CONFIG.EmailTemplateJsonFileId;
    
    // Zugriff auf die Datei in Google Drive anhand der ID
    const file = DriveApp.getFileById(templateFileId);
    
    // Liest den Inhalt der Datei als String
    const jsonContent = file.getBlob().getDataAsString();
    
    // Parst den JSON-Inhalt in ein JavaScript-Objekt
    const templates = JSON.parse(jsonContent);
    
    // Loggt alle verfügbaren Template-Namen zur Überprüfung
    console.log('Vorhandene Templates:', Object.keys(templates));
    
    // Überprüft, ob das gewünschte Template existiert
    if (templates[templateName]) {
      return templates[templateName];
    } else {
      // Schreibt eine Fehlermeldung in das "Mission_Control" Blatt und wirft einen Fehler
      schreibeTextInZelleInMissionControl('D8', `Template "${templateName}" nicht gefunden.`);
      throw new Error(`Template "${templateName}" nicht gefunden.`);
    }
  }
  
  
  // Funktion zum Versenden von Rechnungs-PDFs
  /**
   * Sendet alle Rechnungs-PDFs an die entsprechenden Empfänger.
   *
   * Diese Funktion durchläuft die Daten im konfigurierten Datenblatt, prüft, welche Mitglieder per Rechnung bezahlt werden,
   * lädt die entsprechenden PDF-Dateien und sendet sie per E-Mail an die Empfänger.
   * Dabei wird auch zwischen Test- und Produktionsmodus unterschieden.
   */
  function sendInvoicePDFsToRecipients() {
    // Schreibt den Startstatus in das "Mission_Control" Blatt
    schreibeTextInZelleInMissionControl('K6', 'Gestartet');
  
    // Lädt die Konfigurationswerte
    let CONFIG = getConfig();
  
    // Konvertiert bestimmte Konfigurationsparameter in die benötigten Typen
    CONFIG.IS_TEST_MODE = CONFIG.TestMode.toString().toLowerCase() === 'true';
    CONFIG.TEST_EMAIL = CONFIG.TestEmail;
    
    // Zugriff auf das Datenblatt, das in der Konfiguration festgelegt ist
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.DataSheetName);
    
    // Liest alle Daten aus dem Datenblatt
    const data = sheet.getDataRange().getValues();
    
    // Zugriff auf den Ordner, in dem die Rechnungs-PDFs gespeichert sind
    const invoiceFolder = DriveApp.getFolderById(CONFIG.InvoiceFolderId);
    
    // Speichert den Testmodus und die Test-E-Mail in separate Variablen für bessere Lesbarkeit
    const IS_TEST_MODE = CONFIG.IS_TEST_MODE;
    const TEST_EMAIL = CONFIG.TEST_EMAIL;
    
    // Initialisiert einen Zähler für die versendeten PDFs
    let counter = 0;
    
    // Aktualisiert den Status im "Mission_Control" Blatt
    schreibeTextInZelleInMissionControl('K6', 'Konfiguration eingelesen');
    
    // Loggt den aktuellen Testmodus und die Test-E-Mail
    Logger.log(`Testmode ist:  ${IS_TEST_MODE} und die gesetzte TestEmail ist: ${TEST_EMAIL}`);
    schreibeTextInZelleInMissionControl('K7', `Testmode ist:  ${IS_TEST_MODE} und die gesetzte TestEmail ist: ${TEST_EMAIL}`);
  
    // Lädt das E-Mail-Template für Rechnungen
    const invoiceTemplate = loadEmailTemplate('invoice', CONFIG);
  
    // Durchläuft alle Datenzeilen im Datenblatt, beginnend bei der zweiten Zeile (Index 1)
    for (let i = 1; i < data.length; i++) {
      // Prüft, ob das Mitglied per Rechnung bezahlt (angenommen, Spalte 14 hat den Index 13)
      if (data[i][13] === 1) {
        const privateMail = data[i][0]; // E-Mail-Adresse des Mitglieds
        const vorname = data[i][3];      // Vorname des Mitglieds
        const nachname = data[i][4];     // Nachname des Mitglieds
        
        // Generiert den erwarteten PDF-Dateinamen basierend auf Vor- und Nachname
        const pdfName = `Rechnung_${vorname}_${nachname}.pdf`;
        
        // Loggt die Suche nach der Datei
        Logger.log(`es wird gesucht nach Dateiname: ${pdfName} in diesem Ordner: ${invoiceFolder}`);
        schreibeTextInZelleInMissionControl('K6', `Gefundenes Paar: Email: ${privateMail} Filename: ${pdfName}`);
  
        // Sucht die PDF-Datei im Rechnungsordner
        const pdfFile = getFileByName(invoiceFolder, pdfName);
  
        if (pdfFile) {
          // Loggt die erfolgreiche Suche der Datei
          Logger.log(`Gefundenes Paar: Email: ${privateMail} Filename: ${pdfName}`);
          
          // Personalisieren des E-Mail-Body-Texts durch Ersetzen von Platzhaltern
          const bodyPersonalized = invoiceTemplate.bodyHtml.replace('{{VORNAME}}', vorname);
  
          // Bestimmt die Empfängeradresse basierend auf dem Testmodus
          let recipient;
          if (IS_TEST_MODE) {
            recipient = TEST_EMAIL; // Im Testmodus wird die Test-E-Mail verwendet
          } else {
            recipient = privateMail; // Andernfalls wird die private E-Mail des Mitglieds verwendet
          }
          
          // Versendet die E-Mail mit dem Anhang (auskommentiert, um den Versand zu steuern)
          sendEmailWithAttachment(recipient, invoiceTemplate.subject, bodyPersonalized, pdfFile);
          
          // Erhöht den Zähler für versendete PDFs
          counter++;
          
          // Loggt den erfolgreichen Versand der PDF
          Logger.log(`Rechnungs-PDF versendet an: ${recipient}. Anzahl: ${counter}`);
          
          // Aktualisiert den Status im "Mission_Control" Blatt
          schreibeTextInZelleInMissionControl('L6', `Rechnungs-PDF versendet an: ${recipient}`);
          schreibeTextInZelleInMissionControl('M6', `Anzahl: ${counter}`);
        } else {
          // Loggt, dass die PDF-Datei nicht gefunden wurde
          Logger.log(`Rechnungs-PDF nicht gefunden für: ${vorname} ${nachname}`);
          
          // Aktualisiert den Status im "Mission_Control" Blatt mit einer Fehlermeldung
          schreibeTextInZelleInMissionControl('K8', `SEPA-PDF nicht gefunden für: ${vorname} ${nachname}`);
        }
      }
    }
    
    // Löscht temporäre Statusmeldungen aus dem "Mission_Control" Blatt
    loescheZellInhaltInMissionControl('L6');
    loescheZellInhaltInMissionControl('M6');
    loescheZellInhaltInMissionControl('K7');
    
    // Schreibt die Abschlussmeldung mit der Gesamtanzahl versendeter PDFs
    schreibeTextInZelleInMissionControl('K6', `Fertig. Insgesammt ${counter} PDFs versendet`);
  }
  
  
  // Funktion zum Versenden von SEPA-PDFs
  /**
   * Sendet alle SEPA-PDFs an die entsprechenden Empfänger.
   *
   * Diese Funktion durchläuft die Daten im konfigurierten Datenblatt, prüft, welche Mitglieder per SEPA bezahlt werden,
   * lädt die entsprechenden PDF-Dateien und sendet sie per E-Mail an die Empfänger.
   * Dabei wird auch zwischen Test- und Produktionsmodus unterschieden.
   */
  function sendSepaPDFsToRecipients() {
    // Schreibt den Startstatus in das "Mission_Control" Blatt
    schreibeTextInZelleInMissionControl('K13', 'Gestartet');
  
    // Lädt die Konfigurationswerte
    let CONFIG = getConfig();
  
    // Konvertiert bestimmte Konfigurationsparameter in die benötigten Typen
    CONFIG.IS_TEST_MODE = CONFIG.TestMode.toString().toLowerCase() === 'true';
    CONFIG.TEST_EMAIL = CONFIG.TestEmail;
    
    // Zugriff auf das Datenblatt, das in der Konfiguration festgelegt ist
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.DataSheetName);
    
    // Liest alle Daten aus dem Datenblatt
    const data = sheet.getDataRange().getValues();
    
    // Zugriff auf den Ordner, in dem die SEPA-PDFs gespeichert sind
    const sepaFolder = DriveApp.getFolderById(CONFIG.SepaFolderId);
    
    // Speichert den Testmodus und die Test-E-Mail in separate Variablen für bessere Lesbarkeit
    const IS_TEST_MODE = CONFIG.IS_TEST_MODE;
    const TEST_EMAIL = CONFIG.TEST_EMAIL;
    
    // Initialisiert einen Zähler für die versendeten PDFs
    let counter = 0;
    
    // Aktualisiert den Status im "Mission_Control" Blatt
    schreibeTextInZelleInMissionControl('K13', 'Konfiguration eingelesen');
  
    // Loggt den aktuellen Testmodus und die Test-E-Mail
    Logger.log(`Testmode ist:  ${IS_TEST_MODE} und die gesetzte TestEmail ist: ${TEST_EMAIL}`);
    schreibeTextInZelleInMissionControl('K14', `Testmode ist:  ${IS_TEST_MODE} und die gesetzte TestEmail ist: ${TEST_EMAIL}`);
  
    // Lädt das E-Mail-Template für SEPA
    const sepaTemplate = loadEmailTemplate('sepa', CONFIG);
  
    // Durchläuft alle Datenzeilen im Datenblatt, beginnend bei der zweiten Zeile (Index 1)
    for (let i = 1; i < data.length; i++) {
      // Prüft, ob das Mitglied per SEPA bezahlt (angenommen, Spalte 14 hat den Index 13)
      if (data[i][13] === 0) {
        const privateMail = data[i][0]; // E-Mail-Adresse des Mitglieds
        const vorname = data[i][3];      // Vorname des Mitglieds
        const nachname = data[i][4];     // Nachname des Mitglieds
        
        // Generiert den erwarteten PDF-Dateinamen basierend auf Vor- und Nachname
        const pdfName = `Einzug_${vorname}_${nachname}.pdf`;
        
        // Loggt die Suche nach der Datei
        Logger.log(`es wird gesucht nach Dateiname: ${pdfName} in diesem Ordner: ${sepaFolder}`);
  
        // Sucht die PDF-Datei im SEPA-Ordner
        const pdfFile = getFileByName(sepaFolder, pdfName);
  
        if (pdfFile) {
          // Loggt die erfolgreiche Suche der Datei
          Logger.log(`Gefundenes Paar: Email: ${privateMail} Filename: ${pdfName}`);
          schreibeTextInZelleInMissionControl('K13', `Gefundenes Paar: Email: ${privateMail} Filename: ${pdfName}`);
  
          // Personalisieren des E-Mail-Body-Texts durch Ersetzen von Platzhaltern
          const bodyPersonalized = sepaTemplate.bodyHtml.replace('{{VORNAME}}', vorname);
  
          // Bestimmt die Empfängeradresse basierend auf dem Testmodus
          let recipient;
          if (IS_TEST_MODE) {
            recipient = TEST_EMAIL; // Im Testmodus wird die Test-E-Mail verwendet
          } else {
            recipient = privateMail; // Andernfalls wird die private E-Mail des Mitglieds verwendet
          }
          
          // Versendet die E-Mail mit dem Anhang (auskommentiert, um den Versand zu steuern)
          sendEmailWithAttachment(recipient, sepaTemplate.subject, bodyPersonalized, pdfFile);
          
          // Erhöht den Zähler für versendete PDFs
          counter++;
          
          // Loggt den erfolgreichen Versand der PDF
          Logger.log(`Rechnungs-PDF versendet an: ${recipient}. Anzahl: ${counter}`);
          
          // Aktualisiert den Status im "Mission_Control" Blatt
          schreibeTextInZelleInMissionControl('L13', `Rechnungs-PDF versendet an: ${recipient}`);
          schreibeTextInZelleInMissionControl('M13', `Anzahl: ${counter}`);
        } else {
          // Loggt, dass die PDF-Datei nicht gefunden wurde
          Logger.log(`SEPA-PDF nicht gefunden für: ${vorname} ${nachname}`);
          
          // Aktualisiert den Status im "Mission_Control" Blatt mit einer Fehlermeldung
          schreibeTextInZelleInMissionControl('K15', `SEPA-PDF nicht gefunden für: ${vorname} ${nachname}`);
        }
      }
    }
    
    // Löscht temporäre Statusmeldungen aus dem "Mission_Control" Blatt
    loescheZellInhaltInMissionControl('L13');
    loescheZellInhaltInMissionControl('M13');
    loescheZellInhaltInMissionControl('K14');
    
    // Schreibt die Abschlussmeldung mit der Gesamtanzahl versendeter PDFs
    schreibeTextInZelleInMissionControl('K13', `Fertig. Insgesammt ${counter} PDFs versendet`);
  }
  
  
  // Funktion zur Erstellung von Rechnungs-PDFs
  /**
   * Erstellt Rechnungs-PDFs für alle Mitglieder, die per Rechnung bezahlen.
   *
   * Diese Funktion durchläuft die Daten im konfigurierten Datenblatt, erstellt für jedes berechtigte Mitglied eine PDF basierend
   * auf einer Google Docs Vorlage und speichert diese im angegebenen Ordner.
   */
  function createInvoicePDFs() {
    // Schreibt den Startstatus in das "Mission_Control" Blatt
    schreibeTextInZelleInMissionControl('D7', 'Gestartet');
    
    // Lädt die Konfigurationswerte
    let CONFIG = getConfig();
    
    // Zugriff auf das Datenblatt, das in der Konfiguration festgelegt ist
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.DataSheetName);
    
    // Liest alle Daten aus dem Datenblatt
    const data = sheet.getDataRange().getValues();
    
    // Ruft die ID der Google Docs Vorlage für Rechnungen aus der Konfiguration ab
    const templateId = CONFIG.InvoiceTemplateId;
    
    // Zugriff auf den Zielordner für die erstellten Rechnungs-PDFs
    const targetFolder = DriveApp.getFolderById(CONFIG.InvoiceFolderId);
    
    // Formatiert das aktuelle Datum im gewünschten Format
    const todaysDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');
    
    // Initialisiert einen Zähler für die erstellten PDFs
    let counter = 0;
  
    // Aktualisiert den Status im "Mission_Control" Blatt
    schreibeTextInZelleInMissionControl('D7', 'Konfigurations eingelesen');
  
    // Durchläuft alle Datenzeilen im Datenblatt, beginnend bei der zweiten Zeile (Index 1)
    for (let i = 1; i < data.length; i++) {
      // Prüft, ob das Mitglied per Rechnung bezahlt (angenommen, Spalte 14 hat den Index 13)
      if (data[i][13] === 1) {
        // Aktualisiert den Status im "Mission_Control" Blatt
        schreibeTextInZelleInMissionControl('D7', 'Bisher erstellt:');
  
        // Extrahiert relevante Informationen aus der aktuellen Datenzeile
        const vorname = data[i][3];
        const nachname = data[i][4];
        const strasse = data[i][9];
        const plz = data[i][10];
        const stadt = data[i][11];
        const mandatsreferenz = data[i][8];
  
        // Kopiert die Google Docs Vorlage und benennt die Kopie entsprechend um
        const tempFile = DriveApp.getFileById(templateId).makeCopy(`Rechnung_${vorname}_${nachname}`, targetFolder);
        
        // Öffnet das kopierte Dokument zur Bearbeitung
        const tempDoc = DocumentApp.openById(tempFile.getId());
        let body = tempDoc.getBody();
  
        // Ersetzt Platzhalter im Dokument mit den tatsächlichen Daten des Mitglieds
        body.replaceText('{{VORNAME}}', vorname);
        body.replaceText('{{NACHNAME}}', nachname);
        body.replaceText('{{STR}}', strasse);
        body.replaceText('{{PLZ}}', plz);
        body.replaceText('{{ORT}}', stadt);
        body.replaceText('{{TODAYS_DATE}}', todaysDate);
        body.replaceText('{{MANDATSREFERENZ}}', mandatsreferenz);
  
        // Speichert und schließt das bearbeitete Dokument
        tempDoc.saveAndClose();
  
        // Konvertiert das bearbeitete Dokument in ein PDF
        const pdf = tempFile.getAs('application/pdf');
        
        // Speichert das PDF im Zielordner mit dem entsprechenden Namen
        targetFolder.createFile(pdf).setName(`Rechnung_${vorname}_${nachname}.pdf`);
  
        // Erhöht den Zähler für erstellte PDFs
        counter++;
        
        // Loggt die erfolgreiche Erstellung der PDF
        Logger.log(`Rechnungs-PDF erstellt für: ${vorname} ${nachname}. Anzahl: ${counter}`);
        
        // Aktualisiert den Status im "Mission_Control" Blatt mit der Anzahl und der neuesten Erstellung
        schreibeTextInZelleInMissionControl('E7', `${counter}`);
        schreibeTextInZelleInMissionControl('F7', `Neueste: für ${vorname}_${nachname}`);
      }
    }
    
    // Löscht temporäre Statusmeldungen aus dem "Mission_Control" Blatt
    loescheZellInhaltInMissionControl('E7');
    loescheZellInhaltInMissionControl('F7');
    
    // Schreibt die Abschlussmeldung mit der Gesamtanzahl erstellter PDFs
    schreibeTextInZelleInMissionControl('D7', `Fertig. Insgesammt ${counter} PDFs erstellt`);
  }
  
  
  // Funktion zur Erstellung von SEPA-PDFs
  /**
   * Erstellt SEPA-PDFs für alle Mitglieder, die per SEPA bezahlen.
   *
   * Diese Funktion durchläuft die Daten im konfigurierten Datenblatt, erstellt für jedes berechtigte Mitglied eine PDF basierend
   * auf einer Google Docs Vorlage und speichert diese im angegebenen Ordner.
   */
  function createSepaPDFs() {
    // Schreibt den Startstatus in das "Mission_Control" Blatt
    schreibeTextInZelleInMissionControl('D13', 'Gestartet');
    
    // Lädt die Konfigurationswerte
    let CONFIG = getConfig();
    
    // Zugriff auf das Datenblatt, das in der Konfiguration festgelegt ist
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.DataSheetName);
    
    // Liest alle Daten aus dem Datenblatt
    const data = sheet.getDataRange().getValues();
    
    // Ruft die ID der Google Docs Vorlage für SEPA aus der Konfiguration ab
    const templateId = CONFIG.SepaTemplateId;
    
    // Zugriff auf den Zielordner für die erstellten SEPA-PDFs
    const targetFolder = DriveApp.getFolderById(CONFIG.SepaFolderId);
    
    // Formatiert das aktuelle Datum im gewünschten Format
    const todaysDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');
    
    // Initialisiert einen Zähler für die erstellten PDFs
    let counter = 0;
  
    // Aktualisiert den Status im "Mission_Control" Blatt
    schreibeTextInZelleInMissionControl('D13', 'Konfigurations eingelesen');
    
    // Durchläuft alle Datenzeilen im Datenblatt, beginnend bei der zweiten Zeile (Index 1)
    for (let i = 1; i < data.length; i++) {
      // Prüft, ob das Mitglied per SEPA bezahlt (angenommen, Spalte 14 hat den Index 13)
      if (data[i][13] === 0) {
        // Aktualisiert den Status im "Mission_Control" Blatt
        schreibeTextInZelleInMissionControl('D13', 'Bisher erstellt:');
        
        // Extrahiert relevante Informationen aus der aktuellen Datenzeile
        const vorname = data[i][3];
        const nachname = data[i][4];
        const iban = data[i][5];
        const ibanLastTwo = iban.slice(-2); // Extrahiert die letzten zwei Ziffern der IBAN
        const bank = data[i][6];
        const kontohalter = data[i][7];
        const lastschriftKuerzel = data[i][8];
        const strasse = data[i][9];
        const plz = data[i][10];
        const stadt = data[i][11];
        const mandatsreferenz = data[i][8]; // Annahme: Mandatsreferenz befindet sich ebenfalls in Spalte 9
  
        // Kopiert die Google Docs Vorlage und benennt die Kopie entsprechend um
        const tempFile = DriveApp.getFileById(templateId).makeCopy(`Einzug_${vorname}_${nachname}`, targetFolder);
        
        // Öffnet das kopierte Dokument zur Bearbeitung
        const tempDoc = DocumentApp.openById(tempFile.getId());
        let body = tempDoc.getBody();
  
        // Ersetzt Platzhalter im Dokument mit den tatsächlichen Daten des Mitglieds
        body.replaceText('{{VORNAME}}', vorname);
        body.replaceText('{{NACHNAME}}', nachname);
        body.replaceText('{{Letzte_Zwei_Ziffern_Von_Iban}}', ibanLastTwo);
        body.replaceText('{{Bank}}', bank);
        body.replaceText('{{Kontohalter}}', kontohalter);
        body.replaceText('{{Lastschrift_Kuerzel}}', lastschriftKuerzel);
        body.replaceText('{{STR}}', strasse);
        body.replaceText('{{PLZ}}', plz);
        body.replaceText('{{ORT}}', stadt);
        body.replaceText('{{TODAYS_DATE}}', todaysDate);
        body.replaceText('{{MANDATSREFERENZ}}', mandatsreferenz);
  
        // Speichert und schließt das bearbeitete Dokument
        tempDoc.saveAndClose();
  
        // Konvertiert das bearbeitete Dokument in ein PDF
        const pdf = tempFile.getAs('application/pdf');
        
        // Speichert das PDF im Zielordner mit dem entsprechenden Namen
        targetFolder.createFile(pdf).setName(`Einzug_${vorname}_${nachname}.pdf`);
  
        // Erhöht den Zähler für erstellte PDFs
        counter++;
        
        // Loggt die erfolgreiche Erstellung der PDF
        Logger.log(`SEPA-PDF erstellt für: ${vorname} ${nachname}. Anzahl: ${counter}`);
        
        // Aktualisiert den Status im "Mission_Control" Blatt mit der Anzahl und der neuesten Erstellung
        schreibeTextInZelleInMissionControl('E13', `${counter}`);
        schreibeTextInZelleInMissionControl('F13', `Neueste: für ${vorname}_${nachname}`);
      }
    }
    
    // Löscht temporäre Statusmeldungen aus dem "Mission_Control" Blatt
    loescheZellInhaltInMissionControl('E13');
    loescheZellInhaltInMissionControl('F13');
    
    // Schreibt die Abschlussmeldung mit der Gesamtanzahl erstellter PDFs
    schreibeTextInZelleInMissionControl('D13', `Fertig. Insgesammt ${counter} PDFs erstellt`);
  }