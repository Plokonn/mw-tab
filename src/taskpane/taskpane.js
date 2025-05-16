/*
 * Excel Add-In für Musicworks PowerPoint-Generator
 */

// Office-Initialisierung beim Laden des Add-Ins
Office.initialize = function (reason) {
    // Wird aufgerufen, wenn das Add-In geladen wird
    $(document).ready(function () {
        // Event-Handler für den Server-Button
        $('#serverButton').click(sendDataToServer);
        
        // Statusmeldung anzeigen
        showStatus("Add-In geladen und bereit.");
    });
};

/**
 * Liest Daten aus Excel und sendet sie an den Server
 */
function sendDataToServer() {
    showStatus("Lese Daten aus Excel...", "normal");
    
    // Excel.run verwenden, um auf das Workbook zuzugreifen
    Excel.run(function (context) {
        // Das aktive Arbeitsblatt finden
        var activeSheet = context.workbook.worksheets.getActiveWorksheet();
        activeSheet.load("name");
        
        // Ausführen und Namen des aktiven Arbeitsblatts bekommen
        return context.sync()
            .then(function () {
                // Den Namen des aktiven Arbeitsblatts bekommen
                var activeSheetName = activeSheet.name;
                console.log("Aktives Arbeitsblatt:", activeSheetName);
                
                // Format aus dem aktiven Arbeitsblatt bestimmen
                var wsFormat = activeSheetName.toLowerCase();
                
                // Eine Liste aller Funktionsbezeichnungen
                const functionNames = [
                    "pax", "time", "mainroom", "sideroom1", "sideroom2", "sideroom3", 
                    "sideroom4", "sideroom5", "sideroom6", "sideroom7", "sideroom8", 
                    "chairs", "setup", "setdown", "manager", "pa", 
                    "extra1", "extra2", "extra3", "extra4", "extra5", 
                    "extraclick1", "extraclick2", "extraclick3", "extraclick4", "extraclick5"
                ];
                
                // Dropdown-Menü-Namen (vorher Checkboxen)
                const dropdownNames = ["pa", "extraclick1", "extraclick2", "extraclick3", "extraclick4", "extraclick5"];
                
                // Funktionen zum Laden der benannten Bereiche und Dropdowns
                var namedItems = context.workbook.names;
                namedItems.load("items/name, items/type");
                
                // Daten für die API vorbereiten
                var data = {
                    ws_format: wsFormat,  // Format aus dem aktiven Arbeitsblatt
                    format: "both"        // Standardwert für Export-Format
                };
                
                // Vorinitialisieren aller Funktionsnamen mit null
                functionNames.forEach(function(fnName) {
                    data[fnName] = null;
                });
                
                return context.sync().then(function() {
                    // Alle erkannten Namen und ihre Typen anzeigen
                    console.log("Gefundene benannte Bereiche:", namedItems.items.map(item => ({
                        name: item.name,
                        type: item.type
                    })));
                    
                    // Zum Sammeln aller Bereiche, die wir laden wollen
                    var namesToProcess = [];
                    
                    // Für jede Funktionsbezeichnung prüfen
                    for (var fnIndex = 0; fnIndex < functionNames.length; fnIndex++) {
                        var fnName = functionNames[fnIndex];
                        var fullName = wsFormat + "_" + fnName;
                        
                        // Prüfen, ob es einen benannten Bereich mit diesem Namen gibt
                        for (var niIndex = 0; niIndex < namedItems.items.length; niIndex++) {
                            if (namedItems.items[niIndex].name.toLowerCase() === fullName) {
                                namesToProcess.push({
                                    name: namedItems.items[niIndex].name,
                                    functionName: fnName,
                                    isDropdown: dropdownNames.includes(fnName)
                                });
                                break;
                            }
                        }
                    }
                    
                    console.log("Zu verarbeitende benannte Bereiche:", namesToProcess);
                    
                    // Jetzt verarbeiten wir die gefundenen Namen nacheinander
                    return processNamedItems(context, namesToProcess, data, 0);
                });
            })
            .then(function(processedData) {
                console.log("Gesammelte Daten:", processedData);
                
                // Status aktualisieren
                showStatus("Daten erfolgreich ausgelesen. Bereite Präsentation vor...", "normal");
                
                // API-URL
                const apiUrl = 'https://musicworksnas.synology.me:5049/api/concept';
                
                // Explizite String-Konvertierung für numerische Werte
                var requestData = {
                    ws_format: processedData.ws_format,
                    format: processedData.format
                };
                
                // Für alle anderen Felder: Dropdown-Werte als Boolean, alles andere als String
                const dropdownNames = ["pa", "extraclick1", "extraclick2", "extraclick3", "extraclick4", "extraclick5"];
                
                Object.keys(processedData).forEach(function(key) {
                    // format und ws_format wurden bereits behandelt
                    if (key !== 'format' && key !== 'ws_format') {
                        if (processedData[key] !== null && processedData[key] !== undefined) {
                            if (dropdownNames.includes(key)) {
                                // Dropdown-Werte als Boolean behalten
                                requestData[key] = processedData[key];
                            } else {
                                // Alle anderen Werte explizit zu Strings konvertieren
                                requestData[key] = String(processedData[key]);
                            }
                        } else {
                            requestData[key] = null;
                        }
                    }
                });
                
                console.log("Daten, die an die API gesendet werden:", requestData);
                
                // Daten an den Server senden
                return fetch(apiUrl, {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify(requestData)
                })
                .then(response => {
                    if (!response.ok) {
                        throw new Error('Netzwerkfehler: ' + response.status);
                    }
                    
                    // Prüfen, ob die Antwort JSON oder eine Datei ist
                    const contentType = response.headers.get('content-type');
                    if (contentType && contentType.includes('application/json')) {
                        return response.json().then(data => {
                            return { type: 'json', data: data };
                        });
                    } else {
                        return response.blob().then(blob => {
                            return { type: 'blob', data: blob };
                        });
                    }
                })
                .then(result => {
                    if (result.type === 'json') {
                        // JSON-Antwort mit Links zu beiden Dateien
                        const data = result.data;
                        
                        // Download-Links für beide Dateien erstellen
                        const serverUrl = new URL(apiUrl).origin; // Basis-URL des Servers
                        const pptxUrl = serverUrl + data.pptx_url;
                        const pdfUrl = serverUrl + data.pdf_url;
                        
                        console.log("Download-URLs:", { pptx: pptxUrl, pdf: pdfUrl });
                        
                        // Download-Links im UI anzeigen
                        createDownloadLinks(pptxUrl, pdfUrl);
                        
                        showStatus("Dateien bereit zum Download.", "success");
                    } else {
                        // Direkter Blob-Download (einzelne Datei)
                        const blob = result.data;
                        const filename = getFilenameFromContentType(
                            blob.type, 
                            response.headers.get('content-disposition')
                        );
                        
                        // Download-Link erstellen
                        const url = URL.createObjectURL(blob);
                        createSingleDownloadLink(url, filename);
                        
                        showStatus("Datei bereit zum Download.", "success");
                    }
                });
            })
            .catch(error => {
                showStatus("Fehler bei der Verarbeitung: " + error.message, "error");
                console.error('Fehler:', error);
            });
    });
}

/**
 * Verarbeitet benannte Bereiche nacheinander, um Sync-Fehler zu vermeiden
 */
function processNamedItems(context, itemList, data, index) {
    // Wenn alle Items verarbeitet wurden, geben wir die Daten zurück
    if (index >= itemList.length) {
        return data;
    }
    
    var item = itemList[index];
    var namedRange;
    
    try {
        namedRange = context.workbook.names.getItem(item.name);
        var range = namedRange.getRange();
        range.load("values");
        
        return context.sync()
            .then(function() {
                try {
                    var value = range.values[0][0];
                    if (value !== null && value !== undefined) {
                        if (item.isDropdown) {
                            // Dropdown-Werte als Boolean interpretieren basierend auf dem ausgewählten Text
                            // Wir nehmen an, dass "Ja", "Yes", "Wahr", "True" oder "1" als true gelten
                            if (typeof value === 'string') {
                                value = value.toLowerCase() === "ja" || 
                                       value.toLowerCase() === "yes" || 
                                       value.toLowerCase() === "wahr" ||
                                       value.toLowerCase() === "true" || 
                                       value === "1";
                            } else {
                                value = Boolean(value);
                            }
                        }
                        
                        // Wert direkt unter dem Funktionsnamen speichern
                        data[item.functionName] = value;
                        console.log(`Wert für ${item.functionName} gefunden:`, value);
                    }
                } catch (valueError) {
                    console.error(`Fehler beim Lesen des Werts für ${item.name}:`, valueError);
                }
                
                // Nächstes Item verarbeiten
                return processNamedItems(context, itemList, data, index + 1);
            });
    } catch (rangeError) {
        console.error(`Fehler beim Laden des Bereichs für ${item.name}:`, rangeError);
        
        // Trotz Fehler zum nächsten Item weitergehen
        return processNamedItems(context, itemList, data, index + 1);
    }
}

/**
 * Liefert Testdaten für die API-Anfrage
 */
function getTestData() {
    return {
        ws_format: "dgbdw",        // Workshop-Format
        pax: "80",                  // Teilnehmerzahl
        time: "4 Stunden",          // Zeitrahmen
        mainroom: "150",            // Hauptraum (m²)
        sideroom1: "45",            // Seitenraum 1 (m²)
        sideroom2: "45",            // Seitenraum 2 (m²)
        sideroom3: "20",            // Seitenraum 3 (m²)
        chairs: "90",               // Stühle
        setup: "90",                // Aufbauzeit (min)
        setdown: "30",              // Abbauzeit (min)
        manager: "JR",              // Manager (JR, JS, RG, IS, Allgemein)
        pa: true,                   // PA benötigt
        extra1: "Testdaten",        // Extra Feld 1
        extra2: "Beispiel",         // Extra Feld 2
        extraclick1: true,          // Extra Checkbox 1
        format: "both"              // Beide Dateiformate anfordern
    };
}

/**
 * Erstellt Download-Links für PPTX und PDF
 */
function createDownloadLinks(pptxUrl, pdfUrl) {
    // Container für Download-Links finden oder erstellen
    var downloadContainer = document.getElementById('downloadContainer');
    
    // Container leeren
    downloadContainer.innerHTML = '';
    
    // PPTX-Link erstellen
    var pptxLink = document.createElement('a');
    pptxLink.href = pptxUrl;
    pptxLink.className = 'download-button pptx-button';
    pptxLink.textContent = 'PowerPoint Download';
    pptxLink.download = 'praesentation.pptx';
    
    // PDF-Link erstellen
    var pdfLink = document.createElement('a');
    pdfLink.href = pdfUrl;
    pdfLink.className = 'download-button pdf-button';
    pdfLink.textContent = 'PDF Download';
    pdfLink.download = 'praesentation.pdf';
    
    // Info-Text
    var infoText = document.createElement('p');
    infoText.className = 'download-info';
    infoText.textContent = 'Klicke mit Rechtsklick auf die Buttons und wähle "Link speichern unter" um die Dateien herunterzuladen.';
    
    // Links zum Container hinzufügen
    downloadContainer.appendChild(pptxLink);
    downloadContainer.appendChild(pdfLink);
    downloadContainer.appendChild(infoText);
    
    // Container sichtbar machen
    downloadContainer.style.display = 'block';
}

/**
 * Erstellt einen einzelnen Download-Link
 */
function createSingleDownloadLink(url, filename) {
    // Container für Download-Links finden oder erstellen
    var downloadContainer = document.getElementById('downloadContainer');
    
    // Container leeren
    downloadContainer.innerHTML = '';
    
    // Download-Link erstellen
    var downloadLink = document.createElement('a');
    downloadLink.href = url;
    downloadLink.className = 'download-button';
    
    // Button-Klasse basierend auf Dateityp setzen
    if (filename.endsWith('.pptx')) {
        downloadLink.className += ' pptx-button';
        downloadLink.textContent = 'PowerPoint-Präsentation herunterladen';
    } else if (filename.endsWith('.pdf')) {
        downloadLink.className += ' pdf-button';
        downloadLink.textContent = 'PDF-Version herunterladen';
    } else {
        downloadLink.textContent = 'Datei herunterladen';
    }
    
    downloadLink.download = filename;
    
    // Info-Text
    var infoText = document.createElement('p');
    infoText.className = 'download-info';
    infoText.textContent = 'Klicken Sie auf den Button, um die Datei herunterzuladen.';
    
    // Link zum Container hinzufügen
    downloadContainer.appendChild(downloadLink);
    downloadContainer.appendChild(infoText);
    
    // Container sichtbar machen
    downloadContainer.style.display = 'block';
}

/**
 * Versucht, den Dateinamen aus dem Content-Type oder Content-Disposition-Header zu extrahieren
 */
function getFilenameFromContentType(contentType, contentDisposition) {
    let filename = 'praesentation';
    
    // Aus Content-Disposition extrahieren, falls vorhanden
    if (contentDisposition) {
        const matches = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/.exec(contentDisposition);
        if (matches && matches[1]) {
            filename = matches[1].replace(/['"]/g, '');
        }
    }
    
    // Erweiterung basierend auf Content-Type hinzufügen
    if (!filename.includes('.')) {
        if (contentType.includes('application/vnd.openxmlformats-officedocument.presentationml.presentation')) {
            filename += '.pptx';
        } else if (contentType.includes('application/pdf')) {
            filename += '.pdf';
        }
    }
    
    return filename;
}

/**
 * Zeigt eine Statusmeldung im UI an
 * @param {string} message - Die anzuzeigende Nachricht
 * @param {string} type - Der Nachrichtentyp (normal, success, error)
 */
function showStatus(message, type) {
    var statusElement = $('#statusMessage');
    
    // Alle Klassen entfernen
    statusElement.removeClass("success error");
    
    // Nachricht setzen
    statusElement.text(message);
    
    // Typ-spezifische Klasse hinzufügen
    if (type === "success") {
        statusElement.addClass("success");
    } else if (type === "error") {
        statusElement.addClass("error");
    }
}