# Excel-Fahrtenbuch-Helfer
# Anmerkung
Dieses Tool stellt kein rechtlich verwendbares Fahrtenbuch dar, sondern dient nur zum Ermitteln von Strecken und Routenzusammenfassungen.
VBA Skript zum Ausrechnen von Entfernungen und Streckenzusammenfassungen mit Hilfe der GoogleCloud.

Mit Hilfe dieses Tools lassen sich nachträglich Einträge für das Fahrtenbuch erzeugen.

# Voraussetzungen
Excel
Internetverbindung
Google Cloud Konto

Um die Google Cloud benutzen zu dürfen, muss man ein Google Cloud Konto haben und seine Kreditkarteninformationen hinterlegt haben.
Die Abfragen kosten grundsätzlich einen minimalen Betrag, jedoch hat man ein monatliches, freies Guthaben von EUR 200,- (Stand Juni 2022)



# Beschreibung
Das Tool ermittelt über eine API zu Google Maps die Entfernung und eine Streckenzusammenfassung zwischen Start und Ziel.
Die Adressen sollten sich möglichst präzise aus "Strasse Hausnummer, PLZ Ort" zusammensetzen.
Die Verbindung zur API wird über 2 Funktionsaufrufe hergestellt.
Konkret werden hier die Daten über die Kundennummern ermittelt.
Man gibt also eine Kundennummer bei Start und eine Kundennummer bei Ziel ein.
GetDistance ermittelt dann die Entfernung und GetRouteSummary die Streckenzusammenfassun
  
  
# Verwendung
  Um das Tool in die eigene Excel-Tebelle zu integrieren muss die Datei "modul1.bas" importiert werden.
  Klicke hierzu auf den Reiter "Entwicklertools", dann "VisualBasic". Es öffnet sich ein neues Fenster.
  Klicke mit der rechten Maustaste in der linken Spalte auf "Module" und wähle im Menü "Datei importieren".
  Ersetze im Code in beiden Funktionen den Ausdruck <DeinGoogleCloudKey> durch deinen Google-Key.
  
  Der Reiter "Entwicklertools" ist standardmäßig ausgeblendet.
  Über "Datei", "Optionen", "Menüband anpassen" kann man diese einblenden.
  
