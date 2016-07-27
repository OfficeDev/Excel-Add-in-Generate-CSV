# Aufgabenbereich-Add-In-Beispiel „CSV-Generator” für Excel 2016

_Gilt für: Excel 2016_

Dieses Aufgabenbereich-Add-In veranschaulicht das Erstellen einer Tabelle aus einer Liste mit Spaltennamen mithilfe von JavaScript-APIs in Excel 2016. Es ist in zwei Versionen verfügbar: Code-Editor und Visual Studio

![CSV-Generator-Beispiel](../Images/ScreenCap1.PNG)

## Probieren Sie es aus
### Code-Editor-Version

Am einfachsten können Sie Ihr Add-In bereitstellen und testen, indem Sie die Dateien in eine Netzwerkfreigabe kopieren.

1.  Erstellen Sie einen Ordner in einer Netzwerkfreigabe (z.ä B. \\\MyShare\Excel_CSV_Generator), und kopieren Sie alle Dateien im Ordner „Code-Editor”. 
2.  Bearbeiten Sie das <SourceLocation>-Element der Manifestdatei, damit es auf den Freigabepfad aus Schrittä 1 zeigt. 
3.  Kopieren Sie das Manifest (TeacherCSVGenerator.xml) in eine Netzwerkfreigabe (z.ä B. \\\MyShare\MyManifests).
4.  Fügen Sie den Freigabepfad, unter dem das Manifest enthalten ist, als vertrauenswürdigen App-Katalog in Excel hinzu.

    a. Starten Sie Excel, und öffnen Sie ein leeres Arbeitsblatt.  
    
    b. Klicken Sie auf die Registerkarte **Datei**, und klicken Sie dann auf **Optionen**.
    
    c. Wählen Sie **Trust Center** aus, und klicken Sie dann auf die Schaltfläche **Einstellungen für das Trust Center**.
    
  d. Klicken Sie auf **Vertrauenswürdige Add-in-Kataloge**.
    
  e. Geben Sie im Feld **Katalog-URL** den Pfad zu der in Schritt 3 erstellten Netzwerkfreigabe ein, und klicken Sie auf **Katalog hinzufügen**.
    
   f. Aktivieren Sie das Kontrollkästchen **Im Menü anzeigen**, und wählen Sie dann **OK**. Eine Meldung wird angezeigt, dass Ihre Einstellungen angewendet werden, wenn Office das nächste Mal gestartet wird. 
        
5.  Testen und führen Sie das Add-In aus. 

  a. Klicken Sie auf der Registerkarte **Einfügen** in Excel 2016 auf **Meine-Add-Ins**. 
    
  b. Wählen Sie im Dialogfenster **Office-Add-Ins** die Option **Freigegebener Ordner** aus.
    
  c. Wählen Sie **Beispiel für CSV-Kursteilnehmerliste für Lehrer**>**Einfügen**. Das Add-In wird in einem Aufgabenbereich geöffnet und erstellt eine Kursteilnehmerliste im CSV-Format im aktiven Arbeitsblatt, wie im folgenden Screenshot dargestellt. 
      
   ![Studien-Haushaltsplan-Verfolgungsbeispiel](../Images/ScreenCap2.PNG) 

  d. Wählen Sie einen Kursraumverwaltungsdienst.
    
  e. Klicken Sie auf die Schaltfläche zum Erstellen einer Teilnehmerliste, um eine leere Teilnehmerliste im aktiven Arbeitsblatt einzufügen.  
    
      ![College Budget Tracker Sample](../Images/ScreenCap3.PNG) 
      
  f. Klicken Sie auf die Schaltfläche für Hilfe zum Exportieren in Excel, um zu erfahren, wie Sie das Arbeitsblatt als CSV-Datei exportieren.  
  
    
### Visual Studio-Version
1.  Kopieren Sie das Projekt in einen lokalen Ordner, und öffnen Sie die Datei „TeacherCSVGenerator.sln” in Visual Studio.
2.  Drücken Sie F5, um das Beispiel-Add-In zu erstellen und bereitzustellen. Excel wird gestartet und das Add-In wird in einem Aufgabenbereich rechts neben einem leeren Arbeitsblatt geöffnet, wie im folgenden Screenshot dargestellt. 
        
  ![Excel-CSV-Generator-Beispiel](../Images/ScreenCap1.PNG) 

3.  Wählen Sie aus der Dropdown-Liste einen Kursraum-Onlineverwaltungsdienst.
4.  Fügen Sie eine Teilnehmerlistentabelle mithilfe der Schaltfläche **Liste erstellen** hinzu, und sehen Sie sich die erstellte Tabelle im aktiven Arbeitsblatt an.

  ![Studien-Haushaltsplan-Verfolgungsbeispiel](../Images/ScreenCap3.PNG) 
5.  Fügen Sie Teilnehmer zur Liste hinzu, indem Sie die Zellen in den Zeilen unterhalb der Tabellenkopfzeile ausfüllen.
6.  Verwenden Sie die Exportfunktion in Excel, um das Arbeitsblatt als CSV-Datei zu speichern. Diese Datei weist über das richtige Format, um in einem beliebigen Dienst importiert zu werden.


### Weitere Informationen

Die Excel-JavaScript-APIs haben viel mehr bei der Entwicklung von Add-Ins zu bieten. Im Folgenden werden nur einige der verfügbaren Ressourcen aufgeführt. 

1.  [Programmierungsübersicht für Excel-Add-Ins](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-programming-overview.md)
2.  [Codeausschnitt-Explorer für Excel](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)
3.  [Codebeispiele zu Excel-Add-Ins](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-code-samples.md) 
4.  [JavaScript-API-Referenz zu Excel-Add-Ins](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-javascript-reference.md)
5.  [Erstellen Ihres ersten Excel-Add-Ins](https://github.com/OfficeDev/office-js-docs/blob/master/excel/build-your-first-excel-add-in.md)

