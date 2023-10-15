# Kurlaendische_Gueterurkunden
Dieser Code wurde im Rahmen einer Auftragsarbeit für das Herder Institut für Ost-Mitteleuropaforschung geschrieben.
Nachdem ein ein Urkundenbestand in Riga gesichtet wurde und der Bestand tabellarisch in einer .docx-Datei dokumentiert.
Meine Aufgabe bestand darin, die Tabelle auf das Datenbankschema abzubilden, das der Webseite zu zugrundeliegt. Dabei wird zwischen 3 Fällen unterschieden.

    Fall 1: Die Urkunde ist noch nicht bekannt.
    Fall 2: Die Urkunde ist bereits bekannt und ein in der Datenbank bereits vorhandener Datensatz wird ergänzt.
    Fall 3: Die Urkunde ist bereits bekannt und hat eine bestimmte ID. Ein in der Datenbank bereits vorhandener Datensatz wird ergänzt.

Für jeden der drei Fällen wird eine eigene CSV-Datei erstellt.
Die Abbildung der Tabelle auf das Datenbankschema erschöpft sich nicht in einem einfachen Mapping der Spalten. Einige der Einträge in der .docx-Tabelle enthalten Einträge, die ggf. aufgetrennt und in unterschiedliche Spalten geschrieben werden müssen.
Zuvor habe ich aber noch ein kleine Veranschaulichung der Fallverteilung erstellt und einige Vorarbeiten durchgeführt.
