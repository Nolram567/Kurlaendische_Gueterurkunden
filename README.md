# Kurlaendische_Gueterurkunden
Dieser Code wurde im Rahmen einer Auftragsarbeit für das Herder-Institut für historische Ostmitteleuropaforschung geschrieben.
Ein Urkundenbestand wurde in Riga gesichtet und tabellarisch in einer .docx-Datei dokumentiert.
Meine Aufgabe bestand darin, die Tabelle auf das Datenbankschema abzubilden, dass der Datenbank [Kurländische Güterurkunden](https://www.herder-institut.de/kurlaendische-gueterurkunden/) zugrunde liegt. Dabei wird zwischen 3 Fällen unterschieden.

- Fall 1: Die Urkunde ist noch nicht bekannt.
- Fall 2: Die Urkunde ist bereits bekannt und ein in der Datenbank bereits vorhandener Datensatz wird ergänzt.
- Fall 3: Die Urkunde ist bereits bekannt und hat eine bestimmte ID. Ein in der Datenbank bereits vorhandener Datensatz wird ergänzt.

Für jeden der drei Fällen wird eine eigene CSV-Tabelle erstellt.
Die Abbildung der Tabelle auf das Datenbankschema erschöpft sich nicht in einem einfachen Mapping der Spalten. Einige der Einträge in der .docx-Tabelle enthalten Einträge, die ggf. aufgetrennt und in unterschiedliche Spalten geschrieben werden müssen.
Es mussten zahlreiche Umformungen und Vorarbeiten durchgeführt werden. das gewünschte Endergebnis wurde mir in Form einer Anforderungsspezifikation beschrieben. Einige Spezialfälle, die aufgrund von unsauberen Daten nicht von der regulären Software erfasst wurden, habe ich händisch angepasst. Die Umformung zu den CSV-Dateien ist daher nicht zu 100% mit der vorliegenden Software reproduzierbar.
