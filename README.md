# stripe-lexoffice-csv

Laedt alle Invoices von Stripe runter und baut eine CSV aus Transaktionen welche fuer den Import genutzt werden kann.

Benötigte Packages:

 - csv
 - stripe

Setup:

1. .env.example kopieren, in .env umbenennen und Parameter einstellen
2. script mit `python3 main.py` testen. Es sollte eine csv im csvs/ Ordner erstellt werden und per E-Mail verschickt werden.
3. Per Cronjob am 1. des Monats laufen lassen. Es wird der vergangene Monat abgerufen.