# Anleitung

## Zusammenfassung

Mit diesem Script, lässt man es per Cron aufrufen, lassen sich die bisherigen Antworten eines Nextcloud-Formulars als xlsx per Email versenden.

## Nötige Anpassungen

Die Datei credentials_copy.py muss umbenannt werden ist credentials.py.

Anschließend müssen in dieser Datei die...

- Zugangsdaten zum Nextcloudkonto, welches das Formular erstellt hat,
- die Zugangsdaten der Mailadresse, von der die Daten verschickt werden sollen,
- die Daten des SMTP-Servers der Mailadresse,
- der Empfänger der Mail,
- der Link zum Download der Formulardaten...

eingetragen werden. 