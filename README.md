# PAD Instruktor App

Aplikacja webowa/PWA dla instruktora Polskiej Armii Dronow do prowadzenia kursu, checklist, sprzetu, listy szkolonych, importu Excel, raportow DOCX oraz wymiany informacji miedzy instruktorami.

## Wersja online

Po wlaczeniu GitHub Pages aplikacja bedzie dostepna pod adresem:

```text
https://adasko1912.github.io/pad-instruktor-app/
```

## Instalacja na Androidzie

1. Otworz adres online w Chrome na Androidzie.
2. Poczekaj kilka sekund, az przegladarka wczyta aplikacje.
3. Kliknij przycisk `Zainstaluj` w aplikacji albo menu Chrome `...`.
4. Wybierz `Zainstaluj aplikacje` lub `Dodaj do ekranu glownego`.

## Uruchomienie lokalnie

```powershell
npm install
npm run dev
```

Adres lokalny:

```text
http://127.0.0.1:5173/
```

## Uruchomienie w sieci lokalnej

```powershell
npm run dev -- --host 0.0.0.0
```

Na telefonie w tej samej sieci Wi-Fi wejdz na:

```text
http://ADRES-IP-KOMPUTERA:5173/
```

## Build produkcyjny

```powershell
npm run build
```
