import React, { useEffect, useMemo, useState } from 'react';
import { saveAs } from 'file-saver';
import * as XLSX from 'xlsx';
import {
  AlignmentType,
  Document,
  HeadingLevel,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from 'docx';
import {
  AlertTriangle,
  Archive,
  Boxes,
  CalendarDays,
  CheckCircle2,
  ClipboardCheck,
  ClipboardList,
  Download,
  Edit3,
  FileSpreadsheet,
  FolderOpen,
  ListChecks,
  Menu,
  PackageCheck,
  Plane,
  Plus,
  Printer,
  Radio,
  RotateCcw,
  Search,
  ShieldCheck,
  Upload,
  MessageSquare,
  Save,
  Trash2,
  UserCheck,
  UserPlus,
  Users,
  Wrench,
  X,
} from 'lucide-react';

const today = new Date().toISOString().slice(0, 10);
const assetPath = (path) => `${import.meta.env.BASE_URL}${path}?v=2`;

const navItems = [
  { id: 'start', label: 'Start', icon: ClipboardCheck },
  { id: 'courses', label: 'Kursy', icon: FolderOpen },
  { id: 'pre-course', label: 'Przed kursem', icon: CheckCircle2 },
  { id: 'plan', label: 'Plan', icon: CalendarDays },
  { id: 'kit', label: 'Do zabrania', icon: PackageCheck },
  { id: 'trainee-kit', label: 'Sprzęt szkolonych', icon: Plane },
  { id: 'issues', label: 'Wydania', icon: Boxes },
  { id: 'trainees', label: 'Szkoleni', icon: Users },
  { id: 'instructors', label: 'Instruktorzy', icon: UserCheck },
];

const schedule = [
  {
    day: 'D1',
    title: 'Teoria i warsztat',
    tasks: [
      ['15:30-16:00', 'Zbiórka i wydanie', 'Podział na Z-1 do Z-10, wydanie sprzętu, podpisy odpowiedzialności materialnej.', 'Pełna obsada 40/40'],
      ['16:00-16:30', 'Otwarcie', 'Powitanie, prezentacja kadry, plan 3-dniowy, zasady BHP.', '100% obecność'],
      ['16:30-17:15', 'BSP - budowa', 'Mavic 3/3T/3 Pro: kadłub, ramiona, gimbal, sensory, akumulator, SD, kontroler.', 'Złożenie/rozłożenie < 60 s'],
      ['17:15-17:50', 'DJI - aplikacja', 'RTH 60 m, failsafe, limit 120 m, kalibracje, format SD, nagrywanie ekranu.', '8 parametrów bez błędu'],
      ['17:50-18:35', 'ATAK - podstawy', 'Warstwy, PZ/PO/OBJ/NAI/TAI/no-go, KML/KMZ, CoT, chat, zdjęcia GPS.', 'Eksport KML + import'],
      ['18:50-19:50', 'ATAK - wtyczki', 'Voice, DataSync, GRG, ExCheck, UAS Tool, SpotPlot, Reports, Multi Video 2.', 'GRG + ExCheck każdy kursant'],
      ['20:30-21:00', 'Łączność PACE', 'Starlink Primary, ATAK chat Alternate, PMR/UKF Contingency, GSM Emergency.', 'Test P-A-C w 30/60 s'],
      ['21:00-21:25', 'Maskowanie Starlinka', 'Odsunięcie min. 50 m od PO, siatka, kabel pod ziemią lub wzdłuż przeszkód.', 'Stanowisko pod skrajem lasu'],
      ['21:25-21:50', 'Streaming RTSP/RTMP', 'OFF tranzyt, ON nad celem, OFF powrót, alias UAS-Z-X w Multi Video 2.', 'Opóźnienie < 2 s'],
      ['21:50-22:15', 'Planowanie zadania', 'Pogoda, trasa, PZ, taktyka podejścia, plan B i briefing zespołu.', 'Plan w 10 min'],
      ['22:15-22:35', 'Tryb ATTI', 'GPS vs ATTI, dryf wiatru, REB/EW, warunki niedopuszczalne.', 'Zrozumienie ograniczeń'],
      ['22:35-22:55', 'BDA + zrzut markera', 'Moduł zrzutu, masa 50-250 g, procedura 6 kroków, HIT/NEAR/MISS.', 'Demo na ziemi'],
      ['22:55-23:20', 'AirData / KML', 'Utrata BSP: ekran, pozycja, AirData, KML do ATAK, sektor poszukiwań.', 'Punkt upadku z korektą'],
      ['23:20-23:45', 'Mini-test + rozkaz D2', 'Test 25 pytań, próg 70%, rozkaz i sektory na dzień 2.', 'Wynik min. 70%'],
    ],
  },
  {
    day: 'D2',
    title: 'Praktyka i start BLUE/RED',
    tasks: [
      ['07:30-08:00', 'Zbiórka', 'Obecność, status sprzętu zielony/żółty/czerwony w ATAK, briefing bezpieczeństwa.', '100% gotowości zespołów'],
      ['08:00-08:30', 'Briefing D2', 'Przydział sektorów, procedury, decyzja o dopuszczeniu do ATTI.', '10/10 sektorów zajętych'],
      ['08:30-11:00', 'Loty pojedyncze', 'Start z ziemi/ręki, hover, lądowanie, POI, lot bokiem, powroty, identyfikacja celu.', 'Min. 4 baterie/kursant'],
      ['11:15-12:30', 'Awarie i zaginiony BSP', 'Zdjęcie ekranu, AirData, KML, ATAK, R150 z korektą wiatru.', 'Lokalizacja < 60 min'],
      ['13:30-15:00', 'Sekcja 3-osobowa', 'Wskazywanie celów, komendy pilot-nawigator, rotacja ról, SALUTE.', 'Każdy w 4 rolach'],
      ['15:00-16:30', 'Lot na zakładkę', 'Sekwencja 8 kroków, ciągłość obserwacji, separacja wertykalna 30 m.', 'Bez utraty kontaktu'],
      ['16:30-18:00', 'Cele ruchome', 'Pojazd pozoracyjny, 30 min śledzenia, trasa w ATAK, stream.', 'Przerwa < 5 s'],
      ['18:00-19:00', 'Awarie + ATTI', 'Niska bateria, utrata sygnału, REB, wykrycie, pogoda, awaria silnika.', 'Min. 1 awaria/zespół'],
      ['19:30-20:00', 'Briefing EX', 'BLUE Z-1-Z-5, RED Z-6-Z-10, sektory, kanały Discord, warstwy ATAK.', 'Dowódcy zatwierdzeni'],
      ['20:00', 'EX-START + FRAGO', 'Każda drużyna otrzymuje FRAGO: sytuacja, misja, fazy A-F, NAI/TAI, łączność.', 'FRAGO odebrane'],
      ['20:00-22:00', 'Faza A - planowanie', 'APP-28(B): misja, IPOE, 2 COA, gra wojenna, wybór, decyzja, rozkaz.', 'Plan zatwierdzony'],
      ['22:00-24:00', 'Faza B - PO', 'Przemieszczenie, stanowisko, maskowanie, Starlink 50-100 m od PO, test PACE.', 'PO przed 23:50'],
    ],
  },
  {
    day: 'D3',
    title: 'Ćwiczenie, AAR i rozliczenie',
    tasks: [
      ['00:00-03:00', 'Faza C - NAI', 'Nieprzerwana obserwacja, GDR, drużyny BSP/FPV, SITREP co 30 min, INTREP po wykryciu.', 'Min. 1 GDR + 1 sekcja BSP'],
      ['03:00-04:00', 'S-01/S-02', 'A1 kartki i A6 zakłócenia EW w układzie BLUE/RED.', 'Kartki + antena WRE'],
      ['04:00-05:00', 'S-03/S-04', 'A2 snajper z balonem i A8 manekiny.', 'Zrzut + 6/6 manekinów'],
      ['05:00-06:00', 'S-05/S-06', 'A3 miny i A12 liczenie pojazdów.', '8/8 min + kolory pojazdów'],
      ['06:00-07:00', 'S-07/S-08', 'RU-1 Lancet i A7 śledzenie pojazdu.', 'CFF + 3 zatrzymania'],
      ['07:00-08:00', 'S-09/S-10', 'RU-2 Sznur i A4 WRE.', 'Operator + BDA WRE'],
      ['08:00-09:00', 'S-11/S-12', 'RU-3 ROSA i A11 zakładka 3-zespołowa.', 'ROSA + 30 min obserwacji'],
      ['09:00-10:00', 'S-13/S-14', 'RU-4 wąwóz i A10 patrol szeroki.', '5/5 punktów'],
      ['10:00-11:00', 'S-15/S-16', 'RU-5 Hangar Alfa i A5 wsparcie szturmu.', 'Stanowisko BSP + CFF'],
      ['11:00-12:00', 'S-17/S-18', 'AWR-01 skład MPS i AWR-02 sieć WRE.', 'INTREP < 5 min x2'],
      ['12:00-13:00', 'S-19/S-20', 'AWR-03 moździerz i AWR-04 zrzut zaopatrzenia.', 'Kaliber + patrol odbiorczy'],
      ['13:00-13:30', 'S-21', 'AWR-05 pojazd opancerzony.', 'Identyfikacja typu'],
      ['13:30-14:00', 'S-22 finałowe', 'Rozpoznanie obiektu ALFA i ostrzeżenie obrońców.', 'Pełne rozpoznanie ALFA'],
      ['14:00-16:00', 'Faza E - szturm ALFA', 'Wsparcie szturmu, naprowadzanie, BDA, ewakuacja, ciągły stream RTSP.', 'Stream 120 min'],
      ['16:00-16:30', 'FREEZE', 'Lądowanie BSP w 5 min, eksport logów, KML AirData, zdjęcia BDA, zwijanie.', '100% sprzętu zwinięte'],
      ['16:30-18:00', 'AAR strukturalny', 'Co miało się stać, co się stało, dlaczego, lekcje, certyfikaty.', '4 fazy AAR'],
      ['18:00-18:30', 'Rozliczenie', 'Zwrot sprzętu, ankiety, podziękowania, zakończenie.', '100% rozliczone'],
    ],
  },
].map((group) => ({
  ...group,
  tasks: group.tasks.map(([time, title, action, metric], index) => ({
    id: `${group.day}-${index + 1}`,
    time,
    title,
    action,
    metric,
  })),
}));

const instructorKit = [
  ['BSP i akcesoria', 'BSP DJI Mavic 3 / 3T / 3 Pro (główny)', '20'],
  ['BSP i akcesoria', 'Kontroler RC Pro / RC2', '10'],
  ['BSP i akcesoria', 'Akumulator BSP (Mavic 3 Intelligent)', '60'],
  ['BSP i akcesoria', 'Karty microSD 128 GB (U3/V30)', '40'],
  ['BSP i akcesoria', 'Tablet z DSI TAK + 8 wtyczek wojskowych', '20'],
  ['BSP i akcesoria', 'Smartfon zespołu (Discord, Signal, ATAK)', '10'],
  ['BSP i akcesoria', 'Komplety śmigieł zapasowych (8 szt. każdy)', '20'],
  ['BSP i akcesoria', 'Zestawy narzędzi (klucze, śrubokręty)', '10'],
  ['BSP i akcesoria', 'Worki/pojemniki LiPo na akumulatory', '20'],
  ['Łączność i zasilanie', 'Terminal Starlink Mini lub Standard', '10'],
  ['Łączność i zasilanie', 'Routery Wi-Fi (do streamu RTSP/RTMP)', '10'],
  ['Łączność i zasilanie', 'Stacja zasilania EcoFlow Delta 2 / Delta 2 Max', '10'],
  ['Łączność i zasilanie', 'Powerbanki 20 000 mAh', '20'],
  ['Łączność i zasilanie', 'Agregaty prądotwórcze 2 kW', '5'],
  ['Łączność i zasilanie', 'Kanistry paliwa 20 l (dla agregatów)', '8'],
  ['Łączność i zasilanie', 'Przedłużacze 25 m + rozdzielacze', '10'],
  ['Łączność i zasilanie', 'Radiotelefony PMR / UKF', '20'],
  ['Łączność i zasilanie', 'Latarki czołowe z czerwonym filtrem', '40'],
  ['Łączność i zasilanie', 'Lampy robocze LED z osłoną', '10'],
  ['Łączność i zasilanie', 'Anteny kierunkowe Yagi / panel', '10'],
  ['Łączność i zasilanie', 'Kable LMR-400 z przejściówkami', '10'],
  ['Materiały pozoracyjne', 'Atrapy granatów F1 (drewniane, malowane na zielono)', '10'],
  ['Materiały pozoracyjne', 'Atrapy granatów RGD-5', '10'],
  ['Materiały pozoracyjne', 'Atrapy granatów RG-42', '10'],
  ['Materiały pozoracyjne', 'Puszki 200 g do testów wagowych zrzutu', '20'],
  ['Materiały pozoracyjne', 'Manekiny szkolne (cele pozorowane)', '8'],
  ['Materiały pozoracyjne', 'Atrapy stanowisk WRE', '4'],
  ['Materiały pozoracyjne', 'Atrapy min plastikowych', '10'],
  ['Materiały pozoracyjne', 'Pojazdy pozoracyjne', '2'],
  ['Materiały pozoracyjne', 'Repliki ASG dla rozjemców', '6'],
  ['Materiały pozoracyjne', 'Balony helowe + butla helu', '10 + 1 butla'],
  ['Materiały pozoracyjne', 'Sznurek nylonowy zielony 800 m', '5 rolek po 200 m'],
  ['Materiały pozoracyjne', 'Atrapy wyrzutni Lancet (RU-1)', '2'],
  ['Materiały pozoracyjne', 'Atrapy radarów ROSA (RU-3)', '2'],
  ['Materiały pozoracyjne', 'Atrapy stanowiska hangaru', '2 zestawy'],
  ['Materiały pozoracyjne', 'Beczki 200 l (AWR-01)', '4'],
  ['Materiały pozoracyjne', 'Cysterna pozoracyjna', '1'],
  ['Materiały pozoracyjne', 'Maszty teleskopowe 4 m + atrapy anten', '3 + 3 atrapy'],
  ['Materiały pozoracyjne', 'Atrapa moździerza 60 mm', '1'],
  ['Materiały pozoracyjne', 'Pojazd terenowy z atrapą wieżyczki', '1'],
  ['Materiały pozoracyjne', 'Dymne do oznaczania pozycji', '20'],
  ['Materiały pozoracyjne', 'Flary do oznaczania PZ', '10'],
  ['Materiały pozoracyjne', 'Kartki A4 z literami do A1', '20 kompletów po 5'],
  ['Administracja i sala', 'Certyfikaty szkolenia', 'dla każdego uczestnika'],
  ['Administracja i sala', 'Dyplom dla właściciela terenu Ocieseki', '1'],
  ['Administracja i sala', 'Worki na śmieci', '10'],
  ['Administracja i sala', 'Miotły i szufelki', '4'],
  ['Administracja i sala', 'Woda', 'do uzupełnienia'],
  ['Administracja i sala', 'CASE / skrzynie transportowe', 'wszystkie'],
  ['Administracja i sala', 'Monitor interaktywny plus stojak', '1'],
  ['Administracja i sala', 'Krzesło turystyczne', '1'],
  ['Administracja i sala', 'Agregaty prądotwórcze', '3'],
  ['Administracja i sala', 'Benzyna kanister', '1'],
  ['Administracja i sala', 'Logbook', '50'],
  ['Administracja i sala', 'Butelki z piaskiem do oznaczania PZ', 'wszystkie'],
].map(([category, name, qty], index) => ({
  id: `kit-${index + 1}`,
  category,
  name,
  qty,
  source: 'base',
}));

const preCourseTasks = [
  ['T-30', 'Planowanie strategiczne', 'Potwierdzić termin, miejsce, kadrę, strukturę kursu i główne potrzeby sprzętowe.'],
  ['T-14', 'Logistyka materiałów pozoracyjnych', 'Zamknąć listę atrap, materiałów, pojazdów pozoracyjnych, pirotechniki i elementów rejonu.'],
  ['T-7', 'Finalizacja kursantów', 'Zweryfikować listę kursantów, braki sprzętu własnego, potrzeby wypożyczeń i obsadę zespołów.'],
  ['T-1', 'Pakowanie i kontrola sprzętu', 'Sprawdzić CASE, ładowanie akumulatorów, dokumenty wydania, listy obecności i materiały drukowane.'],
  ['T-0', 'Rano D1', 'Sprawdzić salę, projektor, łączność, miejsce zbiórki, oznaczenia i gotowość kadry do wydania sprzętu.'],
].map(([phase, title, description], index) => ({
  id: `pre-${index + 1}`,
  phase,
  title,
  description,
  source: 'base',
}));

const caseInventory = [
  ['CASE 1', 'Zestaw DJI MAVIC 3 ENTERPRISE 3E/T', 'Dron, kontroler, 5 baterii, 2 kable zasilające, ładowarka 200 W, karty pamięci, lokalizator FINDER'],
  ['CASE 2', 'Starlink Mini', 'Antena, zasilacz, kabel zasilający'],
  ['CASE 3', 'EcoFlow RIVER 2 MAX', 'Stacja zasilająca, instrukcja, 3 kable'],
  ['CASE 4', 'EcoFlow DELTA 2 (1)', 'Stacja zasilająca, 3 kable, instrukcja'],
  ['CASE 5', 'EcoFlow DELTA 2 (2)', 'Stacja zasilająca, 3 kable, instrukcja'],
  ['CASE 6', 'Starlink Standard (plecak)', 'Antena, router, zasilacz, kabel sygnałowy, maskowanie anteny, plecak'],
  ['CASE 7', 'DJI Air 2S + 2 kontrolery + 4Hawks', 'Dron, 2 kontrolery, 5 baterii, ładowarka, antena 4Hawks, zrzut, kable USB'],
  ['CASE 8', 'DJI AVATA 2 (FPV)', 'Dron, RC Motion 3, Goggles 3, 3 baterie, ładowarka, FPV Remote Controller 3'],
  ['CASE 9', 'DJI AIR 2 ze zrzutem', 'Dron, kontroler, 3 baterie, ładowarka, zrzut, FINDER, Case NILS'],
  ['CASE 10', 'Walizka z monitorem (1)', 'Monitor, HDMI-mini HDMI, adapter, USB-C, walizka'],
  ['CASE 11', 'Walizka z monitorem (2)', 'Monitor, HDMI-mini, adapter HDMI, USB-C, walizka'],
  ['CASE 12', 'Antena 4HAWKS RAPTOR XR', 'Antena, 2 kable 15 m, Case NILS'],
  ['CASE 13', 'Zrzuty Mavic 3T + Matrice 4T', 'Zrzut standardowy, Antylka, 2 zrzuty podczepiane, zrzut Matrice'],
  ['CASE 14', 'Tablety (5 sztuk + Graber)', 'Samsung Galaxy A9+ x4, Lenovo Yoga 13, Graber, Case NILS'],
  ['CASE 15', 'MAVIC 3 ENTERPRISE 3E/T (drugi)', 'Dron, kontroler RC Pro, 4 baterie, kable UK/EUR, ładowarka, FINDER'],
  ['CASE 16', 'DJI MATRICE 4T', 'Dron Matrice 4T, RC Plus 2, 4 baterie, karta 128 GB, ładowarki, lokalizator'],
  ['CASE 17', 'MAVIC 3T "Górlice"', 'Dron Mavic 3T, RC Pro, 4 baterie, GSM, RTK, 2 ładowarki, adapter, lokalizator'],
  ['CASE 18', 'DJI AIR 2S (T. Oksiński)', 'Dron Air 2S, 2 kontrolery, 3 baterie, zasilacz, listwa, adapter, lokalizator, torba'],
  ['CASE 19', 'Monitor interaktywny AVTEK 65', 'Monitor 65 cali, stojak, półka, kabel, pilot, pisaki, HDMI, klucze, instrukcje'],
  ['CASE 20', 'Maszt antenowy zestaw', 'Trójnóg, maszt, elementy mocujące, dodatkowy uchwyt'],
  ['CASE 21', 'Maskowanie - 7 siatek Woodland', '5 siatek Woodland, 2 dodatkowe siatki maskujące'],
  ['CASE 22', 'Agregat TRYTON (1)', 'Agregat, olejarka, klucz do świecy, śrubokręt, przewód, instrukcja'],
  ['CASE 23', 'Agregat TRYTON (2)', 'Agregat, olejarka, klucz do świecy, śrubokręt, przewód, instrukcja'],
  ['CASE 24', 'Agregat TRYTON (3)', 'Agregat, klucz do świecy, śrubokręt, przewód, olejarka, instrukcja'],
  ['CASE 25', 'ALLPOWERS 1500 Lite (1)', 'Stacja zasilająca, kabel 230 V, instrukcja'],
  ['CASE 26', 'ALLPOWERS 1500 Lite (2)', 'Stacja zasilająca, kabel 230 V, instrukcja'],
  ['CASE 27', 'Torby przeciwpożarowe TORVOL', '3 torby TORVOL LiPo'],
  ['CASE 28', 'Stolik turystyczny', 'Stolik turystyczny składany, pudło kartonowe'],
  ['CASE 29', 'Kable Ethernet 100 m', '2 kable Ethernet 100 m'],
].map(([caseNo, name, contents]) => ({
  id: caseNo.toLowerCase().replace(' ', '-'),
  caseNo,
  name,
  contents,
  issuedTo: '',
  issuedDate: '',
  returnDate: '',
  status: 'magazyn',
  notes: '',
}));

const traineeKit = [
  ['BSP', 'BSP klasy DJI Mavic 3 / 3T / 3 Pro lub równoważny'],
  ['BSP', 'BSP zapasowy lub szkoleniowy'],
  ['BSP', 'Kontroler RC Pro / RC2 z ekranem'],
  ['BSP', 'Karty microSD 128 GB U3/V30 - 4 szt.'],
  ['BSP', 'Komplet śmigieł zapasowych - 16 szt.'],
  ['BSP', 'Worek/pojemnik LiPo - 2 szt.'],
  ['BSP', 'Zestaw narzędzi serwisowych'],
  ['Zasilanie', 'Akumulatory BSP - min. 6 szt.'],
  ['Zasilanie', 'Szybkoładowarka min. 65 W/slot'],
  ['Zasilanie', 'Stacja zasilania min. 1000 Wh'],
  ['Zasilanie', 'Powerbank min. 20 000 mAh'],
  ['Zasilanie', 'Przedłużacz 25 m + listwa'],
  ['Łączność', 'Terminal Starlink Mini lub Standard'],
  ['Łączność', 'Tablet Android/iOS do DSI TAK + DJI'],
  ['Łączność', 'Smartfon do Discord / Signal / ATAK'],
  ['Łączność', 'Router Wi-Fi i kable HDMI/USB-C'],
  ['Stanowisko', 'Monitor zewnętrzny 7-10 cali'],
  ['Stanowisko', 'Siatka maskująca 3 x 6 m'],
  ['Stanowisko', 'Tyczki / maszty / paracord'],
  ['Stanowisko', 'Latarka czołowa z czerwonym filtrem'],
  ['Osobiste', 'IFAK i leki własne'],
  ['Osobiste', 'Wyżywienie 9000 kcal i woda 9 l'],
  ['Osobiste', 'Śpiwór, karimata, plandeka'],
  ['Osobiste', 'Mundur, buty, odzież przeciwdeszczowa'],
].map(([category, name], index) => ({
  id: `tk-${index + 1}`,
  category,
  name,
}));

const equipmentStatuses = ['nie sprawdzono', 'ma', 'brak', 'do wypożyczenia', 'wydano'];
const issueStatuses = ['magazyn', 'wydane', 'zwrócone', 'uszkodzone', 'brak'];
const teams = Array.from({ length: 10 }, (_, index) => `Z-${index + 1}`);
const emptyTrainee = {
  rank: '',
  firstName: '',
  lastName: '',
  phone: '',
  unit: '',
  team: 'Z-1',
  group: 'BLUE',
  role: 'Pilot',
};

const emptyCourseForm = {
  name: '',
  dateFrom: today,
  dateTo: today,
  location: '',
};

const emptyCustomKitForm = {
  category: 'Własne',
  name: '',
  qty: '1',
};

const emptyCustomTaskForm = {
  phase: 'T-7',
  title: '',
  description: '',
};

const emptyInstructorProfile = {
  rank: '',
  firstName: '',
  lastName: '',
  phone: '',
  unit: '',
};

const emptyInstructorForm = {
  rank: '',
  firstName: '',
  lastName: '',
  phone: '',
  unit: '',
  role: 'Instruktor',
  participating: true,
};

const emptyUpdateForm = {
  category: 'Postęp szkolenia',
  message: '',
};

const traineeColumns = [
  ['Stopień', 'rank'],
  ['Imię', 'firstName'],
  ['Nazwisko', 'lastName'],
  ['Telefon', 'phone'],
  ['Jednostka', 'unit'],
  ['Zespół', 'team'],
  ['Grupa', 'group'],
  ['Rola', 'role'],
];

const traineeColumnAliases = {
  rank: ['stopien', 'stopie', 'stop'],
  firstName: ['imie', 'imi'],
  lastName: ['nazwisko', 'nazw'],
  phone: ['telefon', 'tel', 'nrtelefonu', 'numertelefonu'],
  unit: ['jednostka', 'jedn'],
  team: ['zespol', 'zespo', 'zesp', 'team'],
  group: ['grupa', 'group'],
  role: ['rola', 'role'],
};

function createCourse(meta = {}) {
  return {
    id: uid('kurs'),
    name: meta.name || 'Kurs BSP',
    dateFrom: meta.dateFrom || today,
    dateTo: meta.dateTo || today,
    location: meta.location || '',
    createdAt: new Date().toISOString(),
    status: 'aktywny',
    startedAt: new Date().toISOString(),
    endedAt: '',
    progress: {},
    preCourseProgress: {},
    carryState: {},
    trainees: [],
    equipmentRegister: {},
    caseLedger: caseInventory,
    instructors: [],
    updates: [],
  };
}

function createMigratedCourses() {
  const fallback = createCourse({ name: 'Kurs BSP', dateFrom: today, dateTo: today });
  try {
    const oldProgress = JSON.parse(window.localStorage.getItem('pad-instruktor.progress.v1') || '{}');
    const oldCarryState = JSON.parse(window.localStorage.getItem('pad-instruktor.carry.v1') || '{}');
    const oldTrainees = JSON.parse(window.localStorage.getItem('pad-instruktor.trainees.v1') || '[]');
    const oldEquipmentRegister = JSON.parse(window.localStorage.getItem('pad-instruktor.trainee-kit.v1') || '{}');
    const oldCaseLedger = JSON.parse(window.localStorage.getItem('pad-instruktor.cases.v1') || 'null') || caseInventory;
    const hasOldData =
      Object.keys(oldProgress).length ||
      Object.keys(oldCarryState).length ||
      oldTrainees.length ||
      Object.keys(oldEquipmentRegister).length ||
      oldCaseLedger.some((item) => item.status !== 'magazyn' || item.issuedTo || item.notes);

    if (!hasOldData) return [fallback];

    return [
      {
        ...fallback,
        name: 'Kurs przeniesiony',
        progress: oldProgress,
        preCourseProgress: {},
        carryState: oldCarryState,
        trainees: oldTrainees,
        equipmentRegister: oldEquipmentRegister,
        caseLedger: oldCaseLedger,
      },
    ];
  } catch {
    return [fallback];
  }
}

function useLocalStorage(key, initialValue) {
  const [value, setValue] = useState(() => {
    try {
      const stored = window.localStorage.getItem(key);
      return stored ? JSON.parse(stored) : initialValue;
    } catch {
      return initialValue;
    }
  });

  const updateValue = (nextValue) => {
    setValue((previous) => {
      const resolved = typeof nextValue === 'function' ? nextValue(previous) : nextValue;
      window.localStorage.setItem(key, JSON.stringify(resolved));
      return resolved;
    });
  };

  return [value, updateValue];
}

function uid(prefix) {
  return `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 7)}`;
}

function percent(done, total) {
  if (!total) return 0;
  return Math.round((done / total) * 100);
}

function normalize(value) {
  return String(value || '').toLowerCase();
}

function fullName(person) {
  return [person.rank, person.firstName, person.lastName].filter(Boolean).join(' ');
}

function isDone(value) {
  return Boolean(value === true || value?.done);
}

function completedAt(value) {
  return value?.completedAt || '';
}

function doneEntry(done) {
  return done ? { done: true, completedAt: new Date().toISOString() } : null;
}

function formatDateTime(value) {
  if (!value) return '-';
  try {
    return new Date(value).toLocaleString('pl-PL');
  } catch {
    return value;
  }
}

function personKey(person) {
  const phone = normalizePhone(person.phone);
  if (phone) return `phone:${phone}`;
  return `name:${normalize(`${person.firstName}|${person.lastName}|${person.unit}`)}`;
}

function teamOrder(team) {
  const match = String(team || '').match(/\d+/);
  return match ? Number(match[0]) : 99;
}

function sortedTrainees(trainees) {
  return [...trainees].sort((a, b) => {
    const groupA = a.group === 'BLUE' ? 0 : a.group === 'RED' ? 1 : 2;
    const groupB = b.group === 'BLUE' ? 0 : b.group === 'RED' ? 1 : 2;
    if (groupA !== groupB) return groupA - groupB;
    const teamDelta = teamOrder(a.team) - teamOrder(b.team);
    if (teamDelta) return teamDelta;
    return `${a.lastName} ${a.firstName}`.localeCompare(`${b.lastName} ${b.firstName}`, 'pl');
  });
}

function groupTrainees(trainees) {
  return sortedTrainees(trainees).reduce((groups, trainee) => {
    const group = trainee.group || 'BRAK';
    const team = trainee.team || 'Bez zespołu';
    groups[group] = groups[group] || {};
    groups[group][team] = [...(groups[group][team] || []), trainee];
    return groups;
  }, {});
}

function assigneeLabel(value, trainees) {
  if (!value) return 'Nie wydano';
  if (value.startsWith('trainee:')) {
    const trainee = trainees.find((item) => item.id === value.replace('trainee:', ''));
    return trainee ? fullName(trainee) : 'Usunięty szkolony';
  }
  if (value.startsWith('team:')) return value.replace('team:', '');
  return value;
}

function exportCsv(filename, rows) {
  const csv = rows
    .map((row) =>
      row
        .map((cell) => `"${String(cell ?? '').replace(/"/g, '""')}"`)
        .join(';'),
    )
    .join('\n');
  const blob = new Blob([`\ufeff${csv}`], { type: 'text/csv;charset=utf-8;' });
  const href = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = href;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(href);
}

function downloadBlob(filename, blob) {
  saveAs(blob, filename);
}

function sanitizeFilename(value) {
  return String(value || 'kurs')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9_-]+/g, '_')
    .replace(/^_+|_+$/g, '')
    .slice(0, 80);
}

function normalizeHeader(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9]+/g, '')
    .trim()
    .toLowerCase();
}

function normalizePhone(value) {
  return String(value || '').replace(/[^\d+]/g, '');
}

function normalizeTeam(value) {
  const raw = String(value || '').trim().toUpperCase();
  const match = raw.match(/\d+/);
  return match ? `Z-${Math.min(Math.max(Number(match[0]), 1), 10)}` : 'Z-1';
}

function normalizeGroup(value) {
  const raw = String(value || '').trim().toUpperCase();
  return raw === 'RED' ? 'RED' : 'BLUE';
}

function normalizeRole(value) {
  const raw = String(value || '').trim();
  return raw || 'Pilot';
}

function dedupeKey(trainee) {
  const phone = normalizePhone(trainee.phone);
  if (phone) return `phone:${phone}`;
  return `name:${normalize(`${trainee.firstName}|${trainee.lastName}|${trainee.unit}`)}`;
}

function parseTraineesWorkbook(arrayBuffer, existingTrainees = []) {
  const workbook = XLSX.read(arrayBuffer, { type: 'array' });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false });
  const firstRow = rows[0] || {};
  const headers = Object.keys(firstRow);
  const headerMap = traineeColumns.reduce((map, [label, field]) => {
    const aliases = [normalizeHeader(label), ...(traineeColumnAliases[field] || [])];
    const found = headers.find((key) => {
      const normalized = normalizeHeader(key);
      return aliases.some((alias) => normalized === alias || normalized.startsWith(alias));
    });
    if (found) map[field] = found;
    return map;
  }, {});
  const missingColumns = traineeColumns
    .filter(([, field]) => !headerMap[field])
    .map(([label]) => label);

  if (missingColumns.length) {
    return {
      added: [],
      skipped: [],
      errors: [`Brak wymaganych kolumn: ${missingColumns.join(', ')}`],
    };
  }

  const seen = new Set(existingTrainees.map(dedupeKey));
  const added = [];
  const skipped = [];
  const errors = [];

  rows.forEach((row, index) => {
    const trainee = traineeColumns.reduce((item, [label, field]) => {
      item[field] = String(row[headerMap[field]] || '').trim();
      return item;
    }, {});

    trainee.team = normalizeTeam(trainee.team);
    trainee.group = normalizeGroup(trainee.group);
    trainee.role = normalizeRole(trainee.role);

    if (!trainee.firstName || !trainee.lastName) {
      errors.push(`Wiersz ${index + 2}: wymagane imię i nazwisko.`);
      return;
    }

    const key = dedupeKey(trainee);
    if (seen.has(key)) {
      skipped.push(`${trainee.firstName} ${trainee.lastName}`);
      return;
    }

    seen.add(key);
    added.push({ id: uid('kursant'), ...trainee });
  });

  return { added, skipped, errors };
}

function docParagraph(text, options = {}) {
  return new Paragraph({
    text,
    spacing: { after: 120 },
    ...options,
  });
}

function docCell(text, bold = false) {
  return new TableCell({
    children: [
      new Paragraph({
        children: [new TextRun({ text: String(text || '-'), bold })],
      }),
    ],
  });
}

function docTable(headers, rows) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({ children: headers.map((header) => docCell(header, true)) }),
      ...rows.map((row) => new TableRow({ children: row.map((cell) => docCell(cell)) })),
    ],
  });
}

async function exportCourseDocx(course, kitItems = instructorKit, allPreCourseTasks = preCourseTasks) {
  const trainees = sortedTrainees(course.trainees || []);
  const instructorRows = (course.instructors || []).map((item) => [
    item.participating ? 'Bierze udział' : 'Nie bierze udziału',
    fullName(item),
    item.role,
    item.phone,
    item.unit,
  ]);
  const updateRows = (course.updates || []).map((item) => [
    new Date(item.createdAt).toLocaleString('pl-PL'),
    item.category,
    item.authorName,
    item.message,
  ]);
  const preCourseRows = allPreCourseTasks.map((task) => [
    task.phase,
    task.title,
    task.description,
    isDone(course.preCourseProgress?.[task.id]) ? 'Wykonano' : 'Do wykonania',
    formatDateTime(completedAt(course.preCourseProgress?.[task.id])),
    task.source === 'custom' ? 'Własne' : 'Stałe',
  ]);
  const taskRows = schedule.flatMap((group) =>
    group.tasks.map((task) => [
      group.day,
      task.time,
      task.title,
      isDone(course.progress?.[task.id]) ? 'Wykonano' : 'Do wykonania',
      formatDateTime(completedAt(course.progress?.[task.id])),
      task.metric,
    ]),
  );
  const kitRows = kitItems.map((item) => {
    const state = course.carryState?.[item.id] || {};
    return [item.category, item.name, item.qty, state.status || (state.packed ? 'zabrane' : 'do zabrania'), state.note || '', item.source === 'custom' ? 'Własne' : 'Stałe'];
  });
  const missingRows = trainees.flatMap((trainee) => {
    const register = course.equipmentRegister?.[trainee.id] || {};
    return traineeKit
      .filter((item) => ['brak', 'do wypożyczenia'].includes(register[item.id]))
      .map((item) => [fullName(trainee), trainee.team, trainee.group, item.category, item.name, register[item.id]]);
  });
  const caseRows = (course.caseLedger || []).map((item) => [
    item.caseNo,
    item.name,
    assigneeLabel(item.issuedTo, trainees),
    item.issuedDate,
    item.returnDate,
    item.status,
    item.notes,
  ]);

  const document = new Document({
    sections: [
      {
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            heading: HeadingLevel.TITLE,
            children: [new TextRun({ text: 'Raport kursu PAD Instruktor', bold: true, size: 34 })],
          }),
          docParagraph(`Nazwa: ${course.name}`),
          docParagraph(`Termin: ${course.dateFrom || '-'} - ${course.dateTo || '-'}`),
          docParagraph(`Miejsce: ${course.location || '-'}`),
          docParagraph(`Status: ${course.status || 'aktywny'}`),
          docParagraph(`Rozpoczęto: ${formatDateTime(course.startedAt || course.createdAt)}`),
          docParagraph(`Zakończono: ${formatDateTime(course.endedAt)}`),
          docParagraph(`Liczba szkolonych: ${trainees.length}`),
          new Paragraph({ text: 'Instruktorzy kursu', heading: HeadingLevel.HEADING_1 }),
          instructorRows.length
            ? docTable(['Udział', 'Instruktor', 'Rola', 'Telefon', 'Jednostka'], instructorRows)
            : docParagraph('Brak przypisanych instruktorów.'),
          new Paragraph({ text: 'Wymiana informacji instruktorów', heading: HeadingLevel.HEADING_1 }),
          updateRows.length
            ? docTable(['Czas', 'Kategoria', 'Autor', 'Informacja'], updateRows)
            : docParagraph('Brak wpisów instruktorów.'),
          new Paragraph({ text: 'Lista szkolonych', heading: HeadingLevel.HEADING_1 }),
          docTable(
            ['Grupa', 'Zespół', 'Stopień', 'Imię', 'Nazwisko', 'Telefon', 'Jednostka', 'Rola'],
            trainees.map((item) => [item.group, item.team, item.rank, item.firstName, item.lastName, item.phone, item.unit, item.role]),
          ),
          new Paragraph({ text: 'Do zrobienia przed kursem', heading: HeadingLevel.HEADING_1 }),
          docTable(['Termin', 'Zadanie', 'Opis', 'Status', 'Źródło'], preCourseRows),
          new Paragraph({ text: 'Postęp planu D1-D3', heading: HeadingLevel.HEADING_1 }),
          docTable(['Dzień', 'Czas', 'Blok', 'Status', 'Miernik'], taskRows),
          new Paragraph({ text: 'Sprzęt do zabrania', heading: HeadingLevel.HEADING_1 }),
          docTable(['Kategoria', 'Element', 'Ilość', 'Status', 'Uwagi', 'Źródło'], kitRows),
          new Paragraph({ text: 'Braki sprzętu szkolonych', heading: HeadingLevel.HEADING_1 }),
          missingRows.length
            ? docTable(['Szkolony', 'Zespół', 'Grupa', 'Kategoria', 'Element', 'Status'], missingRows)
            : docParagraph('Brak zarejestrowanych braków.'),
          new Paragraph({ text: 'Wydania i zwroty CASE', heading: HeadingLevel.HEADING_1 }),
          docTable(['CASE', 'Walizka', 'Dla kogo', 'Data wydania', 'Data zwrotu', 'Status', 'Uwagi'], caseRows),
        ],
      },
    ],
  });

  const blob = await Packer.toBlob(document);
  downloadBlob(`PAD_raport_kursu_${course.dateFrom || today}_${sanitizeFilename(course.name)}.docx`, blob);
}

function IconButton({ icon: Icon, label, variant = 'default', ...props }) {
  return (
    <button className={`icon-button ${variant}`} type="button" title={label} aria-label={label} {...props}>
      <Icon size={17} />
    </button>
  );
}

function ActionButton({ icon: Icon, children, variant = 'primary', ...props }) {
  return (
    <button className={`action-button ${variant}`} type="button" {...props}>
      {Icon ? <Icon size={17} /> : null}
      <span>{children}</span>
    </button>
  );
}

function SearchBox({ value, onChange, placeholder = 'Szukaj' }) {
  return (
    <label className="search-box">
      <Search size={16} />
      <input value={value} onChange={(event) => onChange(event.target.value)} placeholder={placeholder} />
    </label>
  );
}

function StatusPill({ value }) {
  const className = `status-pill ${normalize(value).replaceAll(' ', '-')}`;
  return <span className={className}>{value}</span>;
}

function SectionHeader({ kicker, title, children }) {
  return (
    <div className="section-header">
      <div>
        <span className="section-kicker">{kicker}</span>
        <h2>{title}</h2>
      </div>
      <div className="section-actions">{children}</div>
    </div>
  );
}

function MetricCard({ icon: Icon, label, value, sublabel, tone = 'olive' }) {
  return (
    <div className={`metric-card ${tone}`}>
      <div className="metric-icon">
        <Icon size={20} />
      </div>
      <div>
        <span>{label}</span>
        <strong>{value}</strong>
        <small>{sublabel}</small>
      </div>
    </div>
  );
}

function ProgressDial({ value, label }) {
  return (
    <div className="progress-dial" style={{ '--progress': `${value}%` }}>
      <div>
        <strong>{value}%</strong>
        <span>{label}</span>
      </div>
    </div>
  );
}

function App() {
  const [activeView, setActiveView] = useState('courses');
  const [isMenuOpen, setIsMenuOpen] = useState(false);
  const [activeDay, setActiveDay] = useState('D1');
  const [courses, setCourses] = useLocalStorage('pad-instruktor.courses.v1', createMigratedCourses());
  const [activeCourseId, setActiveCourseId] = useLocalStorage('pad-instruktor.active-course.v1', courses[0]?.id || '');
  const [instructorProfile, setInstructorProfile] = useLocalStorage('pad-instruktor.instructor-profile.v1', emptyInstructorProfile);
  const [customKitItems, setCustomKitItems] = useLocalStorage('pad-instruktor.custom-kit.v1', []);
  const [customPreCourseTasks, setCustomPreCourseTasks] = useLocalStorage('pad-instruktor.custom-pre-course.v1', []);
  const [search, setSearch] = useState('');
  const [kitFilter, setKitFilter] = useState('wszystko');
  const [form, setForm] = useState(emptyTrainee);
  const [courseForm, setCourseForm] = useState(emptyCourseForm);
  const [customKitForm, setCustomKitForm] = useState(emptyCustomKitForm);
  const [customTaskForm, setCustomTaskForm] = useState(emptyCustomTaskForm);
  const [instructorForm, setInstructorForm] = useState(emptyInstructorForm);
  const [updateForm, setUpdateForm] = useState(emptyUpdateForm);
  const [editingId, setEditingId] = useState(null);
  const [selectedTraineeId, setSelectedTraineeId] = useState('');
  const [traineeViewMode, setTraineeViewMode] = useState('table');
  const [importReport, setImportReport] = useState(null);
  const [installPrompt, setInstallPrompt] = useState(null);
  const [isInstalled, setIsInstalled] = useState(() => window.matchMedia?.('(display-mode: standalone)').matches || window.navigator.standalone);

  const activeCourse = courses.find((course) => course.id === activeCourseId) || courses[0] || createCourse();
  const progress = activeCourse.progress || {};
  const carryState = activeCourse.carryState || {};
  const trainees = activeCourse.trainees || [];
  const equipmentRegister = activeCourse.equipmentRegister || {};
  const caseLedger = activeCourse.caseLedger || caseInventory;
  const participatingInstructors = (activeCourse.instructors || []).filter((item) => item.participating);
  const profileCourseEntry = (activeCourse.instructors || []).find((item) => item.profileId === 'self');
  const allKitItems = useMemo(() => [...instructorKit, ...customKitItems], [customKitItems]);
  const allPreCourseTasks = useMemo(() => [...preCourseTasks, ...customPreCourseTasks], [customPreCourseTasks]);

  const flatTasks = useMemo(() => schedule.flatMap((group) => group.tasks), []);
  const doneTasks = flatTasks.filter((task) => isDone(progress[task.id])).length;
  const preCourseDone = allPreCourseTasks.filter((task) => isDone(activeCourse.preCourseProgress?.[task.id])).length;
  const packedItems = allKitItems.filter((item) => carryState[item.id]?.packed).length;
  const openIssues = caseLedger.filter((item) => item.status === 'wydane').length;
  const missingGear = trainees.reduce((sum, trainee) => {
    const register = equipmentRegister[trainee.id] || {};
    return sum + traineeKit.filter((item) => ['brak', 'do wypożyczenia'].includes(register[item.id])).length;
  }, 0);
  const activeTraineeId = selectedTraineeId || trainees[0]?.id || '';
  const activeTrainee = trainees.find((item) => item.id === activeTraineeId);

  const filteredKit = allKitItems.filter((item) => {
    const textMatch = normalize(`${item.category} ${item.name} ${item.qty}`).includes(normalize(search));
    const state = carryState[item.id];
    const statusMatch =
      kitFilter === 'wszystko' ||
      (kitFilter === 'zabrane' && state?.packed) ||
      (kitFilter === 'braki' && state?.status === 'brak') ||
      (kitFilter === 'do zabrania' && !state?.packed);
    return textMatch && statusMatch;
  });

  const filteredTrainees = sortedTrainees(
    trainees.filter((trainee) =>
      normalize(`${fullName(trainee)} ${trainee.phone} ${trainee.unit} ${trainee.team} ${trainee.group} ${trainee.role}`).includes(normalize(search)),
    ),
  );

  const filteredCases = caseLedger.filter((item) =>
    normalize(`${item.caseNo} ${item.name} ${item.contents} ${assigneeLabel(item.issuedTo, trainees)} ${item.status}`).includes(normalize(search)),
  );

  useEffect(() => {
    const handleBeforeInstallPrompt = (event) => {
      event.preventDefault();
      setInstallPrompt(event);
    };
    const handleInstalled = () => {
      setIsInstalled(true);
      setInstallPrompt(null);
    };

    window.addEventListener('beforeinstallprompt', handleBeforeInstallPrompt);
    window.addEventListener('appinstalled', handleInstalled);

    return () => {
      window.removeEventListener('beforeinstallprompt', handleBeforeInstallPrompt);
      window.removeEventListener('appinstalled', handleInstalled);
    };
  }, []);

  const installApp = async () => {
    if (!installPrompt) return;
    installPrompt.prompt();
    const choice = await installPrompt.userChoice;
    if (choice.outcome === 'accepted') {
      setIsInstalled(true);
    }
    setInstallPrompt(null);
  };

  const updateActiveCourse = (updater) => {
    setCourses((previous) =>
      previous.map((course) => (course.id === activeCourse.id ? { ...course, ...updater(course) } : course)),
    );
  };

  const updateCourseField = (field, updater) => {
    updateActiveCourse((course) => ({
      [field]: typeof updater === 'function' ? updater(course[field]) : updater,
    }));
  };

  const toggleTask = (taskId) => {
    updateCourseField('progress', (previous) => ({ ...previous, [taskId]: doneEntry(!isDone(previous[taskId])) }));
  };

  const updateCarryItem = (itemId, patch) => {
    updateCourseField('carryState', (previous) => ({
      ...previous,
      [itemId]: { packed: false, status: 'do zabrania', note: '', ...(previous[itemId] || {}), ...patch },
    }));
  };

  const updateCase = (caseId, patch) => {
    updateCourseField('caseLedger', (previous) => previous.map((item) => (item.id === caseId ? { ...item, ...patch } : item)));
  };

  const updateEquipment = (traineeId, itemId, status) => {
    updateCourseField('equipmentRegister', (previous) => ({
      ...previous,
      [traineeId]: { ...(previous[traineeId] || {}), [itemId]: status },
    }));
  };

  const togglePreCourseTask = (taskId) => {
    updateCourseField('preCourseProgress', (previous) => ({ ...previous, [taskId]: doneEntry(!isDone(previous[taskId])) }));
  };

  const addCustomKitItem = (event) => {
    event.preventDefault();
    if (!customKitForm.name.trim()) return;
    setCustomKitItems((previous) => [
      ...previous,
      {
        id: uid('custom-kit'),
        category: customKitForm.category.trim() || 'Własne',
        name: customKitForm.name.trim(),
        qty: customKitForm.qty.trim() || '1',
        source: 'custom',
      },
    ]);
    setCustomKitForm(emptyCustomKitForm);
  };

  const deleteCustomKitItem = (itemId) => {
    setCustomKitItems((previous) => previous.filter((item) => item.id !== itemId));
    setCourses((previous) =>
      previous.map((course) => {
        const next = { ...(course.carryState || {}) };
        delete next[itemId];
        return { ...course, carryState: next };
      }),
    );
  };

  const addCustomPreCourseTask = (event) => {
    event.preventDefault();
    if (!customTaskForm.title.trim()) return;
    setCustomPreCourseTasks((previous) => [
      ...previous,
      {
        id: uid('custom-pre'),
        phase: customTaskForm.phase.trim() || 'T-7',
        title: customTaskForm.title.trim(),
        description: customTaskForm.description.trim(),
        source: 'custom',
      },
    ]);
    setCustomTaskForm(emptyCustomTaskForm);
  };

  const deleteCustomPreCourseTask = (taskId) => {
    setCustomPreCourseTasks((previous) => previous.filter((task) => task.id !== taskId));
    setCourses((previous) =>
      previous.map((course) => {
        const next = { ...(course.preCourseProgress || {}) };
        delete next[taskId];
        return { ...course, preCourseProgress: next };
      }),
    );
  };

  const saveInstructorProfile = (event) => {
    event.preventDefault();
    setInstructorProfile(instructorProfile);
  };

  const markProfileParticipating = () => {
    if (!instructorProfile.firstName.trim() || !instructorProfile.lastName.trim()) {
      alert('Najpierw wpisz imię i nazwisko w profilu instruktora.');
      return;
    }
    updateActiveCourse((course) => {
      const existing = course.instructors || [];
      const selfEntry = {
        id: existing.find((item) => item.profileId === 'self')?.id || uid('instr'),
        profileId: 'self',
        rank: instructorProfile.rank,
        firstName: instructorProfile.firstName,
        lastName: instructorProfile.lastName,
        phone: instructorProfile.phone,
        unit: instructorProfile.unit,
        role: 'Instruktor',
        participating: true,
      };
      return {
        instructors: [...existing.filter((item) => item.profileId !== 'self'), selfEntry],
      };
    });
  };

  const addInstructorToCourse = (event) => {
    event.preventDefault();
    if (!instructorForm.firstName.trim() || !instructorForm.lastName.trim()) return;
    updateActiveCourse((course) => {
      const existing = course.instructors || [];
      const key = personKey(instructorForm);
      const withoutDuplicate = existing.filter((item) => personKey(item) !== key);
      return {
        instructors: [
          ...withoutDuplicate,
          {
            id: uid('instr'),
            ...instructorForm,
          },
        ],
      };
    });
    setInstructorForm(emptyInstructorForm);
  };

  const updateInstructorInCourse = (instructorId, patch) => {
    updateCourseField('instructors', (previous) =>
      (previous || []).map((item) => (item.id === instructorId ? { ...item, ...patch } : item)),
    );
  };

  const deleteInstructorFromCourse = (instructorId) => {
    updateCourseField('instructors', (previous) => (previous || []).filter((item) => item.id !== instructorId));
  };

  const addInstructorUpdate = (event) => {
    event.preventDefault();
    if (!updateForm.message.trim()) return;
    if (!profileCourseEntry?.participating) {
      alert('Wpisz profil instruktora i oznacz, że bierzesz udział w tym szkoleniu.');
      return;
    }
    updateCourseField('updates', (previous) => [
      {
        id: uid('info'),
        createdAt: new Date().toISOString(),
        category: updateForm.category,
        message: updateForm.message.trim(),
        authorId: profileCourseEntry.id,
        authorName: fullName(profileCourseEntry),
      },
      ...(previous || []),
    ]);
    setUpdateForm(emptyUpdateForm);
  };

  const deleteInstructorUpdate = (updateId) => {
    updateCourseField('updates', (previous) => (previous || []).filter((item) => item.id !== updateId));
  };

  const saveTrainee = (event) => {
    event.preventDefault();
    if (!form.firstName.trim() || !form.lastName.trim()) return;

    if (editingId) {
      updateCourseField('trainees', (previous) => previous.map((item) => (item.id === editingId ? { ...item, ...form } : item)));
      setEditingId(null);
    } else {
      const id = uid('kursant');
      updateActiveCourse((course) => ({
        trainees: sortedTrainees([...(course.trainees || []), { id, ...form }]),
        equipmentRegister: { ...(course.equipmentRegister || {}), [id]: {} },
      }));
      setSelectedTraineeId(id);
    }

    setForm(emptyTrainee);
  };

  const editTrainee = (trainee) => {
    setForm({
      rank: trainee.rank || '',
      firstName: trainee.firstName || '',
      lastName: trainee.lastName || '',
      phone: trainee.phone || '',
      unit: trainee.unit || '',
      team: trainee.team || 'Z-1',
      group: trainee.group || 'BLUE',
      role: trainee.role || 'Pilot',
    });
    setEditingId(trainee.id);
    setActiveView('trainees');
  };

  const deleteTrainee = (traineeId) => {
    updateActiveCourse((course) => {
      const next = { ...(course.equipmentRegister || {}) };
      delete next[traineeId];
      return {
        trainees: (course.trainees || []).filter((item) => item.id !== traineeId),
        equipmentRegister: next,
        caseLedger: (course.caseLedger || []).map((item) =>
          item.issuedTo === `trainee:${traineeId}` ? { ...item, issuedTo: '', status: 'magazyn' } : item,
        ),
      };
    });
    if (selectedTraineeId === traineeId) setSelectedTraineeId('');
  };

  const resetAll = () => {
    if (!window.confirm('Wyczyścić zapisane dane aplikacji?')) return;
    const course = createCourse({ name: 'Kurs BSP', dateFrom: today, dateTo: today });
    setCourses([course]);
    setActiveCourseId(course.id);
    setSelectedTraineeId('');
    setEditingId(null);
    setForm(emptyTrainee);
    setImportReport(null);
  };

  const exportTrainees = () => {
    exportCsv('lista_szkolonych_pad.csv', [
      ['Stopień', 'Imię', 'Nazwisko', 'Telefon', 'Jednostka', 'Zespół', 'Grupa', 'Rola'],
      ...sortedTrainees(trainees).map((item) => [item.rank, item.firstName, item.lastName, item.phone, item.unit, item.team, item.group, item.role]),
    ]);
  };

  const exportIssues = () => {
    exportCsv('wydania_sprzetu_pad.csv', [
      ['CASE', 'Walizka', 'Zawartość', 'Dla kogo', 'Data wydania', 'Data zwrotu', 'Status', 'Uwagi'],
      ...caseLedger.map((item) => [
        item.caseNo,
        item.name,
        item.contents,
        assigneeLabel(item.issuedTo, trainees),
        item.issuedDate,
        item.returnDate,
        item.status,
        item.notes,
      ]),
    ]);
  };

  const createNewCourse = (event) => {
    event.preventDefault();
    const course = createCourse({
      name: courseForm.name || 'Kurs BSP',
      dateFrom: courseForm.dateFrom,
      dateTo: courseForm.dateTo,
      location: courseForm.location,
    });
    setCourses((previous) => [course, ...previous]);
    setActiveCourseId(course.id);
    setCourseForm({ ...emptyCourseForm, dateFrom: today, dateTo: today });
    setSearch('');
    setSelectedTraineeId('');
    setImportReport(null);
    setActiveView('start');
    setIsMenuOpen(false);
  };

  const updateCourseMeta = (courseId, patch) => {
    setCourses((previous) => previous.map((course) => (course.id === courseId ? { ...course, ...patch } : course)));
  };

  const switchCourse = (courseId) => {
    setActiveCourseId(courseId);
    setSelectedTraineeId('');
    setSearch('');
    setImportReport(null);
    setActiveView('start');
    setIsMenuOpen(false);
  };

  const deleteCourse = (courseId) => {
    const course = courses.find((item) => item.id === courseId);
    if (!course) return;
    if (!window.confirm(`Usunąć kurs "${course.name}"? Tej operacji nie da się cofnąć.`)) return;

    const remaining = courses.filter((item) => item.id !== courseId);
    if (!remaining.length) {
      const freshCourse = createCourse({ name: 'Kurs BSP', dateFrom: today, dateTo: today });
      setCourses([freshCourse]);
      setActiveCourseId(freshCourse.id);
    } else {
      setCourses(remaining);
      if (courseId === activeCourse.id) {
        setActiveCourseId(remaining[0].id);
      }
    }
    setSelectedTraineeId('');
    setSearch('');
    setImportReport(null);
    setActiveView('courses');
  };

  const goToView = (view) => {
    setActiveView(view);
    setIsMenuOpen(false);
  };

  const finishCourse = async () => {
    const endedAt = new Date().toISOString();
    updateActiveCourse(() => ({ status: 'zakończony', endedAt }));
    await exportCourseDocx({ ...activeCourse, status: 'zakończony', endedAt }, allKitItems, allPreCourseTasks);
  };

  const importTraineesFromExcel = async (file) => {
    if (!file) return;
    try {
      const result = parseTraineesWorkbook(await file.arrayBuffer(), trainees);
      if (result.added.length) {
        updateActiveCourse((course) => {
          const newRegister = result.added.reduce((register, trainee) => ({ ...register, [trainee.id]: {} }), {});
          return {
            trainees: sortedTrainees([...(course.trainees || []), ...result.added]),
            equipmentRegister: { ...(course.equipmentRegister || {}), ...newRegister },
          };
        });
      }
      setImportReport(result);
    } catch (error) {
      setImportReport({ added: [], skipped: [], errors: [`Nie udało się odczytać pliku: ${error.message}`] });
    }
  };

  return (
    <div className="app-shell">
      <aside className="sidebar">
        <div className="brand">
          <img className="brand-logo" src={assetPath('pad-logo-green.png')} alt="Polska Armia Dronów" />
          <div>
            <strong>Instruktor BSP</strong>
            <span>Kurs intensywny 3 dni</span>
          </div>
        </div>
        <button className="menu-toggle" type="button" onClick={() => setIsMenuOpen((open) => !open)}>
          <Menu size={18} />
          <span>Menu</span>
        </button>
        <nav className={isMenuOpen ? 'open' : ''}>
          {navItems.map(({ id, label, icon: Icon }) => (
            <button key={id} className={activeView === id ? 'active' : ''} type="button" onClick={() => goToView(id)}>
              <Icon size={18} />
              <span>{label}</span>
            </button>
          ))}
        </nav>
        <div className="sidebar-footer">
          <ActionButton icon={Printer} variant="ghost" onClick={() => window.print()}>
            Druk
          </ActionButton>
          <IconButton icon={RotateCcw} label="Wyczyść dane" variant="danger" onClick={resetAll} />
        </div>
      </aside>

      <main className="workspace">
        <header className="topbar">
          <div className="topbar-title">
            <img className="topbar-logo" src={assetPath('pad-logo-wide.jpeg')} alt="Polska Armia Dronów" />
            <h1>Panel instruktora</h1>
            <p>{activeCourse.name} / {activeCourse.dateFrom || '-'} - {activeCourse.dateTo || '-'} / {activeCourse.location || 'bez miejsca'}</p>
          </div>
          <div className="topbar-actions">
            {installPrompt && !isInstalled ? (
              <ActionButton icon={Download} variant="primary" onClick={installApp}>
                Zainstaluj
              </ActionButton>
            ) : null}
            <ActionButton icon={Archive} variant="secondary" onClick={() => exportCourseDocx(activeCourse, allKitItems, allPreCourseTasks)}>
              DOCX kursu
            </ActionButton>
            <ActionButton icon={Download} variant="secondary" onClick={exportTrainees}>
              Lista szkolonych
            </ActionButton>
            <ActionButton icon={FileSpreadsheet} variant="secondary" onClick={exportIssues}>
              Wydania
            </ActionButton>
            {activeCourse.status === 'zakończony' ? (
              <ActionButton icon={CheckCircle2} variant="secondary" onClick={() => exportCourseDocx(activeCourse, allKitItems, allPreCourseTasks)}>
                Kurs zakończony
              </ActionButton>
            ) : (
              <ActionButton icon={CheckCircle2} variant="primary" onClick={finishCourse}>
                Zakończ kurs
              </ActionButton>
            )}
          </div>
        </header>

        {activeView === 'start' && (
          <StartView
            doneTasks={doneTasks}
            totalTasks={flatTasks.length}
            preCourseDone={preCourseDone}
            totalPreCourse={allPreCourseTasks.length}
            packedItems={packedItems}
            totalKit={allKitItems.length}
            openIssues={openIssues}
            trainees={trainees}
            missingGear={missingGear}
            instructorsCount={participatingInstructors.length}
            setActiveView={setActiveView}
            schedule={schedule}
            progress={progress}
            carryState={carryState}
            caseLedger={caseLedger}
            allKitItems={allKitItems}
          />
        )}

        {activeView === 'courses' && (
          <CoursesView
            courses={courses}
            activeCourse={activeCourse}
            courseForm={courseForm}
            setCourseForm={setCourseForm}
            createNewCourse={createNewCourse}
            switchCourse={switchCourse}
            updateCourseMeta={updateCourseMeta}
            deleteCourse={deleteCourse}
            exportCourseDocx={(course) => exportCourseDocx(course, allKitItems, allPreCourseTasks)}
            openActiveCourse={() => goToView('start')}
          />
        )}

        {activeView === 'pre-course' && (
          <PreCourseView
            tasks={allPreCourseTasks}
            progress={activeCourse.preCourseProgress || {}}
            toggleTask={togglePreCourseTask}
            customTaskForm={customTaskForm}
            setCustomTaskForm={setCustomTaskForm}
            addCustomTask={addCustomPreCourseTask}
            deleteCustomTask={deleteCustomPreCourseTask}
          />
        )}

        {activeView === 'plan' && (
          <PlanView
            activeDay={activeDay}
            setActiveDay={setActiveDay}
            progress={progress}
            toggleTask={toggleTask}
            schedule={schedule}
          />
        )}

        {activeView === 'kit' && (
          <KitView
            filteredKit={filteredKit}
            kitFilter={kitFilter}
            setKitFilter={setKitFilter}
            search={search}
            setSearch={setSearch}
            carryState={carryState}
            updateCarryItem={updateCarryItem}
            customKitForm={customKitForm}
            setCustomKitForm={setCustomKitForm}
            addCustomKitItem={addCustomKitItem}
            deleteCustomKitItem={deleteCustomKitItem}
          />
        )}

        {activeView === 'trainee-kit' && (
          <TraineeKitView
            trainees={trainees}
            activeTrainee={activeTrainee}
            activeTraineeId={activeTraineeId}
            setSelectedTraineeId={setSelectedTraineeId}
            equipmentRegister={equipmentRegister}
            updateEquipment={updateEquipment}
            setActiveView={setActiveView}
          />
        )}

        {activeView === 'issues' && (
          <IssuesView
            search={search}
            setSearch={setSearch}
            filteredCases={filteredCases}
            updateCase={updateCase}
            trainees={trainees}
            exportIssues={exportIssues}
          />
        )}

        {activeView === 'trainees' && (
          <TraineesView
            form={form}
            setForm={setForm}
            saveTrainee={saveTrainee}
            editingId={editingId}
            setEditingId={setEditingId}
            trainees={filteredTrainees}
            allTrainees={trainees}
            search={search}
            setSearch={setSearch}
            editTrainee={editTrainee}
            deleteTrainee={deleteTrainee}
            exportTrainees={exportTrainees}
            importTraineesFromExcel={importTraineesFromExcel}
            importReport={importReport}
            traineeViewMode={traineeViewMode}
            setTraineeViewMode={setTraineeViewMode}
          />
        )}

        {activeView === 'instructors' && (
          <InstructorsView
            instructorProfile={instructorProfile}
            setInstructorProfile={setInstructorProfile}
            saveInstructorProfile={saveInstructorProfile}
            markProfileParticipating={markProfileParticipating}
            profileCourseEntry={profileCourseEntry}
            instructors={activeCourse.instructors || []}
            instructorForm={instructorForm}
            setInstructorForm={setInstructorForm}
            addInstructorToCourse={addInstructorToCourse}
            updateInstructorInCourse={updateInstructorInCourse}
            deleteInstructorFromCourse={deleteInstructorFromCourse}
            updates={activeCourse.updates || []}
            updateForm={updateForm}
            setUpdateForm={setUpdateForm}
            addInstructorUpdate={addInstructorUpdate}
            deleteInstructorUpdate={deleteInstructorUpdate}
          />
        )}
      </main>
    </div>
  );
}

function CoursesView({ courses, activeCourse, courseForm, setCourseForm, createNewCourse, switchCourse, updateCourseMeta, deleteCourse, exportCourseDocx, openActiveCourse }) {
  const updateForm = (field, value) => setCourseForm((previous) => ({ ...previous, [field]: value }));

  return (
    <section className="courses-layout">
      <div className="panel form-panel">
        <SectionHeader kicker="Po uruchomieniu" title="Aktywny kurs" />
        <div className="active-course-card">
          <strong>{activeCourse.name}</strong>
          <span>{activeCourse.dateFrom || '-'} - {activeCourse.dateTo || '-'} / {activeCourse.location || 'bez miejsca'}</span>
          <StatusPill value={activeCourse.status || 'aktywny'} />
          <div className="row-actions">
            <ActionButton icon={FolderOpen} variant="primary" onClick={openActiveCourse}>
              Przejdź do kursu
            </ActionButton>
            <ActionButton icon={Archive} variant="secondary" onClick={() => exportCourseDocx(activeCourse)}>
              Eksport DOCX
            </ActionButton>
          </div>
        </div>
      </div>

      <div className="panel form-panel">
        <SectionHeader kicker="Nowy cykl" title="Rozpocznij nowy kurs" />
        <form className="trainee-form" onSubmit={createNewCourse}>
          <label className="wide">
            <span>Nazwa kursu</span>
            <input value={courseForm.name} onChange={(event) => updateForm('name', event.target.value)} placeholder="np. Kurs BSP Ociesęki" required />
          </label>
          <label>
            <span>Data od</span>
            <input type="date" value={courseForm.dateFrom} onChange={(event) => updateForm('dateFrom', event.target.value)} required />
          </label>
          <label>
            <span>Data do</span>
            <input type="date" value={courseForm.dateTo} onChange={(event) => updateForm('dateTo', event.target.value)} required />
          </label>
          <label className="wide">
            <span>Miejsce</span>
            <input value={courseForm.location} onChange={(event) => updateForm('location', event.target.value)} placeholder="np. Ociesęki / Nowa Dęba" />
          </label>
          <div className="form-actions">
            <button className="submit-button" type="submit">
              <Plus size={17} />
              <span>Utwórz kurs</span>
            </button>
          </div>
        </form>
      </div>

      <div className="panel table-panel wide-panel">
        <SectionHeader kicker={`${courses.length} zapisanych`} title="Aktywny kurs i archiwum" />
        <div className="course-list">
          {courses.map((course) => (
            <div className={`course-row ${course.id === activeCourse.id ? 'active' : ''}`} key={course.id}>
              <div className="course-meta">
                <strong>{course.name}</strong>
                <span>{course.dateFrom || '-'} - {course.dateTo || '-'} / {course.location || 'bez miejsca'}</span>
                <small>{course.status || 'aktywny'} / {(course.trainees || []).length} szkolonych / {(course.caseLedger || []).filter((item) => item.status === 'wydane').length} aktywnych wydań</small>
                {course.endedAt ? <small>Zakończono: {formatDateTime(course.endedAt)}</small> : null}
              </div>
              <div className="course-edit-grid">
                <input value={course.name} onChange={(event) => updateCourseMeta(course.id, { name: event.target.value })} />
                <input type="date" value={course.dateFrom || ''} onChange={(event) => updateCourseMeta(course.id, { dateFrom: event.target.value })} />
                <input type="date" value={course.dateTo || ''} onChange={(event) => updateCourseMeta(course.id, { dateTo: event.target.value })} />
                <input value={course.location || ''} onChange={(event) => updateCourseMeta(course.id, { location: event.target.value })} placeholder="Miejsce" />
              </div>
              <div className="row-actions">
                <ActionButton icon={FolderOpen} variant={course.id === activeCourse.id ? 'primary' : 'secondary'} onClick={() => switchCourse(course.id)}>
                  {course.id === activeCourse.id ? 'Aktywny' : 'Otwórz'}
                </ActionButton>
                <IconButton icon={Archive} label="Zapisz DOCX" onClick={() => exportCourseDocx(course)} />
                <IconButton icon={Trash2} label="Usuń kurs" variant="danger" onClick={() => deleteCourse(course.id)} />
              </div>
            </div>
          ))}
        </div>
      </div>
    </section>
  );
}

function StartView({ doneTasks, totalTasks, preCourseDone, totalPreCourse, packedItems, totalKit, openIssues, trainees, missingGear, instructorsCount, setActiveView, schedule, progress, carryState, caseLedger, allKitItems }) {
  const nextTasks = schedule
    .flatMap((group) => group.tasks.map((task) => ({ ...task, day: group.day })))
    .filter((task) => !isDone(progress[task.id]))
    .slice(0, 5);
  const notPacked = allKitItems.filter((item) => !carryState[item.id]?.packed).slice(0, 6);
  const issued = caseLedger.filter((item) => item.status === 'wydane').slice(0, 5);

  return (
    <section className="view-stack">
      <div className="dashboard-grid">
        <MetricCard
          icon={ListChecks}
          label="Wykonano"
          value={`${doneTasks}/${totalTasks}`}
          sublabel="pozycji planu"
          tone="olive"
        />
        <MetricCard icon={CheckCircle2} label="Przed kursem" value={`${preCourseDone}/${totalPreCourse}`} sublabel="zadań przygot." tone="olive" />
        <MetricCard icon={PackageCheck} label="Zabrane" value={`${packedItems}/${totalKit}`} sublabel="pozycji sprzętu" tone="amber" />
        <MetricCard icon={Users} label="Szkoleni" value={trainees.length} sublabel="osób na liście" tone="steel" />
        <MetricCard icon={UserCheck} label="Instruktorzy" value={instructorsCount} sublabel="bierze udział" tone="steel" />
      </div>

      <div className="overview-layout">
        <div className="panel progress-panel">
          <SectionHeader kicker="Stan kursu" title="Postęp wykonania">
            <ActionButton icon={ClipboardList} variant="secondary" onClick={() => setActiveView('plan')}>
              Plan
            </ActionButton>
            <ActionButton icon={CheckCircle2} variant="secondary" onClick={() => setActiveView('pre-course')}>
              Przed kursem
            </ActionButton>
          </SectionHeader>
          <div className="dial-row">
            <ProgressDial value={percent(doneTasks, totalTasks)} label="plan" />
            <ProgressDial value={percent(packedItems, totalKit)} label="sprzęt" />
          </div>
          <div className="quick-list">
            {nextTasks.map((task) => (
              <div className="quick-row" key={task.id}>
                <span>{task.day}</span>
                <strong>{task.time}</strong>
                <p>{task.title}</p>
              </div>
            ))}
          </div>
        </div>

        <div className="panel">
          <SectionHeader kicker="Logistyka" title="Jeszcze do zabrania">
            <ActionButton icon={PackageCheck} variant="secondary" onClick={() => setActiveView('kit')}>
              Lista
            </ActionButton>
          </SectionHeader>
          <div className="compact-list">
            {notPacked.map((item) => (
              <div key={item.id}>
                <span>{item.category}</span>
                <strong>{item.name}</strong>
                <small>{item.qty}</small>
              </div>
            ))}
          </div>
        </div>

        <div className="panel">
          <SectionHeader kicker="Ewidencja" title="Wydane CASE">
            <ActionButton icon={Boxes} variant="secondary" onClick={() => setActiveView('issues')}>
              Wydania
            </ActionButton>
          </SectionHeader>
          <div className="compact-list">
            {issued.length ? (
              issued.map((item) => (
                <div key={item.id}>
                  <span>{item.caseNo}</span>
                  <strong>{item.name}</strong>
                  <small>{item.issuedDate || 'bez daty'}</small>
                </div>
              ))
            ) : (
              <div className="empty-line">Brak aktywnych wydań.</div>
            )}
          </div>
        </div>
      </div>
    </section>
  );
}

function PreCourseView({ tasks, progress, toggleTask, customTaskForm, setCustomTaskForm, addCustomTask, deleteCustomTask }) {
  const updateForm = (field, value) => setCustomTaskForm((previous) => ({ ...previous, [field]: value }));
  const done = tasks.filter((task) => isDone(progress[task.id])).length;

  return (
    <section className="prep-layout">
      <div className="panel form-panel">
        <SectionHeader kicker="Szablon na kolejne kursy" title="Dodaj własne zadanie" />
        <form className="trainee-form" onSubmit={addCustomTask}>
          <label>
            <span>Termin</span>
            <input value={customTaskForm.phase} onChange={(event) => updateForm('phase', event.target.value)} placeholder="np. T-7" />
          </label>
          <label>
            <span>Nazwa</span>
            <input value={customTaskForm.title} onChange={(event) => updateForm('title', event.target.value)} placeholder="Zadanie przygotowawcze" required />
          </label>
          <label className="wide">
            <span>Opis</span>
            <input value={customTaskForm.description} onChange={(event) => updateForm('description', event.target.value)} placeholder="Krótka instrukcja wykonania" />
          </label>
          <div className="form-actions">
            <button className="submit-button" type="submit">
              <Plus size={17} />
              <span>Dodaj zadanie</span>
            </button>
          </div>
        </form>
      </div>

      <div className="panel">
        <SectionHeader kicker={`${done}/${tasks.length} wykonane`} title="Do zrobienia przed kursem" />
        <div className="timeline-list">
          {tasks.map((task) => (
            <label className={`timeline-item prep-item ${isDone(progress[task.id]) ? 'done' : ''}`} key={task.id}>
              <input type="checkbox" checked={isDone(progress[task.id])} onChange={() => toggleTask(task.id)} />
              <span className="time">{task.phase}</span>
              <div>
                <strong>{task.title}</strong>
                <p>{task.description || 'Bez opisu.'}</p>
                {isDone(progress[task.id]) ? <small>Wykonano: {formatDateTime(completedAt(progress[task.id]))}</small> : null}
                <small>{task.source === 'custom' ? 'Własne - będzie dostępne w kolejnych kursach' : 'Stałe z checklisty instruktora'}</small>
              </div>
              {task.source === 'custom' ? (
                <IconButton icon={Trash2} label="Usuń zadanie z szablonu" variant="danger" onClick={(event) => {
                  event.preventDefault();
                  deleteCustomTask(task.id);
                }} />
              ) : null}
            </label>
          ))}
        </div>
      </div>
    </section>
  );
}

function PlanView({ activeDay, setActiveDay, progress, toggleTask, schedule }) {
  const activeGroup = schedule.find((group) => group.day === activeDay);

  return (
    <section className="view-stack">
      <SectionHeader kicker="Co wykonano" title="Plan szkolenia D1-D3">
        <div className="segmented">
          {schedule.map((group) => (
            <button key={group.day} type="button" className={activeDay === group.day ? 'selected' : ''} onClick={() => setActiveDay(group.day)}>
              {group.day}
            </button>
          ))}
        </div>
      </SectionHeader>

      <div className="panel">
        <div className="table-title">
          <div>
            <strong>{activeGroup.title}</strong>
            <span>{activeGroup.tasks.filter((task) => isDone(progress[task.id])).length}/{activeGroup.tasks.length} wykonane</span>
          </div>
          <ProgressDial value={percent(activeGroup.tasks.filter((task) => isDone(progress[task.id])).length, activeGroup.tasks.length)} label={activeDay} />
        </div>
        <div className="timeline-list">
          {activeGroup.tasks.map((task) => (
            <label className={`timeline-item ${isDone(progress[task.id]) ? 'done' : ''}`} key={task.id}>
              <input type="checkbox" checked={isDone(progress[task.id])} onChange={() => toggleTask(task.id)} />
              <span className="time">{task.time}</span>
              <div>
                <strong>{task.title}</strong>
                <p>{task.action}</p>
                {isDone(progress[task.id]) ? <small>Wykonano: {formatDateTime(completedAt(progress[task.id]))}</small> : null}
                <small>{task.metric}</small>
              </div>
            </label>
          ))}
        </div>
      </div>
    </section>
  );
}

function KitView({ filteredKit, kitFilter, setKitFilter, search, setSearch, carryState, updateCarryItem, customKitForm, setCustomKitForm, addCustomKitItem, deleteCustomKitItem }) {
  const updateForm = (field, value) => setCustomKitForm((previous) => ({ ...previous, [field]: value }));

  return (
    <section className="view-stack">
      <SectionHeader kicker="Logistyka przed szkoleniem" title="Sprzęt do zabrania">
        <SearchBox value={search} onChange={setSearch} placeholder="Szukaj sprzętu" />
      </SectionHeader>

      <div className="panel">
        <SectionHeader kicker="Szablon na kolejne kursy" title="Dodaj ręcznie sprzęt" />
        <form className="inline-form" onSubmit={addCustomKitItem}>
          <label>
            <span>Kategoria</span>
            <input value={customKitForm.category} onChange={(event) => updateForm('category', event.target.value)} placeholder="np. Medyczne" />
          </label>
          <label className="wide">
            <span>Element</span>
            <input value={customKitForm.name} onChange={(event) => updateForm('name', event.target.value)} placeholder="Nazwa sprzętu" required />
          </label>
          <label>
            <span>Ilość</span>
            <input value={customKitForm.qty} onChange={(event) => updateForm('qty', event.target.value)} placeholder="1" />
          </label>
          <button className="submit-button" type="submit">
            <Plus size={17} />
            <span>Dodaj sprzęt</span>
          </button>
        </form>
      </div>

      <div className="toolbar">
        {['wszystko', 'do zabrania', 'zabrane', 'braki'].map((filter) => (
          <button className={kitFilter === filter ? 'selected' : ''} type="button" key={filter} onClick={() => setKitFilter(filter)}>
            {filter}
          </button>
        ))}
      </div>

      <div className="panel table-panel">
        <table>
          <thead>
            <tr>
              <th className="check-col">OK</th>
              <th>Kategoria</th>
              <th>Element</th>
              <th>Ilość</th>
              <th>Status</th>
              <th>Uwagi</th>
              <th>Szablon</th>
            </tr>
          </thead>
          <tbody>
            {filteredKit.map((item) => {
              const state = carryState[item.id] || { packed: false, status: 'do zabrania', note: '' };
              return (
                <tr key={item.id}>
                  <td>
                    <input
                      type="checkbox"
                      checked={state.packed}
                      onChange={(event) =>
                        updateCarryItem(item.id, {
                          packed: event.target.checked,
                          status: event.target.checked ? 'zabrane' : 'do zabrania',
                        })
                      }
                    />
                  </td>
                  <td><span className="category-label">{item.category}</span></td>
                  <td><strong>{item.name}</strong></td>
                  <td>{item.qty}</td>
                  <td>
                    <select value={state.status} onChange={(event) => updateCarryItem(item.id, { status: event.target.value })}>
                      {['do zabrania', 'zabrane', 'brak', 'sprawdzić'].map((status) => (
                        <option key={status}>{status}</option>
                      ))}
                    </select>
                  </td>
                  <td>
                    <input value={state.note} onChange={(event) => updateCarryItem(item.id, { note: event.target.value })} placeholder="-" />
                  </td>
                  {item.source === 'custom' ? (
                    <td className="row-actions">
                      <IconButton icon={Trash2} label="Usuń sprzęt z szablonu" variant="danger" onClick={() => deleteCustomKitItem(item.id)} />
                    </td>
                  ) : (
                    <td><span className="muted-cell">stałe</span></td>
                  )}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </section>
  );
}

function TraineeKitView({ trainees, activeTrainee, activeTraineeId, setSelectedTraineeId, equipmentRegister, updateEquipment, setActiveView }) {
  const grouped = traineeKit.reduce((groups, item) => {
    groups[item.category] = [...(groups[item.category] || []), item];
    return groups;
  }, {});

  if (!trainees.length) {
    return (
      <section className="view-stack">
        <SectionHeader kicker="Sprzęt szkolonych" title="Kontrola braków">
          <ActionButton icon={UserPlus} variant="primary" onClick={() => setActiveView('trainees')}>
            Dodaj szkolonego
          </ActionButton>
        </SectionHeader>
        <div className="empty-state">
          <ShieldCheck size={28} />
          <strong>Lista szkolonych jest pusta.</strong>
          <span>Po dodaniu osób pojawi się matryca sprzętu własnego i braków do wypożyczenia.</span>
        </div>
      </section>
    );
  }

  const register = equipmentRegister[activeTraineeId] || {};
  const missing = traineeKit.filter((item) => ['brak', 'do wypożyczenia'].includes(register[item.id])).length;

  return (
    <section className="view-stack">
      <SectionHeader kicker="Sprzęt szkolonych" title="Matryca wyposażenia kursanta">
        <select className="person-select" value={activeTraineeId} onChange={(event) => setSelectedTraineeId(event.target.value)}>
          {trainees.map((trainee) => (
            <option key={trainee.id} value={trainee.id}>
              {fullName(trainee)} / {trainee.team}
            </option>
          ))}
        </select>
      </SectionHeader>

      <div className="equipment-summary">
        <MetricCard icon={Users} label="Szkolony" value={activeTrainee ? fullName(activeTrainee) : '-'} sublabel={activeTrainee?.unit || 'bez jednostki'} tone="steel" />
        <MetricCard icon={AlertTriangle} label="Braki" value={missing} sublabel="pozycji do decyzji" tone={missing ? 'red' : 'olive'} />
        <MetricCard icon={Radio} label="Zespół" value={activeTrainee?.team || '-'} sublabel={activeTrainee?.group || '-'} tone="amber" />
      </div>

      <div className="kit-grid">
        {Object.entries(grouped).map(([category, items]) => (
          <div className="panel kit-group" key={category}>
            <h3>{category}</h3>
            {items.map((item) => {
              const value = register[item.id] || 'nie sprawdzono';
              return (
                <div className="kit-check-row" key={item.id}>
                  <div>
                    <strong>{item.name}</strong>
                    <StatusPill value={value} />
                  </div>
                  <select value={value} onChange={(event) => updateEquipment(activeTraineeId, item.id, event.target.value)}>
                    {equipmentStatuses.map((status) => (
                      <option key={status}>{status}</option>
                    ))}
                  </select>
                </div>
              );
            })}
          </div>
        ))}
      </div>
    </section>
  );
}

function IssuesView({ search, setSearch, filteredCases, updateCase, trainees, exportIssues }) {
  return (
    <section className="view-stack">
      <SectionHeader kicker="Odpowiedzialność materialna" title="Wydanie i zwrot walizek CASE">
        <div className="section-actions">
          <SearchBox value={search} onChange={setSearch} placeholder="Szukaj CASE" />
          <ActionButton icon={Download} variant="secondary" onClick={exportIssues}>
            CSV
          </ActionButton>
        </div>
      </SectionHeader>

      <div className="panel table-panel">
        <table>
          <thead>
            <tr>
              <th>CASE</th>
              <th>Walizka</th>
              <th>Dla kogo</th>
              <th>Wydano</th>
              <th>Zwrot</th>
              <th>Status</th>
              <th>Uwagi</th>
            </tr>
          </thead>
          <tbody>
            {filteredCases.map((item) => (
              <tr key={item.id}>
                <td><strong>{item.caseNo}</strong></td>
                <td>
                  <strong>{item.name}</strong>
                  <small className="muted-cell">{item.contents}</small>
                </td>
                <td>
                  <select value={item.issuedTo} onChange={(event) => updateCase(item.id, { issuedTo: event.target.value, status: event.target.value ? 'wydane' : 'magazyn', issuedDate: event.target.value ? item.issuedDate || today : '' })}>
                    <option value="">Nie wydano</option>
                    {teams.map((team) => (
                      <option key={team} value={`team:${team}`}>{team}</option>
                    ))}
                    {trainees.map((trainee) => (
                      <option key={trainee.id} value={`trainee:${trainee.id}`}>{fullName(trainee)}</option>
                    ))}
                  </select>
                </td>
                <td><input type="date" value={item.issuedDate} onChange={(event) => updateCase(item.id, { issuedDate: event.target.value })} /></td>
                <td><input type="date" value={item.returnDate} onChange={(event) => updateCase(item.id, { returnDate: event.target.value, status: event.target.value ? 'zwrócone' : item.status })} /></td>
                <td>
                  <select value={item.status} onChange={(event) => updateCase(item.id, { status: event.target.value })}>
                    {issueStatuses.map((status) => (
                      <option key={status}>{status}</option>
                    ))}
                  </select>
                </td>
                <td><input value={item.notes} onChange={(event) => updateCase(item.id, { notes: event.target.value })} placeholder="-" /></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
  );
}

function TraineesView({
  form,
  setForm,
  saveTrainee,
  editingId,
  setEditingId,
  trainees,
  allTrainees,
  search,
  setSearch,
  editTrainee,
  deleteTrainee,
  exportTrainees,
  importTraineesFromExcel,
  importReport,
  traineeViewMode,
  setTraineeViewMode,
}) {
  const updateForm = (field, value) => setForm((previous) => ({ ...previous, [field]: value }));
  const grouped = groupTrainees(trainees);

  return (
    <section className="trainees-layout">
      <div className="panel form-panel">
        <SectionHeader kicker="Lista szkolonych" title={editingId ? 'Edycja osoby' : 'Nowy szkolony'} />
        <form onSubmit={saveTrainee} className="trainee-form">
          <label>
            <span>Stopień</span>
            <input value={form.rank} onChange={(event) => updateForm('rank', event.target.value)} placeholder="np. st. szer." />
          </label>
          <label>
            <span>Imię</span>
            <input value={form.firstName} onChange={(event) => updateForm('firstName', event.target.value)} required />
          </label>
          <label>
            <span>Nazwisko</span>
            <input value={form.lastName} onChange={(event) => updateForm('lastName', event.target.value)} required />
          </label>
          <label>
            <span>Telefon</span>
            <input value={form.phone} onChange={(event) => updateForm('phone', event.target.value)} inputMode="tel" placeholder="+48 ..." />
          </label>
          <label className="wide">
            <span>Jednostka</span>
            <input value={form.unit} onChange={(event) => updateForm('unit', event.target.value)} placeholder="Jednostka / organizacja" />
          </label>
          <label>
            <span>Zespół</span>
            <select value={form.team} onChange={(event) => updateForm('team', event.target.value)}>
              {teams.map((team) => <option key={team}>{team}</option>)}
            </select>
          </label>
          <label>
            <span>Grupa</span>
            <select value={form.group} onChange={(event) => updateForm('group', event.target.value)}>
              <option>BLUE</option>
              <option>RED</option>
            </select>
          </label>
          <label>
            <span>Rola</span>
            <select value={form.role} onChange={(event) => updateForm('role', event.target.value)}>
              <option>Pilot</option>
              <option>Nawigator</option>
              <option>Obserwator</option>
              <option>Dowódca</option>
            </select>
          </label>
          <div className="form-actions">
            {editingId ? (
              <ActionButton
                icon={X}
                variant="ghost"
                onClick={() => {
                  setEditingId(null);
                  setForm(emptyTrainee);
                }}
              >
                Anuluj
              </ActionButton>
            ) : null}
            <button className="submit-button" type="submit">
              <Plus size={17} />
              <span>{editingId ? 'Zapisz' : 'Dodaj'}</span>
            </button>
          </div>
        </form>
      </div>

      <div className="panel table-panel">
        <SectionHeader kicker={`${allTrainees.length} osób`} title="Ewidencja szkolonych">
          <div className="section-actions">
            <SearchBox value={search} onChange={setSearch} placeholder="Szukaj osoby" />
            <label className="file-button">
              <Upload size={17} />
              <span>Excel</span>
              <input
                type="file"
                accept=".xlsx,.xls"
                onChange={(event) => {
                  importTraineesFromExcel(event.target.files?.[0]);
                  event.target.value = '';
                }}
              />
            </label>
            <ActionButton icon={Download} variant="secondary" onClick={exportTrainees}>
              CSV
            </ActionButton>
          </div>
        </SectionHeader>

        <div className="toolbar">
          <button className={traineeViewMode === 'table' ? 'selected' : ''} type="button" onClick={() => setTraineeViewMode('table')}>
            Tabela
          </button>
          <button className={traineeViewMode === 'groups' ? 'selected' : ''} type="button" onClick={() => setTraineeViewMode('groups')}>
            Według grup
          </button>
        </div>

        {importReport ? (
          <div className={`import-report ${importReport.errors.length ? 'warning' : 'success'}`}>
            <strong>Import Excel: dodano {importReport.added.length}, pominięto duplikaty {importReport.skipped.length}, błędy {importReport.errors.length}</strong>
            {importReport.errors.length ? <span>{importReport.errors.slice(0, 3).join(' | ')}</span> : null}
            {importReport.skipped.length ? <span>Pominięto: {importReport.skipped.slice(0, 5).join(', ')}</span> : null}
          </div>
        ) : null}

        {traineeViewMode === 'table' ? (
          <table>
            <thead>
              <tr>
                <th>Stopień / imię i nazwisko</th>
                <th>Telefon</th>
                <th>Jednostka</th>
                <th>Zespół</th>
                <th>Grupa</th>
                <th>Rola</th>
                <th></th>
              </tr>
            </thead>
            <tbody>
              {trainees.map((trainee) => (
                <tr key={trainee.id}>
                  <td><strong>{fullName(trainee)}</strong></td>
                  <td>{trainee.phone || '-'}</td>
                  <td>{trainee.unit || '-'}</td>
                  <td>{trainee.team}</td>
                  <td><StatusPill value={trainee.group} /></td>
                  <td>{trainee.role}</td>
                  <td className="row-actions">
                    <IconButton icon={Edit3} label="Edytuj" onClick={() => editTrainee(trainee)} />
                    <IconButton icon={Trash2} label="Usuń" variant="danger" onClick={() => deleteTrainee(trainee.id)} />
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        ) : (
          <div className="grouped-trainees">
            {Object.entries(grouped).map(([group, teamsMap]) => (
              <div className="group-section" key={group}>
                <h3>{group}</h3>
                {Object.entries(teamsMap).map(([team, items]) => (
                  <div className="team-section" key={team}>
                    <h4>{team} / {items.length} osób</h4>
                    <div className="team-list">
                      {items.map((trainee) => (
                        <div className="team-person" key={trainee.id}>
                          <div>
                            <strong>{fullName(trainee)}</strong>
                            <span>{trainee.role} / {trainee.unit || 'bez jednostki'} / {trainee.phone || 'bez telefonu'}</span>
                          </div>
                          <div className="row-actions">
                            <IconButton icon={Edit3} label="Edytuj" onClick={() => editTrainee(trainee)} />
                            <IconButton icon={Trash2} label="Usuń" variant="danger" onClick={() => deleteTrainee(trainee.id)} />
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                ))}
              </div>
            ))}
          </div>
        )}
        {!trainees.length && <div className="empty-line">Brak osób dla bieżącego wyszukiwania.</div>}
      </div>
    </section>
  );
}

function InstructorsView({
  instructorProfile,
  setInstructorProfile,
  saveInstructorProfile,
  markProfileParticipating,
  profileCourseEntry,
  instructors,
  instructorForm,
  setInstructorForm,
  addInstructorToCourse,
  updateInstructorInCourse,
  deleteInstructorFromCourse,
  updates,
  updateForm,
  setUpdateForm,
  addInstructorUpdate,
  deleteInstructorUpdate,
}) {
  const updateProfile = (field, value) => setInstructorProfile((previous) => ({ ...previous, [field]: value }));
  const updateInstructorForm = (field, value) => setInstructorForm((previous) => ({ ...previous, [field]: value }));
  const updateInfoForm = (field, value) => setUpdateForm((previous) => ({ ...previous, [field]: value }));
  const participating = instructors.filter((item) => item.participating);

  return (
    <section className="instructors-layout">
      <div className="panel form-panel">
        <SectionHeader kicker="Twój profil" title="Profil instruktora" />
        <form className="trainee-form" onSubmit={saveInstructorProfile}>
          <label>
            <span>Stopień</span>
            <input value={instructorProfile.rank} onChange={(event) => updateProfile('rank', event.target.value)} placeholder="np. kpt." />
          </label>
          <label>
            <span>Imię</span>
            <input value={instructorProfile.firstName} onChange={(event) => updateProfile('firstName', event.target.value)} required />
          </label>
          <label>
            <span>Nazwisko</span>
            <input value={instructorProfile.lastName} onChange={(event) => updateProfile('lastName', event.target.value)} required />
          </label>
          <label>
            <span>Telefon</span>
            <input value={instructorProfile.phone} onChange={(event) => updateProfile('phone', event.target.value)} inputMode="tel" />
          </label>
          <label className="wide">
            <span>Jednostka / organizacja</span>
            <input value={instructorProfile.unit} onChange={(event) => updateProfile('unit', event.target.value)} />
          </label>
          <div className="form-actions">
            <ActionButton icon={Save} variant="secondary" type="submit">
              Zapisz profil
            </ActionButton>
            <ActionButton icon={UserCheck} variant={profileCourseEntry?.participating ? 'primary' : 'secondary'} onClick={markProfileParticipating}>
              {profileCourseEntry?.participating ? 'Biorę udział' : 'Biorę udział w kursie'}
            </ActionButton>
          </div>
        </form>
      </div>

      <div className="panel form-panel">
        <SectionHeader kicker="Kadra kursu" title="Dodaj instruktora do szkolenia" />
        <form className="trainee-form" onSubmit={addInstructorToCourse}>
          <label>
            <span>Stopień</span>
            <input value={instructorForm.rank} onChange={(event) => updateInstructorForm('rank', event.target.value)} />
          </label>
          <label>
            <span>Imię</span>
            <input value={instructorForm.firstName} onChange={(event) => updateInstructorForm('firstName', event.target.value)} required />
          </label>
          <label>
            <span>Nazwisko</span>
            <input value={instructorForm.lastName} onChange={(event) => updateInstructorForm('lastName', event.target.value)} required />
          </label>
          <label>
            <span>Telefon</span>
            <input value={instructorForm.phone} onChange={(event) => updateInstructorForm('phone', event.target.value)} />
          </label>
          <label>
            <span>Rola</span>
            <input value={instructorForm.role} onChange={(event) => updateInstructorForm('role', event.target.value)} />
          </label>
          <label>
            <span>Jednostka</span>
            <input value={instructorForm.unit} onChange={(event) => updateInstructorForm('unit', event.target.value)} />
          </label>
          <label className="wide checkbox-label">
            <input
              type="checkbox"
              checked={instructorForm.participating}
              onChange={(event) => updateInstructorForm('participating', event.target.checked)}
            />
            <span>Bierze udział w tym szkoleniu</span>
          </label>
          <div className="form-actions">
            <button className="submit-button" type="submit">
              <Plus size={17} />
              <span>Dodaj instruktora</span>
            </button>
          </div>
        </form>
      </div>

      <div className="panel table-panel wide-panel">
        <SectionHeader kicker={`${participating.length}/${instructors.length} bierze udział`} title="Instruktorzy przypisani do kursu" />
        <div className="instructor-list">
          {instructors.map((item) => (
            <div className={`instructor-row ${item.participating ? 'active' : ''}`} key={item.id}>
              <div>
                <strong>{fullName(item)}</strong>
                <span>{item.role || 'Instruktor'} / {item.unit || 'bez jednostki'} / {item.phone || 'bez telefonu'}</span>
              </div>
              <label className="mini-toggle">
                <input
                  type="checkbox"
                  checked={Boolean(item.participating)}
                  onChange={(event) => updateInstructorInCourse(item.id, { participating: event.target.checked })}
                />
                <span>Bierze udział</span>
              </label>
              <IconButton icon={Trash2} label="Usuń instruktora" variant="danger" onClick={() => deleteInstructorFromCourse(item.id)} />
            </div>
          ))}
          {!instructors.length ? <div className="empty-line">Brak instruktorów przypisanych do kursu.</div> : null}
        </div>
      </div>

      <div className="panel wide-panel">
        <SectionHeader kicker="Wymiana informacji" title="Postępy i uwagi instruktorów">
          <StatusPill value={profileCourseEntry?.participating ? 'aktywny instruktor' : 'najpierw oznacz udział'} />
        </SectionHeader>
        <form className="update-form" onSubmit={addInstructorUpdate}>
          <select value={updateForm.category} onChange={(event) => updateInfoForm('category', event.target.value)}>
            <option>Postęp szkolenia</option>
            <option>Problem</option>
            <option>Sprzęt</option>
            <option>Bezpieczeństwo</option>
            <option>Decyzja</option>
          </select>
          <input value={updateForm.message} onChange={(event) => updateInfoForm('message', event.target.value)} placeholder="Wpisz informację dla instruktorów biorących udział w szkoleniu" />
          <button className="submit-button" type="submit">
            <MessageSquare size={17} />
            <span>Dodaj wpis</span>
          </button>
        </form>
        <div className="updates-list">
          {updates.map((item) => (
            <div className="update-row" key={item.id}>
              <div>
                <span>{new Date(item.createdAt).toLocaleString('pl-PL')} / {item.category}</span>
                <strong>{item.authorName}</strong>
                <p>{item.message}</p>
              </div>
              <IconButton icon={Trash2} label="Usuń wpis" variant="danger" onClick={() => deleteInstructorUpdate(item.id)} />
            </div>
          ))}
          {!updates.length ? <div className="empty-line">Brak wpisów. Dodaj pierwszy meldunek o postępach, problemach albo sprzęcie.</div> : null}
        </div>
      </div>
    </section>
  );
}

export default App;
