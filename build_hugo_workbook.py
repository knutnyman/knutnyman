#!/usr/bin/env python3
"""Build an Excel summary of every Medtronic Hugo RAS placement found via deep research.

Data compiled 2026-06-05 from regional web-research agents (US, Western Europe,
Eastern Europe/Middle East/Africa, Latin America, Asia-Pacific) plus a global
milestones/regulatory backbone. Sources are Medtronic press releases, hospital
newsrooms, local news, and peer-reviewed clinical reports. See the 'Notes &
Caveats' sheet for confidence and methodology.
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Master placement data
# Columns: Institution, City, State/Province, Country, Region, Date, Specialty,
#          Type, Notable First, Confidence, Source URL(s), Notes
# ---------------------------------------------------------------------------
PLACEMENTS = [
    # ---------------- UNITED STATES ----------------
    ["Cleveland Clinic (Glickman Urologic Institute)", "Cleveland", "OH", "USA", "United States",
     "17 Feb 2026 (commercial); 2023–2025 (Expand URO trial)", "Urology (prostatectomy)",
     "Commercial + Clinical Trial", "First U.S. commercial Hugo surgery", "High",
     "https://news.medtronic.com/2026-02-17-Medtronic-announces-first-surgery-with-Hugo-TM-robotic-assisted-surgery-system-in-the-U-S-performed-at-Cleveland-Clinic ; https://newsroom.clevelandclinic.org/2026/02/17/cleveland-clinic-first-hospital-in-the-us-to-perform-robotic-assisted-prostate-surgery-using-newly-fda-cleared-system",
     "Prostatectomy by Dr. Jihad Kaouk; patient discharged next day. Also Expand URO IDE trial site."],
    ["Duke University Hospital", "Durham", "NC", "USA", "United States",
     "15 Dec 2022 (urology trial); later hernia trial", "Urology (prostatectomy); Hernia",
     "Clinical Trial", "First U.S. Expand URO patient", "High",
     "https://news.medtronic.com/2022-12-15-Medtronic-announces-first-patient-enrolled-in-U-S-clinical-trial-for-Hugo-TM-robotic-assisted-surgery-system ; https://www.prnewswire.com/news-releases/study-of-medtronic-hugo-robotic-assisted-surgery-system-in-hernia-repair-meets-safety-and-effectiveness-endpoints-302546743.html",
     "First Expand URO patient (Dr. Michael R. Abern, national PI). Also first Enable Hernia Repair case (Dr. Jacob Greenberg)."],
    ["Swedish Medical Center", "Seattle", "WA", "USA", "United States",
     "2022–2025 (Expand URO trial)", "Urology (prostatectomy, nephrectomy, cystectomy)",
     "Clinical Trial", "", "High",
     "https://www.prnewswire.com/news-releases/medtronic-announces-first-patient-enrolled-in-us-clinical-trial-for-hugo-robotic-assisted-surgery-system-301703533.html",
     "Dr. James Porter, Expand URO PI (later Medtronic CMO for robotics)."],
    ["Atrium Health Wake Forest Baptist – High Point Medical Center", "High Point", "NC", "USA", "United States",
     "~23 Feb 2026", "Urology (pyeloplasty, prostatectomy)",
     "Commercial", "First U.S. non-trial Hugo install; first U.S. pyeloplasty", "High",
     "https://www.paxtonmedia.com/news/high_point_enterprise/hospital-gets-new-robotic-system/article_cfb9af90-cd53-5d4c-8be3-850332eb973d.html",
     "Dr. Zachary McDowell. First U.S. hospital to install Hugo outside the IDE study."],
    ["Allegheny Health Network – West Penn Hospital", "Pittsburgh", "PA", "USA", "United States",
     "~Oct/Nov 2025 (Embrace Gynecology trial)", "Gynecology (total hysterectomy)",
     "Clinical Trial", "First U.S. Hugo hysterectomy", "High",
     "https://www.prnewswire.com/news-releases/ahn-west-penn-first-in-country-to-perform-hysterectomy-with-new-robotic-assisted-surgery-system-as-part-of-investigative-clinical-trial-302608891.html ; https://news.medtronic.com/2025-10-08-Medtronic-initiates-U-S-IDE-clinical-study-evaluating-Hugo-TM-robotic-assisted-surgery-system-for-gynecological-procedures",
     "Drs. Sarah Crafton & Eirwen Miller. Embrace Gynecology study: up to 70 patients across up to 5 U.S. hospitals."],
    ["Undisclosed Expand URO trial sites (×3)", "—", "—", "USA", "United States",
     "2022–2025", "Urology", "Clinical Trial", "", "Medium (count only)",
     "https://news.medtronic.com/2025-04-26-Medtronic-Expand-URO-U-S-clinical-trial-meets-safety-and-effectiveness-primary-endpoints-for-Hugo-TM-robotic-assisted-surgery-system",
     "Expand URO ran across 6 hospitals (137 patients, 11 surgeons); 3 sites not publicly named."],
    ["Undisclosed Enable Hernia / Embrace Gynecology sites", "—", "—", "USA", "United States",
     "2024–2026", "Hernia; Gynecology", "Clinical Trial", "", "Low (count only)",
     "https://www.prnewswire.com/news-releases/study-of-medtronic-hugo-robotic-assisted-surgery-system-in-hernia-repair-meets-safety-and-effectiveness-endpoints-302546743.html",
     "Enable Hernia (193 patients, multi-center) and Embrace Gynecology (up to 5 sites); only Duke and AHN West Penn named."],

    # ---------------- WESTERN EUROPE ----------------
    ["OLV Hospital (Onze-Lieve-Vrouw)", "Aalst", "—", "Belgium", "Western Europe",
     "2 Feb 2022", "Urology (radical prostatectomy)", "Commercial",
     "First Hugo procedure in Europe", "High",
     "https://www.surgicalroboticstechnology.com/news/first-procedure-performed-in-europe-with-medtronic-hugo-surgical-robotic-system/ ; https://www.fiercebiotech.com/medtech/medtronic-s-hugo-robotic-assisted-surgery-system-used-first-european-procedure",
     "Prof. Alexandre Mottrie (ORSI Academy). First site worldwide with both Hugo + Mazor."],
    ["Policlinico Universitario A. Gemelli IRCCS", "Rome", "—", "Italy", "Western Europe",
     "8 Apr 2022", "Urology (prostate)", "Commercial", "First in Italy", "High",
     "https://www.policlinicogemelli.it/en/news-events/al-gemelli-la-prima-performance-urologica-in-italia-per-il-robot-hugo/ ; https://www.ansa.it/canale_saluteebenessere/notizie/medicina/2022/04/08/al-gemelli-prima-performance-urologica-robot-chirurgo-hugo_1bb121d4-d1d7-460e-bc20-ef221cc74a77.html",
     "Team of Prof. Pierfrancesco Bassi. Italian training centre."],
    ["Evangelisches Krankenhaus (EvK) Herne", "Herne", "—", "Germany", "Western Europe",
     "Jul 2022", "General/visceral surgery", "Commercial", "First clinic in Germany with Hugo", "High",
     "https://ku-gesundheitsmanagement.de/2022/07/27/mit-op-roboter-hugo-haelt-innovative-spitzentechnik-einzug-im-evk-herne/ ; https://www.evk-herne.de/presse-service/themen/hugo",
     "~€1.6M investment."],
    ["Elisabeth-TweeSteden Ziekenhuis (ETZ)", "Tilburg", "—", "Netherlands", "Western Europe",
     "28 Nov 2022 (system since summer 2022)", "Gynaecology; urology; colorectal", "Commercial",
     "First hospital in the Netherlands", "High",
     "https://tilburg.com/nieuws/primeur-in-nederland-eerste-operatie-in-tilburgs-ziekenhuis-met-robot-hugo/ ; https://fmtgezondheidszorg.nl/eerste-nederlandse-ziekenhuis-met-operatierobot-hugo/",
     "First case by gynaecologist Petra Janssen."],
    ["CHU de Rennes", "Rennes", "—", "France", "Western Europe",
     "6 Dec 2022 (urology); 12 Dec 2022 (digestive)", "Urology (prostatectomy); digestive", "Commercial",
     "First in France", "High",
     "https://www.dhmagazine.fr/au-chu-de-rennes-les-equipes-de-chirurgie-urologique-et-digestive-realisent-les-deux-premieres-interventions-francaises-assistees-du-robot-hugo-tm",
     "Prof. Romain Mathieu (urology), Prof. Laurent Sulpice (digestive)."],
    ["HUS / Helsinki University Hospital", "Helsinki", "—", "Finland", "Western Europe",
     "12 Jan 2023", "Urology (prostate); kidney/bladder/ureter", "Commercial", "First in Finland", "High",
     "https://www.hus.fi/ajankohtaista/uusi-hugo-ras-leikkausrobotti-vahvistaa-robottiavusteisen-kirurgian-roolia-husissa-0 ; https://www.laakarilehti.fi/terveydenhuolto/husissa-kayttoon-uusi-hugo-ras-leikkausrobotti/",
     "One of the first in Europe."],
    ["Mútua Terrassa (HUMT)", "Terrassa", "—", "Spain", "Western Europe",
     "Jan 2023 (urology); Feb 2023 (gynae); Mar 2023 (colorectal)", "Urology; gynaecology; general/digestive",
     "Commercial", "First gynaecological Hugo surgery in Spain", "High",
     "https://www.mutuaterrassa.com/es/primera-intervencion-hugo-ras ; https://www.mutuaterrassa.com/es/100-intervenciones-hugo-ras",
     "Dr. Jordi Cassadó. Established Europe's first Hugo RAS surgical committee."],
    ["St. Josef-Hospital (Katholisches Klinikum Bochum)", "Bochum", "—", "Germany", "Western Europe",
     "13 Feb 2023", "Colorectal (sigmoidectomy); general", "Commercial",
     "First in-human colorectal Hugo procedure in Germany", "High",
     "https://www.klinikum-bochum.de/medien/aktuelles-details/neue-generation-der-op-roboter-startet-in-bochum.html ; https://medecon.ruhr/2023/02/neue-generation-der-op-roboter-startet-in-bochum/",
     ""],
    ["UKSH Campus Lübeck (Univ. Hospital Schleswig-Holstein)", "Lübeck", "—", "Germany", "Western Europe",
     "Mar 2023", "Urology (partial nephrectomy); visceral", "Commercial",
     "First Germany-wide Hugo kidney partial resection", "High",
     "https://www.uksh.de/Service/Presse/Presseinformationen/2024/ ; https://www.kgsh.online/post/deutschlandweit-erste-nieren-teilresektion-mit-modernster-roboterchirurgie-am-campus-l%C3%BCbeck",
     "~50 procedures by report date."],
    ["Oslo University Hospital, Rikshospitalet", "Oslo", "—", "Norway", "Western Europe",
     "2023", "Adrenalectomy; urology; gynae; general", "Commercial", "First in Norway", "High",
     "https://www.oslo-universitetssykehus.no/om-oss/innsikt/opererer-med-ny-robot-pa-ous-rikshospitalet/",
     "Proctored by British and Belgian surgeons."],
    ["Hospital de Santo António (CHUdSA)", "Porto", "—", "Portugal", "Western Europe",
     "Mar 2023", "Urology (prostate cancer)", "Commercial", "Early Portuguese adopter", "Medium",
     "https://www.pcguia.pt/2023/03/hospital-de-santo-antonio-comeca-a-operar-cancros-da-prostata-com-robot-cirurgiao-hugo/",
     ""],
    ["Guy's and St Thomas' NHS Foundation Trust", "London", "—", "United Kingdom", "Western Europe",
     "Jun 2023", "Urology (prostatectomy); multispecialty", "Commercial", "First UK NHS trust to adopt Hugo", "High",
     "https://www.guysandstthomas.nhs.uk/news/new-surgical-robot-makes-uk-debut-guys-and-st-thomas ; https://www.medicaldevice-network.com/news/medtronics-hugo-robot-debuts-at-guys-and-st-thomas-hospital/",
     "Largest NHS robotic programme (>1,500 cases/yr). Running IDEAL comparative study."],
    ["Universitätsklinikum Carl Gustav Carus Dresden", "Dresden", "—", "Germany", "Western Europe",
     "17 Oct 2023", "Urology (prostate adenomectomy)", "Commercial", "First urological Hugo use in Germany", "High",
     "https://oiger.de/2023/10/18/uniklinik-dresden-setzt-erstmals-op-roboter-hugo-in-der-urologie-ein/188612 ; https://uroforum.de/op-roboter-hugo-kommt-erstmals-in-deutschland-in-der-urologie-zum-einsatz/",
     "Patient aged 78; ~€1.7M device; only such device in Central Germany at the time."],
    ["Hospital Clínic de Barcelona", "Barcelona", "—", "Spain", "Western Europe",
     "2023", "Urology (kidney surgery)", "Commercial", "First public multispecialty hospital in Spain", "High",
     "https://www.clinicbarcelona.org/en/news/the-hospital-clinic-the-first-public-hospital-in-spain-to-operate-with-the-hugo-robot",
     "Dr. Antonio Alcaraz."],
    ["Policlinico Universitario Campus Bio-Medico", "Rome", "—", "Italy", "Western Europe",
     "Oct 2022 onward", "Urology; gynaecology; general", "Commercial", "", "Medium",
     "https://www.policlinicocampusbiomedico.it/news/innovazione-debutta-la-chirurgia-robotica-alla-fondazione-policlinico-universitario-campus-bio-medico",
     "Kidney/prostate/bladder as first procedures."],
    ["Ospedale Generale Regionale F. Miulli", "Acquaviva delle Fonti (Bari)", "—", "Italy", "Western Europe",
     "2023–2024", "Gynaecology; urology", "Commercial", "", "Medium",
     "https://www.miulli.it/chirurgia-robotica-allospedale-miulli-arriva-hugo-ras-una-delle-strumentazioni-mediche-piu-avanzate-al-mondo/",
     "Among first in world in Medtronic 'Partners in Possibility' program."],
    ["Pineta Grande Hospital", "Castel Volturno (Caserta)", "—", "Italy", "Western Europe",
     "2023–2024", "Urology (radical nephrectomy of pelvic kidney)", "Commercial", "", "Medium",
     "https://www.donatodente.it/chirurgia-robotica-urologia-pineta-grande-hospital/",
     "Reported first Italian Hugo radical nephrectomy of a pelvic kidney."],
    ["IRCCS Ospedale Sacro Cuore Don Calabria", "Negrar (Verona)", "—", "Italy", "Western Europe",
     "19 Mar 2024 (first case); 7 Jun 2024 (inaugurated)", "Urology (prostatectomy); multispecialty", "Commercial", "", "Medium",
     "https://www.sacrocuore.it/news/con-la-festa-del-sacro-cuore-taglio-del-nastro-del-nuovo-robot-chirurgico-hugo/",
     "Urology director Dr. Stefano Cavalleri."],
    ["Hospital General Universitario de Elche", "Elche", "—", "Spain", "Western Europe",
     "Sep 2024", "General/digestive (colon cancer)", "Commercial",
     "First colon cancer resection in Valencia region with Hugo", "Medium",
     "https://elche.san.gva.es/es/noticias/",
     "~€2.38M regional investment."],
    ["Hospital Universitario de La Ribera", "Alzira (Valencia)", "—", "Spain", "Western Europe",
     "2024–2025", "Digestive (cholecystectomy); gynaecology", "Commercial", "", "Medium",
     "https://laribera.san.gva.es/es/",
     "Digestive unit led by Dr. Gloria Báguena."],
    ["HM Hospitales (group)", "Madrid", "—", "Spain", "Western Europe",
     "2024–2025", "Multispecialty", "Commercial", "Designated Hugo centre of excellence for training", "Medium",
     "https://www.hmhospitales.com/comunicado-prensa/hm-hospitales-completa-sus-diez-primeras-cirugias-roboticas-con-hugo-ras-system-con-una-gran-satisfaccion-por-su-potencial/",
     "Private group; 10+ first robotic cases."],
    ["Medway NHS Foundation Trust", "Gillingham (Medway)", "—", "United Kingdom", "Western Europe",
     "2024 (training from Feb)", "Gynaecology (hysterectomy)", "Commercial",
     "First robotic-assisted surgery at Medway", "Medium",
     "https://www.medway.nhs.uk/news/gynaecology-patients-benefit-from-new-robotic-assisted-surgical-device/",
     "16 hysterectomies via Hugo."],
    ["East Kent Hospitals (William Harvey, Ashford; + QEQM, Margate)", "Ashford / Margate", "—", "United Kingdom", "Western Europe",
     "2024", "Colorectal; multispecialty", "Commercial", "", "Medium",
     "https://www.ekhuft.nhs.uk/news/first-patients-have-robotic-surgery-at-william-harvey-hospital/",
     "Hugo installed at William Harvey for complex colorectal; second Hugo at QEQM Margate same year."],
    ["Clinique de Genolier (Swiss Medical Network)", "Genolier", "—", "Switzerland", "Western Europe",
     "15 Apr 2025 (welcomed Feb 2025)", "Urology/general", "Commercial",
     "First private clinic in French-speaking Switzerland with Hugo", "High",
     "https://www.swissmedical.net/en/news-events/20250415_cdg_news_1ere_op_hugo ; https://www.swissmedical.net/en/news-events/20250206_cdg_bienvenue_hugo",
     "First facility on the Lake Geneva arc with Hugo."],
    ["Polyclinique Pau Pyrénées", "Pau", "—", "France", "Western Europe",
     "Feb 2025", "Urology (functional urology + prostatectomy)", "Commercial", "", "Medium",
     "https://polycliniquepaupyrenees.fr/actualites/arrivee-du-robot-hugo-de-medtronic-la-polyclinique-pau-pyrenees",
     "Drs. Philippe Chevallier, Julien Casenave; later passed 100 procedures."],
    ["CHU de Lille (Hôpital Claude Huriez)", "Lille", "—", "France", "Western Europe",
     "~2024", "Urology", "Commercial", "", "Medium",
     "https://www.facebook.com/chulille/posts/855843636569480/",
     "Prof. Arnaud Villers / Dr. Jonathan Olivier. Reached 100 Hugo operations."],
    ["Groupe Urologie Saint-Augustin (Clinique Saint-Augustin, ELSAN)", "Bordeaux", "—", "France", "Western Europe",
     "2024", "Urology", "Commercial", "", "Medium",
     "https://www.groupe-urologie.com/le-groupe-urologie-saint-augustin-sequipe-du-systeme-de-chirurgie-robotique-hugo/ ; https://www.elsan.care/fr/clinique-saint-augustin/nos-actualites/decouvrez-le-nouveau-robot-hugo",
     "Operates both Da Vinci and Hugo; ERUS 2024 host city."],
    ["ULS Viseu Dão-Lafões (Hospital São Teotónio)", "Viseu", "—", "Portugal", "Western Europe",
     "12 Feb 2026", "General/colorectal; gynae; urology", "Commercial",
     "First robotic surgery at this unit", "High",
     "https://recuperarportugal.gov.pt/2026/02/12/prr-reforca-inovacao-cirurgica-na-uls-viseu-dao-lafoes-com-primeira-cirurgia-robotica-2/ ; https://www.theportugalnews.com/pt/noticias/2026-02-12/hospital-portugues-realiza-com-exito-a-primeira-cirurgia-robotica/958072",
     "€1.9M PRR-funded; LigaSure RAS also debuted."],
    ["Rigshospitalet", "Copenhagen", "—", "Denmark", "Western Europe",
     "2023–2024 (exact date not pinned)", "Urology", "Commercial", "First in Denmark (likely)", "Medium",
     "https://www.verdensmaal.org/nyheder/robotten-hugos-h%C3%A6nder-hj%C3%A6lper-kirurgerne-pa-operationsbordet",
     "Prof. Andreas Røder. Confirmed Hugo user; first-surgery date not located."],

    # ---------------- EASTERN EUROPE / MIDDLE EAST / AFRICA ----------------
    ["Wojewódzki Szpital Specjalistyczny Nr 2", "Jastrzębie-Zdrój", "—", "Poland", "Eastern Europe",
     "9 Dec 2024", "Urology (prostatectomy); colorectal; gynaecology", "Commercial", "First Hugo RAS in Poland", "High",
     "https://www.wss2.pl/wykonalismy-juz-50-zabiegow-z-wykorzystaniem-robota-hugo-ras/ ; https://www.rynekzdrowia.pl/Aparatura-i-wyposazenie/Juz-nie-tylko-Da-Vinci-W-Jastrzebiu-Zdroju-przeprowadzono-pierwsze-zabiegi-HUGO-RAS,266862,5.html",
     "50 cases by mid-2025; targeted 200+/yr."],
    ["Szpital św. Łukasza (St. Luke's Hospital)", "Bolesławiec", "—", "Poland", "Eastern Europe",
     "31 Mar 2025", "General/colorectal (right hemicolectomy); urology", "Commercial",
     "First Polish county ('powiat') hospital with own robot", "High",
     "https://www.isbzdrowie.pl/2025/05/boleslawiec-ma-robota-chirurgicznego/ ; https://powiatboleslawiecki.pl/10400-2/",
     "2nd in Poland. Drs. W. Hap, K. Janiszewski."],
    ["Narodowy Instytut Onkologii (National Institute of Oncology), Kraków branch", "Kraków", "—", "Poland", "Eastern Europe",
     "Autumn 2025 (installed mid-2025)", "Oncologic surgery; urology; gynaecology; general", "Commercial",
     "First in Małopolska region", "High",
     "https://krakow.nio.gov.pl/robot-operacyjny-hugo-ras-nowy-rozdzial-w-krakowskiej-chirurgii-onkologicznej/ ; https://www.rynekzdrowia.pl/Aparatura-i-wyposazenie/Milowy-krok-trzeci-taki-robot-medyczny-w-Polsce-Operacje-w-Krakowie-juz-jesienia,274815,5.html",
     "3rd in Poland. ~17M PLN, EU-funded."],
    ["Wojewódzki Szpital im. Kard. Stefana Wyszyńskiego", "Łomża", "—", "Poland", "Eastern Europe",
     "7–9 Apr 2026 (installed Oct 2025)", "Urology (radical prostatectomy)", "Commercial", "", "High",
     "https://wspolczesna.pl/robotem-w-raka-prostaty-po-raz-pierwszy-w-szpitalu-w-lomzy-przeprowadzono-operacje-z-wykorzystaniem-robota-chirurgicznego-hugo-ras/ar/c14p2-28897693 ; https://www.mp.pl/kurier/402946,lomza-szpital-wkroczyl-w-ere-robotyki",
     "4th in Poland. ~10.2M PLN, KPO (EU Recovery Fund) financed."],
    ["Szpital SPZOZ MSWiA (Ministry of Interior Hospital)", "Gdańsk", "—", "Poland", "Eastern Europe",
     "5–6 Mar 2026", "General/colorectal; oncologic; urology; gynaecology", "Commercial", "First in Pomerania", "High",
     "https://radiogdansk.pl/wiadomosci/region/trojmiasto/2026/03/06/pierwsze-zabiegi-na-pomorzu-z-uzyciem-robota-hugo-ras-to-kolejny-etap-rozwoju-chirurgii/ ; https://onkopedia.pl/news/robot-hugo-ras-w-gdanskim-szpitalu-mswia-zwiekszy-precyzje-operacji-onkologicznych/",
     "5th in Poland. ~12.3M PLN, KPO-funded; 3 surgeons trained."],
    ["Wojewódzki Szpital Zespolony im. L. Perzyny", "Kalisz", "—", "Poland", "Eastern Europe",
     "28–29 Jan 2026 (demo/evaluation)", "Multi-specialty (trial)", "Demo / Evaluation", "", "Medium",
     "https://alertmedyczny.pl/robot-hugo-ras-w-kolejnej-sali-operacyjnej-kalisz-testuje-system-medtronic/ ; https://kalisz24.info.pl/operacja-hugo-nowoczesny-system-chirurgii-robotycznej-w-rekach-kaliskich-lekarzy/",
     "Evaluation only as of report; considering ~12M PLN acquisition."],
    ["Ovidius Clinical Hospital (OCH)", "Constanța", "—", "Romania", "Eastern Europe",
     "Dec 2025", "Urology (nephrectomy); general; gynaecology", "Commercial", "First Hugo RAS in Romania", "High",
     "https://ovidius-ch.ro/evenimente/447-ovidius-clinical-hospital-aduce-in-premiera-in-romania-robotul-chirurgical-hugotm-ras-de-la-medtronic ; https://www.ctnews.ro/premiera-nationala-la-constanta-ovidius-clinical-hospital-realizeaza-primele-operatii-robotice-cu-sistemul-hugo-ras/",
     "Only Romanian hospital with Hugo at the time. Drs. I. Dragomirișteanu, O. Azis."],
    ["Koç University Hospital + RMK AIMES simulation center", "Istanbul", "—", "Turkey", "Middle East",
     "4 Feb 2025 (partnership)", "Robotic surgery training (urology focus)", "Training Center",
     "First Hugo RAS placement in Turkey", "High",
     "https://enyakinhastane.com/medtronic-koc-universitesi-hastanesi-ve-rmk-aimes-robotik-cerrahi-alaninda-turkiyede-bir-ilke-imza-atiyor/ ; https://www.businessmed.com.tr/ilac-sektoru/medtronic-koc-universitesi-hastanesi-ve-rmk-aimes-robotik-asiste-cerrahi-alanindaki-is-birligi-ile-turkiyede-bir-ilke-imza-atiyor/",
     "Selected as a Medtronic regional training center (transcontinental)."],
    ["BDF Royal Medical Services (Bahrain Defence Force)", "Riffa / Manama", "—", "Bahrain", "Middle East",
     "~Aug 2024", "Urology; general surgery; obstetrics & gynaecology", "Commercial",
     "First Hugo RAS in the entire Middle East & North Africa (MENA)", "High",
     "https://www.newsofbahrain.com/bahrain/116474.html ; https://www.gdnonline.com/Details/1359614 ; https://www.gdnonline.com/Details/1319853/Surgical-robot-first-in-ops-to-remove-malignant-tumours",
     "First MENA institution to reach 100 Hugo cases."],

    # ---------------- LATIN AMERICA ----------------
    ["Clínica Santa María", "Santiago", "—", "Chile", "Latin America",
     "19 Jun 2021", "Urology (radical prostatectomy)", "Commercial", "FIRST Hugo RAS surgery in the WORLD", "High",
     "https://news.medtronic.com/2021-06-22-First-Procedure-in-the-World-with-Medtronic-Hugo-TM-Robotic-Assisted-Surgery-System-Performed-at-Clinica-Santa-Maria-in-Chile ; https://www.prnewswire.com/news-releases/first-procedure-in-the-world-with-medtronic-hugo-robotic-assisted-surgery-system-performed-at-clinica-santa-maria-in-chile-301316903.html",
     "Dr. Rubén Olivares (and Dr. Alfredo Velasco). Launched the Hugo global clinical registry."],
    ["Hospital Pacífica Salud", "Panama City", "—", "Panama", "Latin America",
     "Jul 2021", "Gynaecology (world-first); urology; colorectal", "Commercial",
     "WORLD-FIRST Hugo gynaecological procedures; first in Central America", "High",
     "https://news.medtronic.com/2021-07-29-First-Gynecological-Procedures-Performed-with-Medtronic-Hugo-TM-Robotic-Assisted-Surgery-System ; https://www.pacificasalud.com/500-cirugias-realizadas-con-el-sistema-robotico-hugo-ras-de-medtronic/",
     "Drs. Salomón Zebede & Juan Carlos López (gyn); Drs. Homero Rodríguez & Moisés Cukier (colorectal). Passed 500 surgeries."],
    ["Hospital Clínico Red de Salud UC CHRISTUS", "Santiago", "—", "Chile", "Latin America",
     "Jul 2021 (announced 20 Jul)", "Urology; general (teaching program)", "Commercial", "", "High",
     "https://news.medtronic.com/2021-07-20-Medtronic-Hugo-TM-Robotic-Assisted-Surgery-System-Cornerstone-of-New-Robotics-Program-in-Latin-America",
     "Joined 'Partners in Possibility.' Affiliated with Pontificia Universidad Católica de Chile."],
    ["Medtronic Surgical Robotics Experience Center (UC Chile / UC CHRISTUS)", "Santiago", "—", "Chile", "Latin America",
     "Jul 2021", "Training/demo (all specialties)", "Training Center",
     "First Medtronic Surgical Robotics Experience Center in Latin America", "High",
     "https://news.medtronic.com/2021-07-20-Medtronic-Hugo-TM-Robotic-Assisted-Surgery-System-Cornerstone-of-New-Robotics-Program-in-Latin-America",
     "1 of 10 worldwide. Regional training hub."],
    ["Hospital Israelita Albert Einstein", "São Paulo", "—", "Brazil", "Latin America",
     "May 2023", "Urology (radical prostatectomy); general; gynaecology", "Commercial", "First Hugo RAS in Brazil", "High",
     "https://www.anahp.com.br/noticias/einstein-incorpora-em-seu-parque-robotico-equipamento-inedito-da-medtronic-no-brasil/ ; https://olhardigital.com.br/2023/05/09/medicina-e-saude/robo-cirurgiao-estreia-no-brasil-transmitindo-operacao-em-3d/",
     "Anvisa approval Apr 2022. First case transmitted in 3D."],
    ["Faculdade de Medicina do ABC (FMABC) – Urologia", "Santo André (SP)", "—", "Brazil", "Latin America",
     "2023", "Urology", "Commercial", "", "Medium",
     "https://www.ncbi.nlm.nih.gov/pmc/articles/PMC10947622/",
     "Co-author institution on Brazil's first Hugo urology series; standalone install date unconfirmed."],
    ["IRCAD América Latina – Unidade Barretos (Hospital de Amor)", "Barretos (SP)", "—", "Brazil", "Latin America",
     "2024", "Training (minimally invasive/robotic)", "Training Center",
     "First Hugo RAS training room in a Latin American training center", "High",
     "https://www.ircadamericalatina.com.br/barretos/noticias/ircad-barretos-inaugura-sala-robotica-experimental-com-a-chegada-da-plataforma-hugo-ras/",
     "14-yr Medtronic–IRCAD partnership."],
    ["Hospital Nove de Julho (Dasa)", "São Paulo", "—", "Brazil", "Latin America",
     "2024–2025", "Urology; general; gynaecology", "Commercial", "", "Medium",
     "https://www.saudebusiness.com/hospitais/hospital-nove-de-julho-investe-em-novo-robo-que-reduz-em-cerca-de-30-os-custos-das",
     "Dasa network. Positioned as ~30% cost reduction vs other robots."],
    ["Hospital Ángeles Acoxpa", "Mexico City", "—", "Mexico", "Latin America",
     "30 Sep 2025 (presented; first surgeries shortly before)", "Urology; gynaecology; general", "Commercial",
     "First Hugo RAS in Mexico", "High",
     "https://blog.hospitalangeles.com/posts/presentan-primer-robot-hugo-ras-en-hospital-angeles-acoxpa/ ; https://www.excelsior.com.mx/nacional/hospital-angeles-acoxpa-lleva-a-cabo-la-inauguracion-de-robot-hugo-ras/1743455",
     "First 4 certified surgeons named. Reached 100 surgeries in <1 yr."],
    ["Hospital Nacional Rosales (new building)", "San Salvador", "—", "El Salvador", "Latin America",
     "1–2 Jun 2026", "General; urology; gynaecology; oncology", "Commercial (unconfirmed as Hugo)",
     "First robotic surgery in a Salvadoran public hospital", "Low",
     "https://www.elsalvador.com/h-entretenimiento/h-tecnologia/tecnologia-hugo-ras-un-hito-en-cirugia-robotica-con-98-5-de-exito/1215670/2025/ ; https://www.infobae.com/el-salvador/2026/06/03/tecnologia-robotica-y-cirugias-complejas-marcan-el-inicio-de-funciones-en-el-nuevo-hospital-rosales-de-el-salvador/",
     "LIKELY but unconfirmed: one 2025 article cited 'Hugo RAS'; official inauguration coverage did not name the manufacturer."],

    # ---------------- ASIA-PACIFIC ----------------
    ["Apollo Hospitals (Greams Road)", "Chennai", "—", "India", "Asia-Pacific",
     "17 Sep 2021", "Urology (prostatectomy); uro-oncology; gynaecology", "Commercial",
     "First Hugo case in all of Asia-Pacific", "High",
     "https://news.medtronic.com/2021-09-27-First-Procedure-in-Asia-Pacific-Performed-with-Medtronic-Hugo-TM-Robotic-Assisted-Surgery-System ; https://www.prnewswire.com/in/news-releases/apollo-hospitals-chennai-performs-worlds-first-robotic-cancer-surgery-for-lymph-node-removal-using-hugo-ras-platform-via-lateral-approach-302785576.html",
     "Dr. N. Ragavan. On 29 May 2026 same site did 'world-first' robotic VEIL inguinal lymphadenectomy for penile cancer."],
    ["Medtronic Surgical Robotics Experience Center (SREC)", "Gurugram", "—", "India", "Asia-Pacific",
     "Sep 2021", "Training/demo (all specialties)", "Training Center",
     "First SREC in Asia-Pacific", "High",
     "https://www.business-standard.com/content/press-releases-ani/medtronic-showcases-hugo-robotic-assisted-surgery-system-at-the-inauguration-of-its-first-surgical-robotics-experience-center-in-india-121091601139_1.html",
     "Inaugurated by Dr. BK Rao (Sir Ganga Ram Hospital)."],
    ["CARE Hospitals (Banjara Hills)", "Hyderabad", "—", "India", "Asia-Pacific",
     "15 Sep 2022", "Gynaecology (hysterectomy, myomectomy)", "Commercial",
     "First gynaecology Hugo procedure in Asia-Pacific", "High",
     "https://telanganatoday.com/care-hospitals-performs-first-gynecology-procedure-in-asia-pacific-using-hugo-ras-system ; https://pmc.ncbi.nlm.nih.gov/articles/PMC11877460/",
     "Dr. Manjula Anagani. First Hugo install in Telangana/AP. 20-patient series published 2025."],
    ["Sir Ganga Ram Hospital", "New Delhi", "—", "India", "Asia-Pacific",
     "~2023–2024", "Urology", "Commercial", "", "Medium",
     "https://sgrh.com/en/facilities-and-technology ; https://www.biospectrumindia.com/news/97/19572/medtronic-launches-robotic-assisted-surgery-platform-hugo-ras-system-in-india.html",
     "Exact first-case date not pinned down."],
    ["Fujita Health University Hospital", "Toyoake (Aichi)", "—", "Japan", "Asia-Pacific",
     "Aug 2023 – Feb 2024", "Urology (robot-assisted radical prostatectomy)", "Commercial",
     "First reported Hugo RARP in Japan", "Medium",
     "https://onlinelibrary.wiley.com/doi/10.1111/ases.13342",
     "Single-center initial experience, 13 cases."],
    ["Tottori University Hospital", "Yonago (Tottori)", "—", "Japan", "Asia-Pacific",
     "Mar–Jun 2023", "Gynaecology (total laparoscopic hysterectomy)", "Commercial",
     "First reported Hugo use in gynaecology in Japan", "Medium",
     "https://pubmed.ncbi.nlm.nih.gov/38070071/ ; https://link.springer.com/article/10.1007/s13304-023-01710-5",
     "'First report of robotic-assisted total hysterectomy using Hugo.'"],
    ["Kyoto University Hospital", "Kyoto", "—", "Japan", "Asia-Pacific",
     "Reported Jan 2026 (publication)", "General/HPB (distal pancreatectomy)", "Commercial",
     "'First nationwide' Japan Hugo distal pancreatectomy", "Medium",
     "https://www.ncbi.nlm.nih.gov/pmc/articles/PMC12805036/",
     "Japan approved Hugo for GI surgery May 2023."],
    ["Sapporo Medical University", "Sapporo", "—", "Japan", "Asia-Pacific",
     "Reported 2024", "Colorectal (cylindrical APR for rectal cancer)", "Commercial",
     "'First-ever' Hugo case report for rectal cancer", "Medium",
     "https://doi.org/10.1111/ases.13321",
     ""],
    ["Seoul National University Hospital (SNUH)", "Seoul", "—", "South Korea", "Asia-Pacific",
     "8 May 2025", "Urology (prostatectomy); HPB (pancreaticoduodenectomy)", "Commercial", "First in Korea to adopt Hugo", "High",
     "https://news.medtronic.com/First-commercial-procedure-using-Medtronic-Hugo-TM-robotic-assisted-surgery-system-in-Korea ; https://www.koreabiomed.com/news/articleView.html?idxno=27570",
     "Prostatectomy by Prof. Chang Wook Jeong; Whipple by Prof. Jin-young Jang. Korea approved Hugo 2024."],
    ["Asan Medical Center", "Seoul", "—", "South Korea", "Asia-Pacific",
     "6 May 2025 (partnership)", "Research/training, multi-specialty", "Training Center", "", "High",
     "https://www.koreabiomed.com/news/articleView.html?idxno=31735 ; https://www.massdevice.com/medtronic-surgical-robotics-collab-korea-hugo/",
     "Strategic partnership with Medtronic to expand Hugo clinical research/training in Korea."],
    ["Tungs' Taichung MetroHarbor Hospital", "Taichung", "—", "Taiwan", "Asia-Pacific",
     "After Nov 2022 (TFDA approval)", "Urology (radical prostatectomy)", "Commercial",
     "Pioneer Hugo experience in Taiwan and Northeast Asia", "Medium",
     "https://onlinelibrary.wiley.com/doi/full/10.1002/rcs.2577 ; https://pmc.ncbi.nlm.nih.gov/articles/PMC10969029/",
     "First 12 RARP cases. Taiwan = 2nd Asian country to clinically adopt Hugo."],
    ["Hurstville Private Hospital", "Sydney", "—", "Australia", "Asia-Pacific",
     "2023", "Colorectal/general; urology; gynaecology", "Commercial", "First hospital in Australia to offer Hugo RAS", "High",
     "https://hurstvilleprivate.com.au/news/2023/hurstville-private-hospital-is-first-in-australia-to-offer-robotic-procedures-with-newest-technology",
     "Recognized Center of Excellence in Robotic/Colorectal/MIS surgery."],
    ["Australian Metabolic & Obesity Surgery (AMOS) clinic", "(NSW)", "—", "Australia", "Asia-Pacific",
     "2023–2024", "Bariatric/metabolic (sleeve gastrectomy)", "Commercial",
     "Reported 'Australia's first bariatric sleeve gastrectomy' with Hugo", "Low",
     "https://amos.clinic/hugo-ras-robot-assisted-gastric-sleeve-surgery/",
     "Exact host facility to verify."],
    ["Southern Cross Wellington Hospital", "Wellington", "—", "New Zealand", "Asia-Pacific",
     "Jun 2025", "Urology (radical prostatectomy)", "Commercial", "First Hugo robotic prostatectomy in New Zealand", "High",
     "https://www.southerncross.co.nz/news/2025/southern-cross-wellington-hospital-introduces-hug-robotic-assisted-surgery ; https://precisionsurgery.co.nz/medtronic-hugo-robotic-surgery/",
     "Surgeon Jim Duthie."],
    ["Royston Hospital", "Hastings (Hawke's Bay)", "—", "New Zealand", "Asia-Pacific",
     "2025", "Urology", "Commercial", "First robotic-assisted surgery in Hawke's Bay", "Medium",
     "https://www.nzdoctor.co.nz/article/undoctored/royston-hospital-surgical-robot-first-hawkes-bay ; https://www.nzherald.co.nz/hawkes-bay-today/news/robotic-assisted-surgery-introduced-in-hawkes-bay-for-the-first-time/premium/P6MYCWRVGZBATBLY4SWMBFNFWU/",
     "Jim Duthie offers both Hugo and da Vinci."],
    ["Chinese University of Hong Kong / Prince of Wales Hospital", "Hong Kong", "—", "Hong Kong (China)", "Asia-Pacific",
     "Study registered 2024 (NCT06669104)", "Thoracic surgery (lung/robotic thoracic)", "Clinical Trial", "", "Medium",
     "https://clinicaltrials.gov/study/NCT06669104 ; https://www.centerwatch.com/clinical-trials/listings/NCT06669104/",
     "Hugo placed for a thoracic surgery clinical study; PI Prof. Calvin Sze Hang Ng."],
]

# ---------------------------------------------------------------------------
# Milestones & regulatory timeline
# ---------------------------------------------------------------------------
MILESTONES = [
    ["19 Jun 2021", "First procedure in the world", "Robotic radical prostatectomy, Clínica Santa María, Santiago, Chile (Dr. Rubén Olivares). Started the Hugo global registry.",
     "https://news.medtronic.com/2021-06-22-First-Procedure-in-the-World-with-Medtronic-Hugo-TM-Robotic-Assisted-Surgery-System-Performed-at-Clinica-Santa-Maria-in-Chile"],
    ["17 Sep 2021", "First procedure in Asia-Pacific", "Robotic prostatectomy, Apollo Hospitals Greams Road, Chennai, India (Dr. N. Ragavan).",
     "https://news.medtronic.com/2021-09-27-First-Procedure-in-Asia-Pacific-Performed-with-Medtronic-Hugo-TM-Robotic-Assisted-Surgery-System"],
    ["11 Oct 2021", "CE Mark (Europe) — urology + gynecology", "First CE Mark authorizing EU sale for urologic and gynecologic procedures.",
     "https://news.medtronic.com/2021-10-11-Medtronic-Hugo-TM-Robotic-Assisted-Surgery-System-Receives-European-CE-Mark-Approval"],
    ["Dec 2021", "Health Canada licence", "Urologic and gynecologic laparoscopic procedures.",
     "https://www.biospace.com/medtronic-hugo-and-8482-robotic-assisted-surgery-system-receives-health-canada-licence"],
    ["2 Feb 2022", "First European procedure", "Performed in Aalst, Belgium (OLV Hospital), following CE Mark.",
     "https://www.fiercebiotech.com/medtech/medtronic-s-hugo-robotic-assisted-surgery-system-used-first-european-procedure"],
    ["May 2023", "Japan MHLW approval — gastroenterological surgery", "Early Japanese indication for GI surgery.",
     "https://www.ncbi.nlm.nih.gov/pmc/articles/PMC12805036/"],
    ["Jan 2024", "Three regulatory clearances bundled", "CE Mark for general surgery (Europe); Health Canada general laparoscopic; Japan MHLW urology + gynecology. Coverage to ~80% of robotic procedures.",
     "https://www.mddionline.com/robotics/medtronic-s-hugo-clears-3-major-regulatory-hurdles"],
    ["2024", "Korea (MFDS) approval", "Laparoscopic/endoscopic procedures incl. radical prostatectomy and cholecystectomy.",
     "https://www.massdevice.com/medtronic-surgical-robotics-collab-korea-hugo/"],
    ["Q1 CY2025", "U.S. FDA urology submission filed", "Hugo urology indication submitted to FDA.",
     "https://www.sec.gov/Archives/edgar/data/0001613103/000161310325000141/exhibit991-fy26q1earningsr.htm"],
    ["26 Apr 2025", "Expand URO meets primary endpoints", "Results at AUA 2025. 137 patients; met safety + effectiveness endpoints.",
     "https://news.medtronic.com/2025-04-26-Medtronic-Expand-URO-U-S-clinical-trial-meets-safety-and-effectiveness-primary-endpoints-for-Hugo-TM-robotic-assisted-surgery-system"],
    ["6 May 2025", "Korea partnership — Asan Medical Center", "Strategic partnership for clinical research, training, technology development.",
     "https://www.koreabiomed.com/news/articleView.html?idxno=31735"],
    ["8 May 2025", "First commercial procedure in Korea", "SNUH — prostatectomy (Prof. Chang Wook Jeong) and Whipple (Prof. Jin-young Jang).",
     "https://news.medtronic.com/First-commercial-procedure-using-Medtronic-Hugo-TM-robotic-assisted-surgery-system-in-Korea"],
    ["Jul 2025", "CE Mark — LigaSure vessel-sealing on Hugo", "Advanced vessel-sealing technology cleared for Hugo in Europe.",
     "https://news.medtronic.com/Medtronic-shaping-future-of-surgery,-secures-CE-Mark-for-LigaSure-TM-technology-on-Hugo-TM-robotic-assisted-surgery-system"],
    ["4 Sep 2025", "Enable Hernia Repair meets endpoints", "First-ever U.S. IDE study for robotic hernia surgery; 193 patients (inguinal + ventral).",
     "https://news.medtronic.com/2025-09-04-Study-of-Medtronic-Hugo-TM-robotic-assisted-surgery-system-in-hernia-repair-meets-safety-and-effectiveness-endpoints"],
    ["8 Oct 2025", "Embrace Gynecology IDE initiated", "Third U.S. IDE study; first hysterectomies at AHN West Penn Hospital. Up to 70 patients across up to 5 U.S. hospitals.",
     "https://news.medtronic.com/2025-10-08-Medtronic-initiates-U-S-IDE-clinical-study-evaluating-Hugo-TM-robotic-assisted-surgery-system-for-gynecological-procedures"],
    ["18 Nov 2025", "FY2026 Q2 earnings", "Hugo named an enterprise growth driver. Q2 revenue $8.961B (+6.6% reported).",
     "https://news.medtronic.com/2025-11-18-Medtronic-reports-strong-second-quarter-fiscal-2026-financial-results"],
    ["3 Dec 2025", "U.S. FDA clearance — urology", "Cleared for minimally invasive urologic procedures (prostatectomy, nephrectomy, cystectomy). ~230,000 such U.S. surgeries/year.",
     "https://news.medtronic.com/2025-12-03-Medtronic-announces-FDA-clearance-of-Hugo-TM-robotic-assisted-surgery-system-for-urologic-surgical-procedures"],
    ["17 Feb 2026", "First U.S. commercial surgery", "Robotic-assisted prostatectomy by Dr. Jihad Kaouk at Cleveland Clinic.",
     "https://news.medtronic.com/2026-02-17-Medtronic-announces-first-surgery-with-Hugo-TM-robotic-assisted-surgery-system-in-the-U-S-performed-at-Cleveland-Clinic"],
    ["3 Jun 2026", "U.S. 510(k) filings — general + gynecologic", "Submissions to expand Hugo into general and gynecologic specialties in the U.S.",
     "https://news.medtronic.com/2026-06-03-Medtronic-submits-510-k-filings-to-expand-Hugo-TM-robotic-assisted-surgery-system-into-general-and-gynecologic-specialties-in-the-United-States"],
]

# ---------------------------------------------------------------------------
# Official totals over time
# ---------------------------------------------------------------------------
TOTALS = [
    ["Dec 2025", "Countries (ex-US)", "'More than 30 countries across 5 continents'",
     "https://news.medtronic.com/2025-12-03-Medtronic-announces-FDA-clearance-of-Hugo-TM-robotic-assisted-surgery-system-for-urologic-surgical-procedures"],
    ["Dec 2025", "Procedures (ex-US, cumulative)", "'Tens of thousands' of urologic, gynecologic, general surgery procedures",
     "https://news.medtronic.com/2025-12-03-Medtronic-announces-FDA-clearance-of-Hugo-TM-robotic-assisted-surgery-system-for-urologic-surgical-procedures"],
    ["Dec 2025", "Countries / experience", "'Nearly five years of commercial experience across more than 35 countries in five continents'",
     "https://www.prnewswire.com/news-releases/medtronic-announces-fda-clearance-of-hugo-robotic-assisted-surgery-system-for-urologic-surgical-procedures-302632328.html"],
    ["End of 2025 (third-party estimate)", "Installed systems", "~150 units; available in ~25 countries (Motley Fool analysis — NOT an official figure)",
     "https://www.fool.com/investing/2025/10/23/medtronic-is-diving-into-the-robotic-assisted-surg/"],
    ["2023–2024 (Medtronic material)", "Installed systems in Europe", "~48–60 Hugo systems installed in Europe",
     "(cited in regional press materials)"],
    ["AUA 2025", "Expand URO enrollment", "137 patients: 55 prostatectomy, 29 cystectomy, 53 nephrectomy",
     "https://www.urotoday.com/conference-highlights/aua-2025/"],
]

# ---------------------------------------------------------------------------
# Styling helpers
# ---------------------------------------------------------------------------
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, size=14, color="1F4E78")
SUBTITLE_FONT = Font(italic=True, size=10, color="595959")
WRAP = Alignment(wrap_text=True, vertical="top")
TOP = Alignment(vertical="top")
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

REGION_COLORS = {
    "United States": "FCE4D6",
    "Western Europe": "DDEBF7",
    "Eastern Europe": "E2EFDA",
    "Middle East": "FFF2CC",
    "Latin America": "FBE2D5",
    "Asia-Pacific": "EDE7F6",
}
CONF_COLORS = {"High": "C6EFCE", "Medium": "FFEB9C", "Low": "FFC7CE"}


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = BORDER


def set_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ---------------------------------------------------------------------------
# Build workbook
# ---------------------------------------------------------------------------
wb = Workbook()

# ----- Sheet 1: All Placements -----
ws = wb.active
ws.title = "All Placements"
ws["A1"] = "Medtronic Hugo™ RAS — Global Placement Tracker"
ws["A1"].font = TITLE_FONT
ws["A2"] = "Deep-research compilation as of 2026-06-05. Sources: Medtronic press releases, hospital newsrooms, local news, peer-reviewed reports. See 'Notes & Caveats'."
ws["A2"].font = SUBTITLE_FONT

headers = ["#", "Institution", "City", "State/Prov.", "Country", "Region",
           "First Case / Install Date", "Specialty", "Type", "Notable 'First'",
           "Confidence", "Source URL(s)", "Notes"]
hrow = 4
for c, h in enumerate(headers, start=1):
    ws.cell(row=hrow, column=c, value=h)
style_header(ws, hrow, len(headers))

for i, rec in enumerate(PLACEMENTS, start=1):
    r = hrow + i
    ws.cell(row=r, column=1, value=i)
    for c, val in enumerate(rec, start=2):
        ws.cell(row=r, column=c, value=val)
    region = rec[4]
    fill = REGION_COLORS.get(region)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=r, column=c)
        cell.alignment = WRAP
        cell.border = BORDER
        if fill and c in (5, 6):  # Country + Region tint
            cell.fill = PatternFill("solid", fgColor=fill)
    conf = rec[9]
    ccell = ws.cell(row=r, column=11)
    if conf.split()[0] in CONF_COLORS:
        ccell.fill = PatternFill("solid", fgColor=CONF_COLORS[conf.split()[0]])

set_widths(ws, [4, 38, 18, 12, 16, 16, 26, 28, 22, 30, 14, 50, 50])
ws.freeze_panes = "A5"
ws.auto_filter.ref = f"A{hrow}:{get_column_letter(len(headers))}{hrow + len(PLACEMENTS)}"
ws.row_dimensions[hrow].height = 30

# ----- Sheet 2: Summary by Country -----
ws2 = wb.create_sheet("Summary by Country")
ws2["A1"] = "Placements by Country & Region"
ws2["A1"].font = TITLE_FONT

# Aggregate
from collections import OrderedDict, Counter
country_region = OrderedDict()
country_counts = Counter()
region_counts = Counter()
for rec in PLACEMENTS:
    country, region = rec[3], rec[4]
    country_counts[country] += 1
    region_counts[region] += 1
    country_region[country] = region

hrow2 = 3
for c, h in enumerate(["Country", "Region", "Placements"], start=1):
    ws2.cell(row=hrow2, column=c, value=h)
style_header(ws2, hrow2, 3)
# sort by region then count desc
rows = sorted(country_counts.items(), key=lambda kv: (country_region[kv[0]], -kv[1], kv[0]))
r = hrow2
for country, cnt in rows:
    r += 1
    ws2.cell(row=r, column=1, value=country)
    ws2.cell(row=r, column=2, value=country_region[country])
    ws2.cell(row=r, column=3, value=cnt)
    fill = REGION_COLORS.get(country_region[country])
    for c in range(1, 4):
        cell = ws2.cell(row=r, column=c)
        cell.border = BORDER
        cell.alignment = TOP
        if fill:
            cell.fill = PatternFill("solid", fgColor=fill)
r += 1
ws2.cell(row=r, column=1, value="TOTAL").font = Font(bold=True)
ws2.cell(row=r, column=3, value=sum(country_counts.values())).font = Font(bold=True)

# Region summary block
r += 3
ws2.cell(row=r, column=1, value="By Region").font = Font(bold=True, size=12, color="1F4E78")
r += 1
for c, h in enumerate(["Region", "Placements"], start=1):
    ws2.cell(row=r, column=c, value=h)
style_header(ws2, r, 2)
for region, cnt in sorted(region_counts.items(), key=lambda kv: -kv[1]):
    r += 1
    ws2.cell(row=r, column=1, value=region)
    ws2.cell(row=r, column=2, value=cnt)
    fill = REGION_COLORS.get(region)
    for c in range(1, 3):
        cell = ws2.cell(row=r, column=c)
        cell.border = BORDER
        if fill:
            cell.fill = PatternFill("solid", fgColor=fill)
set_widths(ws2, [20, 18, 14])
ws2.freeze_panes = "A4"

# ----- Sheet 3: Milestones & Regulatory -----
ws3 = wb.create_sheet("Milestones & Regulatory")
ws3["A1"] = "Hugo RAS — Milestones & Regulatory Timeline"
ws3["A1"].font = TITLE_FONT
hrow3 = 3
for c, h in enumerate(["Date", "Milestone", "Detail", "Source URL"], start=1):
    ws3.cell(row=hrow3, column=c, value=h)
style_header(ws3, hrow3, 4)
for i, rec in enumerate(MILESTONES, start=1):
    r = hrow3 + i
    for c, val in enumerate(rec, start=1):
        cell = ws3.cell(row=r, column=c, value=val)
        cell.alignment = WRAP
        cell.border = BORDER
set_widths(ws3, [16, 38, 60, 60])
ws3.freeze_panes = "A4"

# ----- Sheet 4: Official Totals -----
ws4 = wb.create_sheet("Official Totals")
ws4["A1"] = "Hugo RAS — Official / Reported Totals Over Time"
ws4["A1"].font = TITLE_FONT
ws4["A2"] = "Medtronic does not publish a precise cumulative install or procedure count; figures are qualitative or third-party estimates."
ws4["A2"].font = SUBTITLE_FONT
hrow4 = 4
for c, h in enumerate(["As-of Date", "Metric", "Value", "Source URL"], start=1):
    ws4.cell(row=hrow4, column=c, value=h)
style_header(ws4, hrow4, 4)
for i, rec in enumerate(TOTALS, start=1):
    r = hrow4 + i
    for c, val in enumerate(rec, start=1):
        cell = ws4.cell(row=r, column=c, value=val)
        cell.alignment = WRAP
        cell.border = BORDER
set_widths(ws4, [28, 26, 55, 60])
ws4.freeze_panes = "A5"

# ----- Sheet 5: Notes & Caveats -----
ws5 = wb.create_sheet("Notes & Caveats")
ws5["A1"] = "Methodology, Scope & Caveats"
ws5["A1"].font = TITLE_FONT
notes = [
    "",
    "SUBJECT: Medtronic Hugo™ RAS (Robotic-Assisted Surgery) system. 'Placement' = a hospital/health-system/training-center where a Hugo system has been installed or used (incl. clinical-trial and demo/evaluation sites).",
    "",
    "COMPILED: 2026-06-05, by six parallel web-research agents (US; Western Europe; Eastern Europe/Middle East/Africa; Latin America; Asia-Pacific; global milestones).",
    "",
    "WHAT'S COUNTED: This workbook lists every distinct Hugo placement that could be verified from public sources. It is NOT exhaustive. Medtronic states Hugo is installed in '30–35+ countries on 5 continents' with an estimated ~150 systems (third-party) by end-2025 — far more than the named sites here. Many community/private hospitals adopted Hugo without notable press coverage and are therefore missing.",
    "",
    "TYPE definitions:",
    "  • Commercial — routine clinical use / hospital purchase.",
    "  • Clinical Trial — Hugo placed for an IDE/registered study (e.g., US Expand URO, Enable Hernia, Embrace Gynecology; HK thoracic study).",
    "  • Training Center — Medtronic Surgical Robotics Experience Center / simulation/training hub.",
    "  • Demo / Evaluation — system on-site for testing, purchase not confirmed.",
    "",
    "CONFIDENCE (color-coded in 'All Placements', col K):",
    "  • High (green) — named institution + date corroborated by Medtronic PR, hospital newsroom, or multiple independent sources.",
    "  • Medium (yellow) — institution confirmed but exact first-case/install date imprecise or from a single source / academic paper.",
    "  • Low (red) — placement plausible but not fully confirmed as Hugo (e.g., El Salvador – Hospital Rosales; AMOS clinic Australia).",
    "",
    "KEY 'FIRSTS':",
    "  • World's first Hugo surgery: Clínica Santa María, Santiago, Chile — 19 Jun 2021.",
    "  • World's first Hugo gynaecology: Hospital Pacífica Salud, Panama — Jul 2021.",
    "  • First in Europe: OLV Hospital, Aalst, Belgium — 2 Feb 2022.",
    "  • First in Asia-Pacific: Apollo Hospitals, Chennai, India — 17 Sep 2021.",
    "  • First in MENA: BDF Royal Medical Services, Bahrain — ~Aug 2024.",
    "  • First U.S. commercial surgery: Cleveland Clinic — 17 Feb 2026 (after FDA urology clearance 3 Dec 2025).",
    "",
    "KNOWN GAPS / NOT CONFIRMED (likely da Vinci or no public record): Argentina, Colombia, Peru and most of South America; Sweden, Austria, Ireland, Luxembourg, Iceland (Western Europe); Gulf states (Saudi, UAE, Qatar, Kuwait), Israel, Greece, the Balkans, Czechia/Slovakia/Hungary; all of Africa (Medtronic lists a South Africa product page but no named installs); mainland China, Singapore, Thailand, Malaysia, SE Asia.",
    "  • Roughly 3 of 6 US Expand URO urology sites, plus most Enable Hernia and Embrace Gynecology sites, are not publicly named.",
    "",
    "METHODOLOGY CAVEAT: Direct page fetches (WebFetch) and ClinicalTrials.gov returned HTTP 403 in this environment, so most dates/details derive from search-engine snippets cross-checked across sources. The dated press-release URLs are correct and citable; verify specific in-body figures against primary sources before publication.",
]
for i, line in enumerate(notes, start=2):
    cell = ws5.cell(row=i, column=1, value=line)
    cell.alignment = Alignment(wrap_text=True, vertical="top")
    if line and not line.startswith(" ") and line.isupper() is False and line.endswith(":"):
        cell.font = Font(bold=True)
    if line.startswith(("SUBJECT", "COMPILED", "WHAT'S COUNTED", "TYPE", "CONFIDENCE", "KEY", "KNOWN GAPS", "METHODOLOGY")):
        cell.font = Font(bold=True, color="1F4E78")
ws5.column_dimensions["A"].width = 130

# Save
out = "Hugo_RAS_Placements_Summary.xlsx"
wb.save(out)
print(f"Wrote {out} with {len(PLACEMENTS)} placements across {len(country_counts)} countries.")
print("Sheets:", wb.sheetnames)
