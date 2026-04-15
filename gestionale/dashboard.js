

// ═══════════════════════════════════════════════
// MAPPE
// ═══════════════════════════════════════════════
// Mappa codice → nome esatto cartella su Nextcloud /Clienti/
const CMAP={
  ERS:'Mavent (ERS-RSF)',RSF:'Mavent (ERS-RSF)',
  KFG:'KF Group (KFG)',
  TRM:'Trime (TRM)',
  ZMN:'Zeminian (ZMN)',
  NTM:'NTM Automazione',
  STP:'STP Scale (STP)',
  FIZ:'Tecnical-FIZ',TCN:'Tecnical-FIZ',
  MBE:'Mabe (MBE)',
  NVL:'Navigazione Laghi (NVL-NAV)',NAV:'Navigazione Laghi (NVL-NAV)',
  CSN:'Cason (CSN)',
  DNT:'Denitel (DNT)',
  FDC:'Fedco (FDC)',
  LGL:'LGL Engineering',
  MRM:'MRM',
  NVT:'Nuova Tecninox (NVT)',
  OAS:'OASI',OASI:'OASI',
  OTT:'Ottino (OTT)',
  CLG:'Icil Group (CLG)',
  CSC:'Condominio Santa Chiara (CSC)',
  RLN:'Orlandi (RLN)',
  TCF:'Tecfluid (TCF)',
  TBU:'Tabu (TBU)',
  TCC:'Tecnocamini (TCC)',
  TFC:'Trasp. F.Corretta (TFC)',
  UMA:'UMAC (UMA)',
  LMD:'La Maddalena (LMD)',
};
const CNAMES={ERS:'Mavent',RSF:'Mavent',KFG:'KF Group',TRM:'Trime',ZMN:'Zeminian',NTM:'NTM Autom.',STP:'STP Scale',FIZ:'Tecnical-FIZ',TCN:'Tecnical-FIZ',MBE:'Mabe',NVL:'Nav.Laghi',NAV:'Nav.Laghi',CSN:'Cason',DNT:'Denitel',FDC:'Fedco',LGL:'LGL Eng.',MRM:'MRM',NVT:'Nuova Tecninox',OASI:'OASI',OAS:'OASI',OTT:'Ottino',CLG:'Icil Group',CSC:'Cond.S.Chiara',RLN:'Orlandi',TCF:'Tecfluid',TBU:'Tabu',TCC:'Tecnocamini',TFC:'F.Corretta',UMA:'UMAC',LMD:'La Maddalena'};
const TIPI={ricerca:'Ricerca componenti',realizzazione:'Realizzazione centralina',prestazione:'Prestazione ore',configuratore:'Configuratore PU',misto:'Misto'};
const SL={bozza:'Bozza',offerta:'Offerta prep.',inviata:'Inviata',confermata:'Confermata','in-corso':'In lavorazione',chiusa:'Chiusa'};
const SLC=SL; // alias
const SAP={daordinare:'Da ordinare',ordinato:'Ordinato',arrivato:'Arrivato',problema:'Problema'};
const TIPO_C={ricerca:'tipo-ricerca',realizzazione:'tipo-realizzazione',prestazione:'tipo-prestazione',configuratore:'tipo-configuratore',misto:'tipo-misto'};

// ═══════════════════════════════════════════════
// LISTINO ARES — estratto da xlsx (prezzi 2022)
// ═══════════════════════════════════════════════
const ARES=[
  // ── Collettori PU05 (Gr.05 — minicentraline compatte) ──
  // Codice M1A = collettore monopompa con valvola VMC1, attacchi P-T G1/4"
  {s:'M1A-W',d:'PU05 Collettore M1A + VMC1-W 10-60 bar — G1/4" P-T Ø96',p:88.0,t:'PU05-collettore',serie:'PU05'},
  {s:'M1A-X',d:'PU05 Collettore M1A + VMC1-X 30-150 bar — G1/4" P-T Ø96',p:88.0,t:'PU05-collettore',serie:'PU05'},
  {s:'M1A-Y',d:'PU05 Collettore M1A + VMC1-Y 50-250 bar — G1/4" P-T Ø96',p:88.0,t:'PU05-collettore',serie:'PU05'},
  {s:'M1A-Z',d:'PU05 Collettore M1A + VMC1-Z 80-360 bar — G1/4" P-T Ø96',p:92.0,t:'PU05-collettore',serie:'PU05'},
  // Codice M2A = collettore M1A con doppio P (NOTA: non compatibile blocchi modulari)
  {s:'M2A-W',d:'PU05 Collettore M2A doppio P + VMC1-W 10-60 bar (NO blocchi modulari)',p:92.0,t:'PU05-collettore',serie:'PU05'},
  {s:'M2A-Y',d:'PU05 Collettore M2A doppio P + VMC1-Y 50-250 bar (NO blocchi modulari)',p:92.0,t:'PU05-collettore',serie:'PU05'},
  // Codice M1R = collettore reversibile con due ingressi P1-P2 G3/8"
  {s:'M1R-W',d:'PU05 Collettore reversibile M1R + VMC1-W 10-60 bar — P1-P2 G3/8"',p:145.0,t:'PU05-collettore',serie:'PU05'},
  {s:'M1R-Y',d:'PU05 Collettore reversibile M1R + VMC1-Y 50-250 bar — P1-P2 G3/8"',p:148.0,t:'PU05-collettore',serie:'PU05'},
  // ── Valvole screw-in PU05 (stesse cavità PU10 cav.13) ──
  {s:'NC-PU05',d:'PU05 Valvola NC pilotata VE6-NC — 350 bar 25 l/min cav.13',p:42.6,t:'valvola',serie:'PU05'},
  {s:'NCE-PU05',d:'PU05 Valvola NC pilotata + emergenza VE6-NC-EM — 350 bar 25 l/min',p:51.5,t:'valvola',serie:'PU05'},
  {s:'NAE-PU05',d:'PU05 Valvola NA pilotata + emergenza VE2-NA-EM — 250 bar 12 l/min',p:49.3,t:'valvola',serie:'PU05'},
  {s:'CDE-PU05',d:'PU05 Valvola NC diretta doppia tenuta VE2-NC-DT-EM — 210 bar 12 l/min',p:57.3,t:'valvola',serie:'PU05'},
  {s:'CPE-PU05',d:'PU05 Valvola proporzionale VE9-NC-EM — 300 bar 30 l/min (NH16)',p:136.9,t:'valvola',serie:'PU05'},
  {s:'PCD-PU05',d:'PU05 Valvola pneumatica doppia tenuta VP1-NC-DT — 300 bar 15 l/min',p:82.4,t:'valvola',serie:'PU05'},
  {s:'CM1A',d:'PU05 Valvola manuale 2 vie a cartuccia senza microswitch',p:28.5,t:'valvola',serie:'PU05'},
  {s:'CM1B',d:'PU05 Valvola manuale 2 vie a cartuccia con microswitch',p:34.0,t:'valvola',serie:'PU05'},
  {s:'TC1-PU05',d:'PU05 Tappo cavità cav.11/13',p:4.2,t:'valvola',serie:'PU05'},
  {s:'TC2-PU05',d:'PU05 Tappo cavità cav.13',p:4.2,t:'valvola',serie:'PU05'},
  {s:'TS1',d:'PU05 Scarico ausiliario 1/4" BSPP cav.13',p:8.5,t:'valvola',serie:'PU05'},
  {s:'TM1',d:'PU05 Mandata ausiliaria 1/4" BSPP cav.13',p:8.5,t:'valvola',serie:'PU05'},
  {s:'VMC1-W',d:'PU05 Valvola massima VMC1-W 10-60 bar cav.11',p:16.8,t:'valvola',serie:'PU05'},
  {s:'VMC1-X',d:'PU05 Valvola massima VMC1-X 30-150 bar cav.11',p:16.8,t:'valvola',serie:'PU05'},
  {s:'VMC1-Y',d:'PU05 Valvola massima VMC1-Y 50-250 bar cav.11',p:16.8,t:'valvola',serie:'PU05'},
  {s:'VMC1-Z',d:'PU05 Valvola massima VMC1-Z 80-360 bar cav.11',p:18.0,t:'valvola',serie:'PU05'},
  {s:'VRF1R-PU05',d:'PU05 Regolatrice flusso compensata regolabile 2-20 l/min 350 bar',p:68.5,t:'valvola',serie:'PU05'},
  {s:'RFR1',d:'PU05 Valvola bidirezionale controllo flusso 300 bar 30 l/min',p:52.0,t:'valvola',serie:'PU05'},
  {s:'VR14',d:'PU05 Ritegno 1/4" BSPP cav.17',p:9.5,t:'valvola',serie:'PU05'},
  {s:'VR38',d:'PU05 Ritegno 3/8" BSPP cav.18',p:11.2,t:'valvola',serie:'PU05'},
  {s:'VRP38',d:'PU05 Ritegno pilotato 3/8" BSPP — rapporto 1:4 cav.19',p:42.0,t:'valvola',serie:'PU05'},
  // ── Bobine per PU05 (N-H13, N-H13R, N-H16) ──
  {s:'OA-PU05',d:'PU05 Bobina NH13 12 Vdc — 18W 100% IP65 DIN43650',p:12.0,t:'bobina',serie:'PU05'},
  {s:'OB-PU05',d:'PU05 Bobina NH13 24 Vdc — 18W 100% IP65 DIN43650',p:12.0,t:'bobina',serie:'PU05'},
  {s:'OC-PU05',d:'PU05 Bobina NH13 48 Vdc',p:14.8,t:'bobina',serie:'PU05'},
  {s:'OL-PU05',d:'PU05 Bobina NH13 24 Vac 50Hz',p:12.0,t:'bobina',serie:'PU05'},
  {s:'ON-PU05',d:'PU05 Bobina NH13 220 Vac 50Hz',p:12.5,t:'bobina',serie:'PU05'},
  // ── Serbatoi PU05 (SPM serie — plastica HDPE+PP) ──
  {s:'SPM01D',d:'PU05 Serbatoio plastica 1L orizzontale 135×158×134 Ø96',p:12.0,t:'serbatoio',serie:'PU05'},
  {s:'SPM02D',d:'PU05 Serbatoio plastica 1.8L orizzontale 180×158×134 Ø96',p:13.5,t:'serbatoio',serie:'PU05'},
  {s:'SPM03D',d:'PU05 Serbatoio plastica 2.5L orizzontale 240×158×134 Ø96',p:14.5,t:'serbatoio',serie:'PU05'},
  {s:'SPM04D',d:'PU05 Serbatoio plastica 3.5L orizzontale 280×158×134 Ø96',p:15.5,t:'serbatoio',serie:'PU05'},
  {s:'SPM05D',d:'PU05 Serbatoio plastica 4L orizzontale 330×158×134 Ø96',p:16.5,t:'serbatoio',serie:'PU05'},
  {s:'SPM01V',d:'PU05 Serbatoio plastica 1L verticale 135×174×134 Ø96',p:12.0,t:'serbatoio',serie:'PU05'},
  {s:'SPM03V',d:'PU05 Serbatoio plastica 2.5L verticale 240×174×134 Ø96',p:14.5,t:'serbatoio',serie:'PU05'},
  {s:'SPM05V',d:'PU05 Serbatoio plastica 4L verticale 330×174×134 Ø96',p:16.5,t:'serbatoio',serie:'PU05'},
  {s:'SPM03H',d:'PU05 Serbatoio plastica 1.3L orizzontale compact 218×115×103 Ø96',p:11.0,t:'serbatoio',serie:'PU05'},
  {s:'SPM11H',d:'PU05 Serbatoio plastica 2L orizzontale compact 336×115×103 Ø96',p:13.5,t:'serbatoio',serie:'PU05'},
  // ── Accessori PU05 ──
  {s:'KS3',d:'PU05 Kit fissaggio serbatoio SPM (tutte le serie) — vite M6',p:5.0,t:'accessorio',serie:'PU05'},
  {s:'FS0',d:'PU05 Adattatore per serbatoi in lamiera Ø123 su collettore PU05',p:18.0,t:'accessorio',serie:'PU05'},
  {s:'FC11',d:'PU05 Giunto flangia motori DC C102-C203 Ø71 L=15.8',p:18.5,t:'giunto',serie:'PU05'},
  {s:'FC12',d:'PU05 Giunto flangia motori DC C104-C205-C211-C401 Ø71 L=23.5',p:20.5,t:'giunto',serie:'PU05'},
  {s:'FAM56',d:'PU05 Flangia motori AC MEC56 (M298/M498)',p:28.0,t:'giunto',serie:'PU05'},
  {s:'FAM63',d:'PU05 Flangia motori AC MEC63 (M299/M200/M499/M400)',p:32.0,t:'giunto',serie:'PU05'},
  {s:'FAM71',d:'PU05 Flangia motori AC MEC71 (M201/M202/M401/M402)',p:35.0,t:'giunto',serie:'PU05'},
  // ── Motori DC PU05 (da catalogo sezione 5.1) ──
  {s:'C102',d:'PU05 Motore DC 12V 500W S2=5min S3=17% IP44',p:68.0,t:'motore',serie:'PU05'},
  {s:'C103',d:'PU05 Motore DC 12V 800W S2=4min S3=9% IP44',p:76.0,t:'motore',serie:'PU05'},
  {s:'C108',d:'PU05 Motore DC 12V 500W S2=4min S3=15% IP44',p:70.0,t:'motore',serie:'PU05'},
  {s:'C113',d:'PU05 Motore DC 12V 800W IP44 con termica',p:80.0,t:'motore',serie:'PU05'},
  {s:'C202',d:'PU05 Motore DC 24V 500W S2=5min S3=17% IP44',p:68.0,t:'motore',serie:'PU05'},
  {s:'C203',d:'PU05 Motore DC 24V 800W S2=2.5min S3=8% IP44',p:76.0,t:'motore',serie:'PU05'},
  {s:'C208',d:'PU05 Motore DC 24V 500W S2=4min S3=15% IP44',p:70.0,t:'motore',serie:'PU05'},
  {s:'C213',d:'PU05 Motore DC 24V 800W IP44 con termica',p:80.0,t:'motore',serie:'PU05'},
  {s:'C104',d:'PU05 Motore DC 12V 1600W S2=2min S3=10% IP54',p:138.0,t:'motore',serie:'PU05'},
  {s:'C105',d:'PU05 Motore DC 12V 1600W IP54 con termica',p:144.0,t:'motore',serie:'PU05'},
  {s:'C204',d:'PU05 Motore DC 24V 2200W S2=2min S3=5% IP54',p:155.0,t:'motore',serie:'PU05'},
  {s:'C205',d:'PU05 Motore DC 24V 2200W IP54 con termica',p:162.0,t:'motore',serie:'PU05'},
  {s:'C111',d:'PU05 Motore DC 12V 1500W IP54 con termica Ø114',p:142.0,t:'motore',serie:'PU05'},
  {s:'C211',d:'PU05 Motore DC 24V 2000W IP54 con termica Ø114',p:158.0,t:'motore',serie:'PU05'},
  {s:'C401',d:'PU05 Motore DC 48V 2000W S2=2min IP54',p:158.0,t:'motore',serie:'PU05'},
  {s:'C402',d:'PU05 Motore DC 48V 2000W S2=3min S3=12% IP54 con termica',p:164.0,t:'motore',serie:'PU05'},
  {s:'C130',d:'PU05 Motore DC 12V 1500W ventilato S2=4min IP20',p:148.0,t:'motore',serie:'PU05'},
  {s:'C230',d:'PU05 Motore DC 24V 2000W ventilato S2=5min IP20',p:158.0,t:'motore',serie:'PU05'},
  // ── Motori AC PU05 ──
  {s:'M298',d:'PU05 Motore AC 1-fase 0.12kW 2900rpm MEC56 B14',p:95.0,t:'motore',serie:'PU05'},
  {s:'M299',d:'PU05 Motore AC 1-fase 0.18kW 2900rpm MEC63 B14',p:102.0,t:'motore',serie:'PU05'},
  {s:'M200',d:'PU05 Motore AC 1-fase 0.25kW 2900rpm MEC63 B14',p:108.0,t:'motore',serie:'PU05'},
  {s:'M201',d:'PU05 Motore AC 1-fase 0.37kW 2900rpm MEC71 B14',p:115.0,t:'motore',serie:'PU05'},
  {s:'M202',d:'PU05 Motore AC 1-fase 0.55kW 2900rpm MEC71 B14',p:124.0,t:'motore',serie:'PU05'},
  {s:'T297',d:'PU05 Motore AC 3-fase 0.09kW 2900rpm MEC56 B14',p:82.0,t:'motore',serie:'PU05'},
  {s:'T298',d:'PU05 Motore AC 3-fase 0.12kW 2900rpm MEC56 B14',p:88.0,t:'motore',serie:'PU05'},
  {s:'T299',d:'PU05 Motore AC 3-fase 0.18kW 2900rpm MEC63 B14',p:95.0,t:'motore',serie:'PU05'},
  {s:'T200',d:'PU05 Motore AC 3-fase 0.25kW 2900rpm MEC63 B14',p:102.0,t:'motore',serie:'PU05'},
  {s:'T201',d:'PU05 Motore AC 3-fase 0.37kW 2900rpm MEC71 B14',p:108.0,t:'motore',serie:'PU05'},
  {s:'T202',d:'PU05 Motore AC 3-fase 0.55kW 2900rpm MEC71 B14',p:115.0,t:'motore',serie:'PU05'},
  // ── Relay avviamento e coperchio protezione PU05 ──
  {s:'FC-B',d:'PU05 Relay avviamento 12V 200A/350A max',p:28.0,t:'accessorio',serie:'PU05'},
  {s:'FC-D',d:'PU05 Relay avviamento 24V 200A/350A max',p:28.0,t:'accessorio',serie:'PU05'},
  {s:'G00',d:'PU05/PU10 Senza supporto montaggio',p:0.0,t:'accessorio',serie:'PU05'},
  {s:'G01',d:'PU05/PU10 Supporto montaggio standard L=150 H=87',p:18.5,t:'accessorio',serie:'PU05'},
  // Posizioni montaggio
  {s:'1H',d:'PU05 Posizione montaggio 1H (serbatoio sotto, motore lato)',p:0.0,t:'accessorio',serie:'PU05'},
  {s:'1V',d:'PU05 Posizione montaggio 1V (verticale)',p:0.0,t:'accessorio',serie:'PU05'},
  // ── Blocchi modulari PU05 (stessa interfaccia PU10) ──
  {s:'B09-PU05',d:'PU05 Distanziale H=18 pass-through P-T',p:19.1,t:'blocco',serie:'PU05'},
  {s:'B01-PU05',d:'PU05 Distanziale H=39',p:25.6,t:'blocco',serie:'PU05'},
  {s:'B92-PU05',d:'PU05 Distanziale H=25',p:23.6,t:'blocco',serie:'PU05'},
  {s:'B15-PU05',d:'PU05 Adattatore rotaz. lato motore H=90 (M 1/4" BSPP)',p:102.5,t:'blocco',serie:'PU05'},
  {s:'B03-PU05',d:'PU05 Blocco CETOP3-NG6 parallelo/serie 1/4"',p:39.2,t:'blocco',serie:'PU05'},
  {s:'B11-PU05',d:'PU05 Blocco CETOP3-NG6 parallelo/serie 3/8"',p:39.2,t:'blocco',serie:'PU05'},
  {s:'B73-PU05',d:'PU05 Blocco riduttrice di pressione 350 bar',p:297.0,t:'blocco',serie:'PU05'},
  {s:'B259A-PU05',d:'PU05 Blocco valvola scarico automatica tipo A',p:134.8,t:'blocco',serie:'PU05'},
  {s:'B38-PU05',d:'PU05 Blocco filtro sullo scarico 15 micron max 20 l/min',p:78.5,t:'blocco',serie:'PU05'},
  // ── Collettori PU10 (Gr.1) ──
  {s:'P002',d:'PU05 Pompa Gr.05 cil. 0.25 cc/giro — 200 bar 0.36 l/min @1500',p:58.0,t:'pompa',serie:'PU05'},
  {s:'P004',d:'PU05 Pompa Gr.05 cil. 0.50 cc/giro — 200 bar 0.72 l/min @1500',p:60.0,t:'pompa',serie:'PU05'},
  {s:'P006',d:'PU05 Pompa Gr.05 cil. 0.75 cc/giro — 200 bar 1.08 l/min @1500',p:62.0,t:'pompa',serie:'PU05'},
  {s:'P008',d:'PU05 Pompa Gr.05 cil. 1.00 cc/giro — 200 bar 1.45 l/min @1500',p:64.0,t:'pompa',serie:'PU05'},
  {s:'P009',d:'PU05 Pompa Gr.05 cil. 1.25 cc/giro — 200 bar 1.80 l/min @1500',p:66.0,t:'pompa',serie:'PU05'},
  {s:'P010',d:'PU05 Pompa Gr.05 cil. 1.50 cc/giro — 175 bar 2.18 l/min @1500',p:68.0,t:'pompa',serie:'PU05'},
  {s:'P012',d:'PU05 Pompa Gr.05 cil. 2.00 cc/giro — 160 bar 2.88 l/min @1500',p:72.0,t:'pompa',serie:'PU05'},
  {s:'KP1',d:'PU05 Adattatore montaggio pompa Gr.05 su collettore Gr.1',p:18.5,t:'accessorio',serie:'PU05'},
  // Motori DC per PU05 (stessa gamma PU10 ma versioni più piccole)
  {s:'C102',d:'PU05 Motore DC 12V 500W 2500rpm IP44 reversibile',p:72.0,t:'motore',serie:'PU05'},
  {s:'C202',d:'PU05 Motore DC 24V 500W 2800rpm IP44 reversibile',p:72.0,t:'motore',serie:'PU05'},
  {s:'C103',d:'PU05 Motore DC 12V 800W 3500rpm IP44 reversibile',p:82.0,t:'motore',serie:'PU05'},
  {s:'C203',d:'PU05 Motore DC 24V 800W 4000rpm IP44 reversibile',p:82.0,t:'motore',serie:'PU05'},
  {s:'C114',d:'PU05 Motore DC 12V 500W IP44 reversibile con termica',p:76.0,t:'motore',serie:'PU05'},
  {s:'C214',d:'PU05 Motore DC 24V 500W IP44 reversibile con termica',p:76.0,t:'motore',serie:'PU05'},
  // Motori AC per PU05
  {s:'M204',d:'PU05 Motore AC 1-fase 0.37kW 2900rpm MEC71 B14',p:115.0,t:'motore',serie:'PU05'},
  {s:'M205',d:'PU05 Motore AC 1-fase 0.55kW 2900rpm MEC71 B14',p:122.0,t:'motore',serie:'PU05'},
  {s:'T204',d:'PU05 Motore AC 3-fase 0.37kW 2900rpm MEC71 B14',p:108.0,t:'motore',serie:'PU05'},
  {s:'T205',d:'PU05 Motore AC 3-fase 0.55kW 2900rpm MEC71 B14',p:115.0,t:'motore',serie:'PU05'},
  // Giunti/flange PU05
  {s:'FC05',d:'PU05 Giunto per motori DC C102/C103/C202/C203 Gr.05',p:20.0,t:'giunto',serie:'PU05'},
  {s:'FC06',d:'PU05 Giunto per motori DC C104/C204/C205 Gr.05',p:22.0,t:'giunto',serie:'PU05'},
  {s:'FA72',d:'PU05 Flangia per motori AC MEC71 Gr.05',p:28.0,t:'giunto',serie:'PU05'},
  {s:'FA82',d:'PU05 Flangia per motori AC MEC80 Gr.05',p:32.0,t:'giunto',serie:'PU05'},
  // Serbatoi PU05 (stessi del PU10 piccoli)
  {s:'SL01H',d:'PU05 Serbatoio lamiera 1L orizzontale Ø123',p:18.5,t:'serbatoio',serie:'PU05'},
  {s:'SL03H',d:'PU05 Serbatoio lamiera 2L orizzontale Ø123',p:22.0,t:'serbatoio',serie:'PU05'},
  {s:'SL04H',d:'PU05 Serbatoio lamiera 3L orizzontale Ø123',p:25.0,t:'serbatoio',serie:'PU05'},
  {s:'SP01D',d:'PU05 Serbatoio plastica 1L orizzontale Ø123 trasparente',p:12.0,t:'serbatoio',serie:'PU05'},
  {s:'SP03D',d:'PU05 Serbatoio plastica 2.5L orizzontale Ø123 trasparente',p:14.5,t:'serbatoio',serie:'PU05'},
  {s:'SP05D',d:'PU05 Serbatoio plastica 4L orizzontale Ø123 trasparente',p:16.5,t:'serbatoio',serie:'PU05'},
  // ── Collettori PU10 (Gr.1) ──
  {s:'A1A',d:'PU10 Collettore monopompa VMC1 — attacchi P-T G1/4"',p:263.8,t:'PU10-collettore',serie:'PU10'},
  {s:'A1B',d:'PU10 Collettore + valvola NC pilotata integrata',p:278.5,t:'PU10-collettore',serie:'PU10'},
  {s:'A1C',d:'PU10 Collettore con ritegno su A',p:266.4,t:'PU10-collettore',serie:'PU10'},
  {s:'A1D',d:'PU10 Collettore doppio effetto — attacchi A-T G1/4"',p:269.7,t:'PU10-collettore',serie:'PU10'},
  {s:'B2A',d:'PU10 Collettore VML1 semplice effetto (P+2 utenze)',p:278.5,t:'PU10-collettore',serie:'PU10'},
  {s:'B3B',d:'PU10 Collettore VML1 doppio effetto A-B',p:278.5,t:'PU10-collettore',serie:'PU10'},
  {s:'B4A',d:'PU10 Collettore VML1 4 vie A-B con check',p:286.0,t:'PU10-collettore',serie:'PU10'},
  {s:'R1A',d:'PU10 Collettore pompa reversibile (P1-P2 G3/8")',p:537.2,t:'PU10-collettore',serie:'PU10'},
  {s:'D1A',d:'PU10 Collettore doppia pompa 2×VMC1',p:978.1,t:'PU10-collettore',serie:'PU10'},
  {s:'D2A',d:'PU10 Collettore doppia pompa 2×VMC1 alternativo',p:992.0,t:'PU10-collettore',serie:'PU10'},
  // ── Collettori PU20 (Gr.2) ──
  {s:'Z1A',d:'PU20 Collettore monopompa VMZ1 — attacchi P G3/8" T G1/2"',p:574.6,t:'PU20-collettore',serie:'PU20'},
  {s:'Z2A',d:'PU20 Collettore doppia pompa VMZ1',p:988.6,t:'PU20-collettore',serie:'PU20'},
  // ── Valvole screw-in PU10 ──
  {s:'NC',d:'Valvola NC pilotata VE6-NC — 350 bar 25 l/min (NH13)',p:42.6,t:'valvola',serie:'PU10'},
  {s:'NCE',d:'Valvola NC pilotata + emergenza VE6-NC-EM',p:51.5,t:'valvola',serie:'PU10'},
  {s:'NAE',d:'Valvola NA pilotata + emergenza VE2-NA-EM — 250 bar 12 l/min',p:49.3,t:'valvola',serie:'PU10'},
  {s:'CDE',d:'Valvola NC diretta doppia tenuta + emergenza VE2-NC-DT-EM',p:57.3,t:'valvola',serie:'PU10'},
  {s:'CPE',d:'Valvola proporzionale NC pilotata VE9-NC-EM — 300 bar 30 l/min (NH16)',p:136.9,t:'valvola',serie:'PU10'},
  {s:'PCD',d:'Valvola pneumatica doppia tenuta VP1-NC-DT — 300 bar 15 l/min',p:82.4,t:'valvola',serie:'PU10'},
  {s:'E43A',d:'Valvola 4 vie 3 posizioni a cursore V43S13A',p:98.5,t:'valvola',serie:'PU10'},
  {s:'E43B',d:'Valvola 4 vie 3 posizioni a cursore V43S13B',p:98.5,t:'valvola',serie:'PU10'},
  {s:'E42B',d:'Valvola 4 vie 3 posizioni a cursore V42S13B',p:98.5,t:'valvola',serie:'PU10'},
  // ── Bobine ──
  {s:'OA',d:'Bobina NH13 12 Vdc — 18W IP65',p:12.0,t:'bobina',serie:'PU10'},
  {s:'OB',d:'Bobina NH13 24 Vdc — 18W IP65',p:12.0,t:'bobina',serie:'PU10'},
  {s:'OC',d:'Bobina NH13 48 Vdc',p:14.8,t:'bobina',serie:'PU10'},
  {s:'OL',d:'Bobina NH13 24 Vac 50Hz',p:12.0,t:'bobina',serie:'PU10'},
  {s:'ON',d:'Bobina NH13 220 Vac 50Hz',p:12.5,t:'bobina',serie:'PU10'},
  // ── Pompe Gr.1 (PU10) ──
  {s:'PG07',d:'Pompa Gr.1 cilindrata 3.15 cc/giro — 250 bar 4.44 l/min @1500',p:98.0,t:'pompa',serie:'PU10'},
  {s:'PG05',d:'Pompa Gr.1 cilindrata 2.50 cc/giro — 250 bar 3.53 l/min @1500',p:92.0,t:'pompa',serie:'PU10'},
  {s:'PG04',d:'Pompa Gr.1 cilindrata 2.00 cc/giro — 250 bar 2.82 l/min @1500',p:88.0,t:'pompa',serie:'PU10'},
  {s:'PG09',d:'Pompa Gr.1 cilindrata 4.20 cc/giro — 250 bar 5.92 l/min @1500',p:104.0,t:'pompa',serie:'PU10'},
  {s:'PG11',d:'Pompa Gr.1 cilindrata 5.00 cc/giro — 250 bar 7.05 l/min @1500',p:110.0,t:'pompa',serie:'PU10'},
  {s:'PG12',d:'Pompa Gr.1 cilindrata 5.70 cc/giro — 200 bar 8.12 l/min @1500',p:115.0,t:'pompa',serie:'PU10'},
  // ── Pompe Gr.2 (PU20) ──
  {s:'P204',d:'Pompa Gr.2 cilindrata 8.2 cc/giro — 250 bar 11.32 l/min @1500',p:145.0,t:'pompa',serie:'PU20'},
  {s:'P208',d:'Pompa Gr.2 cilindrata 14 cc/giro — 250 bar 19.95 l/min @1500',p:162.0,t:'pompa',serie:'PU20'},
  {s:'P212',d:'Pompa Gr.2 cilindrata 19 cc/giro — 200 bar 27.36 l/min @1500',p:178.0,t:'pompa',serie:'PU20'},
  {s:'P214',d:'Pompa Gr.2 cilindrata 22 cc/giro — 180 bar 31.68 l/min @1500',p:185.0,t:'pompa',serie:'PU20'},
  // ── Serbatoi (PU10) ──
  {s:'SL47H',d:'Serbatoio lamiera 7L orizzontale Ø123 — H/V',p:38.5,t:'serbatoio',serie:'PU10'},
  {s:'SL44H',d:'Serbatoio lamiera 4L orizzontale Ø123',p:31.2,t:'serbatoio',serie:'PU10'},
  {s:'SL48H',d:'Serbatoio lamiera 8L orizzontale Ø123',p:42.0,t:'serbatoio',serie:'PU10'},
  {s:'SP07H',d:'Serbatoio plastica 3L orizzontale Ø123',p:22.4,t:'serbatoio',serie:'PU10'},
  {s:'SP08H',d:'Serbatoio plastica 5L orizzontale Ø123',p:26.8,t:'serbatoio',serie:'PU10'},
  // ── Serbatoi (PU20) ──
  {s:'SZ23V',d:'Serbatoio PU20 lamiera 30L verticale Ø175',p:88.0,t:'serbatoio',serie:'PU20'},
  {s:'SZ24V',d:'Serbatoio PU20 lamiera 45L verticale Ø175',p:102.0,t:'serbatoio',serie:'PU20'},
  {s:'SZ22H',d:'Serbatoio PU20 lamiera 30L orizzontale Ø175',p:88.0,t:'serbatoio',serie:'PU20'},
  // ── Motori DC (PU10) ──
  {s:'C204',d:'Motore DC 24V 2200W 2600rpm IP54 — non reversibile',p:162.0,t:'motore',serie:'PU10'},
  {s:'C205',d:'Motore DC 24V 2200W 2600rpm IP54 con termica',p:168.0,t:'motore',serie:'PU10'},
  {s:'C104',d:'Motore DC 12V 1600W 2600rpm IP54',p:148.0,t:'motore',serie:'PU10'},
  {s:'C202',d:'Motore DC 24V 500W 2800rpm IP44 reversibile',p:88.0,t:'motore',serie:'PU10'},
  // ── Motori AC (PU10) ──
  {s:'T206',d:'Motore AC 3-fase 0.75kW 2900rpm MEC80 B14',p:145.0,t:'motore',serie:'PU10'},
  {s:'T207',d:'Motore AC 3-fase 1.1kW 2900rpm MEC80 B14',p:158.0,t:'motore',serie:'PU10'},
  {s:'T208',d:'Motore AC 3-fase 1.5kW 2900rpm MEC90 B14',p:175.0,t:'motore',serie:'PU10'},
  {s:'M206',d:'Motore AC 1-fase 0.75kW 2900rpm MEC80 B14',p:155.0,t:'motore',serie:'PU10'},
  {s:'M207',d:'Motore AC 1-fase 1.1kW 2900rpm MEC80 B14',p:168.0,t:'motore',serie:'PU10'},
  // ── Motori AC (PU20) ──
  {s:'T258',d:'Motore AC 3-fase 7.5kW 3000rpm MEC132 B5',p:385.0,t:'motore',serie:'PU20'},
  {s:'T255',d:'Motore AC 3-fase 3.0kW 3000rpm MEC100 B5',p:265.0,t:'motore',serie:'PU20'},
  {s:'T256',d:'Motore AC 3-fase 4.0kW 3000rpm MEC112 B5',p:298.0,t:'motore',serie:'PU20'},
  // ── Giunti/Flange (PU10) ──
  {s:'FC02',d:'Giunto per motori DC C104-C204-C205-C215 Gr.1',p:28.5,t:'giunto',serie:'PU10'},
  {s:'FC01',d:'Giunto per motori DC C102-C103-C202 Gr.1',p:22.4,t:'giunto',serie:'PU10'},
  {s:'FA80',d:'Flangia per motori AC MEC80 Gr.1',p:35.2,t:'giunto',serie:'PU10'},
  {s:'FA90',d:'Flangia per motori AC MEC90 Gr.1',p:38.8,t:'giunto',serie:'PU10'},
  // ── Giunti/Flange (PU20) ──
  {s:'FZ132',d:'Flangia+giunto PU20 per motori MEC132 B5',p:128.0,t:'giunto',serie:'PU20'},
  {s:'FZ100',d:'Flangia+giunto PU20 per motori MEC110/112 B5',p:115.0,t:'giunto',serie:'PU20'},
  // ── Blocchi modulari (comuni) ──
  {s:'B09',d:'Distanziale H=18 — pass-through P-T',p:19.1,t:'blocco',serie:'PU10'},
  {s:'B01',d:'Distanziale H=39',p:25.6,t:'blocco',serie:'PU10'},
  {s:'B92',d:'Distanziale H=25',p:23.6,t:'blocco',serie:'PU10'},
  {s:'B02',d:'Distanziale H=69',p:32.0,t:'blocco',serie:'PU10'},
  {s:'B15',d:'Adattatore rotazione lato motore H=90 (M 1/4" BSPP)',p:102.5,t:'blocco',serie:'PU10'},
  {s:'B51',d:'Adattatore rotazione lato motore H=60',p:48.3,t:'blocco',serie:'PU10'},
  {s:'B26',d:'Adattatore rotazione lato serbatoio H=60',p:48.3,t:'blocco',serie:'PU10'},
  {s:'B76',d:'Adattatore rotazione lato serbatoio H=90',p:55.0,t:'blocco',serie:'PU10'},
  {s:'B33',d:'Adattatore rotazione 90°',p:48.3,t:'blocco',serie:'PU10'},
  {s:'B128',d:'Blocco 4 attacchi P supplementari M 1/4"',p:38.3,t:'blocco',serie:'PU10'},
  {s:'B96',d:'Blocco attacchi P-T 3/8" + M 1/4" BSPP',p:40.5,t:'blocco',serie:'PU10'},
  {s:'B163',d:'Blocco attacchi P-T 1/2" + M 1/4" BSPP',p:45.3,t:'blocco',serie:'PU10'},
  {s:'B03',d:'Blocco CETOP3-NG6 parallelo/serie — attacchi 1/4"',p:39.2,t:'blocco',serie:'PU10'},
  {s:'B11',d:'Blocco CETOP3-NG6 parallelo/serie — attacchi 3/8"',p:39.2,t:'blocco',serie:'PU10'},
  {s:'B85',d:'Blocco CETOP3-NG6 parallelo con uscite 1/4" BSPP',p:45.0,t:'blocco',serie:'PU10'},
  {s:'B37',d:'Blocco CETOP5 parallelo (PU20)',p:68.0,t:'blocco',serie:'PU20'},
  {s:'B30_06',d:'Divisore flusso 50/50 — 1-6 l/min',p:251.1,t:'blocco',serie:'PU10'},
  {s:'B30_10',d:'Divisore flusso 50/50 — 5-10 l/min',p:251.1,t:'blocco',serie:'PU10'},
  {s:'B30_20',d:'Divisore flusso 50/50 — 10-20 l/min',p:251.1,t:'blocco',serie:'PU10'},
  {s:'B30_40',d:'Divisore flusso 50/50 — 20-40 l/min',p:268.0,t:'blocco',serie:'PU10'},
  {s:'B41',d:'Blocco valvola prioritaria compensata — 350 bar 50 l/min',p:206.6,t:'blocco',serie:'PU10'},
  {s:'B73',d:'Blocco riduttrice di pressione — 350 bar 30 l/min',p:297.0,t:'blocco',serie:'PU10'},
  {s:'B259A',d:'Blocco valvola scarico automatica tipo A',p:134.8,t:'blocco',serie:'PU10'},
  {s:'B259B',d:'Blocco valvola scarico automatica tipo B',p:134.8,t:'blocco',serie:'PU10'},
  {s:'B38',d:'Blocco filtro sullo scarico 20 micron — max 20 l/min',p:78.5,t:'blocco',serie:'PU10'},
  {s:'B39_25',d:'Blocco filtro sulla mandata 25 micron — max 10 l/min',p:82.0,t:'blocco',serie:'PU10'},
  {s:'B20',d:'Pompa a mano modulare 6.5cc — max 350 bar',p:112.0,t:'blocco',serie:'PU10'},
  // ── Valvole varie / accessori ──
  {s:'VRF1R',d:'Regolatrice flusso compensata regolabile 2-20 l/min 350 bar',p:68.5,t:'valvola',serie:'PU10'},
  {s:'VU1',d:'Ritegno a cartuccia — 3/4" 16 UNF cav.12',p:18.2,t:'valvola',serie:'PU10'},
  {s:'VMC1-Y',d:'Valvola di massima VMC1 taratura 50-250 bar cav.11',p:24.8,t:'valvola',serie:'PU10'},
  {s:'VMZ1-Y',d:'Valvola di massima VMZ1 taratura 50-220 bar (PU20)',p:38.5,t:'valvola',serie:'PU20'},
  {s:'TC1',d:'Tappo cavità cav.05/11/13',p:4.2,t:'valvola',serie:'PU10'},
  {s:'TC2',d:'Tappo cavità cav.04/05/10/12',p:4.2,t:'valvola',serie:'PU10'},
  {s:'G01',d:'Supporto di montaggio standard',p:18.5,t:'accessorio',serie:'PU10'},
];

function aresSearch(q,serie=''){
  if(!q||q.length<2)return[];
  const r=q.toLowerCase();
  return ARES.filter(a=>{
    const match=a.s.toLowerCase().includes(r)||a.d.toLowerCase().includes(r);
    const matchSerie=!serie||a.serie===serie||a.t==='bobina'||a.t==='accessorio';
    return match&&matchSerie;
  }).slice(0,8);
}
const TIPO_BADGE={
  'PU05-collettore':'background:#fef3c7;color:#92400e;border:1px solid #f59e0b',
  'PU10-collettore':'background:#dcfce7;color:#166534;border:1px solid #6ee7a0',
  'PU20-collettore':'background:#dbeafe;color:#1e40af;border:1px solid #7fb3f5',
  'valvola':'background:#ede9fe;color:#5b21b6;border:1px solid #a78bfa',
  'bobina':'background:#fef3c7;color:#78350f;border:1px solid #f59e0b',
  'pompa':'background:#ccfbf1;color:#115e59;border:1px solid #2dd4bf',
  'serbatoio':'background:#f3f4f6;color:#374151;border:1px solid #d1d5db',
  'motore':'background:#fee2e2;color:#991b1b;border:1px solid #f87171',
  'giunto':'background:#fef3c7;color:#92400e;border:1px solid #fbbf24',
  'blocco':'background:#f0f9ff;color:#0369a1;border:1px solid #7dd3fc',
  'accessorio':'background:#f9fafb;color:#6b7280;border:1px solid #d1d5db',
};

// ═══════════════════════════════════════════════
// STORAGE
// ═══════════════════════════════════════════════
const CK='asg_config';
function lCfg(){try{return JSON.parse(localStorage.getItem(CK)||'{}');}catch{return{};}}
function sCfg(c){localStorage.setItem(CK,JSON.stringify(c));}
// Legge valore da select nativo O da custom dropdown
function getElVal(id){
  const el=document.getElementById(id);
  if(!el)return'';
  if(el.classList&&el.classList.contains('csel'))return window._cselData?.[id]?.val||el.dataset.value||'';
  return el.value||'';
}

function getCfg(){const c=lCfg();return{nc_url:c.nc_url||'',nc_user:c.nc_user||'',nc_pass:c.nc_pass||'',ok:!!(c.nc_url&&c.nc_user&&c.nc_pass)};}
const K={comm:'asg_commesse',conti:'asg_conti',scad:'asg_scadenze',todo:'asg_todos',det:'asg_det',fatt:'asg_fatture',movimenti_fineco:'asg_mov_fineco',movimenti_sella:'asg_mov_sella'};

// ═══════════════════════════════════════════════
// NAVIGAZIONE SIDEBAR v13
// ═══════════════════════════════════════════════
const VIEW_TITLES={
  home:'Home',commesse:'Commesse',approv:'Approvvigionamenti',
  contabilita:'Contabilità','import-fatture':'Import Fatture',
  magazzino:'Magazzino',fornitori:'Fornitori & Componenti',detail:'Commessa'
};
let _currentView='home';

function navTo(viewId){
  // Nascondi tutte le view
  document.querySelectorAll('.view').forEach(v=>v.classList.remove('active'));
  // Mostra quella richiesta
  const el=document.getElementById('view-'+viewId);
  if(el)el.classList.add('active');
  _currentView=viewId;
  // Aggiorna sidebar
  document.querySelectorAll('.nav-item').forEach(n=>{
    n.classList.toggle('active',n.dataset.view===viewId);
  });
  // Aggiorna topbar
  const tt=document.getElementById('topbar-title');
  if(tt)tt.textContent=VIEW_TITLES[viewId]||viewId;
  // Azioni topbar per vista
  const ta=document.getElementById('topbar-actions');
  if(ta){
    if(viewId==='home')ta.innerHTML='<span id="clk" style="font-size:12px;color:var(--text3)">'+new Date().toLocaleDateString('it-IT',{weekday:'long',day:'2-digit',month:'long'})+'</span>';
    else if(viewId==='commesse')ta.innerHTML='';
    else if(viewId==='approv')ta.innerHTML='<button class="btn btn-ghost btn-sm" onclick="apriRDOTutti()">✉ Genera tutti RDO</button>';
    else ta.innerHTML='';
  }
  // Render specifico
  if(viewId==='home')renderHome();
  else if(viewId==='commesse')renderCommesse();
  else if(viewId==='approv')renderApprov();
  else if(viewId==='contabilita')renderContabilita();
  else if(viewId==='magazzino'){renderMag();renderMagMetrics();}
  else if(viewId==='fornitori')renderFornCards(FORNITORI);
  else if(viewId==='import-fatture')updateImpStatus();
}

function contTab(tab){
  document.querySelectorAll('.cont-tab').forEach((t,i)=>{
    const tabs=['fatture','scadenze','cassa'];
    t.classList.toggle('active',tabs[i]===tab);
  
  if(tab==='primanota'){renderPrimaNota();}
});
  document.querySelectorAll('.cont-panel').forEach(p=>p.classList.remove('active'));
  const el=document.getElementById('cont-'+tab);
  if(el)el.classList.add('active');
  if(tab==='scadenze'){renderScadContabilita();if(!_flussoCassaVisible){toggleFlussoCassa();}}
  else if(tab==='cassa')renderCassa();
  else if(tab==='fatture')renderFatt();
}


// ═══════════════════════════════════════════════
// PRIMA NOTA — Import, Storage, Render
// ═══════════════════════════════════════════════
let _pnImportRows = [];

function lPN(){ try{return JSON.parse(localStorage.getItem('asg_primanota')||'[]');}catch{return[];} }
function sPN(d){ localStorage.setItem('asg_primanota', JSON.stringify(d)); }

// Aggiorna filtro conti disponibili
function updatePNContoFilter(){
  const data = lPN();
  const conti = [...new Set(data.map(r=>r.conto).filter(Boolean))].sort();
  const sel = document.getElementById('pn-filtro-conto');
  if(!sel) return;
  const cur = sel.value;
  sel.innerHTML = '<option value="">Tutti i conti</option>' +
    conti.map(c=>`<option value="${c}"${c===cur?' selected':''}>${c}</option>`).join('');
}

// Render tabella Prima Nota
function renderPrimaNota(){
  updatePNContoFilter();
  const data = lPN();
  const filtConto = document.getElementById('pn-filtro-conto')?.value||'';
  const filtDa   = document.getElementById('pn-filtro-da')?.value||'';
  const filtA    = document.getElementById('pn-filtro-a')?.value||'';

  let rows = data;
  if(filtConto) rows = rows.filter(r=>r.conto===filtConto);
  if(filtDa)   rows = rows.filter(r=>r.data_pag>=filtDa+'-01');
  if(filtA)    rows = rows.filter(r=>r.data_pag<=filtA+'-31');

  // Ordina per data pagamento
  rows = [...rows].sort((a,b)=>a.data_pag.localeCompare(b.data_pag));

  // Saldi per conto
  const saldiEl = document.getElementById('pn-saldi');
  if(saldiEl){
    const conti = [...new Set(data.map(r=>r.conto).filter(Boolean))];
    saldiEl.innerHTML = conti.map(c=>{
      const mov = filtConto ? rows : data.filter(r=>r.conto===c);
      const ent = mov.reduce((s,r)=>s+r.entrata,0);
      const usc = mov.reduce((s,r)=>s+r.uscita,0);
      const sal = ent - usc;
      return `<div style="background:var(--bg1);border:1px solid var(--border);border-radius:var(--r-sm);padding:8px 12px;font-size:12px;">
        <div style="font-weight:700;color:var(--text2);margin-bottom:2px">${c}</div>
        <div style="display:flex;gap:10px;">
          <span style="color:var(--green)">↑ €${fmt(ent)}</span>
          <span style="color:var(--red)">↓ €${fmt(usc)}</span>
          <span style="font-weight:700;color:${sal>=0?'var(--blue)':'var(--red)'}">= €${fmt(sal)}</span>
        </div>
      </div>`;
    }).join('');
  }

  const tbody = document.getElementById('pn-tbody');
  const tfoot = document.getElementById('pn-tfoot');
  if(!tbody) return;

  if(!rows.length){
    tbody.innerHTML='<tr><td colspan="10" class="empty">Nessun movimento per questo filtro</td></tr>';
    if(tfoot) tfoot.innerHTML='';
    return;
  }

  let saldoRun = 0;
  const totEnt = rows.reduce((s,r)=>s+r.entrata,0);
  const totUsc = rows.reduce((s,r)=>s+r.uscita,0);

  tbody.innerHTML = rows.map(r=>{
    saldoRun += r.entrata - r.uscita;
    return `<tr style="border-bottom:1px solid var(--border);">
      <td style="padding:6px 10px;font-family:monospace;font-size:11px">${fmtD(r.data_pag)}</td>
      <td style="padding:6px 8px;font-size:11px"><span class="bdg" style="font-size:10px">${r.tipo||'—'}</span></td>
      <td style="padding:6px 8px;font-size:11px;font-family:monospace">${r.num_doc||'—'}</td>
      <td style="padding:6px 8px;font-size:11px;font-family:monospace">${fmtD(r.data_doc)}</td>
      <td style="padding:6px 8px;font-size:12px;max-width:150px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${(r.cliente||r.fornitore||'').replace(/"/g,'')}">${r.cliente||r.fornitore||'—'}</td>
      <td style="padding:6px 8px;font-size:11px;color:var(--text3);max-width:120px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${(r.nota||'').replace(/"/g,'')}">${r.nota||''}</td>
      <td style="padding:6px 8px;font-size:11px"><span style="background:var(--blue-lt);color:var(--blue);padding:1px 6px;border-radius:10px;font-size:10px;font-weight:600">${r.conto||'—'}</span></td>
      <td style="padding:6px 8px;text-align:right;font-weight:600;color:var(--green)">${r.entrata>0?'€'+fmt2(r.entrata):''}</td>
      <td style="padding:6px 8px;text-align:right;font-weight:600;color:var(--red)">${r.uscita>0?'€'+fmt2(r.uscita):''}</td>
      <td style="padding:6px 8px;text-align:right;font-family:monospace;font-size:11px;color:${saldoRun>=0?'var(--text2)':'var(--red)'}">€${fmt2(saldoRun)}</td>
    </tr>`;
  }).join('');

  if(tfoot) tfoot.innerHTML = `<tr style="background:var(--bg2);font-weight:700;border-top:2px solid var(--border);">
    <td colspan="7" style="padding:7px 10px;font-size:12px;color:var(--text3)">TOTALE — ${rows.length} movimenti</td>
    <td style="padding:7px 8px;text-align:right;color:var(--green)">€${fmt2(totEnt)}</td>
    <td style="padding:7px 8px;text-align:right;color:var(--red)">€${fmt2(totUsc)}</td>
    <td style="padding:7px 8px;text-align:right;font-weight:800">€${fmt2(totEnt-totUsc)}</td>
  </tr>`;
}

// Handler drag & drop
function handlePNDrop(e){
  e.preventDefault();
  handlePNFiles(e.dataTransfer.files);
}

// Parser XLS Prima Nota Fattura24
async function handlePNFiles(files){
  _pnImportRows = [];
  const preview = document.getElementById('pn-import-preview');
  const btn = document.getElementById('pn-import-btn');
  if(preview) preview.innerHTML = '<div style="color:var(--text3);font-size:12px;padding:8px">Lettura file…</div>';

  for(const file of files){
    try{
      const rows = await parsePrimaNotaFile(file);
      _pnImportRows.push(...rows);
    } catch(err){
      if(preview) preview.innerHTML += `<div style="color:var(--red);font-size:12px;padding:4px">✗ ${file.name}: ${err.message}</div>`;
    }
  }

  if(_pnImportRows.length){
    if(preview) preview.innerHTML = `<div style="background:var(--green-lt);border:1px solid var(--green-bd);border-radius:var(--r-sm);padding:8px 12px;font-size:12px;color:var(--green)">
      ✓ Trovati <strong>${_pnImportRows.length}</strong> movimenti da importare<br>
      Conti: ${[...new Set(_pnImportRows.map(r=>r.conto))].filter(Boolean).join(', ')}
    </div>`;
    if(btn) btn.style.display='';
  }
}

async function parsePrimaNotaFile(file){
  return new Promise((resolve, reject)=>{
    const reader = new FileReader();
    reader.onload = e => {
      try{
        const wb = XLSX.read(e.target.result, {type:'array', cellDates:true, dateNF:'dd/mm/yyyy'});
        const ws = wb.Sheets[wb.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(ws, {header:1, defval:'', raw:false, dateNF:'dd/mm/yyyy'});

        // Trova riga header (contiene "DATA PAG." o "TIPO DOC.")
        let hdrIdx = -1;
        for(let i=0;i<raw.length;i++){
          const r = raw[i].map(c=>String(c).toUpperCase().trim());
          if(r.includes('DATA PAG.') || r.includes('TIPO DOC.')){
            hdrIdx = i; break;
          }
        }
        if(hdrIdx<0) throw new Error('Header non trovato. Usa export "Prima Nota" da Fattura24.');

        const headers = raw[hdrIdx].map(c=>String(c).trim());
        const rows = [];

        for(let i=hdrIdx+1; i<raw.length; i++){
          const row = {};
          headers.forEach((h,j)=>{ row[h]=String(raw[i][j]??'').trim(); });

          // Converti serial Excel in data se necessario
          const excelDateToStr = v => {
            if(!v)return'';
            if(typeof v==='number'&&v>40000&&v<60000){
              // Serial number Excel: giorni dal 1/1/1900
              const d=new Date(Math.round((v-25569)*86400*1000));
              return d.getUTCFullYear()+'-'+String(d.getUTCMonth()+1).padStart(2,'0')+'-'+String(d.getUTCDate()).padStart(2,'0');
            }
            return normDate(String(v));
          };
          const dataPag = excelDateToStr(row['DATA PAG.']||row['Data pagamento']||'');
          if(!dataPag) continue; // salta righe vuote

          const toNum = v => {
            if(v===null||v===undefined||v===''||v==='-')return 0;
            // SheetJS raw:true restituisce numeri JS nativi
            if(typeof v==='number') return isFinite(v)?v:0;
            const s=String(v).trim();
            if(!s||s==='-')return 0;
            // Rimuovi simbolo €  e spazi
            const clean=s.replace(/[€\s]/g,'');
            // Formato italiano: punto migliaia, virgola decimale → "1.234,56"
            if(/^\d{1,3}(\.\d{3})*(,\d+)?$/.test(clean))
              return parseFloat(clean.replace(/\./g,'').replace(',','.'))||0;
            // Formato con solo virgola decimale → "54,89"
            if(/^\d+,\d+$/.test(clean))
              return parseFloat(clean.replace(',','.'))||0;
            // Già formato decimale con punto → "54.89"
            return parseFloat(clean.replace(/,/g,''))||0;
          };

          rows.push({
            id: 'pn_'+(row['ID']||row['Cod.']||Math.random().toString(36).slice(2)),
            data_pag: dataPag,
            tipo: row['TIPO DOC.']||row['Tipo']||'',
            num_doc: row['NUM. DOC.']||row['Numero']||'',
            data_doc: excelDateToStr(row['DATA DOC.']||row['Data doc.']||''),
            nota: row['NOTA']||row['Descrizione']||'',
            conto: row['CONTO']||row['Conto']||'',
            cliente: row['CLIENTE']||row['Cliente']||'',
            fornitore: row['FORNITORE']||row['Fornitore']||'',
            entrata: toNum(row['ENTRATA']||row['Entrata']||0),
            uscita: toNum(row['USCITA']||row['Uscita']||0),
          });
        }

        if(!rows.length) throw new Error('Nessun movimento trovato nel file.');
        resolve(rows);
      } catch(err){ reject(err); }
    };
    reader.onerror = ()=>reject(new Error('Errore lettura file'));
    reader.readAsArrayBuffer(file);
  });
}

function confirmImportPN(){
  if(!_pnImportRows.length) return;
  const existing = lPN();
  // Deduplica per id
  const existingIds = new Set(existing.map(r=>r.id));
  const nuovi = _pnImportRows.filter(r=>!existingIds.has(r.id));
  const merged = [...existing, ...nuovi];
  sPN(merged);
  closeM();
  showToast(`✓ Importati ${nuovi.length} movimenti (${_pnImportRows.length-nuovi.length} duplicati saltati)`);
  // Sincronizza su Nextcloud
  wPutJSON('primanota.json', merged).catch(()=>{});
  contTab('primanota');
}

// Export PDF Prima Nota
function exportPrimaNotaPDF(){
  const data = lPN();
  const filtConto = document.getElementById('pn-filtro-conto')?.value||'';
  const filtDa    = document.getElementById('pn-filtro-da')?.value||'';
  const filtA     = document.getElementById('pn-filtro-a')?.value||'';
  let rows = data;
  if(filtConto) rows = rows.filter(r=>r.conto===filtConto);
  if(filtDa)    rows = rows.filter(r=>r.data_pag>=filtDa+'-01');
  if(filtA)     rows = rows.filter(r=>r.data_pag<=filtA+'-31');
  rows = [...rows].sort((a,b)=>a.data_pag.localeCompare(b.data_pag));

  const totEnt = rows.reduce((s,r)=>s+r.entrata,0);
  const totUsc = rows.reduce((s,r)=>s+r.uscita,0);
  const periodo = filtDa||filtA ? `${filtDa||'…'} / ${filtA||'…'}` : 'Tutto il periodo';
  const contoLbl = filtConto||'Tutti i conti';

  const win = window.open('','_blank');
  win.document.write(`<!DOCTYPE html><html><head><meta charset="utf-8">
  <title>Prima Nota — ASG LAB</title>
  <style>
    body{font-family:sans-serif;font-size:11px;color:#111;margin:20px;}
    h1{font-size:16px;margin-bottom:4px}
    .sub{color:#666;font-size:11px;margin-bottom:16px}
    table{width:100%;border-collapse:collapse;font-size:10px;}
    th{background:#f0f0f0;padding:5px 6px;text-align:left;border-bottom:2px solid #ccc;font-size:10px;}
    td{padding:4px 6px;border-bottom:1px solid #eee;}
    .num{text-align:right;}
    .green{color:#166534;font-weight:600}
    .red{color:#991b1b;font-weight:600}
    tfoot td{font-weight:700;background:#f7f7f7;border-top:2px solid #ccc;}
    @media print{body{margin:0}button{display:none}}
  </style></head><body>
  <button onclick="window.print()" style="margin-bottom:12px;padding:6px 16px;background:#1a56db;color:#fff;border:none;border-radius:5px;cursor:pointer">🖨 Stampa / Salva PDF</button>
  <h1>Prima Nota — ASG LAB</h1>
  <div class="sub">Conto: <strong>${contoLbl}</strong> · Periodo: <strong>${periodo}</strong> · ${rows.length} movimenti</div>
  <table>
    <thead><tr>
      <th>Data pag.</th><th>Tipo</th><th>N° Doc.</th><th>Data doc.</th>
      <th>Cliente/Fornitore</th><th>Nota</th><th>Conto</th>
      <th class="num" style="color:green">Entrata</th>
      <th class="num" style="color:red">Uscita</th>
    </tr></thead>
    <tbody>
    ${rows.map(r=>`<tr>
      <td>${fmtD(r.data_pag)}</td>
      <td>${r.tipo||'—'}</td>
      <td>${r.num_doc||'—'}</td>
      <td>${fmtD(r.data_doc)}</td>
      <td>${r.cliente||r.fornitore||'—'}</td>
      <td style="color:#666">${r.nota||''}</td>
      <td>${r.conto||'—'}</td>
      <td class="num green">${r.entrata>0?'€'+fmt2(r.entrata):''}</td>
      <td class="num red">${r.uscita>0?'€'+fmt2(r.uscita):''}</td>
    </tr>`).join('')}
    </tbody>
    <tfoot><tr>
      <td colspan="7">TOTALE</td>
      <td class="num" style="color:green">€${fmt2(totEnt)}</td>
      <td class="num" style="color:red">€${fmt2(totUsc)}</td>
    </tr></tfoot>
  </table>
  <div style="margin-top:12px;font-size:10px;color:#666">
    Saldo periodo: <strong style="color:${totEnt-totUsc>=0?'#166534':'#991b1b'}">€${fmt2(totEnt-totUsc)}</strong>
    · Estratto il ${new Date().toLocaleDateString('it-IT')}
  </div>
  </body></html>`);
  win.document.close();
}

function renderContabilita(){
  // Renderizza il tab attivo
  const activeTab=document.querySelector('.cont-tab.active');
  const tabs=['fatture','scadenze','cassa'];
  const active=tabs[document.querySelectorAll('.cont-tab').length?
    [...document.querySelectorAll('.cont-tab')].findIndex(t=>t.classList.contains('active')):0]||'fatture';
  if(active==='scadenze')renderScadContabilita();
  else if(active==='cassa')renderCassa();
  else renderFatt();
}

// ═══════════════════════════════════════════════
// RENDER HOME
// ═══════════════════════════════════════════════
function renderHome(){
  const fatt=SD.fatt||lll('fatt');
  const conti=SD.cassa||lll('conti');
  const comm=SD.comm||lll('comm');
  const scad=[...(SD.scad||[]),...lll('scad').filter(l=>!(SD.scad||[]).find(n=>n.id===l.id))];

  // Metriche
  const tot=conti.reduce((s,c)=>s+parseFloat(c.saldo||0),0);
  const incassi=fatt.filter(f=>f.tipo==='FE'&&f.stato==='nonpagata');
  const pagamenti=fatt.filter(f=>f.tipo==='FA'&&f.stato==='nonpagata');
  const totInc=incassi.reduce((s,f)=>s+parseImporto(f.importo),0);
  const totPag=pagamenti.reduce((s,f)=>s+parseImporto(f.importo),0);
  const oggi=new Date();oggi.setHours(0,0,0,0);
  const fine30=new Date(oggi);fine30.setDate(fine30.getDate()+30);
  const scadFiscali30=scad.filter(s=>s._fiscale&&s.pagato!=='si'&&parseFloat(s.importo||0)>0&&new Date(s.data)<=fine30);
  const totFis30=scadFiscali30.reduce((s,x)=>s+parseFloat(x.importo||0),0);

  const setEl=(id,v)=>{const e=document.getElementById(id);if(e)e.textContent=v;};
  setEl('h-liquidita',conti.length?'€'+fmt(tot):'€—');
  setEl('h-liquidita-sub',conti.length?conti.length+' conti · agg. '+fmtD(conti.sort((a,b)=>b.data_rilevazione?.localeCompare(a.data_rilevazione||'')||0)[0]?.data_rilevazione):'aggiorna saldi banca');
  setEl('h-incassare',fatt.length?'€'+fmt(totInc):'€—');
  setEl('h-incassare-sub',incassi.length+' fatture emesse aperte');
  setEl('h-pagare',fatt.length?'€'+fmt(totPag+totFis30):'€—');
  setEl('h-pagare-sub',pagamenti.length+' fornitori · €'+fmt(totFis30)+' fisco 30gg');

  // Commesse in corso
  const inCorso=(comm.length?comm:lll('comm')).filter(c=>c.stato==='in-corso'||c.stato==='confermata'||c.stato==='inviata'||c.stato==='offerta').sort((a,b)=>(b.data_apertura||'').localeCompare(a.data_apertura||'')).slice(0,6);
  setEl('h-comm-count',inCorso.length||'0');
  const cl=document.getElementById('h-comm-list');
  if(cl){
    if(!inCorso.length){cl.innerHTML='<div class="empty">Nessuna commessa in corso</div>';}
    else{
      const STATI_COMM={confermata:'badge-green','in-corso':'badge-green',offerta:'badge-amber',ricerca:'badge-blue',chiusa:'badge-gray'};
      cl.innerHTML=inCorso.map(c=>{
        const st=c.stato||'';
        const badge=STATI_COMM[st]||'badge-gray';
        return`<div class="list-row" onclick="openDetail('${c.numero}')">
          <div class="sdot ${st==='confermata'||st==='in-corso'?'sdot-green':st==='offerta'?'sdot-amber':'sdot-gray'}"></div>
          <div class="list-row-main"><div class="list-row-title">${c.numero}</div><div class="list-row-sub">${CMAP[getCod(c.numero)]||c.cliente||''} · ${c.oggetto||'—'}</div></div>
          <span class="badge ${badge}">${st}</span>
        </div>`;
      }).join('');
    }
  }
  // Badge navbar
  const nb=document.getElementById('nb-commesse');
  if(nb){if(inCorso.length>0){nb.style.display='';nb.textContent=inCorso.length;}else nb.style.display='none';}

  // Scadenze urgenti (entro 14gg, non pagate)
  const urgenti=scad.filter(s=>s.pagato!=='si'&&dTo(s.data)>=0&&dTo(s.data)<=14).sort((a,b)=>new Date(a.data)-new Date(b.data)).slice(0,6);
  setEl('h-scad-count',urgenti.length||'0');
  const sl=document.getElementById('h-scad-list');
  if(sl){
    if(!urgenti.length){sl.innerHTML='<div class="empty">Nessuna scadenza urgente</div>';}
    else{
      sl.innerHTML=urgenti.map(s=>{
        const d=dTo(s.data);
        const imp=parseFloat(s.importo||0);
        const isInc=s.categoria==='incasso';
        return`<div class="list-row" onclick="navTo('contabilita');setTimeout(()=>{contTab('scadenze');openModal('edit-scad','${s.id}')},80)">
          <div class="sdot ${d<=3?'sdot-red':d<=7?'sdot-amber':'sdot-gray'}"></div>
          <div class="list-row-main"><div class="list-row-title" style="font-size:12px">${s.descrizione||'—'}</div><div class="list-row-sub">${fmtD(s.data)}</div></div>
          ${imp>0?'<div class="list-row-val '+(isInc?'green':'red')+'">'+( isInc?'+':'−')+'€'+fmt(imp)+'</div>':'<div class="list-row-val" style="color:var(--text4);font-size:11px">da def.</div>'}
        </div>`;
      }).join('');
    }
  }

  // Approvvigionamenti da ordinare
  const approvPend=getApprovAggregati('daordinare');
  const totApprov=Object.values(approvPend).reduce((t,a)=>t+a.length,0);
  setEl('h-approv-count',totApprov||'0');
  const al=document.getElementById('h-approv-list');
  if(al){
    if(!totApprov){al.innerHTML='<div class="empty">Nessun componente da ordinare</div>';}
    else{
      const fornitori=Object.keys(approvPend).slice(0,4);
      al.innerHTML=fornitori.map(f=>{
        const items=approvPend[f];
        return`<div class="list-row" onclick="navTo('approv')">
          <div class="sdot sdot-amber"></div>
          <div class="list-row-main"><div class="list-row-title">${f}</div><div class="list-row-sub">${items.length} componenti da ${[...new Set(items.map(i=>i.numero))].join(', ')}</div></div>
          <span class="badge badge-amber">${items.length} pz</span>
        </div>`;
      }).join('');
    }
  }
  const nba=document.getElementById('nb-approv');
  if(nba){if(totApprov>0){nba.style.display='';nba.textContent=totApprov;}else nba.style.display='none';}

  // Magazzino + azioni rapide
  const mag=lll('mag');
  const sottoScorta=mag.filter(m=>parseFloat(m.qty_min||0)>0&&parseFloat(m.qty||0)<parseFloat(m.qty_min||0));
  const todo=lll('todo').filter(t=>!t.done);
  const nbm=document.getElementById('nb-mag');
  if(nbm){if(sottoScorta.length>0){nbm.style.display='';nbm.textContent=sottoScorta.length;}else nbm.style.display='none';}
  const ma=document.getElementById('h-mag-actions');
  if(ma){
    let html='';
    if(sottoScorta.length){
      html+=sottoScorta.slice(0,3).map(m=>`<div class="list-row" onclick="navTo('magazzino')">
        <div class="sdot sdot-red"></div>
        <div class="list-row-main"><div class="list-row-title" style="font-size:12px">⚠ ${m.descrizione||m.codice||'—'}</div><div class="list-row-sub">${m.qty||0} pz · scorta min. ${m.qty_min}</div></div>
      </div>`).join('');
    }
    todo.slice(0,3).forEach(t=>{
      html+=`<div class="list-row"><div class="sdot sdot-${t.pri==='alta'?'red':t.pri==='media'?'amber':'gray'}"></div>
        <div class="list-row-main"><div class="list-row-title" style="font-size:12px">${t.testo}</div></div>
        <button class="xbtn" onclick="togTodo('${t.id}');renderHome();event.stopPropagation()">✓</button>
      </div>`;
    });
    if(!html)html='<div class="list-row"><div class="sdot sdot-green"></div><div class="list-row-main"><div class="list-row-title" style="font-size:12px">Tutto in ordine</div></div></div>';
    ma.innerHTML=html+`<div style="padding:8px 14px;border-top:1px solid var(--border);display:flex;gap:8px;flex-wrap:wrap;">
      <button class="btn btn-ghost btn-sm" onclick="openModal('todo')">+ Azione</button>
      <button class="btn btn-ghost btn-sm" onclick="openModal('conto')">+ Saldo banca</button>
    </div>`;
  }
}

// ═══════════════════════════════════════════════
// APPROVVIGIONAMENTI
// ═══════════════════════════════════════════════
let _approvFiltro='daordinare';

function setFiltroApprov(el,val){
  document.querySelectorAll('[data-appf]').forEach(e=>e.classList.remove('active'));
  el.classList.add('active');
  _approvFiltro=val;
  renderApprov();
}

function getApprovAggregati(filtroStato){
  const det=ll('det');
  const aggregato={}; // fornitore → [{numero,descrizione,qty,stato_approv,...}]
  Object.entries(det).forEach(([numero,cd])=>{
    if(!cd.distinta)return;
    cd.distinta.forEach(r=>{
      if(!r.fornitore)return;
      const st=r.stato_approv||'daordinare';
      if(filtroStato&&filtroStato!=='tutti'&&st!==filtroStato)return;
      const f=r.fornitore.trim();
      if(!aggregato[f])aggregato[f]=[];
      aggregato[f].push({...r,numero,_commessa:numero});
    });
  });
  return aggregato;
}

function renderApprov(){
  const q=(document.getElementById('approv-search')||{}).value?.toLowerCase()||'';
  const aggr=getApprovAggregati(_approvFiltro==='tutti'?null:_approvFiltro||'daordinare');
  const el=document.getElementById('approv-content');if(!el)return;

  const SAP_LABEL={'daordinare':'Da ordinare','ordinato':'Ordinato','ricevuto':'Ricevuto','in-ritardo':'In ritardo','nonrichiesto':'Non richiesto'};
  const SAP_BADGE={'daordinare':'badge-amber','ordinato':'badge-blue','ricevuto':'badge-green','in-ritardo':'badge-red','nonrichiesto':'badge-gray'};

  let html='';
  const fornitori=Object.keys(aggr).sort();
  let totale=0;
  fornitori.forEach(forn=>{
    let items=aggr[forn];
    if(q)items=items.filter(i=>(i.descrizione||'').toLowerCase().includes(q)||(i.codice||'').toLowerCase().includes(q));
    if(!items.length)return;
    totale+=items.length;
    const nDaOrd=items.filter(i=>(!i.stato_approv||i.stato_approv==='daordinare')).length;
    html+=`<div class="approv-group">
      <div class="approv-group-head">
        <span class="approv-group-nome">${forn}</span>
        <span class="approv-group-count">${items.length} componenti${nDaOrd>0?' · '+nDaOrd+' da ordinare':''}</span>
        <span style="flex:1"></span>
        ${nDaOrd>0?'<button class="btn btn-primary btn-sm" data-forn="'+encodeURIComponent(forn)+'" onclick="apriRDO(decodeURIComponent(this.dataset.forn))">✉ RDO</button>':''}
      </div>
      ${items.map(r=>`<div class="approv-row">
        <div class="approv-desc">
          <div class="approv-title">${r.descrizione||r.codice||'—'}</div>
          <div class="approv-commesse">${r._commessa} · ${r.codice||'—'}</div>
        </div>
        <span class="badge ${SAP_BADGE[r.stato_approv||'daordinare']}">${SAP_LABEL[r.stato_approv||'daordinare']}</span>
        <select class="stato-approv-select" onchange="aggiornaStatoApprov('${r._commessa}','${r.id}',this.value)">${Object.entries(SAP_LABEL).map(([k,v])=>`<option value="${k}"${(r.stato_approv||'daordinare')===k?' selected':''}>${v}</option>`).join('')}</select>
      </div>`).join('')}
    </div>`;
  });

  if(!html){el.innerHTML='<div class="empty"><div class="empty-icon">✓</div>Nessun componente per il filtro selezionato</div>';return;}
  el.innerHTML=html;
  setTimeout(()=>upgradeSelects(el),0);
}

function aggiornaStatoApprov(numero,rid,nuovoStato){
  const det=ll('det');
  if(!det[numero])return;
  const idx=det[numero].distinta.findIndex(r=>r.id===rid);
  if(idx<0)return;
  det[numero].distinta[idx].stato_approv=nuovoStato;
  ls('det',det);
  syncDetBackground();
  renderApprov();
  renderHome();
}

function apriRDOTutti(){
  const aggr=getApprovAggregati('daordinare');
  const fornitori=Object.keys(aggr);
  if(!fornitori.length){showToast('Nessun componente da ordinare');return;}
  // Apre il primo RDO come esempio
  apriRDO(fornitori[0]);
}

// Scadenze nel tab contabilità

// ═══════════════════════════════════════════════
// SCADENZE — FILTRI E FLUSSO DI CASSA
// ═══════════════════════════════════════════════
let _filtroScad='tutti';
let _flussoCassaVisible=false;

function setFiltroScad(el,val){
  document.querySelectorAll('[data-sf]').forEach(e=>e.classList.remove('active'));
  if(el)el.classList.add('active');
  _filtroScad=val;
  renderScadContabilita();
}

function toggleFlussoCassa(){
  _flussoCassaVisible=!_flussoCassaVisible;
  const el=document.getElementById('scad-flusso-cont');
  const btn=document.getElementById('btn-flusso');
  if(!el)return;
  if(_flussoCassaVisible){
    el.style.display='block';
    if(btn)btn.style.background='var(--blue-lt)';
    renderFlussoCassa();
  } else {
    el.style.display='none';
    if(btn)btn.style.background='';
  }
}

function renderFlussoCassa(){
  const el=document.getElementById('scad-flusso-cont');
  if(!el)return;

  const conti=SD.cassa||lll('conti');
  const saldoAttuale=conti.reduce((s,c)=>s+parseFloat(c.saldo||0),0);

  const nc=SD.scad||[];const loc=lll('scad');
  const scad=[...nc,...loc.filter(l=>!nc.find(n=>n.id===l.id))];
  const pendenti=scad.filter(s=>s.pagato!=='si'&&s.data&&parseImporto(s.importo)>0);

  const oggi=new Date();oggi.setHours(0,0,0,0);

  // Costruisci settimane (12 settimane = ~90gg)
  const settimane=[];
  for(let i=0;i<12;i++){
    const inizio=new Date(oggi);inizio.setDate(inizio.getDate()+i*7);
    const fine=new Date(inizio);fine.setDate(fine.getDate()+6);
    settimane.push({inizio,fine,incassi:[],pagamenti:[],fiscali:[]});
  }

  // Distribuisci scadenze nelle settimane
  pendenti.forEach(s=>{
    const ds=new Date(s.data+'T00:00:00');
    if(ds<oggi)return; // già scadute — non proiettare
    const imp=parseImporto(s.importo);
    const isInc=s.categoria==='incasso';
    const isFis=['iva','ires','inps','ritenuta','sostituto','dichiarazione','societario'].includes(s.categoria)||s._fiscale;
    settimane.forEach(w=>{
      if(ds>=w.inizio&&ds<=w.fine){
        if(isInc)w.incassi.push(imp);
        else if(isFis)w.fiscali.push(imp);
        else w.pagamenti.push(imp);
      }
    });
  });

  // Calcola saldo cumulativo
  let saldo=saldoAttuale;
  const punti=settimane.map(w=>{
    const inc=w.incassi.reduce((s,v)=>s+v,0);
    const pag=w.pagamenti.reduce((s,v)=>s+v,0);
    const fis=w.fiscali.reduce((s,v)=>s+v,0);
    saldo=saldo+inc-pag-fis;
    return{...w,inc,pag,fis,saldo};
  });

  const minSaldo=Math.min(saldoAttuale,...punti.map(p=>p.saldo));
  const maxSaldo=Math.max(saldoAttuale,...punti.map(p=>p.saldo));
  const range=Math.max(maxSaldo-minSaldo,1000);

  // Barra grafico SVG semplice
  const W=100/punti.length; // % larghezza per settimana
  const barH=p=>{
    if(maxSaldo<=0)return 10;
    return Math.max(3,Math.round((p.saldo-minSaldo)/range*80));
  };
  const barCol=p=>p.saldo<0?'var(--red)':p.saldo<saldoAttuale*0.3?'var(--amber)':'var(--green)';

  const mesi=['gen','feb','mar','apr','mag','giu','lug','ago','set','ott','nov','dic'];
  const fmtW=d=>`${d.getDate()} ${mesi[d.getMonth()]}`;

  el.innerHTML=`
  <div style="background:var(--bg1);border:1px solid var(--border);border-radius:var(--r);padding:14px;margin-bottom:12px;">
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;">
      <div style="font-size:13px;font-weight:600;color:var(--text)">Flusso di cassa — proiezione 90 giorni</div>
      <div style="font-size:12px;color:var(--text3)">Saldo attuale: <strong style="color:${saldoAttuale>=0?'var(--green)':'var(--red)'}">€${fmt(saldoAttuale)}</strong></div>
    </div>

    <!-- Grafico a barre -->
    <div style="position:relative;height:100px;display:flex;align-items:flex-end;gap:2px;margin-bottom:6px;border-bottom:1px solid var(--border);">
      ${punti.map((p,i)=>`
        <div style="flex:1;display:flex;flex-direction:column;align-items:center;gap:1px;" title="${fmtW(p.inizio)}–${fmtW(p.fine)}: saldo €${fmt(p.saldo)}${p.inc?' +€'+fmt(p.inc):''}${p.pag?' −€'+fmt(p.pag):''}${p.fis?' fisco −€'+fmt(p.fis):''}">
          <div style="font-size:9px;color:${barCol(p)};font-weight:700;margin-bottom:1px">${p.saldo<0?'−':''}€${Math.abs(p.saldo)>=1000?(Math.abs(p.saldo)/1000).toFixed(0)+'k':Math.abs(p.saldo).toFixed(0)}</div>
          <div style="width:100%;background:${barCol(p)};border-radius:2px 2px 0 0;opacity:.85;height:${barH(p)}px;min-height:3px;transition:height .2s;"></div>
        </div>`).join('')}
    </div>
    <div style="display:flex;gap:2px;margin-bottom:12px;">
      ${punti.map((p,i)=>`<div style="flex:1;text-align:center;font-size:9px;color:${i%2===0?'var(--text3)':'var(--text4)'};">${i%3===0?fmtW(p.inizio):''}</div>`).join('')}
    </div>

    <!-- Legenda e movimenti rilevanti -->
    <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;margin-bottom:10px;">
      <div style="font-size:11px;padding:6px 10px;background:var(--green-lt);border-radius:var(--r-sm);color:var(--green);">
        <div style="font-weight:700">Attesi in entrata</div>
        <div style="font-size:14px;font-weight:800">+€${fmt(punti.reduce((s,p)=>s+p.inc,0))}</div>
        <div style="opacity:.7">${pendenti.filter(s=>s.categoria==='incasso').length} fatture</div>
      </div>
      <div style="font-size:11px;padding:6px 10px;background:var(--red-lt);border-radius:var(--r-sm);color:var(--red);">
        <div style="font-weight:700">Uscite previste</div>
        <div style="font-size:14px;font-weight:800">−€${fmt(punti.reduce((s,p)=>s+p.pag+p.fis,0))}</div>
        <div style="opacity:.7">fornitori + fisco</div>
      </div>
      <div style="font-size:11px;padding:6px 10px;background:${minSaldo<0?'var(--red-lt)':minSaldo<saldoAttuale*0.2?'var(--amber-lt)':'var(--green-lt)'};border-radius:var(--r-sm);color:${minSaldo<0?'var(--red)':minSaldo<saldoAttuale*0.2?'var(--amber)':'var(--green)'};">
        <div style="font-weight:700">Saldo minimo</div>
        <div style="font-size:14px;font-weight:800">${minSaldo<0?'−':''}€${fmt(Math.abs(minSaldo))}</div>
        <div style="opacity:.7">${minSaldo<0?'⚠ attenzione deficit':'proiettato 90gg'}</div>
      </div>
    </div>

    ${minSaldo<0?'<div style="background:var(--red-lt);border:1px solid var(--red-bd);border-radius:var(--r-sm);padding:8px 12px;font-size:12px;color:var(--red);font-weight:600;">⚠ Attenzione: la proiezione mostra un possibile deficit di cassa. Considera di anticipare incassi o posticipare pagamenti.</div>':''}
    ${minSaldo>=0&&minSaldo<saldoAttuale*0.2?'<div style="background:var(--amber-lt);border:1px solid var(--amber-bd);border-radius:var(--r-sm);padding:8px 12px;font-size:12px;color:var(--amber);font-weight:600;">⚡ Il saldo scenderà significativamente — tieni monitorati gli incassi.</div>':''}
  </div>`;
}

function renderScadContabilita(){
  const nc=SD.scad||[];const loc=lll('scad');
  const data=[...nc,...loc.filter(l=>!nc.find(n=>n.id===l.id))];
  const pendenti=data.filter(s=>s.pagato!=='si').sort((a,b)=>new Date(a.data)-new Date(b.data));

  const oggi=new Date();oggi.setHours(0,0,0,0);
  const fine30=new Date(oggi);fine30.setDate(fine30.getDate()+30);
  const catFiscali=['iva','ires','inps','ritenuta','sostituto','dichiarazione','societario'];
  const isFiscale=s=>catFiscali.includes(s.categoria)||(!s._fatt_key&&s._fiscale);
  const isFattura=s=>s.categoria==='incasso'||s.categoria==='pagamento'||!!s._fatt_key;
  const fiscali  =pendenti.filter(s=>isFiscale(s)&&!isFattura(s));
  const incassi  =pendenti.filter(s=>s.categoria==='incasso');
  const pagamenti=pendenti.filter(s=>s.categoria==='pagamento');
  const altro    =pendenti.filter(s=>!isFiscale(s)&&!isFattura(s));
  const somma=arr=>arr.reduce((t,s)=>t+parseImporto(s.importo),0);
  const totInc=somma(incassi),totPag=somma(pagamenti);
  const totFis30=somma(fiscali.filter(s=>new Date(s.data)<=fine30));
  const nFisNota=fiscali.filter(s=>new Date(s.data)<=fine30&&parseFloat(s.importo||0)===0).length;

  // Filtro attivo
  const filtro=_filtroScad||'tutti';
  let visible=pendenti;
  if(filtro==='incasso')visible=incassi;
  else if(filtro==='pagamento')visible=pagamenti;
  else if(filtro==='fiscale')visible=fiscali;
  else if(filtro==='urgente')visible=pendenti.filter(s=>dTo(s.data)>=0&&dTo(s.data)<=7);
  else if(filtro==='prossimi30')visible=pendenti.filter(s=>dTo(s.data)>=0&&dTo(s.data)<=30);

  const catBadge={iva:'background:#fef3c7;color:#78350f',ires:'background:#fee2e2;color:#991b1b',inps:'background:#ede9fe;color:#5b21b6',incasso:'background:#dcfce7;color:#166534',pagamento:'background:#dbeafe;color:#1e40af',dichiarazione:'background:#f3f4f6;color:#374151',societario:'background:#ccfbf1;color:#115e59',sostituto:'background:#fef9c3;color:#713f12',ritenuta:'background:#fff7ed;color:#9a3412'};

  const riga=s=>{
    const d=dTo(s.data);
    const st=catBadge[s.categoria]||'background:#f3f4f6;color:#374151';
    const catLabel=s.categoria?`<span class="scad-cat" style="${st}">${s.categoria}</span>`:'';
    const imp=parseImporto(s.importo);
    const isInc=s.categoria==='incasso';
    const scadAlert=d<0?'border-left:3px solid var(--red);':d<=3?'border-left:3px solid var(--red);':d<=7?'border-left:3px solid var(--amber);':'';
    return`<div class="scad-item" style="${scadAlert}cursor:pointer" onclick="openModal('edit-scad','${s.id||''}')">
      <div class="scad-date ${d<0?'red':d<=7?'amber':''}">${fmtD(s.data)}<div style="font-size:9px;color:inherit;opacity:.7">${d<0?'scaduta'+(Math.abs(d)>1?' '+Math.abs(d)+'gg':''):d===0?'oggi':d===1?'domani':d+'gg'}</div></div>
      <div class="scad-body">${catLabel}<span class="scad-desc">${s.descrizione||'—'}</span></div>
      <div class="scad-amount" style="color:${imp>0?(isInc?'var(--green)':'var(--red)'):'var(--text4)'}">
        ${imp>0?(isInc?'+':'−')+'€'+fmt(imp):'<em style="font-size:11px;font-weight:400">da definire</em>'}
      </div>
      <button class="xbtn" onclick="delLoc('scad','${s.id||''}');renderScadContabilita();event.stopPropagation()">✕</button>
    </div>`;
  };

  const sez=(titolo,items,sottotitolo='')=>{
    if(!items.length)return'';
    return`<div style="margin-bottom:14px">
      <div style="font-size:11px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.06em;padding:6px 0 4px;border-bottom:2px solid var(--border);display:flex;justify-content:space-between;">
        <span>${titolo}</span><span style="font-weight:400;font-size:10px;color:var(--text4)">${sottotitolo}</span>
      </div>
      ${items.map(riga).join('')}
    </div>`;
  };

  const summary=document.getElementById('scad-summary-cont');
  const list=document.getElementById('scad-list-cont');

  // Pannello "posso pagare?" — confronta saldo vs uscite prossimi 30gg
  const conti=SD.cassa||lll('conti');
  const saldoAtt=conti.reduce((s,c)=>s+parseFloat(c.saldo||0),0);
  const uscite30=[...pagamenti,...fiscali].filter(s=>dTo(s.data)>=0&&new Date(s.data+'T00:00:00')<=fine30);
  const totUscite30=somma(uscite30);
  const incassi30=incassi.filter(s=>dTo(s.data)>=0&&new Date(s.data+'T00:00:00')<=fine30);
  const totIncassi30=somma(incassi30);
  const saldoProiettato30=saldoAtt+totIncassi30-totUscite30;

  const cashPanel=saldoAtt>0||totUscite30>0?`
  <div style="background:var(--bg1);border:1px solid var(--border);border-radius:var(--r);padding:12px 14px;margin-bottom:14px;">
    <div style="font-size:11px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.06em;margin-bottom:10px">Liquidità — prossimi 30 giorni</div>
    <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center;">
      <div style="flex:1;min-width:100px;">
        <div style="font-size:10px;color:var(--text3)">Saldo attuale</div>
        <div style="font-size:18px;font-weight:800;color:${saldoAtt>=0?'var(--green)':'var(--red)'}">€${fmt(saldoAtt)}</div>
      </div>
      <div style="color:var(--text3);font-size:18px">+</div>
      <div style="flex:1;min-width:100px;">
        <div style="font-size:10px;color:var(--text3)">Incassi attesi 30gg</div>
        <div style="font-size:18px;font-weight:800;color:var(--green)">+€${fmt(totIncassi30)}</div>
        <div style="font-size:10px;color:var(--text3)">${incassi30.length} fatture</div>
      </div>
      <div style="color:var(--text3);font-size:18px">−</div>
      <div style="flex:1;min-width:100px;">
        <div style="font-size:10px;color:var(--text3)">Uscite previste 30gg</div>
        <div style="font-size:18px;font-weight:800;color:var(--red)">−€${fmt(totUscite30)}</div>
        <div style="font-size:10px;color:var(--text3)">${uscite30.length} pagamenti</div>
      </div>
      <div style="color:var(--text3);font-size:18px">=</div>
      <div style="flex:1;min-width:120px;padding:8px 12px;background:${saldoProiettato30<0?'var(--red-lt)':saldoProiettato30<saldoAtt*0.2?'var(--amber-lt)':'var(--green-lt)'};border-radius:var(--r-sm);">
        <div style="font-size:10px;font-weight:700;color:${saldoProiettato30<0?'var(--red)':saldoProiettato30<saldoAtt*0.2?'var(--amber)':'var(--green)'}">Saldo proiettato</div>
        <div style="font-size:20px;font-weight:800;color:${saldoProiettato30<0?'var(--red)':saldoProiettato30<saldoAtt*0.2?'var(--amber)':'var(--green)'}">€${fmt(saldoProiettato30)}</div>
        <div style="font-size:10px;color:${saldoProiettato30<0?'var(--red)':saldoProiettato30<saldoAtt*0.2?'var(--amber)':'var(--green)'}">${saldoProiettato30<0?'⚠ deficit atteso':saldoProiettato30<totUscite30*0.5?'⚡ tieni monitorato':'✓ ok'}</div>
      </div>
    </div>
    ${totUscite30>0?'<div style="margin-top:10px;padding-top:8px;border-top:1px solid var(--border);display:flex;gap:6px;flex-wrap:wrap;">'+
      '<button class="filter-chip" onclick="setFiltroScad(this,\'pagamento\')" style="font-size:11px">\u2193 Vedi pagamenti da fare</button>'+
      '<button class="filter-chip" onclick="setFiltroScad(this,\'incasso\')" style="font-size:11px">\u2191 Vedi incassi attesi</button>'+
    '</div>':''}
  </div>`:'';

  if(summary)summary.innerHTML=(cashPanel||'')+`<div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;margin-bottom:14px;">
    <div class="scad-box inc" style="cursor:pointer" onclick="setFiltroScad(document.querySelector('[data-sf=incasso]'),'incasso')">
      <div class="scad-box-label">Da incassare</div>
      <div class="scad-box-val">+€${fmt(totInc)}</div>
      <div class="scad-box-sub">${incassi.length} fatture aperte</div>
    </div>
    <div class="scad-box pag" style="cursor:pointer" onclick="setFiltroScad(document.querySelector('[data-sf=pagamento]'),'pagamento')">
      <div class="scad-box-label">Da pagare</div>
      <div class="scad-box-val">−€${fmt(totPag)}</div>
      <div class="scad-box-sub">${pagamenti.length} fatture passive</div>
    </div>
    <div class="scad-box fis" style="cursor:pointer" onclick="setFiltroScad(document.querySelector('[data-sf=fiscale]'),'fiscale')">
      <div class="scad-box-label">Fiscali 30gg</div>
      <div class="scad-box-val">${totFis30>0?'−€'+fmt(totFis30):'—'}</div>
      <div class="scad-box-sub">${nFisNota>0?nFisNota+' da definire':fiscali.filter(s=>new Date(s.data)<=fine30).length+' scadenze'}</div>
    </div>
  </div>`;

  if(list){
    if(filtro!=='tutti'){
      list.innerHTML=visible.length?visible.map(riga).join(''):'<div class="empty">Nessuna scadenza per questo filtro</div>';
    } else {
      list.innerHTML=(
        sez('Urgenti — entro 7 giorni',pendenti.filter(s=>dTo(s.data)>=0&&dTo(s.data)<=7),'priorità alta')+
        sez('Da incassare (clienti)',incassi.filter(s=>dTo(s.data)>7),'+€'+fmt(totInc))+
        sez('Da pagare (fornitori)',pagamenti.filter(s=>dTo(s.data)>7),'−€'+fmt(totPag))+
        sez('Scadenze fiscali',fiscali,'clicca per inserire importo')+
        sez('Altro',altro)
      )||'<div class="empty">Nessuna scadenza pendente</div>';
    }
  }
}

// Override renderScad per usare renderScadContabilita quando siamo in contabilità
function renderScad(){
  if(_currentView==='contabilita')renderScadContabilita();
  if(_currentView==='home')renderHome();
}

// Filtri fatture
let _filtroFatt='';
const _sortState={fatt:{col:null,dir:1},scad:{col:null,dir:1}};
function setFiltroFatt(el,val){
  document.querySelectorAll('[data-ft]').forEach(e=>e.classList.remove('active'));
  el.classList.add('active');
  _filtroFatt=val;
  renderFatt();
}

// Filtri commesse
let _filtroCommStato='';
let _filtroCommCliente='';
function setFiltroComm(el,tipo,val){
  if(tipo==='stato'){
    document.querySelectorAll('[data-f-stato]').forEach(e=>e.classList.remove('active'));
    el.classList.add('active');
    _filtroCommStato=val;
  }
  // Reset select anno e cliente quando si clicca chip stato
  renderCommesse();
}

function renderAll(){
  if(_currentView==='home')renderHome();
  else if(_currentView==='commesse')renderCommesse();
  else if(_currentView==='approv')renderApprov();
  else if(_currentView==='contabilita')renderContabilita();
  else if(_currentView==='magazzino'){renderMag();renderMagMetrics();}
  else if(_currentView==='fornitori')renderFornCards(FORNITORI);
  // Sempre aggiorna badge
  setTimeout(()=>{
    const fatt=SD.fatt||lll('fatt');
    const comm=SD.comm||lll('comm');
    const inc=comm.filter(c=>c.stato==='in-corso'||c.stato==='confermata').length;
    const nb=document.getElementById('nb-commesse');
    if(nb){if(inc>0){nb.style.display='';nb.textContent=inc;}else nb.style.display='none';}
  },0);
}


function ll(k){try{return JSON.parse(localStorage.getItem(K[k])||'{}');}catch{return{};}}
function lll(k){try{return JSON.parse(localStorage.getItem(K[k])||'[]');}catch{return[];}}
function ls(k,v){localStorage.setItem(K[k],JSON.stringify(v));}
function uid(){return Date.now().toString(36)+Math.random().toString(36).slice(2,5);}
function fmt(n){return new Intl.NumberFormat('it-IT',{minimumFractionDigits:0,maximumFractionDigits:2}).format(n||0);}
function fmt2(n){return new Intl.NumberFormat('it-IT',{minimumFractionDigits:2,maximumFractionDigits:2}).format(n||0);}
function parseImporto(v){
  // Parsing robusto importo: gestisce "782.12" (punto dec.), "1.234,56" (italiano), numeri puri
  if(!v&&v!==0)return 0;
  if(typeof v==='number')return isNaN(v)?0:v;
  const s=String(v).trim();
  if(!s)return 0;
  if(/,\d{1,2}$/.test(s))return parseFloat(s.replace(/\./g,'').replace(',','.'))||0;
  return parseFloat(s.replace(/,/g,''))||0;
}
function fmtD(d){
  if(!d||d==='—'||d==='null')return'—';
  // Normalizza prima con normDate (gestisce dd/mm/yyyy, yyyy-mm-dd, ecc.)
  const iso=normDate(d);
  if(!iso)return'—';
  const dt=new Date(iso+'T00:00:00');
  if(isNaN(dt))return d;
  return dt.toLocaleDateString('it-IT',{day:'2-digit',month:'short',year:'2-digit'});
}
function dTo(d){if(!d)return 9999;const s=d.includes('/')?d.split('/').reverse().join('-'):d;const n=new Date();n.setHours(0,0,0,0);return Math.round((new Date(s+'T00:00:00')-n)/86400000);}
function parseCSV(t){
  if(!t||!t.trim())return[];
  // Parser CSV con supporto virgolette (RFC 4180)
  function parseLine(line){
    const fields=[];let cur='';let inQ=false;
    for(let i=0;i<line.length;i++){
      const ch=line[i];
      if(inQ){
        if(ch==='"'&&line[i+1]==='"'){cur+='"';i++;}
        else if(ch==='"'){inQ=false;}
        else cur+=ch;
      } else {
        if(ch==='"'){inQ=true;}
        else if(ch===','){fields.push(cur.trim());cur='';}
        else cur+=ch;
      }
    }
    fields.push(cur.trim());
    return fields;
  }
  const lines=t.trim().split('\n').map(l=>l.replace(/\r$/,'')).filter(l=>l.trim());
  if(lines.length<2)return[];
  const h=parseLine(lines[0]);
  return lines.slice(1).map(l=>{
    const v=parseLine(l);
    const o={};h.forEach((x,i)=>o[x]=v[i]||'');
    return o;
  }).filter(o=>Object.values(o).some(v=>v));
}
function getAnno(n){const m=n.match(/(\d{4})$/);return m?m[1]:'';}
function getCod(n){const m=n.match(/^([A-Z]+)/);return m?m[1]:'';}
function folderUrl(n,url){
  const a=getAnno(n),c=getCod(n),k=CMAP[c];
  if(!k||!a)return null;
  const ncBase='https://cloud.asglab.it';
  // Prova a puntare alla cartella commessa specifica
  const seqNum=parseInt((n||'').split('-')[1]||'0');
  const comm=lll('comm').find(x=>x.numero===n);
  if(comm?.oggetto){
    const nomeCartella=`${c}-${seqNum}-${a} (${comm.oggetto})`;
    return ncBase+'/apps/files/?dir='+encodeURIComponent('/Clienti/'+k+'/'+a+'/'+nomeCartella);
  }
  return ncBase+'/apps/files/?dir='+encodeURIComponent('/Clienti/'+k+'/'+a);
}

// ═══════════════════════════════════════════════
// WEBDAV
// ═══════════════════════════════════════════════
async function wGet(f){const c=getCfg();if(!c.ok)return null;try{const r=await fetch(`${c.nc_url}/remote.php/dav/files/${c.nc_user}/Dashboard/${f}`,{headers:{'Authorization':'Basic '+btoa(c.nc_user+':'+c.nc_pass)}});if(!r.ok)return null;return parseCSV(await r.text());}catch{return null;}}

async function wListNCDirs(path){
  // Lista solo le sottocartelle in un percorso Nextcloud
  const c=getCfg();if(!c.ok)return[];
  try{
    const url=c.nc_url+'/remote.php/dav/files/'+c.nc_user+encodeURI(path);
    const r=await fetch(url,{method:'PROPFIND',headers:{
      'Authorization':'Basic '+btoa(c.nc_user+':'+c.nc_pass),
      'Depth':'1','Content-Type':'application/xml'
    },body:`<?xml version="1.0"?><d:propfind xmlns:d="DAV:"><d:prop><d:displayname/><d:resourcetype/></d:prop></d:propfind>`});
    if(!r.ok)return[];
    const xml=await r.text();
    const doc=new DOMParser().parseFromString(xml,'text/xml');
    return[...doc.querySelectorAll('response')].slice(1)
      .filter(resp=>resp.querySelector('collection'))
      .map(resp=>{
        const href=resp.querySelector('href')?.textContent||'';
        const name=resp.querySelector('displayname')?.textContent||decodeURIComponent(href.split('/').filter(Boolean).pop()||'');
        return{name,href};
      }).filter(f=>f.name);
  }catch{return[];}
}

async function wListNC(path){
  // Lista file in una cartella Nextcloud via WebDAV PROPFIND
  const c=getCfg();if(!c.ok)return[];
  try{
    const url=c.nc_url+'/remote.php/dav/files/'+c.nc_user+encodeURI(path);
    const r=await fetch(url,{
      method:'PROPFIND',
      headers:{
        'Authorization':'Basic '+btoa(c.nc_user+':'+c.nc_pass),
        'Depth':'1',
        'Content-Type':'application/xml'
      },
      body:`<?xml version="1.0"?><d:propfind xmlns:d="DAV:"><d:prop><d:displayname/><d:getcontenttype/><d:getlastmodified/><d:getcontentlength/></d:prop></d:propfind>`
    });
    if(!r.ok)return[];
    const xml=await r.text();
    const parser=new DOMParser();
    const doc=parser.parseFromString(xml,'text/xml');
    const responses=[...doc.querySelectorAll('response')].slice(1); // salta la cartella stessa
    return responses.map(resp=>{
      const href=resp.querySelector('href')?.textContent||'';
      const name=resp.querySelector('displayname')?.textContent||href.split('/').pop();
      const type=resp.querySelector('getcontenttype')?.textContent||'';
      const size=parseInt(resp.querySelector('getcontentlength')?.textContent||0);
      const mod=resp.querySelector('getlastmodified')?.textContent||'';
      const isDir=type===''||href.endsWith('/');
      return{name,href,type,size,mod,isDir};
    }).filter(f=>!f.isDir&&f.name);
  }catch(e){console.warn('wListNC:',e);return[];}
}

async function wGetJSON(f){
  const c=getCfg();if(!c.ok)return null;
  try{
    const r=await fetch(c.nc_url+'/remote.php/dav/files/'+c.nc_user+'/Dashboard/'+f,
      {headers:{'Authorization':'Basic '+btoa(c.nc_user+':'+c.nc_pass)}});
    if(!r.ok)return null;
    return JSON.parse(await r.text());
  }catch{return null;}
}

async function wPutJSON(f,obj){
  const c=getCfg();if(!c.ok)return false;
  try{
    const r=await fetch(c.nc_url+'/remote.php/dav/files/'+c.nc_user+'/Dashboard/'+f,{
      method:'PUT',
      headers:{'Authorization':'Basic '+btoa(c.nc_user+':'+c.nc_pass),'Content-Type':'application/json; charset=utf-8'},
      body:JSON.stringify(obj)
    });
    return r.ok||r.status===201||r.status===204;
  }catch{return false;}
}

// Salva det su Nextcloud in background (non bloccante)
let _detSyncTimer=null;
function syncDetBackground(){
  clearTimeout(_detSyncTimer);
  _detSyncTimer=setTimeout(async()=>{
    const det=ll('det');
    await wPutJSON('det.json',det);
  },2000); // debounce 2s — evita write flood su ogni keystroke
}
async function wPut(f,rows,fields){
  const c=getCfg();if(!c.ok)return false;
  const esc=v=>{const s=String(v==null?'':v);return(s.includes(',')||s.includes('"')||s.includes('\n'))?'"'+s.replace(/"/g,'""')+'"':s;};
  const lines=[fields.join(',')].concat(rows.map(r=>fields.map(x=>esc(r[x]||'')).join(',')));
  const body=lines.join('\n')+'\n';
  // Chunking per file grandi: se >300KB usa compressione testo (nessuna, Nextcloud accetta fino a 100MB)
  // Il problema reale era timeout — usiamo un timeout più lungo
  try{
    const ctrl=new AbortController();
    const timer=setTimeout(()=>ctrl.abort(),30000); // 30s timeout
    const r=await fetch(c.nc_url+'/remote.php/dav/files/'+c.nc_user+'/Dashboard/'+f,{
      method:'PUT',
      headers:{'Authorization':'Basic '+btoa(c.nc_user+':'+c.nc_pass),'Content-Type':'text/csv; charset=utf-8'},
      body,
      signal:ctrl.signal
    });
    clearTimeout(timer);
    return r.ok||r.status===201||r.status===204;
  }catch(e){console.warn('wPut error:',e.message);return false;}
};

// Crea cartella su Nextcloud via WebDAV MKCOL (ignora se esiste già)
async function wMkdir(path){
  const c=getCfg();if(!c.ok)return false;
  try{
    const r=await fetch(c.nc_url+'/remote.php/dav/files/'+c.nc_user+encodeURI(path),{
      method:'MKCOL',
      headers:{'Authorization':'Basic '+btoa(c.nc_user+':'+c.nc_pass)}
    });
    return r.ok||r.status===405; // 405 = già esiste, va bene lo stesso
  }catch{return false;}
}

// Crea la struttura cartelle per una commessa: /Clienti/Cliente/Anno/Numero (oggetto)/
async function createCommFolder(numero, oggetto){
  const anno=getAnno(numero), cod=getCod(numero), cliente=CMAP[cod];
  if(!anno||!cliente)return false;
  const seqNum=parseInt((numero||'').split('-')[1]||'0');
  const nomeCartella=cod+'-'+seqNum+'-'+anno+(oggetto?' ('+oggetto+')':'');
  // Crea in sequenza (ogni livello deve esistere prima del successivo)
  await wMkdir('/Clienti');
  await wMkdir('/Clienti/'+cliente);
  await wMkdir('/Clienti/'+cliente+'/'+anno);
  const fullPath='/Clienti/'+cliente+'/'+anno+'/'+nomeCartella;
  const ok=await wMkdir(fullPath);
  return ok?fullPath:false;
}

// ═══════════════════════════════════════════════
// SYNC
// ═══════════════════════════════════════════════
let SD={comm:null,fatt:null,scad:null,cassa:null,det:null,mag:null};
function setSyncUI(state,msg){
  const dot=document.getElementById('sync-dot');
  const lbl=document.getElementById('sync-label');
  if(dot){dot.className='sync-dot'+(state==='ok'?' ok':state==='err'?' err':state==='loading'?' loading':'');}
  if(lbl)lbl.textContent=msg||'';
}
async function syncNow(){
  const c=getCfg();
  if(!c.ok){setSyncUI('','non configurato');return;}
  setSyncUI('loading','sync…');
  document.getElementById('cfg-banner').classList.remove('show');
  try{
    const[co,fa,sc,ca,dj,mg,pn]=await Promise.all([
      wGet('commesse.csv'),wGet('fatture.csv'),
      wGet('scadenze.csv'),wGet('cassa.csv'),
      wGetJSON('det.json'),wGet('magazzino.csv'),
      wGetJSON('primanota.json')
    ]);
    let ok=0;
    // Aggiorna SD + salva in localStorage per persistenza multi-dispositivo
    if(co){SD.comm=co;ls('comm',co);ok++;}
    // Fatture: salva da NC solo se ha più dati del locale (evita perdita dati)
    {const localFatt=lll('fatt');
     if(fa&&fa.length>0&&fa.length>=localFatt.length){SD.fatt=fa;ls('fatt',fa);ok++;}
     else if(localFatt.length>0){SD.fatt=localFatt;} // mantieni locale se NC ha meno dati
     else if(fa&&fa.length>0){SD.fatt=fa;ls('fatt',fa);ok++;}} // NC ha qualcosa, locale vuoto
    if(sc){SD.scad=sc;ls('scad',sc);ok++;}
    if(ca){SD.cassa=ca;ls('conti',ca);ok++;}
    if(dj){SD.det=dj;ls('det',dj);ok++;}
    if(mg&&mg.length){SD.mag=mg;ls('mag',mg);ok++;}
    if(pn&&pn.length){sPN(pn);ok++;}
    if(ok){
      setSyncUI('ok','sync '+new Date().toLocaleTimeString('it-IT',{hour:'2-digit',minute:'2-digit'}));
      localStorage.setItem('asg_ls',new Date().toISOString());
    } else {
      setSyncUI('err','nessun dato — verifica credenziali');
    }
  }catch(e){
    setSyncUI('err','errore: '+e.message);
  }
  renderAll();
}

async function importDaNextcloud(){
  const c=getCfg();
  if(!c.ok){openModal('config');showToast('Configura prima Nextcloud');return;}
  setSyncUI('loading','importazione…');
  try{
    // Carica commesse.csv da Nextcloud e merge con localStorage
    const nc=await wGet('commesse.csv');
    if(!nc||!nc.length){showToast('Nessuna commessa trovata su Nextcloud (commesse.csv)');setSyncUI('ok','ok');openModal('guida-import-nc');return;}
    const loc=lll('comm');
    const locNums=new Set(loc.map(c=>c.numero));
    const nuove=nc.filter(r=>r.numero&&!locNums.has(r.numero));
    const aggiornate=nc.filter(r=>r.numero&&locNums.has(r.numero));
    // Merge: aggiorna quelle esistenti, aggiunge le nuove
    const merged=[...loc];
    aggiornate.forEach(ncr=>{const idx=merged.findIndex(x=>x.numero===ncr.numero);if(idx>=0)merged[idx]={...merged[idx],...ncr};});
    nuove.forEach(r=>{if(!r.id)r.id=uid();merged.unshift(r);});
    ls('comm',merged);SD.comm=merged;
    renderAll();
    setSyncUI('ok','importato');
    showToast(`✓ ${nuove.length} nuove + ${aggiornate.length} aggiornate da Nextcloud`);
  }catch(e){setSyncUI('err','errore');showToast('Errore importazione: '+e.message);}
}

// ═══════════════════════════════════════════════
// MAGAZZINO
// ═══════════════════════════════════════════════
const K_MAG='asg_magazzino';
function lMag(){try{return JSON.parse(localStorage.getItem(K_MAG)||'[]');}catch{return[];}}
function sMag(v){localStorage.setItem(K_MAG,JSON.stringify(v));}

function saveMag(existingId){
  const all=lMag();
  const row={
    id:existingId||uid(),
    codice:(document.getElementById('mg-cod')||{}).value||'',
    descrizione:(document.getElementById('mg-desc')||{}).value||'',
    categoria:(document.getElementById('mg-cat')||{}).value||'vario',
    qty:parseFloat((document.getElementById('mg-qty')||{}).value||0),
    qty_min:parseFloat((document.getElementById('mg-min')||{}).value||0),
    costo_cad:parseFloat((document.getElementById('mg-costo')||{}).value||0),
    fornitore:(document.getElementById('mg-forn')||{}).value||'',
    ubicazione:(document.getElementById('mg-ubic')||{}).value||'',
    note:(document.getElementById('mg-note')||{}).value||'',
    aggiornato:new Date().toISOString().slice(0,10)
  };
  if(existingId){const idx=all.findIndex(x=>x.id===existingId);if(idx>=0)all[idx]=row;else all.push(row);}
  else all.push(row);
  sMag(all);closeM();renderMag();
  showToast(existingId?'✓ Voce aggiornata':'✓ Voce aggiunta al magazzino');
}

function delMag(id){
  if(!confirm('Eliminare questa voce dal magazzino?'))return;
  sMag(lMag().filter(x=>x.id!==id));renderMag();
  showToast('Voce eliminata');
}

function aggMagQty(id,delta){
  const all=lMag();const r=all.find(x=>x.id===id);if(!r)return;
  r.qty=Math.max(0,(r.qty||0)+delta);
  r.aggiornato=new Date().toISOString().slice(0,10);
  sMag(all);renderMag();
}

async function syncMagazzino(){
  const cfg=getCfg();if(!cfg.ok){showToast('Configura Nextcloud prima');return;}
  setSyncUI('loading','sync magazzino…');
  const all=lMag();
  const fields=['id','codice','descrizione','categoria','qty','qty_min','costo_cad','fornitore','ubicazione','note','aggiornato'];
  const ok=await wPut('magazzino.csv',all,fields);
  const nc=await wGet('magazzino.csv');
  if(nc&&nc.length){
    const locIds=new Set(all.map(x=>x.id));
    nc.filter(r=>r.id&&!locIds.has(r.id)).forEach(r=>{
      r.qty=parseFloat(r.qty||0);r.qty_min=parseFloat(r.qty_min||0);r.costo_cad=parseFloat(r.costo_cad||0);
      all.push(r);
    });
    sMag(all);
  }
  setSyncUI(ok?'ok':'err',ok?'sync ok':'errore sync');
  renderMag();
  showToast(ok?'✓ Magazzino sincronizzato':'⚠ Sync fallito — dati salvati localmente');
}

function renderMag(){
  const all=lMag();
  const q=(document.getElementById('mag-search')||{}).value?.toLowerCase()||'';
  const cat=(document.getElementById('mag-filt-cat')||{}).value||'';
  const filt=all.filter(r=>{
    if(cat&&r.categoria!==cat)return false;
    if(q&&!r.codice?.toLowerCase().includes(q)&&!r.descrizione?.toLowerCase().includes(q)&&!r.fornitore?.toLowerCase().includes(q))return false;
    return true;
  });
  const totValore=all.reduce((s,r)=>s+(parseFloat(r.qty||0)*parseFloat(r.costo_cad||0)),0);
  const sottoScorta=all.filter(r=>parseFloat(r.qty_min||0)>0&&parseFloat(r.qty||0)<parseFloat(r.qty_min||0));
  const vm=document.getElementById('mag-m-voci');if(vm)vm.textContent=all.length;
  const vv=document.getElementById('mag-m-val');if(vv)vv.textContent='€'+fmt2(totValore);
  const vl=document.getElementById('mag-m-low');if(vl)vl.textContent=sottoScorta.length;
  const el=document.getElementById('mag-tbody');if(!el)return;
  if(!filt.length){el.innerHTML=`<tr><td colspan="10" class="empty">${all.length?'Nessuna voce corrisponde al filtro':'Magazzino vuoto — aggiungi le prime voci con il bottone + in alto'}</td></tr>`;return;}
  const catColor={valvola:'#ede9fe',pompa:'#ccfbf1',motore:'#fee2e2',serbatoio:'#f3f4f6',blocco:'#f0f9ff',giunto:'#fef3c7',raccorderia:'#dbeafe',elettrica:'#fef9c3',vario:'#f9fafb'};
  el.innerHTML=filt.map(r=>{
    const low=parseFloat(r.qty_min||0)>0&&parseFloat(r.qty||0)<parseFloat(r.qty_min||0);
    const qtyColor=low?'color:var(--red);font-weight:800':'color:var(--green);font-weight:700';
    return`<tr>
      <td class="mono" style="font-size:12px">${r.codice||'—'}</td>
      <td style="font-size:12px;max-width:200px">${r.descrizione||'—'}</td>
      <td><span style="font-size:10px;font-weight:700;padding:2px 7px;border-radius:20px;background:${catColor[r.categoria]||'#f3f4f6'};color:var(--text2)">${r.categoria||'—'}</span></td>
      <td style="font-size:12px">${r.fornitore||'—'}</td>
      <td style="text-align:center">
        <div style="display:flex;align-items:center;gap:5px;justify-content:center;">
          <button onclick="aggMagQty('${r.id}',-1)" style="border:1px solid var(--border);background:#fff;border-radius:5px;width:22px;height:22px;cursor:pointer;font-weight:700;font-size:14px;line-height:1;">−</button>
          <span class="mono" style="${qtyColor};min-width:28px;text-align:center;">${r.qty||0}</span>
          <button onclick="aggMagQty('${r.id}',+1)" style="border:1px solid var(--border);background:#fff;border-radius:5px;width:22px;height:22px;cursor:pointer;font-weight:700;font-size:14px;line-height:1;">+</button>
        </div>
        ${low?'<div style="font-size:10px;color:var(--red);font-weight:700;margin-top:2px;">⚠ sotto scorta</div>':''}
      </td>
      <td class="mono" style="text-align:center;color:var(--text3)">${r.qty_min||0}</td>
      <td class="mono">€${fmt2(r.costo_cad||0)}</td>
      <td style="font-size:12px;color:var(--text3)">${r.ubicazione||'—'}</td>
      <td style="font-size:12px;color:var(--text3);max-width:100px">${r.note||''}</td>
      <td><div style="display:flex;gap:4px;"><button class="xbtn" style="color:var(--blue)" onclick="openModal('add-mag','${r.id}')">✎</button><button class="xbtn" onclick="delMag('${r.id}')">✕</button></div></td>
    </tr>`;
  }).join('');
}
// (showView definita nel modulo Fornitori)

// ═══════════════════════════════════════════════
// RENDER DASHBOARD
// ═══════════════════════════════════════════════


function renderComm(){renderCommesse();}

function renderFatt(){
  // Leggi SEMPRE da localStorage per garantire freschezza
  const raw=localStorage.getItem(K.fatt);
  let all=[];
  try{all=raw?JSON.parse(raw):[];}catch(e){console.error('parse err:',e);all=[];}
  if(all.length)SD.fatt=all;
  let data=all.filter(f=>{
    if(_filtroFatt==='FE')return f.tipo==='FE';
    if(_filtroFatt==='FA')return f.tipo==='FA';
    if(_filtroFatt==='nonpagata')return f.stato==='nonpagata';
    return true;
  });
  // Ordinamento
  data=_sortData(data,_sortState.fatt.col,_sortState.fatt.dir);
  const el=document.getElementById('fatt-tbody');
  if(!el)return;
  // Aggiorna header con frecce ordinamento
  const thead=el.closest('table')?.querySelector('thead tr');
  if(thead)thead.innerHTML=
    `<th style="padding:7px 14px;text-align:left;font-size:11px;color:var(--text3);font-weight:600">Tipo</th>`+
    _thSort('fatt','numero','N°')+
    _thSort('fatt','cliente','Cliente/Fornitore')+
    _thSort('fatt','importo','€')+
    _thSort('fatt','data','Data')+
    _thSort('fatt','data_scad','Scadenza')+
    _thSort('fatt','stato','Stato')+
    `<th style="padding:7px 8px;text-align:center;font-size:11px;color:var(--text3);font-weight:600">📁</th>`;
  if(!data.length){el.innerHTML='<tr><td colspan="8" class="empty">Nessuna fattura — importa da Fattura24</td></tr>';return;}
  el.innerHTML=data.slice(0,100).map(f=>{
    const isFE=f.tipo==='FE';
    const pagata=f.stato==='pagata';
    const link=fattUrl(f);
    const scadColor=!pagata&&f.data_scad?(dTo(f.data_scad)<0?'var(--red)':dTo(f.data_scad)<=7?'var(--amber)':'var(--text3)'):'var(--text3)';
    return`<tr style="cursor:pointer;font-size:12px" onmouseover="this.style.background='var(--bg2)'" onmouseout="this.style.background=''">
      <td style="padding:7px 14px"><span class="badge ${isFE?'badge-green':'badge-amber'}">${f.tipo}</span></td>
      <td style="padding:7px 8px;font-family:monospace">${f.numero||'—'}</td>
      <td style="padding:7px 8px;max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${f.cliente||'—'}</td>
      <td style="padding:7px 8px;text-align:right;font-weight:700;color:${isFE?'var(--green)':'var(--red)'}">€${fmt(f.importo)}</td>
      <td style="padding:7px 8px;color:var(--text3)">${fmtD(f.data)}</td>
      <td style="padding:7px 8px;color:${scadColor}">${fmtD(f.data_scad)||'—'}</td>
      <td style="padding:7px 8px"><span class="badge ${pagata?'badge-green':'badge-amber'}">${pagata?'Pagata':'Aperta'}</span></td>
      <td style="padding:7px 8px;font-size:11px;color:var(--text3)">${f.conto||'—'}</td>
      <td style="padding:7px 8px;text-align:center">${link?'<a href="'+link+'" target="_blank" style="font-size:14px;text-decoration:none;opacity:.6">📁</a>':'—'}</td>
    </tr>`;
  }).join('');
}


function fattUrl(f){
  if(!f||!f.numero||!f.data)return null;
  const anno=f.data.slice(0,4);
  const mese=f.data.slice(5,7);
  const base='https://cloud.asglab.it/apps/files/?dir=';
  if(f.tipo==='FE'){
    // /Amministrazione/Contabilità/Fatture attive/Fatture attive ANNO/ANNO-MM/
    const dir=`/Amministrazione/Contabilità/Fatture attive/Fatture attive ${anno}/${anno}-${mese}`;
    return base+encodeURIComponent(dir);
  } else if(f.tipo==='FA'){
    // /Amministrazione/Contabilità/Fatture passive/Fatture passive ANNO/ANNO-MM/
    const dir=`/Amministrazione/Contabilità/Fatture passive/Fatture passive ${anno}/${anno}-${mese}`;
    return base+encodeURIComponent(dir);
  }
  return null;
}

// Aggiunge automaticamente scadenze da fatture importate (anti-duplicati)
function aggiungiScadenzeDaFatture(fatture){
  const scadenze=lll('scad');
  // Chiave unicità: tipo_data_scad_id_f24 (NO importo — così reimportando si aggiorna)
  const esistentiMap=new Map(scadenze.filter(s=>s._fatt_key).map(s=>[s._fatt_key,s]));
  let aggiunte=0;
  const nuove=[];
  fatture.forEach(f=>{
    if(!f.data_scad||f.data_scad==='')return;
    if(f.stato==='pagata')return;
    const key=`${f.tipo}_${f.data_scad}_${f.id_f24}`; // senza importo
    const isFE=f.tipo==='FE';
    const nuovaScad={
      id:esistentiMap.has(key)?(esistentiMap.get(key).id):uid(),
      data:f.data_scad,
      descrizione:isFE
        ?`Incasso fattura ${f.numero||''} — ${f.cliente||''}`
        :`Pagamento fattura ${f.numero||''} — ${f.cliente||''}`,
      importo:f.importo||'',
      categoria:isFE?'incasso':'pagamento',
      pagato:'no',
      _fiscale:false,
      _fatt_key:key,
    };
    if(esistentiMap.has(key)){
      // Aggiorna importo se cambiato (fix bug precedente)
      esistentiMap.set(key,nuovaScad);
    } else {
      nuove.push(nuovaScad);
      aggiunte++;
    }
  });
  // Ricostruisce array: scadenze non-fattura + scadenze-fattura aggiornate + nuove
  const nonFatt=scadenze.filter(s=>!s._fatt_key);
  const fatt=[...esistentiMap.values(),...nuove];
  ls('scad',[...nonFatt,...fatt]);
  renderScad();
  return aggiunte;
}

function renderCassa(){
  const data=SD.cassa||lll('conti');
  const tot=data.reduce((s,c)=>s+parseFloat(c.saldo||0),0);

  // Proiezione da fatture (leggi sempre fresco da localStorage)
  const fattRaw=localStorage.getItem(K.fatt);
  const fatt=fattRaw?JSON.parse(fattRaw):(SD.fatt||[]);
  const parseImp=v=>{
    if(!v)return 0;
    if(typeof v==='number')return v;
    const s=String(v).trim();
    // Formato italiano con virgola decimale: "1.234,56" → ha virgola + max 2 cifre dopo
    if(/,\d{1,2}$/.test(s))return parseFloat(s.replace(/\./g,'').replace(',','.'))||0;
    // Già formato anglosassone o numero puro: "782.12", "782" → parse diretto
    return parseFloat(s.replace(/,/g,''))||0;
  };
  const daIncassare=fatt.filter(f=>f.tipo==='FE'&&f.stato==='nonpagata'&&f.data_scad);
  const daPagare=fatt.filter(f=>f.tipo==='FA'&&f.stato==='nonpagata'&&f.data_scad);
  const totIncasso=daIncassare.reduce((s,f)=>s+parseImp(f.importo),0);
  const totUscite=daPagare.reduce((s,f)=>s+parseImp(f.importo),0);
  const scad=lll('scad');
  const fiscali=scad.filter(s=>s._fiscale&&parseFloat(s.importo||0)>0&&dTo(s.data)>=0&&dTo(s.data)<=90);
  const totFiscali=fiscali.reduce((s,x)=>s+parseFloat(x.importo||0),0);

  const colori={'sella':'#1a56db','fineco':'#e55c1b','revolut':'#191c1f','cassa':'#166534','contanti':'#166534'};
  const colConto=n=>{const nl=(n||'').toLowerCase();for(const[k,v]of Object.entries(colori))if(nl.includes(k))return v;return'#64748b';};

  const g=document.getElementById('cassa-grid');
  const pEl=document.getElementById('cassa-proiezione');

  if(!g)return;
  if(!data.length){
    g.innerHTML=`<div style="grid-column:1/-1;padding:14px;text-align:center;font-size:13px;color:var(--text3)">Nessun conto — <a onclick="openModal('conto')" style="color:var(--blue);cursor:pointer;font-weight:600">aggiungi saldo</a></div>`;
    if(pEl)pEl.innerHTML='';
    return;
  }

  // Saldi conti
  g.innerHTML=data.map(c=>`<div class="cc" style="border-left:3px solid ${colConto(c.conto||c.nome||'')}">
    <div class="ccl">${c.conto||c.nome||'Conto'}</div>
    <div class="ccv" style="color:${parseFloat(c.saldo)<0?'var(--red)':'var(--text)'}">€${fmt(c.saldo)}</div>
    <div class="ccd">${c.data_rilevazione?'agg. '+fmtD(c.data_rilevazione):''}</div>
  </div>`).join('')
  +`<div class="cassa-tot"><span style="font-size:12px;font-weight:700;color:var(--blue)">Liquidità totale</span><span style="font-size:20px;font-weight:800;color:var(--blue)">€${fmt(tot)}</span></div>`;

  // Proiezione separata
  if(pEl){
    pEl.innerHTML=(totIncasso>0||totUscite>0||totFiscali>0)?`
    <div style="background:var(--bg2);border:1px solid var(--border);border-radius:var(--r);padding:12px 14px;margin-top:12px;">
      <div style="font-size:10px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.07em;margin-bottom:8px">Proiezione liquidità (fatture aperte)</div>
      <div style="display:flex;gap:16px;flex-wrap:wrap;margin-bottom:8px;">
        ${totIncasso>0?'<div><div style="font-size:10px;color:var(--text3)">Da incassare (${daIncassare.length} FE)</div><div style="font-size:16px;font-weight:700;color:var(--green)">+€${fmt(totIncasso)}</div></div>':''}
        ${totUscite>0?'<div><div style="font-size:10px;color:var(--text3)">Da pagare (${daPagare.length} FA)</div><div style="font-size:16px;font-weight:700;color:var(--red)">−€${fmt(totUscite)}</div></div>':''}
        ${totFiscali>0?'<div><div style="font-size:10px;color:var(--text3)">Scadenze fiscali 90gg</div><div style="font-size:16px;font-weight:700;color:var(--amber)">−€${fmt(totFiscali)}</div></div>':''}
      </div>
      <div style="border-top:1px solid var(--border);padding-top:8px;display:flex;justify-content:space-between;align-items:center;">
        <span style="font-size:12px;color:var(--text3)">Proiezione netta</span>
        <span style="font-size:18px;font-weight:800;color:${(tot+totIncasso-totUscite-totFiscali)>=0?'var(--green)':'var(--red)'}">€${fmt(tot+totIncasso-totUscite-totFiscali)}</span>
      </div>
    </div>`:
    `<div style="padding:8px 0;font-size:12px;color:var(--text4)">Importa le fatture per vedere la proiezione liquidità</div>`;
  }
}

function renderTodo(){
  const data=lll('todo');
  const sorted=[...data].sort((a,b)=>(a.done?1:0)-(b.done?1:0)||({'alta':0,'media':1,'bassa':2}[a.pri]||1)-({'alta':0,'media':1,'bassa':2}[b.pri]||1));
  const list=document.getElementById('todo-list');
  if(!sorted.length){list.innerHTML='<div class="empty">Nessuna azione pendente</div>';return;}
  list.innerHTML=sorted.map(t=>`<div class="item"><input type="checkbox" class="tcb" ${t.done?'checked':''} onchange="togTodo('${t.id}')"><div class="ttx ${t.done?'done':''}">${t.testo}</div>${t.pri&&t.pri!=='bassa'?'<span class="tpr pr-${t.pri}">${t.pri}</span>':''}<button class="xbtn" onclick="delLoc('todo','${t.id}')">✕</button></div>`).join('');
}


// ═══════════════════════════════════════════════
// DETTAGLIO
// ═══════════════════════════════════════════════
let CC=null; // current commessa

function openDetail(numero){
  const all=SD.comm||lll('comm');
  const c=all.find(x=>x.numero===numero);
  if(!c)return;
  CC={...c,_det:(ll('det')[numero])||{distinta:[],ore:[],note:''}};
  document.getElementById('det-num').textContent=numero;
  document.getElementById('det-title').textContent=c.oggetto||'—';
  const tipo=c.tipo_commessa||'';
  const stato=c.stato||'bozza';

  // Flow bar stati
  const FLOW=['bozza','offerta','inviata','confermata','in-corso','chiusa'];
  const FLOW_LABEL={bozza:'Bozza',offerta:'Offerta',inviata:'Inviata',confermata:'Confermata','in-corso':'In lavoro',chiusa:'Chiusa'};
  const FLOW_DESC={
    bozza:'Commessa aperta, offerta non ancora preparata',
    offerta:'Offerta preparata, non ancora inviata al cliente',
    inviata:'Offerta inviata, in attesa di risposta',
    confermata:'Cliente ha confermato — ordina il materiale',
    'in-corso':'Lavorazione attiva',
    chiusa:'Commessa conclusa e fatturata'
  };
  const curIdx=FLOW.indexOf(stato);
  const flowHTML=`<div style="display:flex;align-items:center;gap:0;margin-top:8px;flex-wrap:nowrap;overflow-x:auto;">
    ${FLOW.map((s,i)=>{
      const active=s===stato;
      const done=i<curIdx;
      const bg=active?'var(--blue)':done?'var(--bg3)':'var(--bg2)';
      const col=active?'#fff':done?'var(--text3)':'var(--text4)';
      const bord=active?'var(--blue)':done?'var(--border)':'var(--border)';
      return`<div onclick="cambiaStatoComm('${numero}','${s}')" title="${FLOW_DESC[s]}"
        style="padding:4px 10px;font-size:11px;font-weight:${active?'700':'500'};
        background:${bg};color:${col};border:1px solid ${bord};cursor:pointer;
        white-space:nowrap;border-radius:0;
        ${i===0?'border-radius:20px 0 0 20px':''}
        ${i===FLOW.length-1?'border-radius:0 20px 20px 0':''}
        border-right:${i<FLOW.length-1?'none':'1px solid '+bord};">
        ${active?'● ':''}${FLOW_LABEL[s]}
      </div>`;
    }).join('')}
  </div>
  <div style="font-size:11px;color:var(--text3);margin-top:4px">${FLOW_DESC[stato]}</div>`;

  document.getElementById('det-meta').innerHTML=`
    ${tipo?'<span class="bdg ${TIPO_C[tipo]||\'\'}">${TIPI[tipo]||tipo}</span>':''}
    <span style="font-size:13px;color:var(--text2);font-weight:500">${CNAMES[getCod(numero)]||c.cliente||'—'}</span>
    ${c.data_apertura?'<span style="font-size:11px;color:var(--text3)">aperta ${fmtD(c.data_apertura)}</span>':''}
    ${flowHTML}
  `;

  const cfg=getCfg();
  const fu=folderUrl(numero,cfg.nc_url);
  const fl=document.getElementById('det-folder-link');
  if(fu){fl.href=fu;fl.style.display='';}else fl.style.display='none';
  buildTabs(tipo,c,CC._det);
  navTo('detail');
  setTimeout(_updateMargine,30);
}

function cambiaStatoComm(numero,nuovoStato){
  const all=lll('comm');
  const idx=all.findIndex(x=>x.numero===numero);
  if(idx<0)return;
  all[idx].stato=nuovoStato;
  ls('comm',all);
  if(SD.comm){const i=SD.comm.findIndex(x=>x.numero===numero);if(i>=0)SD.comm[i].stato=nuovoStato;}
  CC={...CC,...all[idx]};
  // Aggiorna UI senza riaprire il dettaglio
  openDetail(numero);
  // Sync in background
  wPut('commesse.csv',all,['numero','cliente','oggetto','tipo_commessa','importo','data_apertura','data_consegna','stato','note']).catch(()=>{});
}

function buildTabs(tipo,c,cd){
  const tabsEl=document.getElementById('det-tabs');
  const contEl=document.getElementById('det-tab-contents');
  const allTabs=[
    {id:'riepilogo',label:'Riepilogo',always:true},
    {id:'offerta',label:'📄 Offerta',always:true},
    {id:'distinta',label:'Distinta',types:['realizzazione','configuratore','misto','ricerca']},
    {id:'approv',label:'Approvvigionamento',types:['realizzazione','ricerca','configuratore','misto']},
    {id:'ore',label:'Ore & trasferte',types:['prestazione','misto']},
    {id:'documenti',label:'📎 Documenti',always:true},
    {id:'note',label:'Note',always:true},
  ];
  const vis=allTabs.filter(t=>t.always||(t.types&&(t.types.includes(tipo)||tipo==='')));
  tabsEl.innerHTML=vis.map((t,i)=>`<div class="tab${i===0?' active':''}" onclick="switchTab('${t.id}')" id="tab-${t.id}">${t.label}</div>`).join('');
  contEl.innerHTML=vis.map((t,i)=>`<div class="tab-content${i===0?' active':''}" id="tc-${t.id}"></div>`).join('');
  vis.forEach(t=>renderTabContent(t.id,c,cd));
}

function switchTab(id){
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));
  document.querySelectorAll('.tab-content').forEach(t=>t.classList.remove('active'));
  const ta=document.getElementById('tab-'+id);const tc=document.getElementById('tc-'+id);
  if(ta)ta.classList.add('active');if(tc)tc.classList.add('active');
}

function renderTabContent(id,c,cd){
  const el=document.getElementById('tc-'+id);if(!el)return;
  if(id==='riepilogo')el.innerHTML=tabRiepilogo(c,cd);
  else if(id==='offerta')el.innerHTML=tabOfferta(c,cd);
  else if(id==='distinta')el.innerHTML=tabDistinta(c,cd);
  else if(id==='approv')el.innerHTML=tabApprov(c,cd);
  else if(id==='ore')el.innerHTML=tabOre(c,cd);
  else if(id==='documenti')el.innerHTML=tabDocumenti(c,cd);
  else if(id==='note')el.innerHTML=tabNote(c,cd);
}

function tabRiepilogo(c,cd){
  const dist=cd.distinta||[];
  const totC=dist.reduce((s,r)=>s+(parseFloat(r.netto_cad||r.costo_cad||0)*parseFloat(r.qty||1)),0);
  const totV=dist.reduce((s,r)=>s+(parseFloat(r.prezzo_vendita||0)*parseFloat(r.qty||1)),0);
  const marg=totV-totC;const percM=totV>0?Math.round((marg/totV)*100):0;
  const totOre=(cd.ore||[]).reduce((s,o)=>s+parseFloat(o.ore||0),0);
  const arrivati=dist.filter(r=>r.stato_approv==='arrivato').length;
  return`<div class="riepilogo">
    <div class="rc" style="border-left:3px solid var(--red)"><div class="rcl">Costo acquisto</div><div class="rcv" style="color:var(--red)">€${fmt2(totC)}</div><div class="rcs">${dist.length} componenti · ${arrivati} arrivati</div></div>
    <div class="rc" style="border-left:3px solid var(--green)"><div class="rcl">Prezzo vendita</div><div class="rcv" style="color:var(--green)">€${fmt2(totV)}</div><div class="rcs">margine ${percM}%</div></div>
    <div class="rc" style="border-left:3px solid var(--blue)"><div class="rcl">Margine lordo</div><div class="rcv" style="color:${marg>=0?'var(--green)':'var(--red)'}">€${fmt2(marg)}</div><div class="rcs">${totOre>0?totOre+' ore lavoro':''}</div></div>
  </div>
  <div class="panel"><div class="info-grid">
    ${[['Numero',c.numero],['Tipo',TIPI[c.tipo_commessa]||'—'],['Stato',SL[c.stato]||'—'],['Cliente',CNAMES[getCod(c.numero)]||c.cliente||'—'],['Apertura',fmtD(c.data_apertura)],['Consegna prev.',fmtD(c.data_consegna)]].map(([l,v])=>`<div><div class="ir-l">${l}</div><div class="ir-v">${v}</div></div>`).join('')}
    ${c.note?'<div style="grid-column:1/-1"><div class="ir-l">Note</div><div class="ir-v">${c.note}</div></div>':''}
  </div></div>`;
}

function tabOfferta(c,cd){
  const off=(cd&&cd.offerta)||{stato:'bozza',testo:'',data_invio:'',data_risposta:'',note_cliente:'',importo_offerto:c.importo||''};
  const stati={bozza:'📝 Bozza',inviata:'📤 Inviata',approvata:'✅ Approvata',rifiutata:'❌ Rifiutata',scaduta:'⏱ Scaduta'};
  const colori={bozza:'#57606a',inviata:'#1a56db',approvata:'#166534',rifiutata:'#991b1b',scaduta:'#78350f'};
  const dist=(cd&&cd.distinta)||[];
  const totV=dist.reduce((s,r)=>s+(parseFloat(r.prezzo_vendita||0)*parseFloat(r.qty||1)),0);
  const cliente=CNAMES[getCod(c.numero)]||c.cliente||'Spettabile Cliente';
  const testoAuto=`Spettabile ${cliente},\n\nIn riferimento alla Vs. richiesta, siamo lieti di sottoporVi la nostra migliore offerta per:\n\n${c.oggetto||'— oggetto —'}\n\n`+
    (dist.filter(r=>!r._intestazione).length?dist.filter(r=>!r._intestazione).map(r=>`• ${r.descrizione||r.codice} (qty ${r.qty||1}) — € ${fmt2(parseFloat(r.prezzo_vendita||0)*parseFloat(r.qty||1))} IVA esclusa`).join('\n')+'\n\n':'')+
    `TOTALE OFFERTA: € ${fmt2(totV||parseFloat(off.importo_offerto||c.importo||0))} IVA esclusa\n\nLa presente offerta è valida 30 giorni dalla data odierna.\nTempi di consegna: da concordare.\nCondizioni di pagamento: 30 giorni d.f.f.m.\n\nPer qualsiasi informazione restiamo a Vostra disposizione.\n\nCordiali saluti,`;
  const statoColor=colori[off.stato]||'#57606a';
  return`
  <div style="display:flex;align-items:center;gap:12px;margin-bottom:16px;flex-wrap:wrap;">
    <div style="flex:1;">
      <div style="font-size:11px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.07em;margin-bottom:4px">Stato offerta</div>
      <select id="off-stato" style="border:2px solid ${statoColor};border-radius:20px;padding:6px 16px;font-size:13px;font-weight:700;color:${statoColor};background:#fff;cursor:pointer;outline:none;" onchange="aggiornaColoreBordoOfferta(this)">
        ${Object.entries(stati).map(([k,v])=>`<option value="${k}"${off.stato===k?' selected':''}>${v}</option>`).join('')}
      </select>
    </div>
    <div style="display:flex;gap:8px;flex-wrap:wrap;">
      <button class="btn-ghost" onclick="generaTestoOfferta('${c.numero}')">♻ Rigenera da distinta</button>
      <button class="btn btn-sm" onclick="salvaOfferta('${c.numero}')">💾 Salva</button>
      <button class="btn btn-sm" style="background:#166534;" onclick="inviaOffertaMail('${c.numero}')">📧 Invia al cliente</button>
    </div>
  </div>
  <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;margin-bottom:16px;">
    <div class="rc" style="border-left:3px solid var(--green)"><div class="rcl">Importo offerto (€)</div><input id="off-importo" type="number" step="0.01" value="${off.importo_offerto||totV||c.importo||''}" style="font-size:18px;font-weight:800;width:100%;border:none;background:transparent;color:var(--green);outline:none;font-family:var(--font);"></div>
    <div class="rc" style="border-left:3px solid var(--blue)"><div class="rcl">Data invio</div><input id="off-data-invio" type="date" value="${off.data_invio||''}" style="font-size:14px;font-weight:600;width:100%;border:none;background:transparent;outline:none;font-family:var(--font);"></div>
    <div class="rc" style="border-left:3px solid var(--violet)"><div class="rcl">Risposta cliente</div><input id="off-data-risposta" type="date" value="${off.data_risposta||''}" style="font-size:14px;font-weight:600;width:100%;border:none;background:transparent;outline:none;font-family:var(--font);"></div>
  </div>
  <div style="margin-bottom:10px;">
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:6px;">
      <label style="font-size:12px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.07em">Testo offerta</label>
      <span style="font-size:11px;color:var(--text3)">Modificabile liberamente</span>
    </div>
    <textarea id="off-testo" style="width:100%;min-height:280px;border:1.5px solid var(--border);border-radius:8px;padding:14px;font-size:13px;line-height:1.8;color:var(--text2);font-family:var(--font);outline:none;resize:vertical;">${off.testo||testoAuto}</textarea>
  </div>
  <div style="margin-bottom:14px;">
    <label style="font-size:12px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.07em;display:block;margin-bottom:6px;">Note interne — risposta cliente</label>
    <textarea id="off-note-cliente" style="width:100%;min-height:60px;border:1.5px solid var(--border);border-radius:8px;padding:10px 12px;font-size:13px;line-height:1.6;color:var(--text2);font-family:var(--font);outline:none;resize:vertical;" placeholder="Annotazioni interne su feedback, trattative, revisioni richieste…">${off.note_cliente||''}</textarea>
  </div>
  ${off.storico&&off.storico.length?'<div style="background:var(--bg2);border:1px solid var(--border);border-radius:8px;padding:12px 14px;">'+
    '<div style="font-size:11px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.07em;margin-bottom:8px">📋 Storico</div>'+
    off.storico.slice(-6).reverse().map(s=>'<div style="font-size:12px;color:var(--text3);padding:4px 0;border-bottom:1px solid var(--border);">'+s.data.slice(0,10)+' — <strong>'+(stati[s.stato]||s.stato)+'</strong>'+(s.nota?' · '+s.nota:'')+' </div>').join('')+
  '</div>':''}`;
}

function aggiornaColoreBordoOfferta(sel){
  const cc={bozza:'#57606a',inviata:'#1a56db',approvata:'#166534',rifiutata:'#991b1b',scaduta:'#78350f'};
  const col=cc[sel.value]||'#57606a';sel.style.borderColor=col;sel.style.color=col;
}

function generaTestoOfferta(numero){
  const c=CC;if(!c)return;
  const cd=CC._det||{distinta:[]};
  const dist=(cd.distinta||[]).filter(r=>!r._intestazione);
  const totV=dist.reduce((s,r)=>s+(parseFloat(r.prezzo_vendita||0)*parseFloat(r.qty||1)),0);
  const cliente=CNAMES[getCod(c.numero)]||c.cliente||'Spettabile Cliente';
  const testo=`Spettabile ${cliente},\n\nIn riferimento alla Vs. richiesta, siamo lieti di sottoporVi la nostra migliore offerta per:\n\n${c.oggetto||'— oggetto —'}\n\n`+
    (dist.length?dist.map(r=>`• ${r.descrizione||r.codice} (qty ${r.qty||1}) — € ${fmt2(parseFloat(r.prezzo_vendita||0)*parseFloat(r.qty||1))} IVA esclusa`).join('\n')+'\n\n':'')+
    `TOTALE OFFERTA: € ${fmt2(totV||parseFloat(c.importo||0))} IVA esclusa\n\nLa presente offerta è valida 30 giorni dalla data odierna.\nTempi di consegna: da concordare.\nCondizioni di pagamento: 30 giorni d.f.f.m.\n\nPer qualsiasi informazione restiamo a Vostra disposizione.\n\nCordiali saluti,`;
  const ta=document.getElementById('off-testo');if(ta)ta.value=testo;
  showToast('Testo rigenerato dalla distinta');
}

function salvaOfferta(numero){
  const det=ll('det');if(!det[numero])det[numero]={distinta:[],ore:[],note:''};
  const old=det[numero].offerta||{};
  const nuovoStato=(document.getElementById('off-stato')||{}).value||'bozza';
  const storico=old.storico||[];
  if(nuovoStato!==old.stato)storico.push({data:new Date().toISOString(),stato:nuovoStato,nota:''});
  det[numero].offerta={stato:nuovoStato,testo:(document.getElementById('off-testo')||{}).value||'',data_invio:(document.getElementById('off-data-invio')||{}).value||'',data_risposta:(document.getElementById('off-data-risposta')||{}).value||'',note_cliente:(document.getElementById('off-note-cliente')||{}).value||'',importo_offerto:(document.getElementById('off-importo')||{}).value||'',storico};
  ls('det',det);syncDetBackground();CC._det=det[numero];
  if(nuovoStato==='approvata'){const all=lll('comm');const idx=all.findIndex(x=>x.numero===numero);if(idx>=0&&all[idx].stato==='offerta'){all[idx].stato='confermata';ls('comm',all);if(SD.comm){const i=SD.comm.findIndex(x=>x.numero===numero);if(i>=0)SD.comm[i].stato='confermata';}renderComm();}}
  showToast('✓ Offerta salvata');
}

function inviaOffertaMail(numero){
  const c=CC;if(!c)return;
  const testo=(document.getElementById('off-testo')||{}).value||'';
  const oggetto=`Offerta ${numero} — ${c.oggetto||''}`;
  const statoEl=document.getElementById('off-stato');
  if(statoEl&&statoEl.value==='bozza'){statoEl.value='inviata';aggiornaColoreBordoOfferta(statoEl);}
  const dataEl=document.getElementById('off-data-invio');if(dataEl&&!dataEl.value)dataEl.value=new Date().toISOString().slice(0,10);
  salvaOfferta(numero);
  window.open(`mailto:?subject=${encodeURIComponent(oggetto)}&body=${encodeURIComponent(testo)}`);
}

function tabDistinta(c,cd){
  const dist=cd.distinta||[];
  const sezioni=[...new Set(dist.map(r=>r.sezione||'Generale'))];
  const rows=sezioni.map(s=>{
    const items=dist.filter(r=>(r.sezione||'Generale')===s);
    const tot=items.reduce((sum,r)=>sum+(parseFloat(r.netto_cad||r.costo_cad||0)*parseFloat(r.qty||1)),0);
    return`<tr><td colspan="11" class="sez-head">${s}<span style="color:var(--blue);font-size:11px;font-weight:700">€${fmt2(tot)}</span></td></tr>`
      +items.map(r=>`<tr>
        <td style="max-width:170px;font-size:12px">${r.descrizione||'—'}</td>
        <td class="mono" style="text-align:center">${r.qty||1}</td>
        <td style="font-size:12px">${r.fornitore||'—'}</td>
        <td class="mono">${r.codice||'—'}</td>
        <td class="mono">€${fmt2(r.costo_cad||0)}</td>
        <td class="mono">${r.sconto?r.sconto+'%':'—'}</td>
        <td class="mono" style="font-weight:700">€${fmt2(r.netto_cad||r.costo_cad||0)}</td>
        <td class="mono" style="color:var(--green);font-weight:700">€${fmt2(r.prezzo_vendita||0)}</td>
        <td><span class="bdg a-${r.stato_approv||'daordinare'}"><select class="stato-sel" onchange="updStatoApprov('${r.id}','${c.numero}',this.value)">${Object.entries(SAP).map(([k,v])=>`<option value="${k}"${r.stato_approv===k?' selected':''}>${v}</option>`).join('')}</select></span></td>
        <td><span onclick="this.nextSibling.style.display='inline';this.style.display='none'" style="font-size:11px;color:var(--text3);cursor:pointer;padding:2px 4px;border-radius:4px;border:1px solid transparent;" onmouseover="this.style.borderColor='var(--border)'" onmouseout="this.style.borderColor='transparent'">${r.data_attesa?fmtD(r.data_attesa):'—'}</span><input type="date" value="${r.data_attesa||''}" onchange="updDataAttesa('${r.id}','${c.numero}',this.value);this.previousSibling.textContent=fmtD(this.value)||'—';this.style.display='none';this.previousSibling.style.display=''" style="display:none;border:1px solid var(--border);border-radius:5px;font-size:11px;padding:2px 5px;width:120px;" onblur="this.style.display='none';this.previousSibling.style.display=''"></td>
        <td style="white-space:nowrap;">
          <button class="xbtn" style="color:var(--blue)" title="Modifica" onclick="openModal('edit-riga','${c.numero}|${r.id}');event.stopPropagation()">✎</button>
          <button class="xbtn" title="Elimina" onclick="delDistinta('${r.id}','${c.numero}');event.stopPropagation()">✕</button>
        </td>
      </tr>`).join('');
  }).join('');
  // Suggerimento leggero solo se distinta vuota
  const emptyHint=!dist.length?`<div style="background:var(--bg2);border:1px solid var(--border);border-radius:8px;padding:10px 14px;margin-bottom:12px;font-size:12px;color:var(--text3);display:flex;align-items:center;gap:8px;"><span>💡</span><span>Usa <strong>+ aggiungi componente</strong> per cercare nel listino Ares, nel magazzino o configurare un tubo flessibile.</span></div>`:'';
  return`${emptyHint}<div class="dist-toolbar"><button class="btn btn-sm" onclick="openModal('add-riga','${c.numero}')">+ aggiungi componente</button><button class="btn-ghost" style="font-size:11px;padding:4px 10px;" onclick="openModal('add-sezione','${c.numero}')">+ sezione</button><button class="btn-ghost" style="font-size:11px;padding:4px 10px;" onclick="exportDistintaPDF('${c.numero}')">📄 Stampa distinta</button></div>
  <div style="overflow-x:auto"><table>
    <thead><tr><th>Descrizione</th><th>Q.</th><th>Fornitore</th><th>Codice</th><th>Costo/cad</th><th>Sc.</th><th>Netto/cad</th><th>Vendita/cad</th><th>Stato</th><th>Data attesa</th><th></th></tr></thead>
    <tbody>${rows||'<tr><td colspan="11" class="empty">Nessun componente</td></tr>'}</tbody>
  </table></div>`;
}


function useAres(a,numero){
  const netto=a.p;const markup=2.25;const vendita=parseFloat((netto*markup).toFixed(4));
  const row={id:uid(),sezione:'Centralina',descrizione:a.d,qty:1,fornitore:'Ares / Tecfluid',codice:a.s,costo_cad:netto,sconto:0,netto_cad:netto,markup,prezzo_vendita:vendita,stato_approv:'daordinare',data_attesa:''};
  const det=ll('det');if(!det[numero])det[numero]={distinta:[],ore:[],note:''};
  det[numero].distinta.push(row);ls('det',det);
  CC._det=det[numero];
  ['distinta','approv','riepilogo'].forEach(id=>renderTabContent(id,CC,det[numero]));
}

function tabApprov(c,cd){
  const dist=cd.distinta||[];
  if(!dist.length)return`<div class="empty">Aggiungi componenti nella distinta</div>`;
  const fornitori={};
  dist.forEach(r=>{const f=r.fornitore||'Non assegnato';if(!fornitori[f])fornitori[f]=[];fornitori[f].push(r);});
  const cardsForn=Object.entries(fornitori).map(([forn,items])=>{
    const tot=items.reduce((s,r)=>s+(parseFloat(r.netto_cad||r.costo_cad||0)*parseFloat(r.qty||1)),0);
    const tutti=items.every(r=>r.stato_approv==='arrivato');
    const qualcuno=items.some(r=>r.stato_approv==='arrivato');
    const prob=items.some(r=>r.stato_approv==='problema');
    const stk=prob?'a-problema':tutti?'a-arrivato':qualcuno?'a-ordinato':'a-daordinare';
    const stl=prob?'Problema':tutti?'Tutto arrivato':qualcuno?'Parziale':'Da ordinare';
    // Trova ID fornitore per RDO
    const fornObj=FORNITORI.find(f2=>f2.nome.toLowerCase().includes(forn.toLowerCase().split(' ')[0])||forn.toLowerCase().includes(f2.nome.toLowerCase().split(' ')[0]));
    const rdoBtn=`<button class="btn-ghost" style="font-size:11px;padding:3px 10px;" onclick="openRDOApprov('${c.numero}','${forn.replace(/'/g,"\\'")}')">✉ RDO a ${forn.split(' ')[0]}</button>`;
    return`<div class="approv-card">
      <div class="approv-head" style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;">
        <span class="approv-name">${forn}</span>
        <span class="bdg ${stk}">${stl}</span>
        <span class="approv-meta">${items.length} art. · €${fmt2(tot)}</span>
        <div style="margin-left:auto">${rdoBtn}</div>
      </div>
    <table><thead><tr><th>Descrizione</th><th>Q.</th><th>Codice</th><th>€ tot</th><th>Stato</th><th>Data attesa</th></tr></thead>
    <tbody>${items.map(r=>`<tr>
      <td style="font-size:12px;max-width:180px">${r.descrizione||'—'}</td>
      <td class="mono" style="text-align:center">${r.qty||1}</td>
      <td class="mono">${r.codice||'—'}</td>
      <td class="mono" style="font-weight:700">€${fmt2(parseFloat(r.netto_cad||r.costo_cad||0)*parseFloat(r.qty||1))}</td>
      <td><span class="bdg a-${r.stato_approv||'daordinare'}"><select class="stato-sel" onchange="updStatoApprov('${r.id}','${c.numero}',this.value)">${Object.entries(SAP).map(([k,v])=>`<option value="${k}"${r.stato_approv===k?' selected':''}>${v}</option>`).join('')}</select></span></td>
      <td><span onclick="this.nextSibling.style.display='inline';this.style.display='none'" style="font-size:11px;color:var(--text3);cursor:pointer;padding:2px 4px;border-radius:4px;border:1px solid transparent;" onmouseover="this.style.borderColor='var(--border)'" onmouseout="this.style.borderColor='transparent'">${r.data_attesa?fmtD(r.data_attesa):'—'}</span><input type="date" value="${r.data_attesa||''}" onchange="updDataAttesa('${r.id}','${c.numero}',this.value);this.previousSibling.textContent=fmtD(this.value)||'—';this.style.display='none';this.previousSibling.style.display=''" style="display:none;border:1px solid var(--border);border-radius:5px;font-size:11px;padding:2px 5px;width:120px;" onblur="this.style.display='none';this.previousSibling.style.display=''"></td>
    </tr>`).join('')}</tbody></table></div>`;
  }).join('');
  return cardsForn;
}

function openRDOApprov(numero,fornNome){
  // Costruisce un testo RDO precompilato con tutti i componenti del fornitore
  const det=ll('det');const cd=det[numero]||{distinta:[]};
  const items=(cd.distinta||[]).filter(r=>!r._intestazione&&(r.fornitore||'Non assegnato')===fornNome);
  const c=(SD.comm||lll('comm')).find(x=>x.numero===numero)||{numero,oggetto:''};
  const oggi=new Date().toLocaleDateString('it-IT',{day:'2-digit',month:'long',year:'numeric'});
  const templates=JSON.parse(localStorage.getItem('asg_rdo_templates')||'{}');
  const fornObj=FORNITORI.find(f=>f.nome.toLowerCase().includes(fornNome.toLowerCase().split(' ')[0])||fornNome.toLowerCase().includes(f.nome.toLowerCase().split(' ')[0]));
  const fornId=fornObj?fornObj.id:'';
  const baseTemplate=templates[fornId]||`Spett.le ${fornNome},\n\nin riferimento ai Vs. prodotti, con la presente siamo a richiedere conferma di disponibilità e tempi di consegna per i seguenti articoli, relativi alla commessa ${numero} — ${c.oggetto||''}:\n\n`;
  const righe=items.map((r,i)=>`${i+1}. ${r.codice?'['+r.codice+'] ':''}${r.descrizione||'—'} — qty ${r.qty||1}`).join('\n');
  const totNetto=items.reduce((s,r)=>s+(parseFloat(r.netto_cad||r.costo_cad||0)*parseFloat(r.qty||1)),0);
  const testo=baseTemplate+righe+`\n\nSi prega di indicare:\n- Prezzi unitari IVA esclusa\n- Disponibilità a magazzino\n- Tempi di consegna\n- Condizioni di resa\n\nValore indicativo ordine: € ${fmt2(totNetto)} IVA esclusa\n\nRingraziamo e restiamo in attesa di riscontro.\n\nCordiali saluti,`;
  const subj=`RDO commessa ${numero} — ${c.oggetto||''}`;
  // Mostra modal
  const email=fornObj?.contatti?.email||'';
  document.getElementById('modal-body').innerHTML=`
    <h3>✉ RDO a ${fornNome} — Commessa ${numero}</h3>
    <div style="font-size:12px;color:var(--text3);margin-bottom:12px;">${items.length} articoli · valore indicativo €${fmt2(totNetto)}</div>
    <div class="form-row"><label>A (email)</label><input id="rdoa-to" value="${email}" placeholder="email@fornitore.it"></div>
    <div class="form-row"><label>Oggetto</label><input id="rdoa-subj" value="${subj}"></div>
    <div class="form-row"><label>Testo</label><textarea id="rdoa-body" style="min-height:260px;font-family:var(--font);font-size:13px;line-height:1.7">${testo}</textarea></div>
    <div class="modal-actions">
      <button class="btn-ghost" onclick="closeM()">Annulla</button>
      <button class="btn-ghost" onclick="navigator.clipboard.writeText(document.getElementById('rdoa-subj').value+'\\n\\n'+document.getElementById('rdoa-body').value).then(()=>showToast('Copiato!'))">📋 Copia</button>
      ${email?'<a class="btn" href="mailto:${email}?subject=${encodeURIComponent(subj)}&body=${encodeURIComponent(testo)}" onclick="closeM()">📧 Apri email</a>':'<button class="btn" onclick="navigator.clipboard.writeText(document.getElementById(\'rdoa-subj\').value+\'\\n\\n\'+document.getElementById(\'rdoa-body\').value).then(()=>{showToast(\'Copiato — incolla nel client email\');closeM();})">📋 Copia tutto</button>'}
    </div>`;
  document.getElementById('modal-overlay').classList.add('show');
}

function tabOre(c,cd){
  const ore=cd.ore||[];
  const totOre=ore.reduce((s,o)=>s+parseFloat(o.ore||0),0);
  const totEur=ore.reduce((s,o)=>s+(parseFloat(o.ore||0)*parseFloat(o.tariffa||0)),0);
  return`<div style="display:flex;gap:10px;margin-bottom:14px;"><button class="btn btn-sm" onclick="openModal('add-ore','${c.numero}')">+ registra ore</button></div>
  <div class="riepilogo" style="margin-bottom:14px;">
    <div class="rc" style="border-left:3px solid var(--violet)"><div class="rcl">Ore totali</div><div class="rcv" style="color:var(--violet)">${fmt(totOre)} h</div></div>
    <div class="rc" style="border-left:3px solid var(--green)"><div class="rcl">Valore ore</div><div class="rcv" style="color:var(--green)">€${fmt2(totEur)}</div></div>
    <div class="rc" style="border-left:3px solid var(--blue)"><div class="rcl">Tariffa media</div><div class="rcv" style="color:var(--blue)">${totOre>0?'€'+fmt2(totEur/totOre)+'/h':'—'}</div></div>
  </div>
  <div style="overflow-x:auto"><table><thead><tr><th>Data</th><th>Descrizione</th><th>Operatore</th><th>Ore</th><th>Tariffa</th><th>Totale</th><th></th></tr></thead>
  <tbody>${ore.length?ore.map(o=>`<tr><td class="mono">${fmtD(o.data)}</td><td style="font-size:12px">${o.descrizione||'—'}</td><td style="font-size:12px">${o.operatore||'—'}</td><td class="mono" style="font-weight:700">${o.ore}h</td><td class="mono">€${fmt2(o.tariffa||0)}/h</td><td class="mono" style="color:var(--green);font-weight:700">€${fmt2(parseFloat(o.ore||0)*parseFloat(o.tariffa||0))}</td><td><button class="xbtn" onclick="delOre('${o.id}','${c.numero}')">✕</button></td></tr>`).join(''):'<tr><td colspan="7" class="empty">Nessuna ora registrata</td></tr>'}</tbody></table></div>`;
}

function tabDocumenti(c,cd){
  const cfg=getCfg();
  const docs=cd.documenti||[];
  const anno=getAnno(c.numero);const cod=getCod(c.numero);const client=CMAP[cod]||cod;
  const ncPath=cod&&anno?'/Clienti/'+client+'/'+anno:null;
  const ncFolder=ncPath?'https://cloud.asglab.it/apps/files/?dir='+encodeURIComponent(ncPath):null;
  const fileListId='nc-files-'+c.numero;
  if(ncPath&&cfg.ok){
    setTimeout(async()=>{
      const el=document.getElementById(fileListId);
      if(!el)return;
      el.innerHTML='<span style="color:var(--text4);font-size:12px">Ricerca cartella commessa…</span>';
      // Cerca la sottocartella della commessa dentro la cartella anno
      // La sottocartella inizia tipicamente con "COD-N-ANNO" o contiene l'oggetto commessa
      let targetPath=ncPath;
      const subfolders=await wListNCDirs(ncPath);
      if(subfolders.length){
        // Numero commessa: KFG-01-2026 → cerca cartella che inizia con KFG-1-2026
        const seqNum=parseInt((c.numero||'').split('-')[1]||'0'); // rimuovi zero padding: 01→1
        const commKey=cod+'-'+seqNum+'-'+anno; // es. KFG-1-2026
        const match=subfolders.find(f=>
          f.name.startsWith(commKey)||
          f.name.toLowerCase().includes((c.oggetto||'').toLowerCase().slice(0,10))||
          f.name.startsWith(cod+'-'+seqNum+' ')
        );
        if(match)targetPath=ncPath+'/'+match.name;
      }
      const folderLinkEl=document.getElementById('nc-folder-link-'+c.numero);
      if(folderLinkEl)folderLinkEl.href='https://cloud.asglab.it/apps/files/?dir='+encodeURIComponent(targetPath);
      el.innerHTML='<span style="color:var(--text4);font-size:12px">Caricamento file…</span>';
      const files=await wListNC(targetPath);
      if(!files.length){el.innerHTML='<span style="color:var(--text4);font-size:12px">Nessun file in '+targetPath+'</span>';return;}
      const ext=n=>{const p=n.lastIndexOf('.');return p>=0?n.slice(p+1).toLowerCase():'';};
      const icon=n=>{const e=ext(n);const m={pdf:'📄',xlsx:'📊',xls:'📊',docx:'📝',doc:'📝',jpg:'🖼',jpeg:'🖼',png:'🖼',dxf:'📐',dwg:'📐',zip:'🗜',rar:'🗜',eml:'📧',msg:'📧'};return m[e]||'📎';};
      const fmtSz=b=>{if(b>1048576)return(b/1048576).toFixed(1)+' MB';if(b>1024)return(b/1024).toFixed(0)+' KB';return b+' B';};
      el.innerHTML=`<div style="display:grid;gap:6px;">${files.map(f=>{
        const ncUrl=cfg.nc_url+'/apps/files/?dir='+encodeURIComponent(ncPath)+'&scrollto='+encodeURIComponent(f.name);
        const dlUrl=cfg.nc_url+'/remote.php/dav/files/'+cfg.nc_user+encodeURI(ncPath+'/'+f.name);
        return`<div style="display:flex;align-items:center;gap:8px;padding:7px 10px;background:var(--bg2);border:1px solid var(--border);border-radius:var(--r-sm);">
          <span style="font-size:18px">${icon(f.name)}</span>
          <div style="flex:1;min-width:0;"><div style="font-size:12px;font-weight:500;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${f.name}</div>
          <div style="font-size:10px;color:var(--text4)">${fmtSz(f.size)} · ${f.mod?new Date(f.mod).toLocaleDateString('it-IT',{day:'2-digit',month:'short',year:'2-digit'}):''}</div></div>
          <a href="${dlUrl}" target="_blank" style="font-size:11px;color:var(--blue);font-weight:600;text-decoration:none;padding:3px 8px;border:1px solid var(--blue-bd);border-radius:20px;">↓</a>
        </div>`;
      }).join('')}</div>`;
    },100);
  }

  return`
  <div style="display:flex;align-items:center;gap:8px;margin-bottom:12px;flex-wrap:wrap;">
    ${ncFolder?'<a id="nc-folder-link-${c.numero}" href="${ncFolder}" target="_blank" class="btn btn-primary btn-sm">📂 Apri cartella NC</a>':''}
    <button class="btn btn-ghost btn-sm" onclick="openModal('add-doc','${c.numero}')">+ Aggiungi riferimento</button>
    ${ncPath&&cfg.ok?'<button class="btn btn-ghost btn-sm" onclick="(async()=>{document.getElementById(\'${fileListId}\').innerHTML=\'<span style=\'font-size:12px;color:var(--text4)\'>Aggiornamento…</span>\';const files=await wListNC(\'${ncPath}\');document.getElementById(\'${fileListId}\').innerHTML=files.length?files.map(f=>\'<div>\'+f.name+\'</div>\').join(\'\'):\'Nessun file\';})()">↻ Aggiorna</button>':''}
  </div>

  ${ncPath&&cfg.ok?'<div style="margin-bottom:12px;">'+
    '<div style="font-size:11px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.06em;margin-bottom:6px">File su Nextcloud \u2014 '+ncPath+'</div>'+
    '<div id="'+fileListId+'"><span style="color:var(--text4);font-size:12px">In caricamento\u2026</span></div>'+
  '</div>':'<div style="background:var(--amber-lt);border:1px solid var(--amber-bd);border-radius:var(--r-sm);padding:10px 12px;font-size:12px;color:var(--amber);margin-bottom:12px;">Configura Nextcloud per vedere i file della commessa</div>'}

  <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:10px;flex-wrap:wrap;gap:8px;">
    <div style="font-size:13px;font-weight:700;color:var(--text)">Riferimenti documenti</div>
    <button class="btn btn-sm" onclick="openModal('add-doc','${c.numero}')">+ aggiungi riferimento</button>
  </div>

  ${docs.length?'<div style="display:grid;gap:8px;">'+
    docs.map(d=>'<div style="background:var(--bg2);border:1px solid var(--border);border-radius:8px;padding:10px 14px;display:flex;align-items:flex-start;gap:12px;">'+
      '<span style="font-size:20px;flex-shrink:0">'+tipoDocEmoji(d.tipo)+'</span>'+
      '<div style="flex:1;">'+
        '<div style="font-size:13px;font-weight:700;color:var(--text);margin-bottom:2px">'+(d.nome||'\u2014')+'</div>'+
        '<div style="font-size:12px;color:var(--text3)">'+(d.tipo||'')+' '+(d.data?'\u00b7 '+fmtD(d.data):'')+' '+(d.note?'\u00b7 '+d.note:'')+'</div>'+
        (d.url?'<a href="'+d.url+'" target="_blank" style="font-size:11px;color:var(--blue);font-weight:600;">&nearr; apri</a>':'')+
      '</div>'+
      '<button class="xbtn" onclick="delDoc(\''+d.id+'\',\''+c.numero+'\')">&times;</button>'+
    '</div>').join('')+
  '</div>':'<div class="empty">Nessun documento collegato. Usa il bottone sopra per aggiungere un riferimento (link Nextcloud, numero schema, nota).</div>'}`;  
}

function tipoDocEmoji(t){const m={schema:'📐',foto:'📷',email:'📧',offerta:'📄',ordine:'🧾',certificato:'📋',disegno:'✏️',altro:'📎'};return m[t]||'📎';}

function delDoc(id,numero){
  const det=ll('det');if(!det[numero])return;
  det[numero].documenti=(det[numero].documenti||[]).filter(d=>d.id!==id);
  ls('det',det);syncDetBackground();CC._det=det[numero];
  renderTabContent('documenti',CC,det[numero]);
}

function saveDoc(numero){
  const det=ll('det');if(!det[numero])det[numero]={distinta:[],ore:[],note:'',documenti:[]};
  if(!det[numero].documenti)det[numero].documenti=[];
  const doc={id:uid(),nome:(document.getElementById('d-nome')||{}).value||'',tipo:(document.getElementById('d-tipo')||{}).value||'altro',url:(document.getElementById('d-url')||{}).value||'',data:(document.getElementById('d-data')||{}).value||'',note:(document.getElementById('d-note')||{}).value||''};
  det[numero].documenti.push(doc);
  ls('det',det);syncDetBackground();CC._det=det[numero];
  closeM();renderTabContent('documenti',CC,det[numero]);
  showToast('✓ Documento aggiunto');
}

function tabNote(c,cd){
  return`<div class="fr"><label>Note interne</label><textarea id="note-area" style="min-height:200px;">${cd.note||''}</textarea></div>
  <div style="margin-top:10px;"><button class="btn btn-sm" onclick="saveNote('${c.numero}')">Salva note</button></div>`;
}

// ═══════════════════════════════════════════════
// AZIONI DETTAGLIO
// ═══════════════════════════════════════════════
function updStatoApprov(rid,num,stato){const det=ll('det');if(!det[num])return;const r=det[num].distinta.find(x=>x.id===rid);if(r){r.stato_approv=stato;ls('det',det);syncDetBackground();}CC._det=det[num];['approv','distinta','riepilogo'].forEach(id=>renderTabContent(id,CC,det[num]));}
function updDataAttesa(rid,num,data){const det=ll('det');if(!det[num])return;const r=det[num].distinta.find(x=>x.id===rid);if(r){r.data_attesa=data;ls('det',det);syncDetBackground();}}
function delDistinta(rid,num){const det=ll('det');if(!det[num])return;det[num].distinta=det[num].distinta.filter(r=>r.id!==rid);ls('det',det);syncDetBackground();CC._det=det[num];['distinta','approv','riepilogo'].forEach(id=>renderTabContent(id,CC,det[num]));setTimeout(()=>upgradeSelects(document.getElementById('det-tab-contents')),50);}
function delOre(id,num){const det=ll('det');if(!det[num])return;det[num].ore=(det[num].ore||[]).filter(o=>o.id!==id);ls('det',det);CC._det=det[num];renderTabContent('ore',CC,det[num]);renderTabContent('riepilogo',CC,det[num]);}
function saveNote(num){const det=ll('det');if(!det[num])det[num]={distinta:[],ore:[],note:''};det[num].note=document.getElementById('note-area').value;ls('det',det);syncDetBackground();const b=event.target;b.textContent='✓ Salvato';setTimeout(()=>b.textContent='Salva note',1500);}
function delLoc(key,id){const arr=lll(key).filter(i=>i.id!==id);ls(key,arr);renderAll();}
function togTodo(id){const d=lll('todo');const i=d.find(t=>t.id===id);if(i){i.done=!i.done;ls('todo',d);renderTodo();}}

async function eliminaCommessa(){
  const c=CC;if(!c)return;
  if(!confirm('Eliminare definitivamente la commessa '+c.numero+'?\n\nQuesta azione non può essere annullata. La distinta e tutti i dati collegati verranno rimossi.'))return;
  // Rimuovi da localStorage
  const all=lll('comm').filter(x=>x.numero!==c.numero);
  ls('comm',all);
  const det=ll('det');delete det[c.numero];ls('det',det);
  // Rimuovi da SD
  if(SD.comm)SD.comm=SD.comm.filter(x=>x.numero!==c.numero);
  // Sincronizza su Nextcloud
  const cfg=getCfg();
  if(cfg.ok){
    await wPut('commesse.csv',all,['numero','cliente','oggetto','tipo_commessa','importo','data_apertura','data_consegna','stato','note']);
  }
  navTo('dashboard');
  renderAll();
  showToast(`✓ Commessa ${c.numero} eliminata`);
}

// ═══════════════════════════════════════════════
// MODAL
// ═══════════════════════════════════════════════
function openModal(type,extra){
  if(type==='add-riga'){currentAddRigaNumero=extra||'';_tfRaSelected={A:null,B:null};_tfStdFiltro='';_tfDnAttuale=0;_raccScioltoFiltro='';_pu10Blocchi=[];_pu20Blocchi=[];}
  document.getElementById('modal-body').innerHTML=getModalHTML(type,extra);
  document.getElementById('modal-overlay').classList.add('show');
  setTimeout(()=>{upgradeSelects();precompilaMarkup();},0);
}

function closeM(){document.getElementById('modal-overlay').classList.remove('show');}
document.addEventListener('keydown',e=>{if(e.key==='Escape')closeM();});

function getModalHTML(type,extra){
  const today=new Date().toISOString().slice(0,10);
  const SC=['bozza','offerta','inviata','confermata','in-corso','chiusa'];
  const act=(ok,cancel='Annulla')=>`<div class="modal-actions"><button class="btn-ghost" onclick="closeM()">${cancel}</button><button class="btn" onclick="${ok}">Salva</button></div>`;

  if(type==='commessa')return`<h3>Nuova commessa</h3><div class="form-inline">
    <div class="fg2"><div class="fr"><label>Numero</label><input id="f-num" placeholder="es. KFG-16-2026"></div><div class="fr"><label>Tipo commessa</label><select id="f-tipo"><option value="">— seleziona —</option>${Object.entries(TIPI).map(([k,v])=>`<option value="${k}">${v}</option>`).join('')}</select></div></div>
    <div class="fr"><label>Oggetto</label><input id="f-ogg"></div>
    <div class="fg3"><div class="fr"><label>Importo (€)</label><input id="f-imp" type="number"></div><div class="fr"><label>Apertura</label><input id="f-dat" type="date" value="${today}"></div><div class="fr"><label>Consegna prev.</label><input id="f-cons" type="date"></div></div>
    <div class="fr"><label>Stato</label><select id="f-stato">${SC.map(s=>`<option value="${s}">${SLC[s]}</option>`).join('')}</select></div>
    <div class="fr"><label>Note</label><textarea id="f-note"></textarea></div>
  </div>${act('saveComm()')}`;

  if(type==='edit-commessa'){const c=CC||{};return`<h3>Modifica commessa</h3><div class="form-inline">
    <div class="fg2"><div class="fr"><label>Numero</label><input id="f-num" value="${c.numero||''}" readonly style="opacity:.6;cursor:not-allowed"></div><div class="fr"><label>Tipo commessa</label><select id="f-tipo"><option value="">— seleziona —</option>${Object.entries(TIPI).map(([k,v])=>`<option value="${k}"${c.tipo_commessa===k?' selected':''}>${v}</option>`).join('')}</select></div></div>
    <div class="fr"><label>Oggetto</label><input id="f-ogg" value="${c.oggetto||''}"></div>
    <div class="fg3"><div class="fr"><label>Importo (€)</label><input id="f-imp" type="number" value="${c.importo||''}"></div><div class="fr"><label>Apertura</label><input id="f-dat" type="date" value="${c.data_apertura||''}"></div><div class="fr"><label>Consegna prev.</label><input id="f-cons" type="date" value="${c.data_consegna||''}"></div></div>
    <div class="fr"><label>Stato</label><select id="f-stato">${SC.map(s=>`<option value="${s}"${c.stato===s?' selected':''}>${SLC[s]}</option>`).join('')}</select></div>
    <div class="fr"><label>Note</label><textarea id="f-note">${c.note||''}</textarea></div>
  </div>${act('updateComm()')}`;
  }

  if(type==='add-riga'){
  const magItems=lMag();
  const magOpts=magItems.map(m=>`<option value="${m.id}">[MAG] ${m.codice||''} — ${m.descrizione} (${m.qty||0} pz)`).join('');
  return`<h3 style="margin-bottom:12px">+ Aggiungi componente</h3>

  <!-- Selettore modalità — grid di bottoni chiari -->
  <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(120px,1fr));gap:6px;margin-bottom:14px;">
    <button id="atab-listino" onclick="switchAddTab('listino')" class="add-mode-btn active" title="Cerca nel catalogo Ares/Tecfluid">
      <span style="font-size:18px">📋</span>
      <span>Listino Ares</span>
    </button>
    <button id="atab-mag" onclick="switchAddTab('mag')" class="add-mode-btn" title="Componenti già in magazzino">
      <span style="font-size:18px">📦</span>
      <span>Magazzino</span>
    </button>
    <button id="atab-tubo" onclick="switchAddTab('tubo')" class="add-mode-btn" title="Configuratore tubo flessibile raccordato">
      <span style="font-size:18px">🔧</span>
      <span>Tubo raccordato</span>
    </button>
    <button id="atab-racc" onclick="switchAddTab('racc')" class="add-mode-btn" title="Raccorderia scelta libera">
      <span style="font-size:18px">🔩</span>
      <span>Raccorderia</span>
    </button>
    <button id="atab-pu10" onclick="switchAddTab('pu10')" class="add-mode-btn" title="Configuratore centralina PU10">
      <span style="font-size:18px">⚙</span>
      <span>Centralina PU10</span>
    </button>
    <button id="atab-pu20" onclick="switchAddTab('pu20')" class="add-mode-btn" title="Configuratore centralina PU20 (pompe Gr.2)">
      <span style="font-size:18px">⚙</span>
      <span>Centralina PU20</span>
    </button>
    <button id="atab-manuale" onclick="switchAddTab('manuale')" class="add-mode-btn" title="Inserimento manuale componente">
      <span style="font-size:18px">✏️</span>
      <span>Manuale</span>
    </button>
  </div>

  <!-- Contenuto pannello attivo -->
  <div id="apanel-listino" style="border-top:1px solid var(--border);padding-top:12px;margin-top:2px;">
    <div style="display:flex;gap:8px;margin-bottom:8px;">
      <input id="ares-q" style="flex:1;border:1.5px solid var(--border);border-radius:7px;padding:7px 10px;font-size:14px;outline:none;" placeholder="Cerca sigla o descrizione… (es. B09, NCE, collettore)" oninput="aresFilter()">
      <select id="ares-serie" class="keep-native" style="border:1.5px solid var(--border);border-radius:7px;padding:7px 9px;font-size:13px;outline:none;" onchange="aresFilter()">
        <option value="">Tutte le serie</option>
        <option value="PU05">PU05</option>
        <option value="PU10">PU10</option>
        <option value="PU20">PU20</option>
      </select>
    </div>
    <div id="ares-results" style="max-height:260px;overflow-y:auto;font-size:13px;"></div>
  </div>

  <!-- Panel: Magazzino -->
  <div id="apanel-mag" style="display:none;border-top:1px solid var(--border);padding-top:12px;margin-top:2px;">
    ${magItems.length?
    '<input id="mag-q" style="width:100%;border:1.5px solid var(--border);border-radius:7px;padding:7px 10px;font-size:14px;outline:none;margin-bottom:8px;" placeholder="\ud83d\udd0d cerca nel magazzino\u2026" oninput="magFilter()">'+
    '<div id="mag-results" style="max-height:280px;overflow-y:auto;">'+
      magItems.map(m=>'<div style="display:flex;align-items:center;gap:10px;padding:8px 4px;border-bottom:1px solid var(--border);font-size:13px;">'+
        '<div style="flex:1;">'+
          '<div style="font-weight:600">'+(m.codice?'<span style="font-family:monospace;color:var(--green);font-weight:700">'+m.codice+'</span> \u2014 ':'')+' '+(m.descrizione||'\u2014')+'</div>'+
          '<div style="font-size:11px;color:var(--text3)">'+(m.fornitore||'')+' \u00b7 giacenza: <strong style="color:'+(parseFloat(m.qty||0)>0?'var(--green)':'var(--red)')+'">'+(m.qty||0)+'</strong></div>'+
        '</div>'+
        '<span style="font-weight:700;white-space:nowrap">\u20ac'+fmt2(m.costo_cad||0)+'</span>'+
        '<button style="background:var(--blue);border:none;color:#fff;font-size:11px;font-weight:700;padding:4px 10px;border-radius:20px;cursor:pointer;" onclick="usaMagazzino('+JSON.stringify(m).replace(/"/g,'&quot;')+',\''+extra+'\')">&plus; usa</button>'+
      '</div>').join('')+
    '</div>':'<div class="empty" style="padding:20px">Magazzino vuoto \u2014 aggiungi componenti nella sezione \ud83d\udce6 Magazzino</div>'}
  </div>

  <!-- Panel: Configuratore tubo flessibile -->
  <div id="apanel-tubo" style="display:none;border-top:1px solid var(--border);padding-top:12px;margin-top:2px;">
    <div style="background:#fff7ed;border:1px solid #fed7aa;border-radius:8px;padding:10px 12px;margin-bottom:12px;font-size:12px;color:#7c2d12;">
      Configuratore tubo flessibile raccordato — Pa.Co Fluid. I raccordi si filtrano automaticamente in base al diametro del tubo scelto.
    </div>

    <!-- Riga 1: Tipo tubo + Lunghezza -->
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:10px;">
      <div>
        <label style="font-size:12px;font-weight:700;color:var(--text3);display:block;margin-bottom:4px">Tipo tubo</label>
        <select id="tf-tipo" class="keep-native" style="width:100%;border:1.5px solid var(--border);border-radius:7px;padding:7px;font-size:12px;outline:none;" onchange="onTuboTipoChange()">
          <option value="">— seleziona —</option>
          <optgroup label="Trecciati — 1 treccia (R1AT/1SN)">
            <option value="1SN|R1AT|100|6|1/4&quot;">DN06 (¼&quot;) · R1AT/1SN · max 100 bar</option>
            <option value="1SN|R1AT|250|10|3/8&quot;">DN10 (⅜&quot;) · R1AT/1SN · max 250 bar</option>
            <option value="1SN|R1AT|200|12|1/2&quot;">DN12 (½&quot;) · R1AT/1SN · max 200 bar</option>
            <option value="1SN|R1AT|160|16|5/8&quot;">DN16 (⅝&quot;) · R1AT/1SN · max 160 bar</option>
            <option value="1SN|R1AT|125|20|3/4&quot;">DN20 (¾&quot;) · R1AT/1SN · max 125 bar</option>
            <option value="1SN|R1AT|100|25|1&quot;">DN25 (1&quot;) · R1AT/1SN · max 100 bar</option>
          </optgroup>
          <optgroup label="Trecciati — 2 trecce (R2AT/2SN)">
            <option value="2SN|R2AT|180|6|1/4&quot;">DN06 (¼&quot;) · R2AT/2SN · max 180 bar</option>
            <option value="2SN|R2AT|330|10|3/8&quot;">DN10 (⅜&quot;) · R2AT/2SN · max 330 bar</option>
            <option value="2SN|R2AT|215|12|1/2&quot;">DN12 (½&quot;) · R2AT/2SN · max 215 bar</option>
            <option value="2SN|R2AT|165|16|5/8&quot;">DN16 (⅝&quot;) · R2AT/2SN · max 165 bar</option>
            <option value="2SN|R2AT|140|20|3/4&quot;">DN20 (¾&quot;) · R2AT/2SN · max 140 bar</option>
            <option value="2SN|R2AT|105|25|1&quot;">DN25 (1&quot;) · R2AT/2SN · max 105 bar</option>
          </optgroup>
          <optgroup label="Spiralati 4 strati (R9R/4SP)">
            <option value="4SP|R9R|350|6|1/4&quot;">DN06 (¼&quot;) · 4SP/R9R · max 350 bar</option>
            <option value="4SP|R9R|420|10|3/8&quot;">DN10 (⅜&quot;) · 4SP/R9R · max 420 bar</option>
            <option value="4SP|R9R|350|12|1/2&quot;">DN12 (½&quot;) · 4SP/R9R · max 350 bar</option>
            <option value="4SP|R9R|320|16|5/8&quot;">DN16 (⅝&quot;) · 4SP/R9R · max 320 bar</option>
            <option value="4SP|R9R|280|20|3/4&quot;">DN20 (¾&quot;) · 4SP/R9R · max 280 bar</option>
            <option value="4SP|R9R|210|25|1&quot;">DN25 (1&quot;) · 4SP/R9R · max 210 bar</option>
          </optgroup>
          <optgroup label="Spiralati 6 strati (R12/6SP) — alta pressione">
            <option value="6SP|R12|420|12|1/2&quot;">DN12 (½&quot;) · 6SP/R12 · max 420 bar</option>
            <option value="6SP|R12|400|19|3/4&quot;">DN19 (¾&quot;) · 6SP/R12 · max 400 bar</option>
            <option value="6SP|R12|350|25|1&quot;">DN25 (1&quot;) · 6SP/R12 · max 350 bar</option>
          </optgroup>
          <optgroup label="Termoplastici / speciali">
            <option value="TF|R7|210|10|3/8&quot;">DN10 (⅜&quot;) · Termoplastico R7 · max 210 bar</option>
            <option value="TF|R7|210|12|1/2&quot;">DN12 (½&quot;) · Termoplastico R7 · max 210 bar</option>
            <option value="R4|AS|10|25|1&quot;">DN25 (1&quot;) · Aspirazione R4 · max 10 bar</option>
            <option value="R6|LP|16|25|1&quot;">DN25 (1&quot;) · Ritorno/drenaggio R6 · max 16 bar</option>
          </optgroup>
        </select>
      </div>
      <div>
        <label style="font-size:12px;font-weight:700;color:var(--text3);display:block;margin-bottom:4px">Lunghezza</label>
        <div style="display:flex;gap:6px;align-items:center;">
          <input id="tf-lung-mm" type="number" min="50" max="10000" step="50" placeholder="mm" style="flex:1;border:1.5px solid var(--border);border-radius:7px;padding:7px;font-size:14px;outline:none;" oninput="syncLung('mm')">
          <span style="font-size:12px;color:var(--text3);flex-shrink:0">mm</span>
          <input id="tf-lung-in" type="number" min="2" max="400" step="0.5" placeholder="inch" style="width:68px;border:1.5px solid var(--border);border-radius:7px;padding:7px;font-size:13px;outline:none;" oninput="syncLung('in')">
          <span style="font-size:12px;color:var(--text3);flex-shrink:0">″</span>
        </div>
        <div id="tf-lung-hint" style="font-size:10px;color:var(--text4);margin-top:2px;"></div>
      </div>
    </div>

    <!-- Riga 2: Raccordi con selettore visivo filtrato -->
    <div id="tf-racc-section">
      <div style="font-size:12px;color:var(--text3);padding:10px;text-align:center;background:var(--bg2);border-radius:7px;">
        Seleziona prima il tipo di tubo per filtrare i raccordi compatibili
      </div>
    </div>

    <!-- Riga 3: Qtà + Orientamento doppia curva -->
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:12px;margin-top:10px;">
      <div>
        <label style="font-size:12px;font-weight:700;color:var(--text3);display:block;margin-bottom:4px">Q.tà</label>
        <input id="tf-qty" type="number" value="1" min="1" style="width:100%;border:1.5px solid var(--border);border-radius:7px;padding:7px;font-size:14px;outline:none;" oninput="aggTuboPreview()">
      </div>
      <div id="tf-orient-row" style="display:none;">
        <label style="font-size:12px;font-weight:700;color:var(--amber);display:block;margin-bottom:4px">⟳ Orientamento doppia curva</label>
        <select id="tf-orient" style="width:100%;border:1.5px solid #f59e0b;border-radius:7px;padding:7px;font-size:13px;outline:none;" onchange="aggTuboPreview()">
          <option value="0">0° — curve parallele, stesso piano</option>
          <option value="45">45°</option>
          <option value="90">90° — curve perpendicolari</option>
          <option value="135">135°</option>
          <option value="180">180° — curve opposte, stesso piano</option>
          <option value="225">225°</option>
          <option value="270">270°</option>
          <option value="315">315°</option>
        </select>
        <div style="font-size:10px;color:var(--text3);margin-top:3px">Angolo di rotazione tra i piani delle due curve (convenzione Pa.Co)</div>
      </div>
    </div>

    <div id="tf-preview" style="background:var(--bg2);border:1px solid var(--border);border-radius:8px;padding:10px 14px;font-size:13px;color:var(--text2);min-height:56px;">
      <span style="color:var(--text4)">Compila i campi sopra per vedere la descrizione generata…</span>
    </div>
    <div style="margin-top:10px;display:flex;gap:8px;justify-content:flex-end;">
      <button class="btn-ghost" onclick="closeM()">Annulla</button>
      <button class="btn" onclick="usaTuboFlessibile('${extra}')">+ aggiungi alla distinta</button>
    </div>
  </div>

  <!-- Panel: Raccorderia sciolta -->
  <div id="apanel-racc" style="display:none;border-top:1px solid var(--border);padding-top:12px;margin-top:2px;">
    <div style="display:flex;gap:6px;flex-wrap:wrap;margin-bottom:10px;align-items:center;">
      <span style="font-size:11px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.06em;">Standard:</span>
      ${['Tutti','BSP','DIN','JIC','ORFS','Bonded','Stucchi'].map((s,i)=>`<button onclick="setRaccScioltoFiltro('${s==='Tutti'?'':s}',this)" style="font-size:11px;font-weight:700;padding:3px 10px;border-radius:20px;border:1.5px solid ${i===0?'var(--blue)':'var(--border)'};background:${i===0?'var(--blue-lt)':'#fff'};color:${i===0?'var(--blue)':'var(--text3)'};cursor:pointer;">${s}</button>`).join('')}
      <input id="racc-q" placeholder="🔍 cerca…" style="flex:1;min-width:100px;border:1.5px solid var(--border);border-radius:7px;padding:5px 9px;font-size:13px;outline:none;" oninput="renderRaccSciolto()">
    </div>
    <div id="racc-results" style="max-height:310px;overflow-y:auto;border:1.5px solid var(--border);border-radius:8px;background:#fff;"></div>
    <div style="font-size:11px;color:var(--text3);margin-top:6px;">Pa.Co Fluid per BSP/DIN/JIC/ORFS · Tecnical-FIZ per Stucchi · Cliccando "+ usa" si aggiunge alla distinta con fornitore preimpostato</div>
  </div>

  <!-- Panel: Configuratore PU10 -->
  <div id="apanel-pu10" style="display:none;border-top:1px solid var(--border);padding-top:12px;margin-top:2px;">
    <div style="background:#f0fdf4;border:1px solid #6ee7a0;border-radius:8px;padding:10px 13px;margin-bottom:12px;font-size:12px;color:#166534;">
      Configuratore centralina PU10 (Ares / Tecfluid) — genera codice ordinazione e distinta completa.
    </div>
    <!-- Step 1: Collettore -->
    <div style="margin-bottom:10px;">
      <div style="font-size:11px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.06em;margin-bottom:5px;">1 — Collettore</div>
      <select id="pu10-coll" class="keep-native" style="width:100%;border:1.5px solid var(--border);border-radius:7px;padding:7px;font-size:13px;outline:none;" onchange="pu10Update()">
        <option value="">— seleziona —</option>
        <option value="A1A|263.8|Collettore monopompa VMC1 P-T G1/4&quot;">A1A — monopompa VMC1, att. G1/4"</option>
        <option value="A1B|278.5|Collettore + valvola NC pilotata integrata">A1B — + valvola NC integrata</option>
        <option value="A1C|266.4|Collettore con ritegno su A">A1C — con ritegno su A</option>
        <option value="A1D|269.7|Collettore doppio effetto A-T G1/4&quot;">A1D — doppio effetto A-T G1/4"</option>
        <option value="B2A|278.5|Collettore semplice effetto P+2 utenze">B2A — semplice effetto, P+2 utenze</option>
        <option value="B3B|278.5|Collettore doppio effetto A-B">B3B — doppio effetto A-B</option>
        <option value="B4A|286.0|Collettore 4 vie A-B con check">B4A — 4 vie A-B con check</option>
        <option value="R1A|537.2|Collettore pompa reversibile P1-P2 G3/8&quot;">R1A — pompa reversibile, att. G3/8"</option>
      </select>
    </div>
    <!-- Step 2: Pompa -->
    <div style="margin-bottom:10px;">
      <div style="font-size:11px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.06em;margin-bottom:5px;">2 — Pompa</div>
      <select id="pu10-pompa" class="keep-native" style="width:100%;border:1.5px solid var(--border);border-radius:7px;padding:7px;font-size:13px;outline:none;" onchange="pu10Update()">
        <option value="">— seleziona —</option>
        <option value="PG04|88.0|Pompa Gr.1 2.00 cc/giro 250 bar 2.82 l/min @1500">PG04 — 2.0 cc/giro · 250 bar · 2.82 l/min @1500</option>
        <option value="PG05|92.0|Pompa Gr.1 2.50 cc/giro 250 bar 3.53 l/min @1500">PG05 — 2.5 cc/giro · 250 bar · 3.53 l/min @1500</option>
        <option value="PG07|98.0|Pompa Gr.1 3.15 cc/giro 250 bar 4.44 l/min @1500">PG07 — 3.15 cc/giro · 250 bar · 4.44 l/min @1500</option>
        <option value="PG09|104.0|Pompa Gr.1 4.20 cc/giro 250 bar 5.92 l/min @1500">PG09 — 4.2 cc/giro · 250 bar · 5.92 l/min @1500</option>
        <option value="PG11|110.0|Pompa Gr.1 5.00 cc/giro 250 bar 7.05 l/min @1500">PG11 — 5.0 cc/giro · 250 bar · 7.05 l/min @1500</option>
        <option value="PG12|115.0|Pompa Gr.1 5.70 cc/giro 200 bar 8.12 l/min @1500">PG12 — 5.7 cc/giro · 200 bar · 8.12 l/min @1500</option>
      </select>
    </div>
    <!-- Step 3: Motore -->
    <div style="margin-bottom:10px;">
      <div style="font-size:11px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.06em;margin-bottom:5px;">3 — Motore</div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;">
        <select id="pu10-mot-tipo" class="keep-native" style="width:100%;border:1.5px solid var(--border);border-radius:7px;padding:7px;font-size:12px;outline:none;" onchange="pu10Update()">
          <option value="">— tipo —</option>
          <optgroup label="DC (corrente continua)">
            <option value="C104|148.0|DC 12V 1600W IP54">C104 — 12V DC 1600W IP54</option>
            <option value="C105|154.0|DC 12V 1600W IP54 con termica">C105 — 12V DC 1600W IP54 + termica</option>
            <option value="C204|162.0|DC 24V 2200W IP54">C204 — 24V DC 2200W IP54</option>
            <option value="C205|168.0|DC 24V 2200W IP54 con termica">C205 — 24V DC 2200W IP54 + termica</option>
            <option value="C211|158.0|DC 24V 2000W IP54 con termica">C211 — 24V DC 2000W IP54 + termica</option>
          </optgroup>
          <optgroup label="AC trifase">
            <option value="T206|145.0|AC 3~ 0.75kW 2900rpm MEC80">T206 — AC 3~ 0.75 kW 2900 rpm MEC80</option>
            <option value="T207|158.0|AC 3~ 1.1kW 2900rpm MEC80">T207 — AC 3~ 1.1 kW 2900 rpm MEC80</option>
            <option value="T208|175.0|AC 3~ 1.5kW 2900rpm MEC90">T208 — AC 3~ 1.5 kW 2900 rpm MEC90</option>
          </optgroup>
          <optgroup label="AC monofase">
            <option value="M206|155.0|AC 1~ 0.75kW 2900rpm MEC80">M206 — AC 1~ 0.75 kW 2900 rpm MEC80</option>
            <option value="M207|168.0|AC 1~ 1.1kW 2900rpm MEC80">M207 — AC 1~ 1.1 kW 2900 rpm MEC80</option>
          </optgroup>
        </select>
        <select id="pu10-giunto" class="keep-native" style="width:100%;border:1.5px solid var(--border);border-radius:7px;padding:7px;font-size:12px;outline:none;" onchange="pu10Update()">
          <option value="">— giunto/flangia —</option>
          <option value="FC02|28.5|Giunto per DC C104-C205">FC02 — per DC C104/C204/C205</option>
          <option value="FC01|22.4|Giunto per DC C202">FC01 — per DC C102/C103/C202</option>
          <option value="FA80|35.2|Flangia AC MEC80">FA80 — flangia AC MEC80</option>
          <option value="FA90|38.8|Flangia AC MEC90">FA90 — flangia AC MEC90</option>
        </select>
      </div>
    </div>
    <!-- Step 4: Serbatoio -->
    <div style="margin-bottom:10px;">
      <div style="font-size:11px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.06em;margin-bottom:5px;">4 — Serbatoio</div>
      <select id="pu10-serb" class="keep-native" style="width:100%;border:1.5px solid var(--border);border-radius:7px;padding:7px;font-size:13px;outline:none;" onchange="pu10Update()">
        <option value="">— seleziona —</option>
        <option value="SL44H|31.2|Serbatoio lamiera 4L orizzontale Ø123">SL44H — 4 L lamiera orizzontale</option>
        <option value="SL47H|38.5|Serbatoio lamiera 7L orizzontale Ø123">SL47H — 7 L lamiera orizzontale</option>
        <option value="SL48H|42.0|Serbatoio lamiera 8L orizzontale Ø123">SL48H — 8 L lamiera orizzontale</option>
        <option value="SP07H|22.4|Serbatoio plastica 3L orizzontale Ø123">SP07H — 3 L plastica orizzontale</option>
        <option value="SP08H|26.8|Serbatoio plastica 5L orizzontale Ø123">SP08H — 5 L plastica orizzontale</option>
      </select>
    </div>
    <!-- Step 5: Blocchi modulari (opzionali) -->
    <div style="margin-bottom:10px;">
      <div style="font-size:11px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.06em;margin-bottom:5px;">5 — Blocchi modulari <span style="font-weight:400;text-transform:none;color:var(--text4)">(opzionali, multipli)</span></div>
      <div id="pu10-blocchi-list" style="display:flex;flex-direction:column;gap:5px;margin-bottom:6px;max-height:120px;overflow-y:auto;"></div>
      <select id="pu10-blocco-add" class="keep-native" style="width:100%;border:1.5px solid var(--border);border-radius:7px;padding:6px;font-size:12px;outline:none;color:var(--text3);" onchange="pu10AddBlocco(this)">
        <option value="">+ aggiungi blocco modulare…</option>
        <option value="B09|19.1|Distanziale H=18 pass-through P-T">B09 — Distanziale H=18</option>
        <option value="B01|25.6|Distanziale H=39">B01 — Distanziale H=39</option>
        <option value="B92|23.6|Distanziale H=25">B92 — Distanziale H=25</option>
        <option value="B15|102.5|Adattatore rotazione lato motore H=90 M 1/4&quot; BSPP">B15 — Adattatore rotaz. motore H=90</option>
        <option value="B03|39.2|Blocco CETOP3-NG6 parallelo/serie 1/4&quot;">B03 — CETOP3-NG6 parallelo/serie 1/4"</option>
        <option value="B11|39.2|Blocco CETOP3-NG6 parallelo/serie 3/8&quot;">B11 — CETOP3-NG6 parallelo/serie 3/8"</option>
        <option value="B85|45.0|Blocco CETOP3-NG6 parallelo con uscite 1/4&quot;">B85 — CETOP3-NG6 con uscite 1/4"</option>
        <option value="B73|297.0|Blocco riduttrice di pressione 350 bar">B73 — Riduttrice di pressione</option>
        <option value="B30_06|251.1|Divisore flusso 50/50 1-6 l/min">B30 — Divisore flusso 50/50</option>
        <option value="B259A|134.8|Blocco scarico automatico tipo A">B259A — Scarico automatico tipo A</option>
        <option value="B38|78.5|Blocco filtro scarico 15 micron">B38 — Filtro sullo scarico 15μ</option>
      </select>
    </div>
    <!-- Preview codice e prezzo -->
    <div id="pu10-preview" style="background:#f0fdf4;border:1px solid #6ee7a0;border-radius:8px;padding:10px 14px;min-height:60px;font-size:12px;color:var(--text2);">
      <span style="color:var(--text4)">Seleziona almeno collettore, pompa e motore…</span>
    </div>
    <div style="margin-top:10px;display:flex;gap:8px;justify-content:flex-end;">
      <button class="btn-ghost" onclick="closeM()">Annulla</button>
      <button class="btn" style="background:#166534;" onclick="usaCentralinaPU10('${extra}')">⚙ Aggiungi centralina alla distinta</button>
    </div>
  </div>


    <!-- Panel: Configuratore PU20 -->
  <div id="apanel-pu20" style="display:none;border-top:1px solid var(--border);padding-top:12px;margin-top:2px;">
    <div style="background:#f0fdf4;border:1px solid #6ee7a0;border-radius:8px;padding:10px 13px;margin-bottom:12px;font-size:12px;color:#166534;">
      Configuratore centralina PU20 (Ares / Tecfluid) — serie più grande, pompe gruppo 2, serbatoio verticale da 8÷20 L.
    </div>
    <!-- Step 1: Collettore -->
    <div style="margin-bottom:10px;">
      <div style="font-size:11px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.06em;margin-bottom:5px;">1 — Collettore</div>
      <select id="pu20-coll" class="keep-native" style="width:100%;border:1.5px solid var(--border);border-radius:7px;padding:7px;font-size:13px;outline:none;" onchange="pu20Update()">
        <option value="">— seleziona —</option>
        <option value="Z1A|574.6|Collettore monopompa VMZ1 P G3/8&quot; T G1/2&quot;">Z1A — monopompa VMZ1, P G3/8" T G1/2"</option>
        <option value="Z2A|988.6|Collettore doppia pompa 2×VMZ1">Z2A — doppia pompa 2×VMZ1</option>
      </select>
    </div>
    <!-- Step 2: Pompa -->
    <div style="margin-bottom:10px;">
      <div style="font-size:11px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.06em;margin-bottom:5px;">2 — Pompa (gruppo 2)</div>
      <select id="pu20-pompa" class="keep-native" style="width:100%;border:1.5px solid var(--border);border-radius:7px;padding:7px;font-size:13px;outline:none;" onchange="pu20Update()">
        <option value="">— seleziona —</option>
        <option value="PG20|165.0|Pompa Gr.2 8.0 cc/giro 200 bar 12.0 l/min @1500">PG20 — 8.0 cc/giro · 200 bar · 12.0 l/min @1500</option>
        <option value="PG21|172.0|Pompa Gr.2 10.0 cc/giro 200 bar 15.0 l/min @1500">PG21 — 10.0 cc/giro · 200 bar · 15.0 l/min @1500</option>
        <option value="PG22|180.0|Pompa Gr.2 12.0 cc/giro 200 bar 18.0 l/min @1500">PG22 — 12.0 cc/giro · 200 bar · 18.0 l/min @1500</option>
        <option value="PG23|188.0|Pompa Gr.2 14.0 cc/giro 200 bar 21.0 l/min @1500">PG23 — 14.0 cc/giro · 200 bar · 21.0 l/min @1500</option>
        <option value="PG24|196.0|Pompa Gr.2 16.0 cc/giro 200 bar 24.0 l/min @1500">PG24 — 16.0 cc/giro · 200 bar · 24.0 l/min @1500</option>
        <option value="PG25|210.0|Pompa Gr.2 19.0 cc/giro 200 bar 28.5 l/min @1500">PG25 — 19.0 cc/giro · 200 bar · 28.5 l/min @1500</option>
      </select>
    </div>
    <!-- Step 3: Motore -->
    <div style="margin-bottom:10px;">
      <div style="font-size:11px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.06em;margin-bottom:5px;">3 — Motore</div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;">
        <select id="pu20-mot-tipo" class="keep-native" style="width:100%;border:1.5px solid var(--border);border-radius:7px;padding:7px;font-size:12px;outline:none;" onchange="pu20Update()">
          <option value="">— tipo —</option>
          <optgroup label="DC (corrente continua)">
            <option value="C204|162.0|DC 24V 2200W IP54">C204 — 24V DC 2200W IP54</option>
            <option value="C205|168.0|DC 24V 2200W IP54 con termica">C205 — 24V DC 2200W + termica</option>
            <option value="C304|175.0|DC 24V 3000W IP54">C304 — 24V DC 3000W IP54</option>
          </optgroup>
          <optgroup label="AC trifase">
            <option value="T208|175.0|AC 3~ 1.5kW 2900rpm MEC90">T208 — AC 3~ 1.5 kW MEC90</option>
            <option value="T209|192.0|AC 3~ 2.2kW 2900rpm MEC90">T209 — AC 3~ 2.2 kW MEC90</option>
            <option value="T210|215.0|AC 3~ 3.0kW 2900rpm MEC100">T210 — AC 3~ 3.0 kW MEC100</option>
            <option value="T211|240.0|AC 3~ 4.0kW 2900rpm MEC112">T211 — AC 3~ 4.0 kW MEC112</option>
          </optgroup>
          <optgroup label="AC monofase">
            <option value="M208|188.0|AC 1~ 1.5kW 2900rpm MEC90">M208 — AC 1~ 1.5 kW MEC90</option>
            <option value="M209|205.0|AC 1~ 2.2kW 2900rpm MEC90">M209 — AC 1~ 2.2 kW MEC90</option>
          </optgroup>
        </select>
        <select id="pu20-giunto" class="keep-native" style="width:100%;border:1.5px solid var(--border);border-radius:7px;padding:7px;font-size:12px;outline:none;" onchange="pu20Update()">
          <option value="">— giunto/flangia —</option>
          <option value="FC03|32.5|Giunto elastico per DC 24V Gr.2">FC03 — elastico DC 24V Gr.2</option>
          <option value="FA90|38.8|Flangia AC MEC90">FA90 — flangia AC MEC90</option>
          <option value="FA100|42.5|Flangia AC MEC100">FA100 — flangia AC MEC100</option>
          <option value="FA112|46.0|Flangia AC MEC112">FA112 — flangia AC MEC112</option>
        </select>
      </div>
    </div>
    <!-- Step 4: Serbatoio -->
    <div style="margin-bottom:10px;">
      <div style="font-size:11px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.06em;margin-bottom:5px;">4 — Serbatoio</div>
      <select id="pu20-serb" class="keep-native" style="width:100%;border:1.5px solid var(--border);border-radius:7px;padding:7px;font-size:13px;outline:none;" onchange="pu20Update()">
        <option value="">— seleziona —</option>
        <option value="SL28V|52.0|Serbatoio lamiera 8L verticale Ø165">SL28V — 8 L lamiera verticale Ø165</option>
        <option value="SL29V|58.0|Serbatoio lamiera 12L verticale Ø165">SL29V — 12 L lamiera verticale Ø165</option>
        <option value="SL30V|65.0|Serbatoio lamiera 16L verticale Ø165">SL30V — 16 L lamiera verticale Ø165</option>
        <option value="SL31V|72.0|Serbatoio lamiera 20L verticale Ø165">SL31V — 20 L lamiera verticale Ø165</option>
        <option value="SP20V|38.0|Serbatoio plastica 8L verticale Ø165">SP20V — 8 L plastica verticale Ø165</option>
        <option value="SP21V|44.0|Serbatoio plastica 12L verticale Ø165">SP21V — 12 L plastica verticale Ø165</option>
      </select>
    </div>
    <!-- Step 5: Accessori -->
    <div style="margin-bottom:10px;">
      <div style="font-size:11px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.06em;margin-bottom:5px;">5 — Accessori <span style="font-weight:400;text-transform:none;color:var(--text4)">(opzionali)</span></div>
      <div id="pu20-blocchi-list" style="display:flex;flex-direction:column;gap:5px;margin-bottom:6px;max-height:120px;overflow-y:auto;"></div>
      <select id="pu20-blocco-add" class="keep-native" style="width:100%;border:1.5px solid var(--border);border-radius:7px;padding:6px;font-size:12px;outline:none;color:var(--text3);" onchange="pu20AddBlocco(this)">
        <option value="">+ aggiungi accessorio…</option>
        <option value="G00|0.0|Senza supporto montaggio">G00 — senza supporto</option>
        <option value="G01|18.5|Supporto montaggio standard L=150 H=87">G01 — supporto standard L=150</option>
        <option value="G02|24.0|Supporto montaggio rinforzato L=200">G02 — supporto rinforzato L=200</option>
        <option value="MG01|12.8|Manometro glicerina 0-160 bar G1/4&quot;">MG01 — manometro 0-160 bar</option>
        <option value="MG02|14.5|Manometro glicerina 0-250 bar G1/4&quot;">MG02 — manometro 0-250 bar</option>
        <option value="FL01|28.0|Filtro aspirazione 100 micron">FL01 — filtro aspirazione 100μ</option>
        <option value="FL02|34.0|Filtro scarico 25 micron">FL02 — filtro scarico 25μ</option>
        <option value="PS01|42.0|Pressostato 1-10 bar G1/4&quot;">PS01 — pressostato 1-10 bar</option>
        <option value="TS01|38.0|Termostato bimetallico 80°C">TS01 — termostato 80°C</option>
        <option value="RV01|55.0|Valvola di massima regolabile 50-350 bar">RV01 — val. massima 50-350 bar</option>
      </select>
    </div>
    <!-- Preview -->
    <div id="pu20-preview" style="background:#f0fdf4;border:1px solid #6ee7a0;border-radius:8px;padding:10px 14px;min-height:60px;font-size:12px;color:var(--text2);">
      <span style="color:var(--text4)">Seleziona almeno collettore, pompa e motore…</span>
    </div>
    <div style="margin-top:10px;display:flex;gap:8px;justify-content:flex-end;">
      <button class="btn-ghost" onclick="closeM()">Annulla</button>
      <button class="btn" style="background:#166534;" onclick="usaCentralinaPU20(currentAddRigaNumero)">⚙ Aggiungi PU20 alla distinta</button>
    </div>
  </div>

<!-- Panel: Manuale -->
  <div id="apanel-manuale" style="display:none;border-top:1px solid var(--border);padding-top:12px;margin-top:2px;">
    <div class="form-inline">
      <div class="fr"><label>Sezione</label><input id="f-sez" placeholder="es. Centralina, Raccorderia" list="sez-opt"><datalist id="sez-opt"><option value="Centralina"><option value="Cilindri"><option value="Raccorderia"><option value="Tubi flessibili"><option value="Elettrica"><option value="Generale"></datalist></div>
      <div class="fr"><label>Descrizione</label><textarea id="f-desc" style="min-height:50px"></textarea></div>
      <div class="fg3"><div class="fr"><label>Q.tà</label><input id="f-qty" type="number" value="1" min="1"></div><div class="fr"><label>Fornitore</label><input id="f-forn" placeholder="es. Pa.Co, GTA…" list="forn-list"><datalist id="forn-list">${FORNITORI.map(f=>`<option value="${f.nome}">`).join('')}</datalist></div><div class="fr"><label>Codice</label><input id="f-cod"></div></div>
      <div class="fg3"><div class="fr"><label>Costo/cad (€)</label><input id="f-costo" type="number" step="0.01"></div><div class="fr"><label>Sconto (%)</label><input id="f-sconto" type="number" min="0" max="100" placeholder="0"></div><div class="fr"><label>Mark-up</label><input id="f-markup" type="number" step="0.01" value=""></div></div>
      <div class="fg2"><div class="fr"><label>Stato approvv.</label><select id="f-stato-a">${Object.entries(SAP).map(([k,v])=>`<option value="${k}">${v}</option>`).join('')}</select></div><div class="fr"><label>Data attesa</label><input id="f-datta" type="date"></div></div>
    </div>
    ${act("saveRiga('"+extra+"')")}
  </div>
  `;
  }
  if(type==='add-sezione')return`<h3>Nuova sezione distinta</h3>
  <p style="font-size:13px;color:var(--text3);margin-bottom:14px;line-height:1.5">Le sezioni raggruppano i componenti nella distinta (es. <em>Centralina</em>, <em>Cilindri</em>, <em>Raccorderia</em>).</p>
  <div class="form-row"><label>Nome sezione</label><input id="f-sez-nome" placeholder="es. Centralina PU10, Cilindri, Elettrica…" list="sez-sugg"><datalist id="sez-sugg"><option value="Centralina PU10"><option value="Centralina PU20"><option value="Centralina PU05"><option value="Cilindri"><option value="Raccorderia"><option value="Tubi flessibili"><option value="Valvole CETOP"><option value="Elettrica"><option value="Sensori"><option value="Struttura"><option value="Varie"></datalist></div>
  <div class="modal-actions"><button class="btn-ghost" onclick="closeM()">Annulla</button><button class="btn" onclick="saveSezione('${extra}')">Crea sezione</button></div>`;

  if(type==='edit-riga'){
    const parts=(extra||'').split('|');
    const numero=parts[0]||'';const rid=parts[1]||'';
    const det=ll('det');const cd=det[numero]||{distinta:[]};
    const r=(cd.distinta||[]).find(x=>x.id===rid)||{};
    return`<h3>✎ Modifica componente</h3>
    <div class="form-inline">
      <div class="fr"><label>Sezione</label><input id="f-sez" value="${r.sezione||''}" placeholder="es. Centralina, Raccorderia" list="sez-opt2"><datalist id="sez-opt2"><option value="Centralina"><option value="Cilindri"><option value="Raccorderia"><option value="Tubi flessibili"><option value="Elettrica"><option value="Generale"></datalist></div>
      <div class="fr"><label>Descrizione</label><textarea id="f-desc" style="min-height:50px">${r.descrizione||''}</textarea></div>
      <div class="fg3"><div class="fr"><label>Q.tà</label><input id="f-qty" type="number" value="${r.qty||1}" min="1"></div><div class="fr"><label>Fornitore</label><input id="f-forn" value="${r.fornitore||''}" list="forn-list2"><datalist id="forn-list2">${FORNITORI.map(f=>`<option value="${f.nome}">`).join('')}</datalist></div><div class="fr"><label>Codice</label><input id="f-cod" value="${r.codice||''}"></div></div>
      <div class="fg3"><div class="fr"><label>Costo/cad (€)</label><input id="f-costo" type="number" step="0.01" value="${r.costo_cad||''}"></div><div class="fr"><label>Sconto (%)</label><input id="f-sconto" type="number" min="0" max="100" value="${r.sconto||0}"></div><div class="fr"><label>Mark-up</label><input id="f-markup" type="number" step="0.01" value="${r.markup||2.25}"></div></div>
      <div class="fg2"><div class="fr"><label>Stato approvv.</label><select id="f-stato-a">${Object.entries(SAP).map(([k,v])=>`<option value="${k}"${r.stato_approv===k?' selected':''}>${v}</option>`).join('')}</select></div><div class="fr"><label>Data attesa</label><input id="f-datta" type="date" value="${r.data_attesa||''}"></div></div>
    </div>
    ${act("updateRiga('"+numero+"','"+rid+"')")}`;
  }

  if(type==='add-ore')return`<h3>Registra ore</h3><div class="form-inline">
    <div class="fg2"><div class="fr"><label>Data</label><input id="f-dat" type="date" value="${today}"></div><div class="fr"><label>Operatore</label><input id="f-op" placeholder="es. Andrea"></div></div>
    <div class="fr"><label>Descrizione attività</label><input id="f-desc"></div>
    <div class="fg2"><div class="fr"><label>Ore</label><input id="f-ore" type="number" step="0.5" min="0.5" value="1"></div><div class="fr"><label>Tariffa (€/h)</label><input id="f-tar" type="number" value="65"></div></div>
  </div>${act("saveOre('"+extra+"')")}`;

  if(type==='add-mag'){
    const m=extra?((JSON.parse(localStorage.getItem('asg_magazzino')||'[]')).find(x=>x.id===extra)||{}):{}; 
    return`<h3>${m.id?'Modifica voce magazzino':'+ Nuova voce magazzino'}</h3>
    <div class="form-inline">
      <div class="fg2">
        <div class="fr"><label>Codice / sigla</label><input id="mg-cod" value="${m.codice||''}" placeholder="es. NC, B09, VR38…"></div>
        <div class="fr"><label>Categoria</label>
          <select id="mg-cat">
            ${['valvola','pompa','motore','serbatoio','blocco','giunto','raccorderia','elettrica','vario'].map(c=>`<option value="${c}"${m.categoria===c?' selected':''}>${c.charAt(0).toUpperCase()+c.slice(1)}</option>`).join('')}
          </select>
        </div>
      </div>
      <div class="fr"><label>Descrizione</label><input id="mg-desc" value="${m.descrizione||''}" placeholder="descrizione completa del componente"></div>
      <div class="fg3">
        <div class="fr"><label>Giacenza attuale</label><input id="mg-qty" type="number" min="0" step="1" value="${m.qty||0}"></div>
        <div class="fr"><label>Scorta minima</label><input id="mg-min" type="number" min="0" step="1" value="${m.qty_min||0}" placeholder="0 = nessun alert"></div>
        <div class="fr"><label>Costo/cad (€)</label><input id="mg-costo" type="number" step="0.01" value="${m.costo_cad||''}"></div>
      </div>
      <div class="fg2">
        <div class="fr"><label>Fornitore abituale</label><input id="mg-forn" value="${m.fornitore||''}" list="mg-forn-list"><datalist id="mg-forn-list">${FORNITORI.map(f=>`<option value="${f.nome}">`).join('')}</datalist></div>
        <div class="fr"><label>Ubicazione fisica</label><input id="mg-ubic" value="${m.ubicazione||''}" placeholder="es. Scaffale A2, Cassetto 3…"></div>
      </div>
      <div class="fr"><label>Note</label><input id="mg-note" value="${m.note||''}"></div>
    </div>
    ${act("saveMag('"+(m.id||'')+"')")}`;
  }

  if(type==='guida-import-nc'){
    return`<h3>☁ Come importare le commesse da Nextcloud</h3>
    <div style="font-size:13px;color:var(--text2);line-height:1.8;">
      <p style="margin-bottom:12px">Il bottone <strong>"☁ importa NC"</strong> legge il file <code style="background:var(--bg3);padding:1px 6px;border-radius:4px">Dashboard/commesse.csv</code> dalla tua Nextcloud e lo unisce con i dati locali.</p>
      <p style="font-weight:700;margin-bottom:6px">Passaggi:</p>
      <ol style="margin-left:18px;margin-bottom:14px;">
        <li>Assicurati che Nextcloud sia configurato (⚙ Configura in alto a destra)</li>
        <li>Su Nextcloud, carica o crea il file <code style="background:var(--bg3);padding:1px 6px;border-radius:4px">Dashboard/commesse.csv</code></li>
        <li>Clicca <strong>☁ importa NC</strong> — la dashboard fa il merge: aggiorna le commesse esistenti e aggiunge le nuove, senza cancellare i dati locali (distinta, ore, note)</li>
      </ol>
      <p style="font-weight:700;margin-bottom:6px">Formato commesse.csv atteso:</p>
      <div style="background:#1e293b;color:#e2e8f0;font-family:monospace;font-size:12px;padding:10px 14px;border-radius:8px;overflow-x:auto;white-space:pre;">numero,cliente,oggetto,tipo_commessa,importo,data_apertura,data_consegna,stato,note
KFG-16-2026,KFG,Centralina PU10 pressa,realizzazione,1850,2026-01-10,,in-corso,
ERS-05-2025,ERS,Ricerca pompe reversibili,ricerca,0,2025-11-01,,chiusa,</div>
      <p style="margin-top:12px;color:var(--text3);font-size:12px;">Il campo <code style="background:var(--bg3);padding:1px 4px;border-radius:3px">tipo_commessa</code> deve essere uno tra: <em>ricerca, realizzazione, prestazione, configuratore, misto</em>.<br>Le date nel formato YYYY-MM-DD.</p>
    </div>
    <div class="modal-actions"><button class="btn" onclick="closeM()">Chiudi</button></div>`;
  }

  if(type==='add-doc'){return`<h3>📎 Aggiungi riferimento documento</h3>
  <div class="form-inline">
    <div class="fr"><label>Nome / descrizione</label><input id="d-nome" placeholder="es. Schema oleodinamico rev.2, Foto montaggio, Email conferma…"></div>
    <div class="fg2">
      <div class="fr"><label>Tipo</label>
        <select id="d-tipo">
          <option value="schema">📐 Schema oleodinamico</option>
          <option value="disegno">✏️ Disegno/DXF</option>
          <option value="foto">📷 Foto</option>
          <option value="offerta">📄 Offerta/preventivo</option>
          <option value="ordine">🧾 Ordine</option>
          <option value="email">📧 Email</option>
          <option value="certificato">📋 Certificato/collaudo</option>
          <option value="altro">📎 Altro</option>
        </select>
      </div>
      <div class="fr"><label>Data</label><input id="d-data" type="date" value="${today}"></div>
    </div>
    <div class="fr"><label>Link (URL Nextcloud o esterno — opzionale)</label><input id="d-url" placeholder="https://cloud.asglab.it/apps/files/?dir=…" type="url"></div>
    <div class="fr"><label>Note</label><input id="d-note" placeholder="es. rev.3 approvata, da verificare…"></div>
  </div>
  <p style="font-size:12px;color:var(--text3);margin-top:4px;line-height:1.5">💡 Per ottenere il link diretto a un file su Nextcloud: tasto destro sul file → <strong>Copia link diretto</strong>. Da mobile usa l'app Nextcloud per caricare e poi incolla qui il link.</p>
  ${act("saveDoc('"+extra+"')")}`;
  }
  if(type==='import-movimenti')return`<h3>📥 Import movimenti bancari</h3>
  <p style="font-size:13px;color:var(--text3);margin-bottom:12px;line-height:1.5">
    Trascina o seleziona il file esportato dalla tua banca. Il saldo viene aggiornato automaticamente.
  </p>
  <div style="display:grid;gap:10px;margin-bottom:14px;">
    <!-- Fineco -->
    <div style="border:1.5px solid var(--border);border-radius:8px;padding:12px 14px;">
      <div style="font-weight:700;font-size:13px;margin-bottom:4px;display:flex;align-items:center;gap:8px;">
        <span style="background:#e55c1b;color:#fff;padding:1px 7px;border-radius:4px;font-size:11px">FINECO</span>
        Movimenti conto XLSX
      </div>
      <div style="font-size:11px;color:var(--text3);margin-bottom:8px">
        Da: Fineco web → Movimenti → Ricerca avanzata → <strong>Esporta risultato</strong> → file .xlsx
      </div>
      <div id="drop-fineco" ondrop="dropMovimenti(event,'fineco')" ondragover="event.preventDefault()"
        style="border:2px dashed var(--border);border-radius:7px;padding:14px;text-align:center;font-size:12px;color:var(--text3);cursor:pointer;transition:border-color .15s;"
        onmouseenter="this.style.borderColor='var(--blue)'" onmouseleave="this.style.borderColor='var(--border)'"
        onclick="document.getElementById('file-fineco').click()">
        📂 Trascina qui o clicca — <em>movements_YYYYMMDD.xlsx</em>
      </div>
      <input type="file" id="file-fineco" accept=".xlsx" style="display:none" onchange="fileMovimenti(this,'fineco')">
      <div id="res-fineco" style="margin-top:6px;font-size:12px;"></div>
    </div>
    <!-- Sella/Tot -->
    <div style="border:1.5px solid var(--border);border-radius:8px;padding:12px 14px;">
      <div style="font-weight:700;font-size:13px;margin-bottom:4px;display:flex;align-items:center;gap:8px;">
        <span style="background:#1a56db;color:#fff;padding:1px 7px;border-radius:4px;font-size:11px">SELLA</span>
        Movimenti conto CSV
      </div>
      <div style="font-size:11px;color:var(--text3);margin-bottom:8px">
        Da: Banca Sella / Tot → Movimenti → <strong>Esporta CSV</strong> → file export-movimenti-conto_*.csv
      </div>
      <div id="drop-sella" ondrop="dropMovimenti(event,'sella')" ondragover="event.preventDefault()"
        style="border:2px dashed var(--border);border-radius:7px;padding:14px;text-align:center;font-size:12px;color:var(--text3);cursor:pointer;transition:border-color .15s;"
        onmouseenter="this.style.borderColor='var(--blue)'" onmouseleave="this.style.borderColor='var(--border)'"
        onclick="document.getElementById('file-sella').click()">
        📂 Trascina qui o clicca — <em>export-movimenti-conto_*.csv</em>
      </div>
      <input type="file" id="file-sella" accept=".csv" style="display:none" onchange="fileMovimenti(this,'sella')">
      <div id="res-sella" style="margin-top:6px;font-size:12px;"></div>
    </div>
    <!-- Revolut -->
    <div style="border:1.5px solid var(--border);border-radius:8px;padding:12px 14px;background:var(--bg2);">
      <div style="font-weight:700;font-size:13px;margin-bottom:4px;display:flex;align-items:center;gap:8px;">
        <span style="background:#191c1f;color:#fff;padding:1px 7px;border-radius:4px;font-size:11px">REVOLUT</span>
        Aggiornamento saldo manuale
      </div>
      <div style="font-size:11px;color:var(--text3);margin-bottom:8px">
        Revolut invia l'estratto conto in PDF via email — non importabile direttamente.<br>
        Leggi il <strong>Saldo finale</strong> dal PDF e aggiornalo qui:
      </div>
      <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;">
        <input id="rev-saldo" type="number" step="0.01" placeholder="saldo finale €" style="flex:1;border:1.5px solid var(--border);border-radius:7px;padding:7px;font-size:14px;outline:none;min-width:120px;">
        <input id="rev-data" type="date" value="${today}" style="border:1.5px solid var(--border);border-radius:7px;padding:7px;font-size:13px;outline:none;">
        <button class="btn" onclick="salvaRevolut()" style="background:#191c1f;">Salva</button>
      </div>
      <div id="res-revolut" style="margin-top:6px;font-size:12px;"></div>
    </div>
  </div>`;

  if(type==='conto'){
    const conti=lll('conti');
    return`<h3>Aggiorna saldo conto</h3>
    <p style="font-size:13px;color:var(--text3);margin-bottom:12px;line-height:1.5">Inserisci il saldo attuale leggendolo dall'estratto conto o dall'app bancaria.</p>
    <div class="form-inline">
      <div class="fr"><label>Conto</label>
        <input id="f-nome" list="conti-list" placeholder="es. Banca Sella, Fineco, Revolut…">
        <datalist id="conti-list">
          <option value="Banca Sella">
          <option value="Banca Fineco">
          <option value="Revolut Andrea">
          <option value="Revolut Sergio">
          <option value="Cassa contanti">
          <option value="Carta Fineco">
          <option value="Carta Sella">
        </datalist>
      </div>
      <div class="fg2">
        <div class="fr"><label>Saldo (€)</label><input id="f-saldo" type="number" step="0.01" placeholder="0.00"></div>
        <div class="fr"><label>Data rilevazione</label><input id="f-dat" type="date" value="${today}"></div>
      </div>
      <div class="fr"><label>Note (opzionale)</label><input id="f-note" placeholder="es. saldo da estratto conto trimestrale"></div>
    </div>
    ${act('saveConto()')}`;
  }
  if(type==='scadenza')return`<h3>Nuova scadenza</h3><div class="form-inline">
    <div class="fr"><label>Descrizione</label><input id="f-desc" placeholder="es. IVA I trimestre, Acconto IRES…"></div>
    <div class="fg2">
      <div class="fr"><label>Data</label><input id="f-dat" type="date"></div>
      <div class="fr"><label>Importo (€)</label><input id="f-imp" type="number" step="0.01" placeholder="0.00"></div>
    </div>
    <div class="fr"><label>Categoria</label>
      <select id="f-cat" style="width:100%;border:1.5px solid var(--border);border-radius:7px;padding:7px;font-size:13px;outline:none;">
        <option value="">— nessuna —</option>
        <option value="iva">IVA</option><option value="ires">IRES/IRAP</option>
        <option value="inps">INPS</option><option value="ritenuta">Ritenuta</option>
        <option value="sostituto">Sostituto d'imposta</option><option value="dichiarazione">Dichiarazione</option>
        <option value="societario">Societario/CCIAA</option>
        <option value="incasso">Incasso fattura</option><option value="pagamento">Pagamento fattura</option>
      </select>
    </div>
  </div>${act('saveScad()')}`;

  if(type==='edit-scad'){
    const sid=extra||'';
    const all=lll('scad');
    const s=all.find(x=>x.id===sid)||{};
    return`<h3>✎ Modifica scadenza</h3>
    <div class="form-inline">
      <div class="fr"><label>Descrizione</label><input id="f-desc" value="${(s.descrizione||'').replace(/"/g,'&quot;')}"></div>
      <div class="fg2">
        <div class="fr"><label>Data</label><input id="f-dat" type="date" value="${s.data||''}"></div>
        <div class="fr"><label>Importo (€)</label><input id="f-imp" type="number" step="0.01" value="${parseFloat(s.importo||0)||''}"></div>
      </div>
      <div class="fr"><label>Categoria</label>
        <select id="f-cat" style="width:100%;border:1.5px solid var(--border);border-radius:7px;padding:7px;font-size:13px;outline:none;">
          <option value="">— nessuna —</option>
          ${['iva','ires','inps','ritenuta','sostituto','dichiarazione','societario','incasso','pagamento'].map(c=>`<option value="${c}"${s.categoria===c?' selected':''}>${c.charAt(0).toUpperCase()+c.slice(1)}</option>`).join('')}
        </select>
      </div>
      <div class="fr"><label>Stato</label>
        <select id="f-pagata" style="width:100%;border:1.5px solid var(--border);border-radius:7px;padding:7px;font-size:13px;outline:none;">
          <option value="no"${(!s.pagato||s.pagato==='no')?' selected':''}>⏳ Da pagare/completare</option>
          <option value="si"${s.pagato==='si'?' selected':''}>✓ Completata/pagata</option>
        </select>
      </div>
    </div>
    <div class="modal-actions">
      <button class="btn-ghost" onclick="closeM()">Annulla</button>
      <button class="btn" onclick="saveScadEdit('${sid}')">Salva</button>
    </div>`;
  }

  if(type==='todo')return`<h3>Nuova azione</h3><div class="form-inline">
    <div class="fr"><label>Azione</label><textarea id="f-desc" style="min-height:70px"></textarea></div>
    <div class="fr"><label>Priorità</label><select id="f-pri"><option value="alta">Alta</option><option value="media" selected>Media</option><option value="bassa">Bassa</option></select></div>
  </div>${act('saveTodo()')}`;

  if(type==='rdo'){
    const f=FORNITORI.find(x=>x.id===extra)||{nome:'Fornitore',categorie:[],contatti:{}};
    const oggi=new Date().toLocaleDateString('it-IT',{day:'2-digit',month:'long',year:'numeric'});
    // Recupera template salvato per questo fornitore
    const templates=JSON.parse(localStorage.getItem('asg_rdo_templates')||'{}');
    const defaultTpl=`Spett.le {FORNITORE},\n\nin riferimento ai Vs. prodotti/servizi, con la presente siamo a richiedere un'offerta per i seguenti articoli:\n\n1. [CODICE/DESCRIZIONE] — qty ___\n2. \n\nSi prega di indicare:\n- Prezzo unitario IVA esclusa\n- Disponibilità a magazzino\n- Tempi di consegna\n- Condizioni di resa\n\nRingraziamo anticipatamente e restiamo in attesa di un Vostro riscontro.\n\nCordiali saluti,`;
    const tpl=templates[extra]||defaultTpl;
    const testo=tpl.replace('{FORNITORE}',f.nome).replace('{DATA}',oggi);
    return`<h3>✉ Richiesta di offerta — ${f.nome}</h3>
    <div style="display:flex;gap:8px;margin-bottom:12px;border-bottom:1px solid var(--border);padding-bottom:12px;">
      <button class="btn-ghost" style="font-size:12px;padding:4px 10px;" onclick="document.getElementById('rdo-panel-send').style.display='';document.getElementById('rdo-panel-tpl').style.display='none';this.style.background='var(--blue)';this.style.color='#fff';document.getElementById('rdo-tab-tpl').style.background='';document.getElementById('rdo-tab-tpl').style.color='';" id="rdo-tab-send">✉ Componi</button>
      <button class="btn-ghost" style="font-size:12px;padding:4px 10px;" onclick="document.getElementById('rdo-panel-tpl').style.display='';document.getElementById('rdo-panel-send').style.display='none';this.style.background='var(--blue)';this.style.color='#fff';document.getElementById('rdo-tab-send').style.background='';document.getElementById('rdo-tab-send').style.color='';" id="rdo-tab-tpl">⚙ Modello testo</button>
    </div>
    <div id="rdo-panel-send">
      <div class="form-row"><label>A (email fornitore)</label><input id="rdo-to" value="${f.contatti?.email||''}" placeholder="email@fornitore.it"></div>
      <div class="form-row"><label>Oggetto</label><input id="rdo-subj" value="Richiesta di offerta — ${oggi}"></div>
      <div class="form-row"><label>Testo</label><textarea id="rdo-body" style="min-height:220px;font-family:var(--font);font-size:13px;line-height:1.7">${testo}</textarea></div>
      <div class="modal-actions">
        <button class="btn-ghost" onclick="closeM()">Annulla</button>
        <button class="btn-ghost" onclick="copyRDO()">📋 Copia testo</button>
        ${f.contatti?.email?'<a class="btn" href="mailto:'+f.contatti.email+'?subject='+encodeURIComponent('Richiesta di offerta \u2014 '+oggi)+'&body='+encodeURIComponent(testo)+'" onclick="closeM()">📧 Apri email</a>':'<button class="btn" onclick="copyRDO()">📋 Copia tutto</button>'}
      </div>
    </div>
    <div id="rdo-panel-tpl" style="display:none;">
      <p style="font-size:13px;color:var(--text3);margin-bottom:10px;line-height:1.5">Personalizza il testo predefinito per <strong>${f.nome}</strong>. Variabili disponibili: <code style="background:var(--bg3);padding:1px 5px;border-radius:3px">{FORNITORE}</code> <code style="background:var(--bg3);padding:1px 5px;border-radius:3px">{DATA}</code></p>
      <div class="form-row"><label>Modello testo per ${f.nome}</label><textarea id="rdo-tpl-body" style="min-height:240px;font-family:var(--font);font-size:13px;line-height:1.7">${templates[extra]||defaultTpl}</textarea></div>
      <div class="modal-actions">
        <button class="btn-ghost" onclick="closeM()">Annulla</button>
        <button class="btn-ghost" onclick="ripristinaTemplatePredefinito('${extra}')">↺ Ripristina predefinito</button>
        <button class="btn" onclick="salvaTemplate('${extra}')">💾 Salva modello</button>
      </div>
    </div>`;
  }

  if(type==='rdo-new'){
    return`<h3>✉ Nuova richiesta di offerta</h3>
    <div class="form-row"><label>Fornitore</label>
      <select id="rdo-forn">${FORNITORI.map(f=>`<option value="${f.id}">${f.nome}</option>`).join('')}</select>
    </div>
    <div class="modal-actions">
      <button class="btn-ghost" onclick="closeM()">Annulla</button>
      <button class="btn" onclick="const v=document.getElementById('rdo-forn').value;closeM();openModal('rdo',v)">Continua →</button>
    </div>`;
  }

  if(type==='import-primanota'){return`<h3>📥 Import Prima Nota Fattura24</h3>
  <p style="font-size:13px;color:var(--text3);line-height:1.6;margin-bottom:14px;">
    Esporta la <strong>Prima Nota</strong> da Fattura24 → Contabilità → Prima Nota → filtra per conto → Esporta XLS.<br>
    Puoi importare più file (uno per conto) — i movimenti vengono uniti e deduplicati.
  </p>
  <div style="border:2px dashed var(--border);border-radius:var(--r);padding:24px;text-align:center;margin-bottom:12px;"
    ondragover="event.preventDefault()" ondrop="handlePNDrop(event)">
    <div style="font-size:32px;margin-bottom:8px">📂</div>
    <div style="font-size:13px;color:var(--text3);margin-bottom:10px">Trascina i file XLS qui oppure</div>
    <input type="file" id="pn-file-input" accept=".xls,.xlsx" multiple style="display:none" onchange="handlePNFiles(this.files)">
    <button class="btn btn-primary" onclick="document.getElementById('pn-file-input').click()">Seleziona file</button>
  </div>
  <div id="pn-import-preview" style="max-height:200px;overflow-y:auto;"></div>
  <div class="modal-actions">
    <button class="btn-ghost" onclick="closeM()">Annulla</button>
    <button class="btn" id="pn-import-btn" onclick="confirmImportPN()" style="display:none">✓ Importa movimenti</button>
  </div>`;}

  if(type==='config'){const c=lCfg();return`<h3>⚙ Configurazione</h3>

  <div style="font-size:12px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.07em;margin-bottom:8px;padding-bottom:6px;border-bottom:1px solid var(--border)">☁ Nextcloud</div>
  <p style="font-size:13px;color:var(--text3);margin-bottom:12px;line-height:1.5">Usa una <strong>App Password</strong> da Nextcloud → Impostazioni → Sicurezza.</p>
  <div class="form-inline" style="margin-bottom:18px;">
    <div class="fr"><label>URL proxy</label><input id="cfg-url" value="${c.nc_url||'https://dash.asglab.it:8766'}"></div>
    <div class="fg2">
      <div class="fr"><label>Utente</label><input id="cfg-user" value="${c.nc_user||'andrea'}"></div>
      <div class="fr"><label>App Password</label><input id="cfg-pass" type="password" value="${c.nc_pass||''}" placeholder="xxxx-xxxx-xxxx-xxxx"></div>
    </div>
  </div>

  <div style="font-size:12px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.07em;margin-bottom:8px;padding-bottom:6px;border-bottom:1px solid var(--border)">🧾 Fattura24</div>
  <p style="font-size:13px;color:var(--text3);margin-bottom:12px;line-height:1.5">
    Trova la tua API key in <strong>Fattura24 → Impostazioni → API</strong>.<br>
    La chiave viene salvata localmente e usata per aggiornare <code style="background:var(--bg3);padding:1px 5px;border-radius:4px;font-size:12px">/opt/asg-dashboard/config.env</code> sul server tramite sync.<br>
    <span style="color:var(--amber)">⚠ Richiede account Fattura24 con piano a pagamento (API attiva).</span>
  </p>
  <div class="form-inline" style="margin-bottom:6px;">
    <div class="fr">
      <label>API Key Fattura24</label>
      <input id="cfg-f24key" type="password" value="${c.f24_key||''}" placeholder="es. abc123def456…" autocomplete="off">
    </div>
    <div class="fg2">
      <div class="fr">
        <label>Tipo documento da importare</label>
        <select id="cfg-f24tipo">
          <option value="all" ${(c.f24_tipo||'all')==='all'?'selected':''}>Tutte (fatture + NDC)</option>
          <option value="fattura" ${c.f24_tipo==='fattura'?'selected':''}>Solo fatture emesse</option>
          <option value="ricevuta" ${c.f24_tipo==='ricevuta'?'selected':''}>Solo ricevute</option>
        </select>
      </div>
      <div class="fr">
        <label>Sync automatico ogni</label>
        <select id="cfg-f24sync">
          <option value="60" ${(c.f24_sync||'60')==='60'?'selected':''}>1 ora</option>
          <option value="30" ${c.f24_sync==='30'?'selected':''}>30 min</option>
          <option value="240" ${c.f24_sync==='240'?'selected':''}>4 ore</option>
          <option value="0" ${c.f24_sync==='0'?'selected':''}>Manuale</option>
        </select>
      </div>
    </div>
    <div id="cfg-f24status" style="font-size:12px;color:var(--text3);margin-top:4px">${c.f24_key?'✓ API key presente — stato: '+(c.f24_ok?'<span style="color:var(--green);font-weight:700">connessa</span>':'<span style="color:var(--amber);font-weight:700">da verificare</span>'):'API key non inserita'}</div>
  </div>

  <div style="font-size:12px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.07em;margin:16px 0 8px;padding-bottom:6px;border-bottom:1px solid var(--border)">💰 Markup per cliente</div>
  <p style="font-size:12px;color:var(--text3);margin-bottom:10px">Ricarico predefinito sui costi materiali per ogni cliente. Viene precompilato nel form "aggiungi componente".</p>
  <div style="display:grid;grid-template-columns:1fr 1fr;gap:6px 12px;" id="cfg-markup-grid">
    ${Object.entries(CMAP).map(([cod,nome])=>'<div style="display:flex;align-items:center;gap:8px;"><label style="font-size:12px;color:var(--text2);flex:1;white-space:nowrap">'+nome.split('(')[0].trim()+' <span style="color:var(--text4);font-size:10px">('+cod+')</span></label><input id="cfg-mk-'+cod+'" type="number" step="0.05" min="1" max="10" value="'+(c.markup_clienti?.[cod]||2.25)+'" style="width:70px;border:1.5px solid var(--border);border-radius:6px;padding:4px 6px;font-size:12px;text-align:right;"></div>').join('')}
  </div>

  ${act('saveCfg2()')}\`;
  }
  return'';
}

// Ricerca Ares nel modal
// ── Tab switcher modal add-riga ──
// Precompila markup con valore per cliente corrente
function precompilaMarkup(){
  const el=document.getElementById('f-markup');
  if(!el||el.value)return; // già compilato
  const cfg=lCfg();
  const cod=getCod(CC?.numero||'');
  const val=cfg.markup_clienti?.[cod]||2.25;
  el.value=val;
}

function switchAddTab(tab){
  ['listino','mag','tubo','racc','pu10','pu20','manuale'].forEach(t=>{
    const btn=document.getElementById('atab-'+t);
    const panel=document.getElementById('apanel-'+t);
    if(btn)btn.classList.toggle('active',t===tab);
    if(panel)panel.style.display=t===tab?'':'none';
  });
  if(tab==='racc')renderRaccSciolto();
  if(tab==='pu10'){_pu10Blocchi=[];pu10Update();const el=document.getElementById('pu10-blocchi-list');if(el)el.innerHTML='';}
  if(tab==='pu20'){_pu20Blocchi=[];pu20Update();const el=document.getElementById('pu20-blocchi-list');if(el)el.innerHTML='';}
  // Upgrade select nel pannello appena reso visibile
  setTimeout(()=>{
    const panel=document.getElementById('apanel-'+tab);
    if(panel) upgradeSelects(panel);
    upgradeSelects(); // anche il resto del modal
  },0);
}

function magFilter(){
  const q=(document.getElementById('mag-q')||{}).value?.toLowerCase()||'';
  const items=lMag();
  const el=document.getElementById('mag-results');if(!el)return;
  const filt=q?items.filter(m=>(m.codice||'').toLowerCase().includes(q)||(m.descrizione||'').toLowerCase().includes(q)||(m.fornitore||'').toLowerCase().includes(q)):items;
  if(!filt.length){el.innerHTML='<div class="empty" style="padding:12px">Nessun risultato</div>';return;}
  el.innerHTML=filt.map(m=>'<div style="display:flex;align-items:center;gap:10px;padding:8px 4px;border-bottom:1px solid var(--border);font-size:13px;">'+'<div style="flex:1;"><div style="font-weight:600">'+(m.codice?'<span style="font-family:monospace;color:var(--green);font-weight:700">'+m.codice+'</span> \u2014 ':'')+(m.descrizione||'\u2014')+'</div>'+'<div style="font-size:11px;color:var(--text3)">'+(m.fornitore||'')+' \u00b7 giacenza: <strong style="color:'+(parseFloat(m.qty||0)>0?'var(--green)':'var(--red)')+'">'+( m.qty||0)+'</strong></div></div>'+'<span style="font-weight:700;white-space:nowrap">\u20ac'+fmt2(m.costo_cad||0)+'</span>'+'<button style="background:var(--blue);border:none;color:#fff;font-size:11px;font-weight:700;padding:4px 10px;border-radius:20px;cursor:pointer;" onclick="usaMagazzino('+JSON.stringify(m).replace(/"/g,'&quot;')+',currentAddRigaNumero)">+ usa</button></div>').join('');
}

let currentAddRigaNumero='';

function usaMagazzino(m,numero){
  const netto=parseFloat(m.costo_cad||0);const markup=2.25;
  const row={id:uid(),sezione:'Magazzino',descrizione:m.descrizione,qty:1,fornitore:m.fornitore||'Magazzino interno',codice:m.codice||'',costo_cad:netto,sconto:0,netto_cad:netto,markup,prezzo_vendita:parseFloat((netto*markup).toFixed(4)),stato_approv:'arrivato',data_attesa:''};
  const det=ll('det');if(!det[numero])det[numero]={distinta:[],ore:[],note:''};
  det[numero].distinta.push(row);ls('det',det);syncDetBackground();CC._det=det[numero];
  closeM();['distinta','approv','riepilogo'].forEach(id=>renderTabContent(id,CC,det[numero]));
  showToast('✓ '+m.descrizione+' aggiunto dalla magazzino');
}

// ── Configuratore tubo flessibile ──
// ════════════════════════════════════════════
// CONFIGURATORE TUBO FLESSIBILE
// ════════════════════════════════════════════

// Database raccordi: {std, misura, dn_compatibili, varianti}
// dn_compatibili: array di DN (mm) su cui il raccordo è montabile nativamente
// Le misure maggiori sono disponibili come riduzione, ma non le elenchiamo per default
const RACCORDI_DB=[
  // BSP (gas)
  {std:'BSP', mis:'1/8"',  dn:[6],        label:'BSP 1/8"'},
  {std:'BSP', mis:'1/4"',  dn:[6,10],     label:'BSP 1/4"'},
  {std:'BSP', mis:'3/8"',  dn:[10,12],    label:'BSP 3/8"'},
  {std:'BSP', mis:'1/2"',  dn:[12,16],    label:'BSP 1/2"'},
  {std:'BSP', mis:'3/4"',  dn:[16,19,20], label:'BSP 3/4"'},
  {std:'BSP', mis:'1"',    dn:[20,25],    label:'BSP 1"'},
  {std:'BSP', mis:'1¼"',   dn:[25,32],    label:'BSP 1¼"'},
  {std:'BSP', mis:'1½"',   dn:[32,38],    label:'BSP 1½"'},
  // ORFS
  {std:'ORFS', mis:'1/4"', dn:[6,10],     label:'ORFS 1/4"'},
  {std:'ORFS', mis:'3/8"', dn:[10,12],    label:'ORFS 3/8"'},
  {std:'ORFS', mis:'1/2"', dn:[12,16],    label:'ORFS 1/2"'},
  {std:'ORFS', mis:'3/4"', dn:[16,19,20], label:'ORFS 3/4"'},
  {std:'ORFS', mis:'1"',   dn:[20,25],    label:'ORFS 1"'},
  // JIC 37°
  {std:'JIC',  mis:'1/4"', dn:[6,10],     label:'JIC 37° 1/4"'},
  {std:'JIC',  mis:'3/8"', dn:[10,12],    label:'JIC 37° 3/8"'},
  {std:'JIC',  mis:'1/2"', dn:[12,16],    label:'JIC 37° 1/2"'},
  {std:'JIC',  mis:'3/4"', dn:[16,19,20], label:'JIC 37° 3/4"'},
  {std:'JIC',  mis:'1"',   dn:[20,25],    label:'JIC 37° 1"'},
  // DIN — metrico
  {std:'DIN',  mis:'M12×1.5', dn:[6],     label:'M12×1.5'},
  {std:'DIN',  mis:'M14×1.5', dn:[6,10],  label:'M14×1.5'},
  {std:'DIN',  mis:'M16×1.5', dn:[10],    label:'M16×1.5'},
  {std:'DIN',  mis:'M18×1.5', dn:[10,12], label:'M18×1.5'},
  {std:'DIN',  mis:'M20×1.5', dn:[12],    label:'M20×1.5'},
  {std:'DIN',  mis:'M22×1.5', dn:[12,16], label:'M22×1.5'},
  {std:'DIN',  mis:'M24×1.5', dn:[16],    label:'M24×1.5'},
  {std:'DIN',  mis:'M26×1.5', dn:[16,20], label:'M26×1.5'},
  {std:'DIN',  mis:'M30×2',   dn:[20,25], label:'M30×2'},
  {std:'DIN',  mis:'M36×2',   dn:[25],    label:'M36×2'},
  // Flangia SAE (solo dritto, no curve)
  {std:'SAE',  mis:'SAE 6000psi 1/2"', dn:[12,16], label:'Flangia SAE 6000psi 1/2"', noCurva:true},
  {std:'SAE',  mis:'SAE 3000psi 3/4"', dn:[16,20], label:'Flangia SAE 3000psi 3/4"', noCurva:true},
  {std:'SAE',  mis:'SAE 3000psi 1"',   dn:[20,25], label:'Flangia SAE 3000psi 1"',   noCurva:true},
  // Nippli / capezzoli
  {std:'NIPPLO', mis:'Capezzolo 1/4" BSP', dn:[6,10], label:'Capezzolo 1/4" BSP', noCurva:true},
  {std:'NIPPLO', mis:'Capezzolo 3/8" BSP', dn:[10,12], label:'Capezzolo 3/8" BSP', noCurva:true},
  {std:'NIPPLO', mis:'Capezzolo 1/2" BSP', dn:[12,16], label:'Capezzolo 1/2" BSP', noCurva:true},
  {std:'NIPPLO', mis:'Capezzolo 3/4" BSP', dn:[16,20], label:'Capezzolo 3/4" BSP', noCurva:true},
  {std:'NIPPLO', mis:'Capezzolo 1" BSP',   dn:[20,25], label:'Capezzolo 1" BSP', noCurva:true},
  {std:'NIPPLO', mis:'Capezzolo inox 1/4"',dn:[6,10],  label:'Capezzolo inox 1/4"', noCurva:true},
  {std:'NIPPLO', mis:'Capezzolo inox 3/8"',dn:[10,12], label:'Capezzolo inox 3/8"', noCurva:true},
  {std:'NIPPLO', mis:'Capezzolo inox 1/2"',dn:[12,16], label:'Capezzolo inox 1/2"', noCurva:true},
  // Terminali ad occhio (swaged) — per tubi raccordati
  {std:'OCCHIO', mis:'Occhio 3/8" Ø11 BSP',    dn:[6,10],     label:'Term. occhio 3/8" Ø11', noCurva:true},
  {std:'OCCHIO', mis:'Occhio 1/2" Ø11 BSP',    dn:[10,12],    label:'Term. occhio 1/2" Ø11', noCurva:true},
  {std:'OCCHIO', mis:'Occhio 1/2" Ø14 BSP',    dn:[10,12],    label:'Term. occhio 1/2" Ø14', noCurva:true},
  {std:'OCCHIO', mis:'Occhio 3/4" Ø14 BSP',    dn:[12,16],    label:'Term. occhio 3/4" Ø14', noCurva:true},
  {std:'OCCHIO', mis:'Occhio 3/4" Ø18 BSP',    dn:[16,20],    label:'Term. occhio 3/4" Ø18', noCurva:true},
  {std:'OCCHIO', mis:'Occhio 1" Ø18 BSP',      dn:[16,20,25], label:'Term. occhio 1" Ø18',   noCurva:true},
  {std:'OCCHIO', mis:'Occhio 1" Ø22 BSP',      dn:[20,25],    label:'Term. occhio 1" Ø22',   noCurva:true},
  {std:'OCCHIO', mis:'Occhio inox 3/8" Ø11',   dn:[6,10],     label:'Occhio inox 3/8" Ø11',  noCurva:true},
  {std:'OCCHIO', mis:'Occhio inox 1/2" Ø14',   dn:[10,12],    label:'Occhio inox 1/2" Ø14',  noCurva:true},
  {std:'OCCHIO', mis:'Occhio inox 3/4" Ø18',   dn:[16,20],    label:'Occhio inox 3/4" Ø18',  noCurva:true},
];

// ════════════════════════════════════════════
// RACCORDERIA SCIOLTA — selettore con filtri
// ════════════════════════════════════════════

// Database raccorderia sciolta per alta pressione oleodinamica
// figura: 'adattatore' | 'gomito' | 'T' | 'nipplo' | 'bonded' | 'stucchi' | 'riduzione'
const RACC_SCIOLTI=[
  // ── BSP maschio/femmina — adattatori ──
  {std:'BSP', fig:'adattatore', mis:'BSP 1/4"M × 1/4"F',   cod:'A-BSP14M-14F', desc:'Adattatore BSP 1/4"M → 1/4"F',   forn:'Pa.Co', p:1.8},
  {std:'BSP', fig:'adattatore', mis:'BSP 3/8"M × 3/8"F',   cod:'A-BSP38M-38F', desc:'Adattatore BSP 3/8"M → 3/8"F',   forn:'Pa.Co', p:2.2},
  {std:'BSP', fig:'adattatore', mis:'BSP 1/2"M × 1/2"F',   cod:'A-BSP12M-12F', desc:'Adattatore BSP 1/2"M → 1/2"F',   forn:'Pa.Co', p:2.8},
  {std:'BSP', fig:'adattatore', mis:'BSP 3/4"M × 3/4"F',   cod:'A-BSP34M-34F', desc:'Adattatore BSP 3/4"M → 3/4"F',   forn:'Pa.Co', p:3.5},
  {std:'BSP', fig:'adattatore', mis:'BSP 1"M × 1"F',        cod:'A-BSP1M-1F',   desc:'Adattatore BSP 1"M → 1"F',       forn:'Pa.Co', p:4.2},
  // BSP riduzioni
  {std:'BSP', fig:'riduzione', mis:'BSP 3/8"M × 1/4"F',    cod:'R-BSP38-14',   desc:'Riduzione BSP 3/8"M × 1/4"F',    forn:'Pa.Co', p:2.0},
  {std:'BSP', fig:'riduzione', mis:'BSP 1/2"M × 3/8"F',    cod:'R-BSP12-38',   desc:'Riduzione BSP 1/2"M × 3/8"F',    forn:'Pa.Co', p:2.4},
  {std:'BSP', fig:'riduzione', mis:'BSP 3/4"M × 1/2"F',    cod:'R-BSP34-12',   desc:'Riduzione BSP 3/4"M × 1/2"F',    forn:'Pa.Co', p:3.0},
  {std:'BSP', fig:'riduzione', mis:'BSP 1"M × 3/4"F',      cod:'R-BSP1-34',    desc:'Riduzione BSP 1"M × 3/4"F',      forn:'Pa.Co', p:3.8},
  // BSP gomiti 90°
  {std:'BSP', fig:'gomito',    mis:'BSP 1/4"M × 90°',      cod:'G90-BSP14',    desc:'Gomito BSP 1/4"M 90°',            forn:'Pa.Co', p:3.5},
  {std:'BSP', fig:'gomito',    mis:'BSP 3/8"M × 90°',      cod:'G90-BSP38',    desc:'Gomito BSP 3/8"M 90°',            forn:'Pa.Co', p:4.0},
  {std:'BSP', fig:'gomito',    mis:'BSP 1/2"M × 90°',      cod:'G90-BSP12',    desc:'Gomito BSP 1/2"M 90°',            forn:'Pa.Co', p:5.0},
  {std:'BSP', fig:'gomito',    mis:'BSP 3/4"M × 90°',      cod:'G90-BSP34',    desc:'Gomito BSP 3/4"M 90°',            forn:'Pa.Co', p:6.2},
  // BSP nippli doppi
  {std:'BSP', fig:'nipplo',    mis:'Nipplo BSP 1/4"M × 1/4"M', cod:'N-BSP14', desc:'Nipplo doppio BSP 1/4"M × 1/4"M', forn:'Pa.Co', p:1.5},
  {std:'BSP', fig:'nipplo',    mis:'Nipplo BSP 3/8"M × 3/8"M', cod:'N-BSP38', desc:'Nipplo doppio BSP 3/8"M × 3/8"M', forn:'Pa.Co', p:1.8},
  {std:'BSP', fig:'nipplo',    mis:'Nipplo BSP 1/2"M × 1/2"M', cod:'N-BSP12', desc:'Nipplo doppio BSP 1/2"M × 1/2"M', forn:'Pa.Co', p:2.2},
  {std:'BSP', fig:'nipplo',    mis:'Nipplo BSP 3/4"M × 3/4"M', cod:'N-BSP34', desc:'Nipplo doppio BSP 3/4"M × 3/4"M', forn:'Pa.Co', p:2.8},
  // ── DIN — metrico ──
  {std:'DIN', fig:'adattatore', mis:'M14×1.5M × BSP 1/4"F', cod:'A-M14-BSP14', desc:'Adatt. M14×1.5M → BSP 1/4"F',   forn:'Pa.Co', p:3.2},
  {std:'DIN', fig:'adattatore', mis:'M16×1.5M × BSP 3/8"F', cod:'A-M16-BSP38', desc:'Adatt. M16×1.5M → BSP 3/8"F',   forn:'Pa.Co', p:3.4},
  {std:'DIN', fig:'adattatore', mis:'M18×1.5M × BSP 3/8"F', cod:'A-M18-BSP38', desc:'Adatt. M18×1.5M → BSP 3/8"F',   forn:'Pa.Co', p:3.6},
  {std:'DIN', fig:'adattatore', mis:'M22×1.5M × BSP 1/2"F', cod:'A-M22-BSP12', desc:'Adatt. M22×1.5M → BSP 1/2"F',   forn:'Pa.Co', p:4.0},
  {std:'DIN', fig:'adattatore', mis:'M26×1.5M × BSP 3/4"F', cod:'A-M26-BSP34', desc:'Adatt. M26×1.5M → BSP 3/4"F',   forn:'Pa.Co', p:4.5},
  {std:'DIN', fig:'nipplo',     mis:'Nipplo M14×1.5 doppio', cod:'N-M14',       desc:'Nipplo doppio M14×1.5M × M14×1.5M', forn:'Pa.Co', p:2.8},
  {std:'DIN', fig:'nipplo',     mis:'Nipplo M16×1.5 doppio', cod:'N-M16',       desc:'Nipplo doppio M16×1.5M × M16×1.5M', forn:'Pa.Co', p:3.0},
  {std:'DIN', fig:'nipplo',     mis:'Nipplo M18×1.5 doppio', cod:'N-M18',       desc:'Nipplo doppio M18×1.5M × M18×1.5M', forn:'Pa.Co', p:3.2},
  {std:'DIN', fig:'nipplo',     mis:'Nipplo M22×1.5 doppio', cod:'N-M22',       desc:'Nipplo doppio M22×1.5M × M22×1.5M', forn:'Pa.Co', p:3.8},
  // ── JIC ──
  {std:'JIC', fig:'adattatore', mis:'JIC 1/4"M × BSP 1/4"F', cod:'A-JIC14-BSP14', desc:'Adatt. JIC 1/4"M → BSP 1/4"F', forn:'Pa.Co', p:3.5},
  {std:'JIC', fig:'adattatore', mis:'JIC 3/8"M × BSP 3/8"F', cod:'A-JIC38-BSP38', desc:'Adatt. JIC 3/8"M → BSP 3/8"F', forn:'Pa.Co', p:3.8},
  {std:'JIC', fig:'adattatore', mis:'JIC 1/2"M × BSP 1/2"F', cod:'A-JIC12-BSP12', desc:'Adatt. JIC 1/2"M → BSP 1/2"F', forn:'Pa.Co', p:4.2},
  {std:'JIC', fig:'adattatore', mis:'JIC 3/4"M × BSP 3/4"F', cod:'A-JIC34-BSP34', desc:'Adatt. JIC 3/4"M → BSP 3/4"F', forn:'Pa.Co', p:4.8},
  // ── Rondelle bonded ──
  {std:'Bonded', fig:'bonded', mis:'Bonded 1/4" BSP',  cod:'BD-14', desc:'Rondella bonded BSP 1/4" (alluminio+NBR)', forn:'Pa.Co', p:0.5},
  {std:'Bonded', fig:'bonded', mis:'Bonded 3/8" BSP',  cod:'BD-38', desc:'Rondella bonded BSP 3/8"',               forn:'Pa.Co', p:0.6},
  {std:'Bonded', fig:'bonded', mis:'Bonded 1/2" BSP',  cod:'BD-12', desc:'Rondella bonded BSP 1/2"',               forn:'Pa.Co', p:0.7},
  {std:'Bonded', fig:'bonded', mis:'Bonded 3/4" BSP',  cod:'BD-34', desc:'Rondella bonded BSP 3/4"',               forn:'Pa.Co', p:0.9},
  {std:'Bonded', fig:'bonded', mis:'Bonded 1" BSP',    cod:'BD-1',  desc:'Rondella bonded BSP 1"',                 forn:'Pa.Co', p:1.1},
  {std:'Bonded', fig:'bonded', mis:'Bonded M14×1.5',   cod:'BD-M14',desc:'Rondella bonded M14×1.5 DIN',            forn:'Pa.Co', p:0.6},
  {std:'Bonded', fig:'bonded', mis:'Bonded M16×1.5',   cod:'BD-M16',desc:'Rondella bonded M16×1.5 DIN',            forn:'Pa.Co', p:0.7},
  {std:'Bonded', fig:'bonded', mis:'Bonded M22×1.5',   cod:'BD-M22',desc:'Rondella bonded M22×1.5 DIN',            forn:'Pa.Co', p:0.9},
  // ── Innesti rapidi Stucchi ── (fornitore: Tecnical/FIZ)
  {std:'Stucchi', fig:'stucchi', mis:'Stucchi ISO-B 3/8"M FF', cod:'S-ISO-B-38FF', desc:'Innesto rapido Stucchi ISO-B DN10 3/8"M flat-face', forn:'FIZ/Tecnical', p:18.0},
  {std:'Stucchi', fig:'stucchi', mis:'Stucchi ISO-B 1/2"M FF', cod:'S-ISO-B-12FF', desc:'Innesto rapido Stucchi ISO-B DN12 1/2"M flat-face', forn:'FIZ/Tecnical', p:22.0},
  {std:'Stucchi', fig:'stucchi', mis:'Stucchi ISO-B 3/4"M FF', cod:'S-ISO-B-34FF', desc:'Innesto rapido Stucchi ISO-B DN20 3/4"M flat-face', forn:'FIZ/Tecnical', p:28.0},
  {std:'Stucchi', fig:'stucchi', mis:'Stucchi ISO-B 3/8"F FF', cod:'S-ISO-B-38FFF',desc:'Innesto rapido Stucchi ISO-B DN10 3/8"F flat-face', forn:'FIZ/Tecnical', p:18.0},
  {std:'Stucchi', fig:'stucchi', mis:'Stucchi ISO-B 1/2"F FF', cod:'S-ISO-B-12FFF',desc:'Innesto rapido Stucchi ISO-B DN12 1/2"F flat-face', forn:'FIZ/Tecnical', p:22.0},
  {std:'Stucchi', fig:'stucchi', mis:'Stucchi poppet DN10 3/8"M', cod:'S-POP-38', desc:'Innesto rapido Stucchi poppet DN10 3/8"M',         forn:'FIZ/Tecnical', p:15.0},
  {std:'Stucchi', fig:'stucchi', mis:'Stucchi poppet DN12 1/2"M', cod:'S-POP-12', desc:'Innesto rapido Stucchi poppet DN12 1/2"M',         forn:'FIZ/Tecnical', p:18.0},
];

let _raccScioltoFiltro='';

function setRaccScioltoFiltro(std,btn){
  _raccScioltoFiltro=std;
  // aggiorna bottoni stile
  document.querySelectorAll('#apanel-racc button[onclick^="setRaccScioltoFiltro"]').forEach(b=>{
    const active=b===btn;
    b.style.borderColor=active?'var(--blue)':'var(--border)';
    b.style.background=active?'var(--blue-lt)':'#fff';
    b.style.color=active?'var(--blue)':'var(--text3)';
  });
  renderRaccSciolto();
}

function renderRaccSciolto(){
  const q=(document.getElementById('racc-q')||{}).value?.toLowerCase()||'';
  const el=document.getElementById('racc-results');if(!el)return;
  const filt=RACC_SCIOLTI.filter(r=>{
    if(_raccScioltoFiltro&&r.std!==_raccScioltoFiltro)return false;
    if(q&&!r.mis.toLowerCase().includes(q)&&!r.desc.toLowerCase().includes(q)&&!r.cod.toLowerCase().includes(q))return false;
    return true;
  });
  if(!filt.length){el.innerHTML='<div class="empty" style="padding:14px">Nessun raccordo trovato</div>';return;}
  const figEmoji={adattatore:'🔗',riduzione:'↕',gomito:'↩',nipplo:'⇔',bonded:'⭕',stucchi:'⚡',default:'🔩'};
  // Raggruppa per standard
  const groups={};
  filt.forEach(r=>{if(!groups[r.std])groups[r.std]=[];groups[r.std].push(r);});
  el.innerHTML=Object.entries(groups).map(([std,items])=>'<div style="background:var(--bg2);padding:5px 10px;font-size:10px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.07em;border-bottom:1px solid var(--border);position:sticky;top:0;">'+std+'</div>'+
    items.map(r=>'<div style="display:flex;align-items:center;gap:8px;padding:7px 10px;border-bottom:1px solid var(--border);font-size:12px;">'+
      '<span style="font-size:14px;flex-shrink:0">'+(figEmoji[r.fig]||figEmoji.default)+'</span>'+
      '<div style="flex:1;"><div style="font-weight:600;color:var(--text)">'+r.mis+'</div>'+
        '<div style="font-size:11px;color:var(--text3)">'+r.cod+' · '+r.forn+'</div></div>'+
      '<span style="font-weight:700;white-space:nowrap;color:var(--text)">€'+fmt2(r.p)+'</span>'+
      '<button style="background:var(--blue);border:none;color:#fff;font-size:11px;font-weight:700;padding:4px 10px;border-radius:20px;cursor:pointer;flex-shrink:0;" onclick="usaRaccordoSciolto('+JSON.stringify(r).replace(/"/g,'&quot;')+')">+ usa</button>'+
    '</div>').join('')
  ).join('');
}

function usaRaccordoSciolto(r){
  const netto=r.p;const markup=2.0;
  const row={id:uid(),sezione:'Raccorderia',descrizione:r.desc,qty:1,fornitore:r.forn==='Pa.Co'?'Pa.Co Fluid':'FIZ / Tecnical Impianti',codice:r.cod,costo_cad:netto,sconto:0,netto_cad:netto,markup,prezzo_vendita:parseFloat((netto*markup).toFixed(4)),stato_approv:'daordinare',data_attesa:''};
  const numero=currentAddRigaNumero;if(!numero){showToast('Apri una commessa per aggiungere componenti');return;}
  const det=ll('det');if(!det[numero])det[numero]={distinta:[],ore:[],note:''};
  det[numero].distinta.push(row);ls('det',det);syncDetBackground();CC._det=det[numero];
  closeM();['distinta','approv','riepilogo'].forEach(id=>renderTabContent(id,CC,det[numero]));
  showToast('✓ '+r.mis+' aggiunto alla distinta');
}

// ════════════════════════════════════════════
// CONFIGURATORE PU10 (Ares / Tecfluid)
// ════════════════════════════════════════════
let _pu10Blocchi=[];

function pu10AddBlocco(sel){
  if(!sel.value)return;
  const parts=sel.value.split('|');
  _pu10Blocchi.push({cod:parts[0],prezzo:parseFloat(parts[1]||0),desc:parts[2]||''});
  sel.value='';
  // Render lista blocchi
  const el=document.getElementById('pu10-blocchi-list');if(!el)return;
  el.innerHTML=_pu10Blocchi.map((b,i)=>'<div style="display:flex;align-items:center;gap:8px;background:var(--bg2);border:1px solid var(--border);border-radius:6px;padding:5px 10px;font-size:12px;">'+
    '<span style="font-family:monospace;font-weight:700;color:var(--green);min-width:50px">'+b.cod+'</span>'+
    '<span style="flex:1;color:var(--text2)">'+b.desc+'</span>'+
    '<span style="font-weight:700">€'+fmt2(b.prezzo)+'</span>'+
    '<button onclick="_pu10Blocchi.splice('+i+',1);document.getElementById(\'pu10-blocchi-list\').innerHTML=\'\';_pu10Blocchi.forEach((_,j)=>_);pu10AddBlocco({value:\'\'});pu10Update();" style="background:none;border:none;color:var(--red);cursor:pointer;font-size:14px;padding:0 3px;">✕</button>'+
    '</div>').join('');
  pu10Update();
}

function pu10Update(){
  const getVal=id=>{const el=document.getElementById(id);if(!el)return'';if(el.classList&&el.classList.contains('csel'))return window._cselData?.[id]?.val||el.dataset.value||'';return el.value||'';};
  const collV=getVal('pu10-coll');const pompaV=getVal('pu10-pompa');
  const motV=getVal('pu10-mot-tipo');const giunto=getVal('pu10-giunto');
  const serbV=getVal('pu10-serb');
  const prev=document.getElementById('pu10-preview');if(!prev)return;
  if(!collV||!pompaV||!motV){prev.innerHTML='<span style="color:var(--text4)">Seleziona almeno collettore, pompa e motore per vedere il riepilogo…</span>';return;}
  const pC=collV.split('|');const pP=pompaV.split('|');const pM=motV.split('|');
  const pG=giunto.split('|');const pS=serbV.split('|');
  const prColl=parseFloat(pC[1]||0),prPompa=parseFloat(pP[1]||0),prMot=parseFloat(pM[1]||0);
  const prGiunto=giunto?parseFloat(pG[1]||0):0,prSerb=serbV?parseFloat(pS[1]||0):0;
  const prBlocchi=_pu10Blocchi.reduce((s,b)=>s+b.prezzo,0);
  const totale=prColl+prPompa+prMot+prGiunto+prSerb+prBlocchi;
  // Codice ordinazione semplificato
  const cod='PU10-'+(pC[0]||'?')+'-'+(pP[0]||'?')+'-'+(pM[0]||'?')+(giunto?'-'+pG[0]:'')+(serbV?'-'+pS[0]:'')+_pu10Blocchi.map(b=>'-'+b.cod).join('');
  prev.innerHTML=
    '<div style="font-weight:700;color:var(--green);font-size:13px;margin-bottom:6px">Centralina PU10 — '+(pC[2]||'')+'</div>'+
    '<div style="display:grid;grid-template-columns:auto 1fr auto;gap:2px 10px;font-size:12px;margin-bottom:8px;">'+
    '<span style="color:var(--text3)">Collettore</span><span>'+(pC[2]||'—')+'</span><span style="text-align:right">€'+fmt2(prColl)+'</span>'+
    '<span style="color:var(--text3)">Pompa</span><span>'+(pP[2]||'—')+'</span><span style="text-align:right">€'+fmt2(prPompa)+'</span>'+
    '<span style="color:var(--text3)">Motore</span><span>'+(pM[2]||'—')+'</span><span style="text-align:right">€'+fmt2(prMot)+'</span>'+
    (giunto?'<span style="color:var(--text3)">Giunto</span><span>'+(pG[2]||'—')+'</span><span style="text-align:right">€'+fmt2(prGiunto)+'</span>':'')+
    (serbV?'<span style="color:var(--text3)">Serbatoio</span><span>'+(pS[2]||'—')+'</span><span style="text-align:right">€'+fmt2(prSerb)+'</span>':'')+
    _pu10Blocchi.map(b=>'<span style="color:var(--text3)">'+b.cod+'</span><span>'+b.desc+'</span><span style="text-align:right">€'+fmt2(b.prezzo)+'</span>').join('')+
    '</div>'+
    '<div style="border-top:1px solid #6ee7a0;padding-top:6px;display:flex;justify-content:space-between;align-items:center;">'+
    '<span style="font-size:11px;font-family:monospace;color:var(--text3)">'+cod+'</span>'+
    '<span style="font-weight:800;font-size:15px;color:var(--green)">TOT. €'+fmt2(totale)+'</span>'+
    '</div>'+
    '<div style="font-size:10px;color:var(--text4);margin-top:3px;">Fornitore: Ares / Tecfluid · Prezzi listino 2022 — da confermare</div>';
  prev.dataset.cod=cod;prev.dataset.tot=totale;
  prev.dataset.desc='Centralina PU10: '+(pC[2]||'')+' · Pompa '+pP[0]+' · Motore '+pM[0]+(giunto?' · '+pG[0]:'')+(serbV?' · '+pS[0]:'')+(  _pu10Blocchi.length?' · Blocchi: '+_pu10Blocchi.map(b=>b.cod).join('+'):'');
}

function usaCentralinaPU10(numero){
  const prev=document.getElementById('pu10-preview');
  if(!prev||!prev.dataset.cod||!prev.dataset.cod.includes('PU10-')){showToast('Completa la configurazione prima');return;}
  const totale=parseFloat(prev.dataset.tot||0);const markup=2.0;
  const row={id:uid(),sezione:'Centralina',descrizione:prev.dataset.desc,qty:1,fornitore:'Ares / Tecfluid',codice:prev.dataset.cod,costo_cad:totale,sconto:0,netto_cad:totale,markup,prezzo_vendita:parseFloat((totale*markup).toFixed(4)),stato_approv:'daordinare',data_attesa:''};
  const det=ll('det');if(!det[numero])det[numero]={distinta:[],ore:[],note:''};
  det[numero].distinta.push(row);ls('det',det);syncDetBackground();CC._det=det[numero];
  closeM();['distinta','approv','riepilogo'].forEach(id=>renderTabContent(id,CC,det[numero]));
  showToast('✓ Centralina PU10 aggiunta alla distinta');
}
let _pu20Blocchi=[];

function pu20AddBlocco(sel){
  if(!sel.value)return;
  const p=sel.value.split('|');
  _pu20Blocchi.push({cod:p[0],prezzo:parseFloat(p[1]||0),desc:p[2]||p[0]});
  sel.value='';
  const list=document.getElementById('pu20-blocchi-list');
  if(list){
    const div=document.createElement('div');
    div.style.cssText='display:flex;align-items:center;gap:6px;font-size:12px;background:var(--bg2);padding:4px 8px;border-radius:5px;';
    const idx=_pu20Blocchi.length-1;
    div.innerHTML='<span style="flex:1">'+p[0]+' — '+(p[2]||p[0])+'</span><span style="color:var(--text3)">€'+fmt2(parseFloat(p[1]||0))+'</span><button class="xbtn" onclick="this.parentNode.remove();_pu20Blocchi.splice('+idx+',1);pu20Update()">✕</button>';
    list.appendChild(div);
  }
  pu20Update();
}

function pu20Update(){
  const getVal=id=>{const el=document.getElementById(id);if(!el)return'';if(el.classList&&el.classList.contains('csel'))return window._cselData?.[id]?.val||el.dataset.value||'';return el.value||'';};
  const collV=getVal('pu20-coll');const pompaV=getVal('pu20-pompa');
  const motV=getVal('pu20-mot-tipo');const giunto=getVal('pu20-giunto');
  const serbV=getVal('pu20-serb');
  const prev=document.getElementById('pu20-preview');if(!prev)return;
  if(!collV||!pompaV||!motV){
    prev.innerHTML='<span style="color:var(--text4)">Seleziona almeno collettore, pompa e motore per vedere il riepilogo…</span>';
    return;
  }
  const pC=collV.split('|');const pP=pompaV.split('|');
  const pM=motV.split('|');const pG=giunto.split('|');const pS=serbV.split('|');
  const prColl=parseFloat(pC[1]||0),prPompa=parseFloat(pP[1]||0),prMot=parseFloat(pM[1]||0);
  const prGiunto=giunto?parseFloat(pG[1]||0):0;
  const prSerb=serbV?parseFloat(pS[1]||0):0;
  const prAcc=_pu20Blocchi.reduce((s,b)=>s+b.prezzo,0);
  const totale=prColl+prPompa+prMot+prGiunto+prSerb+prAcc;
  const cod='PU20-'+(pC[0]||'?')+'-'+(pP[0]||'?')+'-'+(pM[0]||'?')+(giunto?'-'+pG[0]:'')+(serbV?'-'+pS[0]:'')+_pu20Blocchi.map(b=>'-'+b.cod).join('');
  prev.innerHTML=
    '<div style="font-weight:700;color:var(--green);font-size:13px;margin-bottom:6px">Centralina PU20 — '+(pC[2]||'')+'</div>'+
    '<div style="display:grid;grid-template-columns:auto 1fr auto;gap:2px 10px;font-size:12px;margin-bottom:8px;">'+
    '<span style="color:var(--text3)">Collettore</span><span>'+(pC[2]||'—')+'</span><span style="text-align:right">€'+fmt2(prColl)+'</span>'+
    '<span style="color:var(--text3)">Pompa</span><span>'+(pP[2]||'—')+'</span><span style="text-align:right">€'+fmt2(prPompa)+'</span>'+
    '<span style="color:var(--text3)">Motore</span><span>'+(pM[2]||'—')+'</span><span style="text-align:right">€'+fmt2(prMot)+'</span>'+
    (giunto?'<span style="color:var(--text3)">Giunto</span><span>'+(pG[2]||'—')+'</span><span style="text-align:right">€'+fmt2(prGiunto)+'</span>':'')+
    (serbV?'<span style="color:var(--text3)">Serbatoio</span><span>'+(pS[2]||'—')+'</span><span style="text-align:right">€'+fmt2(prSerb)+'</span>':'')+
    _pu20Blocchi.map(b=>'<span style="color:var(--text3)">'+b.cod+'</span><span>'+b.desc+'</span><span style="text-align:right">€'+fmt2(b.prezzo)+'</span>').join('')+
    '</div>'+
    '<div style="border-top:1px solid #6ee7a0;padding-top:6px;display:flex;justify-content:space-between;align-items:center;">'+
    '<span style="font-size:11px;font-family:monospace;color:var(--text3)">'+cod+'</span>'+
    '<span style="font-weight:800;font-size:15px;color:var(--green)">TOT. €'+fmt2(totale)+'</span>'+
    '</div>'+
    '<div style="font-size:10px;color:var(--text4);margin-top:3px;">Fornitore: Ares / Tecfluid · Prezzi listino 2022 — da confermare</div>';
  prev.dataset.cod=cod;prev.dataset.tot=totale;
  prev.dataset.desc='Centralina PU20: '+(pC[2]||'')+' · Pompa '+pP[0]+' · Motore '+pM[0]+(giunto?' · '+pG[0]:'')+(serbV?' · '+pS[0]:'')+(  _pu20Blocchi.length?' · Accessori: '+_pu20Blocchi.map(b=>b.cod).join('+'):'');
}

function usaCentralinaPU20(numero){
  const prev=document.getElementById('pu20-preview');
  if(!prev||!prev.dataset.cod||!prev.dataset.cod.includes('PU20-')){showToast('Completa la configurazione prima');return;}
  const totale=parseFloat(prev.dataset.tot||0);const markup=2.0;
  const row={id:uid(),sezione:'Centralina',descrizione:prev.dataset.desc,qty:1,
    fornitore:'Ares / Tecfluid',codice:prev.dataset.cod,costo_cad:totale,
    sconto:0,netto_cad:totale,markup,prezzo_vendita:parseFloat((totale*markup).toFixed(4)),
    stato_approv:'daordinare',data_attesa:''};
  const det=ll('det');if(!det[numero])det[numero]={distinta:[],ore:[],note:''};
  det[numero].distinta.push(row);ls('det',det);syncDetBackground();CC._det=det[numero];
  closeM();['distinta','approv','riepilogo'].forEach(id=>renderTabContent(id,CC,det[numero]));
  showToast('✓ Centralina PU20 aggiunta alla distinta');
}

let _tfRaSelected={A:null,B:null};
let _tfStdFiltro='';
let _tfDnAttuale=0;

function onTuboTipoChange(){
  const sel=document.getElementById('tf-tipo');
  if(!sel)return;
  const val=sel.value;
  if(!val){_tfDnAttuale=0;renderRaccordiSection();aggTuboPreview();return;}
  const p=val.split('|');
  _tfDnAttuale=parseInt(p[3]||0);
  _tfRaSelected={A:null,B:null};
  _tfStdFiltro='';
  renderRaccordiSection();
  aggTuboPreview();
}

function syncLung(da){
  const mmEl=document.getElementById('tf-lung-mm');
  const inEl=document.getElementById('tf-lung-in');
  const hint=document.getElementById('tf-lung-hint');
  if(da==='mm'&&mmEl){
    const mm=parseFloat(mmEl.value||0);
    if(inEl&&mm)inEl.value=(mm/25.4).toFixed(1);
    if(hint&&mm)hint.textContent=mm+' mm = '+(mm/25.4).toFixed(1)+'" = '+(mm/10).toFixed(0)+' cm';
  } else if(da==='in'&&inEl){
    const ins=parseFloat(inEl.value||0);
    if(mmEl&&ins)mmEl.value=Math.round(ins*25.4);
    if(hint&&ins)hint.textContent=(ins*25.4).toFixed(0)+' mm = '+ins+'" = '+(ins*2.54).toFixed(1)+' cm';
  }
  aggTuboPreview();
}

function renderRaccordiSection(){
  const el=document.getElementById('tf-racc-section');if(!el)return;
  if(!_tfDnAttuale){
    el.innerHTML='<div style="font-size:12px;color:var(--text3);padding:10px;text-align:center;background:var(--bg2);border-radius:7px;">Seleziona prima il tipo di tubo per filtrare i raccordi compatibili</div>';
    return;
  }
  const compatibili=RACCORDI_DB.filter(r=>r.dn.includes(_tfDnAttuale));
  const lista=_tfStdFiltro?compatibili.filter(r=>r.std===_tfStdFiltro):compatibili;
  const stds=[...new Set(compatibili.map(r=>r.std))];

  // Bottoni filtro standard
  const btnTutti='<button onclick="setRaccFiltro('')" style="font-size:11px;font-weight:700;padding:3px 10px;border-radius:20px;border:1.5px solid '+(!_tfStdFiltro?'var(--blue)':'var(--border)')+'";background:'+(!_tfStdFiltro?'var(--blue-lt)':'#fff')+'";color:'+(!_tfStdFiltro?'var(--blue)':'var(--text3)')+';cursor:pointer;">Tutti</button>';
  const btnStds=stds.map(s=>'<button onclick="setRaccFiltro(''+s+'')" style="font-size:11px;font-weight:700;padding:3px 10px;border-radius:20px;border:1.5px solid '+(_tfStdFiltro===s?'var(--blue)':'var(--border)')+'";background:'+(_tfStdFiltro===s?'var(--blue-lt)':'#fff')+'";color:'+(_tfStdFiltro===s?'var(--blue)':'var(--text3)')+';cursor:pointer;">'+s+'</button>').join('');

  // Colonne A e B
  const colHtml=['A','B'].map(lato=>{
    const selInfo=_tfRaSelected[lato];
    const selBadge=selInfo?'<span style="background:var(--green-lt);color:var(--green);border:1px solid var(--green-bd);padding:1px 7px;border-radius:20px;font-size:10px;margin-left:4px;">✓ '+selInfo.label.split(' ').slice(0,2).join(' ')+' '+( selInfo.curva||'')+'</span>':'';
    const rows=lista.map(r=>{
      const varianti=[{k:'dritto',label:'↔ dritto'},...(!r.noCurva?[{k:'45°',label:'↗ curva 45°'},{k:'90°',label:'↕ curva 90°'}]:[])];
      return varianti.map(v=>{
        const isSel=selInfo&&selInfo.raccId===r.std+'_'+r.mis&&selInfo.curva===v.k;
        const bg=isSel?'var(--blue-lt)':'transparent';
        const col=isSel?'var(--blue)':'var(--text2)';
        const colCode=isSel?'var(--blue)':'var(--text3)';
        const colV=isSel?'var(--blue)':'var(--text4)';
        return '<div onclick="selRacc(''+lato+'',''+r.std+'',''+r.mis.replace(/"/g,'&quot;')+'',''+r.label.replace(/"/g,'&quot;')+'',''+v.k+'')" style="padding:6px 10px;cursor:pointer;border-bottom:1px solid var(--border);font-size:12px;display:flex;align-items:center;justify-content:space-between;background:'+bg+';color:'+col+';transition:background .1s;">'+
          '<span><span style="font-family:monospace;font-weight:700;font-size:11px;color:'+colCode+'">'+r.std+'</span> '+r.mis+' <span style="font-size:11px;color:'+colV+'">'+v.label+'</span></span>'+
          (isSel?'<span style="font-size:14px;">✓</span>':'')+
          '</div>';
      }).join('');
    }).join('');
    return '<div>'+
      '<div style="font-size:11px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.06em;margin-bottom:5px;">Raccordo lato '+lato+selBadge+'</div>'+
      '<div style="max-height:190px;overflow-y:auto;border:1.5px solid var(--border);border-radius:8px;background:#fff;">'+rows+'</div>'+
      '</div>';
  }).join('');

  el.innerHTML=
    '<div style="margin-bottom:8px;">'+
      '<div style="font-size:11px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.06em;margin-bottom:5px;">Raccordi compatibili con DN'+_tfDnAttuale+' — filtra per standard:</div>'+
      '<div style="display:flex;gap:5px;flex-wrap:wrap;margin-bottom:10px;">'+btnTutti+btnStds+'</div>'+
    '</div>'+
    '<div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;">'+colHtml+'</div>';
}

function setRaccFiltro(std){
  _tfStdFiltro=std;
  renderRaccordiSection();
}

function selRacc(lato,std,mis,label,curva){
  _tfRaSelected[lato]={raccId:std+'_'+mis,std,mis,label,curva};
  renderRaccordiSection();
  aggTuboPreview();
}

function aggTuboPreview(){
  const tipo=(document.getElementById('tf-tipo')||{}).value||'';
  const lung=parseInt((document.getElementById('tf-lung-mm')||{}).value||0);
  const ra=_tfRaSelected.A;
  const rb=_tfRaSelected.B;
  const qty=parseInt((document.getElementById('tf-qty')||{}).value||1);
  const prev=document.getElementById('tf-preview');if(!prev)return;

  const doubleCurva=ra&&rb&&ra.curva!=='dritto'&&rb.curva!=='dritto';
  const orientRow=document.getElementById('tf-orient-row');
  if(orientRow)orientRow.style.display=doubleCurva?'':'none';
  const orient=doubleCurva?(document.getElementById('tf-orient')||{}).value||'0':'';

  if(!tipo||!lung||!ra||!rb){
    prev.innerHTML='<span style="color:var(--text4)">Compila tutti i campi e seleziona entrambi i raccordi per vedere la descrizione…</span>';
    return;
  }

  const p=tipo.split('|');
  const codT=p[0]||'';const normaT=p[1]||'';const pressT=p[2]||'';const dn=parseInt(p[3]||0);const pollici=p[4]||'';

  const lungIn=(lung/25.4).toFixed(1);
  const raStr=ra.label+' '+(ra.curva==='dritto'?'↔':'curva '+ra.curva);
  const rbStr=rb.label+' '+(rb.curva==='dritto'?'↔':'curva '+rb.curva);

  let desc='Tubo flessibile raccordato \u2014 '+normaT+' DN'+dn+' ('+pollici+') \u00b7 '+lung+' mm ('+lungIn+'")\n  Lato A: '+raStr+'\n  Lato B: '+rbStr;
  if(doubleCurva)desc+='\n  Orientamento doppia curva: '+orient+'°';

  const raShort=ra.std+ra.mis.replace(/[^0-9/]/g,'').replace('/','');
  const rbShort=rb.std+rb.mis.replace(/[^0-9/]/g,'').replace('/','');
  const orientSuffix=doubleCurva?'-OR'+orient:'';
  const codice='TF-'+codT+'-DN'+dn+'-'+lung+'mm-'+raShort+'-'+(ra.curva!=='dritto'?ra.curva.replace('°','d'):'')+'-'+rbShort+'-'+(rb.curva!=='dritto'?rb.curva.replace('°','d'):'')+orientSuffix;

  const costoTubo=lung/1000*3.8;
  const costoRa=ra.curva!=='dritto'?6.2:4.5;
  const costoRb=rb.curva!=='dritto'?6.2:4.5;
  const prezzoStimato=costoTubo+costoRa+costoRb;

  prev.innerHTML=
    '<div style="font-weight:700;color:var(--text);white-space:pre-line;margin-bottom:6px;">'+desc+'</div>'+
    '<div style="font-size:12px;color:var(--text3)">Norma: '+normaT+' · Pressione max: '+pressT+' bar · Fornitore: Pa.Co Fluid</div>'+
    '<div style="font-size:12px;color:var(--text3);margin-top:3px;font-family:monospace">'+codice+'</div>'+
    '<div style="font-size:12px;color:var(--text2);margin-top:4px">Prezzo stimato: <strong>\u20ac'+fmt2(prezzoStimato)+'</strong> cad \u2014 <em>da confermare con Pa.Co Fluid</em></div>'+
    (doubleCurva?'<div style="font-size:11px;background:#fef3c7;border:1px solid #f59e0b;border-radius:5px;padding:4px 8px;margin-top:6px;color:#78350f;">\u26a0 Doppia curva orientamento '+orient+'\u00b0 \u2014 comunicare esplicitamente a Pa.Co nell\'ordine.</div>':'');

  prev.dataset.desc=desc.replace(/\n/g,' | ');
  prev.dataset.cod=codice;prev.dataset.prezzo=prezzoStimato;prev.dataset.qty=qty;
}

function usaTuboFlessibile(numero){
  const prev=document.getElementById('tf-preview');if(!prev||!prev.dataset.desc){showToast('Completa prima il configuratore');return;}
  const netto=parseFloat(prev.dataset.prezzo||8);const markup=2.0;
  const row={id:uid(),sezione:'Raccorderia',descrizione:prev.dataset.desc,qty:parseInt(prev.dataset.qty||1),fornitore:'Pa.Co Fluid',codice:prev.dataset.cod,costo_cad:netto,sconto:0,netto_cad:netto,markup,prezzo_vendita:parseFloat((netto*markup).toFixed(4)),stato_approv:'daordinare',data_attesa:''};
  const det=ll('det');if(!det[numero])det[numero]={distinta:[],ore:[],note:''};
  det[numero].distinta.push(row);ls('det',det);syncDetBackground();CC._det=det[numero];
  closeM();['distinta','approv','riepilogo'].forEach(id=>renderTabContent(id,CC,det[numero]));
  showToast('✓ Tubo flessibile aggiunto alla distinta');
}

function aresFilter(){
  const q=(document.getElementById('ares-q')||{}).value||'';
  const serie=(document.getElementById('ares-serie')||{}).value||'';
  const res=aresSearch(q,serie);
  const el=document.getElementById('ares-results');
  if(!el)return;
  if(!res.length){el.innerHTML=q.length>1?'<div style="padding:6px;font-size:13px;color:var(--text3)">Nessun risultato</div>':'';return;}
  el.innerHTML=res.map(a=>{
    const st=TIPO_BADGE[a.t]||'background:#f3f4f6;color:#374151';
    return '<div style="display:flex;align-items:center;gap:8px;padding:7px 4px;border-bottom:1px solid var(--border);font-size:13px;flex-wrap:wrap;">'+
      '<span style="font-family:monospace;font-weight:700;color:var(--green);min-width:68px;font-size:12px">'+a.s+'</span>'+
      '<span style="font-size:10px;font-weight:700;padding:2px 6px;border-radius:20px;'+st+';white-space:nowrap">'+(a.serie||a.t)+'</span>'+
      '<span style="flex:1;color:var(--text2);min-width:120px">'+a.d+'</span>'+
      '<span style="font-weight:700;white-space:nowrap">€'+fmt2(a.p)+'</span>'+
      '<button style="background:var(--green);border:none;color:#fff;font-size:11px;font-weight:700;padding:3px 10px;border-radius:20px;cursor:pointer;" onclick="fillAres('+JSON.stringify(a).replace(/"/g,'&quot;')+')">usa</button>'+
    '</div>';
  }).join('');
}
function fillAres(a){
  document.getElementById('f-desc').value=a.d;
  document.getElementById('f-forn').value='Ares / Tecfluid';
  document.getElementById('f-cod').value=a.s;
  document.getElementById('f-costo').value=a.p;
  document.getElementById('ares-results').innerHTML='';
  document.getElementById('ares-q').value='';
}

// ═══════════════════════════════════════════════
// SAVE
// ═══════════════════════════════════════════════
async function saveComm(){
  const numero=(document.getElementById('f-num')||{}).value?.trim();if(!numero)return;
  const getF=id=>{const el=document.getElementById(id);if(!el)return'';if(el.classList?.contains('csel'))return window._cselData?.[id]?.val||'';return el.value||'';};
  const row={id:uid(),numero,cliente:getCod(numero),oggetto:getF('f-ogg'),tipo_commessa:getF('f-tipo'),importo:getF('f-imp'),data_apertura:getF('f-dat'),data_consegna:getF('f-cons'),stato:getF('f-stato'),note:getF('f-note')};
  const data=lll('comm');data.unshift(row);ls('comm',data);
  const det=ll('det');if(!det[numero])det[numero]={distinta:[],ore:[],note:''};ls('det',det);
  await wPut('commesse.csv',[row,...(SD.comm||[])],['numero','cliente','oggetto','tipo_commessa','importo','data_apertura','data_consegna','stato','note']);
  // Crea cartella Nextcloud in background
  const cfg=getCfg();
  if(cfg.ok){
    createCommFolder(numero, row.oggetto).then(path=>{
      if(path) showToast('✓ Commessa salvata · Cartella NC creata');
      else showToast('✓ Commessa salvata (cartella NC non creata)');
    });
  } else {
    showToast('✓ Commessa salvata');
  }
  closeM();await syncNow();
}

async function updateComm(){
  const c=CC;if(!c)return;
  const updated={...c,oggetto:document.getElementById('f-ogg').value.trim(),tipo_commessa:document.getElementById('f-tipo').value,importo:document.getElementById('f-imp').value,data_apertura:document.getElementById('f-dat').value,data_consegna:document.getElementById('f-cons').value,stato:document.getElementById('f-stato').value,note:document.getElementById('f-note').value.trim()};
  const all=lll('comm');const idx=all.findIndex(x=>x.numero===c.numero);
  if(idx>=0)all[idx]=updated;else all.unshift(updated);ls('comm',all);
  if(SD.comm){const i=SD.comm.findIndex(x=>x.numero===c.numero);if(i>=0)SD.comm[i]=updated;}
  CC={...updated,_det:c._det};
  await wPut('commesse.csv',all,['numero','cliente','oggetto','tipo_commessa','importo','data_apertura','data_consegna','stato','note']);
  closeM();
  // aggiorna header senza ricaricare tutto
  document.getElementById('det-title').textContent=updated.oggetto||'—';
  const tipo=updated.tipo_commessa||'';
  document.getElementById('det-meta').innerHTML='<span class="bdg s-'+(updated.stato||'offerta').replace(' ','-')+'">'+(SL[updated.stato]||'\u2014')+'</span>'+(tipo?'<span class="bdg '+(TIPO_C[tipo]||'')+'">'+(TIPI[tipo]||tipo)+'</span>':'')+'<span style="font-size:12px;color:var(--text2)">'+(CNAMES[getCod(c.numero)]||c.cliente||'\u2014')+'</span>';
  buildTabs(tipo,updated,CC._det);
}

function saveRiga(numero){
  const costo=parseFloat(document.getElementById('f-costo').value||0);
  const sconto=parseFloat(document.getElementById('f-sconto').value||0);
  const markup=parseFloat(document.getElementById('f-markup').value||2.25);
  const netto=costo*(1-sconto/100);
  const row={id:uid(),sezione:document.getElementById('f-sez').value.trim()||'Generale',descrizione:document.getElementById('f-desc').value.trim(),qty:parseInt(document.getElementById('f-qty').value||1),fornitore:document.getElementById('f-forn').value.trim(),codice:document.getElementById('f-cod').value.trim(),costo_cad:costo,sconto,netto_cad:parseFloat(netto.toFixed(4)),markup,prezzo_vendita:parseFloat((netto*markup).toFixed(4)),stato_approv:document.getElementById('f-stato-a').value,data_attesa:document.getElementById('f-datta').value};
  const det=ll('det');if(!det[numero])det[numero]={distinta:[],ore:[],note:''};
  det[numero].distinta.push(row);ls('det',det);syncDetBackground();CC._det=det[numero];
  closeM();['distinta','approv','riepilogo'].forEach(id=>renderTabContent(id,CC,det[numero]));
}

function updateRiga(numero,rid){
  const costo=parseFloat(document.getElementById('f-costo').value||0);
  const sconto=parseFloat(document.getElementById('f-sconto').value||0);
  const markup=parseFloat(document.getElementById('f-markup').value||2.25);
  const netto=costo*(1-sconto/100);
  const det=ll('det');if(!det[numero])return;
  const idx=det[numero].distinta.findIndex(r=>r.id===rid);if(idx<0)return;
  det[numero].distinta[idx]={
    ...det[numero].distinta[idx],
    sezione:document.getElementById('f-sez').value.trim()||det[numero].distinta[idx].sezione,
    descrizione:document.getElementById('f-desc').value.trim(),
    qty:parseInt(document.getElementById('f-qty').value||1),
    fornitore:document.getElementById('f-forn').value.trim(),
    codice:document.getElementById('f-cod').value.trim(),
    costo_cad:costo,sconto,
    netto_cad:parseFloat(netto.toFixed(4)),
    markup,
    prezzo_vendita:parseFloat((netto*markup).toFixed(4)),
    stato_approv:document.getElementById('f-stato-a').value,
    data_attesa:document.getElementById('f-datta').value
  };
  ls('det',det);syncDetBackground();CC._det=det[numero];
  closeM();['distinta','approv','riepilogo'].forEach(id=>renderTabContent(id,CC,det[numero]));
  showToast('✓ Componente aggiornato');
}

function saveSezione(numero){
  const nome=(document.getElementById('f-sez-nome').value||'').trim();
  if(!nome){alert('Inserisci un nome per la sezione');return;}
  // Aggiunge una riga-intestazione sezione nella distinta
  const row={id:uid(),sezione:nome,_intestazione:true,descrizione:'',qty:0,fornitore:'',codice:'',costo_cad:0,sconto:0,netto_cad:0,markup:1,prezzo_vendita:0,stato_approv:'',data_attesa:''};
  const det=ll('det');if(!det[numero])det[numero]={distinta:[],ore:[],note:''};
  det[numero].distinta.push(row);ls('det',det);syncDetBackground();CC._det=det[numero];
  closeM();['distinta','approv','riepilogo'].forEach(id=>renderTabContent(id,CC,det[numero]));
}

function saveOre(numero){
  const row={id:uid(),data:document.getElementById('f-dat').value,operatore:document.getElementById('f-op').value.trim(),descrizione:document.getElementById('f-desc').value.trim(),ore:parseFloat(document.getElementById('f-ore').value||0),tariffa:parseFloat(document.getElementById('f-tar').value||0)};
  const det=ll('det');if(!det[numero])det[numero]={distinta:[],ore:[],note:''};
  (det[numero].ore=det[numero].ore||[]).push(row);ls('det',det);syncDetBackground();CC._det=det[numero];
  closeM();renderTabContent('ore',CC,det[numero]);renderTabContent('riepilogo',CC,det[numero]);
}

// ════════════════════════════════════════════
// IMPORT MOVIMENTI BANCARI
// ════════════════════════════════════════════

function dropMovimenti(ev, banca){
  ev.preventDefault();
  const file=ev.dataTransfer.files[0];
  if(file)parseMovimenti(file,banca);
}
function fileMovimenti(input,banca){
  if(input.files[0])parseMovimenti(input.files[0],banca);
}

async function parseMovimenti(file, banca){
  const resEl=document.getElementById('res-'+banca);
  if(resEl)resEl.innerHTML='<span style="color:var(--text3)">⏳ Lettura in corso…</span>';
  try{
    let movimenti=[];
    let nomeConto='';
    let saldoFinale=null;
    let dataUltimo=null;

    if(banca==='fineco'){
      ({movimenti,nomeConto,saldoFinale,dataUltimo}=await parseFinecoXLSX(file));
    } else if(banca==='sella'){
      ({movimenti,nomeConto,saldoFinale,dataUltimo}=await parseSellaCSV(file));
    }

    if(!movimenti.length){
      if(resEl)resEl.innerHTML='<span style="color:var(--amber)">⚠ Nessun movimento trovato nel file.</span>';
      return;
    }

    // Salva movimenti con anti-duplicati
    const existing=lll('movimenti_'+banca);
    const keys=new Set(existing.map(m=>m._key||''));
    const nuovi=movimenti.filter(m=>!keys.has(m._key));
    const merged=[...nuovi,...existing].sort((a,b)=>b.data.localeCompare(a.data));
    ls('movimenti_'+banca, merged);

    // Aggiorna saldo conto automaticamente
    if(saldoFinale!==null && dataUltimo){
      const contiAll=lll('conti');
      const nomeContoNorm=nomeConto||({fineco:'Banca Fineco',sella:'Banca Sella'}[banca]||banca);
      const idx=contiAll.findIndex(c=>(c.conto||'').toLowerCase().includes(banca.toLowerCase()));
      const row={id:idx>=0?contiAll[idx].id:uid(),conto:nomeContoNorm,saldo:saldoFinale.toFixed(2),data_rilevazione:dataUltimo,note:'Import automatico da '+file.name};
      if(idx>=0)contiAll[idx]=row; else contiAll.push(row);
      ls('conti',contiAll);
      await wPut('cassa.csv',contiAll,['conto','saldo','data_rilevazione','note']);
      renderCassa();
    }

    const msg='✓ '+nuovi.length+' nuovi mov. su '+movimenti.length+' totali'+(saldoFinale!==null?' · saldo aggiornato: €'+fmt(saldoFinale):'');
    if(resEl)resEl.innerHTML='<span style="color:var(--green);font-weight:600">'+msg+'</span>';
    showToast(msg);

  }catch(err){
    if(resEl)resEl.innerHTML='<span style="color:var(--red)">✗ '+err.message+'</span>';
    console.error('Import movimenti error:',err);
  }
}

async function parseFinecoXLSX(file){
  return new Promise((resolve,reject)=>{
    const reader=new FileReader();
    reader.onload=async e=>{
      try{
        if(typeof XLSX==='undefined'){
          await new Promise((res,rej)=>{
            const s=document.createElement('script');
            s.src='https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
            s.onload=res;s.onerror=()=>rej(new Error('Impossibile caricare SheetJS'));
            document.head.appendChild(s);
          });
        }
        const wb=XLSX.read(new Uint8Array(e.target.result),{type:'array',cellDates:true});
        const ws=wb.Sheets[wb.SheetNames[0]];
        const aoa=XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:null});

        // Estrai intestazione: R0=conto, R2=periodo, R4=saldo
        let nomeConto='Banca Fineco';
        let saldoFinale=null;
        for(let i=0;i<10;i++){
          const r=aoa[i];if(!r)continue;
          const v=String(r[0]||'');
          if(v.includes('Conto Corrente:'))nomeConto='Banca Fineco ('+v.split(':')[1].trim()+')';
          if(v.includes('Saldo Finale:')){
            const m=v.match(/Saldo Finale:\s*([\d.,]+)/);
            if(m)saldoFinale=parseFloat(m[1].replace(/\./g,'').replace(',','.'));
          }
        }

        // Trova riga header (Data_Operazione)
        let hdrIdx=-1;
        for(let i=0;i<aoa.length;i++){
          if(aoa[i]&&String(aoa[i][0]||'').includes('Data_Operazione')){hdrIdx=i;break;}
        }
        if(hdrIdx<0)throw new Error('Header "Data_Operazione" non trovato — controlla il file Fineco.');

        const movimenti=[];
        let dataUltimo=null;

        for(let i=hdrIdx+1;i<aoa.length;i++){
          const r=aoa[i];if(!r||!r[0])continue;
          // Data: può essere Date object o stringa
          let data='';
          if(r[0] instanceof Date){
            data=r[0].toISOString().slice(0,10);
          } else if(typeof r[0]==='number'){
            // Excel serial date
            const d=XLSX.SSF.parse_date_code(r[0]);
            data=`${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`;
          } else {
            const s=String(r[0]);
            const m=s.match(/(\d{2})\/(\d{2})\/(\d{4})/);
            if(m)data=`${m[3]}-${m[2].padStart(2,"0")}-${m[1].padStart(2,"0")}`;
          }
          if(!data)continue;

          const entrate=parseFloat(r[2])||0;
          const uscite=parseFloat(r[3])||0; // già negativo
          const importo=entrate>0?entrate:uscite; // uscite < 0
          const desc=String(r[4]||'').trim();
          const descCompl=String(r[5]||'').trim();
          const key=`fineco_${data}_${importo}_${desc.slice(0,20)}`;

          movimenti.push({
            banca:'fineco',data,
            entrata:entrate>0?entrate:0,
            uscita:uscite<0?Math.abs(uscite):0,
            importo,
            descrizione:desc,
            dettaglio:descCompl,
            stato:String(r[6]||''),
            _key:key
          });
          if(!dataUltimo||data>dataUltimo)dataUltimo=data;
        }

        resolve({movimenti,nomeConto,saldoFinale,dataUltimo});
      }catch(err){reject(err);}
    };
    reader.onerror=()=>reject(new Error('Errore lettura file'));
    reader.readAsArrayBuffer(file);
  });
}

async function parseSellaCSV(file){
  return new Promise((resolve,reject)=>{
    const reader=new FileReader();
    reader.onload=e=>{
      try{
        const text=e.target.result;
        const lines=text.split(/\r?\n/).filter(l=>l.trim());

        // Riga 0 = header
        const headers=lines[0].split(',').map(h=>h.trim().replace(/^"|"$/g,''));

        const iData=headers.findIndex(h=>h.toLowerCase().includes('data operazione'));
        const iDesc=headers.findIndex(h=>h.toLowerCase()==='descrizione');
        const iContr=headers.findIndex(h=>h.toLowerCase()==='controparte');
        const iDeb=headers.findIndex(h=>h.toLowerCase()==='debito');
        const iCred=headers.findIndex(h=>h.toLowerCase()==='credito');
        const iImp=headers.findIndex(h=>h.toLowerCase().includes('importo'));

        // Cerca righe saldo (Descrizione="Saldo al...") per estrarre saldo finale
        let saldoFinale=null;
        let dataUltimo=null;
        const movimenti=[];

        for(let i=1;i<lines.length;i++){
          const cols=lines[i].split(',').map(c=>c.trim().replace(/^"|"$/g,''));
          const desc=cols[iDesc]||'';

          // Riga saldo — usa parseIta (formato italiano con punto migliaia)
          if(desc.startsWith('Saldo al')){
            const raw=(cols[iImp]||'').trim();
            const s=parseFloat(raw.replace(/\./g,'').replace(',','.'));
            if(!isNaN(s))saldoFinale=s;
            continue;
          }

          const dataRaw=cols[iData]||'';
          if(!dataRaw||!dataRaw.match(/\d{2}\/\d{2}\/\d{4}/))continue;
          const dm=dataRaw.match(/(\d{2})\/(\d{2})\/(\d{4})/);
          const data=dm?`${dm[3]}-${dm[2].padStart(2,'0')}-${dm[1].padStart(2,'0')}`:'';
          if(!data)continue;

          // parseIta: converte numero formato italiano (1.234,56) in float
          const parseIta=v=>{
            if(v===null||v===undefined||v==='')return 0;
            const s=String(v).trim();
            if(!s)return 0;
            // Se già numero JS valido (es. da SheetJS) restituisce direttamente
            if(!isNaN(Number(s))&&!s.includes(','))return parseFloat(s)||0;
            // Formato italiano: rimuovi punti migliaia, converti virgola decimale
            return parseFloat(s.replace(/\./g,'').replace(',','.'))||0;
          };

          const deb=parseIta(cols[iDeb]);
          const cred=parseIta(cols[iCred]);
          const imp=parseIta(cols[iImp]);
          const controparte=cols[iContr]||'';
          const key=`sella_${data}_${imp}_${desc.slice(0,20)}`;

          movimenti.push({
            banca:'sella',data,
            entrata:cred>0?cred:0,
            uscita:deb>0?deb:0,
            importo:imp,
            descrizione:desc+(controparte?' — '+controparte:''),
            dettaglio:controparte,
            stato:'',_key:key
          });
          if(!dataUltimo||data>dataUltimo)dataUltimo=data;
        }

        resolve({movimenti,nomeConto:'Banca Sella',saldoFinale,dataUltimo});
      }catch(err){reject(err);}
    };
    reader.onerror=()=>reject(new Error('Errore lettura file'));
    reader.readAsText(file,'utf-8');
  });
}

async function salvaRevolut(){
  const saldo=parseFloat(document.getElementById('rev-saldo')?.value||'');
  const data=document.getElementById('rev-data')?.value||'';
  if(isNaN(saldo)||!data){showToast('Inserisci saldo e data');return;}
  const conti=lll('conti');
  const idx=conti.findIndex(c=>(c.conto||'').toLowerCase().includes('revolut'));
  const row={id:idx>=0?conti[idx].id:uid(),conto:'Revolut',saldo:saldo.toFixed(2),data_rilevazione:data,note:'Aggiornato da estratto conto PDF'};
  if(idx>=0)conti[idx]=row; else conti.push(row);
  ls('conti',conti);
  await wPut('cassa.csv',conti,['conto','saldo','data_rilevazione','note']);
  renderCassa();
  const el=document.getElementById('res-revolut');
  if(el)el.innerHTML=`<span style="color:var(--green);font-weight:600">✓ Saldo Revolut aggiornato: €${fmt(saldo)}</span>`;
  showToast('✓ Saldo Revolut aggiornato');
}

async function saveConto(){
  const nome=(document.getElementById('f-nome')||{}).value?.trim()||'';
  const saldo=(document.getElementById('f-saldo')||{}).value||'0';
  const data_rilevazione=(document.getElementById('f-dat')||{}).value||'';
  const note=(document.getElementById('f-note')||{}).value?.trim()||'';
  if(!nome){showToast('Inserisci il nome del conto');return;}
  const all=lll('conti');
  // Upsert: aggiorna se esiste già un conto con lo stesso nome
  const idx=all.findIndex(c=>(c.conto||c.nome||'').toLowerCase()===nome.toLowerCase());
  const row={id:idx>=0?all[idx].id:uid(),conto:nome,saldo,data_rilevazione,note};
  if(idx>=0)all[idx]=row;else all.push(row);
  ls('conti',all);
  await wPut('cassa.csv',all,['conto','saldo','data_rilevazione','note']);
  closeM();renderCassa();
  showToast('✓ Saldo aggiornato: '+nome);
}

async function saveScad(){
  const row={
    id:uid(),
    data:document.getElementById('f-dat').value,
    descrizione:(document.getElementById('f-desc')||{}).value?.trim()||'',
    importo:document.getElementById('f-imp').value||'0',
    categoria:(document.getElementById('f-cat')||{}).value||'',
    pagato:'no',
    _fiscale:false
  };
  if(!row.data){showToast('Inserisci la data');return;}
  const data=lll('scad');data.push(row);ls('scad',data);
  await wPut('scadenze.csv',[row,...(SD.scad||[])],['data','descrizione','importo','categoria','pagato']);
  closeM();renderScad();
}

function rigeneraScadenzeFatture(){
  const scadenze=lll('scad');
  const nonFatt=scadenze.filter(s=>!s._fatt_key);
  ls('scad',nonFatt);
  const fatture=SD.fatt||lll('fatt');
  if(!fatture.length){showToast('Nessuna fattura importata');return;}
  const n=aggiungiScadenzeDaFatture(fatture);
  showToast(`✓ Rigenerate ${n} scadenze da ${fatture.length} fatture`);
}

function saveScadEdit(sid){
  const all=lll('scad');
  const idx=all.findIndex(x=>x.id===sid);
  if(idx<0){showToast('Scadenza non trovata');return;}
  all[idx]={
    ...all[idx],
    descrizione:(document.getElementById('f-desc')||{}).value?.trim()||all[idx].descrizione,
    data:document.getElementById('f-dat').value||all[idx].data,
    importo:document.getElementById('f-imp').value||'0',
    categoria:(document.getElementById('f-cat')||{}).value||all[idx].categoria||'',
    pagato:(document.getElementById('f-pagata')||{}).value||'no',
  };
  ls('scad',all);
  closeM();renderScad();
  showToast('✓ Scadenza aggiornata');
}

function saveTodo(){
  const data=lll('todo');data.push({id:uid(),testo:document.getElementById('f-desc').value.trim(),pri:document.getElementById('f-pri').value,done:false});
  ls('todo',data);closeM();renderTodo();
}

function saveCfg2(){
  const f24key=(document.getElementById('cfg-f24key')||{}).value||'';
  const f24tipo=(document.getElementById('cfg-f24tipo')||{}).value||'all';
  const f24sync=(document.getElementById('cfg-f24sync')||{}).value||'60';
  const oldCfg=lCfg();
  // Salva markup per cliente
  const markup_clienti={};
  Object.keys(CMAP).forEach(cod=>{
    const el=document.getElementById('cfg-mk-'+cod);
    if(el&&el.value)markup_clienti[cod]=parseFloat(el.value)||2.25;
  });
  sCfg({
    nc_url:document.getElementById('cfg-url').value.trim().replace(/\/$/,''),
    nc_user:document.getElementById('cfg-user').value.trim(),
    nc_pass:document.getElementById('cfg-pass').value.trim(),
    f24_key:f24key||oldCfg.f24_key||'',
    f24_tipo:f24tipo,
    f24_sync:f24sync,
    f24_ok:oldCfg.f24_ok||false,
    markup_clienti
  });
  closeM();
  if(f24key&&f24key!==oldCfg.f24_key){
    testF24Key(f24key,f24tipo,f24sync);
  } else {
    syncNow();
  }
}

async function testF24Key(key,tipo,syncMin){
  const cfg=getCfg();
  if(!cfg.ok){showToast('Configura prima Nextcloud');return;}
  showToast('Salvataggio API key Fattura24…');
  // Scrive f24_config.txt in Nextcloud — lo script Python lo legge al prossimo ciclo
  await pushF24Config(key,tipo,syncMin);
  // Prova a verificare via proxy (opzionale, non bloccante)
  try{
    const r=await fetch(`${cfg.nc_url}/f24test`,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({api_key:key}),signal:AbortSignal.timeout(5000)});
    if(r.ok){
      const d=await r.json().catch(()=>({}));
      if(d.ok){const c=lCfg();c.f24_ok=true;sCfg(c);showToast('✓ API key Fattura24 attiva e verificata!');}
      else showToast('API key salvata — verifica su Fattura24 se attiva');
    } else {
      showToast('API key salvata in Nextcloud (/Dashboard/f24_config.txt)');
    }
  }catch{
    showToast('API key salvata — lo script la leggerà al prossimo sync');
  }
  syncNow();
}

async function pushF24Config(key,tipo,syncMin){
  const cfg=getCfg();
  if(!cfg.ok)return;
  const content=`F24_API_KEY=${key}\nF24_TIPO=${tipo}\nF24_SYNC_MIN=${syncMin}\n`;
  try{
    await fetch(`${cfg.nc_url}/remote.php/dav/files/${cfg.nc_user}/Dashboard/f24_config.txt`,{
      method:'PUT',
      headers:{'Authorization':'Basic '+btoa(cfg.nc_user+':'+cfg.nc_pass),'Content-Type':'text/plain'},
      body:content
    });
  }catch{/* silenzioso */}
}


function exportAll(){
  const blob=new Blob([JSON.stringify({commesse:lll('comm'),conti:lll('conti'),scadenze:lll('scad'),todos:lll('todo'),dettagli:ll('det'),exported:new Date().toISOString()},null,2)],{type:'application/json'});
  const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=`asg-backup-${new Date().toISOString().slice(0,10)}.json`;a.click();
}

// ═══════════════════════════════════════════════
// DATABASE FORNITORI
// ═══════════════════════════════════════════════
const FORNITORI=[
  {
    id:'ARES',nome:'Ares Elettronica / Tecfluid',zona:'Reggio Emilia (RE)',colore:'#166534',emoji:'⚙',
    url:'https://www.areselettronica.it',
    categorie:['Minicentraline PU05/PU10/PU20','Collettori modulari','Valvole integrate','Pompe Gr.1/2','Motori DC/AC','Serbatoi','Blocchi modulari','Distanziali','Bobine','Controllo elettronico'],
    note:'Fornitore principale per minicentraline modulari ASG. Ottimi prezzi e disponibilità su serie PU. Listino 2022 disponibile. Stesso gruppo di Tecfluid (RE). Tecfluid produce anche sistemi completi per applicazioni industriali (sollevatori, dock leveller, forbici) e mobile (sponde camion, ribaltabili, barche).',
    accordo:'Listino accordato 2022',
    contatti:{tel:'+39 0522 926 369',email:'info@tecfluid.it'},
    applicazioni:[
      {cat:'Garage & Sollevamento',emoji:'🚗',items:['Ponte sollevatore 2 colonne (max 5.5 kW, 4000 PSI, 5 GPM)','Ponte HD a 2 colonne con livellamento elettronico','Colonna mobile (12/24V DC + AC, Energy Recovery)','Piattaforma sollevatore HD (fino a 13 kW, 6 GPM, serbatoio 25 GAL)','Forbice sollevatore (max 4 kW, livellamento opzionale)','Cambia pneumatici truck (5 assi, pompa silenzio)','Cambia pneumatici HD','Sollevatore moto (telecomando opzionale)']},
      {cat:'Logistica & Dock',emoji:'🏭',items:['Dock leveller swing lip (max 1.5 kW, 3000 PSI)','Dock leveller telescopic lip','Combi leveller truck & van','Truck immobilizer (soluzione CETOP + cartucce)']},
      {cat:'Industriale Generale',emoji:'🔧',items:['Compattatore carta (soft-switch CETOP)','Colonnina mobile (IP67 submerged option)','Car parking system (pompa silenzio)','Tavola elevatrice (remote control, versione sommersa)']},
      {cat:'Mobile — Camion',emoji:'🚚',items:['Sponda idraulica (12/24V DC, 4000 PSI)','Sponda US version','Ribaltabile tipper (wireless remote)','Wing body (wired/wireless)','Snowplow pick-up (wireless)','Wheelchair lift']},
      {cat:'Mobility',emoji:'🚐',items:['Stabilizzatore caravan (livellamento elettronico)','Platform scissor lift (brushless, energy recovery)','Pick-up tow truck','Pick-up scissor lift','Forklift (brushless, energy recovery)','Sweeper']},
      {cat:'Marine',emoji:'⛵',items:['Passerella (gangway, wireless)','Tender lift','Boat lift (wired/wireless)','Boat steering system (24V brushless)']},
    ]
  },
  {
    id:'CPA',nome:'CPA Fluid',zona:'Rivoli (TO)',colore:'#1a56db',emoji:'🔵',
    url:'https://www.cpafluid.eu',
    categorie:['Elettrovalvole Atos','Distributori CETOP','Pompe ingranaggi','Pompe palette/vite','Pompe variabili','Valvole di massima','Valvole riduttrici','Valvole di flusso','Valvole di ritegno','Controbilanciamento','Giunti e lanterne','Filtri','Accumulatori'],
    note:'Vasto magazzino con spedizione rapida DHL. Distributore principale Atos. Ottimo per componenti standard con consegna urgente. E-shop con prezzi visibili.',
    accordo:'Prezzi concordati — accesso e-shop',
    contatti:{tel:'+39 011 9591521',email:'info@cpafluid.eu'}
  },
  {
    id:'FIZ',nome:'FIZ Srl / Tecnical Impianti',zona:'Verona + Bergamo',colore:'#5b21b6',emoji:'🟣',
    url:'https://www.fizsrl.it',
    categorie:['Bosch Rexroth','Elettrovalvole proporzionali','Distributori DN10/16/25','Pompe pistoni assiali','Pompe palette variabili','Motori idraulici','Cilindri','Centraline su misura','Tubi flessibili','Pneumatica','SKF','Festo','Siemens'],
    note:'Gruppo ampio con Tecnical Impianti (partner diretto). Accesso a catalogo completo Bosch Rexroth a ottimi prezzi. Molti costruttori alternativi gestiti nel gruppo. Buona disponibilità e tempi.',
    accordo:'Accordo tramite Tecnical Impianti',
    contatti:{tel:'',email:''}
  },
  {
    id:'PACO',nome:'Pa.Co Fluid',zona:'Legnano (MI) + Parma (PR)',colore:'#0369a1',emoji:'🔧',
    url:'https://pacofluid.it',
    categorie:['Tubi flessibili raccordati','Raccorderia BSP/UNF/JIC/ORFS','Adattatori','Tubi rigidi','Valvole a sfera','Valvole di ritegno','Scambiatori di calore','Pneumatica','Cilindri pneumatici','Accessori serraggio'],
    note:'Specialista tubi flessibili raccordati e raccorderia. Ampio magazzino ~2000 articoli. Assemblaggio e collaudo in sede. Ottimo per urgenze su tubi e raccordi.',
    accordo:'Prezzi concordati — consegna rapida',
    contatti:{tel:'+39 0331 455644',email:'info@pacosrl.it'}
  },
  {
    id:'COLFER',nome:'Colfer Srl',zona:'Monza (MB)',colore:'#78350f',emoji:'🏭',
    url:'',
    categorie:['Carpenteria su disegno','Serbatoi oleodinamici speciali','Lamierazione','Verniciatura epossidica','Strutture saldate','Basamenti motori-pompa'],
    note:'Partner carpenteria e verniciatura. Serbatoi e strutture a disegno in tempi brevi. Verniciatura RAL epossidica/poliuretanica. Ideale per serbatoi speciali fuori standard e basamenti personalizzati.',
    accordo:'Collaborazione diretta',
    contatti:{tel:'',email:''}
  },
  {
    id:'MORATTI',nome:'M.B. Moratti',zona:'Rovato (BS)',colore:'#0f766e',emoji:'🔩',
    url:'',
    categorie:['Cilindri oleodinamici saldati','Cilindri a disegno','Cilindri grandi dimensioni','Cilindri telescopici','Riparazione cilindri'],
    note:'Produzione cilindri oleodinamici saldati su misura, anche di grandi dimensioni. Ottimo per cilindri non standard e riparazioni. Tempi di consegna competitivi.',
    accordo:'Collaborazione diretta',
    contatti:{tel:'',email:''}
  },
  {
    id:'GTA',nome:'GTA Lombardia',zona:'Lombardia',colore:'#64748b',emoji:'📦',
    url:'',
    categorie:['Catalogo ampio generico','Valvole','Pompe','Motori','Cilindri','Raccorderia','Filtri','Serbatoi','Accessori'],
    note:'Catalogo molto ampio. Prezzi e tempi non sempre ottimali ma utile per componenti difficilmente reperibili altrove o per completare un ordine.',
    accordo:'Listino standard',
    contatti:{tel:'',email:''}
  },
  {
    id:'SMEM',nome:'Smem Srl',zona:'Nord Italia',colore:'#b91c1c',emoji:'⚡',
    url:'',
    categorie:['Motori elettrici AC trifase','Motori AC monofase','Motori IE3','Motori ATEX','Motori freno','Riduttori'],
    note:'Ottima disponibilità e prezzi competitivi su motori elettrici AC. Riferimento principale per motori standard da catalogo con consegna rapida.',
    accordo:'Prezzi concordati',
    contatti:{tel:'',email:''}
  },
  {
    id:'XF',nome:'XF Srl (Fox Group)',zona:'Italia',colore:'#92400e',emoji:'📊',
    url:'',
    categorie:['Pressostati','Accumulatori a vescica','Accumulatori a membrana','Accumulatori a pistone','Trasduttori di pressione','Manometri digitali'],
    note:'Società del gruppo Fox Srl. Riferimento per pressostati e accumulatori. Buona disponibilità su gamma standard.',
    accordo:'Listino concordato',
    contatti:{tel:'',email:''}
  },
  {
    id:'PIOTTI',nome:'Piotti Oleodinamica',zona:'Italia',colore:'#4338ca',emoji:'🛢',
    url:'',
    categorie:['Componenti oleodinamica generici','Valvole','Pompe','Filtri','Raccorderia'],
    note:'Distributore oleodinamica generale.',
    accordo:'Da definire',
    contatti:{tel:'',email:''}
  },
  {
    id:'BART',nome:'Bart Srl',zona:'Milano (MI)',colore:'#0891b2',emoji:'🔬',
    url:'https://www.bart-e.com',
    categorie:['Manometri a secco','Manometri a glicerina','Manometri ATEX','Manometri isometrici','Salvamanometri / esclusori','Termometri','Valvole a sfera inox','Raccordi strumentazione'],
    note:'Specialista manometri e strumentazione. Gamma completa da 0 a 1000 bar. Prodotti certificati CE/Accredia. Versioni ATEX disponibili. Salvamanometri esclusori di vari tipi.',
    accordo:'Listino concordato',
    contatti:{tel:'',email:'info@bart-e.com'}
  },
];

// ═══════════════════════════════════════════════
// TASSONOMIA CATEGORIE OLEODINAMICA
// ═══════════════════════════════════════════════
const CAT_TREE=[
  {id:'valvole',label:'Valvole',emoji:'⚡',colore:'#5b21b6',sub:[
    {id:'val-direz',label:'Direzionali (Distributori)',fornitori:['CPA','FIZ','GTA'],sub2:['A piastra CETOP3 (DN6)','A piastra CETOP5 (DN10)','A piastra CETOP7 (DN16)','In linea / a cartuccia','Proporzionali','Manuali/meccaniche','Sicurezza monitorate']},
    {id:'val-press',label:'Pressione',fornitori:['CPA','FIZ','ARES','GTA'],sub2:['Valvole di massima (in linea)','Valvole di massima (a piastra)','Proporzionali di pressione','Riduttrici di pressione','Sequenza','Scarico / unloading']},
    {id:'val-flusso',label:'Flusso',fornitori:['CPA','ARES','FIZ','GTA'],sub2:['Regolatrici compensate (in linea)','Regolatrici (a piastra)','Proporzionali di flusso','Divisori di flusso 50/50','Prioritarie']},
    {id:'val-ritegno',label:'Ritegno & Bilanciamento',fornitori:['CPA','ARES','FIZ','PACO','GTA'],sub2:['Ritegni a cartuccia','Ritegni in linea','Ritegni pilotati','Valvole di bilanciamento (controbilanciamento)','Ritegni CETOP']},
    {id:'val-blocco',label:'Valvole di blocco / tenuta',fornitori:['CPA','ARES','FIZ'],sub2:['A cartuccia singola tenuta NC','A cartuccia doppia tenuta','Elettriche a cartuccia (V42/V43)','Pneumatiche doppia tenuta']},
  ]},
  {id:'pompe',label:'Pompe',emoji:'🔄',colore:'#0f766e',sub:[
    {id:'p-ingr',label:'Ad ingranaggi esterni',fornitori:['ARES','CPA','FIZ','GTA'],sub2:['Gruppo 05 (max 2cc)','Gruppo 1 (max 9.8cc)','Gruppo 2 (max 36cc)','Gruppo 3','Pompe reversibili Gr.1','Pompe doppie Gr.2+Gr.1','Pompe a rotori elicoidali (silenziose)']},
    {id:'p-pale',label:'A palette',fornitori:['CPA','FIZ'],sub2:['Cilindrata fissa','Cilindrata variabile','Bassa rumorosità']},
    {id:'p-pist',label:'A pistoni assiali',fornitori:['FIZ','CPA'],sub2:['Cilindrata fissa','Cilindrata variabile','Con regolatore di pressione','Con regolatore Load Sensing']},
    {id:'p-spec',label:'Speciali',fornitori:['CPA','GTA'],sub2:['Pompe a vite','Pompe a membrana','Hand pump (manuali)']},
  ]},
  {id:'motori-el',label:'Motori elettrici',emoji:'⚡',colore:'#b91c1c',sub:[
    {id:'mot-ac',label:'AC trifase (B14/B5)',fornitori:['SMEM','ARES','FIZ'],sub2:['MEC71 0.25-0.55kW','MEC80 0.75-1.1kW','MEC90 1.5-2.2kW','MEC100-112 3-4kW','MEC132 5.5-9.2kW','IE3 alta efficienza','ATEX','Freno','Flangia quadra (FQ)']},
    {id:'mot-mono',label:'AC monofase',fornitori:['SMEM','ARES'],sub2:['MEC71 0.25-0.55kW','MEC80 0.75-1.1kW','MEC90 1.5-2.2kW','Condensatore avviamento']},
    {id:'mot-dc',label:'DC (12/24/48V)',fornitori:['ARES'],sub2:['12V 500W-2400W','24V 500W-3000W','48V','Con ventilazione forzata','Reversibili IP44','Non reversibili IP54']},
    {id:'mot-rid',label:'Riduttori',fornitori:['FIZ','SMEM'],sub2:['Riduttori coassiali','Riduttori ortogonali']},
  ]},
  {id:'mot-idr',label:'Motori idraulici',emoji:'🔁',colore:'#0369a1',sub:[
    {id:'mi-ingr',label:'Ad ingranaggi',fornitori:['FIZ','CPA','GTA'],sub2:['Gr.1 piccoli','Gr.2 medi','Gr.3 grandi']},
    {id:'mi-orb',label:'Orbitali (gerotori)',fornitori:['FIZ','CPA'],sub2:['50-800cc/giro','Con freno','Con riduttore']},
    {id:'mi-pist',label:'A pistoni assiali',fornitori:['FIZ'],sub2:['Cilindrata fissa','Cilindrata variabile']},
  ]},
  {id:'cilindri',label:'Cilindri idraulici',emoji:'📐',colore:'#92400e',sub:[
    {id:'cil-sal',label:'Saldati su misura',fornitori:['MORATTI'],sub2:['Alesaggio 40-500mm','Stelo singolo','Stelo passante','Telescopici','Grandi dimensioni']},
    {id:'cil-tir',label:'A tiranti (standard)',fornitori:['FIZ','CPA','GTA'],sub2:['ISO 6020/1','ISO 6020/2','ISO 6022','Compatti']},
    {id:'cil-rep',label:'Riparazione cilindri',fornitori:['MORATTI'],sub2:['Revisione guarnizioni','Rettifica stelo','Rettifica camicia']},
  ]},
  {id:'tubi-fless',label:'Tubi flessibili',emoji:'🌀',colore:'#0891b2',sub:[
    {id:'tf-rac',label:'Raccordati su misura',fornitori:['PACO','FIZ'],sub2:['DN6 (1/4")','DN10 (3/8")','DN12 (1/2")','DN16 (5/8")','DN19 (3/4")','DN25 (1")','DN32 (1"1/4)','Acciaio al carbonio','Acciaio inox','PVC']},
    {id:'tf-rig',label:'Tubi rigidi',fornitori:['PACO'],sub2:['In barre 6m','Rivestiti','Inox']},
  ]},
  {id:'raccord',label:'Raccorderia',emoji:'🔩',colore:'#0891b2',sub:[
    {id:'rac-bsp',label:'BSP (gas)',fornitori:['PACO','CPA','GTA'],sub2:['Maschio-maschio','Femmina-femmina','Riduzioni','A gomito 90°','A T']},
    {id:'rac-unf',label:'UNF (SAE)',fornitori:['PACO','ARES','CPA'],sub2:['7/16-20 UNF','9/16-18 UNF','3/4-16 UNF','7/8-14 UNF','1-1/16-12 UNF']},
    {id:'rac-jic',label:'JIC / 37°',fornitori:['PACO','CPA'],sub2:['Maschio JIC','Femmina JIC']},
    {id:'rac-orfs',label:'ORFS (O-ring face seal)',fornitori:['PACO','CPA'],sub2:['Maschio ORFS','Femmina ORFS','Perni']},
    {id:'rac-metr',label:'Metrici',fornitori:['PACO','CPA','GTA'],sub2:['M12×1.5','M14×1.5','M16×1.5','M18×1.5','M22×1.5','M26×1.5','M33×2']},
    {id:'rac-adatt',label:'Adattatori misti',fornitori:['PACO','CPA','GTA'],sub2:['BSP → UNF','BSP → Metrico','JIC → ORFS','Riduzioni universali']},
  ]},
  {id:'serbatoi',label:'Serbatoi & Vasche',emoji:'🛢',colore:'#78350f',sub:[
    {id:'serb-std',label:'Standard serie ASG',fornitori:['ARES'],sub2:['Lamiera 1-3L (PU05/10)','Lamiera 4-12L (PU10)','Plastica 1-12L','Lamiera Ø175 30-90L (PU20)']},
    {id:'serb-spec',label:'Speciali a disegno',fornitori:['COLFER'],sub2:['Vasche quadre/rettangolari','Con flange custom','Con divisori interni','Verniciatura epossidica RAL','Con kit filtraggio completo']},
  ]},
  {id:'filtri',label:'Filtri',emoji:'🔽',colore:'#166534',sub:[
    {id:'fil-asp',label:'In aspirazione',fornitori:['ARES','CPA','GTA'],sub2:['G1/4" fino 5 l/min','G3/8" fino 15 l/min','1/2" fino 30+ l/min','Magnetici']},
    {id:'fil-man',label:'In mandata (alta pressione)',fornitori:['CPA','FIZ','GTA'],sub2:['10 micron','25 micron','Corpo alluminio','Corpo ghisa']},
    {id:'fil-rit',label:'In ritorno (bassa pressione)',fornitori:['CPA','ARES','GTA'],sub2:['20 micron','10 micron','Con indicatore intasamento']},
    {id:'fil-line',label:'In linea',fornitori:['CPA','FIZ'],sub2:['Doppi (change-over)','Con bypass termico']},
  ]},
  {id:'strument',label:'Strumentazione',emoji:'📏',colore:'#0891b2',sub:[
    {id:'str-man',label:'Manometri',fornitori:['BART','CPA','GTA'],sub2:['A secco Ø63 fino 400 bar','A glicerina Ø63','Isometrici alta sovrappressione','ATEX','Con contatti elettrici','0-60 / 0-150 / 0-250 / 0-400 bar']},
    {id:'str-salva',label:'Salvamanometri / esclusori',fornitori:['BART','CPA'],sub2:['1/4" BSP diritti','1/4" BSP a 90°','Con smorzatore']},
    {id:'str-press',label:'Pressostati',fornitori:['XF','CPA'],sub2:['Regolabili NC/NA','Con isteresi','4-20mA']},
    {id:'str-trasd',label:'Trasduttori pressione',fornitori:['XF','CPA','FIZ'],sub2:['0-10V','4-20mA','IO-Link','ATEX']},
    {id:'str-temp',label:'Termometri & sensori T°',fornitori:['BART','CPA'],sub2:['A bimetallo','A espansione gas','Elettronici PT100']},
    {id:'str-flow',label:'Misuratori di portata',fornitori:['CPA','FIZ'],sub2:['Ad ingranaggi','Elettromagnetici','Ultrasuoni']},
  ]},
  {id:'accum',label:'Accumulatori',emoji:'🔋',colore:'#4338ca',sub:[
    {id:'acc-vesg',label:'A vescica',fornitori:['XF','CPA','FIZ'],sub2:['0.75-50L','Alta pressione 350-500 bar','Bassa pressione']},
    {id:'acc-memb',label:'A membrana',fornitori:['XF','CPA'],sub2:['Piccoli 0.075-1L','Compatti']},
    {id:'acc-pist',label:'A pistone',fornitori:['XF','FIZ'],sub2:['Grandi volumi','Alta pressione']},
    {id:'acc-acc',label:'Accessori accumulatori',fornitori:['XF','CPA'],sub2:['Blocchi di sicurezza','Valvole di isolamento','Caricatori gas N2']},
  ]},
  {id:'scamb',label:'Scambiatori di calore',emoji:'🌡',colore:'#0891b2',sub:[
    {id:'sc-olio-aria',label:'Olio-aria',fornitori:['CPA','GTA'],sub2:['A ventilazione naturale','A ventilazione forzata 12/24V DC','Trifase 230/400V']},
    {id:'sc-olio-acq',label:'Olio-acqua',fornitori:['CPA','PACO','GTA'],sub2:['A fascio tubiero','A piastre']},
  ]},
  {id:'carpent',label:'Carpenteria speciale',emoji:'🏭',colore:'#78350f',sub:[
    {id:'carp-serb',label:'Serbatoi su disegno',fornitori:['COLFER'],sub2:['Lamiera nera verniciata','Acciaio inox','Con flange SAE/BSP']},
    {id:'carp-basam',label:'Basamenti e strutture',fornitori:['COLFER'],sub2:['Basamento motore-pompa','Supporti centralina','Quadri di controllo']},
  ]},
];

// ═══════════════════════════════════════════════
// STATO RICERCA FORNITORI
// ═══════════════════════════════════════════════
let _catSel=''; // categoria selezionata
let _fornSel=''; // fornitore chip selezionato

function showView(v){
  // Mappa vecchi ID view ai nuovi
  const map={'dashboard':'home','magazzino':'magazzino','fornitori':'fornitori','import':'import-fatture'};
  navTo(map[v]||v);
}

function copyRDO(){
  const txt=(document.getElementById('rdo-body')||{}).value||'';
  const subj=(document.getElementById('rdo-subj')||{}).value||'';
  const full=subj+'\n\n'+txt;
  navigator.clipboard.writeText(full).then(()=>showToast('Testo copiato negli appunti!'));
}

function salvaTemplate(fornId){
  const tpl=(document.getElementById('rdo-tpl-body')||{}).value||'';
  const templates=JSON.parse(localStorage.getItem('asg_rdo_templates')||'{}');
  templates[fornId]=tpl;
  localStorage.setItem('asg_rdo_templates',JSON.stringify(templates));
  showToast('✓ Modello salvato per questo fornitore');
  // Aggiorna il testo nel pannello componi se è visibile
  const f=FORNITORI.find(x=>x.id===fornId)||{nome:'Fornitore'};
  const oggi=new Date().toLocaleDateString('it-IT',{day:'2-digit',month:'long',year:'numeric'});
  const bodyEl=document.getElementById('rdo-body');
  if(bodyEl)bodyEl.value=tpl.replace('{FORNITORE}',f.nome).replace('{DATA}',oggi);
}

function ripristinaTemplatePredefinito(fornId){
  const templates=JSON.parse(localStorage.getItem('asg_rdo_templates')||'{}');
  delete templates[fornId];
  localStorage.setItem('asg_rdo_templates',JSON.stringify(templates));
  const el=document.getElementById('rdo-tpl-body');
  if(el)el.value=`Spett.le {FORNITORE},\n\nin riferimento ai Vs. prodotti/servizi, con la presente siamo a richiedere un'offerta per i seguenti articoli:\n\n1. [CODICE/DESCRIZIONE] — qty ___\n2. \n\nSi prega di indicare:\n- Prezzo unitario IVA esclusa\n- Disponibilità a magazzino\n- Tempi di consegna\n- Condizioni di resa\n\nRingraziamo anticipatamente e restiamo in attesa di un Vostro riscontro.\n\nCordiali saluti,`;
  showToast('Modello ripristinato al predefinito');
}


function renderSidebar(){
  const el=document.getElementById('forn-sidebar');
  el.innerHTML=`<div style="padding:8px 16px 6px;font-size:12px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.07em">Categorie</div>`+
  CAT_TREE.map(g=>`
    <div class="cat-group">
      <div class="cat-group-hd" onclick="toggleCatGroup(this)">
        <span>${g.emoji} ${g.label}</span><span class="arrow">›</span>
      </div>
      <div class="cat-sub">
        ${g.sub.map(s=>`<div class="cat-item" onclick="selectCat('${s.id}','${g.label} › ${s.label}')" data-id="${s.id}">
          ${s.label}
          <span class="cnt">${s.fornitori.length}</span>
        </div>`).join('')}
      </div>
    </div>`).join('');
}

function toggleCatGroup(el){
  el.classList.toggle('open');
}

function selectCat(id,label){
  _catSel=id;
  // evidenzia item attivo
  document.querySelectorAll('.cat-item').forEach(x=>x.classList.toggle('active',x.dataset.id===id));
  // trova sottocategoria
  let found=null;
  for(const g of CAT_TREE){
    const s=g.sub.find(x=>x.id===id);
    if(s){found={gruppo:g,sub:s};break;}
  }
  if(!found)return;
  // filtra fornitori rilevanti
  const ids=found.sub.fornitori;
  const fornFilt=FORNITORI.filter(f=>ids.includes(f.id));
  // aggiorna titolo
  document.getElementById('forn-section-title').textContent=label;
  // mostra sezione descrittiva
  const sub2=found.sub.sub2||[];
  document.getElementById('comp-results').innerHTML=sub2.length?`
    <div style="background:var(--bg2);border:1px solid var(--border);border-radius:var(--r-sm);padding:10px 14px;margin-bottom:14px;display:flex;flex-wrap:wrap;gap:6px;align-items:center">
      <span style="font-size:11px;font-weight:700;color:var(--text3);text-transform:uppercase;margin-right:4px">Sottotipi:</span>
      ${sub2.map(s=>`<span style="font-size:12px;background:#fff;border:1px solid var(--border);padding:3px 9px;border-radius:20px;color:var(--text2)">${s}</span>`).join('')}
    </div>`:'';
  renderFornCards(fornFilt.length?fornFilt:FORNITORI);
  document.getElementById('forn-cards-section').style.display='block';
  // pulizia chip
  document.querySelectorAll('.src-chip').forEach(x=>x.classList.remove('active'));
  document.querySelector('.src-chip').classList.add('active');
}

function renderFornCards(list){
  const el=document.getElementById('forn-grid');
  el.innerHTML=list.map(f=>{
    const cats=f.categorie.slice(0,5).map(c=>`<span class="fcat">${c}</span>`).join('');
    const extra=f.categorie.length>5?`<span class="fcat">+${f.categorie.length-5}</span>`:'';
    const urlBtn=f.url?`<a class="forn-btn" href="${f.url}" target="_blank">🌐 Sito</a>`:'';
    const telBtn=f.contatti.tel?`<a class="forn-btn" href="tel:${f.contatti.tel}">📞 Chiama</a>`:'';
    const rdoBtn=`<button class="forn-btn primary" onclick="openRDO('${f.id}')">✉ RDO</button>`;
    // Schede applicazione (solo per fornitori che le hanno)
    const appl=f.applicazioni&&f.applicazioni.length?`
      <div style="margin-top:10px;border-top:1px solid var(--border);padding-top:8px;">
        <div style="font-size:10px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.07em;margin-bottom:6px;">Applicazioni Tecfluid</div>
        <div style="display:grid;gap:4px;">
          ${f.applicazioni.map(a=>`
          <details style="border:1px solid var(--border);border-radius:6px;overflow:hidden;">
            <summary style="padding:5px 9px;font-size:12px;font-weight:700;cursor:pointer;list-style:none;display:flex;align-items:center;gap:6px;background:var(--bg2);">
              <span>${a.emoji}</span><span>${a.cat}</span><span style="margin-left:auto;font-size:10px;color:var(--text4)">${a.items.length} sistemi</span>
            </summary>
            <div style="padding:6px 10px;background:#fff;">
              ${a.items.map(i=>`<div style="font-size:11px;color:var(--text2);padding:2px 0;border-bottom:1px solid var(--border);">· ${i}</div>`).join('')}
            </div>
          </details>`).join('')}
        </div>
      </div>`:'';
    return`<div class="forn-card">
      <div class="forn-card-hd">
        <div class="forn-logo" style="background:${f.colore}">${f.emoji}</div>
        <div>
          <div class="forn-name">${f.nome}</div>
          <div class="forn-zone">📍 ${f.zona}</div>
          ${f.accordo?'<div style="font-size:10px;color:var(--green);font-weight:700;margin-top:2px">✓ ${f.accordo}</div>':''}
        </div>
      </div>
      <div class="forn-body">
        <div class="forn-cats">${cats}${extra}</div>
        <div class="forn-note">${f.note}</div>
        <div class="forn-actions">${urlBtn}${telBtn}${rdoBtn}</div>
        ${appl}
      </div>
    </div>`;
  }).join('');
}


function searchComp(){
  const q=(document.getElementById('comp-search').value||'').toLowerCase().trim();
  const resEl=document.getElementById('comp-results');
  const cardsEl=document.getElementById('forn-cards-section');
  if(q.length<2){
    resEl.innerHTML='';
    cardsEl.style.display='block';
    return;
  }
  cardsEl.style.display='none';
  // Cerca in categorie + sottocategorie + fornitori
  const hits=[];
  for(const g of CAT_TREE){
    for(const s of g.sub){
      const matchSub=s.label.toLowerCase().includes(q)||g.label.toLowerCase().includes(q);
      const matchSub2=(s.sub2||[]).filter(x=>x.toLowerCase().includes(q));
      if(matchSub||matchSub2.length){
        hits.push({tipo:'cat',gruppo:g,sub:s,matchSub2,score:matchSub2.length?2:1});
      }
    }
  }
  // Cerca in fornitori (nome, note, categorie)
  for(const f of FORNITORI){
    if(f.nome.toLowerCase().includes(q)||f.note.toLowerCase().includes(q)||f.categorie.some(c=>c.toLowerCase().includes(q))){
      hits.push({tipo:'forn',f,score:f.nome.toLowerCase().includes(q)?3:1});
    }
  }
  // Cerca nel listino Ares
  const aresHits=ARES.filter(a=>a.s.toLowerCase().includes(q)||a.d.toLowerCase().includes(q)).slice(0,6);

  hits.sort((a,b)=>b.score-a.score);

  let html='';
  if(aresHits.length){
    html+=`<div style="font-size:11px;font-weight:700;color:var(--green);text-transform:uppercase;letter-spacing:.07em;margin-bottom:6px">📋 Listino Ares / Tecfluid</div>`;
    html+=aresHits.map(a=>{
      const st=TIPO_BADGE[a.t]||'background:#f3f4f6;color:#374151;border:1px solid #d1d5db';
      return`<div class="comp-result">
        <div class="comp-cat-dot" style="background:${TIPO_BADGE[a.t]?'#166534':'#6b7280'}"></div>
        <div class="comp-main">
          <div class="comp-name">${a.s} — ${a.d}</div>
          <div class="comp-sub">Ares / Tecfluid · <span style="${st};font-size:10px;font-weight:700;padding:1px 6px;border-radius:20px">${a.serie||a.t}</span></div>
        </div>
        <div class="comp-actions">
          <span style="font-weight:700;color:var(--text)">€${fmt2(a.p)}</span>
          <button class="forn-btn primary" onclick="addAresFromSearch(${JSON.stringify(a).replace(/"/g,'&quot;')})">+ distinta</button>
        </div>
      </div>`;
    }).join('');
  }

  if(hits.filter(x=>x.tipo==='cat').length){
    html+=`<div style="font-size:11px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.07em;margin:12px 0 6px">🗂 Categorie</div>`;
    html+=hits.filter(x=>x.tipo==='cat').slice(0,6).map(h=>{
      const fIds=h.sub.fornitori;
      const tags=fIds.map(id=>{const f=FORNITORI.find(x=>x.id===id);return f?`<span class="comp-forn-tag" style="background:${f.colore}22;color:${f.colore};border-color:${f.colore}44">${f.nome.split(' ')[0]}</span>`:''}).join('');
      return`<div class="comp-result" onclick="selectCat('${h.sub.id}','${h.gruppo.label} › ${h.sub.label}')" style="cursor:pointer">
        <div class="comp-cat-dot" style="background:${h.gruppo.colore}"></div>
        <div class="comp-main">
          <div class="comp-name">${h.gruppo.emoji} ${h.gruppo.label} › ${h.sub.label}</div>
          <div class="comp-sub">${h.matchSub2.length?'Sottotipi: '+h.matchSub2.slice(0,3).join(', ')+(h.matchSub2.length>3?' …':''):''}</div>
          <div class="comp-forn" style="margin-top:5px">${tags}</div>
        </div>
        <div class="comp-actions"><span style="font-size:12px;color:var(--text3)">→</span></div>
      </div>`;
    }).join('');
  }

  if(hits.filter(x=>x.tipo==='forn').length){
    html+=`<div style="font-size:11px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.07em;margin:12px 0 6px">🏢 Fornitori</div>`;
    html+=hits.filter(x=>x.tipo==='forn').slice(0,4).map(h=>{
      const f=h.f;
      return`<div class="comp-result">
        <div class="comp-cat-dot" style="background:${f.colore}"></div>
        <div class="comp-main">
          <div class="comp-name">${f.emoji} ${f.nome}</div>
          <div class="comp-sub">📍 ${f.zona} ${f.accordo?'· ✓ '+f.accordo:''}</div>
        </div>
        <div class="comp-actions">
          ${f.url?'<a class="forn-btn" href="${f.url}" target="_blank">🌐</a>':''}
          <button class="forn-btn primary" onclick="openRDO('${f.id}')">✉ RDO</button>
        </div>
      </div>`;
    }).join('');
  }

  resEl.innerHTML=html||`<div class="empty">Nessun risultato per "<strong>${q}</strong>"</div>`;
}

function addAresFromSearch(a){
  // Se c'è una commessa aperta, aggiunge alla distinta, altrimenti apre modal commessa
  if(currentCommessa){
    useAres(a,currentCommessa.numero);
    showToast('Aggiunto alla distinta di '+currentCommessa.numero);
  } else {
    showToast('Apri una commessa per aggiungere il componente alla distinta');
  }
}

function openRDO(fornId){
  const f=FORNITORI.find(x=>x.id===fornId);
  if(!f)return;
  openModal('rdo',fornId);
}

function showToast(msg){
  let t=document.getElementById('_toast');
  if(!t){t=document.createElement('div');t.id='_toast';t.style.cssText='position:fixed;bottom:24px;left:50%;transform:translateX(-50%);background:#1e293b;color:#f1f5f9;font-size:13px;font-weight:600;padding:10px 20px;border-radius:30px;z-index:9999;opacity:0;transition:opacity .2s;white-space:nowrap;box-shadow:0 4px 16px rgba(0,0,0,.3)';document.body.appendChild(t);}
  t.textContent=msg;t.style.opacity='1';
  clearTimeout(t._to);t._to=setTimeout(()=>{t.style.opacity='0';},2500);
}

// ═══════════════════════════════════════════════
// IMPORT FATTURE — browser-side
// ═══════════════════════════════════════════════

let _importRows=[];

function handleDrop(e){
  e.preventDefault();
  document.getElementById('drop-zone').classList.remove('drag');
  const files=Array.from(e.dataTransfer.files).filter(f=>f.name.match(/\.(xls|xlsx|csv)$/i));
  if(files.length)processImportFiles(files);
}
function handleFileSelect(e){
  const files=Array.from(e.target.files);
  if(files.length)processImportFiles(files);
}

async function processImportFiles(files){
  setImpProgress(true,'Lettura file…');
  _importRows=[];
  document.getElementById('imp-preview-area').innerHTML='';
  for(const file of files){
    try{const rows=await parseF24File(file);_importRows.push(...rows);}
    catch(e){showToast('Errore su '+file.name+': '+e.message);}
  }
  setImpProgress(false);
  if(!_importRows.length){
    document.getElementById('imp-actions').style.display='none';
    document.getElementById('imp-preview-area').innerHTML=`<div class="empty" style="padding:16px">Nessuna riga riconosciuta. Usa "Esporta lista semplice in Excel" da Fattura24.</div>`;
    return;
  }
  renderImportPreview(_importRows);
  const actEl=document.getElementById('imp-actions');
  actEl.style.display='flex';
  document.getElementById('btn-confirm-import').textContent=`✓ Importa ${_importRows.length} fatture`;
}

async function parseF24File(file){
  return new Promise((resolve,reject)=>{
    const reader=new FileReader();
    reader.onload=async e=>{
      try{
        // Carica SheetJS dinamicamente se non presente
        if(typeof XLSX==='undefined'){
          await new Promise((res,rej)=>{
            const s=document.createElement('script');
            s.src='https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
            s.onload=res;s.onerror=()=>rej(new Error('Impossibile caricare SheetJS'));
            document.head.appendChild(s);
          });
        }
        const data=new Uint8Array(e.target.result);
        const wb=XLSX.read(data,{type:'array',cellDates:true,dateNF:'yyyy-mm-dd'});
        const ws=wb.Sheets[wb.SheetNames[0]];
        // Converti in array di array (raw, senza header)
        const aoa=XLSX.utils.sheet_to_json(ws,{header:1,raw:false,dateNF:'yyyy-mm-dd'});

        // Cerca tipo documento nelle prime righe
        const top=aoa.slice(0,8).map(r=>(r||[]).join(' ').toUpperCase()).join('\n');
        const tipoBase=top.includes('RICEVUT')||top.includes('ACQUIST')?'FA':'FE';

        // Trova la riga header (prima colonna = 'ID')
        let hdrIdx=-1;
        for(let i=0;i<aoa.length;i++){
          const r=aoa[i];
          if(r&&(r[0]==='ID'||(typeof r[0]==='string'&&r[0].trim()==='ID'))){
            hdrIdx=i;break;
          }
        }
        if(hdrIdx<0)throw new Error('Header "ID" non trovato. Usa "Esporta lista semplice in Excel" da Fattura24 (non "per righe").');

        const headers=aoa[hdrIdx].map(h=>String(h||'').trim());
        const rows=[];

        for(let i=hdrIdx+1;i<aoa.length;i++){
          const arow=aoa[i];if(!arow||!arow.length)continue;
          const idVal=String(arow[0]||'').trim();
          if(!idVal||!/^\d+$/.test(idVal))continue;

          const row={};
          headers.forEach((h,j)=>{row[h]=String(arow[j]??'').trim();});

          // toNum: gestisce sia numeri JS nativi (da SheetJS) sia stringhe formato italiano
          const toNum=v=>{
            if(v===null||v===undefined||v==='')return 0;
            if(typeof v==='number')return v; // SheetJS ha già convertito
            const s=String(v).trim();
            if(!s||s==='0')return 0;
            // Formato italiano: punto migliaia, virgola decimale → es. "1.234,56"
            // Formato anglosassone: virgola migliaia, punto decimale → es. "1,234.56"
            const hasCommaDecimal=/\d,\d{1,2}$/.test(s); // virgola negli ultimi 2 cifre = decimale italiano
            if(hasCommaDecimal){
              return parseFloat(s.replace(/\./g,'').replace(',','.'))||0;
            }
            // Altrimenti rimuovi virgole (migliaia anglosassoni) e parse normale
            return parseFloat(s.replace(/,/g,''))||0;
          };
          const importo=toNum(row['Totale documento']);
          const imponibile=toNum(row['Imponibile']);
          const iva=toNum(row['IVA (EUR)']);

          // Data: SheetJS con dateNF restituisce stringa dd/mm/yyyy
          const normD=d=>normDate(d); // usa normDate globale

          const incassato=(row['Incassato']||row['Versato']||'').toUpperCase().trim();
          const liquidato=(row['Liquidato']||'').toUpperCase().trim();
          const pagata=liquidato==='SI'||(incassato&&incassato!==''&&incassato!=='0'&&incassato!=='NULL'&&incassato!=='NAN'&&incassato!=='UNDEFINED');

          rows.push({
            tipo:tipoBase,
            numero:row['N. documento']||row['Numero']||'',
            data:normD(row['Data documento']||row['Data']||''),
            cliente:row['Ragione Sociale']||row['Cliente']||'',
            importo:importo.toFixed(2),
            imponibile:imponibile.toFixed(2),
            iva:iva.toFixed(2),
            stato:pagata?'pagata':'nonpagata',
            data_scad:normD(row['Data scadenza']||''),
            oggetto:row['Oggetto']||'',
            conto:row['Conto']||row['Metodo pagamento']||row['Banca']||'',
            id_f24:row['ID']||idVal,
            sync_ts:new Date().toISOString().slice(0,19),
          });
        }
        if(!rows.length)throw new Error('Nessuna riga trovata nel file. Controlla che il file sia un export "lista semplice" di Fattura24.');
        resolve(rows);
      }catch(err){reject(err);}
    };
    reader.onerror=()=>reject(new Error('Errore lettura file'));
    reader.readAsArrayBuffer(file);  // XLS binario → ArrayBuffer
  });
}


function normDate(d){
  if(!d||d==='null'||d==='NaN'||d==='undefined')return'';
  // yyyy-mm-dd già ok
  if(/^\d{4}-\d{2}-\d{2}$/.test(d))return d;
  // dd/mm/yyyy (italiano) → yyyy-mm-dd
  const mIT=d.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if(mIT){
    const giorno=parseInt(mIT[1]),mese=parseInt(mIT[2]);
    // Disambigua: se primo campo > 12 è sicuramente dd/mm, altrimenti ambiguo
    // Per sicurezza trattiamo sempre come dd/mm (formato italiano)
    return`${mIT[3]}-${mIT[2].padStart(2,'0')}-${mIT[1].padStart(2,'0')}`;
  }
  // mm/dd/yyyy (americano, da SheetJS) → yyyy-mm-dd
  const mUS=d.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  // già gestito sopra (stesso pattern) — SheetJS con dateNF:'yyyy-mm-dd' non dovrebbe produrlo
  // dd-mm-yyyy
  const mDash=d.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
  if(mDash)return`${mDash[3]}-${mDash[2].padStart(2,'0')}-${mDash[1].padStart(2,'0')}`;
  return d;
}

function setImpProgress(show,msg=''){
  const el=document.getElementById('imp-progress');
  if(!el)return;
  el.classList.toggle('show',show);
  if(msg){const t=document.getElementById('imp-progress-text');if(t)t.textContent=msg;}
}

function renderImportPreview(rows){
  const el=document.getElementById('imp-preview-area');
  if(!el)return;
  const fe=rows.filter(r=>r.tipo==='FE'),fa=rows.filter(r=>r.tipo==='FA');
  const totFE=fe.reduce((s,r)=>s+parseFloat(r.importo||0),0);
  const totFA=fa.reduce((s,r)=>s+parseFloat(r.importo||0),0);
  const preview=rows.slice(0,8);
  el.innerHTML=`
  <div style="display:flex;gap:12px;flex-wrap:wrap;margin:12px 0">
    ${fe.length?'<div style="background:#dcfce7;border:1px solid #6ee7a0;border-radius:8px;padding:8px 14px;font-size:13px"><strong style="color:#166534">${fe.length} fatture emesse</strong><span style="color:#166534;margin-left:6px">€ ${totFE.toFixed(2)}</span></div>':''}
    ${fa.length?'<div style="background:#dbeafe;border:1px solid #7fb3f5;border-radius:8px;padding:8px 14px;font-size:13px"><strong style="color:#1e40af">${fa.length} fatture ricevute</strong><span style="color:#1e40af;margin-left:6px">€ ${totFA.toFixed(2)}</span></div>':''}
  </div>
  <div class="imp-preview">
    <div class="imp-preview-hd">📋 Anteprima — prime ${Math.min(8,rows.length)} righe su ${rows.length}</div>
    <div style="overflow-x:auto"><table>
      <thead><tr><th>Tipo</th><th>Numero</th><th>Data</th><th>Cliente/Fornitore</th><th style="text-align:right">Importo</th><th>Stato</th><th>Scadenza</th></tr></thead>
      <tbody>${preview.map(r=>`<tr>
        <td><span class="imp-badge ${r.tipo==='FE'?'fe':'fa'}">${r.tipo}</span></td>
        <td style="font-family:monospace;font-size:11px">${r.numero}</td>
        <td style="white-space:nowrap">${fmtD(r.data)}</td>
        <td style="max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${r.cliente}</td>
        <td style="text-align:right;font-weight:700">€ ${parseFloat(r.importo).toFixed(2)}</td>
        <td><span style="font-size:10px;font-weight:700;padding:2px 6px;border-radius:20px;${r.stato==='pagata'?'background:#dcfce7;color:#166534':'background:#fee2e2;color:#991b1b'}">${r.stato}</span></td>
        <td style="color:var(--text3);white-space:nowrap">${fmtD(r.data_scad)||'—'}</td>
      </tr>`).join('')}</tbody>
    </table></div>
    ${rows.length>8?'<div style="padding:6px 12px;font-size:11px;color:var(--text3)">… e altre ${rows.length-8} righe</div>':''}
  </div>`;
}

async function confirmImport(){
  if(!_importRows.length){showToast('Nessuna riga da importare');return;}
  try{
    const existing=lll('fatt'); // leggi sempre da localStorage, non da SD che potrebbe essere stale
    const newRows=_importRows.filter(r=>r&&r.tipo);
    const merged=[...newRows,...existing.filter(r=>!newRows.find(n=>n.id_f24&&n.id_f24===r.id_f24))];
    merged.sort((a,b)=>(b.data||'').localeCompare(a.data||''));

    // Salva in localStorage E aggiorna SD
    ls('fatt',merged);
    SD.fatt=merged; // aggiorna in-memory immediatamente

    // Scadenze automatiche
    let nScad=0;
    try{nScad=aggiungiScadenzeDaFatture(newRows);}catch(e){console.warn('scadenze:',e);}

    // Nascondi bottoni import e svuota righe
    _importRows=[];
    const actEl=document.getElementById('imp-actions');
    if(actEl)actEl.style.display='none';

    // Feedback
    showToast(`✓ ${merged.length} fatture totali salvate${nScad?' · '+nScad+' scadenze':''}`);

    // Aggiorna status nella view import
    try{updateImpStatus();}catch(e){console.warn('updateImpStatus:',e);}

    // Upload Nextcloud in background
    const cfg=getCfg();
    const ncEl=document.getElementById('imp-nc-status');
    if(cfg.ok){
      if(ncEl)ncEl.textContent='⏳ Sincronizzazione Nextcloud…';
      try{
        const fields=['tipo','numero','data','cliente','importo','imponibile','iva','stato','data_scad','oggetto','id_f24','sync_ts'];
        const ok=await wPut('fatture.csv',merged,fields);
        if(ncEl)ncEl.textContent=ok?`✓ Sync OK · ${merged.length} fatture`:'⚠ Salvataggio solo locale';
      }catch(e){if(ncEl)ncEl.textContent='⚠ Errore sync: '+e.message;}
    } else {
      if(ncEl)ncEl.textContent='✓ Salvato in locale (Nextcloud non configurato)';
    }
  }catch(err){
    console.error('confirmImport error:',err);
    showToast('Errore import: '+err.message);
  }
}

function resetImport(){
  _importRows=[];
  const p=document.getElementById('imp-preview-area');if(p)p.innerHTML='';
  const a=document.getElementById('imp-actions');if(a)a.style.display='none';
  const fi=document.getElementById('file-input');if(fi)fi.value='';
}

function updateImpStatus(){
  const el=document.getElementById('imp-status-body');
  if(!el)return;
  const fatt=SD.fatt||lll('fatt');
  if(!fatt||!fatt.length){el.innerHTML=`<div style="font-size:13px;color:var(--text3)">Nessuna fattura importata ancora.</div>`;return;}
  const fe=fatt.filter(r=>r.tipo==='FE'),fa=fatt.filter(r=>r.tipo==='FA');
  // parseNum gestisce sia numeri che stringhe con virgola (es. "1.234,56")
  const parseNum=v=>{if(!v)return 0;const s=String(v).replace(/\./g,'').replace(',','.');return parseFloat(s)||0;};
  const totFE=fe.reduce((s,r)=>s+parseNum(r.importo),0);
  const totFA=fa.reduce((s,r)=>s+parseNum(r.importo),0);
  const lastSync=fatt[0]?.sync_ts||'—';
  el.innerHTML=`
  <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:12px;">
    <div style="background:var(--bg2);border:1px solid var(--border);border-radius:8px;padding:12px 14px">
      <div style="font-size:10px;font-weight:700;color:var(--text3);text-transform:uppercase;margin-bottom:4px">Fatture emesse</div>
      <div style="font-size:20px;font-weight:800;color:var(--green)">${fe.length}</div>
      <div style="font-size:12px;color:var(--text3)">€ ${fmt2(totFE)}</div>
    </div>
    <div style="background:var(--bg2);border:1px solid var(--border);border-radius:8px;padding:12px 14px">
      <div style="font-size:10px;font-weight:700;color:var(--text3);text-transform:uppercase;margin-bottom:4px">Fatture ricevute</div>
      <div style="font-size:20px;font-weight:800;color:var(--blue)">${fa.length}</div>
      <div style="font-size:12px;color:var(--text3)">€ ${fmt2(totFA)}</div>
    </div>
    <div style="background:var(--bg2);border:1px solid var(--border);border-radius:8px;padding:12px 14px">
      <div style="font-size:10px;font-weight:700;color:var(--text3);text-transform:uppercase;margin-bottom:4px">Ultimo import</div>
      <div style="font-size:13px;font-weight:600;color:var(--text)">${lastSync!=='—'?fmtD(lastSync.slice(0,10)):'—'}</div>
      <div style="font-size:12px;color:var(--text3)">${lastSync.slice(11,16)||''}</div>
    </div>
  </div>
  <button class="btn btn-ghost btn-sm" onclick="navTo('contabilita')">→ Vai a Contabilità per vedere le fatture</button>`;
}

function initScadenzeFiscali(){
  // Aggiunge scadenze fiscali standard per SRL con IVA trimestrale se non già presenti
  const existing=lll('scad');
  if(existing.some(s=>s._fiscale))return; // già inizializzato
  const y=new Date().getFullYear();
  const scadenze=[
    // IVA trimestrale (liquidazione + versamento entro il 16 del mese successivo al trim.)
    {data:`${y}-02-16`,descrizione:'Liquidazione IVA IV trim. anno prec. (16 feb)',importo:0,categoria:'iva',_fiscale:true},
    {data:`${y}-05-16`,descrizione:'Liquidazione IVA I trim. (16 mag) — acconto F24',importo:0,categoria:'iva',_fiscale:true},
    {data:`${y}-08-20`,descrizione:'Liquidazione IVA II trim. (entro 20 ago — proroga agosto)',importo:0,categoria:'iva',_fiscale:true},
    {data:`${y}-11-16`,descrizione:'Liquidazione IVA III trim. (16 nov)',importo:0,categoria:'iva',_fiscale:true},
    // Dichiarazione IVA annuale
    {data:`${y}-04-30`,descrizione:'Dichiarazione IVA annuale (entro 30 apr)',importo:0,categoria:'dichiarazione',_fiscale:true},
    // IRES / IRAP acconti e saldo
    {data:`${y}-06-30`,descrizione:'IRES/IRAP — saldo anno prec. + I acconto (30 giu)',importo:0,categoria:'ires',_fiscale:true},
    {data:`${y}-07-30`,descrizione:'IRES/IRAP — saldo con maggiorazione 0.4% (entro 30 lug)',importo:0,categoria:'ires',_fiscale:true},
    {data:`${y}-11-30`,descrizione:'IRES/IRAP — II acconto (30 nov)',importo:0,categoria:'ires',_fiscale:true},
    // INPS amministratore (compenso amministratore — versamento trimestrale)
    {data:`${y}-02-16`,descrizione:'INPS gestione separata — IV trim. anno prec. (Andrea)',importo:0,categoria:'inps',_fiscale:true},
    {data:`${y}-05-16`,descrizione:'INPS gestione separata — I trim. (Andrea)',importo:0,categoria:'inps',_fiscale:true},
    {data:`${y}-08-20`,descrizione:'INPS gestione separata — II trim. (Andrea)',importo:0,categoria:'inps',_fiscale:true},
    {data:`${y}-11-16`,descrizione:'INPS gestione separata — III trim. (Andrea)',importo:0,categoria:'inps',_fiscale:true},
    // Dichiarazione redditi SRL (modello Redditi SC)
    {data:`${y}-10-31`,descrizione:'Dichiarazione redditi SRL Modello Redditi SC (31 ott)',importo:0,categoria:'dichiarazione',_fiscale:true},
    // Bilancio / deposito CCIAA
    {data:`${y}-06-30`,descrizione:'Approvazione bilancio assemblea soci (entro 120 gg chiusura esercizio)',importo:0,categoria:'societario',_fiscale:true},
    {data:`${y}-07-30`,descrizione:'Deposito bilancio CCIAA (30 giorni da approvazione)',importo:0,categoria:'societario',_fiscale:true},
    // Libro unico del lavoro / CU compensi amministratore
    {data:`${y}-03-16`,descrizione:'CU compenso amministratore — consegna (16 mar)',importo:0,categoria:'sostituto',_fiscale:true},
    {data:`${y}-03-31`,descrizione:'CU compenso amministratore — invio telematico Agenzia Entrate (31 mar)',importo:0,categoria:'sostituto',_fiscale:true},
    // Ritenuta d'acconto su compenso amministratore (mensile se erogato)
    {data:`${y}-01-16`,descrizione:'F24 ritenuta compenso amm. dicembre anno prec.',importo:0,categoria:'ritenuta',_fiscale:true},
    // 770 sostituto d'imposta
    {data:`${y}-10-31`,descrizione:'Modello 770 sostituto d\'imposta (31 ott)',importo:0,categoria:'dichiarazione',_fiscale:true},
  ];
  const all=[...existing,...scadenze.map(s=>({...s,id:uid()}))];
  ls('scad',all);
}

function updateClock(){const el=document.getElementById('clk');if(el)el.textContent=new Date().toLocaleDateString('it-IT',{weekday:'short',day:'2-digit',month:'short'});}
setInterval(updateClock,60000);
updateClock();
window.addEventListener('DOMContentLoaded',async()=>{
  initScadenzeFiscali();
  navTo('home');
  renderAll();
  // Sync automatico all'apertura se Nextcloud è configurato
  setTimeout(()=>{
    const c=getCfg();
    if(c.ok){
      setSyncUI('loading','sincronizzazione…');
      syncNow().catch(()=>setSyncUI('err','errore sync'));
    } else {
      setSyncUI('','non configurato');
    }
  }, 800);
  const cfg=getCfg();
  if(cfg.ok){document.getElementById('cfg-banner').classList.remove('show');await syncNow();}
  else{document.getElementById('cfg-banner').classList.add('show');}
  setInterval(syncNow,300000);
});
function renderCommesse(){
  const el=document.getElementById('comm-list');if(!el)return;
  const comm=SD.comm||lll('comm');
  const q=(document.getElementById('comm-search')||{}).value?.toLowerCase()||'';
  const annoSel=(document.getElementById('comm-anno')||{}).value||'';
  const clienteSel=(document.getElementById('comm-cliente')||{}).value||'';

  // Popola select anno (solo se vuoto)
  const selAnno=document.getElementById('comm-anno');
  if(selAnno&&selAnno.options.length<=1){
    const anni=[...new Set(comm.map(c=>c.numero?.split('-').pop()).filter(Boolean))].sort((a,b)=>b-a);
    anni.forEach(a=>{const o=document.createElement('option');o.value=a;o.textContent=a;selAnno.appendChild(o);});
  }

  // Popola select cliente (solo se vuoto)
  const selCli=document.getElementById('comm-cliente');
  if(selCli&&selCli.options.length<=1){
    const codici=[...new Set(comm.map(c=>getCod(c.numero||'')).filter(Boolean))].sort();
    codici.forEach(cod=>{
      const o=document.createElement('option');
      o.value=cod;
      o.textContent=CMAP[cod]||CNAMES[cod]||cod;
      selCli.appendChild(o);
    });
  }

  // Aggiorna variabile dal select (per cambio da select)
  if(selCli)_filtroCommCliente=selCli.value;

  let filtered=comm.filter(c=>{
    if(_filtroCommStato&&c.stato!==_filtroCommStato)return false;
    if(annoSel&&!c.numero?.endsWith(annoSel))return false;
    if(_filtroCommCliente&&getCod(c.numero||'')!==_filtroCommCliente)return false;
    if(q&&!(c.numero||'').toLowerCase().includes(q)&&!(c.oggetto||'').toLowerCase().includes(q)&&!(CMAP[getCod(c.numero||'')]||c.cliente||'').toLowerCase().includes(q))return false;
    return true;
  }).sort((a,b)=>(b.data_apertura||b.numero||'').localeCompare(a.data_apertura||a.numero||''));
  if(!filtered.length){el.innerHTML='<div class="empty">Nessuna commessa trovata</div>';return;}
  const STATI_COMM={bozza:'badge-gray',offerta:'badge-amber',inviata:'badge-blue',confermata:'badge-teal','in-corso':'badge-green',chiusa:'badge-gray'};
  el.innerHTML=filtered.map(c=>{
    const st=c.stato||'';
    const badge=STATI_COMM[st]||'badge-gray';
    const cliente=CMAP[getCod(c.numero||'')]||c.cliente||'—';
    // Calcola margine se disponibile
    const det=ll('det');const cd=det[c.numero]||{distinta:[]};
    const costiDist=(cd.distinta||[]).reduce((s,r)=>s+parseFloat(r.netto_cad||r.costo_cad||0)*(r.qty||1),0);
    const importo=parseFloat(c.importo||0);
    const margPct=importo>0?Math.round((importo-costiDist)/importo*100):null;
    const fu=folderUrl(c.numero,'https://cloud.asglab.it');
    return`<div class="list-row" onclick="openDetail('${c.numero}')">
      <div class="sdot ${st==='in-corso'?'sdot-green':st==='confermata'?'sdot-teal':st==='inviata'?'sdot-blue':st==='offerta'?'sdot-amber':st==='chiusa'?'sdot-gray':'sdot-gray'}"></div>
      <div class="list-row-main">
        <div class="list-row-title">${c.numero} <span class="text-muted text-sm" style="font-weight:400">· ${cliente}</span></div>
        <div class="list-row-sub">${c.oggetto||'—'} · ${c.tipo_commessa||''}</div>
      </div>
      ${importo>0?'<div class="list-row-val">${\'€\'+fmt(importo)}</div>':''}
      ${margPct!==null?'<div class="list-row-val ${margPct>=30?\'green\':margPct>=15?\'amber\':\'red\'}" style="font-size:11px">${margPct}%</div>':''}
      <span class="badge ${badge}">${st}</span>
      ${fu?'<a href="'+fu+'" target="_blank" onclick="event.stopPropagation()" style="font-size:14px;text-decoration:none;opacity:.5;padding:2px 6px;" title="Apri cartella Nextcloud">📁</a>':'<span style="width:22px"></span>'}
    </div>`;
  }).join('');
}

// Aggiornamento det-margine quando si apre il dettaglio
const _origBuildTabs = typeof buildTabs !== 'undefined' ? buildTabs : null;
function _updateMargine(){
  if(!CC)return;
  const el=document.getElementById('det-margine');if(!el)return;
  const cd=CC._det||{distinta:[]};
  const dist=cd.distinta||[];
  const oreArr=cd.ore||[];

  // Costi distinta (netto × qty)
  const costiDist=dist.reduce((s,r)=>s+(parseFloat(r.netto_cad||r.costo_cad||0)*parseFloat(r.qty||1)),0);
  // Valore vendita distinta (prezzo_vendita × qty)
  const vendDist=dist.reduce((s,r)=>s+(parseFloat(r.prezzo_vendita||0)*parseFloat(r.qty||1)),0);
  // Ore valorizzate
  const costiOre=oreArr.reduce((s,o)=>s+(parseFloat(o.ore||0)*parseFloat(o.tariffa||0)),0);
  const totOreH=oreArr.reduce((s,o)=>s+parseFloat(o.ore||0),0);
  // Importo offerta (fonte prioritaria: offerta.importo_offerto, poi importo commessa)
  const importoOff=parseFloat(CC._det?.offerta?.importo_offerto||CC.importo||0);
  // Riferimento valore: usa offerta se disponibile, altrimenti vendita distinta
  const valRef=importoOff>0?importoOff:vendDist;
  const costiTot=costiDist+costiOre;
  const margLordo=valRef-costiDist;   // senza ore
  const margNetto=valRef-costiTot;    // con ore
  const percLordo=valRef>0?((margLordo/valRef)*100):null;
  const percNetto=valRef>0?((margNetto/valRef)*100):null;
  const col=percNetto===null?'var(--text3)':percNetto>=30?'var(--green)':percNetto>=15?'var(--amber)':'var(--red)';

  el.innerHTML=`
  <div style="font-size:10px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.06em;margin-bottom:6px">Margine</div>
  <div style="font-size:26px;font-weight:800;color:${col};line-height:1">${percNetto!==null?percNetto.toFixed(0)+'%':'—'}</div>
  <div style="font-size:10px;color:var(--text3);margin-top:2px;margin-bottom:6px">netto (con ore)</div>
  <div class="margine-bar" style="margin-bottom:8px"><div class="margine-fill ${!percNetto||percNetto<15?'low':percNetto<30?'mid':''}" style="width:${Math.min(100,Math.max(0,percNetto||0))}%"></div></div>
  <div style="display:grid;grid-template-columns:1fr 1fr;gap:3px 8px;font-size:10px;color:var(--text3);margin-bottom:6px;">
    <span>Offerta</span><span style="text-align:right;font-weight:600;color:var(--text2)">€${fmt2(valRef)}</span>
    <span>Costi materiali</span><span style="text-align:right;color:var(--red)">−€${fmt2(costiDist)}</span>
    ${costiOre>0?'<span>Ore (${totOreH}h)</span><span style="text-align:right;color:var(--red)">−€${fmt2(costiOre)}</span>':''}
    <span style="font-weight:600;color:var(--text2)">Margine lordo</span><span style="text-align:right;font-weight:600;color:${margLordo>=0?'var(--green)':'var(--red)'}">€${fmt2(margLordo)} ${percLordo!==null?'('+percLordo.toFixed(0)+'%)':''}</span>
    ${costiOre>0?'<span style="font-weight:600;color:var(--text2)">Margine netto</span><span style="text-align:right;font-weight:600;color:${col}">€${fmt2(margNetto)}</span>':''}
  </div>
  <button class="btn-ghost btn-sm" onclick="exportDistintaPDF(CC?.numero)" style="width:100%;margin-top:4px;font-size:10px">📄 Stampa distinta</button>`;
}


// Export PDF commesse con margini
function exportReportCommesse(){
  const allComms=lll('comm');
  const det=ll('det');
  // Rispetta filtri attivi
  const filtroStato=document.getElementById('comm-anno')?.closest('.filter-bar')?.querySelector('[data-f-stato].active')?.dataset?.fStato||'';
  const filtroAnno=(document.getElementById('comm-anno')||{}).value||'';
  const filtroCliente=(document.getElementById('comm-cliente')||{}).value||'';
  const comms=allComms.filter(c=>{
    if(filtroStato&&c.stato!==filtroStato)return false;
    if(filtroAnno&&!(c.numero||'').includes('-'+filtroAnno))return false;
    if(filtroCliente&&c.cliente!==filtroCliente)return false;
    return true;
  });
  const anno=new Date().getFullYear();

  // Calcola margine per ogni commessa
  const rows=comms.map(c=>{
    const cd=det[c.numero]||{distinta:[],ore:[]};
    const dist=cd.distinta||[];
    const oreArr=cd.ore||[];
    const costiDist=dist.reduce((s,r)=>s+(parseFloat(r.netto_cad||r.costo_cad||0)*parseFloat(r.qty||1)),0);
    const vendDist=dist.reduce((s,r)=>s+(parseFloat(r.prezzo_vendita||0)*parseFloat(r.qty||1)),0);
    const costiOre=oreArr.reduce((s,o)=>s+(parseFloat(o.ore||0)*parseFloat(o.tariffa||0)),0);
    const totOreH=oreArr.reduce((s,o)=>s+parseFloat(o.ore||0),0);
    const valRef=parseFloat(cd.offerta?.importo_offerto||c.importo||0)||vendDist;
    const costiTot=costiDist+costiOre;
    const margNetto=valRef-costiTot;
    const perc=valRef>0?(margNetto/valRef*100):null;
    return{...c,costiDist,costiOre,costiTot,vendDist,valRef,margNetto,perc,totOreH};
  }).sort((a,b)=>b.valRef-a.valRef);

  const totVal=rows.reduce((s,r)=>s+r.valRef,0);
  const totCosti=rows.reduce((s,r)=>s+r.costiTot,0);
  const totMarg=totVal-totCosti;

  const win=window.open('','_blank');
  win.document.write(`<!DOCTYPE html><html><head><meta charset="utf-8">
  <title>Report Commesse — ASG LAB</title>
  <style>
    body{font-family:sans-serif;font-size:11px;color:#111;margin:20px}
    h1{font-size:16px;margin-bottom:4px}
    .sub{color:#666;font-size:11px;margin-bottom:16px}
    table{width:100%;border-collapse:collapse;font-size:10px}
    th{background:#f0f0f0;padding:5px 6px;text-align:left;border-bottom:2px solid #ccc;font-weight:700}
    td{padding:4px 6px;border-bottom:1px solid #eee;vertical-align:top}
    .num{text-align:right}
    .green{color:#166534;font-weight:600}
    .red{color:#991b1b;font-weight:600}
    .amber{color:#92400e;font-weight:600}
    tfoot td{font-weight:700;background:#f7f7f7;border-top:2px solid #ccc}
    .stato{display:inline-block;padding:1px 5px;border-radius:8px;font-size:9px;font-weight:700}
    .st-chiusa{background:#dcfce7;color:#166534}
    .st-in-corso{background:#dbeafe;color:#1e40af}
    .st-other{background:#f3f4f6;color:#374151}
    @media print{button{display:none}}
  </style></head><body>
  <button onclick="window.print()" style="margin-bottom:12px;padding:6px 16px;background:#1a56db;color:#fff;border:none;border-radius:5px;cursor:pointer">🖨 Stampa / Salva PDF</button>
  <h1>Report Commesse — ASG LAB</h1>
  <div class="sub">Tutte le commesse · ${rows.length} totali · Estratto il ${new Date().toLocaleDateString('it-IT')}</div>
  <table>
    <thead><tr>
      <th>N°</th><th>Cliente</th><th>Oggetto</th><th>Anno</th><th>Stato</th>
      <th class="num">Offerta</th>
      <th class="num">Costi mat.</th>
      <th class="num">Ore (€)</th>
      <th class="num">Costi tot.</th>
      <th class="num">Margine</th>
      <th class="num">%</th>
    </tr></thead>
    <tbody>
    ${rows.map(r=>{
      const pClass=r.perc===null?'':r.perc>=30?'green':r.perc>=15?'amber':'red';
      const stClass=r.stato==='chiusa'?'st-chiusa':r.stato==='in-corso'?'st-in-corso':'st-other';
      return`<tr>
        <td style="font-family:monospace;font-size:10px">${r.numero}</td>
        <td>${CMAP[r.cliente]||r.cliente||'—'}</td>
        <td style="max-width:180px">${r.oggetto||'—'}</td>
        <td>${(r.numero||'').split('-')[2]||'—'}</td>
        <td><span class="stato ${stClass}">${r.stato||'—'}</span></td>
        <td class="num">${r.valRef>0?'€'+fmt2(r.valRef):'—'}</td>
        <td class="num red">${r.costiDist>0?'€'+fmt2(r.costiDist):'—'}</td>
        <td class="num">${r.costiOre>0?'€'+fmt2(r.costiOre)+' ('+r.totOreH+'h)':'—'}</td>
        <td class="num red">${r.costiTot>0?'€'+fmt2(r.costiTot):'—'}</td>
        <td class="num ${pClass}">${r.margNetto!==0?'€'+fmt2(r.margNetto):'—'}</td>
        <td class="num ${pClass}">${r.perc!==null?r.perc.toFixed(0)+'%':'—'}</td>
      </tr>`;
    }).join('')}
    </tbody>
    <tfoot><tr>
      <td colspan="5">TOTALE (${rows.filter(r=>r.valRef>0).length} commesse con offerta)</td>
      <td class="num">€${fmt2(totVal)}</td>
      <td class="num red">€${fmt2(rows.reduce((s,r)=>s+r.costiDist,0))}</td>
      <td class="num">€${fmt2(rows.reduce((s,r)=>s+r.costiOre,0))}</td>
      <td class="num red">€${fmt2(totCosti)}</td>
      <td class="num ${totMarg>=0?'green':'red'}">€${fmt2(totMarg)}</td>
      <td class="num ${totMarg>=0?'green':'red'}">${totVal>0?(totMarg/totVal*100).toFixed(0)+'%':'—'}</td>
    </tr></tfoot>
  </table>
  </body></html>`);
  win.document.close();
}

function exportDistintaPDF(numero){
  const comms=lll('comm');
  const c=comms.find(x=>x.numero===numero)||{numero,oggetto:'',cliente:'',importo:0};
  const det=ll('det');
  const cd=det[numero]||{distinta:[],ore:[]};
  const dist=cd.distinta||[];
  const ore=cd.ore||[];
  const totC=dist.reduce((s,r)=>s+(parseFloat(r.netto_cad||r.costo_cad||0)*parseFloat(r.qty||1)),0);
  const totV=dist.reduce((s,r)=>s+(parseFloat(r.prezzo_vendita||0)*parseFloat(r.qty||1)),0);
  const totOreEur=ore.reduce((s,o)=>s+(parseFloat(o.ore||0)*parseFloat(o.tariffa||0)),0);
  const totOreH=ore.reduce((s,o)=>s+parseFloat(o.ore||0),0);
  const importoOff=parseFloat(cd.offerta?.importo_offerto||c.importo||0)||totV;
  const margine=importoOff>0?((importoOff-totC-totOreEur)/importoOff*100):null;
  const sezioni={};
  dist.forEach(r=>{const s=r.sezione||'Generale';if(!sezioni[s])sezioni[s]=[];sezioni[s].push(r);});
  const win=window.open('','_blank');
  win.document.write(`<!DOCTYPE html><html><head><meta charset="utf-8">
  <title>Distinta ${numero}</title>
  <style>
    body{font-family:sans-serif;font-size:11px;color:#111;margin:20px;}
    h1{font-size:15px;margin:0 0 2px}
    .sub{color:#666;font-size:11px;margin-bottom:16px;}
    h2{font-size:12px;font-weight:700;margin:14px 0 4px;padding:4px 8px;background:#f0f0f0;border-radius:3px;}
    table{width:100%;border-collapse:collapse;font-size:10px;margin-bottom:8px;}
    th{background:#f7f7f7;padding:4px 6px;text-align:left;border-bottom:2px solid #ccc;font-size:10px;}
    td{padding:3px 6px;border-bottom:1px solid #eee;vertical-align:top;}
    .num{text-align:right;}
    .tot{font-weight:700;background:#f7f7f7;border-top:2px solid #ccc;}
    .summary{margin-top:16px;padding:10px;background:#f7f7f7;border-radius:5px;}
    .green{color:#166534;font-weight:700;} .red{color:#991b1b;font-weight:700;}
    @media print{button{display:none}}
  </style></head><body>
  <button onclick="window.print()" style="margin-bottom:12px;padding:6px 16px;background:#1a56db;color:#fff;border:none;border-radius:5px;cursor:pointer">🖨 Stampa / Salva PDF</button>
  <h1>Distinta materiali — ${numero}</h1>
  <div class="sub">${c.oggetto||'—'} · ${CMAP[c.cliente]||c.cliente||'—'} · ${new Date().toLocaleDateString('it-IT')}</div>
  ${Object.entries(sezioni).map(([sez,righe])=>
    '<h2>'+sez+'</h2><table><thead><tr><th>Descrizione</th><th>Q.</th><th>Fornitore</th><th>Codice</th><th class=num>Costo</th><th class=num>Netto</th><th class=num>Vendita</th><th class=num>Tot.vendita</th><th>Stato</th></tr></thead><tbody>'+
    righe.map(r=>'<tr><td>'+r.descrizione+'</td><td class=num>'+r.qty+'</td><td>'+(r.fornitore||'')+'</td><td style="font-family:monospace;font-size:9px">'+(r.codice||'')+'</td><td class=num>€'+fmt2(r.costo_cad||0)+'</td><td class=num>€'+fmt2(r.netto_cad||r.costo_cad||0)+'</td><td class=num>€'+fmt2(r.prezzo_vendita||0)+'</td><td class=num>€'+fmt2(parseFloat(r.prezzo_vendita||0)*parseFloat(r.qty||1))+'</td><td style="font-size:9px">'+(r.stato_approv||'')+'</td></tr>').join('')+
    '<tr class=tot><td colspan=7 style="text-align:right">Totale sezione:</td><td class=num>€'+fmt2(righe.reduce((s,r)=>s+(parseFloat(r.prezzo_vendita||0)*parseFloat(r.qty||1)),0))+'</td><td></td></tr>'+
    '</tbody></table>'
  ).join('')}
  ${ore.length?'<h2>Ore di lavoro</h2><table><thead><tr><th>Data</th><th>Descrizione</th><th>Operatore</th><th class=num>Ore</th><th class=num>Tariffa</th><th class=num>Totale</th></tr></thead><tbody>'+ore.map(o=>'<tr><td>'+fmtD(o.data)+'</td><td>'+(o.descrizione||'')+'</td><td>'+(o.operatore||'')+'</td><td class=num>'+o.ore+'h</td><td class=num>€'+fmt2(o.tariffa||0)+'/h</td><td class=num>€'+fmt2(parseFloat(o.ore||0)*parseFloat(o.tariffa||0))+'</td></tr>').join('')+'<tr class=tot><td colspan=5 style="text-align:right">Totale ore:</td><td class=num>€'+fmt2(totOreEur)+' ('+totOreH+'h)</td></tr></tbody></table>':''}
  <div class="summary">
    Costi materiali: €${fmt2(totC)} | Costi ore: €${fmt2(totOreEur)} | <strong>Totale costi: €${fmt2(totC+totOreEur)}</strong> | 
    Valore vendita: €${fmt2(totV)} ${importoOff>0?' | Offerta: €'+fmt2(importoOff)+' | <span class="'+(margine>=20?'green':'red')+'">Margine: '+(margine!==null?margine.toFixed(1)+'%':'—')+'</span>':''}
  </div>
  </body></html>`);
  win.document.close();
}



// Tab styling fix: .tab → .det-tab
document.addEventListener('DOMContentLoaded',()=>{
  // Override tab CSS per det-tabs
  const style=document.createElement('style');
  style.textContent=`
    #det-tabs .tab{padding:9px 14px;font-size:13px;font-weight:600;color:var(--text3);cursor:pointer;border-bottom:2px solid transparent;white-space:nowrap;background:none;border-top:none;border-left:none;border-right:none;}
    #det-tabs .tab.active{color:var(--blue);border-bottom-color:var(--blue);}
    #det-tabs .tab:hover:not(.active){background:var(--bg2);}
    .tab-content{display:none;padding:14px 20px;}
    .tab-content.active{display:block;}
  `;
  document.head.appendChild(style);
});


// ═══════════════════════════════════════════════
// FUNZIONI MANCANTI / ALIAS v13
// ═══════════════════════════════════════════════

function renderMagMetrics(){
  const data=lll('mag');
  const tot=data.length;
  const low=data.filter(m=>parseFloat(m.qty_min||0)>0&&parseFloat(m.qty||0)<parseFloat(m.qty_min||0)).length;
  const val=data.reduce((s,m)=>s+parseFloat(m.costo_cad||0)*parseFloat(m.qty||0),0);
  const setEl=(id,v)=>{const e=document.getElementById(id);if(e)e.textContent=v;};
  setEl('mag-tot',tot);
  setEl('mag-low',low||'0');
  setEl('mag-val','€'+fmt(val));
  const nb=document.getElementById('nb-mag');
  if(nb){if(low>0){nb.style.display='';nb.textContent=low;}else nb.style.display='none';}
  // Popola categorie
  const sel=document.getElementById('mag-cat');
  if(sel&&sel.options.length<=1){
    const cats=[...new Set(data.map(m=>m.categoria||'').filter(Boolean))].sort();
    cats.forEach(c=>{const o=document.createElement('option');o.value=c;o.textContent=c;sel.appendChild(o);});
  }
}

function apriRDO(fornId){
  // fornId può essere il nome del fornitore (da approvvigionamenti) o l'ID (da fornitori)
  const f=FORNITORI.find(x=>x.id===fornId||x.nome===fornId);
  if(f){
    // Precompila il template con i componenti da ordinare per questo fornitore
    const aggr=getApprovAggregati('daordinare');
    const items=aggr[f.nome]||aggr[f.id]||[];
    if(items.length){
      const templates=JSON.parse(localStorage.getItem('asg_rdo_templates')||'{}');
      if(!templates[f.id]){
        const righe=items.map(r=>`- ${r.descrizione||r.codice||'—'} (rif. ${r._commessa}) × ${r.qty||1}`).join('\n');
        const tplConArticoli=`Spett.le {FORNITORE},\n\nCon la presente richediamo offerta per:\n\n${righe}\n\nSi prega di indicare prezzi, disponibilità e tempi di consegna.\n\nCordiali saluti,`;
        const tmp=JSON.parse(localStorage.getItem('asg_rdo_templates')||'{}');
        tmp[f.id]=tplConArticoli;
        localStorage.setItem('asg_rdo_templates',JSON.stringify(tmp));
      }
    }
    openModal('rdo',f.id);
    return;
  }
  // Fornitore non in rubrica — modal semplice con textarea
  const aggr=getApprovAggregati('daordinare');
  const items=aggr[fornId]||[];
  const righe=items.map(r=>`- ${r.descrizione||r.codice||'—'} (rif. ${r._commessa}) × ${r.qty||1}`).join('\n');
  const testo=`Spett.le ${fornId},\n\nSi richiede offerta per:\n\n${righe||'[articoli]'}\n\nCordiali saluti,\nASG LAB SRL`;
  document.getElementById('modal-body').innerHTML=`<h3>✉ RDO — ${fornId}</h3>
    <textarea style="width:100%;min-height:200px;border:1.5px solid var(--border);border-radius:var(--r-sm);padding:10px;font-size:13px;font-family:var(--font);line-height:1.7">${testo}</textarea>
    <div class="modal-actions">
      <button class="btn-ghost" onclick="closeM()">Chiudi</button>
      <button class="btn btn-primary" onclick="navigator.clipboard.writeText(this.closest('#modal-box').querySelector('textarea').value);showToast('✓ Copiato')">📋 Copia</button>
    </div>`;
  document.getElementById('modal-overlay').classList.add('show');
}
if(typeof openRDO==='undefined'){
  function openRDO(fornId){
    const f=FORNITORI.find(x=>x.id===fornId);
    if(f)apriRDO(f.nome);
  }
}

// Se non esiste apriRDOApprov dal vecchio codice
if(typeof apriRDOApprov==='undefined'){
  function apriRDOApprov(numero,forn){apriRDO(forn);}
}

// Modal RDO manuale


// Funzione per aprire dettaglio (alias)
if(typeof openDetail==='function'&&typeof apriDettaglio==='undefined'){
  function apriDettaglio(n){openDetail(n);}
}

// renderTodo alias (per home)
if(typeof renderTodo==='undefined'){
  function renderTodo(){}
}

// Aggiorna DOMContentLoaded per navTo('home') invece di showView
// già gestito da window.addEventListener sotto


// ═══════════════════════════════════════════════
// ORDINAMENTO TABELLE
// ═══════════════════════════════════════════════
// _sortState dichiarato all'inizio con le variabili globali (riga ~1406)

function sortTable(tabella,col){
  if(_sortState[tabella].col===col){
    _sortState[tabella].dir*=-1;
  } else {
    _sortState[tabella].col=col;
    _sortState[tabella].dir=1;
  }
  if(tabella==='fatt')renderFatt();
  else if(tabella==='scad')renderScadContabilita();
}

function _sortData(data,col,dir){
  if(!col)return data;
  return [...data].sort((a,b)=>{
    let va=a[col]||'',vb=b[col]||'';
    // Numeri
    const na=parseFloat(va),nb=parseFloat(vb);
    if(!isNaN(na)&&!isNaN(nb))return(na-nb)*dir;
    // Stringhe/date
    return va.toString().localeCompare(vb.toString())*dir;
  });
}

function _thSort(tabella,col,label){
  const s=_sortState[tabella];
  const cls=s.col===col?(s.dir===1?'asc':'desc'):'';
  return`<th class="sortable ${cls}" style="padding:7px 8px;text-align:left;cursor:pointer" onclick="sortTable('${tabella}','${col}')">${label}</th>`;
}



// ═══════════════════════════════════════════════
// upgradeSelects — converte <select> in custom dropdown
// ═══════════════════════════════════════════════
function upgradeSelects(container){
  const box=container||document.getElementById('modal-box');
  if(!box)return;
  box.querySelectorAll('select:not(.keep-native)').forEach(sel=>{
    if(sel.closest('.csel'))return; // già convertito
    if(sel.style.display==='none')return;

    const options=[];
    Array.from(sel.children).forEach(child=>{
      if(child.tagName==='OPTGROUP'){
        const grp={group:child.label,items:[]};
        Array.from(child.children).forEach(o=>grp.items.push({value:o.value,label:o.textContent.trim()}));
        options.push(grp);
      } else if(child.tagName==='OPTION'){
        options.push({value:child.value,label:child.textContent.trim()});
      }
    });

    const value=sel.value||'';
    const id='csel-'+Math.random().toString(36).slice(2);
    const isSmall=sel.classList.contains('stato-approv-select')||sel.classList.contains('stato-sel');
    const w=sel.style.width||sel.offsetWidth>50?(sel.offsetWidth+'px'):'';

    // Ricava la stringa onChange dal listener onchange
    const onchangeAttr=sel.getAttribute('onchange')||'';

    const html=cselHTML(id,options,value,
      onchangeAttr.replace(/this\.value/g,'__VAL__').replace(/this\.options\[this\.selectedIndex\]\.text/g,'"__VAL__"'),
      {small:isSmall,width:w}
    );
    const div=document.createElement('div');
    div.innerHTML=html;
    const csel=div.firstElementChild;

    // Copia classi/stili rilevanti
    if(sel.style.width)csel.style.width=sel.style.width;

    sel.parentNode.replaceChild(csel,sel);
  });
}


// ═══════════════════════════════════════════════
// CUSTOM DROPDOWN v3 — overlay backdrop
// ═══════════════════════════════════════════════
(function(){
  let _popup = null;
  let _backdrop = null;
  let _popupOwner = null;

  function _hidePopup() {
    if (_backdrop) { _backdrop.remove(); _backdrop = null; }
    if (_popup) { _popup.remove(); _popup = null; }
    if (_popupOwner) {
      _popupOwner.querySelector('.csel-btn')?.classList.remove('open');
      _popupOwner = null;
    }
  }

  window._cselToggle = function(btn, id) {
    const owner = btn.closest('.csel');
    if (_popupOwner === owner) { _hidePopup(); return; }
    _hidePopup();

    const data = window._cselData?.[id];
    if (!data) return;

    // Backdrop trasparente — cattura click fuori senza interferire col popup
    const backdrop = document.createElement('div');
    backdrop.style.cssText = 'position:fixed;inset:0;z-index:99998;';
    backdrop.addEventListener('click', _hidePopup);
    document.body.appendChild(backdrop);
    _backdrop = backdrop;

    // Crea popup
    const popup = document.createElement('div');
    popup.className = 'csel-popup show';
    popup.style.zIndex = '99999';
    popup.innerHTML = data.opts.map(o => {
      if (o.group) {
        return `<div class="csel-group">${o.group}</div>` +
          (o.items||[]).map(i => `<div class="csel-item${i.v===data.val?' active':''}" tabindex="0">${i.l}</div>`).join('');
      }
      return `<div class="csel-item${o.v===data.val?' active':''}${o.v===''?' ph':''}" tabindex="0">${o.l}</div>`;
    }).join('');

    // Event delegation sul popup — un solo listener, nessun conflitto
    popup.addEventListener('click', function(e) {
      const item = e.target.closest('.csel-item');
      if (!item) return;
      e.stopPropagation();
      // Trova valore dall'indice
      const items = [...popup.querySelectorAll('.csel-item')];
      const idx = items.indexOf(item);
      const flat = [];
      data.opts.forEach(o => { if(o.items) o.items.forEach(i=>flat.push(i)); else flat.push(o); });
      const opt = flat[idx];
      if (opt) window._cselPick(id, opt.v, opt.l);
    });

    document.body.appendChild(popup);
    _popup = popup;
    _popupOwner = owner;
    btn.classList.add('open');

    // Posiziona
    const rect = btn.getBoundingClientRect();
    const spaceBelow = window.innerHeight - rect.bottom - 8;
    const spaceAbove = rect.top - 8;
    const ph = Math.min(280, 400);
    if (spaceBelow < 150 && spaceAbove > spaceBelow) {
      popup.style.bottom = (window.innerHeight - rect.top + 4) + 'px';
      popup.style.top = 'auto';
    } else {
      popup.style.top = (rect.bottom + 4) + 'px';
      popup.style.bottom = 'auto';
    }
    popup.style.left = rect.left + 'px';
    popup.style.minWidth = rect.width + 'px';
  };

  window._cselPick = function(id, value, label) {
    const data = window._cselData?.[id];
    if (!data) { _hidePopup(); return; }
    data.val = value;
    const owner = document.getElementById(id);
    if (owner) {
      const lbl = owner.querySelector('.csel-lbl');
      if (lbl) lbl.textContent = label;
    }
    _hidePopup();
    if (data.cb) {
      try { data.cb(value); } catch(e) { console.warn('csel cb:', e); }
    }
  };

  window._cselData = {};

  window.cselCreate = function(id, options, value, callback, {small=false, width='', placeholder='— seleziona —'}={}) {
    let label = placeholder;
    const flat = [];
    options.forEach(o => { if(o.items) o.items.forEach(i=>flat.push(i)); else flat.push(o); });
    const found = flat.find(o => o.v === value);
    if (found) label = found.l;
    window._cselData[id] = { opts: options, val: value, cb: callback };
    const w = width ? `style="width:${width}"` : '';
    return `<div class="csel${small?' small':''}" id="${id}" ${w}>
      <button type="button" class="csel-btn" onclick="window._cselToggle(this,'${id}')">
        <span class="csel-lbl">${label.replace(/</g,'&lt;')}</span>
        <span class="csel-arrow">▼</span>
      </button>
    </div>`;
  };

  window.upgradeSelects = function(container) {
    const box = container || document.getElementById('modal-box');
    if (!box) return;
    box.querySelectorAll('select:not(.keep-native)').forEach(sel => {
      if (sel.closest('.csel')) return;
      // Non saltare per display:none — il pannello potrebbe essere nascosto ma il select va upgradato
      const opts = [];
      Array.from(sel.children).forEach(child => {
        if (child.tagName === 'OPTGROUP') {
          const grp = { group: child.label, items: [] };
          Array.from(child.children).forEach(o => grp.items.push({v: o.value, l: o.textContent.trim()}));
          opts.push(grp);
        } else if (child.tagName === 'OPTION') {
          opts.push({v: child.value, l: child.textContent.trim()});
        }
      });
      const value = sel.value || '';
      const id = 'csel-' + Math.random().toString(36).slice(2,8);
      const onchangeStr = sel.getAttribute('onchange') || '';
      const isSmall = sel.classList.contains('stato-approv-select') || sel.classList.contains('stato-sel') || sel.classList.contains('small');
      const w = sel.style.width || (sel.offsetWidth > 50 ? sel.offsetWidth + 'px' : '');
      const cb = onchangeStr ? function(v) {
        try { (new Function(onchangeStr.replace(/this\.value/g, JSON.stringify(v)))).call({}); }
        catch(e) { try { eval(onchangeStr.replace(/this\.value/g, JSON.stringify(v))); } catch(e2) {} }
      } : null;
      const html = window.cselCreate(id, opts, value, cb, {small: isSmall, width: w});
      const div = document.createElement('div');
      div.innerHTML = html;
      const csel = div.firstElementChild;
      sel.parentNode.replaceChild(csel, sel);
    });
  };
})();

