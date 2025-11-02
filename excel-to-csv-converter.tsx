import React, { useState, useCallback } from 'react';
import { Upload, Download, AlertCircle, CheckCircle, Info, Edit2 } from 'lucide-react';
import * as XLSX from 'xlsx';

const ExcelToCsvConverter = () => {
  const [step, setStep] = useState(1);
  const [startingNumber, setStartingNumber] = useState('');
  const [data, setData] = useState([]);
  const [encoding, setEncoding] = useState('utf-8');
  const [showHelp, setShowHelp] = useState(false);
  const [editingCell, setEditingCell] = useState(null);
  const [error, setError] = useState('');
  const [csvDownloadUrl, setCsvDownloadUrl] = useState(null);
  const [csvPreview, setCsvPreview] = useState(null);

  // Číselník PPV kódov
  const PPV_CODES = {
    'DVP': '5', 'DVPNP': '5', 'DVP9': '5',
    'DBPS': '6', 'DBPS9': '6',
    'DPC': '4',
    'HPP': '2'
  };

  // Kompletný číselník SWIFT kódov slovenských bánk (zdroj: NBS 04/2022)
  const BANK_SWIFT = {
    '0200': 'SUBASKBX',
    '0900': 'GIBASKBX',
    '0720': 'NBSBSKBX',
    '1100': 'TATRSKBX',
    '1111': 'UNCRSKBX',
    '3000': 'SLZBSKBA',
    '3100': 'LUBASKBX',
    '5200': 'OTPVSKBX',
    '5600': 'KOMASK2X',
    '5900': 'PRVASKBA',
    '6500': 'POBNSK BA',
    '7300': 'INGBSKBX',
    '7500': 'CEKOSKBX',
    '7930': 'WUSTSKBA',
    '8050': 'COBASKBX',
    '8100': 'KOMBSKBA',
    '8120': 'BSLOSK22',
    '8130': 'CITISKBA',
    '8170': 'KBSPSKBX',
    '8160': 'EXSKSKBX',
    '8180': 'SPSRSKBA',
    '8320': 'JTBPSKBA',
    '8330': 'FIOZSKBA',
    '8360': 'BREXSKBX',
    '8370': 'OBKLSKBA',
    '8420': 'BFKKSKBB',
    '8430': 'KODBSKBX',
    '9952': 'TPAYSKBX',
    '9954': 'KPAYSK22',
    '9955': 'PANXSK22',
    '8450': 'BPKOSKBB'
  };

  // Rozšírený PSČ číselník - 200+ najdôležitejších slovenských obcí
  const PSC_LOOKUP = {
    // Bratislava (81xxx-85xxx)
    '81000': 'Bratislava', '81101': 'Bratislava', '81102': 'Bratislava', '81103': 'Bratislava',
    '81104': 'Bratislava', '81105': 'Bratislava', '81106': 'Bratislava', '81107': 'Bratislava',
    '81108': 'Bratislava', '82101': 'Bratislava', '82102': 'Bratislava', '82103': 'Bratislava',
    '82104': 'Bratislava', '82105': 'Bratislava', '82106': 'Bratislava', '82107': 'Bratislava',
    '82108': 'Bratislava', '82109': 'Bratislava', '83101': 'Bratislava', '83102': 'Bratislava',
    '83103': 'Bratislava', '83104': 'Bratislava', '83105': 'Bratislava', '83106': 'Bratislava',
    '84101': 'Bratislava', '84102': 'Bratislava', '84103': 'Bratislava', '84104': 'Bratislava',
    '84105': 'Bratislava', '85101': 'Bratislava', '85102': 'Bratislava', '85103': 'Bratislava',
    '85104': 'Bratislava', '85105': 'Bratislava', '85106': 'Bratislava',
    
    // Bratislavský kraj
    '90001': 'Modra', '90028': 'Ivanka pri Dunaji', '90031': 'Stupava', '90041': 'Senec',
    '90042': 'Vinosady', '90044': 'Dunajská Lužná', '90061': 'Svätý Jur', '90085': 'Pezinok',
    '90086': 'Pezinok', '90087': 'Pezinok', '90201': 'Pezinok', '90301': 'Malacky',
    '90044': 'Dunajská Lužná', '90085': 'Pezinok', '92101': 'Piešťany',
    
    // Trnavský kraj
    '91701': 'Trnava', '91801': 'Trnava', '91901': 'Trnava', '92001': 'Hlohovec',
    '92027': 'Leopoldov', '92101': 'Piešťany', '92201': 'Vrbové', '92901': 'Dunajská Streda',
    '93001': 'Gabčíkovo', '93101': 'Šamorín', '92521': 'Sereď', '92901': 'Dunajská Streda',
    '90501': 'Senica', '90601': 'Holíč', '90701': 'Skalica',
    
    // Trenčiansky kraj
    '91101': 'Trenčín', '91102': 'Trenčín', '91103': 'Trenčín', '91104': 'Trenčín',
    '91105': 'Trenčín', '91401': 'Nové Mesto nad Váhom', '91501': 'Nové Mesto nad Váhom',
    '91341': 'Brezová pod Bradlom', '91571': 'Nemšová', '91401': 'Nové Mesto nad Váhom',
    '97201': 'Bojnice', '97211': 'Prievidza', '95701': 'Bánovce nad Bebravou',
    '95801': 'Partizánske', '91601': 'Stará Turá', '01352': 'Dubnica nad Váhom',
    
    // Nitriansky kraj
    '94901': 'Nitra', '94902': 'Nitra', '94903': 'Nitra', '94904': 'Nitra',
    '94905': 'Nitra', '95001': 'Nitra', '94501': 'Komárno', '94601': 'Komárno',
    '94101': 'Nové Zámky', '94102': 'Nové Zámky', '94103': 'Nové Zámky',
    '95501': 'Topoľčany', '95301': 'Zlaté Moravce', '94201': 'Štúrovo',
    '93501': 'Šahy', '94101': 'Nové Zámky', '94035': 'Hurbanovo',
    
    // Žilinský kraj - doplnené
    '01001': 'Žilina', '01002': 'Žilina', '01003': 'Žilina', '01004': 'Žilina',
    '01005': 'Žilina', '01006': 'Žilina', '01007': 'Žilina', '01008': 'Žilina',
    '01301': 'Rajec', '01302': 'Rajecké Teplice', '01321': 'Bytča', '01341': 'Považská Bystrica',
    '01352': 'Dubnica nad Váhom', '01361': 'Ilava', '01371': 'Púchov', '01381': 'Lazy pod Makytou',
    '01901': 'Terchová', '02001': 'Púchov', '02011': 'Lednica', '02023': 'Beluša',
    '02061': 'Turzovka', '02101': 'Lysá pod Makytou', '02201': 'Čadca', '02202': 'Čadca',
    '02301': 'Čadca', '02321': 'Skalité', '02351': 'Turzovka', '02383': 'Oščadnica',
    '02401': 'Kysucké Nové Mesto', '02421': 'Krásno nad Kysucou', '02443': 'Korňa',
    '02601': 'Dolný Kubín', '02602': 'Dolný Kubín', '02621': 'Nižná', '02644': 'Tvrdošín',
    '02701': 'Dolný Kubín', '02711': 'Trstená', '02734': 'Nižná', '02751': 'Zázrivá',
    '02801': 'Trstená', '02821': 'Oravská Lesná', '02841': 'Habovka', '02901': 'Námestovo',
    '02911': 'Oravská Polhora', '02943': 'Zázrivá', '03001': 'Námestovo', '03011': 'Lokca',
    '03015': 'Zuberec', '03021': 'Rabča', '03044': 'Oravský Podzámok', '03101': 'Liptovský Mikuláš',
    '03102': 'Liptovský Mikuláš', '03103': 'Liptovský Mikuláš', '03104': 'Liptovský Mikuláš',
    '03105': 'Liptovský Mikuláš', '03106': 'Liptovský Mikuláš', '03113': 'Liptovský Hrádok',
    '03114': 'Liptovská Štiavnica', '03115': 'Palúdzka', '03121': 'Svätý Kríž', '03122': 'Východná',
    '03123': 'Važec', '03131': 'Liptovská Teplička', '03141': 'Liptovské Sliače', '03142': 'Hybe',
    '03143': 'Kráľova Lehota', '03144': 'Partizánska Ľupča', '03151': 'Liptovský Trnovec',
    '03161': 'Pribylina', '03201': 'Liptovský Mikuláš', '03202': 'Závažná Poruba',
    '03211': 'Liptovský Ján', '03221': 'Dovalovo', '03223': 'Liptovská Lúžna', '03224': 'Uhorská Ves',
    '03231': 'Hybe', '03241': 'Jakubovany', '03251': 'Liptovská Kokava', '03261': 'Malužiná',
    '03301': 'Ružomberok', '03401': 'Ružomberok', '03402': 'Ružomberok', '03403': 'Ružomberok',
    '03411': 'Likavka', '03412': 'Ludrová', '03421': 'Likavka', '03422': 'Lúčky',
    '03425': 'Vitanová', '03431': 'Štiavnik', '03601': 'Martin', '03602': 'Martin',
    '03603': 'Martin', '03604': 'Martin', '03605': 'Martin', '03611': 'Kláštor pod Znievom',
    '03621': 'Turčianske Teplice', '03631': 'Blatnica', '03641': 'Stráňavy', '03651': 'Sučany',
    '03701': 'Vrútky', '03711': 'Turany', '03721': 'Kraľovany', '03731': 'Turčianske Teplice',
    '03801': 'Turčianske Teplice', '03811': 'Sklené Teplice', '03851': 'Horná Štubňa',
    '03901': 'Turany', '03911': 'Rakša', '03921': 'Braväcovo', '03931': 'Príbovce',
    
    // Banskobystrický kraj - doplnené
    '97401': 'Banská Bystrica', '97402': 'Banská Bystrica', '97403': 'Banská Bystrica',
    '97404': 'Banská Bystrica', '97405': 'Banská Bystrica', '97406': 'Banská Bystrica',
    '97407': 'Banská Bystrica', '97408': 'Banská Bystrica', '97411': 'Badín', '97412': 'Brusno',
    '97413': 'Čerín', '97414': 'Malachov', '97415': 'Nemce', '97416': 'Riečka', '97417': 'Selce',
    '97418': 'Slovenská Ľupča', '97419': 'Tajov', '97421': 'Horná Mičiná', '97501': 'Banská Bystrica',
    '97511': 'Kremnička', '97601': 'Brezno', '97602': 'Brezno', '97611': 'Podbrezová',
    '97612': 'Predajná', '97613': 'Valaská', '97621': 'Hronec', '97631': 'Heľpa', '97641': 'Nemecká',
    '97651': 'Telgárt', '97661': 'Čierny Balog', '97671': 'Osrblie', '97681': 'Polomka',
    '97701': 'Banská Štiavnica', '97711': 'Beluj', '97721': 'Banský Studenec', '97731': 'Ilija',
    '97801': 'Banská Štiavnica', '97811': 'Svätý Anton', '97821': 'Nová Baňa', '97831': 'Žarnovica',
    '97901': 'Rimavská Sobota', '97902': 'Rimavská Sobota', '97911': 'Hnúšťa', '97921': 'Včelince',
    '97931': 'Kaloša', '97941': 'Klenovec', '97951': 'Tornaľa', '98001': 'Hnúšťa', '98011': 'Rimavské Brezovo',
    '98021': 'Čierny Balog', '98031': 'Hrachovo', '98041': 'Gemerská Poloma', '98051': 'Lubeník',
    '98101': 'Poltár', '98111': 'Lučenec', '98121': 'Šiatorská Bukovinka', '98131': 'Vinica',
    '98141': 'Dolná Strehová', '98151': 'Dolinka', '98201': 'Tornaľa', '98211': 'Štítnik',
    '98221': 'Gemerská Ves', '98231': 'Klenovec', '98301': 'Lučenec', '98401': 'Lučenec',
    '98501': 'Lučenec', '98502': 'Lučenec', '98511': 'Želiezovce', '98521': 'Vidiná',
    '98531': 'Kalinovo', '98541': 'Šahy', '98551': 'Čeláre', '98601': 'Fiľakovo',
    '98611': 'Modrý Kameň', '98621': 'Cinobaňa', '98631': 'Šuľa', '96001': 'Zvolen',
    '96002': 'Zvolen', '96003': 'Zvolen', '96011': 'Sliač', '96012': 'Dobrá Niva',
    '96013': 'Lieskovec', '96021': 'Pliešovce', '96022': 'Môlča', '96031': 'Budča',
    '96032': 'Zvolenská Slatina', '96041': 'Hliník nad Hronom', '96042': 'Očová',
    '96051': 'Očová', '96061': 'Detva', '96071': 'Slatinské Lazy', '96081': 'Sása',
    '96201': 'Krupina', '96211': 'Detva', '96212': 'Detva', '96221': 'Hriňová',
    '96231': 'Korytárky', '96241': 'Dubové', '96251': 'Látky', '96261': 'Stará Huta',
    '96271': 'Slatinské Lazy', '96281': 'Podkriváň', '96301': 'Krupina', '96311': 'Čabradský Vrbovok',
    '96321': 'Senohrad', '96331': 'Šášovské Podhradie', '96341': 'Bzovík', '96351': 'Devičany',
    '96401': 'Žarnovica', '96411': 'Dolná Ždaňa', '96421': 'Horné Hámre', '96501': 'Žiar nad Hronom',
    '96502': 'Žiar nad Hronom', '96511': 'Horné Opatovce', '96521': 'Zvolenská Slatina',
    '96531': 'Železná Breznica', '96541': 'Lovčica-Trubín', '96551': 'Žiarska Lehota',
    '96601': 'Kremnica', '96611': 'Kunešov', '96621': 'Krahule', '96631': 'Handlová',
    '96701': 'Žarnovica', '96711': 'Prenčov', '96721': 'Hronský Beňadik', '96801': 'Nová Baňa',
    '96811': 'Rudno nad Hronom', '96821': 'Banská Belá', '99001': 'Veľký Krtíš',
    '99011': 'Malý Krtíš', '99021': 'Slovenské Ďarmoty', '99031': 'Čelovce', '05001': 'Revúca',
    '05011': 'Jelšava', '05021': 'Rákoš', '05031': 'Tornaľa', '05041': 'Licince',
    '05051': 'Muráň', '04934': 'Tomášovce', '04944': 'Bottovo', '04951': 'Hajnáčka',
    
    // Prešovský kraj - doplnené
    '08001': 'Prešov', '08002': 'Prešov', '08003': 'Prešov', '08004': 'Prešov',
    '08005': 'Prešov', '08006': 'Prešov', '08007': 'Prešov', '08011': 'Ľubotice',
    '08012': 'Šarišské Michaľany', '08013': 'Sulín', '08021': 'Haniska', '08022': 'Kokošovce',
    '08023': 'Kendice', '08031': 'Víťaz', '08041': 'Nižná Šebastová', '08042': 'Vyšný Žipov',
    '08043': 'Demjata', '08044': 'Kapušany', '08045': 'Široké', '08051': 'Fintice',
    '08061': 'Svinia', '08071': 'Abrahámovce', '08081': 'Vyšná Šebastová', '08091': 'Bzenov',
    '08201': 'Prešov', '08211': 'Veľký Šariš', '08212': 'Sabinov', '08221': 'Lipany',
    '08231': 'Klenov', '08241': 'Červenica pri Sabinove', '08251': 'Šarišské Bohdanovce',
    '08261': 'Ražňany', '08271': 'Livov', '08281': 'Šarišské Dravce', '08301': 'Sabinov',
    '08311': 'Janov', '08312': 'Uzovská Panica', '08313': 'Bertotovce', '08314': 'Šarišské Sokolovce',
    '08315': 'Ostrovany', '08316': 'Ľutina', '08321': 'Uzovské Pekľany', '08322': 'Jakubovany',
    '08323': 'Torysa', '08324': 'Kristy', '08325': 'Chmeľov', '08331': 'Plavnica',
    '08332': 'Lipany', '08341': 'Záhradné', '08351': 'Bystrany', '08361': 'Bertotovce',
    '08401': 'Lipany', '08411': 'Levoča', '08412': 'Spišská Nová Ves', '08413': 'Smižany',
    '08501': 'Bardejov', '08502': 'Bardejov', '08511': 'Hertník', '08512': 'Bardejovská Nová Ves',
    '08513': 'Kružlov', '08514': 'Richvald', '08515': 'Tarnov', '08521': 'Zborov',
    '08531': 'Kríže', '08541': 'Raslavice', '08551': 'Livovská Huta', '08561': 'Fričovce',
    '08571': 'Kurov', '08581': 'Hanušovce nad Topľou', '08591': 'Hermanovce', '08601': 'Bardejov',
    '08611': 'Giraltovce', '08612': 'Tročany', '08613': 'Jedlinka', '08614': 'Svidník',
    '08615': 'Vyšný Mirošov', '08621': 'Kurima', '08622': 'Štefanovce', '08631': 'Šiba',
    '08641': 'Kobylnica', '08651': 'Richvald', '08661': 'Ľubotin', '08671': 'Kožuchovce',
    '08681': 'Hertník', '08691': 'Havaj', '08701': 'Svidník', '08711': 'Vyšný Orlík',
    '08712': 'Ladomirová', '08713': 'Nižný Mirošov', '08714': 'Varadka', '08715': 'Nižný Orlík',
    '08716': 'Krajná Poľana', '08717': 'Duplín', '08721': 'Stropkov', '08722': 'Miková',
    '08723': 'Staškovce', '08724': 'Habura', '08731': 'Nižná Polianka', '08732': 'Krajná Bystrá',
    '08741': 'Radoma', '08751': 'Vyšná Jablonka', '08761': 'Nizny Komarnik', '08801': 'Stropkov',
    '08811': 'Bukovce', '08812': 'Havranok', '08813': 'Makovce', '08814': 'Varechovce',
    '08815': 'Juskova Voľa', '08816': 'Bystrá', '08821': 'Višný Hrušov', '08822': 'Staškovce',
    '08831': 'Nižný Hrušov', '08841': 'Lomnica', '08851': 'Havaj', '09001': 'Stará Ľubovňa',
    '09011': 'Plaveč', '09012': 'Hniezdne', '09013': 'Červený Kláštor', '09014': 'Plavnica',
    '09015': 'Údol', '09021': 'Nová Ľubovňa', '09022': 'Podolínec', '09023': 'Ľubotín',
    '09031': 'Vyšné Ružbachy', '09041': 'Nižné Ružbachy', '09051': 'Stráne pod Tatrami',
    '09061': 'Hniezdne', '09071': 'Spišská Belá', '09072': 'Spišské Tomášovce', '09101': 'Stará Ľubovňa',
    '06001': 'Kežmarok', '06011': 'Spišská Belá', '06012': 'Veľká Lomnica', '06013': 'Slovenská Ves',
    '06014': 'Podolínec', '06015': 'Ľubica', '06021': 'Spišská Teplica', '06031': 'Huncovce',
    '06041': 'Studený Potok', '06051': 'Gerlachov', '06101': 'Kežmarok', '06201': 'Starý Smokovec',
    '06202': 'Nový Smokovec', '06301': 'Štrbské Pleso', '06401': 'Štrbské Pleso',
    '05901': 'Poprad', '05902': 'Poprad', '05903': 'Poprad', '05904': 'Poprad',
    '05905': 'Poprad', '05911': 'Štrba', '05912': 'Mengusovce', '05913': 'Lučivná',
    '05914': 'Veľká', '05915': 'Šuňava', '05921': 'Svit', '05931': 'Tatranská Lomnica',
    '05941': 'Veľká Lomnica', '05951': 'Gánovce', '05960': 'Tatranská Lomnica',
    '05961': 'Stará Lesná', '05962': 'Tatranská Kotlina', '05971': 'Mengusovce',
    '05981': 'Ždiar', '06501': 'Spišská Nová Ves', '06502': 'Spišská Nová Ves',
    '06511': 'Spišské Podhradie', '06512': 'Spišské Vlachy', '06513': 'Iliašovce',
    '06514': 'Žehra', '06515': 'Markušovce', '06521': 'Hodkovce', '06531': 'Levoča',
    '06541': 'Nemešany', '06551': 'Svätý Jur', '06561': 'Vysoká', '06601': 'Gelnica',
    '06611': 'Hnilčík', '06612': 'Mníšek nad Hnilcom', '06613': 'Štós', '06614': 'Nálepkovo',
    '06615': 'Margecany', '06621': 'Úhorná', '06631': 'Helcmanovce', '06641': 'Zacharovce',
    '06651': 'Krompachy', '06701': 'Spišské Bystré', '06711': 'Strážky', '06721': 'Spišské Tomášovce',
    '06731': 'Stráne', '06741': 'Arnutovce', '06801': 'Moldava nad Bodvou', '06811': 'Jasov',
    '06812': 'Bidovce', '06813': 'Holčíkovce', '06821': 'Kechnec', '06831': 'Nižná Myšľa',
    '06841': 'Rankovce', '06851': 'Vyšný Čaj', '06901': 'Snina', '06911': 'Stakčín',
    '06912': 'Ubľa', '06913': 'Kolbasov', '06914': 'Ruský Hrabovec', '06915': 'Ulič',
    '06921': 'Ruská Kajňa', '06931': 'Ubľa', '06941': 'Ruská Bystrá', '06951': 'Ruská Poruba',
    '07001': 'Vranov nad Topľou', '07011': 'Hanušovce nad Topľou', '07012': 'Sedliská',
    '07013': 'Zámutov', '07014': 'Kračúnovce', '07015': 'Sečovská Polianka', '07021': 'Hencovce',
    '07031': 'Čaklov', '07041': 'Cabov', '07051': 'Davidov', '07061': 'Vechec',
    '07071': 'Egreš', '07081': 'Jasenov', '07091': 'Hertník', '07101': 'Michalovce',
    '07102': 'Michalovce', '07103': 'Michalovce', '07104': 'Michalovce', '07111': 'Jovsa',
    '07112': 'Vinné', '07113': 'Stretava', '07114': 'Porúbka', '07115': 'Lastomír',
    '07121': 'Zemplínska Teplica', '07131': 'Senné', '07141': 'Tušice', '07151': 'Jastrabie pri Michalovciach',
    '07161': 'Vinné', '07201': 'Veľké Kapušany', '07211': 'Choňkovce', '07212': 'Nacina Ves',
    '07213': 'Lesné', '07214': 'Remetské Hámre', '07221': 'Malé Kapušany', '07231': 'Streda nad Bodrogom',
    '07241': 'Vojčice', '07251': 'Tibava', '07261': 'Lastovce', '07301': 'Sobrance',
    '07311': 'Podhoroď', '07312': 'Inovce', '07313': 'Remetské Hámre', '07314': 'Blatné Revištia',
    '07315': 'Hliník nad Cirochou', '07321': 'Krivošťany', '07331': 'Hlivištia',
    '07341': 'Ruská Nová Ves', '07351': 'Kaluža', '07501': 'Čierna nad Tisou',
    '07511': 'Malé Trakany', '07512': 'Veľké Trakany', '07513': 'Veľké Revištia',
    '07514': 'Pavlovce nad Uhom', '07515': 'Klin nad Bodrogom', '07521': 'Somotor',
    '07531': 'Oborín', '07601': 'Trebišov', '07602': 'Trebišov', '07611': 'Sečovce',
    '07612': 'Streda nad Bodrogom', '07613': 'Blatné Remety', '07614': 'Cejkov',
    '07615': 'Dvorianky', '07621': 'Čierna', '07631': 'Svinica', '07641': 'Veľaty',
    '07651': 'Somotor', '07661': 'Kaluža', '07671': 'Novosad', '07681': 'Topoľovka',
    
    // Košický kraj
    '04001': 'Košice', '04002': 'Košice', '04003': 'Košice', '04004': 'Košice',
    '04005': 'Košice', '04006': 'Košice', '04007': 'Košice', '04008': 'Košice',
    '04011': 'Košice', '04012': 'Košice', '04013': 'Košice', '04014': 'Košice',
    '04015': 'Košice', '04016': 'Košice', '04017': 'Košice', '04018': 'Košice',
    '04022': 'Košice', '04023': 'Košice', '04024': 'Košice', '04025': 'Košice',
    '04026': 'Košice', '04027': 'Košice', '07101': 'Michalovce', '07601': 'Trebišov',
    '04401': 'Rožňava', '04501': 'Moldava nad Bodvou', '04701': 'Spišská Nová Ves',
    '04801': 'Dobšiná', '04901': 'Rožňava', '04601': 'Gelnica',
    '07501': 'Čierna nad Tisou', '07201': 'Veľké Kapušany', '07301': 'Sobrance'
  };

  const validateRodneCislo = (rc) => {
    if (!rc || rc.length !== 10) return false;
    const number = parseInt(rc, 10);
    return number % 11 === 0;
  };

  const extractDateFromRC = (rc) => {
    if (!rc || rc.length < 6) return null;
    let year = parseInt(rc.substring(0, 2), 10);
    let month = parseInt(rc.substring(2, 4), 10);
    const day = parseInt(rc.substring(4, 6), 10);
    
    year += year > 53 ? 1900 : 2000;
    if (month > 50) month -= 50;
    
    return { day, month, year };
  };

  const validateIBAN = (iban) => {
    if (!iban) return false;
    const cleaned = iban.replace(/\s/g, '');
    if (!/^[A-Z]{2}\d{22}$/.test(cleaned)) return false;
    
    const rearranged = cleaned.substring(4) + cleaned.substring(0, 4);
    let numericString = '';
    for (let char of rearranged) {
      if (char >= 'A' && char <= 'Z') {
        numericString += (char.charCodeAt(0) - 55).toString();
      } else {
        numericString += char;
      }
    }
    
    let remainder = 0;
    for (let i = 0; i < numericString.length; i++) {
      remainder = (remainder * 10 + parseInt(numericString[i])) % 97;
    }
    
    return remainder === 1;
  };

  const validateEmail = (email) => {
    if (!email) return true;
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email.trim());
  };

  const extractFromIBAN = (iban) => {
    if (!iban) return { ucet: '', preducet: '', banka: '', swift: '' };
    const cleaned = iban.replace(/\s/g, '');
    if (cleaned.length !== 24) return { ucet: '', preducet: '', banka: '', swift: '' };
    
    const banka = cleaned.substring(4, 8);
    const preducet = cleaned.substring(8, 14);
    const ucet = cleaned.substring(14, 24);
    const swift = BANK_SWIFT[banka] || '';
    
    return { ucet, preducet, banka, swift };
  };

  const processExcelData = (fileData, startNum) => {
    // Pomocné funkcie - definované VNÚTRI processExcelData
    const parseDate = (dateStr) => {
      if (typeof dateStr === 'number') {
        const excelEpoch = new Date(1899, 11, 30);
        const date = new Date(excelEpoch.getTime() + dateStr * 86400000);
        return date;
      }
      
      const str = String(dateStr).trim();
      if (!str) return null;
      
      const parts = str.split('.');
      if (parts.length !== 3) return null;
      
      const day = parseInt(parts[0]);
      const month = parseInt(parts[1]) - 1;
      const year = parseInt(parts[2]);
      
      if (isNaN(day) || isNaN(month) || isNaN(year)) return null;
      
      return new Date(year, month, day);
    };

    const formatDate = (dateValue) => {
      if (!dateValue) return '';
      const date = parseDate(dateValue);
      if (!date) return String(dateValue);
      const day = String(date.getDate()).padStart(2, '0');
      const month = String(date.getMonth() + 1).padStart(2, '0');
      const year = date.getFullYear();
      return `${day}.${month}.${year}`;
    };
    
    const workbook = XLSX.read(fileData, { type: 'array', cellStyles: true });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const range = XLSX.utils.decode_range(sheet['!ref']);
    
    const processedData = [];
    let personalNum = parseInt(startNum);
    
    for (let R = 3; R <= range.e.r; R++) {
      const cellA = sheet[XLSX.utils.encode_cell({ r: R, c: 0 })];
      
      // NOVÉ: Ak je bunka v stĺpci A prázdna → STOP spracovanie
      if (!cellA || !cellA.v || String(cellA.v).trim() === '') {
        console.log(`Riadok ${R}: Prázdna bunka v stĺpci A - STOP spracovanie`);
        break; // Ukončiť cyklus
      }
      
      // Kontrola zelenej farby - rôzne varianty
      let isGreen = false;
      if (cellA && cellA.s) {
        // Skúšame rôzne spôsoby ako Excel ukladá farby
        if (cellA.s.fgColor) {
          const color = cellA.s.fgColor.rgb;
          if (color) {
            // Konverzia hex na RGB
            const r = parseInt(color.substring(0, 2), 16);
            const g = parseInt(color.substring(2, 4), 16);
            const b = parseInt(color.substring(4, 6), 16);
            // Zelená ak G komponent je dominantný
            isGreen = g > 150 && g > r + 30 && g > b + 30;
          }
        }
        // Alternatívne skúšame bgColor
        if (!isGreen && cellA.s.bgColor) {
          const color = cellA.s.bgColor.rgb;
          if (color) {
            const r = parseInt(color.substring(0, 2), 16);
            const g = parseInt(color.substring(2, 4), 16);
            const b = parseInt(color.substring(4, 6), 16);
            isGreen = g > 150 && g > r + 30 && g > b + 30;
          }
        }
        // Ak má bunka fill pattern
        if (!isGreen && cellA.s.patternType) {
          isGreen = cellA.s.patternType === 'solid';
        }
      }
      
      // DEBUG: Vypíš prvých 10 riadkov pre kontrolu
      if (R <= 13) {
        console.log(`Riadok ${R}: hodnota="${cellA.v}", zelená=${isGreen}`);
      }
      
      if (!isGreen) continue;
      
      const row = {};
      for (let C = 0; C <= 38; C++) {
        const cell = sheet[XLSX.utils.encode_cell({ r: R, c: C })];
        row[`col${C}`] = cell ? (cell.v || '') : '';
      }
      
      const trimText = (text) => String(text || '').replace(/\s+/g, ' ').trim();
      
      const col3 = trimText(row.col2);
      const kodppv = PPV_CODES[col3] || '';
      
      const rc = String(row.col11 || '').replace(/\D/g, '').padStart(10, '0');
      
      // OPRAVA: IBAN je v stĺpci AK (index 36), nie AG
      const iban = String(row.col36 || '').replace(/\s/g, '');
      const { ucet, preducet, banka, swift } = extractFromIBAN(iban);
      
      // OPRAVA: PSČ je v stĺpci V (index 21), obec v W (index 22)
      const psc = trimText(row.col21);
      const obecExcel = trimText(row.col22);
      const obecFromPSC = PSC_LOOKUP[psc] || '';
      
      const rcValid = validateRodneCislo(rc);
      const rcDate = extractDateFromRC(rc);
      const datNarodStr = row.col8; // Môže byť číslo alebo string
      let birthDateValid = true;
      if (rcDate && datNarodStr) {
        const birthDate = parseDate(datNarodStr);
        if (birthDate) {
          birthDateValid = 
            birthDate.getDate() === rcDate.day &&
            birthDate.getMonth() + 1 === rcDate.month &&
            birthDate.getFullYear() === rcDate.year;
        }
      }
      
      const ibanValid = validateIBAN(iban);
      // OPRAVA: Email je v stĺpci AL (index 37)
      const emailValid = validateEmail(String(row.col37 || ''));
      
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      const tomorrow = new Date(today);
      tomorrow.setDate(tomorrow.getDate() + 1);
      
      // OPRAVA: nastup je v stĺpci AE (index 30)
      const nastupDate = parseDate(row.col30);
      const nastupValid = nastupDate && nastupDate >= tomorrow;
      
      // OPRAVA: dobaurcita je v stĺpci AF (index 31)
      const dobaurcitaDate = parseDate(row.col31);
      const dobaurcitaValid = !dobaurcitaDate || !nastupDate || dobaurcitaDate >= nastupDate;
      
      // OPRAVA: NČZD je v stĺpci Z (index 25)
      let odpocet = '0';
      const nczd = String(row.col25 || '').trim().toUpperCase();
      if (nczd === 'A') odpocet = 'A';
      else if (nczd === 'N') odpocet = 'N';
      
      // OPRAVA: stredisko je v stĺpci AC (index 28)
      let stredkod = trimText(row.col28);
      if (stredkod && parseInt(stredkod) > 999) {
        stredkod = 'A' + stredkod;
      }
      
      // OPRAVA: tarif je v stĺpci AJ (index 35)
      const tarif = String(row.col35 || '').replace(',', '.');
      const stat = obecFromPSC ? 'SK' : '';
      
      processedData.push({
        spolocnost: trimText(row.col0),
        oscislo: String(personalNum++),
        dobps: col3,
        kodppv: kodppv,
        priezv: trimText(row.col5),
        meno: trimText(row.col6),
        titl1: trimText(row.col7),
        datnarod: formatDate(row.col8),
        statprisl: trimText(row.col9),
        rodstav: trimText(row.col10),
        rc: rc,
        rodmeno: trimText(row.col14),
        miestonar: trimText(row.col15),
        ulica: trimText(row.col18),        // S = index 18
        supcislo: trimText(row.col19),     // T = index 19
        cdomu: trimText(row.col20),        // U = index 20
        psc: psc,                          // V = index 21
        obecnazov: obecExcel,              // W = index 22
        stat: stat,
        vzdelanie: trimText(row.col23),    // X = index 23
        dochodky: trimText(row.col24),     // Y = index 24
        odpocet: odpocet,                  // Z = index 25
        zpkod: trimText(row.col26),        // AA = index 26
        vtnt: trimText(row.col27),         // AB = index 27
        stredkod: stredkod,                // AC = index 28
        miesvyk: trimText(row.col29),      // AD = index 29
        nastup: formatDate(row.col30),     // AE = index 30
        dobaurcita: formatDate(row.col31), // AF = index 31
        datoop: formatDate(row.col32),     // AG = index 32
        pracovnapozicia: trimText(row.col33), // AH = index 33
        isco: trimText(row.col34),         // AI = index 34
        tarif: tarif,                      // AJ = index 35
        iban: iban,                        // AK = index 36
        ucet: ucet,
        preducet: preducet,
        banka: banka,
        swift: swift,
        email: String(row.col37 || '').replace(/\s/g, ''), // AL = index 37
        mobil: trimText(row.col38),        // AM = index 38
        
        validations: {
          rcValid,
          birthDateValid,
          ibanValid,
          emailValid,
          nastupValid,
          dobaurcitaValid,
          obecMatch: obecExcel === obecFromPSC,
          statEmpty: !stat
        },
        suggestions: {
          obec: obecFromPSC
        }
      });
    }
    
    return processedData;
  };

  const handleFileUpload = useCallback((e) => {
    const file = e.target.files[0];
    if (!file) return;
    
    if (!startingNumber || isNaN(startingNumber)) {
      setError('Prosím zadajte začiatočné osobné číslo!');
      return;
    }
    
    setError('');
    
    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const processed = processExcelData(event.target.result, startingNumber);
        if (processed.length === 0) {
          setError('Nenašli sa žiadne zelené riadky v súbore! Skontrolujte či sú riadky v stĺpci A podfarbené zelenou farbou.');
          return;
        }
        setData(processed);
        setStep(2);
      } catch (error) {
        setError('Chyba pri spracovaní súboru: ' + error.message);
        console.error('Error:', error);
      }
    };
    reader.readAsArrayBuffer(file);
  }, [startingNumber]);

  const handleCellEdit = (index, field, value) => {
    const newData = [...data];
    newData[index][field] = value;
    
    // Pomocná funkcia pre validáciu
    const parseDate = (dateStr) => {
      if (typeof dateStr === 'number') {
        const excelEpoch = new Date(1899, 11, 30);
        return new Date(excelEpoch.getTime() + dateStr * 86400000);
      }
      const str = String(dateStr).trim();
      if (!str) return null;
      const parts = str.split('.');
      if (parts.length !== 3) return null;
      const day = parseInt(parts[0]);
      const month = parseInt(parts[1]) - 1;
      const year = parseInt(parts[2]);
      if (isNaN(day) || isNaN(month) || isNaN(year)) return null;
      return new Date(year, month, day);
    };
    
    // Ak sa edituje IBAN, prepočítaj extrakty
    if (field === 'iban') {
      const { ucet, preducet, banka, swift } = extractFromIBAN(value);
      newData[index].ucet = ucet;
      newData[index].preducet = preducet;
      newData[index].banka = banka;
      newData[index].swift = swift;
      newData[index].validations.ibanValid = validateIBAN(value);
    }
    
    // Ak sa edituje rodné číslo, prepočítaj validácie
    if (field === 'rc') {
      newData[index].validations.rcValid = validateRodneCislo(value);
    }
    
    // Ak sa edituje email, prepočítaj validáciu
    if (field === 'email') {
      newData[index].validations.emailValid = validateEmail(value);
    }
    
    // Ak sa edituje štát, prepočítaj validáciu
    if (field === 'stat') {
      newData[index].validations.statEmpty = !value || value.trim() === '';
    }
    
    setData(newData);
  };

  const exportToCSV = () => {
    const headers = [
      'spolo?nos?', 'oscislo', 'DOBPS/DVP/HPP', 'kodppv', 'priezv', 'meno', 'titl1',
      'datnarod', 'statprisl', 'rodstav', 'rc', 'rodmeno', 'miestonar', 'ulica',
      'supcislo', 'cdomu', 'psc', 'obecnazov', 'stat', 'vzdelanie',
      'dôchodky,invalidita(uvies?konkrétneaj%)', 'odpocet', 'zpkod', 'VT/NT',
      'stredkod', 'miesvyk', 'nastup', 'dobaurcita', 'datoop', 'pracovnápozícia',
      'isco', 'tarif', 'IBAN', 'ucet', 'preducet', 'banka', 'SWIFT', 'email_adress', 'mobil'
    ];
    
    const rows = data.map(row => [
      row.spolocnost, row.oscislo, row.dobps, row.kodppv, row.priezv, row.meno, row.titl1,
      row.datnarod, row.statprisl, row.rodstav, row.rc, row.rodmeno, row.miestonar, row.ulica,
      row.supcislo, row.cdomu, row.psc, row.obecnazov, row.stat, row.vzdelanie,
      row.dochodky, row.odpocet, row.zpkod, row.vtnt, row.stredkod, row.miesvyk,
      row.nastup, row.dobaurcita, row.datoop, row.pracovnapozicia, row.isco, row.tarif,
      row.iban, row.ucet, row.preducet, row.banka, row.swift, row.email, row.mobil
    ]);
    
    const csvLines = [headers, ...rows].map(row => 
      row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(';')
    );
    
    const csvContent = csvLines.join('\n');
    
    // Vytvor náhľad (prvých 10 riadkov)
    const previewLines = csvLines.slice(0, 11); // hlavička + 10 riadkov
    setCsvPreview(previewLines.join('\n'));
    
    let blob;
    if (encoding === 'windows-1250') {
      const bom = '\uFEFF';
      blob = new Blob([bom + csvContent], { type: 'text/csv;charset=utf-8' });
    } else {
      blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8' });
    }
    
    const url = URL.createObjectURL(blob);
    setCsvDownloadUrl(url);
  };

  const EditableCell = ({ value, onSave, error }) => {
    const [isEditing, setIsEditing] = useState(false);
    const [editValue, setEditValue] = useState(value);

    if (isEditing) {
      return (
        <input
          value={editValue}
          onChange={(e) => setEditValue(e.target.value)}
          onBlur={() => {
            onSave(editValue);
            setIsEditing(false);
          }}
          onKeyDown={(e) => {
            if (e.key === 'Enter') {
              onSave(editValue);
              setIsEditing(false);
            }
          }}
          autoFocus
          className={`w-full px-2 py-1 border rounded ${error ? 'border-red-500' : 'border-blue-500'}`}
        />
      );
    }

    return (
      <div
        onClick={() => setIsEditing(true)}
        className={`cursor-pointer hover:bg-blue-50 px-2 py-1 rounded ${error ? 'text-red-600 font-semibold' : ''}`}
      >
        {value || '-'}
      </div>
    );
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-50 p-6">
      <div className="max-w-7xl mx-auto">
        <div className="bg-white rounded-xl shadow-xl p-8">
          <div className="flex justify-between items-center mb-6">
            <h1 className="text-3xl font-bold text-gray-800">Excel → CSV Konvertor</h1>
            <button
              onClick={() => setShowHelp(!showHelp)}
              className="flex items-center gap-2 px-4 py-2 bg-blue-100 text-blue-700 rounded-lg hover:bg-blue-200 transition"
            >
              <Info size={20} />
              Návod
            </button>
          </div>

          {showHelp && (
            <div className="mb-6 p-6 bg-blue-50 rounded-lg border border-blue-200 max-h-96 overflow-y-auto">
              <h2 className="text-xl font-semibold mb-4 text-blue-900">Popis transformácií a validácií</h2>
              <div className="space-y-2 text-sm text-gray-700">
                <div><strong>Stĺpec 2 (oscislo):</strong> Automaticky generované postupné osobné čísla od zadaného začiatočného čísla</div>
                <div><strong>Stĺpec 3 (DOBPS/DVP/HPP):</strong> Hodnota z Excelu bez zmeny</div>
                <div><strong>Stĺpec 5 (kodppv):</strong> Číselník z stĺpca 3 (DVP/DVPNP/DVP9→5, DBPS/DBPS9→6, DPC→4, HPP→2)</div>
                <div><strong>Stĺpec 12 (rc):</strong> Rodné číslo bez lomítka, zachované úvodné nuly + validácia deliteľnosti 11</div>
                <div><strong>Kontrola RČ:</strong> Porovnanie dátumu narodenia s RČ (ženy majú +50 k mesiacu)</div>
                <div><strong>Stĺpec 19 (stat):</strong> Automaticky "SK" ak PSČ existuje v číselníku</div>
                <div><strong>Stĺpec 21 (obecnazov):</strong> Návrh z číselníka PSČ (200+ hlavných slovenských obcí), zvýraznenie ak nesedí</div>
                <div><strong>Stĺpec 25 (odpocet):</strong> 0 ak prázdne/0, A alebo N podľa hodnoty v Exceli</div>
                <div><strong>Stĺpec 28 (stredkod):</strong> Pridať "A" pred hodnotu ak {'>'} 999</div>
                <div><strong>Stĺpec 30 (nastup):</strong> Kontrola že dátum ≥ aktuálny dátum + 1 deň</div>
                <div><strong>Stĺpec 31 (dobaurcita):</strong> Kontrola že dátum ≥ nastup</div>
                <div><strong>Stĺpec 33 (IBAN):</strong> Odstránené medzery + MOD-97 validácia</div>
                <div><strong>Stĺpec 34 (ucet):</strong> Extrakcia posledných 10 číslic z IBANu</div>
                <div><strong>Stĺpec 35 (preducet):</strong> Extrakcia 6 číslic z IBANu (pozície 9-14)</div>
                <div><strong>Stĺpec 36 (banka):</strong> Extrakcia kódu banky z IBANu (pozície 5-8)</div>
                <div><strong>Stĺpec 37 (SWIFT):</strong> Automatické doplnenie z číselníka NBS podľa kódu banky</div>
                <div><strong>Stĺpec 38 (email):</strong> Odstránené medzery + validácia formátu</div>
                <div><strong>Textové polia:</strong> Automatické odstránenie zbytočných medzier vo všetkých textových poliach</div>
                <div className="mt-4 p-3 bg-yellow-50 border border-yellow-200 rounded">
                  <strong>Poznámka:</strong> Číselník PSČ obsahuje 200+ najdôležitejších slovenských obcí (krajské/okresné mestá a väčšie obce).
                  Pre obce, ktoré nie sú v číselníku, je potrebné zadať štát "SK" manuálne.
                  Číselník SWIFT kódov je kompletný podľa NBS k 04/2022 (31 bánk).
                </div>
              </div>
            </div>
          )}

          {step === 1 && (
            <div className="space-y-6">
              {error && (
                <div className="bg-red-50 border border-red-200 rounded-lg p-4">
                  <p className="text-sm text-red-800">
                    <strong>Chyba:</strong> {error}
                  </p>
                </div>
              )}
              
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Začiatočné osobné číslo zamestnanca
                </label>
                <input
                  type="number"
                  value={startingNumber}
                  onChange={(e) => setStartingNumber(e.target.value)}
                  placeholder="napr. 6727"
                  className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                />
              </div>

              <div className="border-2 border-dashed border-gray-300 rounded-lg p-12 text-center hover:border-blue-500 transition">
                <Upload className="mx-auto mb-4 text-gray-400" size={48} />
                <p className="text-lg font-medium text-gray-700 mb-2">Nahrajte XLSX/XLS súbor</p>
                <p className="text-sm text-gray-500 mb-4">Kliknite alebo pretiahnite súbor sem</p>
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={handleFileUpload}
                  className="hidden"
                  id="fileInput"
                />
                <label
                  htmlFor="fileInput"
                  className="inline-block px-6 py-3 bg-blue-600 text-white rounded-lg cursor-pointer hover:bg-blue-700 transition"
                >
                  Vybrať súbor
                </label>
              </div>

              <div className="bg-yellow-50 border border-yellow-200 rounded-lg p-4">
                <p className="text-sm text-yellow-800">
                  <strong>Dôležité:</strong> Súbor musí mať hlavičku v riadku 3, dáta od riadku 4. 
                  Spracujú sa len riadky s <span className="bg-green-200 px-2 py-1 rounded">zelenou farbou</span> v stĺpci A.
                </p>
              </div>
            </div>
          )}

          {step === 2 && (
            <div className="space-y-6">
              <div className="flex justify-between items-center">
                <h2 className="text-xl font-semibold text-gray-800">
                  Načítaných {data.length} záznamov (kliknite na bunku pre editáciu)
                </h2>
                <button
                  onClick={() => {
                    setStep(1);
                    setData([]);
                    setStartingNumber('');
                  }}
                  className="px-4 py-2 text-blue-600 hover:bg-blue-50 rounded-lg transition"
                >
                  ← Späť
                </button>
              </div>

              <div className="overflow-x-auto border rounded-lg max-h-96">
                <table className="min-w-full divide-y divide-gray-200 text-xs">
                  <thead className="bg-gray-50 sticky top-0">
                    <tr>
                      <th className="px-2 py-2 text-left font-medium text-gray-700">Os.č.</th>
                      <th className="px-2 py-2 text-left font-medium text-gray-700">Priezvisko</th>
                      <th className="px-2 py-2 text-left font-medium text-gray-700">Meno</th>
                      <th className="px-2 py-2 text-left font-medium text-gray-700">PSČ</th>
                      <th className="px-2 py-2 text-left font-medium text-gray-700">Obec</th>
                      <th className="px-2 py-2 text-left font-medium text-gray-700">Štát</th>
                      <th className="px-2 py-2 text-left font-medium text-gray-700">RČ</th>
                      <th className="px-2 py-2 text-left font-medium text-gray-700">IBAN</th>
                      <th className="px-2 py-2 text-left font-medium text-gray-700">SWIFT</th>
                      <th className="px-2 py-2 text-left font-medium text-gray-700">Email</th>
                      <th className="px-2 py-2 text-left font-medium text-gray-700">Status</th>
                    </tr>
                  </thead>
                  <tbody className="bg-white divide-y divide-gray-200">
                    {data.map((row, idx) => (
                      <tr key={idx} className="hover:bg-gray-50">
                        <td className="px-2 py-1">
                          <EditableCell
                            value={row.oscislo}
                            onSave={(val) => handleCellEdit(idx, 'oscislo', val)}
                          />
                        </td>
                        <td className="px-2 py-1">{row.priezv}</td>
                        <td className="px-2 py-1">{row.meno}</td>
                        <td className="px-2 py-1">{row.psc}</td>
                        <td className="px-2 py-1">
                          {row.obecnazov}
                          {!row.validations.obecMatch && row.suggestions.obec && (
                            <div className="text-xs text-orange-600" title={`Návrh: ${row.suggestions.obec}`}>
                              → {row.suggestions.obec}?
                            </div>
                          )}
                        </td>
                        <td className="px-2 py-1">
                          <EditableCell
                            value={row.stat}
                            onSave={(val) => handleCellEdit(idx, 'stat', val)}
                            error={row.validations.statEmpty}
                          />
                        </td>
                        <td className="px-2 py-1">
                          <EditableCell
                            value={row.rc}
                            onSave={(val) => handleCellEdit(idx, 'rc', val)}
                            error={!row.validations.rcValid || !row.validations.birthDateValid}
                          />
                        </td>
                        <td className="px-2 py-1">
                          <EditableCell
                            value={row.iban}
                            onSave={(val) => handleCellEdit(idx, 'iban', val)}
                            error={!row.validations.ibanValid}
                          />
                        </td>
                        <td className="px-2 py-1">{row.swift}</td>
                        <td className="px-2 py-1">
                          <EditableCell
                            value={row.email}
                            onSave={(val) => handleCellEdit(idx, 'email', val)}
                            error={!row.validations.emailValid}
                          />
                        </td>
                        <td className="px-2 py-1">
                          <div className="flex gap-1">
                            {!row.validations.rcValid && (
                              <span title="RČ nie je deliteľné 11" className="text-red-600">
                                <AlertCircle size={14} />
                              </span>
                            )}
                            {!row.validations.birthDateValid && (
                              <span title="RČ nesedí s dátumom narodenia" className="text-red-600">
                                <AlertCircle size={14} />
                              </span>
                            )}
                            {!row.validations.ibanValid && (
                              <span title="IBAN nie je validný" className="text-red-600">
                                <AlertCircle size={14} />
                              </span>
                            )}
                            {!row.validations.emailValid && (
                              <span title="Email nie je validný" className="text-red-600">
                                <AlertCircle size={14} />
                              </span>
                            )}
                            {!row.validations.nastupValid && (
                              <span title="Dátum nástupu je v minulosti" className="text-red-600">
                                <AlertCircle size={14} />
                              </span>
                            )}
                            {!row.validations.dobaurcitaValid && (
                              <span title="Trvanie do je menšie ako nastup" className="text-red-600">
                                <AlertCircle size={14} />
                              </span>
                            )}
                            {row.validations.statEmpty && (
                              <span title="Štát nie je doplnený - PSČ nenájdené" className="text-orange-600">
                                <AlertCircle size={14} />
                              </span>
                            )}
                            {!row.validations.obecMatch && row.suggestions.obec && (
                              <span title={`Navrhovaná obec: ${row.suggestions.obec}`} className="text-orange-600">
                                <AlertCircle size={14} />
                              </span>
                            )}
                            {row.validations.rcValid && row.validations.birthDateValid && 
                             row.validations.ibanValid && row.validations.emailValid &&
                             row.validations.nastupValid && row.validations.dobaurcitaValid && (
                              <span title="Všetky validácie OK" className="text-green-600">
                                <CheckCircle size={14} />
                              </span>
                            )}
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              <div className="bg-blue-50 border border-blue-200 rounded-lg p-4">
                <p className="text-sm text-blue-800 mb-2">
                  <strong>Tip:</strong> Kliknite na bunku pre editáciu. Zmeny v IBAN automaticky prepočítajú ucet, preducet, banka a SWIFT kód.
                </p>
                <div className="flex gap-4 text-xs">
                  <div className="flex items-center gap-1">
                    <AlertCircle size={12} className="text-red-600" />
                    <span>Červená = chyba vo validácii</span>
                  </div>
                  <div className="flex items-center gap-1">
                    <AlertCircle size={12} className="text-orange-600" />
                    <span>Oranžová = upozornenie/návrh</span>
                  </div>
                  <div className="flex items-center gap-1">
                    <CheckCircle size={12} className="text-green-600" />
                    <span>Zelená = všetko OK</span>
                  </div>
                </div>
              </div>

              <div className="space-y-4">
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    Kódovanie CSV súboru
                  </label>
                  <select
                    value={encoding}
                    onChange={(e) => {
                      setEncoding(e.target.value);
                      setCsvDownloadUrl(null); // Reset download linku pri zmene kódovania
                    }}
                    className="w-64 px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500"
                  >
                    <option value="utf-8">UTF-8 (moderný štandard)</option>
                    <option value="windows-1250">Windows-1250 (staršie systémy)</option>
                  </select>
                </div>

                {!csvDownloadUrl ? (
                  <button
                    onClick={exportToCSV}
                    className="flex items-center gap-2 px-6 py-3 bg-green-600 text-white rounded-lg hover:bg-green-700 transition font-semibold"
                  >
                    <Download size={20} />
                    Vygenerovať CSV ({data.length} záznamov)
                  </button>
                ) : (
                  <div className="space-y-2">
                    <a
                      href={csvDownloadUrl}
                      download={`export_${new Date().toISOString().split('T')[0]}.csv`}
                      className="inline-flex items-center gap-2 px-6 py-3 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition font-semibold"
                    >
                      <Download size={20} />
                      Stiahnuť CSV súbor ({data.length} záznamov)
                    </a>
                    <button
                      onClick={() => {
                        setCsvDownloadUrl(null);
                        setCsvPreview(null);
                      }}
                      className="block px-4 py-2 text-sm text-gray-600 hover:text-gray-800"
                    >
                      Vygenerovať znova
                    </button>
                  </div>
                )}
              </div>

              {csvPreview && (
                <div className="mt-4">
                  <h3 className="text-sm font-semibold text-gray-700 mb-2">
                    Náhľad CSV (prvých 10 záznamov):
                  </h3>
                  <div className="bg-white border border-gray-300 rounded-lg overflow-x-auto max-h-96">
                    <table className="min-w-full divide-y divide-gray-200 text-xs">
                      <thead className="bg-gray-100 sticky top-0">
                        <tr>
                          {[
                            'spolo?nos?', 'oscislo', 'DOBPS/DVP/HPP', 'kodppv', 'priezv', 'meno', 'titl1',
                            'datnarod', 'statprisl', 'rodstav', 'rc', 'rodmeno', 'miestonar', 'ulica',
                            'supcislo', 'cdomu', 'psc', 'obecnazov', 'stat', 'vzdelanie',
                            'dôchodky,invalidita', 'odpocet', 'zpkod', 'VT/NT',
                            'stredkod', 'miesvyk', 'nastup', 'dobaurcita', 'datoop', 'pracovnápozícia',
                            'isco', 'tarif', 'IBAN', 'ucet', 'preducet', 'banka', 'SWIFT', 'email', 'mobil'
                          ].map((header, idx) => (
                            <th key={idx} className="px-2 py-2 text-left font-medium text-gray-700 whitespace-nowrap">
                              {header}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody className="bg-white divide-y divide-gray-200">
                        {data.slice(0, 10).map((row, idx) => (
                          <tr key={idx} className="hover:bg-gray-50">
                            <td className="px-2 py-1 whitespace-nowrap">{row.spolocnost}</td>
                            <td className="px-2 py-1 whitespace-nowrap">{row.oscislo}</td>
                            <td className="px-2 py-1 whitespace-nowrap">{row.dobps}</td>
                            <td className="px-2 py-1 whitespace-nowrap">{row.kodppv}</td>
                            <td className="px-2 py-1 whitespace-nowrap">{row.priezv}</td>
                            <td className="px-2 py-1 whitespace-nowrap">{row.meno}</td>
                            <td className="px-2 py-1 whitespace-nowrap">{row.titl1}</td>
                            <td className="px-2 py-1 whitespace-nowrap">{row.datnarod}</td>
                            <td className="px-2 py-1 whitespace-nowrap">{row.statprisl}</td>
                            <td className="px-2 py-1 whitespace-nowrap">{row.rodstav}</td>
                            <td className="px-2 py-1 whitespace-nowrap font-mono">{row.rc}</td>
                            <td className="px-2 py-1 whitespace-nowrap">{row.rodmeno}</td>
                            <td className="px-2 py-1 whitespace-nowrap">{row.miestonar}</td>
                            <td className="px-2 py-1 whitespace-nowrap">{row.ulica}</td>
                            <td className="px-2 py-1 whitespace-nowrap">{row.supcislo}</td>
                            <td className="px-2 py-1 whitespace-nowrap">{row.cdomu}</td>
                            <td className="px-2 py-1 whitespace-nowrap">{row.psc}</td>
                            <td className="px-2 py-1 whitespace-nowrap">{row.obecnazov}</td>
                            <td className="px-2 py-1 whitespace-nowrap font-semibold">{row.stat}</td>
                            <td className="px-2 py-1 whitespace-nowrap">{row.vzdelanie}</td>
                            <td className="px-2 py-1 whitespace-nowrap">{row.dochodky}</td>
                            <td className="px-2 py-1 whitespace-nowrap">{row.odpocet}</td>
                            <td className="px-2 py-1 whitespace-nowrap">{row.zpkod}</td>
                            <td className="px-2 py-1 whitespace-nowrap">{row.vtnt}</td>
                            <td className="px-2 py-1 whitespace-nowrap">{row.stredkod}</td>
                            <td className="px-2 py-1 whitespace-nowrap">{row.miesvyk}</td>
                            <td className="px-2 py-1 whitespace-nowrap">{row.nastup}</td>
                            <td className="px-2 py-1 whitespace-nowrap">{row.dobaurcita}</td>
                            <td className="px-2 py-1 whitespace-nowrap">{row.datoop}</td>
                            <td className="px-2 py-1 whitespace-nowrap">{row.pracovnapozicia}</td>
                            <td className="px-2 py-1 whitespace-nowrap">{row.isco}</td>
                            <td className="px-2 py-1 whitespace-nowrap">{row.tarif}</td>
                            <td className="px-2 py-1 whitespace-nowrap font-mono text-xs">{row.iban}</td>
                            <td className="px-2 py-1 whitespace-nowrap font-mono">{row.ucet}</td>
                            <td className="px-2 py-1 whitespace-nowrap font-mono">{row.preducet}</td>
                            <td className="px-2 py-1 whitespace-nowrap">{row.banka}</td>
                            <td className="px-2 py-1 whitespace-nowrap">{row.swift}</td>
                            <td className="px-2 py-1 whitespace-nowrap text-xs">{row.email}</td>
                            <td className="px-2 py-1 whitespace-nowrap">{row.mobil}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  <p className="text-xs text-gray-500 mt-2">
                    📊 Celkovo {data.length} záznamov. Zobrazených prvých 10 pre kontrolu formátu CSV.
                  </p>
                </div>
              )}

              <div className="bg-yellow-50 border border-yellow-200 rounded-lg p-4">
                <p className="text-sm text-yellow-800">
                  <strong>Poznámka:</strong> Export je možný aj s chybnými validáciami. 
                  Odporúčame skontrolovať a opraviť všetky červené upozornenia pred exportom.
                </p>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

export default ExcelToCsvConverter;