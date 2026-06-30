// ============================================================
// LOGIKA.GS — ROKO Računi
// Vse funkcije za delovanje sistema izdajanja računov.
// ============================================================


// ============================================================
// CELIČNE REFERENCE — račun
// Če spremenite postavitev lista Račun, posodobite tukaj.
// ============================================================

var REF = {
  // Podatki o kupcu (izpolni isciKupca / dodajKupca)
  sifraKupca:  'D9',   // šifra kupca — vnosno polje
  imeKupca:    'D11',  // ime kupca (merge D-H)
  ulica:       'D12',  // ulica
  hisnaSt:     'F12',  // hišna številka
  postKoda:    'D13',  // poštna koda
  postNaziv:   'E13',  // naziv pošte

  // Glava računa (desno)
  datumRacuna: 'K9',   // datum računa — nastavi onOpen()
  krajIzdaje:  'K10',  // kraj izdaje — ročni vpis
  valuta:      'K11',  // valuta (rok plačila) — valuta30() / valuta90()
  storitev:    'K12',  // datum opravljene storitve — ročni vpis
  idDdv:       'K13',  // ID DDV kupca — izpolni isciKupca()

  // Številka računa
  stevilka:    'J15',  // format LLXX (npr. 2601) — plain text celica!

  // Postavke (vrstice 18-36)
  vrstica1:    18,
  vrsticaN:    36,
  colKolicina: 3,   // C — količina
  colEM:       4,   // D — enota mere
  colOpis:     5,   // E — opis (merge E-H)
  colCena:     9,   // I — cena na enoto
  colVrednost: 11,  // K — vrednost (formula: =IF(C18="","",ROUND(C18*I18,2)))

  // Seštevki (formule)
  skupaj:      'K37',  // =SUM(K18:K36)
  zaPlacilo:   'K38',  // =K37

  // Sklic za plačilo (formula)
  // ="Plačilo na TRR : SI56 0000 0000 0000 000, sklic : 00 - "&J15&" - "&D9
  sklic:       'E39'
};


// ============================================================
// KONSTANTE — šifrant
// ============================================================

var SIFRANT_VNOSNA_VRSTICA = 3;    // vrstica z vnosnimi polji za novega kupca
var SIFRANT_PODATKI_ZACETEK = 8;   // prva vrstica s podatki kupcev

// Stolpci šifranta (1 = A, 2 = B, ...)
var SIFRANT_COLS = {
  sifra:   1,  // A — šifra kupca
  idDdv:   2,  // B — ID za DDV
  kupec:   3,  // C — ime kupca (velike tiskane črke)
  ulica:   4,  // D — ulica
  hisnaSt: 5,  // E — hišna številka
  postKod: 6,  // F — poštna koda
  posta:   7,  // G — naziv pošte (velike tiskane črke)
  banka:   8   // H — banka (opcijsko)
};


// ============================================================
// ON EDIT — avtomatske akcije ob urejanju
// ============================================================

function onEdit(e) {
  if (!e) return;
  var sheet  = e.range.getSheet();
  var celica = e.range.getA1Notation();

  // Na listu Šifrant: ime kupca (C3) in pošta (G3) se
  // avtomatsko pretvorita v velike tiskane črke
  if (sheet.getName() === 'Šifrant') {
    if (celica === 'C3' || celica === 'G3') {
      var val = e.range.getValue();
      if (val) e.range.setValue(val.toString().toUpperCase());
    }
  }
}


// ============================================================
// ON OPEN — ob odprtju dokumenta
// ============================================================

function onOpen() {
  // Nastavi današnji datum na računu (kot fiksna vrednost, ne formula)
  // Datum se posodobi samo ob vsakem odprtju — med delom ostane fiksen
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var racun = ss.getSheetByName('Račun');
  if (racun) {
    racun.getRange('K9').setValue(
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'd.M.yyyy')
    );
  }

  // Zgradi meni "ROKO Računi" v menijski vrstici
  SpreadsheetApp.getUi()
    .createMenu('🧾 ROKO Računi')
    .addItem('Dodaj kupca', 'dodajKupca')           // šifrant → račun
    .addSeparator()
    .addItem('Išči kupca', 'isciKupca')             // D9 → izpolni kupca
    .addItem('Valuta +30', 'valuta30')              // K9 + 30 dni → K11
    .addItem('Valuta +90', 'valuta90')              // K9 + 90 dni → K11
    .addItem('Shrani v knjigo + izvozi PDF', 'shraniInIzvozi')
    .addSeparator()
    .addItem('Išči v knjigi', 'isciVKnjigi')        // filtrira knjigo računov
    .addItem('Cel dnevnik', 'celDnevnik')           // pokaže vse vrstice
    .addToUi();
}


// ============================================================
// DODAJ KUPCA
// Prebere vnosna polja (vrstica 3) na listu Šifrant,
// doda kupca na konec šifranta, prenese podatke na račun
// in predlaga naslednjo razpoložljivo šifro.
// Gumb na listu Šifrant → dodajKupca
// ============================================================

function dodajKupca() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var sifrant = ss.getSheetByName('Šifrant');
  var racun   = ss.getSheetByName('Račun');

  // Preberi vnosna polja A3:H3
  var vnos = sifrant.getRange(SIFRANT_VNOSNA_VRSTICA, 1, 1, 8).getValues()[0];

  // Preveri da sta vsaj šifra in ime izpolnjena
  if (!vnos[SIFRANT_COLS.sifra - 1] && !vnos[SIFRANT_COLS.kupec - 1]) {
    SpreadsheetApp.getUi().alert('Vpiši vsaj šifro in ime kupca!');
    return;
  }

  // Dodaj na konec šifranta (minimalno vrstica 8)
  var zadnja = Math.max(sifrant.getLastRow() + 1, SIFRANT_PODATKI_ZACETEK);
  sifrant.getRange(zadnja, 1, 1, 8).setValues([vnos])
    .setFontSize(10)
    .setVerticalAlignment('middle');

  // Prenesi podatke novega kupca direktno na račun
  racun.getRange(REF.sifraKupca).setValue(vnos[SIFRANT_COLS.sifra - 1]);
  racun.getRange(REF.imeKupca).setValue(vnos[SIFRANT_COLS.kupec - 1]);
  racun.getRange(REF.ulica).setValue(vnos[SIFRANT_COLS.ulica - 1]);
  racun.getRange(REF.hisnaSt).setValue(vnos[SIFRANT_COLS.hisnaSt - 1]);
  racun.getRange(REF.postKoda).setValue(vnos[SIFRANT_COLS.postKod - 1]);
  racun.getRange(REF.postNaziv).setValue(vnos[SIFRANT_COLS.posta - 1]);
  racun.getRange(REF.idDdv).setValue(vnos[SIFRANT_COLS.idDdv - 1] || 0);

  // Počisti vnosna polja in predlagaj naslednjo šifro
  sifrant.getRange(SIFRANT_VNOSNA_VRSTICA, 1, 1, 8).clearContent();
  predlagajNaslednjoSifro(sifrant);

  // Preklopi na list Račun
  ss.setActiveSheet(racun);

  SpreadsheetApp.getUi().alert('✓ Kupec dodan in prenesen na račun!');
}

// Poišče največjo obstoječo šifro v šifrantu in vpiše +1 v vnosno polje
function predlagajNaslednjoSifro(sifrant) {
  var zadnja = sifrant.getLastRow();
  if (zadnja < SIFRANT_PODATKI_ZACETEK) {
    sifrant.getRange(SIFRANT_VNOSNA_VRSTICA, 1).setValue(1);
    return;
  }

  var sifre = sifrant.getRange(SIFRANT_PODATKI_ZACETEK, 1, zadnja - SIFRANT_PODATKI_ZACETEK + 1, 1).getValues();
  var max = 0;
  for (var i = 0; i < sifre.length; i++) {
    var s = parseInt(sifre[i][0]);
    if (!isNaN(s) && s > max) max = s;
  }
  sifrant.getRange(SIFRANT_VNOSNA_VRSTICA, 1).setValue(max + 1);
}


// ============================================================
// IŠČI KUPCA
// Prebere šifro iz D9, poišče kupca v šifrantu in
// izpolni polja kupca na računu.
// Gumb na listu Račun → isciKupca
// ============================================================

function isciKupca() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var racun   = ss.getSheetByName('Račun');
  var sifrant = ss.getSheetByName('Šifrant');

  // Preberi šifro iz vnosnega polja
  var sifra = String(racun.getRange(REF.sifraKupca).getValue()).trim();
  if (!sifra || sifra === '') {
    SpreadsheetApp.getUi().alert('Vpiši šifro kupca v celico D9!');
    return;
  }

  var zadnja = sifrant.getLastRow();
  if (zadnja < SIFRANT_PODATKI_ZACETEK) {
    SpreadsheetApp.getUi().alert('Šifrant je prazen!');
    return;
  }

  // Preišči šifrant po stolpcu A
  var podatki = sifrant.getRange(SIFRANT_PODATKI_ZACETEK, 1, zadnja - SIFRANT_PODATKI_ZACETEK + 1, 8).getValues();
  var najden  = null;

  for (var i = 0; i < podatki.length; i++) {
    if (String(podatki[i][SIFRANT_COLS.sifra - 1]).trim() === sifra) {
      najden = podatki[i];
      break;
    }
  }

  if (!najden) {
    SpreadsheetApp.getUi().alert('Kupec s šifro ' + sifra + ' ni najden!');
    return;
  }

  // Izpolni polja kupca na računu
  racun.getRange(REF.imeKupca).setValue(najden[SIFRANT_COLS.kupec - 1]);
  racun.getRange(REF.ulica).setValue(najden[SIFRANT_COLS.ulica - 1]);
  racun.getRange(REF.hisnaSt).setValue(najden[SIFRANT_COLS.hisnaSt - 1]);
  racun.getRange(REF.postKoda).setValue(najden[SIFRANT_COLS.postKod - 1]);
  racun.getRange(REF.postNaziv).setValue(najden[SIFRANT_COLS.posta - 1]);
  racun.getRange(REF.idDdv).setValue(najden[SIFRANT_COLS.idDdv - 1] || 0);
}


// ============================================================
// VALUTA
// Izračuna rok plačila kot datum računa + 30 ali 90 dni.
// Gumb na listu Račun → valuta30 / valuta90
// ============================================================

function valuta30() {
  var sh    = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Račun');
  var datum = new Date(sh.getRange('K9').getValue());
  datum.setDate(datum.getDate() + 30);
  sh.getRange('K11').setValue(Utilities.formatDate(datum, Session.getScriptTimeZone(), 'd.M.yyyy'));
}

function valuta90() {
  var sh    = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Račun');
  var datum = new Date(sh.getRange('K9').getValue());
  datum.setDate(datum.getDate() + 90);
  sh.getRange('K11').setValue(Utilities.formatDate(datum, Session.getScriptTimeZone(), 'd.M.yyyy'));
}


// ============================================================
// SHRANI V KNJIGO + IZVOZI PDF
// Prebere podatke z računa, zapiše vrstico v knjigo računov,
// izvozi PDF v Drive mapo "Računi" in poveča številko računa.
// Gumb na listu Račun → shraniInIzvozi
// ============================================================

function shraniInIzvozi() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var racun  = ss.getSheetByName('Račun');
  var knjiga = ss.getSheetByName('Knjiga računov');
  var ui     = SpreadsheetApp.getUi();

  // Preberi podatke z računa
  var stevilka = racun.getRange(REF.stevilka).getValue();
  var datum    = racun.getRange(REF.datumRacuna).getValue();
  var sifra    = racun.getRange(REF.sifraKupca).getValue();
  var kupec    = racun.getRange(REF.imeKupca).getValue();
  var idDdv    = racun.getRange(REF.idDdv).getValue();
  var znesek   = racun.getRange(REF.zaPlacilo).getValue();

  // Osnovna validacija
  if (!stevilka || !kupec || !znesek) {
    ui.alert('Manjkajo podatki! Preveri številko računa, kupca in znesek.');
    return;
  }

  // Preveri podvojen račun
  var zadnjaVrstica = knjiga.getLastRow();
  if (zadnjaVrstica >= 5) {
    var obstojeceStevike = knjiga.getRange(5, 2, zadnjaVrstica - 4, 1).getValues();
    for (var i = 0; i < obstojeceStevike.length; i++) {
      if (String(obstojeceStevike[i][0]).trim() === String(stevilka).trim()) {
        var odgovor = ui.alert(
          'Opozorilo',
          'Račun ' + stevilka + ' je že vpisan v knjigo računov!\n\nVseeno nadaljujem?',
          ui.ButtonSet.YES_NO
        );
        if (odgovor !== ui.Button.YES) return;
        break;
      }
    }
  }

  // Zapiši vrstico v knjigo računov (stolpci B-G = 2-7, podatki od vrstice 5)
  var novaVrstica = Math.max(knjiga.getLastRow() + 1, 5);
  knjiga.getRange(novaVrstica, 2).setValue(stevilka);
  knjiga.getRange(novaVrstica, 3).setValue(datum);
  knjiga.getRange(novaVrstica, 4).setValue(sifra);
  knjiga.getRange(novaVrstica, 5).setValue(kupec);
  knjiga.getRange(novaVrstica, 6).setValue(idDdv || 0);
  knjiga.getRange(novaVrstica, 7).setValue(znesek);
  knjiga.getRange(novaVrstica, 7).setNumberFormat('#,##0.00');

  // Izvozi PDF (počisti zelena polja, izvozi, vrni barve)
  var rezultat = izvozPdf(ss, racun, stevilka, kupec);

  // Poveča številko računa za 1 (format LLXX → LL(XX+1))
  var stevilkaStr = stevilka.toString();
  var leto        = stevilkaStr.slice(0, 2);
  var zap         = parseInt(stevilkaStr.slice(2)) || 0;
  var nova        = leto + String(zap + 1).padStart(2, '0');
  racun.getRange(REF.stevilka).setValue(nova);

  // Obvesti uporabnika in odpri Drive mapo v novem zavihku
  if (rezultat.uspeh) {
    ui.alert('✓ Račun ' + stevilka + ' shranjen!');
    var html = HtmlService.createHtmlOutput(
      '<script>window.open("' + rezultat.mapaUrl + '");google.script.host.close();</script>'
    );
    SpreadsheetApp.getUi().showModalDialog(html, 'Odpiranje mape...');
  } else {
    ui.alert('✓ Račun ' + stevilka + ' shranjen v knjigo.\n⚠ PDF izvoz ni uspel.');
  }
}


// ============================================================
// IZVOZ PDF
// Začasno odstrani zelena ozadja vnosnih polj (da ne gredo
// v PDF), izvozi list Račun kot PDF v Drive mapo "Računi",
// nato vrne barve nazaj.
// Print area: stolpci C-K (c1=2, c2=11), vrstice 1-47 (r1=0, r2=46)
// ============================================================

function izvozPdf(ss, racun, stevilka, kupec) {
  // Zelena vnosna polja ki jih začasno počistimo pred izvozom
  var zeleni = ['D9', 'K9', 'K11', 'K12', 'J15', 'C18:I36'];

  try {
    var mapa = najdiAliUstvari('Računi', DriveApp.getRootFolder());

    // Začasno odstrani zelena ozadja
    zeleni.forEach(function(r) { racun.getRange(r).setBackground(null); });
    SpreadsheetApp.flush();

    // Sestavi URL za PDF izvoz
    var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() +
      '/export?format=pdf&gid=' + racun.getSheetId() +
      '&size=A4&portrait=true&fitw=true' +
      '&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false' +
      '&top_margin=0.5&bottom_margin=0.5&left_margin=0.5&right_margin=0.5' +
      '&r1=0&c1=2&r2=52&c2=11';  // print area: vrstice 1-47, stolpci C-K

    // Prenesi in poimenuj PDF
    var pdf = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() }
    }).getBlob();

    pdf.setName(stevilka + ' - ' + kupec.trim() + '.pdf');
    mapa.createFile(pdf);

    // Vrni zelena ozadja
    zeleni.forEach(function(r) { racun.getRange(r).setBackground('#ccffcc'); });

    return { uspeh: true, mapaUrl: 'https://drive.google.com/drive/folders/' + mapa.getId() };

  } catch(e) {
    // Vrni barve tudi ob napaki
    try {
      zeleni.forEach(function(r) { racun.getRange(r).setBackground('#ccffcc'); });
    } catch(e2) {}
    Logger.log('PDF napaka: ' + e.toString());
    return { uspeh: false, mapaUrl: null };
  }
}

// Pomožna funkcija: poišče mapo po imenu ali jo ustvari
function najdiAliUstvari(ime, stars) {
  var it = stars.getFoldersByName(ime);
  return it.hasNext() ? it.next() : stars.createFolder(ime);
}


// ============================================================
// ISKANJE V KNJIGI RAČUNOV
// Filtrira vrstice glede na iskalni niz v C2.
// Išče po vseh stolpcih hkrati (številka, datum, kupec...).
// Gumb na listu Knjiga računov → isciVKnjigi
// ============================================================

function isciVKnjigi() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var sh      = ss.getSheetByName('Knjiga računov');
  var iskanje = sh.getRange('C2').getValue().toString().trim().toLowerCase();

  var zadnja = sh.getLastRow();
  if (zadnja < 5) return;

  // Najprej pokaži vse vrstice
  sh.showRows(5, zadnja - 4);
  if (iskanje === '') return;

  // Skrij vrstice ki ne vsebujejo iskalnega niza
  var podatki = sh.getRange(5, 2, zadnja - 4, 6).getValues();
  for (var i = 0; i < podatki.length; i++) {
    if (podatki[i].join(' ').toLowerCase().indexOf(iskanje) === -1) {
      sh.hideRows(5 + i);
    }
  }
}

// Pokaže vse vrstice in počisti iskalno polje
// Gumb na listu Knjiga računov → celDnevnik
function celDnevnik() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var sh     = ss.getSheetByName('Knjiga računov');
  var zadnja = sh.getLastRow();
  if (zadnja >= 5) sh.showRows(5, zadnja - 4);
  sh.getRange('C2').clearContent();
  sh.getRange('B5').activate();  // skoči na prvo vrstico podatkov
}
