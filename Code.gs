// ============================================================
// Dashboard Mommy 'n Me 2026 — Code.gs (Backend)
// PT Bank Negara Indonesia (Persero) Tbk | Dept. RTR Divisi RPP
// Created by HQ/63927
// ============================================================

var SPREADSHEET_ID = "1QMYyjWEh9_u_JhNXXYS2xxBdQXSrQPt3wjPyW2PgviY"; // <-- GANTI DENGAN ID SPREADSHEET ANDA

var SHEET_NEW    = "Data_Utama_Baru";
var SHEET_OLD    = "Data_Utama_Lama";
var SHEET_ACQ    = "Data_Akuisisi";
var SHEET_TICKET = "Data_Ticket";
var SHEET_CONFIG = "Config";

// ----------------------------------------------------------
// doGet: Entry point Web App
// ----------------------------------------------------------
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Dashboard Mommy 'n Me 2026")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

// ----------------------------------------------------------
// getDataDashboard: Fungsi utama pengambilan & agregasi data
// ----------------------------------------------------------
function getDataDashboard() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    var newData    = _getSheetData(ss, SHEET_NEW);
    var oldData    = _getSheetData(ss, SHEET_OLD);
    var acqData    = _getSheetData(ss, SHEET_ACQ);
    var ticketData = _getSheetData(ss, SHEET_TICKET);
    var configData = _getSheetData(ss, SHEET_CONFIG);

    var target = _getTarget(configData);

    // Normalisasi kolom nama agar seragam
    var newRows = _normalizeTransactions(newData);
    var oldRows = _normalizeTransactions(oldData);

    var halls = ["ALL", "MAIN LOBBY", "ASSEMBLY", "PLENARY", "CENDRAWASIH", "HALL A", "HALL B"];

    var hallResults = {};
    halls.forEach(function(hall) {
      var nRows = hall === "ALL" ? newRows : newRows.filter(function(r){ return r.hall === hall; });
      var oRows = hall === "ALL" ? oldRows : oldRows.filter(function(r){ return r.hall === hall; });
      hallResults[hall] = _buildHallMetrics(nRows, oRows, target);
    });

    var acqRows   = _buildAcquisition(acqData);
    var ticketSummary = _buildTicketSummary(ticketData);

    return {
      success: true,
      target: target,
      halls: hallResults,
      acquisition: acqRows,
      ticket: ticketSummary,
      lastUpdated: new Date().toLocaleString("id-ID", { timeZone: "Asia/Jakarta" })
    };

  } catch (err) {
    return { success: false, error: err.message, stack: err.stack };
  }
}

// ----------------------------------------------------------
// _getSheetData: Baca sheet dan kembalikan array of objects
// ----------------------------------------------------------
function _getSheetData(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error("Sheet tidak ditemukan: " + sheetName);
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0].map(function(h){ return String(h).trim().toUpperCase(); });
  return data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i){ obj[h] = row[i]; });
    return obj;
  });
}

// ----------------------------------------------------------
// _getTarget: Ambil nilai Target Sales Volume dari Config
// ----------------------------------------------------------
function _getTarget(configRows) {
  for (var i = 0; i < configRows.length; i++) {
    var param = String(configRows[i]["PARAMETER"] || configRows[i]["Parameter"] || "").trim().toUpperCase();
    if (param === "TARGET SALES VOLUME") {
      return Number(configRows[i]["VALUE"] || configRows[i]["Value"] || 0);
    }
  }
  return 0;
}

// ----------------------------------------------------------
// _normalizeTransactions: Standarisasi kolom transaksi
// ----------------------------------------------------------
function _normalizeTransactions(rows) {
  return rows.map(function(r) {
    var dateRaw = r["TRANSACTION DATE TIME"] || r["TRX_DATE"] || r["TRANSACTION DATE"] || "";
    var dateObj  = _excelDateToJS(dateRaw);
    return {
      merchantId:   String(r["MERCHANT ID"]   || r["MERCHANT_ID"]   || ""),
      merchantName: String(r["MERCHANT NAME"] || r["MER_NM"]        || ""),
      date:         dateObj,
      dateStr:      _formatDate(dateObj),
      cardType:     String(r["CARD TYPE"]  || r["CARD_TYP"]  || "").trim(),
      trxType:      String(r["TANSACTION TYPE"] || r["TRX TYPE"] || r["TRX_TYPE"] || "").trim().toUpperCase(),
      amount:       Number(r["SALES VOLUME"] || r["AUTH_TRNS_AMT"] || 0),
      hall:         String(r["HALL"] || r["Hall"] || "").trim().toUpperCase()
    };
  }).filter(function(r){ return r.amount > 0; });
}

// ----------------------------------------------------------
// _excelDateToJS: Konversi serial date Excel ke Date JS
// ----------------------------------------------------------
function _excelDateToJS(val) {
  if (val instanceof Date) return val;
  if (typeof val === "number" && val > 1000) {
    return new Date(Math.round((val - 25569) * 86400 * 1000));
  }
  var d = new Date(val);
  return isNaN(d) ? new Date(0) : d;
}

function _formatDate(d) {
  if (!d || isNaN(d.getTime())) return "";
  return d.getFullYear() + "-" + _pad(d.getMonth()+1) + "-" + _pad(d.getDate());
}

function _pad(n) { return n < 10 ? "0"+n : ""+n; }

// ----------------------------------------------------------
// _buildHallMetrics: Bangun semua metrik untuk satu hall
// ----------------------------------------------------------
function _buildHallMetrics(nRows, oRows, target) {

  var totalVolNew  = nRows.reduce(function(s,r){ return s + r.amount; }, 0);
  var totalVolOld  = oRows.reduce(function(s,r){ return s + r.amount; }, 0);
  var totalFreqNew = nRows.length;
  var totalFreqOld = oRows.length;
  var achievement  = target > 0 ? (totalVolNew / target) * 100 : 0;

  // Time-series harian
  var tsNew = _timeSeries(nRows);
  var tsOld = _timeSeries(oRows);

  // Bar Chart: Volume per CardType x TrxType
  var cardBar = _cardTypeBreakdown(nRows, oRows);

  // Doughnut: Frekuensi per CardType
  var donutNew = _cardTypeFreq(nRows);
  var donutOld = _cardTypeFreq(oRows);

  // Growth YoY
  var growthFD   = _growthPercent(totalVolNew, totalVolOld);
  var growthDays = _growthByDay(nRows, oRows);

  // Top 10 / Bottom 10 Merchant
  var merchantRank = _merchantRank(nRows);

  // Top 10 ON US / OFF US
  var onUsTop  = _trxTypeTop(nRows, "ON US");
  var offUsTop = _trxTypeTop(nRows, "OFF US");

  // Top 10 Credit / Debit
  var creditTop = _cardTop(nRows, "Credit");
  var debitTop  = _cardTop(nRows, "Debit");

  // Sparkline data (7 titik terakhir)
  var sparkVol  = _sparkline(tsNew, "volume");
  var sparkFreq = _sparkline(tsNew, "freq");

  return {
    totalVolNew:  totalVolNew,
    totalVolOld:  totalVolOld,
    totalFreqNew: totalFreqNew,
    totalFreqOld: totalFreqOld,
    achievement:  achievement,
    timeSeries:   { new: tsNew, old: tsOld },
    cardBar:      cardBar,
    donutNew:     donutNew,
    donutOld:     donutOld,
    growthFD:     growthFD,
    growthDays:   growthDays,
    merchantTop:  merchantRank.top,
    merchantBot:  merchantRank.bottom,
    onUsTop:      onUsTop,
    offUsTop:     offUsTop,
    creditTop:    creditTop,
    debitTop:     debitTop,
    sparkVol:     sparkVol,
    sparkFreq:    sparkFreq
  };
}

// ----------------------------------------------------------
// _timeSeries: Agregasi volume & frekuensi per hari
// ----------------------------------------------------------
function _timeSeries(rows) {
  var map = {};
  rows.forEach(function(r) {
    var d = r.dateStr;
    if (!d) return;
    if (!map[d]) map[d] = { date: d, volume: 0, freq: 0 };
    map[d].volume += r.amount;
    map[d].freq   += 1;
  });
  return Object.values(map).sort(function(a,b){ return a.date < b.date ? -1 : 1; });
}

// ----------------------------------------------------------
// _cardTypeBreakdown: Volume per CardType x TrxType (stacked)
// ----------------------------------------------------------
function _cardTypeBreakdown(nRows, oRows) {
  var cardTypes = ["Credit","Debit","QR Credit","QR Non Credit","TapCash"];
  var trxTypes  = ["ON US","OFF US"];
  var result = { cardTypes: cardTypes, trxTypes: trxTypes, new: {}, old: {} };
  cardTypes.forEach(function(ct) {
    result.new[ct] = {}; result.old[ct] = {};
    trxTypes.forEach(function(tt) {
      result.new[ct][tt] = 0;
      result.old[ct][tt] = 0;
    });
  });
  nRows.forEach(function(r) {
    if (result.new[r.cardType] && r.trxType) {
      var tt = trxTypes.includes(r.trxType) ? r.trxType : "ON US";
      result.new[r.cardType][tt] = (result.new[r.cardType][tt] || 0) + r.amount;
    }
  });
  oRows.forEach(function(r) {
    if (result.old[r.cardType] && r.trxType) {
      var tt = trxTypes.includes(r.trxType) ? r.trxType : "ON US";
      result.old[r.cardType][tt] = (result.old[r.cardType][tt] || 0) + r.amount;
    }
  });
  return result;
}

// ----------------------------------------------------------
// _cardTypeFreq: Frekuensi per CardType
// ----------------------------------------------------------
function _cardTypeFreq(rows) {
  var cardTypes = ["Credit","Debit","QR Credit","QR Non Credit","TapCash"];
  var map = {};
  cardTypes.forEach(function(ct){ map[ct] = 0; });
  rows.forEach(function(r){
    if (map[r.cardType] !== undefined) map[r.cardType]++;
  });
  return map;
}

// ----------------------------------------------------------
// _growthPercent: Hitung pertumbuhan YoY
// ----------------------------------------------------------
function _growthPercent(newVal, oldVal) {
  if (oldVal === 0) return newVal > 0 ? 100 : 0;
  return ((newVal - oldVal) / oldVal) * 100;
}

// ----------------------------------------------------------
// _growthByDay: Pertumbuhan per hari (Day 1, 2, 3)
// ----------------------------------------------------------
function _growthByDay(nRows, oRows) {
  var nDates = _uniqueSortedDates(nRows);
  var oDates = _uniqueSortedDates(oRows);
  var result = { day1: null, day2: null, day3: null };

  ["day1","day2","day3"].forEach(function(key, idx) {
    var nDate = nDates[idx] || null;
    var oDate = oDates[idx] || null;
    var nVol  = nDate ? nRows.filter(function(r){ return r.dateStr === nDate; }).reduce(function(s,r){ return s+r.amount; },0) : 0;
    var oVol  = oDate ? oRows.filter(function(r){ return r.dateStr === oDate; }).reduce(function(s,r){ return s+r.amount; },0) : 0;
    result[key] = { growth: _growthPercent(nVol, oVol), nVol: nVol, oVol: oVol, nDate: nDate, oDate: oDate };
  });
  return result;
}

function _uniqueSortedDates(rows) {
  var set = {};
  rows.forEach(function(r){ if(r.dateStr) set[r.dateStr] = true; });
  return Object.keys(set).sort();
}

// ----------------------------------------------------------
// _merchantRank: Top 10 & Bottom 10 Merchant by volume
// ----------------------------------------------------------
function _merchantRank(rows) {
  var map = {};
  rows.forEach(function(r) {
    var key = r.merchantId + "||" + r.merchantName;
    if (!map[key]) map[key] = { id: r.merchantId, name: r.merchantName, volume: 0, freq: 0 };
    map[key].volume += r.amount;
    map[key].freq   += 1;
  });
  var arr = Object.values(map).sort(function(a,b){ return b.volume - a.volume; });
  return {
    top:    arr.slice(0, 10),
    bottom: arr.slice(-10).reverse()
  };
}

// ----------------------------------------------------------
// _trxTypeTop: Top 10 per TrxType
// ----------------------------------------------------------
function _trxTypeTop(rows, trxType) {
  var filtered = rows.filter(function(r){ return r.trxType === trxType; });
  return _merchantRank(filtered).top;
}

// ----------------------------------------------------------
// _cardTop: Top 10 per CardType
// ----------------------------------------------------------
function _cardTop(rows, cardType) {
  var filtered = rows.filter(function(r){ return r.cardType === cardType; });
  return _merchantRank(filtered).top;
}

// ----------------------------------------------------------
// _sparkline: Ambil 7 titik terakhir untuk sparkline chart
// ----------------------------------------------------------
function _sparkline(timeSeries, field) {
  var last7 = timeSeries.slice(-7);
  return last7.map(function(d){ return d[field] || 0; });
}

// ----------------------------------------------------------
// _buildAcquisition: Bangun data akuisisi merchant
// ----------------------------------------------------------
function _buildAcquisition(rows) {
  return rows.map(function(r) {
    return {
      hall:          String(r["HALL"]                  || r["Hall"]                  || ""),
      jumlahTenant:  Number(r["JUMLAH TENANT"]         || r["Jumlah Tenant"]         || 0),
      totalDoneFU:   Number(r["TOTAL DONE FU"]         || r["Total Done FU"]         || 0),
      distributor:   Number(r["DISTRIBUTOR"]           || r["Distributor"]           || 0),
      tokoOnline:    Number(r["TOKO ONLINE"]           || r["Toko Online"]           || 0),
      offlinePot:    Number(r["OFFLINE/BERPOTENSI"]    || r["Offline/Berpotensi"]    || 0),
      existingMer:   Number(r["EXISTING MERCHANT"]     || r["Existing Merchant"]     || 0),
      brandPartner:  Number(r["BRAND PARTNER"]         || r["Brand Partner"]         || 0),
      tidakBerminat: Number(r["TIDAK BERMINAT"]        || r["Tidak Berminat"]        || 0),
      gap:           Number(r["GAP (BELUM FU)"]        || r["GAP"]                   || 0),
      totalAkuisisi: Number(r["TOTAL AKUISISI"]        || r["Total Akuisisi"]        || 0)
    };
  });
}

// ----------------------------------------------------------
// _buildTicketSummary: Bangun ringkasan data tiket
// ----------------------------------------------------------
function _buildTicketSummary(rows) {
  var totalVol = 0, totalQty = 0;
  var detail = [];
  rows.forEach(function(r) {
    var vol = Number(r["SALES VOLUME"] || r["TICKET_VOL"] || 0);
    var qty = Number(r["TICKET_QTY"]  || 1);
    totalVol += vol;
    totalQty += qty;
    detail.push({
      id:   String(r["MERCHANT ID"] || r["TICKET_ID"] || ""),
      name: String(r["MERCHANT NAME"] || ""),
      date: _formatDate(_excelDateToJS(r["TRANSACTION DATE TIME"] || r["SALES_DATE"] || "")),
      vol:  vol,
      qty:  qty
    });
  });
  return { totalVol: totalVol, totalQty: detail.length, detail: detail };
}