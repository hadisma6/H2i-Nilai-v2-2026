const fs = require('fs');
let gs = fs.readFileSync('Code.gs', 'utf-8');

gs = gs.replace(
    "case 'getJurnal':",
    "case 'getPoint': r = getPoint(p); break;\n      case 'simpanPoint': r = simpanPoint(p); break;\n      case 'hapusPoint': r = hapusPoint(p); break;\n      case 'getJurnal':"
);

const point_funcs = 
// --- POINT REWARD & PELANGGARAN ---
function getPoint() {
  try {
    var s = getSheet("DataPoint");
    var d = s.getDataRange().getValues();
    var o = [];
    for(var i=1; i<d.length; i++){
      if(d[i][0]) o.push({
        id: String(d[i][0]),
        tanggal: d[i][1],
        nama: d[i][2],
        kelas: d[i][3],
        tipe: d[i][4],
        kategori: d[i][5],
        poin: d[i][6],
        keterangan: d[i][7] // Added keterangan field
      });
    }
    return {status:'success', data:o};
  } catch(err) { return {status:'error', message:err.toString()}; }
}

function simpanPoint(p) {
  try {
    var sheet = getSheet("DataPoint");
    if(sheet.getLastRow() === 0) {
      sheet.appendRow(["ID_POINT", "TANGGAL", "NAMA_SISWA", "KELAS", "TIPE", "KATEGORI", "POIN", "KETERANGAN"]);
    }
    var isNew = !p.id;
    var pointId = isNew ? "PT-" + new Date().getTime() : p.id;
    var tgFormat = "";
    if (p.tanggal) {
       var tParts = p.tanggal.split("-");
       if (tParts.length == 3) tgFormat = tParts[2]+"/"+tParts[1]+"/"+tParts[0];
       else tgFormat = p.tanggal;
    }
    var rowData = [pointId, tgFormat, p.nama, p.kelas, p.tipe, p.kategori, p.poin, p.keterangan || "-"];
    
    if (isNew) {
      sheet.appendRow(rowData);
    } else {
      var d = sheet.getDataRange().getValues();
      var found = false;
      for (var i=1; i<d.length; i++){
        if (d[i][0] == p.id) {
          sheet.getRange(i+1, 1, 1, rowData.length).setValues([rowData]);
          found = true;
          break;
        }
      }
      if (!found) return {status:'error', message:'Data Point tidak ditemukan.'};
    }
    return {status:'success', message:'Data point berhasil disimpan.'};
  } catch(err) { return {status:'error', message:err.toString()}; }
}

function hapusPoint(p) {
  try {
    var sheet = getSheet("DataPoint");
    var d = sheet.getDataRange().getValues();
    for (var i=1; i<d.length; i++) {
        if(d[i][0] == p.id) {
            sheet.deleteRow(i+1);
            return {status:'success', message:'Data berhasil dihapus'};
        }
    }
    return {status:'error', message:'Data tidak ditemukan'};
  } catch(err) { return {status:'error', message:err.toString()}; }
}
;

const idx = gs.indexOf("function getJurnal");
if (idx !== -1) {
    gs = gs.substring(0, idx) + point_funcs + '\n\n' + gs.substring(idx);
}

fs.writeFileSync('Code.gs', gs, 'utf-8');
console.log('Code.gs patched!');
