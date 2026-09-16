// ================== DÍA 1: MEMORIA ==================
var SHEET_NAME = 'CPU';
var MEM_ROW0 = 3;
var MEM_COL0 = 2;
var COLOR_CODE_SEG = '#EEEDFE';
var COLOR_DATA_SEG = '#F1EFE8';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('CPU Simulador')
    .addItem('1. Inicializar memoria', 'initMemory')
    .addItem('Probar Read/Write (demo)', 'testReadWrite')
    .addToUi();
}

function initMemory() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  sheet.clear();
  sheet.getRange(1, 1).setValue('Simulador de CPU 8 bits — Memoria RAM (00h-FFh)').setFontWeight('bold');

  for (var c = 0; c < 16; c++) {
    sheet.getRange(MEM_ROW0 - 1, MEM_COL0 + c).setValue(c.toString(16).toUpperCase()).setFontWeight('bold').setHorizontalAlignment('center');
  }
  for (var r = 0; r < 16; r++) {
    sheet.getRange(MEM_ROW0 + r, MEM_COL0 - 1).setValue((r * 16).toString(16).toUpperCase() + 'h').setFontWeight('bold');
    for (var c = 0; c < 16; c++) {
      var addr = r * 16 + c;
      var cell = sheet.getRange(MEM_ROW0 + r, MEM_COL0 + c);
      cell.setValue(0).setHorizontalAlignment('center').setBorder(true, true, true, true, true, true);
      cell.setBackground(addr < 0x20 ? COLOR_CODE_SEG : COLOR_DATA_SEG);
    }
  }
  sheet.getRange(MEM_ROW0 - 2, MEM_COL0).setValue('Azul claro = código (00-1F)  |  Gris = datos (20-FF)').setFontStyle('italic');
  sheet.setColumnWidths(MEM_COL0, 16, 35);
}

// ---- Subrutinas primitivas requeridas por el examen ----
function Read(address) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var r = Math.floor(address / 16), c = address % 16;
  return sheet.getRange(MEM_ROW0 + r, MEM_COL0 + c).getValue();
}
function Write(address, value) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var r = Math.floor(address / 16), c = address % 16;
  sheet.getRange(MEM_ROW0 + r, MEM_COL0 + c).setValue(value);
}

// ---- Prueba manual para validar hoy mismo que Read/Write funcionan ----
function testReadWrite() {
  Write(0x10, 42);
  var valor = Read(0x10);
  SpreadsheetApp.getUi().alert('Prueba Read/Write: escribí 42 en 0x10 y leí de vuelta: ' + valor +
    (valor === 42 ? ' ✅ funciona' : ' ❌ algo falló'));
}

// ---- Utilidades que ya vamos a necesitar más adelante ----
function hex(n) { return '0x' + Number(n).toString(16).toUpperCase().padStart(2, '0'); }
function bin(n) { return Number(n).toString(2).padStart(8, '0'); }
