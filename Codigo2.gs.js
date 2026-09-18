// ================== DÍA 1: MEMORIA ==================
var SHEET_NAME = 'CPU';
var MEM_ROW0 = 3;
var MEM_COL0 = 2;
var COLOR_CODE_SEG = '#EEEDFE';
var COLOR_DATA_SEG = '#F1EFE8';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('CPU Simulador')
    .addItem('1. Inicializar todo', 'initTodo')
    .addSeparator()
    .addItem('Step', 'stepCPU')
    .addItem('Run', 'runCPU')
    .addItem('Reset', 'resetCPU')
    .addItem('Recargar programa', 'loadProgram')
    .addToUi();
}

function initTodo() {
  initMemory();
  initRegistros();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var log = ss.getSheetByName(LOG_SHEET_NAME) || ss.insertSheet(LOG_SHEET_NAME);
  log.clear();
  log.getRange(1, 1, 1, 3).setValues([['Paso', 'Fase', 'Detalle']]).setFontWeight('bold');
  log.setColumnWidth(3, 500);
  resetCPU();
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








// ================== DÍA 2: REGISTROS Y FLAGS ==================
var REG_COL_LABEL = 19, REG_COL_VALUE = 20;
var ROW_PC = 3, ROW_IR_OP = 4, ROW_IR1 = 5, ROW_IR2 = 6,
    ROW_MAR = 7, ROW_MDR = 8, ROW_AX = 9, ROW_BX = 10,
    ROW_ZF = 11, ROW_CF = 12, ROW_SF = 13, ROW_ESTADO = 14;

function initRegistros() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var labels = [
    [ROW_PC, 'PC'], [ROW_IR_OP, 'IR (opcode)'], [ROW_IR1, 'IR (op1)'], [ROW_IR2, 'IR (op2)'],
    [ROW_MAR, 'MAR'], [ROW_MDR, 'MDR'], [ROW_AX, 'AX'], [ROW_BX, 'BX'],
    [ROW_ZF, 'ZF'], [ROW_CF, 'CF'], [ROW_SF, 'SF'], [ROW_ESTADO, 'Estado']
  ];
  sheet.getRange(2, REG_COL_LABEL).setValue('REGISTROS').setFontWeight('bold');
  labels.forEach(function (item) {
    sheet.getRange(item[0], REG_COL_LABEL).setValue(item[1]).setFontWeight('bold');
    sheet.getRange(item[0], REG_COL_VALUE).setBorder(true, true, true, true, true, true);
  });
  sheet.setColumnWidth(REG_COL_LABEL, 170);
  sheet.setColumnWidth(REG_COL_VALUE, 90);
  resetRegistros();
}

function resetRegistros() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  [ROW_PC, ROW_IR_OP, ROW_IR1, ROW_IR2, ROW_MAR, ROW_MDR, ROW_AX, ROW_BX, ROW_ZF, ROW_CF, ROW_SF]
    .forEach(function (row) { setCell(sheet, row, 0); });
  setCell(sheet, ROW_ESTADO, 'LISTO');
}

function getCell(sheet, row) { return sheet.getRange(row, REG_COL_VALUE).getValue(); }
function setCell(sheet, row, val) { sheet.getRange(row, REG_COL_VALUE).setValue(val); }
function getReg(sheet, code) { return getCell(sheet, code === 0 ? ROW_AX : ROW_BX); }
function setReg(sheet, code, val) { setCell(sheet, code === 0 ? ROW_AX : ROW_BX, val); }
function regRow(code) { return code === 0 ? ROW_AX : ROW_BX; }
function regName(code) { return code === 0 ? 'AX' : 'BX'; }






// ================== DÍA 3: CICLO DE INSTRUCCIÓN ==================
var LOG_SHEET_NAME = 'Log';
var ROW_FASE = 15, ROW_MNEMO = 16;

var PROGRAMA = [
  [0x00, 0x01, 0x00, 0x05], // MOV AX,5
  [0x03, 0x01, 0x01, 0x03], // MOV BX,3
  [0x06, 0x05, 0x00, 0x01], // ADD AX,BX
  [0x09, 0xFF, 0x00, 0x00]  // HLT
];

function loadProgram() {
  for (var a = 0; a < 0x20; a++) Write(a, 0);
  PROGRAMA.forEach(function (instr) {
    Write(instr[0], instr[1]); Write(instr[0] + 1, instr[2]); Write(instr[0] + 2, instr[3]);
  });
}

function resetCPU() {
  resetRegistros();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  setCell(sheet, ROW_FASE, 0);
  setCell(sheet, ROW_MNEMO, '-');
  loadProgram();
  var log = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
  log.getRange(2, 1, Math.max(log.getLastRow() - 1, 1), 3).clearContent();
}

function stepCPU() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (getCell(sheet, ROW_ESTADO) === 'DETENIDO') { appendLog('-', 'CPU detenida. Usa Reset.'); return; }
  var fase = getCell(sheet, ROW_FASE);
  if (fase === 0) { doFetch(sheet); setCell(sheet, ROW_FASE, 1); }
  else if (fase === 1) { doDecode(sheet); setCell(sheet, ROW_FASE, 2); }
  else if (fase === 2) { doExecute(sheet); setCell(sheet, ROW_FASE, 3); }
  else { doStore(sheet); setCell(sheet, ROW_FASE, 0); }
}

function runCPU() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var i = 0;
  while (getCell(sheet, ROW_ESTADO) !== 'DETENIDO' && i < 200) {
    stepCPU(); SpreadsheetApp.flush(); Utilities.sleep(250); i++;
  }
}

function doFetch(sheet) {
  var pc = getCell(sheet, ROW_PC);
  setCell(sheet, ROW_MAR, pc);
  var opcode = Read(pc), op1 = Read(pc + 1), op2 = Read(pc + 2);
  setCell(sheet, ROW_MDR, opcode);
  setCell(sheet, ROW_IR_OP, opcode); setCell(sheet, ROW_IR1, op1); setCell(sheet, ROW_IR2, op2);
  setCell(sheet, ROW_PC, pc + 3);
  appendLog('FETCH', 'PC=' + hex(pc) + ' -> IR=[' + hex(opcode) + ',' + hex(op1) + ',' + hex(op2) + ']');
}

function doDecode(sheet) {
  var op = getCell(sheet, ROW_IR_OP), o1 = getCell(sheet, ROW_IR1), o2 = getCell(sheet, ROW_IR2);
  var m = mnemonic(op, o1, o2);
  setCell(sheet, ROW_MNEMO, m);
  appendLog('DECODE', 'Instrucción: ' + m);
}

var EXEC_RESULT = 0;
function doExecute(sheet) {
  var op = getCell(sheet, ROW_IR_OP), o1 = getCell(sheet, ROW_IR1), o2 = getCell(sheet, ROW_IR2);
  var detalle = '';
  if (op === 0x01) { EXEC_RESULT = o2; detalle = 'Preparado ' + regName(o1) + ' <- ' + hex(o2); }
  else if (op === 0x05) {
    var raw = getReg(sheet, o1) + getReg(sheet, o2);
    EXEC_RESULT = raw % 256;
    setCell(sheet, ROW_ZF, EXEC_RESULT === 0 ? 1 : 0);
    setCell(sheet, ROW_CF, raw > 255 ? 1 : 0);
    setCell(sheet, ROW_SF, (EXEC_RESULT & 0x80) ? 1 : 0);
    detalle = 'ALU: ' + regName(o1) + ' + ' + regName(o2);
  } else if (op === 0xFF) { detalle = 'HLT: se detendrá en Store'; }
  appendLog('EXECUTE', detalle);
}

function doStore(sheet) {
  var op = getCell(sheet, ROW_IR_OP), o1 = getCell(sheet, ROW_IR1);
  var detalle = '';
  if (op === 0x01 || op === 0x05) { setReg(sheet, o1, EXEC_RESULT); detalle = regName(o1) + ' = ' + hex(EXEC_RESULT); }
  else if (op === 0xFF) { setCell(sheet, ROW_ESTADO, 'DETENIDO'); detalle = 'CPU DETENIDA'; }
  appendLog('STORE', detalle);
}

function mnemonic(op, o1, o2) {
  if (op === 0x01) return 'MOV ' + regName(o1) + ', ' + hex(o2);
  if (op === 0x05) return 'ADD ' + regName(o1) + ', ' + regName(o2);
  if (op === 0xFF) return 'HLT';
  return 'DESCONOCIDO (' + hex(op) + ')';
}

function appendLog(fase, texto) {
  var log = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
  var fila = log.getLastRow() + 1;
  log.getRange(fila, 1, 1, 3).setValues([[fila - 1, fase, texto]]);
}