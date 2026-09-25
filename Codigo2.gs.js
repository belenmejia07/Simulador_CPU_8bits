/*
 * SIMULADOR DE CPU DE 8 BITS - Google Apps Script
 * Parcial 1 - Arquitectura de Computadoras
 *
 * TARJETA 1: paneles Control Unit / Registers separados y coloreados,
 * banner de FASE ACTUAL, contador de PASO.
 *
 * Ciclo de instrucción: FETCH -> DECODE -> EXECUTE -> STORE
 * Cada clic en "Step" avanza UNA fase (no una instrucción completa).
 *
 * Formato de instrucción: 3 bytes fijos [opcode, operando1, operando2].
 *
 * ISA implementada:
 *   0x01 MOV reg, imm     op1=regDestino  op2=valor inmediato
 *   0x02 MOV reg, reg     op1=regDestino  op2=regOrigen
 *   0x03 LOAD reg, addr   op1=regDestino  op2=dirección memoria
 *   0x04 STORE addr, reg  op1=dirección   op2=regOrigen
 *   0x05 ADD reg, reg     op1=regDestino(y origen1) op2=regOrigen2
 *   0x06 SUB reg, reg     igual que ADD
 *   0x07 INC reg          op1=reg
 *   0x08 DEC reg          op1=reg
 *   0x09 CMP reg, reg     op1=reg1 op2=reg2 (solo actualiza flags)
 *   0x0A JMP addr         op1=dirección destino
 *   0x0B JZ addr          salta si ZF=1
 *   0x0C JNZ addr         salta si ZF=0
 *   0xFF HLT              detiene la CPU
 *
 * Códigos de registro: 0 = AX, 1 = BX
 *
 * Mapa de memoria (256 direcciones, 00h-FFh):
 *   0x00 - 0x1F  Segmento de CÓDIGO (32 bytes)
 *   0x20 - 0xFF  Segmento de DATOS
 */

// ---------------------- CONFIGURACIÓN DE CELDAS ----------------------
var SHEET_NAME = 'CPU';
var LOG_SHEET_NAME = 'Log';

var MEM_ROW0 = 3;
var MEM_COL0 = 2;

var REG_COL_LABEL = 19;
var REG_COL_VALUE = 20;

var ROW_PASO = 1;
var ROW_FASE_TEXTO = 2;
var ROW_CU_HEADER = 3;
var ROW_PC = 4, ROW_IR_OP = 5, ROW_IR1 = 6, ROW_IR2 = 7, ROW_MAR = 8, ROW_MDR = 9;
var ROW_REG_HEADER = 10;
var ROW_AX = 11, ROW_BX = 12;
var ROW_ZF = 13, ROW_CF = 14, ROW_SF = 15;
var ROW_ESTADO = 16;
var ROW_FASE = 17;
var ROW_MNEMO = 18;

// Panel "celda activa" (detalle Hex/Bin/Dec de la dirección en MAR) — sin chocar con ROW_MNEMO
var ROW_DET_ADDR = 20, ROW_DET_HEX = 21, ROW_DET_BIN = 22, ROW_DET_DEC = 23;

var COLOR_FASE = ['#B5D4F4', '#FAC775', '#F0997B', '#9FE1CB']; // Fetch, Decode, Execute, Store
var NOMBRES_FASE = ['FETCH', 'DECODE', 'EXECUTE', 'STORE'];

var PROGRAMA = [
  [0x00, 0x01, 0x00, 0x00], // MOV AX,0
  [0x03, 0x01, 0x01, 0x05], // MOV BX,5
  [0x06, 0x05, 0x00, 0x01], // ADD AX,BX
  [0x09, 0x08, 0x01, 0x00], // DEC BX
  [0x0C, 0x0C, 0x06, 0x00], // JNZ 0x06
  [0x0F, 0x04, 0x20, 0x00], // STORE 0x20,AX
  [0x12, 0xFF, 0x00, 0x00]  // HLT
];

var COLOR_FETCH = '#B5D4F4';
var COLOR_DECODE = '#FAC775';
var COLOR_EXECUTE = '#F0997B';
var COLOR_STORE = '#9FE1CB';
var COLOR_CODE_SEG = '#EEEDFE';
var COLOR_DATA_SEG = '#F1EFE8';

// ---------------------- MENÚ ----------------------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('CPU Simulador')
    .addItem('1. Inicializar todo (ejecutar primero)', 'initMemory')
    .addSeparator()
    .addItem('Step (avanzar una fase)', 'stepCPU')
    .addItem('Run (ejecutar hasta HLT)', 'runCPU')
    .addItem('Reset', 'resetCPU')
    .addItem('Recargar programa demo', 'loadProgram')
    .addToUi();
}

// ---------------------- SETUP INICIAL (correr una sola vez) ----------------------
function initMemory() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  sheet.clear();
  sheet.getRange(1, 1, 1, 1).setValue('Simulador de CPU 8 bits — Memoria RAM (00h-FFh)').setFontWeight('bold');

  for (var c = 0; c < 16; c++) {
    sheet.getRange(MEM_ROW0 - 1, MEM_COL0 + c).setValue(c.toString(16).toUpperCase()).setFontWeight('bold').setHorizontalAlignment('center');
  }
  for (var r = 0; r < 16; r++) {
    sheet.getRange(MEM_ROW0 + r, MEM_COL0 - 1).setValue((r * 16).toString(16).toUpperCase() + 'h').setFontWeight('bold');
    for (var c2 = 0; c2 < 16; c2++) {
      var addr = r * 16 + c2;
      var cell = sheet.getRange(MEM_ROW0 + r, MEM_COL0 + c2);
      cell.setValue(0).setHorizontalAlignment('center').setBorder(true, true, true, true, true, true);
      cell.setBackground(addr < 0x20 ? COLOR_CODE_SEG : COLOR_DATA_SEG);
    }
  }
  sheet.getRange(MEM_ROW0 - 2, MEM_COL0).setValue('Azul claro = segmento de código (00-1F)   |   Gris = segmento de datos (20-FF)').setFontStyle('italic');
  sheet.setColumnWidths(MEM_COL0, 16, 35);

  var log = ss.getSheetByName(LOG_SHEET_NAME) || ss.insertSheet(LOG_SHEET_NAME);
  log.clear();
  log.getRange(1, 1, 1, 3).setValues([['Paso', 'Fase', 'Detalle']]).setFontWeight('bold');
  log.setColumnWidth(3, 500);

  initRegistros();  // <-- FALTABA esta llamada (Problema 1)
  resetCPU();
}

// ---------------------- PANEL DE REGISTROS / CONTROL UNIT / FASE / PASO ----------------------
function initRegistros() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  sheet.getRange(ROW_PASO, REG_COL_LABEL).setValue('PASO').setFontWeight('bold');
  sheet.getRange(ROW_PASO, REG_COL_VALUE).setBorder(true, true, true, true, true, true).setFontWeight('bold');

  sheet.getRange(ROW_FASE_TEXTO, REG_COL_LABEL, 1, 2).merge().setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');

  sheet.getRange(ROW_CU_HEADER, REG_COL_LABEL, 1, 2).merge().setValue('CONTROL UNIT (CU)').setFontWeight('bold').setBackground('#F6C9CE').setHorizontalAlignment('center');
  [[ROW_PC, 'PC'], [ROW_IR_OP, 'IR (opcode)'], [ROW_IR1, 'IR (op1)'], [ROW_IR2, 'IR (op2)'], [ROW_MAR, 'MAR'], [ROW_MDR, 'MDR']]
    .forEach(function (item) {
      sheet.getRange(item[0], REG_COL_LABEL).setValue(item[1]).setFontWeight('bold');
      sheet.getRange(item[0], REG_COL_VALUE).setBorder(true, true, true, true, true, true);
    });

  sheet.getRange(ROW_REG_HEADER, REG_COL_LABEL, 1, 2).merge().setValue('REGISTERS').setFontWeight('bold').setBackground('#C7EAD1').setHorizontalAlignment('center');
  [[ROW_AX, 'AX'], [ROW_BX, 'BX'], [ROW_ZF, 'ZF'], [ROW_CF, 'CF'], [ROW_SF, 'SF'], [ROW_ESTADO, 'Estado'], [ROW_FASE, 'Fase (interno)'], [ROW_MNEMO, 'Instrucción']]
    .forEach(function (item) {
      sheet.getRange(item[0], REG_COL_LABEL).setValue(item[1]).setFontWeight('bold');
      sheet.getRange(item[0], REG_COL_VALUE).setBorder(true, true, true, true, true, true);
    });

  // Panel de celda activa — vive aquí, ya no duplicado en initMemory (Problema 3)
  sheet.getRange(ROW_DET_ADDR - 1, REG_COL_LABEL).setValue('CELDA ACTIVA (según MAR)').setFontWeight('bold');
  ['Dirección', 'Hex', 'Binario', 'Decimal'].forEach(function (t, i) {
    sheet.getRange(ROW_DET_ADDR + i, REG_COL_LABEL).setValue(t);
    sheet.getRange(ROW_DET_ADDR + i, REG_COL_VALUE).setBorder(true, true, true, true, true, true);
  });

  sheet.setColumnWidth(REG_COL_LABEL, 170);
  sheet.setColumnWidth(REG_COL_VALUE, 100);
  resetRegistros();
}

function resetRegistros() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  [ROW_PC, ROW_IR_OP, ROW_IR1, ROW_IR2, ROW_MAR, ROW_MDR, ROW_AX, ROW_BX, ROW_ZF, ROW_CF, ROW_SF, ROW_FASE]
    .forEach(function (row) { setCell(sheet, row, 0); });
  setCell(sheet, ROW_ESTADO, 'LISTO');
  setCell(sheet, ROW_PASO, 0);
  setFaseDisplay(sheet, 0);
}

function setFaseDisplay(sheet, faseNum) {
  sheet.getRange(ROW_FASE_TEXTO, REG_COL_LABEL).setValue('FASE: ' + NOMBRES_FASE[faseNum]).setBackground(COLOR_FASE[faseNum]);
}

function updatePasoCounter() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  setCell(ss.getSheetByName(SHEET_NAME), ROW_PASO, Math.max(ss.getSheetByName(LOG_SHEET_NAME).getLastRow() - 1, 0));
}

// ---------------------- CARGA DE PROGRAMA ----------------------
function loadProgram() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  for (var a = 0; a < 0x20; a++) writeMem(sheet, a, 0);
  PROGRAMA.forEach(function (instr) {
    var addr = instr[0];
    writeMem(sheet, addr, instr[1]);
    writeMem(sheet, addr + 1, instr[2]);
    writeMem(sheet, addr + 2, instr[3]);
  });
}

// ---------------------- RESET (simplificado: reutiliza resetRegistros) ----------------------
function resetCPU() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  clearHighlights(sheet);
  resetRegistros();               // Problema 6: ya no duplica los setCell uno por uno
  setCell(sheet, ROW_MNEMO, '-');
  loadProgram();
  updateDetailPanel(sheet, 0);

  var log = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
  log.getRange(2, 1, Math.max(log.getLastRow() - 1, 1), 3).clearContent();
}

// ---------------------- STEP: avanza UNA fase del ciclo ----------------------
function stepCPU() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (getCell(sheet, ROW_ESTADO) === 'DETENIDO') {
    appendLog('-', 'CPU detenida (HLT). Usa Reset para reiniciar.');
    return;
  }
  clearHighlights(sheet);
  var fase = getCell(sheet, ROW_FASE);
  setFaseDisplay(sheet, fase);    // Problema 4: ahora sí se llama
  if (fase === 0) { doFetch(sheet); setCell(sheet, ROW_FASE, 1); }
  else if (fase === 1) { doDecode(sheet); setCell(sheet, ROW_FASE, 2); }
  else if (fase === 2) { doExecute(sheet); setCell(sheet, ROW_FASE, 3); }
  else { doStore(sheet); setCell(sheet, ROW_FASE, 0); }
}

// ---------------------- RUN: ejecuta hasta HLT ----------------------
function runCPU() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var maxIter = 400;
  var i = 0;
  while (getCell(sheet, ROW_ESTADO) !== 'DETENIDO' && i < maxIter) {
    stepCPU();
    SpreadsheetApp.flush();
    Utilities.sleep(250);
    i++;
  }
}

// ---------------------- FASES DEL CICLO ----------------------
function doFetch(sheet) {
  var pc = getCell(sheet, ROW_PC);
  setCell(sheet, ROW_MAR, pc);
  var opcode = readMem(sheet, pc);
  var op1 = readMem(sheet, pc + 1);
  var op2 = readMem(sheet, pc + 2);
  setCell(sheet, ROW_MDR, opcode);
  setCell(sheet, ROW_IR_OP, opcode);
  setCell(sheet, ROW_IR1, op1);
  setCell(sheet, ROW_IR2, op2);
  setCell(sheet, ROW_PC, pc + 3);

  highlightMemRange(sheet, pc, 3, COLOR_FETCH);
  highlightReg(sheet, ROW_PC, COLOR_FETCH);
  highlightReg(sheet, ROW_MAR, COLOR_FETCH);
  highlightReg(sheet, ROW_MDR, COLOR_FETCH);
  updateDetailPanel(sheet, pc);
  appendLog('FETCH', 'PC=' + hex(pc) + ' -> IR=[' + hex(opcode) + ',' + hex(op1) + ',' + hex(op2) + ']  PC actualizado a ' + hex(pc + 3));
}

function doDecode(sheet) {
  var opcode = getCell(sheet, ROW_IR_OP);
  var op1 = getCell(sheet, ROW_IR1);
  var op2 = getCell(sheet, ROW_IR2);
  var mnem = mnemonic(opcode, op1, op2);
  setCell(sheet, ROW_MNEMO, mnem);
  highlightReg(sheet, ROW_IR_OP, COLOR_DECODE);
  highlightReg(sheet, ROW_IR1, COLOR_DECODE);
  highlightReg(sheet, ROW_IR2, COLOR_DECODE);
  appendLog('DECODE', 'Instrucción interpretada: ' + mnem);
}

function doExecute(sheet) {
  var opcode = getCell(sheet, ROW_IR_OP);
  var op1 = getCell(sheet, ROW_IR1);
  var op2 = getCell(sheet, ROW_IR2);
  var detalle = '';

  switch (opcode) {
    case 0x01:
      PropertiesService.getScriptProperties().setProperty('EXEC_RESULT', String(op2));
      detalle = 'Preparado: ' + regName(op1) + ' <- ' + hex(op2);
      highlightReg(sheet, regRow(op1), COLOR_EXECUTE);
      break;
    case 0x02:
      var vSrc = getReg(sheet, op2);
      PropertiesService.getScriptProperties().setProperty('EXEC_RESULT', String(vSrc));
      detalle = 'Preparado: ' + regName(op1) + ' <- ' + regName(op2) + ' (' + hex(vSrc) + ')';
      highlightReg(sheet, regRow(op1), COLOR_EXECUTE);
      highlightReg(sheet, regRow(op2), COLOR_EXECUTE);
      break;
    case 0x03:
      var vMem = readMem(sheet, op2);
      setCell(sheet, ROW_MAR, op2);
      setCell(sheet, ROW_MDR, vMem);
      PropertiesService.getScriptProperties().setProperty('EXEC_RESULT', String(vMem));
      detalle = 'Preparado: ' + regName(op1) + ' <- RAM[' + hex(op2) + '] = ' + hex(vMem);
      highlightMemRange(sheet, op2, 1, COLOR_EXECUTE);
      updateDetailPanel(sheet, op2);
      break;
    case 0x04:
      var vReg = getReg(sheet, op2);
      setCell(sheet, ROW_MAR, op1);
      setCell(sheet, ROW_MDR, vReg);
      PropertiesService.getScriptProperties().setProperty('EXEC_RESULT', String(vReg));
      detalle = 'Preparado: RAM[' + hex(op1) + '] <- ' + regName(op2) + ' (' + hex(vReg) + ')';
      highlightReg(sheet, regRow(op2), COLOR_EXECUTE);
      break;
    case 0x05:
      execALU(sheet, op1, op2, 'ADD');
      detalle = 'ALU: ' + regName(op1) + ' + ' + regName(op2);
      highlightReg(sheet, regRow(op1), COLOR_EXECUTE);
      highlightReg(sheet, regRow(op2), COLOR_EXECUTE);
      break;
    case 0x06:
      execALU(sheet, op1, op2, 'SUB');
      detalle = 'ALU: ' + regName(op1) + ' - ' + regName(op2);
      highlightReg(sheet, regRow(op1), COLOR_EXECUTE);
      highlightReg(sheet, regRow(op2), COLOR_EXECUTE);
      break;
    case 0x07:
      execALU(sheet, op1, -1, 'INC');
      detalle = 'ALU: ' + regName(op1) + ' + 1';
      highlightReg(sheet, regRow(op1), COLOR_EXECUTE);
      break;
    case 0x08:
      execALU(sheet, op1, -1, 'DEC');
      detalle = 'ALU: ' + regName(op1) + ' - 1';
      highlightReg(sheet, regRow(op1), COLOR_EXECUTE);
      break;
    case 0x09:
      execALU(sheet, op1, op2, 'CMP');
      detalle = 'ALU: comparar ' + regName(op1) + ' con ' + regName(op2) + ' (solo flags)';
      highlightReg(sheet, regRow(op1), COLOR_EXECUTE);
      highlightReg(sheet, regRow(op2), COLOR_EXECUTE);
      break;
    case 0x0A:
      setCell(sheet, ROW_PC, op1);
      detalle = 'Salto incondicional -> PC=' + hex(op1);
      highlightReg(sheet, ROW_PC, COLOR_EXECUTE);
      break;
    case 0x0B:
      if (getCell(sheet, ROW_ZF) === 1) { setCell(sheet, ROW_PC, op1); detalle = 'ZF=1 -> salta a ' + hex(op1); }
      else { detalle = 'ZF=0 -> no salta'; }
      highlightReg(sheet, ROW_ZF, COLOR_EXECUTE);
      break;
    case 0x0C:
      if (getCell(sheet, ROW_ZF) === 0) { setCell(sheet, ROW_PC, op1); detalle = 'ZF=0 -> salta a ' + hex(op1); }
      else { detalle = 'ZF=1 -> no salta'; }
      highlightReg(sheet, ROW_ZF, COLOR_EXECUTE);
      break;
    case 0xFF:
      detalle = 'HLT: la CPU se detendrá en la fase Store';
      break;
    default:
      detalle = 'Opcode desconocido: ' + hex(opcode);
  }
  appendLog('EXECUTE', detalle);
}

function doStore(sheet) {
  var opcode = getCell(sheet, ROW_IR_OP);
  var op1 = getCell(sheet, ROW_IR1);
  var op2 = getCell(sheet, ROW_IR2);
  var props = PropertiesService.getScriptProperties();
  var result = props.getProperty('EXEC_RESULT');
  var detalle = '';

  switch (opcode) {
    case 0x01: case 0x02:
      setReg(sheet, op1, Number(result));
      detalle = regName(op1) + ' = ' + hex(Number(result));
      highlightReg(sheet, regRow(op1), COLOR_STORE);
      break;
    case 0x03:
      setReg(sheet, op1, Number(result));
      detalle = regName(op1) + ' = ' + hex(Number(result)) + ' (desde memoria)';
      highlightReg(sheet, regRow(op1), COLOR_STORE);
      break;
    case 0x04:
      writeMem(sheet, op1, Number(result));
      detalle = 'RAM[' + hex(op1) + '] = ' + hex(Number(result));
      highlightMemRange(sheet, op1, 1, COLOR_STORE);
      updateDetailPanel(sheet, op1);
      break;
    case 0x05: case 0x06: case 0x07: case 0x08:
      setReg(sheet, op1, Number(result));
      detalle = regName(op1) + ' = ' + hex(Number(result));
      highlightReg(sheet, regRow(op1), COLOR_STORE);
      break;
    case 0x09: case 0x0A: case 0x0B: case 0x0C:
      detalle = 'Sin escritura en esta fase (solo control de flujo / flags)';
      break;
    case 0xFF:
      setCell(sheet, ROW_ESTADO, 'DETENIDO');
      detalle = 'CPU DETENIDA';
      break;
    default:
      detalle = '-';
  }
  props.deleteProperty('EXEC_RESULT');
  appendLog('STORE', detalle);
}

// ---------------------- ALU ----------------------
function execALU(sheet, op1, op2, tipo) {
  var a = getReg(sheet, op1);
  var b = (tipo === 'INC' || tipo === 'DEC') ? 1 : getReg(sheet, op2);
  var raw, result;
  var cf = 0;

  if (tipo === 'ADD') { raw = a + b; }
  else if (tipo === 'SUB' || tipo === 'CMP') { raw = a - b; }
  else if (tipo === 'INC') { raw = a + 1; }
  else if (tipo === 'DEC') { raw = a - 1; }

  if (raw > 255) cf = 1;
  if (raw < 0) cf = 1;
  result = ((raw % 256) + 256) % 256;

  setCell(sheet, ROW_ZF, result === 0 ? 1 : 0);
  setCell(sheet, ROW_CF, cf);
  setCell(sheet, ROW_SF, (result & 0x80) ? 1 : 0);

  if (tipo !== 'CMP') {
    PropertiesService.getScriptProperties().setProperty('EXEC_RESULT', String(result));
  } else {
    PropertiesService.getScriptProperties().setProperty('EXEC_RESULT', String(a));
  }
}

// ---------------------- MEMORIA ----------------------
function readMem(sheet, addr) {
  var r = Math.floor(addr / 16), c = addr % 16;
  return sheet.getRange(MEM_ROW0 + r, MEM_COL0 + c).getValue();
}
function writeMem(sheet, addr, val) {
  var r = Math.floor(addr / 16), c = addr % 16;
  sheet.getRange(MEM_ROW0 + r, MEM_COL0 + c).setValue(val);
}
function highlightMemRange(sheet, addrStart, len, color) {
  for (var i = 0; i < len; i++) {
    var addr = addrStart + i;
    var r = Math.floor(addr / 16), c = addr % 16;
    sheet.getRange(MEM_ROW0 + r, MEM_COL0 + c).setBackground(color);
  }
}
function clearHighlights(sheet) {
  for (var addr = 0; addr < 256; addr++) {
    var r = Math.floor(addr / 16), c = addr % 16;
    sheet.getRange(MEM_ROW0 + r, MEM_COL0 + c).setBackground(addr < 0x20 ? COLOR_CODE_SEG : COLOR_DATA_SEG);
  }
  [ROW_PC, ROW_IR_OP, ROW_IR1, ROW_IR2, ROW_MAR, ROW_MDR, ROW_AX, ROW_BX, ROW_ZF, ROW_CF, ROW_SF].forEach(function (row) {
    sheet.getRange(row, REG_COL_VALUE).setBackground(null);
  });
}
function highlightReg(sheet, row, color) {
  sheet.getRange(row, REG_COL_VALUE).setBackground(color);
}

// ---------------------- REGISTROS ----------------------
function getReg(sheet, code) { return getCell(sheet, code === 0 ? ROW_AX : ROW_BX); }
function setReg(sheet, code, val) { setCell(sheet, code === 0 ? ROW_AX : ROW_BX, val); }
function regRow(code) { return code === 0 ? ROW_AX : ROW_BX; }
function regName(code) { return code === 0 ? 'AX' : 'BX'; }

function getCell(sheet, row) { return sheet.getRange(row, REG_COL_VALUE).getValue(); }
function setCell(sheet, row, val) { sheet.getRange(row, REG_COL_VALUE).setValue(val); }

// ---------------------- PANEL DE CELDA ACTIVA ----------------------
function updateDetailPanel(sheet, addr) {
  var val = readMem(sheet, addr);
  sheet.getRange(ROW_DET_ADDR, REG_COL_VALUE).setValue(hex(addr));
  sheet.getRange(ROW_DET_HEX, REG_COL_VALUE).setValue(hex(val));
  sheet.getRange(ROW_DET_BIN, REG_COL_VALUE).setValue(bin(val));
  sheet.getRange(ROW_DET_DEC, REG_COL_VALUE).setValue(val);
}

// ---------------------- UTILIDADES ----------------------
function hex(n) { return '0x' + Number(n).toString(16).toUpperCase().padStart(2, '0'); }
function bin(n) { return Number(n).toString(2).padStart(8, '0'); }

function mnemonic(opcode, op1, op2) {
  switch (opcode) {
    case 0x01: return 'MOV ' + regName(op1) + ', ' + hex(op2);
    case 0x02: return 'MOV ' + regName(op1) + ', ' + regName(op2);
    case 0x03: return 'LOAD ' + regName(op1) + ', [' + hex(op2) + ']';
    case 0x04: return 'STORE [' + hex(op1) + '], ' + regName(op2);
    case 0x05: return 'ADD ' + regName(op1) + ', ' + regName(op2);
    case 0x06: return 'SUB ' + regName(op1) + ', ' + regName(op2);
    case 0x07: return 'INC ' + regName(op1);
    case 0x08: return 'DEC ' + regName(op1);
    case 0x09: return 'CMP ' + regName(op1) + ', ' + regName(op2);
    case 0x0A: return 'JMP ' + hex(op1);
    case 0x0B: return 'JZ ' + hex(op1);
    case 0x0C: return 'JNZ ' + hex(op1);
    case 0xFF: return 'HLT';
    default: return 'DESCONOCIDO (' + hex(opcode) + ')';
  }
}

// ---------------------- LOG ----------------------
function appendLog(fase, texto) {
  var log = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
  var fila = log.getLastRow() + 1;
  log.getRange(fila, 1, 1, 3).setValues([[fila - 1, fase, texto]]);
  updatePasoCounter();  // Problema 5: ahora sí se llama
}