# Simulador de CPU de 8 bits — Google Sheets / Apps Script

Parcial 1 — Arquitectura de Computadoras, UCB Santa Cruz

## 1. Arquitectura

### Mapa de memoria (256 direcciones, 00h-FFh)
- **Segmento de código:** 0x00 - 0x1F (32 bytes) — instrucciones del programa
- **Segmento de datos:** 0x20 - 0xFF — resultados almacenados con `STORE`

### Registros
| Registro | Función |
|---|---|
| PC | Dirección de la próxima instrucción |
| IR (opcode, op1, op2) | Instrucción actual, decodificada en 3 partes |
| MAR | Dirección de memoria que se está accediendo |
| MDR | Dato leído/escrito de esa dirección |
| AX, BX | Registros de propósito general (código 0 y 1) |
| ZF, CF, SF | Flags: cero, acarreo/overflow, signo |

### Decisión de diseño: instrucciones de longitud fija (3 bytes)
Se eligió que **toda** instrucción ocupe siempre 3 bytes (opcode + operando1 + operando2), en vez de longitud variable. Esto desperdicia memoria en instrucciones sin operandos (como `HLT`), pero simplifica el cálculo del PC (`PC = PC + 3` siempre) y reduce el riesgo de errores en el ciclo Fetch — una prioridad dado que el simulador debe explicarse y modificarse en vivo durante la defensa.

### Ciclo de instrucción
Cada clic en **Step** avanza una fase (no una instrucción completa), mostrando explícitamente:
1. **Fetch:** PC → MAR; `Read(MAR)` → MDR → IR; PC += 3
2. **Decode:** IR se traduce a mnemónico legible (ej. `ADD AX, BX`)
3. **Execute:** la ALU calcula (o se resuelve el salto), actualiza flags
4. **Store:** el resultado se escribe en el registro o en memoria con `Write`

## 2. Repertorio de instrucciones (ISA)

| Mnemónico | Opcode | Operandos | Descripción |
|---|---|---|---|
| MOV reg, imm | 0x01 | reg, valor | Copia un valor inmediato a un registro |
| MOV reg, reg | 0x02 | reg, reg | Copia el valor de un registro a otro |
| LOAD reg, [dir] | 0x03 | reg, dirección | Carga desde memoria a un registro |
| STORE [dir], reg | 0x04 | dirección, reg | Guarda un registro en memoria |
| ADD reg, reg | 0x05 | reg, reg | Suma (actualiza ZF, CF, SF) |
| SUB reg, reg | 0x06 | reg, reg | Resta (actualiza flags) |
| INC reg | 0x07 | reg | Incrementa en 1 |
| DEC reg | 0x08 | reg | Decrementa en 1 |
| CMP reg, reg | 0x09 | reg, reg | Compara (solo actualiza flags) |
| JMP dir | 0x0A | dirección | Salto incondicional |
| JZ dir | 0x0B | dirección | Salta si ZF=1 |
| JNZ dir | 0x0C | dirección | Salta si ZF=0 |
| HLT | 0xFF | — | Detiene la CPU |

Códigos de registro: `0 = AX`, `1 = BX`.

## 3. Manual de usuario

1. Abre la hoja de cálculo y ve al menú **"CPU Simulador"**
2. **"1. Inicializar todo"** — construye la memoria, registros, ALU y Cola de Instrucciones, y crea la pestaña **"Programa"**
3. Edita la pestaña **Programa** (columna "Código") para escribir tus propias instrucciones en texto — el simulador las traduce solas a bytes al recalcular
4. Usa los botones (o el menú): **Step** avanza una fase, **Run** ejecuta hasta HLT, **Acelerar/Desacelerar** ajustan la velocidad de Run, **Reset** recarga todo desde la pestaña Programa
5. La hoja **"Log"** guarda el historial completo de cada micro-operación

## 4. Traza del programa demostrativo

Programa: suma 5+4+3+2+1 usando un bucle, y guarda el resultado en memoria.

```
0x00  MOV AX, 0        ; AX = 0
0x03  MOV BX, 5        ; BX = 5
0x06  ADD AX, BX        <- LOOP
0x09  DEC BX
0x0C  JNZ 0x06
0x0F  STORE 0x20, AX
0x12  HLT
```

| Iteración | AX antes | BX antes | AX después (ADD) | BX después (DEC) | ¿Salta? |
|---|---|---|---|---|---|
| 1 | 0 | 5 | 5 | 4 | Sí (BX≠0) |
| 2 | 5 | 4 | 9 | 3 | Sí |
| 3 | 9 | 3 | 12 | 2 | Sí |
| 4 | 12 | 2 | 14 | 1 | Sí |
| 5 | 14 | 1 | 15 | 0 | No — sale del bucle |

Resultado final: **AX = 15**, guardado en la dirección **0x20** de memoria.

## 5. Casos borde probados

| Caso | Resultado |
|---|---|
| Overflow (250 + 10) | AX = 4, CF = 1 |
| Resultado cero (5 - 5) | AX = 0, ZF = 1 |
| Resultado negativo (3 - 5) | AX = 254 (0xFE), SF = 1, CF = 1 (el diseño marca CF también en "borrow", no solo en overflow hacia arriba) |

## 6. Alcance del proyecto

Este simulador implementa un ciclo de instrucción secuencial de un solo programa, sin memoria virtual, sin pipeline y sin unidades de entrada/salida — fuera del alcance de este examen, correspondiente a temas de sistemas operativos y arquitecturas avanzadas que se cubren en el Parcial 2.
