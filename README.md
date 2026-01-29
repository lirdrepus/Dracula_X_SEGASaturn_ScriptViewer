# Castlevania: Symphony of the Night (Saturn JP) – CHR Font Viewer & Text Layout Preview Tool

A lightweight utility designed specifically for **Castlevania: Symphony of the Night (Saturn, Japanese version)** to visualize in‑game text, decode PTN script data, preview tile‑based layouts, and manage text‑to‑font address mappings.  
This tool is intended for ROM hackers, fan translators, and developers working with Saturn‑era CHR font graphics and PTN text encoding.

---

## Features

### 🔹 Text ↔ Font Address Mapping
- Add, delete, import, and export mapping pairs
- CSV import support
- Internal arrays store all address pairs (`TXTadd()` / `TILEadd()`)

### 🔹 CHR Font Rendering
- Reads 32‑byte tiles from `.CHR` font files used by the Saturn version of SOTN
- Supports two flip modes:
  - **8‑block flip** (reverse 8 × 4‑byte blocks)
  - **4‑byte flip** (reverse inside each block and swap pixel values 1 ↔ 16)
- Renders 8×8 tiles using VB6 `Line` graphics

### 🔹 PTN Script Decoding
- Reads `.PTN` encoded text files from the Saturn game data
- Each character uses 2 bytes:
  - Byte 1 → flip mode  
  - Byte 2 → tile index
- Renders full text pages in an 8×16 tile layout
- Useful for previewing in‑game dialogue, menus, and system text

### 🔹 Language Switching
- One‑click UI toggle (Chinese ↔ English)
- Implemented via Caption/Tag swapping for all controls

### 🔹 Data Export
- Saves mapping table to a text file for further editing or ROM hacking workflows

---

## File Types

| Extension | Description |
|----------|-------------|
| `.CHR` | 32‑byte tile font file (8×8 pixels per tile) used by SOTN Saturn |
| `.PTN` | Encoded text file (flip mode + tile index pairs) |
| `.CSV` | Mapping table (text address, tile address) |

---

## How It Works

1. **User enters or imports text/tile address pairs**  
   Stored in `TXTadd()` and `TILEadd()`.

2. **User selects a text address and clicks “PrintOut”**  
   Program reads the `.PTN` script starting from the selected address.

3. **PTN data is decoded**  
   Every 2 bytes → flip mode + tile index.

4. **Each tile is drawn using `fPrint()`**  
   - Reads tile from `.CHR`  
   - Applies flip mode  
   - Draws pixels on the form

5. **Full text page is rendered**  
   Tiles arranged in an 8×16 layout, matching the Saturn game's text rendering style.

---

## Requirements

- Visual Basic 6.0
- CHR font file extracted from SOTN Saturn
- PTN text file extracted from SOTN Saturn
- Optional CSV mapping file

---

## Notes

- Tile size (scaling) is configurable.
- This tool is tailored for **Castlevania: Symphony of the Night (Saturn JP)** 
- Language switching is implemented using Caption/Tag swapping for simplicity.

---

## Purpose

This tool was created to assist with:
- Reverse‑engineering SOTN Saturn’s text system  
- Previewing in‑game Japanese text layouts  
- Debugging CHR font graphics  

It provides a fast and visual way to inspect how the game maps PTN script data to CHR tiles.

---
