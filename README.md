# ERPIE PriceFinder

**Expected Resale Price calculator voor refurbished IT equipment — door PlanBit**

## Features

- **Quick Quote** — Bereken direct de ERP voor een enkel apparaat
- **Batch Upload** — Analyseer een volledige Excel/CSV lijst met apparaten
- **HTML Report** — Genereer een professioneel rapport voor klanten
- **Pricing Engine** — Gebaseerd op Rule Book v8.4 met CPU-generatie logica

## Tech Stack

- Backend: Node.js + Express
- Frontend: Vanilla HTML/CSS/JS (geen frameworks)
- Excel parsing: `xlsx`
- File upload: `multer`

## Lokaal draaien

```bash
npm install
npm start
# → http://localhost:8000
```

## API Routes

### `POST /api/quote`
Bereken prijs voor één apparaat.

**Body:**
```json
{
  "model": "HP EliteBook 840 G8",
  "ram": 16,
  "ssd": 256,
  "grade": "A",
  "battery": "good",
  "region": "EU"
}
```

### `POST /api/analyze`
Batch analyse van Excel/CSV bestand.

**Form-data:**
- `file` — xlsx of csv bestand
- `region` — EU / UK / INTL

### `POST /api/report`
Genereer HTML rapport.

**Body:**
```json
{
  "dealName": "Deal XYZ Corp",
  "devices": [...],
  "summary": {...}
}
```

## Pricing Logic

Gebaseerd op **Rule Book v8.4**:

1. **Base price** per CPU-generatie (Gen4–Gen14)
2. **RAM correctie** (baseline 8GB)
3. **SSD correctie** (baseline 256GB)
4. **Battery correctie** (good/bad/missing)
5. **Grade vermenigvuldiger** (A1=1.40 t/m D=0.40)
6. **Brand tier vermenigvuldiger**
7. **WATCH caps** (Gen8 max €120, Gen9 max €140)
8. **Region adjustment** (UK -16%, INTL -15%)

### Status classificatie
- **GO** — Gen 10 en nieuwer
- **WATCH** — Gen 8-9 (met prijscap)
- **NO-GO** — Gen 7 en ouder

## Deploy op Render.com

1. Push naar GitHub
2. Maak nieuw Web Service aan op render.com
3. Koppel je repo — `render.yaml` regelt de rest

## Test Resultaten (verwacht)

| Model | RAM | SSD | Grade | Verwacht |
|-------|-----|-----|-------|----------|
| HP EliteBook 840 G8 | 16GB | 256GB | A | ~€275, GO |
| MacBook Air M1 | 8GB | 256GB | A | €450, GO |
| Dell Latitude 7490 | 8GB | 256GB | A | €120, WATCH (cap) |
