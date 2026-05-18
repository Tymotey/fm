# Forge Master — Build Calculator & Comparison

Documentation for **`ForgeMaster_Build_Calculator.xlsx`** and its generator
**`scripts/build_excel_calculator.py`**.

The workbook compares two builds side-by-side and shows the exact percentage
change of every stat and every derived metric. Combat formulas mirror
[`src/utils/statEngine.ts`](src/utils/statEngine.ts) — function
`finalizeCalculation`.

---

## 1. Files

| File | Purpose |
|---|---|
| `ForgeMaster_Build_Calculator.xlsx` | The calculator. Open in Excel / Google Sheets / LibreOffice. |
| `scripts/build_excel_calculator.py` | Generator. Re-run to rebuild the workbook after a formula change. |
| `ForgeMaster_Build_Calculator.md` | This document. |

Rebuild the workbook:

```bash
python scripts/build_excel_calculator.py
```

The script needs `openpyxl` (`pip install openpyxl`).

---

## 2. Workbook layout

Three sheets.

### `Calculator`

The main table starts at the top of the sheet. Game constants live in a
separate block on the **right** (columns O–Q).

| Col | Meaning |
|---|---|
| **A** | Stat / metric label |
| **B** | Build A — the baseline |
| **C, E, G, I, K** | Builds B, C, D, E, F |
| **D, F, H, J, L** | `vs A` — each build's % difference from Build A on that row |
| **M** | Max-per-roll cap / notes |
| **O – Q** | Game constants block (label, value, note) |

Every build B–F is followed by a narrow **`vs A`** column: it shows that
build's percentage difference from Build A for the same row — **green** when
higher than A, **red** when lower. Build A is the baseline and has no `vs A`
column. The **MATCHUP** section runs a 1-on-1 duel for **Build A vs Build B
only**; C–F are comparison columns and do not duel.

Cell colours:

- 🟡 **Yellow** — input you type
- ⬜ **Grey** — computed by formula
- 🟩 **Green** — key result (Power, Crit Hit, Real APS, Real DPS, HPS, Winner)
- 🟦 **Blue** — game constant
- 🟢/🔴 **`vs A` cell** — green = better than Build A, red = worse

### `Documentation`
An in-workbook copy of the formulas and the maintenance guide.

### `Stat Reference`
The 13 stats with their max value per single substat roll, plus the build-setup
notes.

---

## 3. What you enter (yellow cells)

There are **6 build columns** (Build A = sheet column B; Builds B–F in C, E, G,
I, K — interleaved with their `vs A` columns). Fill the yellow cells in each
build's column. The sheet keeps **only the inputs that change a result** — per
build that is just the Build Setup and the 13 stat %s.

**Build Setup**

| Input | Notes |
|---|---|
| Weapon Type | `Melee` or `Range` — pick from the dropdown |
| Max Damage | The final **Damage** from the in-game character screen |
| Max Health | The final **Health** from the in-game character screen |

Max Damage and Max Health accept either a plain number **or a short form with an
`M` / `B` / `T` suffix**:

| You type | Value used |
|---|---|
| `1.2B` | 1,200,000,000 |
| `640M` | 640,000,000 |
| `5T` | 5,000,000,000,000 |
| `41080` | 41,080 |

The grey **`Max Damage (number)`** / **`Max Health (number)`** row directly below
each input shows the parsed value the formulas actually use.

These are the **final totals**, so all gear, pet, mount, skill and ascension
bonuses are already baked into them — the calculator does not add them again.
That is why pet / mount / skill selectors were removed: they did not change any
result.

**Substat Totals** — the 13 stats as percentage points **summed across all 24
substat rolls** (12 items × 2). Type `403.8` for 403.8 %, *not* `4.038`. Totals
legitimately exceed the single-roll cap in column D, because they are a sum.

---

## 4. The formulas

All names match the row labels on the `Calculator` sheet.

### Power

```
POWER = ( (MaxDamage − 10) × 8 + (MaxHealth − 80) ) × 3
```

Reverse-engineered from the game binary. `10` / `80` are base damage / health,
`8` weights damage against health, `3` is the final scale.

### Damage & crit

```
Crit Damage Mult    = 1 + 0.20 + CritDamage%
Critical Hit Damage = Max Damage × Crit Damage Mult
Crit Mult on hit    = 1 + CritChance × (Crit Damage Mult − 1)
Double Mult on hit  = 1 + DoubleChance%
```

### Attack Speed & breakpoints

Weapon Duration and Windup come from the Melee or Ranged constants, selected by
Weapon Type.

```
Attack Speed Mult = 1 + AttackSpeed%             (122% → 2.22×)
Theoretical APS   = Attack Speed Mult / Weapon Duration
Stepped Windup    = FLOOR( Windup / SpeedMult , 0.1 )
Stepped Recovery  = FLOOR( (Duration − Windup) / SpeedMult , 0.1 )
Stepped Cycle     = MAX( 0.4 , Stepped Windup + Stepped Recovery + 0.2 )
Double-Hit Cycle  = Stepped Cycle + MAX( 0.1 , FLOOR( 0.25 / SpeedMult , 0.1 ) )
Avg Real Cycle    = (1 − Double%) × Cycle + Double% × DoubleHitCycle
Real APS          = (1 + Double%) / Avg Real Cycle
```

The game floors animation times to **0.1 s steps**, so attack speed improves in
jumps (*breakpoints*), not smoothly. *Real APS* is the honest in-combat rate;
*Theoretical APS* ignores breakpoints.

### Skill cooldown

```
Cooldown Reduction = Skill Cooldown%
Cooldown Mult      = MAX( 0.1 , 1 − Reduction )
```

Derived purely from the *Skill Cooldown %* stat — multiply a skill's base
cooldown by *Cooldown Mult* to get its real cooldown.

### DPS & HPS

```
Theoretical Wpn DPS = Max Damage × Theoretical APS × Crit Mult × Double Mult
Real Weapon DPS     = Max Damage × Real APS × Crit Mult   (Double already in Real APS)
Block Factor        = 1 / (1 − MIN(Block%, 95%))
Lifesteal HPS       = Real Weapon DPS × Lifesteal%
TOTAL HPS (real)    = (Health Regen/s + Lifesteal HPS) × Block Factor
```

`Health Regen/s = Max Health × HealthRegen%`.

### Matchup — who wins

The **MATCHUP** section settles a straight 1-on-1 duel between **Build A and
Build B only**. Builds C–F are comparison columns and do not duel.

```
Net DPS dealt   = your Real Weapon DPS − opponent's Total HPS
Time to defeat  = opponent's Max Health / your Net DPS dealt
WINNER          = the build with the shorter Time to defeat
```

If your Net DPS is ≤ 0 the opponent out-heals your damage and *Time to defeat*
shows `never`. Both sides unkillable, or an exact tie → **DRAW**. To duel a
different pair, copy a build's column values into the Build A or Build B column.

This is a deliberately simple model — it ignores attack timing, skill rotations
and who strikes first. It is a fair *relative* comparison, not a frame-perfect
simulation.

---

## 5. Stat reference

13 stats; the value is the **maximum per single substat roll**.

| Stat | Max / roll | Effect |
|---|---|---|
| Crit Damage | 80 | Damage of critical hits |
| Crit Chance | 12 | Chance an attack crits |
| Health Regen | 4 | Health/s as % of max HP |
| Lifesteal | 20 | % of weapon damage healed |
| Double Chance | 20 | Chance of a bonus double hit |
| Damage | 15 | Global damage multiplier |
| Melee Damage | 50 | Extra damage, melee weapons only |
| Range Damage | 15 | Extra damage, ranged weapons only |
| Attack Speed | 40 | Speeds up the attack cycle |
| Skill Cooldown | 7 | Reduces active-skill cooldown |
| Skill Damage | 30 | Multiplier on skill damage & healing |
| Health | 15 | Global health multiplier |
| Block Chance | 5 | Chance to block; raises effective HPS |

**Build setup:** 12 items = 8 equipment + 1 mount + 3 pets, each with 2 substats
→ 24 rolls. Gear, pets, mount, skills and ascension are all already reflected in
the Max Damage / Max Health you read off the character screen, so the calculator
asks only for the inputs that change a result.

---

## 6. Verification

The generated workbook ships with dummy data in all 6 builds and was
recalculated in Excel. Build A (Weapon Type Range, Max Damage `1.2B`, Max Health
`640M`, Crit Dmg 403.8 %, Attack Speed 122 %, Lifesteal 61.7 %) produces:

| Metric | Build A | Build B |
|---|---|---|
| Max Damage (number) | 1,200,000,000 | 1,450,000,000 |
| Power | 30,719,999,520 | 37,139,999,520 |
| Critical Hit Damage | 6,285,600,000 | 7,250,000,000 |
| Real Weapon DPS | 4,807,936,117 | 6,141,176,471 |
| Total HPS (real) | 2,966,496,584 | 3,498,192,844 |

The `M` / `B` / `T` parsing is confirmed (`1.2B` → 1,200,000,000), all 6 build
columns compute, and the A-vs-B duel returns **"BUILD B WINS"**.

---

## 7. Maintaining the workbook

**Do not hand-edit cells for permanent changes** — they are lost on the next
rebuild. Edit the generator instead.

`scripts/build_excel_calculator.py` builds the `Calculator` sheet from a list of
entry helpers:

| Helper | Adds |
|---|---|
| `inp(label, key, vals, note, fmt)` | A yellow input row — `vals` is one value per build, via `six(...)` |
| `out(label, key, fmt, formula, hilite)` | A computed row (formula applied to all 6 builds + their `vs A`) |
| `matchup(label, key, fmt, fa, fb)` / `verdict(...)` | The A-vs-B duel rows |
| `section(label)` / `header()` / `blank()` | Layout rows |

Game constants are **not** in the row flow — they are the `CONSTS` list, drawn
as the block in columns O–Q.

`six(...)` expands to one value per build: `six("Range")` fills all 6,
`six(9, 12, 10, 8, 13, 9)` gives each build its own value.

**Formula tokens** inside `out(...)` / `matchup(...)`:

- `@key` → the cell for that key **in the build's own column**
- `#key` → a game constant, resolved to its cell in the right-side block
- `~key` → **Build A** column (B) · `^key` → **Build B** column (C) — for the duel

So `"=@max_dmg*@crit_dmg_mult"` resolves per build column (B, C, E, G, I, K).
The `vs A` columns (D, F, H, J, L) are generated automatically as
`=(buildcell − Bcell)/Bcell`. Game constants are defined in the `CONSTS` list
and rendered as the right-side block; `#key` points at the matching cell.

`Max Damage` / `Max Health` are entered as text (so `1.2B` is accepted) and a
`parse_formula()` helper builds the formula that turns the `M` / `B` / `T`
suffix into a number — edit that helper to add more suffixes.

### When the game maths changes

1. Update the matching formula in `src/utils/statEngine.ts` (the web tool).
2. Update the matching `out(...)` line in `scripts/build_excel_calculator.py`.
3. Update the formula text in section 4 above **and** in the workbook's
   `Documentation` sheet (`d_text(...)` lines).
4. Re-run `python scripts/build_excel_calculator.py`.
5. Open the result in Excel, force a recalculation, and confirm the green
   result rows are sane.

Keeping all four in sync ensures the spreadsheet and the web calculator always
agree.
