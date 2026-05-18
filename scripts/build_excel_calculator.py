# -*- coding: utf-8 -*-
"""
Forge Master - Build Calculator & Comparison (Excel generator)

Generates ForgeMaster_Build_Calculator.xlsx in the repo root.

Layout of the Calculator sheet:
    - The main table starts at the top: label column A, then 6 builds (A-F).
    - GAME CONSTANTS sit in a separate block on the right (columns J-L).
    - The MATCHUP section runs a 1-on-1 duel for Build A vs Build B.
    - The final section duels every build B-F against Build A.

Per build you enter only:
    - Weapon Type (Melee / Range)
    - Max Damage and Max Health (accept a plain number or an M / B / T suffix)
    - the 13 stat % totals
Everything else is derived. Combat formulas mirror src/utils/statEngine.ts.

Re-run after changing anything:  python scripts/build_excel_calculator.py
"""

import re
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------
# Builds  -  one column per build, A-F laid out contiguously
# --------------------------------------------------------------------------
BUILDS = [
    {"name": "BUILD A", "idx": 0, "col": 2, "letter": "B"},
    {"name": "BUILD B", "idx": 1, "col": 3, "letter": "C"},
    {"name": "BUILD C", "idx": 2, "col": 4, "letter": "D"},
    {"name": "BUILD D", "idx": 3, "col": 5, "letter": "E"},
    {"name": "BUILD E", "idx": 4, "col": 6, "letter": "F"},
    {"name": "BUILD F", "idx": 5, "col": 7, "letter": "G"},
]
NUM_BUILDS = len(BUILDS)
NOTE_COL   = 8                        # column H
LAST_COL   = NOTE_COL
LAST_LET   = get_column_letter(LAST_COL)

CONST_LABEL_COL = 10                  # column J
CONST_VAL_COL   = 11                  # column K
CONST_NOTE_COL  = 12                  # column L

# --------------------------------------------------------------------------
# Colours / styles
# --------------------------------------------------------------------------
C_TITLE   = "1F3864"
C_SECTION = "2F5597"
C_HEADER  = "8EA9DB"
C_INPUT   = "FFF2CC"
C_OUTPUT  = "F2F2F2"
C_HILITE  = "C6E0B4"
C_CONST   = "DDEBF7"

thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

FILL_TITLE   = PatternFill("solid", fgColor=C_TITLE)
FILL_SECTION = PatternFill("solid", fgColor=C_SECTION)
FILL_HEADER  = PatternFill("solid", fgColor=C_HEADER)
FILL_INPUT   = PatternFill("solid", fgColor=C_INPUT)
FILL_OUTPUT  = PatternFill("solid", fgColor=C_OUTPUT)
FILL_HILITE  = PatternFill("solid", fgColor=C_HILITE)
FILL_CONST   = PatternFill("solid", fgColor=C_CONST)

NF = {
    "big":  "#,##0",
    "mult": "0.0000",
    "sec":  '0.000" s"',
    "pct":  "0.00%",
    "aps":  "0.000",
    "int":  "0",
    "dec2": "0.00",
    "ttk":  '0.0" s"',
}

WHITE = Font(color="FFFFFF", bold=True)
CTR   = Alignment(horizontal="center")

# --------------------------------------------------------------------------
# Game constants  (rendered as a block on the right)
# --------------------------------------------------------------------------
CONSTS = [
    ("Base Damage",                "base_dmg",    10,   "Player base attack damage", "int"),
    ("Base Health",                "base_hp",     80,   "Player base health", "int"),
    ("Base Crit Damage",           "base_crit",   0.20, "+20% -> crit mult starts at 1.20x", "dec2"),
    ("Power: Damage Weight",       "pow_dmg",     8,    "Damage weight in the Power formula", "int"),
    ("Power: Final Multiplier",    "pow_fin",     3,    "Outer multiplier in the Power formula", "int"),
    ("Attack Breakpoint Step (s)", "step",        0.10, "Animation times floored to this step", "dec2"),
    ("Attack Recovery Pad (s)",    "pad",         0.20, "Fixed time added to every attack cycle", "dec2"),
    ("Double-Hit Base Delay (s)",  "dhd",         0.25, "Delay before the bonus double hit", "dec2"),
    ("Melee Attack Duration (s)",  "melee_dur",   1.50, "Default melee attack animation length", "dec2"),
    ("Melee Windup (s)",           "melee_wind",  0.50, "Default melee time-to-hit", "dec2"),
    ("Ranged Attack Duration (s)", "ranged_dur",  1.50, "Default ranged attack animation length", "dec2"),
    ("Ranged Windup (s)",          "ranged_wind", 0.50, "Default ranged time-to-hit", "dec2"),
]
# constant key -> its absolute cell ($P$row), rows 2.. in the right block
CONST_CELL = {}
for _i, (_lab, _key, _v, _n, _f) in enumerate(CONSTS):
    CONST_CELL[_key] = f"${get_column_letter(CONST_VAL_COL)}${2 + _i}"

# --------------------------------------------------------------------------
# Calculator sheet definition
#
# Formula tokens:  @key = the build's own column ;  #key = a game constant ;
#                  ~key = Build A column (B) ;  ^key = Build B column (C)
# --------------------------------------------------------------------------
E = []

def title(s):    E.append({"t": "title", "label": s})
def subtitle(s): E.append({"t": "subtitle", "label": s})
def blank():     E.append({"t": "blank"})
def section(s):  E.append({"t": "section", "label": s})
def header():    E.append({"t": "header"})
def inp(label, key, vals, note, fmt="dec2"):
    assert len(vals) == NUM_BUILDS, f"{key}: need {NUM_BUILDS} values"
    E.append({"t": "input", "label": label, "key": key, "vals": list(vals),
              "note": note, "fmt": fmt})
def out(label, key, fmt, formula, hilite=False, skip_a=False):
    # skip_a=True leaves Build A's own column blank (a build cannot duel itself)
    E.append({"t": "output", "label": label, "key": key, "fmt": fmt,
              "formula": formula, "hilite": hilite, "skip_a": skip_a})
def matchup(label, key, fmt, fa, fb, hilite=False):
    E.append({"t": "matchup", "label": label, "key": key, "fmt": fmt,
              "fa": fa, "fb": fb, "hilite": hilite})
def verdict(label, key, formula):
    E.append({"t": "verdict", "label": label, "key": key, "formula": formula})

def parse_formula(inkey):
    """Excel formula: read an input that is a plain number or a string with an
    M / B / T suffix (e.g. 1.2B) and return its numeric value."""
    v = f"@{inkey}"
    t = f"TRIM(@{inkey})"
    base = f"VALUE(LEFT({t},LEN({t})-1))"
    return (f'=IF(ISNUMBER({v}),{v},'
            f'IF(RIGHT(UPPER({t}),1)="M",{base}*1000000,'
            f'IF(RIGHT(UPPER({t}),1)="B",{base}*1000000000,'
            f'IF(RIGHT(UPPER({t}),1)="T",{base}*1000000000000,'
            f'VALUE({v})))))')

def six(*vals):
    """Expand to NUM_BUILDS values; a single value fills every build."""
    if len(vals) == 1:
        return [vals[0]] * NUM_BUILDS
    return list(vals)

# ---- title -------------------------------------------------------------
title("FORGE MASTER  -  BUILD CALCULATOR & COMPARISON")
subtitle("6 builds (A-F) side by side. Game constants are in the block on the right. "
         "MATCHUP duels A vs B; the final section duels every build B-F against Build A.")
blank()
header()

# ---- build setup -------------------------------------------------------
section("BUILD SETUP  -  the core numbers for each build")
inp("Weapon Type", "wpn_type", six("Range"), "Melee or Range (pick from the dropdown)", "text")
inp("Max Damage",  "max_dmg_in",
    six("2.4B", "2.0B", "1.5B", "1.1B", "1.5B", "1.25B"),
    "Final Damage from the character screen - accepts M / B / T (e.g. 1.2B) or a plain number", "text")
inp("Max Health",  "max_hp_in",
    six("500M", "600M", "1.2B", "600M", "820M", "660M"),
    "Final Health from the character screen - accepts M / B / T (e.g. 640M) or a plain number", "text")
blank()

# ---- substat totals ----------------------------------------------------
section("SUBSTAT TOTALS  -  % summed from all 24 rolls (12 items x 2 substats)")
# Builds A-C are the 3 worked example builds (Crit Burst / Attack-Speed DPS /
# Lifesteal Bruiser) - each spends exactly 24 substat rolls. D-F are spare slots.
inp("Damage %",        "s_dmg",        six(30.00, 30.00, 0.00, 8.00, 13.20, 9.50),    "Max per roll: 15")
inp("Melee Damage %",  "s_melee",      six(0.00, 0.00, 0.00, 0.00, 0.00, 0.00),       "Max per roll: 50  (used only if weapon is melee)")
inp("Range Damage %",  "s_range",      six(30.00, 30.00, 0.00, 12.90, 14.50, 13.60),  "Max per roll: 15  (used only if weapon is ranged)")
inp("Crit Damage %",   "s_critdmg",    six(640.00, 320.00, 320.00, 360.00, 410.00, 388.00), "Max per roll: 80")
inp("Crit Chance %",   "s_critchance", six(48.00, 36.00, 36.00, 29.00, 36.00, 32.00), "Max per roll: 12")
inp("Double Chance %", "s_double",     six(80.00, 100.00, 40.00, 42.00, 52.00, 46.00), "Max per roll: 20")
inp("Attack Speed %",  "s_atkspeed",   six(160.00, 320.00, 120.00, 115.00, 140.00, 125.00), "Max per roll: 40")
inp("Lifesteal %",     "s_lifesteal",  six(0.00, 0.00, 120.00, 64.00, 52.00, 60.00),  "Max per roll: 20")
inp("Health Regen %",  "s_healthreg",  six(0.00, 0.00, 4.00, 0.00, 3.00, 0.50),       "Max per roll: 4")
inp("Skill Cooldown %","s_skillcd",    six(0.00, 0.00, 0.00, 2.90, 5.00, 3.50),       "Max per roll: 7   (cooldown reduction)")
inp("Skill Damage %",  "s_skilldmg",   six(0.00, 0.00, 0.00, 0.00, 22.00, 8.00),      "Max per roll: 30")
inp("Health %",        "s_health",     six(0.00, 0.00, 60.00, 0.00, 10.00, 3.00),     "Max per roll: 15")
inp("Block Chance %",  "s_block",      six(0.00, 0.00, 5.00, 0.00, 4.00, 2.00),       "Max per roll: 5")
blank()

# ---- derived: max damage & health numbers -----------------------------
section("DERIVED  -  MAX DAMAGE & HEALTH (numbers)")
out("Max Damage (number)", "max_dmg", "big", parse_formula("max_dmg_in"))
out("Max Health (number)", "max_hp", "big", parse_formula("max_hp_in"))
blank()

# ---- derived: power ----------------------------------------------------
section("DERIVED  -  POWER")
out("POWER", "power", "big",
    "=ROUND(((@max_dmg-#base_dmg)*#pow_dmg+(@max_hp-#base_hp))*#pow_fin,0)", hilite=True)
blank()

# ---- derived: damage & crit -------------------------------------------
section("DERIVED  -  DAMAGE & CRIT")
out("Crit Damage Multiplier x",       "crit_dmg_mult", "mult", "=1+#base_crit+@s_critdmg/100")
out("Critical Hit Damage",            "crit_hit",      "big",  "=@max_dmg*@crit_dmg_mult", hilite=True)
out("Crit Chance (cap 100%)",         "cchance",       "pct",  "=MIN(@s_critchance/100,1)")
out("Crit Multiplier on average hit", "crit_mult_hit", "mult", "=1+@cchance*(@crit_dmg_mult-1)")
out("Double Damage Chance (cap 100%)","dchance",       "pct",  "=MIN(@s_double/100,1)")
out("Double Multiplier on hit",       "double_mult",   "mult", "=1+@dchance")
blank()

# ---- derived: attack speed --------------------------------------------
section("DERIVED  -  ATTACK SPEED & BREAKPOINTS")
out("Attack Speed Multiplier x",      "atk_speed_mult", "mult", "=1+@s_atkspeed/100")
out("Weapon Attack Duration (s)",     "wpn_dur",        "sec",  '=IF(@wpn_type="Range",#ranged_dur,#melee_dur)')
out("Weapon Windup (s)",              "wpn_wind",       "sec",  '=IF(@wpn_type="Range",#ranged_wind,#melee_wind)')
out("Theoretical APS (attacks/sec)",  "theo_aps",       "aps",  "=@atk_speed_mult/@wpn_dur")
out("Base Recovery (s)",              "base_recovery",  "sec",  "=MAX(0,@wpn_dur-@wpn_wind)")
out("Stepped Windup (s)",             "stepped_windup", "sec",  "=FLOOR(@wpn_wind/@atk_speed_mult,#step)")
out("Stepped Recovery (s)",           "stepped_recovery","sec", "=FLOOR(@base_recovery/@atk_speed_mult,#step)")
out("Stepped Attack Cycle (s)",       "stepped_cycle",  "sec",  "=MAX(0.4,@stepped_windup+@stepped_recovery+#pad)")
out("Stepped Double-Hit Delay (s)",   "stepped_dhd",    "sec",  "=MAX(0.1,FLOOR(#dhd/@atk_speed_mult,#step))")
out("Double-Hit Cycle (s)",           "dh_cycle",       "sec",  "=@stepped_cycle+@stepped_dhd")
out("Average Real Cycle (s)",         "avg_real_cycle", "sec",  "=(1-@dchance)*@stepped_cycle+@dchance*@dh_cycle")
out("Real APS (weighted)",            "real_aps",       "aps",  "=(1+@dchance)/@avg_real_cycle", hilite=True)
blank()

# ---- derived: skill cooldown ------------------------------------------
section("DERIVED  -  SKILL COOLDOWN")
out("Skill Cooldown Reduction",       "cd_reduction",   "pct",  "=@s_skillcd/100")
out("Cooldown Multiplier",            "cd_mult",        "mult", "=MAX(0.1,1-@cd_reduction)")
blank()

# ---- derived: dps & hps ------------------------------------------------
section("DERIVED  -  DPS & HPS")
out("Theoretical Weapon DPS",         "theo_wpn_dps",   "big",  "=@max_dmg*@theo_aps*@crit_mult_hit*@double_mult")
out("Real Weapon DPS",                "real_wpn_dps",   "big",  "=@max_dmg*@real_aps*@crit_mult_hit", hilite=True)
out("Health Regen per second",        "hp_regen_s",     "big",  "=@max_hp*@s_healthreg/100")
out("Block Chance (cap 95%)",         "block_cap",      "pct",  "=MIN(@s_block/100,0.95)")
out("Block Factor",                   "block_factor",   "mult", "=1/(1-@block_cap)")
out("Lifesteal HPS",                  "lifesteal_hps",  "big",  "=@real_wpn_dps*@s_lifesteal/100")
out("TOTAL HPS (real)",               "total_hps",      "big",  "=(@hp_regen_s+@lifesteal_hps)*@block_factor", hilite=True)
blank()

# ---- matchup: A vs B ---------------------------------------------------
section("MATCHUP  -  WHO WINS, BUILD A vs BUILD B  (1-on-1 duel; C-F are comparison only)")
matchup("Net DPS dealt to opponent", "net_dps", "big",
        "=~real_wpn_dps-^total_hps",
        "=^real_wpn_dps-~total_hps")
matchup("Time to defeat opponent (s)", "ttk", "ttk",
        '=IF(@net_dps<=0,"never",^max_hp/@net_dps)',
        '=IF(@net_dps<=0,"never",~max_hp/@net_dps)')
verdict("WINNER  (A vs B)", "winner",
        '=IF(AND(~ttk="never",^ttk="never"),"DRAW - neither can kill",'
        'IF(~ttk="never","BUILD B WINS",IF(^ttk="never","BUILD A WINS",'
        'IF(~ttk<^ttk,"BUILD A WINS",IF(~ttk>^ttk,"BUILD B WINS","DRAW")))))')
blank()

# ---- does this build beat Build A? ------------------------------------
section("DOES THIS BUILD BEAT BUILD A?  -  1-on-1 duel of each build B-F vs Build A")
out("Net DPS dealt to Build A",           "net_vs_a",   "big",
    "=@real_wpn_dps-~total_hps", skip_a=True)
out("Net DPS taken from Build A",         "net_from_a", "big",
    "=~real_wpn_dps-@total_hps", skip_a=True)
out("Time to defeat Build A (s)",         "ttk_vs_a",   "ttk",
    '=IF(@net_vs_a<=0,"never",~max_hp/@net_vs_a)', skip_a=True)
out("Time for Build A to defeat this (s)", "ttk_from_a", "ttk",
    '=IF(@net_from_a<=0,"never",@max_hp/@net_from_a)', skip_a=True)
out("VERDICT  vs BUILD A", "verdict_vs_a", "text",
    '=IF(AND(@ttk_vs_a="never",@ttk_from_a="never"),"DRAW - neither can kill",'
    'IF(@ttk_vs_a="never","LOSES vs A",IF(@ttk_from_a="never","BEATS A",'
    'IF(@ttk_vs_a<@ttk_from_a,"BEATS A",'
    'IF(@ttk_vs_a>@ttk_from_a,"LOSES vs A","DRAW")))))',
    hilite=True, skip_a=True)

# --------------------------------------------------------------------------
# Pass 1 - assign row numbers
# --------------------------------------------------------------------------
R = {}
row = 1
header_row = None
for e in E:
    e["row"] = row
    if e["t"] == "header":
        header_row = row
    if "key" in e:
        R[e["key"]] = row
    row += 1
last_row = row - 1

def resolve(formula, col):
    """Tokens:  @key = column `col` ;  #key = a game constant cell ;
                ~key = Build A column (B) ;  ^key = Build B column (C)."""
    def repl(m):
        sym, key = m.group(1), m.group(2)
        if sym == "#":
            return CONST_CELL[key]
        r = R[key]
        if sym == "~": return f"B{r}"
        if sym == "^": return f"C{r}"
        return f"{col}{r}"
    return re.sub(r"([@#~^])([a-z_0-9]+)", repl, formula)

# --------------------------------------------------------------------------
# Build the workbook
# --------------------------------------------------------------------------
wb = Workbook()
ws = wb.active
ws.title = "Calculator"

ws.column_dimensions["A"].width = 30
for b in BUILDS:
    ws.column_dimensions[b["letter"]].width = 15
ws.column_dimensions[LAST_LET].width = 44                                   # H notes
ws.column_dimensions[get_column_letter(LAST_COL + 1)].width = 3             # I spacer
ws.column_dimensions[get_column_letter(CONST_LABEL_COL)].width = 24         # J const label
ws.column_dimensions[get_column_letter(CONST_VAL_COL)].width = 11           # K const value
ws.column_dimensions[get_column_letter(CONST_NOTE_COL)].width = 44          # L const note

def mspan(rw):
    return f"A{rw}:{LAST_LET}{rw}"

for e in E:
    t, rw = e["t"], e["row"]

    if t == "title":
        ws.merge_cells(mspan(rw))
        c = ws.cell(rw, 1, e["label"])
        c.font = Font(color="FFFFFF", bold=True, size=14)
        c.fill = FILL_TITLE
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[rw].height = 26

    elif t == "subtitle":
        ws.merge_cells(mspan(rw))
        c = ws.cell(rw, 1, e["label"])
        c.font = Font(italic=True, size=9, color="404040")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[rw].height = 28

    elif t == "blank":
        pass

    elif t == "section":
        ws.merge_cells(mspan(rw))
        c = ws.cell(rw, 1, e["label"])
        c.font = WHITE
        c.fill = FILL_SECTION
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[rw].height = 20

    elif t == "header":
        h = ws.cell(rw, 1, "STAT / METRIC")
        h.font = WHITE; h.fill = FILL_HEADER; h.border = BORDER
        h.alignment = Alignment(horizontal="center", vertical="center")
        for b in BUILDS:
            c = ws.cell(rw, b["col"], b["name"])
            c.font = WHITE; c.fill = FILL_HEADER; c.border = BORDER
            c.alignment = Alignment(horizontal="center", vertical="center")
        n = ws.cell(rw, NOTE_COL, "MAX PER ROLL / NOTES")
        n.font = WHITE; n.fill = FILL_HEADER; n.border = BORDER
        n.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[rw].height = 20

    elif t == "input":
        ws.cell(rw, 1, e["label"])
        is_text = e["fmt"] == "text"
        for b in BUILDS:
            c = ws.cell(rw, b["col"], e["vals"][b["idx"]])
            c.fill = FILL_INPUT; c.border = BORDER
            if not is_text:
                c.number_format = NF[e["fmt"]]
            c.alignment = CTR
        n = ws.cell(rw, NOTE_COL, e["note"]); n.font = Font(italic=True, size=9, color="404040")

    elif t == "output":
        a = ws.cell(rw, 1, e["label"])
        fill = FILL_HILITE if e["hilite"] else FILL_OUTPUT
        a.font = Font(bold=e["hilite"])
        is_text = e["fmt"] == "text"
        for b in BUILDS:
            if e.get("skip_a") and b["idx"] == 0:
                c = ws.cell(rw, b["col"], "-")
                c.fill = fill; c.border = BORDER; c.alignment = CTR
                c.font = Font(bold=e["hilite"])
                continue
            c = ws.cell(rw, b["col"], resolve(e["formula"], b["letter"]))
            c.fill = fill; c.border = BORDER
            if not is_text:
                c.number_format = NF[e["fmt"]]
            c.alignment = CTR
            c.font = Font(bold=e["hilite"])

    elif t == "matchup":
        a = ws.cell(rw, 1, e["label"])
        fill = FILL_HILITE if e["hilite"] else FILL_OUTPUT
        a.font = Font(bold=e["hilite"])
        is_text = e["fmt"] == "text"
        for col, letter, frm in ((2, "B", e["fa"]), (3, "C", e["fb"])):
            c = ws.cell(rw, col, resolve(frm, letter))
            c.fill = fill; c.border = BORDER
            if not is_text:
                c.number_format = NF[e["fmt"]]
            c.alignment = CTR
            c.font = Font(bold=e["hilite"])

    elif t == "verdict":
        a = ws.cell(rw, 1, e["label"])
        a.font = Font(bold=True, size=12)
        ws.merge_cells(f"B{rw}:C{rw}")
        c = ws.cell(rw, 2, resolve(e["formula"], "B"))
        c.fill = FILL_HILITE
        c.font = Font(bold=True, size=12, color="1F3864")
        c.alignment = Alignment(horizontal="center", vertical="center")
        for col in (2, 3):
            ws.cell(rw, col).border = BORDER
        ws.row_dimensions[rw].height = 24

# ---- game constants block (right side, columns J-L) --------------------
ct = ws.cell(1, CONST_LABEL_COL, "GAME CONSTANTS")
ws.merge_cells(start_row=1, start_column=CONST_LABEL_COL,
               end_row=1, end_column=CONST_NOTE_COL)
ct.font = WHITE; ct.fill = FILL_TITLE
ct.alignment = Alignment(horizontal="left", vertical="center", indent=1)
for i, (label, key, val, note, fmt) in enumerate(CONSTS):
    rr = 2 + i
    lc = ws.cell(rr, CONST_LABEL_COL, label); lc.font = Font(bold=True, size=9)
    lc.border = BORDER
    vc = ws.cell(rr, CONST_VAL_COL, val)
    vc.fill = FILL_CONST; vc.border = BORDER
    vc.number_format = NF[fmt]; vc.alignment = CTR
    nc = ws.cell(rr, CONST_NOTE_COL, note)
    nc.font = Font(italic=True, size=9, color="404040")

# ---- weapon-type dropdown ---------------------------------------------
dv = DataValidation(type="list", formula1='"Melee,Range"', allow_blank=False)
ws.add_data_validation(dv)
for b in BUILDS:
    dv.add(ws[f"{b['letter']}{R['wpn_type']}"])

ws.freeze_panes = f"B{header_row+1}"
ws.sheet_view.showGridLines = False

# ==========================================================================
# Documentation sheet
# ==========================================================================
doc = wb.create_sheet("Documentation")
doc.column_dimensions["A"].width = 118
doc.sheet_view.showGridLines = False

def d_title(s):
    r = doc.max_row + 1 if doc["A1"].value else 1
    c = doc.cell(r, 1, s)
    c.font = Font(color="FFFFFF", bold=True, size=13)
    c.fill = FILL_TITLE
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    doc.row_dimensions[r].height = 24

def d_head(s):
    r = doc.max_row + 1
    c = doc.cell(r, 1, s)
    c.font = WHITE
    c.fill = FILL_SECTION
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    doc.row_dimensions[r].height = 19

def d_text(s, mono=False):
    r = doc.max_row + 1
    c = doc.cell(r, 1, s)
    f = {"size": 10}
    if mono:
        f["name"] = "Consolas"; f["size"] = 9
    c.font = Font(**f)
    c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1)
    lines = max(1, len(s) // 110 + s.count("\n") + 1)
    doc.row_dimensions[r].height = 15 * lines

d_title("FORGE MASTER BUILD CALCULATOR  -  DOCUMENTATION")
d_text("")
d_text("This workbook compares 6 builds (A-F). The main table starts at the top of the "
       "Calculator sheet; GAME CONSTANTS sit in a block on the right (columns J-L). Combat "
       "formulas mirror src/utils/statEngine.ts (function finalizeCalculation).")
d_text("")

d_head("1.  SHEET LAYOUT")
d_text("- Column A          : stat / metric label.")
d_text("- Columns B-G       : the 6 builds, one column each. Build A = B, Build B = C, "
       "Build C = D, Build D = E, Build E = F, Build F = G.")
d_text("- Column H          : max-per-roll cap / notes.")
d_text("- Columns J-L       : the GAME CONSTANTS block (shared values, rarely changed).")
d_text("")

d_head("2.  WHAT YOU ENTER  (yellow cells)")
d_text("Per build: Weapon Type (Melee/Range dropdown); Max Damage and Max Health (the final "
       "character-screen numbers - they accept a plain number or an M / B / T suffix, e.g. "
       "1.2B = 1,200,000,000); and the 13 substat % totals summed across all 24 rolls "
       "(12 items x 2). Type 403.8 for 403.8%, not 4.038.")
d_text("Max Damage / Max Health are FINAL totals - all gear, pet, mount, skill and ascension "
       "bonuses are already inside them.")
d_text("")

d_head("3.  POWER")
d_text("POWER = ( (MaxDamage - 10) x 8 + (MaxHealth - 80) ) x 3", mono=True)
d_text("Reverse-engineered from the game binary.")
d_text("")

d_head("4.  DAMAGE & CRIT")
d_text("Crit Damage Mult   = 1 + 0.20 + CritDamage%", mono=True)
d_text("Critical Hit Damage= Max Damage x Crit Damage Mult", mono=True)
d_text("Crit Mult on hit   = 1 + CritChance x (Crit Damage Mult - 1)", mono=True)
d_text("Double Mult on hit = 1 + DoubleChance%", mono=True)
d_text("")

d_head("5.  ATTACK SPEED & BREAKPOINTS")
d_text("Attack Speed Mult  = 1 + AttackSpeed%                    (122% -> 2.22x)", mono=True)
d_text("Weapon Duration/Windup come from the Melee or Ranged constants by Weapon Type.", mono=True)
d_text("Theoretical APS    = Attack Speed Mult / Weapon Duration", mono=True)
d_text("Stepped Windup     = FLOOR( Windup / SpeedMult , 0.1 )", mono=True)
d_text("Stepped Recovery   = FLOOR( (Duration - Windup) / SpeedMult , 0.1 )", mono=True)
d_text("Stepped Cycle      = MAX( 0.4 , Stepped Windup + Stepped Recovery + 0.2 )", mono=True)
d_text("Double-Hit Cycle   = Stepped Cycle + MAX(0.1, FLOOR(0.25 / SpeedMult , 0.1))", mono=True)
d_text("Avg Real Cycle     = (1 - Double%) x Cycle + Double% x DoubleHitCycle", mono=True)
d_text("Real APS           = (1 + Double%) / Avg Real Cycle", mono=True)
d_text("The game floors animation times to 0.1s steps, so attack speed improves in jumps.")
d_text("")

d_head("6.  SKILL COOLDOWN")
d_text("Cooldown Reduction = Skill Cooldown%", mono=True)
d_text("Cooldown Mult      = MAX( 0.1 , 1 - Reduction )", mono=True)
d_text("")

d_head("7.  DPS & HPS")
d_text("Theoretical Wpn DPS= Max Damage x Theoretical APS x Crit Mult x Double Mult", mono=True)
d_text("Real Weapon DPS    = Max Damage x Real APS x Crit Mult   (Double already in Real APS)", mono=True)
d_text("Block Factor       = 1 / (1 - MIN(Block%, 95%))", mono=True)
d_text("Lifesteal HPS      = Real Weapon DPS x Lifesteal%", mono=True)
d_text("TOTAL HPS (real)   = (Health Regen/s + Lifesteal HPS) x Block Factor", mono=True)
d_text("")

d_head("8.  WHO WINS  -  THE MATCHUP  (Build A vs Build B)")
d_text("Net DPS dealt   = your Real Weapon DPS - the opponent's Total HPS", mono=True)
d_text("Time to defeat  = opponent's Max Health / your Net DPS dealt", mono=True)
d_text("WINNER          = the build with the shorter Time to defeat", mono=True)
d_text("If your Net DPS is 0 or below, Time to defeat shows 'never'. Both unkillable, or an "
       "exact tie, gives a DRAW. The duel is Build A vs Build B only; C-F are comparison "
       "columns. To duel a different pair, copy a build's values into the A or B column.")
d_text("")

d_head("9.  DOES THIS BUILD BEAT BUILD A?")
d_text("The final section duels every build B-F against Build A one-on-one, in that build's "
       "own column:", mono=True)
d_text("Net DPS dealt to A   = this build's Real Weapon DPS - Build A's Total HPS", mono=True)
d_text("Net DPS taken from A = Build A's Real Weapon DPS - this build's Total HPS", mono=True)
d_text("Time to defeat A     = Build A's Max Health / Net DPS dealt to A", mono=True)
d_text("VERDICT vs Build A   = BEATS A / LOSES vs A / DRAW (shorter time-to-kill wins)", mono=True)
d_text("Build A's own column shows a dash - a build cannot duel itself. 'never' appears when "
       "a Net DPS is 0 or below; two unkillable builds give a DRAW.")
d_text("")

d_head("10.  MAINTAINING THIS FILE")
d_text("Do not edit cells by hand for permanent changes. Edit scripts/build_excel_calculator.py "
       "and re-run:  python scripts/build_excel_calculator.py")
d_text("Tokens in formulas: @key = the build's own column, #key = a game constant, "
       "~key = Build A, ^key = Build B. Constants live in the CONSTS list (rendered as the "
       "right-side block); inputs use inp(...) with one value per build via six(...).")

# ==========================================================================
# Stat Reference sheet
# ==========================================================================
ref = wb.create_sheet("Stat Reference")
ref.sheet_view.showGridLines = False
ref.column_dimensions["A"].width = 22
ref.column_dimensions["B"].width = 16
ref.column_dimensions["C"].width = 64

rc = ref.cell(1, 1, "STAT REFERENCE  -  max value per single substat roll")
ref.merge_cells("A1:C1")
rc.font = Font(color="FFFFFF", bold=True, size=12)
rc.fill = FILL_TITLE
rc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
ref.row_dimensions[1].height = 24

for i, lab in enumerate(["STAT", "MAX / ROLL", "NOTES"], start=1):
    c = ref.cell(2, i, lab)
    c.font = WHITE; c.fill = FILL_HEADER; c.border = BORDER
    c.alignment = CTR

stat_rows = [
    ("Crit Damage",   80, "Increases damage dealt by critical hits"),
    ("Crit Chance",   12, "Probability that an attack is a critical hit"),
    ("Health Regen",  4,  "Health restored per second, as % of max health"),
    ("Lifesteal",     20, "% of weapon damage dealt converted to healing"),
    ("Double Chance", 20, "Probability of a bonus double hit"),
    ("Damage",        15, "Global damage multiplier"),
    ("Melee Damage",  50, "Extra damage, melee weapons only"),
    ("Range Damage",  15, "Extra damage, ranged weapons only"),
    ("Attack Speed",  40, "Speeds up the attack animation cycle"),
    ("Skill Cooldown",7,  "Reduces active-skill cooldown"),
    ("Skill Damage",  30, "Multiplier on active-skill damage and healing"),
    ("Health",        15, "Global health multiplier"),
    ("Block Chance",  5,  "Probability to block; raises effective HPS"),
]
r = 3
for name, mx, note in stat_rows:
    ref.cell(r, 1, name).border = BORDER
    cc = ref.cell(r, 2, mx); cc.border = BORDER; cc.alignment = CTR
    nn = ref.cell(r, 3, note); nn.border = BORDER
    nn.alignment = Alignment(horizontal="left", indent=1)
    r += 1

r += 1
for line in [
    "Build setup notes:",
    "- 12 items total = 8 equipment items + 1 mount + 3 pets, each with exactly 2 substats.",
    "- That is 24 substat rolls; the SUBSTAT TOTALS on the Calculator are their summed %.",
    "- Gear, pets, mount, skills and ascension are all already reflected in the",
    "  Max Damage / Max Health you read off the character screen.",
]:
    c = ref.cell(r, 1, line)
    ref.merge_cells(f"A{r}:C{r}")
    c.font = Font(size=10, bold=line.endswith(":"))
    c.alignment = Alignment(horizontal="left", indent=1)
    r += 1

# ==========================================================================
out_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                        "ForgeMaster_Build_Calculator.xlsx")
wb.save(out_path)
print("Saved:", out_path)
print("Calculator rows:", last_row, " header row:", header_row, " builds:", NUM_BUILDS)
