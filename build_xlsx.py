"""Builds Hero-Ledger.xlsx from data.py. All character numbers are live Excel formulas."""
import json, os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
import data as D

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "Hero-Ledger.xlsx")
FONT = "Arial"
NAVY, BLUE, PALE, GRID, RED, GOLD = "16324F", "2F6DB5", "E8F0F8", "B9CBDD", "B3261E", "FFF6D6"
f = lambda **k: Font(name=FONT, size=k.pop("size", 10), **k)
fill = lambda c: PatternFill("solid", fgColor=c)
thin = Side(style="thin", color=GRID)
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
WRAP = Alignment(vertical="top", wrap_text=True)

wb = Workbook()
wsG = wb.active; wsG.title = "Guide"
wsB = wb.create_sheet("Builder"); wsS = wb.create_sheet("Sheet"); wsP = wb.create_sheet("Spells"); wsI = wb.create_sheet("Inventory")
wsC = wb.create_sheet("Calc")
for ws, c in ((wsG, "6B7C8F"), (wsB, RED), (wsS, BLUE), (wsP, BLUE), (wsI, BLUE), (wsC, "6B7C8F")):
    ws.sheet_properties.tabColor = c
    ws.sheet_view.showGridLines = False

def name(n, ref):
    wb.defined_names[n] = DefinedName(n, attr_text=ref)

# ---------- data tables
class T:
    def __init__(self, title, rows, fields, key=None):
        self.ws = wb.create_sheet(title); self.ws.sheet_properties.tabColor = "B9CBDD"
        self.title, self.fields, self.n = title, fields, len(rows)
        self.col = {}
        for j, fld in enumerate(fields, 1):
            self.col[fld] = CL(j)
            c = self.ws.cell(1, j, fld); c.font = f(bold=True, color="FFFFFF"); c.fill = fill(NAVY)
            self.ws.column_dimensions[CL(j)].width = max(10, min(26, len(str(fld)) + 2))
        for i, r in enumerate(rows, 2):
            for j, v in enumerate(r, 1):
                c = self.ws.cell(i, j, v if v != "" else None); c.font = f()
        self.ws.freeze_panes = "A2"
        self.key = key or fields[0]
    def r(self, fld):
        c = self.col[fld]
        return f"{self.title}!${c}$2:${c}${self.n + 1}"
    def lk(self, fld, key, text=False):
        s = f"INDEX({self.r(fld)},MATCH({key},{self.r(self.key)},0))"
        return s + '&""' if text else s
    def mx(self, a, b):
        return f"{self.title}!${self.col[a]}$2:${self.col[b]}${self.n + 1}"

AB = D.AB
tCl = T("Classes", [[c["name"], c["die"], c["caster"], c["ability"], c["skillN"], "|".join(c["skills24"]), "|".join(c["skills14"]), c["req"][0], c["req"][1] if len(c["req"]) > 1 else "", c["reqMode"], c["armor24"], c["armor14"], "|".join(c["saves"])] for c in D.CLASSES],
        ["name", "die", "caster", "ability", "skillN", "skills24", "skills14", "req1", "req2", "reqMode", "armor24", "armor14", "saves"])
LVH = [f"L{i}" for i in range(1, 21)]
tSp = T("SpellTbl", [[f'{s["era"]}|{s["name"]}', s["mode"]] + s["cantrips"] + s["prepared"] for s in D.spellrows],
        ["key", "mode"] + ["C" + x for x in LVH] + ["P" + x for x in LVH])
tSl = T("Slots", [[i + 1] + r + [D.PACT_N[i], D.PACT_L[i], D.XP[i]] for i, r in enumerate(D.SLOTS)],
        ["level"] + [f"S{i}" for i in range(1, 10)] + ["pactN", "pactL", "xp"])
tSu = T("Subclasses", [[f'{s["era"]}|{s["cls"]}|{s["name"]}', s["level"], s["third"], s["hpPerLevel"], s["ac"]] for s in D.SUBCLASSES], ["key", "level", "third", "hpPerLevel", "ac"])
tSe = T("Species", [[f'{s["era"]}|{s["name"]}', s["size"], s["speed"], s["dv"], s["resist"], s["traits"]] + [s["fixed"][a] for a in AB] + s["picks"] + s["tasha"] + [s["hp"], s["carry"], s["skill"], s["extra"], s["feat"], s["noPick"]] for s in D.SPECIES],
        ["key", "size", "speed", "dv", "resist", "traits"] + AB + ["pick1", "pick2", "pick3", "tasha1", "tasha2", "tasha3", "hp", "carry", "skill", "extra", "feat", "noPick"])
tBg = T("Backgrounds", [[f'{b["era"]}|{b["name"]}', "|".join(b["abilities"]), b["feat"], b["skills"][0], b["skills"][1], b["tool"]] for b in D.BACKGROUNDS], ["key", "abilities", "feat", "sk1", "sk2", "tool"])
FX = ["hpPerLevel", "hpFlat", "speed", "initPB", "initFlat", "passive", "mam", "skills", "expert", "boon"]
tFe = T("Feats", [[f'{x["era"]}|{x["name"]}', x["cat"], x["points"], x["opts"]] + [x[k] for k in FX] for x in D.FEATS], ["key", "cat", "points", "opts"] + FX)
tAr = T("Armor", [[a["name"], a["kind"], a["base"], a["dexCap"], a["str"], a["stealth"], a["lb"], a["gp"]] for a in D.ARMOR], ["name", "kind", "base", "dexCap", "str", "stealth", "lb", "gp"])
WF = ["name", "cat", "range", "type", "props", "mastery", "finesse", "thrown", "twoH", "monk24", "monk14", "prof24", "prof14", "txt24", "vtxt24", "avg24", "vavg24", "gwf24", "vgwf24", "txt14", "vtxt14", "avg14", "vavg14", "gwf14", "vgwf14", "hasV24", "hasV14"]
for w in D.WEAPONS:
    w["hasV24"] = 1 if w["ver24"] else 0
    w["hasV14"] = 1 if w["ver14"] else 0
tWe = T("Weapons", [[w[k] for k in WF] for w in D.WEAPONS], WF)

# Lists: every dropdown is a contiguous block keyed in column A
lst = []
def add_list(k, vals):
    for v in vals: lst.append([k, v])
add_list("Rulebook", [r["id"] for r in D.RULESETS])
for r in D.RULESETS:
    era, rid = r["era"], r["id"]
    cn = [c["name"] for c in D.CLASSES if c["name"] != "Artificer" or r["artificer"]]
    add_list(f"Class|{rid}", cn)
    add_list(f"Species|{rid}", [s["name"] for s in D.SPECIES if s["era"] == era])
    add_list(f"Background|{rid}", [b["name"] for b in D.BACKGROUNDS if b["era"] == era])
    add_list(f"Feat|{rid}", [x["name"] for x in D.FEATS if x["era"] == era and (not x["tasha"] or r["tasha"])])
    for c in cn:
        add_list(f"Sub|{rid}|{c}", [s["name"] for s in D.SUBCLASSES if s["era"] == era and s["cls"] == c and (not s["tasha"] or r["tasha"])])
add_list("Armor", [a["name"] for a in D.ARMOR]); add_list("Weapon", [w["name"] for w in D.WEAPONS]); add_list("Style", D.FIGHTING_STYLES)
tLi = T("Lists", lst, ["key", "item"])
LN = len(lst) + 1
LA, LB = f"Lists!$A$1:$A${LN}", "Lists!$B$1"
def dyn(keyexpr):
    return f'OFFSET({LB},MATCH({keyexpr},{LA},0)-1,0,COUNTIF({LA},{keyexpr}),1)'

# Features and resources get running-count helper columns for the Sheet tab
CLS_N, CLS_L = "Calc!$A$4:$A$6", "Calc!$B$4:$B$6"
tFt = T("Features", [[x["era"], x["cls"], x["level"], x["name"], None, None, None] for x in D.FEATURES], ["era", "cls", "level", "name", "cur", "run", "label"])
for i in range(2, tFt.n + 2):
    w = tFt.ws
    w[f"E{i}"] = f"=IF(A{i}=ERA,SUMIF({CLS_N},B{i},{CLS_L}),0)"
    w[f"F{i}"] = f"=IF(AND(E{i}>=C{i},E{i}>0),1,0)+" + ("0" if i == 2 else f"F{i-1}")
    w[f"G{i}"] = f'=B{i}&" "&C{i}&": "&D{i}'
tRe = T("Resources", [[x["era"], x["cls"], x["name"], x["rest"], None, None, None, None] + x["v"] for x in D.RES], ["era", "cls", "name", "rest", "cur", "raw", "value", "run"] + LVH)
for i in range(2, tRe.n + 2):
    w = tRe.ws
    w[f"E{i}"] = f"=IF(A{i}=ERA,SUMIF({CLS_N},B{i},{CLS_L}),0)"
    w[f"F{i}"] = f'=IF(E{i}>0,INDEX($I{i}:$AB{i},1,E{i}),0)'
    w[f"G{i}"] = f'=IF(F{i}="CHA",MAX(1,mCHA),IF(F{i}="INT",MAX(1,mINT),IF(F{i}="WIS",MAX(1,mWIS),IF(F{i}="CHA+1",MAX(1,mCHA+1),F{i}))))'
    w[f"H{i}"] = f'=IF(AND(E{i}>0,G{i}<>0),1,0)+' + ("0" if i == 2 else f"H{i-1}")

# ---------- Builder helpers
def inp(ws, addr, v=None):
    c = ws[addr]
    if v is not None: c.value = v
    c.font = f(color="0000FF"); c.fill = fill(GOLD); c.border = BOX; c.alignment = Alignment(vertical="center", wrap_text=False)
    return c
def out(ws, addr, v, bold=False, fmt=None, center=False):
    c = ws[addr]; c.value = v; c.font = f(bold=bold); c.border = BOX
    if fmt: c.number_format = fmt
    if center: c.alignment = CENTER
    return c
def label(ws, addr, t, bold=False):
    c = ws[addr]; c.value = t; c.font = f(bold=bold); return c
def section(ws, row, t, a="B", b="M"):
    ca, cb = ws[f"{a}{row}"].column, ws[f"{b}{row}"].column
    for j in range(ca, cb + 1): ws.cell(row, j).fill = fill(NAVY)
    c = ws[f"{a}{row}"]; c.value = t; c.font = f(bold=True, color="FFFFFF", size=11)
def header(ws, row, a, labels):
    j0 = ws[f"{a}{row}"].column
    for j, t in enumerate(labels):
        c = ws.cell(row, j0 + j, t); c.font = f(bold=True); c.fill = fill(PALE); c.border = BOX; c.alignment = CENTER
def dv(ws, addr, formula, warn=True):
    v = DataValidation(type="list", formula1=formula, allow_blank=True)
    v.errorStyle = "warning" if warn else "stop"; v.showErrorMessage = True
    v.errorTitle = "Not on the list"; v.error = "That value is not in this rulebook's list."
    ws.add_data_validation(v); v.add(addr)
ABL = '"STR,DEX,CON,INT,WIS,CHA"'

B = wsB
B.column_dimensions["A"].width = 2
for c, w in zip("BCDEFGHIJKLM", [30, 26, 12, 12, 12, 12, 12, 10, 10, 10, 11, 10]): B.column_dimensions[c].width = w
B.column_dimensions["N"].width = 2; B.column_dimensions["O"].width = 15; B.column_dimensions["P"].width = 58
B["B1"] = "Hero Ledger: character builder"; B["B1"].font = f(bold=True, size=16, color=NAVY)
B["B2"] = "Fill the yellow cells with blue text. Everything else calculates. Sample character below: overwrite it."; B["B2"].font = f(italic=True, color="555555")

section(B, 4, "1  Rulebook and identity")
label(B, "B5", "Rulebook", True); inp(B, "C5", D.RULESETS[0]["id"]); dv(B, "C5", dyn('"Rulebook"'), warn=False)
B.merge_cells("C5:E5")
label(B, "B6", "Rules era"); out(B, "C6", "=" + "IF(LEFT(C5,4)=\"2024\",2024,2014)")
for i, (t, v) in enumerate([("Character name", "Mira Thorne (example)"), ("Player", ""), ("Alignment", "Lawful Good"), ("Experience points", 23000)]):
    label(B, f"B{7+i}", t); inp(B, f"C{7+i}", v or None)

section(B, 12, "2  Classes (first row is your starting class)")
header(B, 13, "B", ["Class", "Subclass", "Level", "Hit die", "Subclass?"])
SAMPLE_CL = [("Paladin", "Oath of Vengeance", 5), ("Warlock", None, 2), (None, None, None)]
for i in range(3):
    r = 14 + i
    inp(B, f"B{r}", SAMPLE_CL[i][0]); inp(B, f"C{r}", SAMPLE_CL[i][1]); inp(B, f"D{r}", SAMPLE_CL[i][2])
    dv(B, f"B{r}", dyn('"Class|"&RS')); dv(B, f"C{r}", dyn(f'"Sub|"&RS&"|"&$B{r}'))
    v = DataValidation(type="whole", operator="between", formula1="1", formula2="20", allow_blank=True); B.add_data_validation(v); v.add(f"D{r}")
    out(B, f"E{r}", f'=IF(Calc!C{4+i}>0,"d"&Calc!C{4+i},"")', center=True)
    out(B, f"F{r}", f'=IF(Calc!A{4+i}="","",IF(C{r}="","none",IF(Calc!F{4+i}=1,"active","not yet")))', center=True)
label(B, "B17", "Total level", True); out(B, "D17", "=SUM(Calc!B4:B6)", True, center=True)
label(B, "E17", "Prof. bonus", True); out(B, "F17", "=IF(D17>0,ROUNDUP(D17/4,0)+1,2)", True, center=True)

section(B, 19, "3  Ability scores")
label(B, "B20", "Generation method"); inp(B, "C20", "Standard Array"); dv(B, "C20", '"Standard Array,Point Buy,Rolled"', False)
header(B, 21, "B", ["Ability", "Base score", "Species", "Origin picks", "Feats / ASI", "Capstone", "Final", "Modifier", "Save", "Save prof", "Point cost"])
SAMPLE_BASE = dict(STR=15, DEX=10, CON=13, INT=8, WIS=12, CHA=14)
SPKEY = 'ERA&"|"&$C$35'
BGKEY = 'ERA&"|"&$C$36'
FR = range(43, 53)   # feat rows
for i, a in enumerate(AB):
    r = 22 + i
    label(B, f"B{r}", D.ABILITY_NAMES[a] + f" ({a})"); B[f"B{r}"].border = BOX
    inp(B, f"C{r}", SAMPLE_BASE[a]); B[f"C{r}"].alignment = CENTER
    out(B, f"D{r}", f'=IF(Calc!$B$20=1,0,IFERROR({tSe.lk(a, SPKEY)},0))', center=True)
    out(B, f"E{r}", f'=IF($C$31="{a}",$D$31,0)+IF($C$32="{a}",$D$32,0)+IF($C$33="{a}",$D$33,0)', center=True)
    out(B, f"F{r}", f'=SUMPRODUCT(($C$43:$C$52="{a}")*($E$43:$E$52>=1))+SUMPRODUCT(($D$43:$D$52="{a}")*($E$43:$E$52>=2))', center=True)
    boost = {"STR": "lvBarbarian>=20", "CON": "lvBarbarian>=20", "DEX": "AND(ERA=2024,lvMonk>=20)", "WIS": "AND(ERA=2024,lvMonk>=20)"}.get(a, "FALSE")
    out(B, f"G{r}", f"=H{r}-MIN(Calc!$B$21,C{r}+D{r}+E{r}+F{r})", center=True)
    out(B, f"H{r}", f"=IF({boost},MAX(MIN(Calc!$B$21,C{r}+D{r}+E{r}+F{r}),MIN(IF(ERA=2024,25,24),MIN(Calc!$B$21,C{r}+D{r}+E{r}+F{r})+4)),MIN(Calc!$B$21,C{r}+D{r}+E{r}+F{r}))", True, center=True)
    out(B, f"I{r}", f"=INT((H{r}-10)/2)", True, "+0;-0;0", True)
    prof = (f'OR(ISNUMBER(SEARCH("{a}",Calc!$B$22)),COUNTIFS($B$43:$B$52,"Resilient",$C$43:$C$52,"{a}")>0,lvMonk>=14,'
            + ('lvRogue>=15' if a == "WIS" else ('AND(ERA=2024,lvRogue>=15)' if a == "CHA" else 'FALSE')) + ')')
    out(B, f"K{r}", f'=IF({prof},"Yes","")', center=True)
    out(B, f"J{r}", f'=I{r}+IF(K{r}="Yes",PB,0)+Calc!$B$23+N($C$87)', True, "+0;-0;0", True)
    out(B, f"L{r}", f"=IF(AND(C{r}>=8,C{r}<=15),CHOOSE(C{r}-7,0,1,2,3,4,5,7,9),0)", center=True)
    name("s" + a, f"Builder!$H${r}"); name("m" + a, f"Builder!$I${r}")
label(B, "K28", "Total", True); out(B, "L28", "=SUM(L22:L27)", True, center=True)
label(B, "B30", "Background bonus split (2024)"); inp(B, "C30", "+2 / +1"); dv(B, "C30", '"+2 / +1,+1 / +1 / +1"', False)
amt = ['IF(ERA=2024,IF($C$30="+2 / +1",2,1),IF(TASHA=1,IFERROR({a},0),IFERROR({b},0)))',
       'IF(ERA=2024,1,IF(TASHA=1,IFERROR({a},0),IFERROR({b},0)))',
       'IF(ERA=2024,IF($C$30="+2 / +1",0,1),IF(TASHA=1,IFERROR({a},0),IFERROR({b},0)))']
for i, v in enumerate(["STR", "CHA", None]):
    r = 31 + i
    label(B, f"B{r}", f"Ability pick {i+1}"); inp(B, f"C{r}", v); dv(B, f"C{r}", ABL, False)
    out(B, f"D{r}", "=" + amt[i].format(a=tSe.lk(f"tasha{i+1}", SPKEY), b=tSe.lk(f"pick{i+1}", SPKEY)), center=True)
    label(B, f"E{r}", "points this pick gives")

section(B, 34, "4  Species and background")
label(B, "B35", "Species", True); inp(B, "C35", "Tiefling (Infernal)"); dv(B, "C35", dyn('"Species|"&RS'))
label(B, "B36", "Background", True); inp(B, "C36", "Noble"); dv(B, "C36", dyn('"Background|"&RS'))
for r, t, v in [(37, "Allowed bonus abilities", f'=IF(ERA=2024,IFERROR({tBg.lk("abilities", BGKEY)},""),"Species sets these")'),
                (38, "Origin feat (automatic)", f'=IF(ERA=2024,IFERROR({tBg.lk("feat", BGKEY, True)},""),"None in 2014 rules")'),
                (39, "Skills granted", f'=IFERROR({tBg.lk("sk1", BGKEY)}&", "&{tBg.lk("sk2", BGKEY)},"")&IF(Calc!B24<>"",", "&Calc!B24,"")'),
                (40, "Tools or languages", f'=IFERROR({tBg.lk("tool", BGKEY, True)},"")')]:
    label(B, f"B{r}", t); out(B, f"C{r}", v); B.merge_cells(f"C{r}:H{r}")

section(B, 41, "5  Feats and ability score improvements")
header(B, 42, "B", ["Feat or ASI", "+1 ability", "2nd +1 (ASI)", "Points", "Allowed", "Valid?"])
for r in FR:
    inp(B, f"B{r}"); inp(B, f"C{r}"); inp(B, f"D{r}")
    dv(B, f"B{r}", dyn('"Feat|"&RS')); dv(B, f"C{r}", ABL, False); dv(B, f"D{r}", ABL, False)
    k = f'ERA&"|"&$B{r}'
    out(B, f"E{r}", f'=IFERROR({tFe.lk("points", k)},0)', center=True)
    out(B, f"F{r}", f'=IFERROR({tFe.lk("opts", k, True)},"")')
    out(B, f"G{r}", f'=IF(B{r}="","",IF(COUNTIF({tFe.r("key")},{k})=0,"No",IF(AND(OR(E{r}<1,AND(C{r}<>"",ISNUMBER(SEARCH(C{r},F{r})))),OR(E{r}<2,AND(D{r}<>"",ISNUMBER(SEARCH(D{r},F{r}))))),"Yes","No")))', center=True)
B["B43"] = "Great Weapon Master"; B["C43"] = "STR"
label(B, "B53", "Used / available", True); out(B, "C53", '=COUNTA(B43:B52)&" of "&Calc!B25')

section(B, 54, "6  Skills: P = proficient, E = expertise (background skills are automatic)")
header(B, 55, "B", ["Skill", "Ability", "Automatic", "Your pick", "Class list", "Misc bonus", "Total", "Passive"])
SK0 = 56
for i, (s, a) in enumerate(D.SKILLS):
    r = SK0 + i
    label(B, f"B{r}", s); B[f"B{r}"].border = BOX
    out(B, f"C{r}", a, center=True)
    out(B, f"D{r}", f'=IF(OR(Calc!$B$26=B{r},Calc!$B$27=B{r},Calc!$B$24=B{r}),"Yes","")', center=True)
    inp(B, f"E{r}"); B[f"E{r}"].alignment = CENTER; dv(B, f"E{r}", '"P,E"', False)
    out(B, f"F{r}", f'=IF(ISNUMBER(SEARCH("|"&B{r}&"|",Calc!$B$28)),"Yes","")', center=True)
    inp(B, f"G{r}"); B[f"G{r}"].alignment = CENTER
    out(B, f"H{r}", f'=INDEX($I$22:$I$27,MATCH(C{r},{{"STR";"DEX";"CON";"INT";"WIS";"CHA"}},0))+IF(E{r}="E",2*PB,IF(OR(E{r}="P",D{r}="Yes"),PB,Calc!$B$29))+N(G{r})', True, "+0;-0;0", True)
    if s in ("Perception", "Investigation"): out(B, f"I{r}", f"=10+H{r}+Calc!$B$47", center=True)
    if s == "Insight": out(B, f"I{r}", f"=10+H{r}", center=True)
SK1 = SK0 + 17
B[f"E{SK0+3}"] = "P"; B[f"E{SK0+7}"] = "P"
label(B, f"B{SK1+1}", "Picks / allowed", True); out(B, f"C{SK1+1}", f'=Calc!B30&" of "&Calc!B31')
label(B, f"D{SK1+1}", "Expertise", True); out(B, f"E{SK1+1}", f'=Calc!B32&" of "&Calc!B33', center=True)

AR = SK1 + 3
section(B, AR, "7  Armor, hit points and adjustments")
fields = [("Armor worn", "Chain Mail"), ("Shield", "No"), ("Mage Armor active", "No"), ("Fighting style", "Great Weapon Fighting"), ("Hit point method", "Average"), ("Rolled HP total (levels after 1st)", None),
          ("AC adjustment", None), ("HP adjustment", None), ("Initiative adjustment", None), ("Speed adjustment", None), ("Saving throw adjustment", None)]
ROW = {}
for i, (t, v) in enumerate(fields):
    r = AR + 1 + i; ROW[t] = r
    label(B, f"B{r}", t); inp(B, f"C{r}", v)
SAVE_ADJ = ROW["Saving throw adjustment"]
dv(B, f'C{ROW["Armor worn"]}', dyn('"Armor"'), False); dv(B, f'C{ROW["Shield"]}', '"Yes,No"', False); dv(B, f'C{ROW["Mage Armor active"]}', '"Yes,No"', False)
dv(B, f'C{ROW["Fighting style"]}', dyn('"Style"'), False); dv(B, f'C{ROW["Hit point method"]}', '"Average,Rolled"', False)
# save adjustment reference was written as $C$87 above; repoint
for i in range(6):
    B[f"J{22+i}"].value = B[f"J{22+i}"].value.replace("$C$87", f"$C${SAVE_ADJ}")
cARM = f'$C${ROW["Armor worn"]}'; cSHD = f'Builder!$C${ROW["Shield"]}'; cMA = f'Builder!$C${ROW["Mage Armor active"]}'; cFS = f'Builder!$C${ROW["Fighting style"]}'

AT = SAVE_ADJ + 2
section(B, AT, "8  Attacks")
label(B, f"B{AT+1}", "Target AC for hit chance"); inp(B, f"C{AT+1}", 15)
label(B, f"B{AT+2}", "Raging"); inp(B, f"C{AT+2}", "No"); dv(B, f"C{AT+2}", '"Yes,No"', False)
label(B, f"E{AT+1}", "Attacks per Attack action", True); out(B, f"H{AT+1}", "=Calc!B34", True, center=True)
header(B, AT + 3, "B", ["Weapon", "Ability", "Proficient", "Magic bonus", "Extra damage", "Two-handed grip", "To hit", "Damage", "Avg hit", "Hit chance", "Dmg / round", "Mastery"])
AT0 = AT + 4
for i in range(5):
    r = AT0 + i; cr = 60 + i
    for c in "BCDEFG": inp(B, f"{c}{r}")
    dv(B, f"B{r}", dyn('"Weapon"')); dv(B, f"C{r}", '"Auto,STR,DEX,CON,INT,WIS,CHA"', False); dv(B, f"D{r}", '"Auto,Yes,No"', False); dv(B, f"G{r}", '"Yes,No"', False)
    out(B, f"H{r}", f'=Calc!J{cr}', True, "+0;-0;0", True)
    out(B, f"I{r}", f'=Calc!S{cr}', center=True)
    out(B, f"J{r}", f'=Calc!T{cr}', fmt="0.0", center=True)
    out(B, f"K{r}", f'=Calc!U{cr}', fmt="0%", center=True)
    out(B, f"L{r}", f'=Calc!W{cr}', fmt="0.0", center=True)
    out(B, f"M{r}", f'=Calc!X{cr}', center=True)
B[f"B{AT0}"] = "Greatsword"

# ---------- names
name("RS", "Builder!$C$5"); name("ERA", "Builder!$C$6"); name("PB", "Builder!$F$17"); name("TL", "Builder!$D$17")
name("TASHA", "Calc!$B$19")

# ---------- Calc
C = wsC
C.column_dimensions["A"].width = 34; C.column_dimensions["B"].width = 16
for j in range(3, 26): C.column_dimensions[CL(j)].width = 13
C["A1"] = "Calculation layer: no inputs here"; C["A1"].font = f(bold=True, size=12, color=NAVY)
heads = ["Class", "Level", "Die", "Sub key", "Sub level", "Sub active", "Third caster", "Caster type", "Spell ability", "Spell key", "Mode", "Cantrips", "Prepared", "Save DC", "Spell attack", "Sub HP", "Sub AC", "Eff single", "Eff multi", "Slot caster", "ASI count", "MC skill", "MC ok", "Armor"]
for j, h in enumerate(heads, 1):
    c = C.cell(3, j, h); c.font = f(bold=True); c.fill = fill(PALE)
MODS = "Builder!$I$22:$I$27"; SCORES = "Builder!$H$22:$H$27"
ABIDX = '{"STR";"DEX";"CON";"INT";"WIS";"CHA"}'
for i in range(3):
    r = 4 + i; br = 14 + i
    C[f"A{r}"] = f'=IF(AND(Builder!B{br}<>"",N(Builder!D{br})>=1,COUNTIF({tCl.r("name")},Builder!B{br})>0),Builder!B{br},"")'
    C[f"B{r}"] = f'=IF(A{r}="",0,Builder!D{br})'
    C[f"C{r}"] = f'=IF(A{r}="",0,{tCl.lk("die", f"A{r}")})'
    C[f"D{r}"] = f'=ERA&"|"&A{r}&"|"&Builder!C{br}'
    C[f"E{r}"] = f'=IFERROR({tSu.lk("level", f"D{r}")},99)'
    C[f"F{r}"] = f'=IF(AND(A{r}<>"",B{r}>=E{r}),1,0)'
    C[f"G{r}"] = f'=IF(F{r}=1,{tSu.lk("third", f"D{r}", True)},"")'
    C[f"H{r}"] = f'=IF(A{r}="","none",IF(G{r}<>"","third",{tCl.lk("caster", f"A{r}")}))'
    C[f"I{r}"] = f'=IF(H{r}="none","",IF(G{r}<>"","INT",{tCl.lk("ability", f"A{r}", True)}))'
    C[f"J{r}"] = f'=ERA&"|"&IF(G{r}<>"",G{r},A{r})'
    C[f"K{r}"] = f'=IF(H{r}="none","",{tSp.lk("mode", f"J{r}")})'
    mod = f'INDEX({MODS},MATCH(I{r},{ABIDX},0))'
    C[f"L{r}"] = f'=IF(H{r}="none",0,INDEX({tSp.mx("CL1","CL20")},MATCH(J{r},{tSp.r("key")},0),B{r}))'
    C[f"M{r}"] = f'=IF(H{r}="none",0,IF(K{r}="level",MAX(1,B{r}+{mod}),IF(K{r}="half",IF(AND(A{r}="Paladin",B{r}<2),0,MAX(1,INT(B{r}/2)+{mod})),INDEX({tSp.mx("PL1","PL20")},MATCH(J{r},{tSp.r("key")},0),B{r}))))'
    C[f"N{r}"] = f'=IF(H{r}="none",0,8+PB+{mod})'
    C[f"O{r}"] = f'=IF(H{r}="none",0,PB+{mod})'
    C[f"P{r}"] = f'=IF(F{r}=1,{tSu.lk("hpPerLevel", f"D{r}")}*B{r},0)'
    C[f"Q{r}"] = f'=IF(F{r}=1,{tSu.lk("ac", f"D{r}", True)},"")'
    C[f"R{r}"] = f'=IF(H{r}="full",B{r},IF(H{r}="half",IF(AND(ERA=2014,A{r}<>"Artificer",B{r}<2),0,ROUNDUP(B{r}/2,0)),IF(H{r}="third",ROUNDUP(B{r}/3,0),0)))'
    C[f"S{r}"] = f'=IF(H{r}="full",B{r},IF(H{r}="half",IF(OR(ERA=2024,A{r}="Artificer"),ROUNDUP(B{r}/2,0),INT(B{r}/2)),IF(H{r}="third",INT(B{r}/3),0)))'
    C[f"T{r}"] = f'=IF(OR(H{r}="full",H{r}="half",H{r}="third"),1,0)'
    C[f"U{r}"] = f'=(B{r}>=4)+(B{r}>=8)+(B{r}>=12)+(B{r}>=16)+(B{r}>=19)+IF(A{r}="Fighter",(B{r}>=6)+(B{r}>=14),0)+IF(A{r}="Rogue",--(B{r}>=10),0)'
    C[f"V{r}"] = f'=IF(OR(A{r}="Bard",A{r}="Ranger",A{r}="Rogue"),1,0)'
    sc = lambda fld: f'INDEX({SCORES},MATCH({tCl.lk(fld, f"A{r}")},{ABIDX},0))'
    C[f"W{r}"] = f'=IF(A{r}="",TRUE,IF({tCl.lk("reqMode", f"A{r}")}="ANY",OR({sc("req1")}>=13,{sc("req2")}>=13),AND({sc("req1")}>=13,IF({tCl.lk("req2", f"A{r}", True)}="",TRUE,{sc("req2")}>=13))))'
    C[f"X{r}"] = f'=IF(A{r}="","",IF(ERA=2024,{tCl.lk("armor24", f"A{r}", True)},{tCl.lk("armor14", f"A{r}", True)}))'
    # class dropdown lives on Builder; level totals by class:
for j, c in enumerate(D.CLASSES):
    r = 4 + j
    C[f"Z{r}"] = c["name"]; C[f"AA{r}"] = f'=SUMIF($A$4:$A$6,Z{r},$B$4:$B$6)'
    name("lv" + c["name"], f"Calc!$AA${r}")
C["Z3"] = "Class"; C["AA3"] = "Levels"
rows = [
 (9, "First class", '=IF(A4<>"",A4,IF(A5<>"",A5,A6))'),
 (10, "First class die", f'=IF(B9="",0,{tCl.lk("die","B9")})'),
 (11, "First class average", "=IF(B10=0,0,INT(B10/2)+1)"),
 (12, "HP from dice", f'=IF(B9="",0,IF(Builder!C{ROW["Hit point method"]}="Rolled",B10+N(Builder!C{ROW["Rolled HP total (levels after 1st)"]}),B10-B11+SUMPRODUCT((INT(C4:C6/2)+(C4:C6>0))*B4:B6)))'),
 (13, "Hit points", f'=MAX(TL,B12+mCON*TL+IFERROR({tSe.lk("hp",SPKEY.replace("$C$35","Builder!$C$35"))},0)*TL+SUM(P4:P6)+B37*TL+B38+N(Builder!C{ROW["HP adjustment"]}))'),
 (14, "Hit dice", '=IFERROR(LEFT(B15,LEN(B15)-3),"")'),
 (15, "Hit dice raw", '=IF(SUMIF(C4:C6,12,B4:B6)>0,SUMIF(C4:C6,12,B4:B6)&"d12 + ","")&IF(SUMIF(C4:C6,10,B4:B6)>0,SUMIF(C4:C6,10,B4:B6)&"d10 + ","")&IF(SUMIF(C4:C6,8,B4:B6)>0,SUMIF(C4:C6,8,B4:B6)&"d8 + ","")&IF(SUMIF(C4:C6,6,B4:B6)>0,SUMIF(C4:C6,6,B4:B6)&"d6 + ","")'),
 (16, "Slot casters", "=SUM(T4:T6)"),
 (17, "Effective caster level", "=IF(B16=1,SUM(R4:R6),SUM(S4:S6))"),
 (19, "Tasha ruleset", '=IF(ISNUMBER(SEARCH("Tasha",RS)),1,0)'),
 (20, "Tasha floats species bonus", f'=IF(AND(B19=1,IFERROR({tSe.lk("tasha1",SPKEY.replace("$C$35","Builder!$C$35"))},0)>0),1,0)'),
 (21, "Ability cap", "=IF(B46>0,30,20)"),
 (22, "First class saves", f'=IF(B9="","",{tCl.lk("saves","B9")})'),
 (23, "Aura of Protection", "=IF(lvPaladin>=6,MAX(1,mCHA),0)"),
 (24, "Species fixed skill", f'=IFERROR({tSe.lk("skill",SPKEY.replace("$C$35","Builder!$C$35"),True)},"")'),
 (25, "Feat slots", f'=SUM(U4:U6)+IFERROR({tSe.lk("feat",SPKEY.replace("$C$35","Builder!$C$35"))},0)'),
 (26, "Background skill 1", f'=IFERROR({tBg.lk("sk1",BGKEY.replace("$C$36","Builder!$C$36"))},"")'),
 (27, "Background skill 2", f'=IFERROR({tBg.lk("sk2",BGKEY.replace("$C$36","Builder!$C$36"))},"")'),
 (28, "Class skill list", f'=IF(B9="","","|"&IF(ERA=2024,{tCl.lk("skills24","B9")},{tCl.lk("skills14","B9")})&"|")'),
 (29, "Jack of All Trades", "=IF(lvBard>=2,INT(PB/2),0)"),
 (30, "Skill picks used", f'=SUMPRODUCT((Builder!E{SK0}:E{SK1}<>"")*(Builder!D{SK0}:D{SK1}<>"Yes"))'),
 (31, "Skill picks allowed", f'=IF(B9="",0,{tCl.lk("skillN","B9")})+SUM(V4:V6)-IF(OR(B9="Bard",B9="Ranger",B9="Rogue"),1,0)+IFERROR({tSe.lk("extra",SPKEY.replace("$C$35","Builder!$C$35"))},0)+B44'),
 (32, "Expertise used", f'=COUNTIF(Builder!E{SK0}:E{SK1},"E")'),
 (33, "Expertise allowed", '=IF(ERA=2024,IF(lvBard>=9,4,IF(lvBard>=2,2,0)),IF(lvBard>=10,4,IF(lvBard>=3,2,0)))+IF(lvRogue>=6,4,IF(lvRogue>=1,2,0))+IF(AND(ERA=2024,lvRanger>=9),2,0)+B45'),
 (34, "Attacks per action", "=IF(lvFighter>=20,4,IF(lvFighter>=11,3,IF(OR(lvFighter>=5,lvBarbarian>=5,lvMonk>=5,lvPaladin>=5,lvRanger>=5),2,1)))"),
 (35, "Rage damage", "=IF(lvBarbarian>=16,4,IF(lvBarbarian>=9,3,IF(lvBarbarian>=1,2,0)))"),
 (36, "Martial Arts avg", "=IF(lvMonk=0,0,IF(ERA=2024,IF(lvMonk>=17,6.5,IF(lvMonk>=11,5.5,IF(lvMonk>=5,4.5,3.5))),IF(lvMonk>=17,5.5,IF(lvMonk>=11,4.5,IF(lvMonk>=5,3.5,2.5)))))"),
]
for r, t, v in rows:
    C[f"A{r}"] = t; C[f"B{r}"] = v
C["A18"] = "Martial Arts die"; C["B18"] = '=IF(lvMonk=0,"",IF(ERA=2024,IF(lvMonk>=17,"1d12",IF(lvMonk>=11,"1d10",IF(lvMonk>=5,"1d8","1d6"))),IF(lvMonk>=17,"1d10",IF(lvMonk>=11,"1d8",IF(lvMonk>=5,"1d6","1d4")))))'
# feat effects: rows 37-47 are totals across the block at D37:N47
FXROW = {"hpPerLevel": 37, "hpFlat": 38, "speed": 39, "initPB": 40, "initFlat": 41, "mam": 42, "skills": 44, "expert": 45, "boon": 46, "passive": 47}
for k, r in FXROW.items():
    col = CL(5 + FX.index(k))
    C[f"A{r}"] = "Feat total: " + k; C[f"B{r}"] = f"=SUM({col}38:{col}48)"
C["D37"] = "Feat key"
for j, k in enumerate(FX): C.cell(37, 5 + j, k).font = f(bold=True)
for i, br in enumerate(list(FR) + [None]):
    r = 38 + i
    C[f"D{r}"] = f'=ERA&"|"&Builder!B{br}' if br else '=IF(ERA=2024,ERA&"|"&Builder!C38,"none")'
    for j, k in enumerate(FX):
        C.cell(r, 5 + j, f'=IFERROR({tFe.lk(k, f"D{r}")},0)')
# armor
ARM = f'Builder!{cARM}'
cal = [
 (49, "Armor kind", f'=IFERROR({tAr.lk("kind",ARM)},"None")'),
 (50, "Armor base", f'=IFERROR({tAr.lk("base",ARM)},10)'),
 (51, "Armor dex cap", f'=IF(B49="Medium",2+B42,IFERROR({tAr.lk("dexCap",ARM)},99))'),
 (52, "Armor str need", f'=IFERROR({tAr.lk("str",ARM)},0)'),
 (53, "Stealth disadvantage", f'=IFERROR({tAr.lk("stealth",ARM)},0)'),
 (54, "Short of Strength", '=IF(AND(B52>0,sSTR<B52,NOT(AND(ERA=2014,LEFT(Builder!$C$35,5)="Dwarf"))),1,0)'),
 (55, "Base AC", f'=IF(B49<>"None",B50+MIN(B51,IF(B49="Heavy",0,mDEX)),MAX(10+mDEX,IF(lvBarbarian>0,10+mDEX+mCON,-99),IF(AND(lvMonk>0,{cSHD}<>"Yes"),10+mDEX+mWIS,-99),IF(COUNTIF(Q4:Q6,"13+DEX")>0,13+mDEX,-99),IF(COUNTIF(Q4:Q6,"10+DEX+CHA")>0,10+mDEX+mCHA,-99),IF({cMA}="Yes",13+mDEX,-99)))'),
 (56, "Armor Class", f'=B55+IF({cSHD}="Yes",2,0)+IF(AND(B49<>"None",{cFS}="Defense"),1,0)+N(Builder!C{ROW["AC adjustment"]})'),
 (57, "Speed", f'=IFERROR({tSe.lk("speed",SPKEY.replace("$C$35","Builder!$C$35"))},30)+IF(AND(lvBarbarian>=5,B49<>"Heavy"),10,0)+IF(AND(lvMonk>=2,B49="None",{cSHD}<>"Yes"),IF(lvMonk>=18,30,IF(lvMonk>=14,25,IF(lvMonk>=10,20,IF(lvMonk>=6,15,10)))),0)+IF(AND(ERA=2024,lvRanger>=6,B49<>"Heavy"),10,0)-IF(B54=1,10,0)+B39+N(Builder!C{ROW["Speed adjustment"]})'),
 (58, "Initiative", f'=mDEX+B29+B40*PB+B41+N(Builder!C{ROW["Initiative adjustment"]})'),
]
for r, t, v in cal:
    C[f"A{r}"] = t; C[f"B{r}"] = v
# attacks at row 60+
hd = ["Weapon", "Found", "Melee", "Monk weapon", "Finesse", "Ability", "Ability mod", "Prof list", "Proficient", "To hit", "Has versatile", "Two-hand", "GWF", "Die avg", "Dice text", "Dueling", "Rage", "Flat", "Damage text", "Avg hit", "Chance", "Expected", "Per round", "Mastery"]
for j, h in enumerate(hd, 1):
    c = C.cell(59, j, h); c.font = f(bold=True); c.fill = fill(PALE)
E = lambda fld, r: f'IF(ERA=2024,{tWe.lk(fld+"24", f"$A{r}")},{tWe.lk(fld+"14", f"$A{r}")})'
for i in range(5):
    r = 60 + i; br = AT0 + i
    C[f"A{r}"] = f'=IF(Builder!B{br}="","",Builder!B{br})'
    C[f"B{r}"] = f'=IF(A{r}="",0,COUNTIF({tWe.r("name")},A{r}))'
    C[f"C{r}"] = f'=IF(B{r}=0,0,IF({tWe.lk("range", f"$A{r}")}="Melee",1,0))'
    C[f"D{r}"] = f'=IF(OR(B{r}=0,lvMonk=0),0,{E("monk", r)})'
    C[f"E{r}"] = f'=IF(B{r}=0,0,{tWe.lk("finesse", f"$A{r}")})'
    C[f"F{r}"] = f'=IF(B{r}=0,"",IF(AND(Builder!C{br}<>"",Builder!C{br}<>"Auto"),Builder!C{br},IF(C{r}=0,"DEX",IF(OR(E{r}=1,D{r}=1),IF(mDEX>mSTR,"DEX","STR"),"STR"))))'
    C[f"G{r}"] = f'=IF(B{r}=0,0,INDEX({MODS},MATCH(F{r},{ABIDX},0)))'
    C[f"H{r}"] = f'=IF(B{r}=0,"","|"&{E("prof", r)}&"|")'
    C[f"I{r}"] = f'=IF(B{r}=0,0,IF(Builder!D{br}="Yes",1,IF(Builder!D{br}="No",0,IF(SUMPRODUCT(($A$4:$A$6<>"")*ISNUMBER(SEARCH("|"&$A$4:$A$6&"|",H{r})))>0,1,0))))'
    C[f"J{r}"] = f'=IF(B{r}=0,"",G{r}+I{r}*PB+N(Builder!E{br})+IF(AND({cFS}="Archery",C{r}=0),2,0))'
    C[f"K{r}"] = f'=IF(B{r}=0,0,{E("hasV", r)})'
    C[f"L{r}"] = f'=IF(AND(Builder!G{br}="Yes",K{r}=1),1,0)'
    C[f"M{r}"] = f'=IF(B{r}=0,0,IF(AND({cFS}="Great Weapon Fighting",C{r}=1,OR({tWe.lk("twoH", f"$A{r}")}=1,L{r}=1)),1,0))'
    raw = f'IF(M{r}=1,IF(L{r}=1,{E("vgwf", r)},{E("gwf", r)}),IF(L{r}=1,{E("vavg", r)},{E("avg", r)}))'
    C[f"N{r}"] = f'=IF(B{r}=0,0,IF(AND(D{r}=1,M{r}=0,$B$36>{raw}),$B$36,{raw}))'
    rtxt = f'IF(L{r}=1,{E("vtxt", r)},{E("txt", r)})'
    C[f"O{r}"] = f'=IF(B{r}=0,"",IF(AND(D{r}=1,M{r}=0,$B$36>{raw}),$B$18,{rtxt}&""))'
    C[f"P{r}"] = f'=IF(AND(B{r}>0,{cFS}="Dueling",C{r}=1,L{r}=0),IF({tWe.lk("twoH", f"$A{r}")}=1,0,2),0)'
    C[f"Q{r}"] = f'=IF(AND(B{r}>0,Builder!$C${AT+2}="Yes",C{r}=1,F{r}="STR"),$B$35,0)'
    C[f"R{r}"] = f'=IF(B{r}=0,0,G{r}+N(Builder!E{br})+N(Builder!F{br})+P{r}+Q{r})'
    C[f"S{r}"] = f'=IF(B{r}=0,"",O{r}&IF(R{r}>=0,"+","")&R{r}&" "&{tWe.lk("type", f"$A{r}")})'
    C[f"T{r}"] = f'=IF(B{r}=0,"",N{r}+R{r})'
    C[f"U{r}"] = f'=IF(B{r}=0,"",MAX(0.05,MIN(0.95,(21+J{r}-N(Builder!$C${AT+1}))/20)))'
    C[f"V{r}"] = f'=IF(B{r}=0,0,(U{r}-0.05)*T{r}+0.05*(T{r}+N{r}))'
    C[f"W{r}"] = f'=IF(B{r}=0,"",V{r}*$B$34)'
    C[f"X{r}"] = f'=IF(OR(B{r}=0,ERA<>2024),"",{tWe.lk("mastery", f"$A{r}", True)})'
# picks check helpers
C["A66"] = "Pick used 1"; C["B66"] = '=IF(AND(Builder!C31<>"",Builder!D31>0),Builder!C31,"")'
C["A67"] = "Pick used 2"; C["B67"] = '=IF(AND(Builder!C32<>"",Builder!D32>0),Builder!C32,"")'
C["A68"] = "Pick used 3"; C["B68"] = '=IF(AND(Builder!C33<>"",Builder!D33>0),Builder!C33,"")'
C["A69"] = "Picks needed"; C["B69"] = "=(Builder!D31>0)+(Builder!D32>0)+(Builder!D33>0)"
C["A70"] = "Picks made"; C["B70"] = '=(B66<>"")+(B67<>"")+(B68<>"")'
C["A71"] = "Picks distinct"; C["B71"] = '=NOT(OR(AND(B66<>"",B66=B67),AND(B66<>"",B66=B68),AND(B67<>"",B67=B68)))'
C["A72"] = "Bg abilities"; C["B72"] = f'=IF(ERA=2024,IFERROR({tBg.lk("abilities",BGKEY.replace("$C$36","Builder!$C$36"))},""),"")'
C["A73"] = "Picks in bg list"; C["B73"] = '=IF(B72="",TRUE,AND(OR(B66="",ISNUMBER(SEARCH(B66,B72))),OR(B67="",ISNUMBER(SEARCH(B67,B72))),OR(B68="",ISNUMBER(SEARCH(B68,B72)))))'
C["A74"] = "No-pick ability"; C["B74"] = f'=IF(AND(ERA=2014,B19=0),IFERROR({tSe.lk("noPick",SPKEY.replace("$C$35","Builder!$C$35"),True)},""),"")'
C["A75"] = "Armor letter"; C["B75"] = '=IF(B49="None","",LEFT(B49,1))'
C["A76"] = "Armor trained"; C["B76"] = '=IF(B75="",TRUE,ISNUMBER(SEARCH(B75,X4&X5&X6)))'
C["A77"] = "Shield trained"; C["B77"] = f'=IF({cSHD}<>"Yes",TRUE,ISNUMBER(SEARCH("S",X4&X5&X6)))'
C["A78"] = "Attunement slots"; C["B78"] = "=IF(lvArtificer>=18,6,IF(lvArtificer>=14,5,IF(lvArtificer>=10,4,3)))"
C["A79"] = "XP for next level"; C["B79"] = f'=IF(AND(TL>=1,TL<20),INDEX({tSl.r("xp")},TL+1),0)'
C["A80"] = "Carry capacity"; C["B80"] = f'=sSTR*15*IFERROR({tSe.lk("carry",SPKEY.replace("$C$35","Builder!$C$35"))},1)'
for row in C.iter_rows(min_row=3, max_row=80):
    for c in row:
        if c.font.name != FONT: c.font = f()

# ---------- Builder side panel
section(B, 4, "Live results", "O", "P")
LIVE = [("Level", "=TL"), ("Prof bonus", "=PB"), ("Hit points", "=Calc!B13"), ("Armor Class", "=Calc!B56"), ("Initiative", "=Calc!B58"), ("Speed (ft)", "=Calc!B57"), ("Passive Perc.", f"=I{SK0+11}")]
for i, (t, v) in enumerate(LIVE):
    label(B, f"O{5+i}", t, True); c = out(B, f"P{5+i}", v, True); c.alignment = Alignment(horizontal="left")
section(B, 13, "Rules check", "O", "P")
sk = lambda p: f"Calc!{p}"
CHK = [
 (f'IF(C20="Point Buy",AND(MIN(C22:C27)>=8,MAX(C22:C27)<=15,L28<=27),IF(C20="Standard Array",AND(COUNTIF(C22:C27,15)=1,COUNTIF(C22:C27,14)=1,COUNTIF(C22:C27,13)=1,COUNTIF(C22:C27,12)=1,COUNTIF(C22:C27,10)=1,COUNTIF(C22:C27,8)=1),AND(MIN(C22:C27)>=3,MAX(C22:C27)<=18)))',
  '=IF(C20="Point Buy","Point buy: "&L28&" of 27 points, each score 8 to 15",IF(C20="Standard Array","Standard array: scores must be 15, 14, 13, 12, 10, 8","Rolled scores: each between 3 and 18"))'),
 ("AND(TL>=1,TL<=20)", '="Total level is "&TL&" (must be 1 to 20)"'),
 (f'SUMPRODUCT((Calc!A4:A6<>"")*(COUNTIFS({LA},"Class|"&RS,Lists!$B$1:$B${LN},Calc!A4:A6)=0))=0', "Classes are available in this rulebook"),
 (f'COUNTIFS({LA},"Species|"&RS,Lists!$B$1:$B${LN},C35)>0', "Species is available in this rulebook"),
 (f'COUNTIFS({LA},"Background|"&RS,Lists!$B$1:$B${LN},C36)>0', "Background is available in this rulebook"),
 ('AND(Calc!B70=Calc!B69,Calc!B71,Calc!B73,OR(Calc!B74="",AND(Calc!B66<>Calc!B74,Calc!B67<>Calc!B74,Calc!B68<>Calc!B74)))', '=IF(ERA=2024,"Background ability bonuses go to different allowed abilities","Species ability choices are complete and different")'),
 ("COUNTA(B43:B52)<=Calc!B25", '="Feats and ASIs: "&COUNTA(B43:B52)&" used of "&Calc!B25&" available"'),
 ('COUNTIF(G43:G52,"No")=0', "Each feat's ability increase is allowed and a feat from this rulebook"),
 ("Calc!B30<=Calc!B31", '="Skills: "&Calc!B30&" chosen of "&Calc!B31&" allowed"'),
 ("Calc!B32<=Calc!B33", '="Expertise: "&Calc!B32&" used of "&Calc!B33&" allowed"'),
 ('OR(COUNTIF(Calc!A4:A6,"?*")<2,AND(Calc!W4,Calc!W5,Calc!W6))', "Multiclass ability requirements (13+) are met"),
 ("AND(Calc!B76,Calc!B77)", "Worn armor and shield match class training"),
 ("Calc!B54=0", "Strength meets the worn armor's requirement"),
 ('SUMPRODUCT((Calc!A4:A6<>"")*(C14:C16<>"")*(Calc!F4:F6=0))=0', "Subclasses are unlocked at their class level"),
]
CH0 = 14
for i, (ok, txt) in enumerate(CHK):
    r = CH0 + i
    c = out(B, f"O{r}", f'=IF({ok},"OK","FIX")', True, center=True)
    out(B, f"P{r}", txt if txt.startswith("=") else txt)
CH1 = CH0 + len(CHK) - 1
B.conditional_formatting.add(f"O{CH0}:O{CH1}", CellIsRule(operator="equal", formula=['"FIX"'], fill=fill("F8D7D3"), font=Font(name=FONT, bold=True, color=RED)))
B.conditional_formatting.add(f"O{CH0}:O{CH1}", CellIsRule(operator="equal", formula=['"OK"'], fill=fill("DCEFDC"), font=Font(name=FONT, bold=True, color="1E6B2B")))
B.conditional_formatting.add("G43:G52", CellIsRule(operator="equal", formula=['"No"'], fill=fill("F8D7D3")))
label(B, f"O{CH1+1}", "Open", True); out(B, f"P{CH1+1}", f'=COUNTIF(O{CH0}:O{CH1},"FIX")&" rule issue(s) to fix"', True)
B.freeze_panes = "A4"

# ---------- Sheet
S = wsS
S.column_dimensions["A"].width = 2
for j, w in enumerate([22, 11, 9, 3, 22, 9, 9, 3, 22, 10, 11, 13], 2): S.column_dimensions[CL(j)].width = w
S["B1"] = "=Builder!C7"; S["B1"].font = f(bold=True, size=18, color=NAVY)
cl = lambda i: f'IF(Calc!A{i}="","",Calc!A{i}&IF(Builder!C{10+i}<>""," ("&Builder!C{10+i}&")","")&" "&Calc!B{i})'
S["B2"] = f'={cl(4)}&IF(Calc!A5<>"","  /  "&{cl(5)},"")&IF(Calc!A6<>"","  /  "&{cl(6)},"")'; S["B2"].font = f(bold=True, size=11)
S["B3"] = '=Builder!C35&"  |  "&Builder!C36&"  |  "&Builder!C9&"  |  "&RS'; S["B3"].font = f(color="555555")
big = [("LEVEL", "=TL"), ("PROF", "=PB"), ("HP", "=Calc!B13"), ("AC", "=Calc!B56"), ("INIT", "=Calc!B58"), ("SPEED", "=Calc!B57"), ("HIT DICE", "=Calc!B14")]
ADDR = ["B", "C", "D", "F", "G", "H", "J"]
for (t, v), c in zip(big, ADDR):
    a = S[f"{c}5"]; a.value = t; a.font = f(bold=True, size=8, color="FFFFFF"); a.fill = fill(NAVY); a.alignment = CENTER
    b = S[f"{c}6"]; b.value = v; b.font = f(bold=True, size=15, color=NAVY); b.alignment = CENTER; b.border = BOX
S["C6"].number_format = "+0"; S["G6"].number_format = "+0;-0;0"
S.row_dimensions[6].height = 30
section(S, 8, "Abilities and saves", "B", "D")
header(S, 9, "B", ["Ability", "Score (mod)", "Save"])
for i, a in enumerate(AB):
    r = 10 + i
    label(S, f"B{r}", D.ABILITY_NAMES[a]); S[f"B{r}"].border = BOX
    out(S, f"C{r}", f'=s{a}&" ("&TEXT(m{a},"+0;-0;0")&")"', True, center=True)
    out(S, f"D{r}", f'=TEXT(Builder!J{22+i},"+0;-0;0")&IF(Builder!K{22+i}="Yes"," *","")', center=True)
label(S, "B16", "* proficient save"); S["B16"].font = f(italic=True, size=8)
section(S, 8, "Skills", "F", "H")
header(S, 9, "F", ["Skill", "Ability", "Bonus"])
for i, (s, a) in enumerate(D.SKILLS):
    r = 10 + i
    out(S, f"F{r}", f'=IF(OR(Builder!E{SK0+i}="E"),"** ",IF(OR(Builder!E{SK0+i}="P",Builder!D{SK0+i}="Yes"),"* ",""))&"{s}"')
    out(S, f"G{r}", a, center=True)
    out(S, f"H{r}", f"=Builder!H{SK0+i}", True, "+0;-0;0", True)
label(S, "B18", "Senses and movement", True)
sens = [("Passive Perception", f"=Builder!I{SK0+11}"), ("Passive Investigation", f"=Builder!I{SK0+8}"), ("Passive Insight", f"=Builder!I{SK0+6}"),
        ("Darkvision (ft)", f'=IFERROR({tSe.lk("dv",SPKEY.replace("$C$35","Builder!$C$35"))},0)'), ("Size", f'=IFERROR({tSe.lk("size",SPKEY.replace("$C$35","Builder!$C$35"))},"")'),
        ("Carry capacity (lb)", "=Calc!B80"), ("Push / drag / lift (lb)", "=Calc!B80*2"), ("Long / high jump (ft)", '=sSTR&" / "&MAX(0,3+mSTR)'), ("Attunement slots", "=Calc!B78"),
        ("Next level at (XP)", '=IF(Calc!B79=0,"max",Calc!B79)')]
for i, (t, v) in enumerate(sens):
    r = 19 + i
    label(S, f"B{r}", t); S[f"B{r}"].border = BOX; out(S, f"C{r}", v, center=True); S.merge_cells(f"C{r}:D{r}")
S["J8"].value = None
section(S, 8, "Attacks", "J", "M")
header(S, 9, "J", ["Weapon", "To hit", "Damage", "Dmg/round"])
for i in range(5):
    r = 10 + i; br = AT0 + i
    out(S, f"J{r}", f'=IF(Builder!B{br}="","",Builder!B{br})'); out(S, f"K{r}", f"=Builder!H{br}", True, "+0;-0;0", True)
    out(S, f"L{r}", f"=Builder!I{br}", center=True); out(S, f"M{r}", f"=Builder!L{br}", fmt="0.0", center=True)
label(S, "J15", f'="Attacks per action: "&Calc!B34&"  |  versus AC "&Builder!C{AT+1}'); S["J15"].font = f(italic=True, size=8)
section(S, 17, "Spellcasting", "J", "M")
header(S, 18, "J", ["Class", "Save DC", "Spell atk", "Cantrip/Spell"])
for i in range(3):
    r = 19 + i; cr = 4 + i
    out(S, f"J{r}", f'=IF(Calc!H{cr}="none","",Calc!A{cr})'); out(S, f"K{r}", f'=IF(Calc!H{cr}="none","",Calc!N{cr})', True, center=True)
    out(S, f"L{r}", f'=IF(Calc!H{cr}="none","",Calc!O{cr})', True, "+0;-0;0", True); out(S, f"M{r}", f'=IF(Calc!H{cr}="none","",Calc!L{cr}&" / "&Calc!M{cr})', center=True)
label(S, "J22", "Spell slots by level", True)
for i in range(9):
    col = "JKLM"[i % 4] ; r = 23 + (i // 4) * 2
    c1 = S[f"{col}{r}"]; c1.value = f"L{i+1}"; c1.font = f(bold=True, size=8); c1.fill = fill(PALE); c1.alignment = CENTER; c1.border = BOX
    out(S, f"{col}{r+1}", f'=IF(Calc!$B$17>0,INDEX({tSl.mx("S1","S9")},MIN(Calc!$B$17,20),{i+1}),0)', True, center=True)
    name(f"slot_{i+1}", f"Sheet!${col}${r+1}")
c = S["K27"]; c.value = "Pact slots"; c.font = f(bold=True, size=8); c.fill = fill(PALE); c.alignment = CENTER; c.border = BOX
out(S, "K28", f'=IF(lvWarlock>0,INDEX({tSl.r("pactN")},lvWarlock)&" at L"&INDEX({tSl.r("pactL")},lvWarlock),"none")', True, center=True)
S.merge_cells("K27:M27"); S.merge_cells("K28:M28")
section(S, 30, "Class resources", "B", "H")
header(S, 31, "B", ["Resource", "Class", "Value"]); header(S, 31, "F", ["Recovers", "", ""]); S.merge_cells("F31:H31")
for i in range(16):
    r = 32 + i; k = i + 1
    run = tRe.r("run")
    m = f"MATCH({k},{run},0)"
    out(S, f"B{r}", f'=IFERROR(INDEX({tRe.r("name")},{m}),"")'); out(S, f"C{r}", f'=IFERROR(INDEX({tRe.r("cls")},{m}),"")', center=True)
    out(S, f"D{r}", f'=IFERROR(INDEX({tRe.r("value")},{m}),"")', True, center=True); out(S, f"F{r}", f'=IFERROR(INDEX({tRe.r("rest")},{m}),"")')
    S.merge_cells(f"F{r}:H{r}")
section(S, 30, "Class features", "J", "M")
for i in range(46):
    r = 31 + i; k = i + 1
    c = S[f"J{r}"]; c.value = f'=IFERROR(INDEX({tFt.r("label")},MATCH({k},{tFt.r("run")},0)),"")'; c.font = f(size=9)
    S.merge_cells(f"J{r}:M{r}")
section(S, 49, "Traits and resistances", "B", "H")
for r, t, v in [(50, "Species traits", f'=IFERROR({tSe.lk("traits",SPKEY.replace("$C$35","Builder!$C$35"))},"")'), (51, "Resistances", f'=IFERROR({tSe.lk("resist",SPKEY.replace("$C$35","Builder!$C$35"),True)},"")'),
                (52, "Origin feat", "=Builder!C38"), (53, "Tools / languages", "=Builder!C40"), (54, "Armor notes", '=IF(Calc!B53=1,"Disadvantage on Stealth from armor. ","")&IF(Calc!B54=1,"Strength too low for this armor: speed -10.","")')]:
    label(S, f"B{r}", t, True); out(S, f"C{r}", v); S[f"C{r}"].alignment = WRAP; S.merge_cells(f"C{r}:H{r}"); S.row_dimensions[r].height = 26 if r == 50 else 15
section(S, 56, "Feats taken", "B", "H")
for i, br in enumerate(FR):
    r = 57 + i
    out(S, f"B{r}", f'=IF(Builder!B{br}="","",Builder!B{br}&IF(Builder!C{br}<>""," (+"&Builder!C{br}&IF(AND(Builder!D{br}<>"",Builder!E{br}>=2),", +"&Builder!D{br},"")&")",""))'); S.merge_cells(f"B{r}:H{r}")
S.page_setup.orientation = "portrait"; S.page_setup.fitToWidth = 1; S.page_setup.fitToHeight = 0
S.sheet_properties.pageSetUpPr.fitToPage = True
S.print_area = "A1:M77"

# ---------- Spells
P = wsP
P["B1"] = "Spell tracker"; P["B1"].font = f(bold=True, size=16, color=NAVY)
P["B2"] = "Write your spells in the yellow rows. The first row is an example. Spell text is not included: look it up in your rulebook."; P["B2"].font = f(italic=True, color="555555")
for c, w in zip("ABCDEFGHI", [2, 28, 8, 14, 13, 11, 11, 9, 34]): P.column_dimensions[c].width = w
header(P, 4, "B", ["Class", "Save DC", "Spell attack", "Cantrips", "Spells (max)", "Marked prepared"])
for i in range(3):
    r = 5 + i; cr = 4 + i
    out(P, f"B{r}", f'=IF(Calc!H{cr}="none","",Calc!A{cr})'); out(P, f"C{r}", f'=IF(Calc!H{cr}="none","",Calc!N{cr})', center=True); out(P, f"D{r}", f'=IF(Calc!H{cr}="none","",Calc!O{cr})', fmt="+0;-0;0", center=True)
    out(P, f"E{r}", f'=IF(Calc!H{cr}="none","",Calc!L{cr})', center=True); out(P, f"F{r}", f'=IF(Calc!H{cr}="none","",Calc!M{cr})', center=True)
    out(P, f"G{r}", f'=IF(B{r}="","",COUNTIFS($E$11:$E$70,B{r},$F$11:$F$70,"Yes"))', center=True)
header(P, 10, "B", ["Spell", "Level", "School", "Class", "Prepared", "Concentr.", "Ritual", "Notes"])
for r in range(11, 71):
    for c in "BCDEFGHI": inp(P, f"{c}{r}")
    dv(P, f"E{r}", dyn('"Class|"&RS')); dv(P, f"F{r}", '"Yes,No"', False); dv(P, f"G{r}", '"Yes,No"', False); dv(P, f"H{r}", '"Yes,No"', False)
for c, v in zip("BCDEFGHI", ["Bless", 1, "Enchantment", "Paladin", "Yes", "Yes", "No", "Example row: overwrite"]): P[f"{c}11"] = v
P.freeze_panes = "A11"

# ---------- Inventory
I = wsI
I["B1"] = "Inventory and wealth"; I["B1"].font = f(bold=True, size=16, color=NAVY)
I["B2"] = "Add gear in the yellow rows. The first row is an example."; I["B2"].font = f(italic=True, color="555555")
for c, w in zip("ABCDEFGH", [2, 30, 8, 11, 11, 11, 11, 3]): I.column_dimensions[c].width = w
header(I, 4, "B", ["Coin", "Amount", "lb", "Value (gp)"])
for i, (t, v) in enumerate([("Platinum (pp)", 10), ("Gold (gp)", 1), ("Electrum (ep)", 0.5), ("Silver (sp)", 0.1), ("Copper (cp)", 0.01)]):
    r = 5 + i
    label(I, f"B{r}", t); I[f"B{r}"].border = BOX; inp(I, f"C{r}", 25 if i == 1 else None)
    out(I, f"D{r}", f"=N(C{r})/50", fmt="0.00", center=True); out(I, f"E{r}", f"=N(C{r})*{v}", fmt="0.00", center=True)
label(I, "B10", "Coin total", True); out(I, "D10", "=SUM(D5:D9)", True, "0.00", True); out(I, "E10", "=SUM(E5:E9)", True, "0.00", True)
header(I, 12, "B", ["Item", "Qty", "lb each", "gp each", "Attuned", "Total lb"])
label(I, "B13", "Worn armor (automatic)"); I["B13"].border = BOX; out(I, "C13", 1, center=True)
out(I, "D13", f'=IFERROR({tAr.lk("lb","Builder!"+cARM)},0)', center=True); out(I, "E13", f'=IFERROR({tAr.lk("gp","Builder!"+cARM)},0)', center=True); out(I, "F13", "", center=True); out(I, "G13", "=C13*D13", center=True)
label(I, "B14", "Shield (automatic)"); I["B14"].border = BOX; out(I, "C14", f'=IF({cSHD}="Yes",1,0)', center=True); out(I, "D14", 6, center=True); out(I, "E14", 10, center=True); out(I, "F14", "", center=True); out(I, "G14", "=C14*D14", center=True)
for r in range(15, 55):
    for c in "BCDEF": inp(I, f"{c}{r}")
    dv(I, f"F{r}", '"Yes,No"', False)
    out(I, f"G{r}", f"=N(C{r})*N(D{r})", center=True)
for c, v in zip("BCDE", ["Greatsword", 1, 6, 50]): I[f"{c}15"] = v
label(I, "I4", "Summary", True); I.column_dimensions["I"].width = 26; I.column_dimensions["J"].width = 14
for r, t, v in [(5, "Total weight (lb)", "=SUM(G13:G54)+D10"), (6, "Carry capacity (lb)", "=Calc!B80"), (7, "Load status", '=IF(J5>J6,"Over capacity",IF(J5>J6*2/3,"Heavy load","Fine"))'),
                (8, "Gear value (gp)", "=SUMPRODUCT(C13:C54*1,E13:E54*1)"), (9, "Attuned items", '=COUNTIF(F15:F54,"Yes")&" of "&Calc!B78')]:
    label(I, f"I{r}", t); out(I, f"J{r}", v, center=True)
I["J8"].value = "=SUMPRODUCT(N(+C13)*N(+E13))+SUMPRODUCT(N(+C14)*N(+E14))+SUMPRODUCT(C15:C54,E15:E54)"
I.freeze_panes = "A13"

# ---------- Guide
G = wsG
G.column_dimensions["A"].width = 2; G.column_dimensions["B"].width = 112
lines = [("Hero Ledger: D&D 5e character builder", "t"),
 ("Unofficial fan tool. Pick a rulebook on the Builder tab, fill the yellow cells, and read the finished character on the Sheet tab.", ""),
 ("How to use it", "h"),
 ("1. Builder tab: choose the Rulebook first. Class, species, background, feat and subclass dropdowns follow that choice.", ""),
 ("2. Yellow cells with blue text are yours. All other cells calculate; leave them alone.", ""),
 ("3. Watch the Rules check panel on the right of the Builder. FIX means a choice breaks the selected rulebook; it never blocks you.", ""),
 ("4. After switching rulebook, re-pick anything the check flags. Old picks stay in place until you change them.", ""),
 ("5. Sheet tab prints on one page wide. Spells and Inventory are trackers. Calc shows the working. Grey-blue tabs hold the rules tables.", ""),
 ("Rulebook options", "h"),
 ("2024 Player's Handbook: the current rules (also called 5.5e or 5e 2024).", ""),
 ("2024 PHB + Forge of the Artificer: adds the 2024 Artificer from Eberron: Forge of the Artificer.", ""),
 ("2014 Player's Handbook: the original 5e rules.", ""),
 ("2014 PHB + Tasha's Cauldron: adds the Artificer, Tasha's subclasses and feats, and Tasha's floating species bonuses.", ""),
 ("What it calculates", "h"),
 ("Ability scores from every source with caps and capstones; modifiers; saves; all 18 skills with expertise and Jack of All Trades; passive senses; hit points and hit dice; AC from armor, Unarmored Defense, Draconic resilience or Mage Armor; speed; initiative.", ""),
 ("Attack bonus, damage text, hit chance and expected damage per round; weapon mastery; fighting styles (Archery, Defense, Dueling, Great Weapon Fighting).", ""),
 ("Spell save DC and attack per class; cantrips and prepared or known spells; multiclass spell slots with the right rounding per rulebook; Pact Magic; class resources and features by level; carrying, jumping, XP to next level, attunement.", ""),
 ("Not automated", "h"),
 ("Subclass features, invocations, spell text and magic items are not included or applied. Use the adjustment cells and Misc bonus column for anything else that changes a number.", ""),
 ("Multiclass armor and weapon proficiencies are simplified: any class's training counts. Set Proficient to Yes or No on an attack row when your case differs.", ""),
 ("Divine Order (Protector), Primal Order (Warden) and feats like Heavily Armored can trigger the armor warning; that check is a prompt to confirm, not a ruling.", ""),
 ("Assumptions", "h"),
 ("Hit chance: d20 + to-hit against the target AC, with natural 1 always missing and natural 20 always hitting. A critical adds the weapon dice once more. Expected damage assumes every attack uses that weapon.", ""),
 ("Average hit points use the fixed value (half the die + 1) for levels after the first. Choose Rolled to enter your own total.", ""),
 ("Sources and credits", "h"),
 ("This work includes material from the System Reference Document 5.2.1 (SRD 5.2.1) and System Reference Document 5.1 (SRD 5.1) by Wizards of the Coast LLC, available at https://www.dndbeyond.com/srd. Both are licensed under the Creative Commons Attribution 4.0 International License (https://creativecommons.org/licenses/by/4.0/legalcode).", ""),
 ("Names and numbers beyond the SRD (some backgrounds, subclasses and feats) are listed only so you can record choices from books you own. No rules text is reproduced. Not affiliated with or endorsed by Wizards of the Coast.", ""),
]
r = 1
for t, k in lines:
    c = G[f"B{r}"]; c.value = t; c.alignment = WRAP
    c.font = f(bold=True, size=18, color=NAVY) if k == "t" else f(bold=True, size=12, color=RED) if k == "h" else f()
    if not k and len(t) > 120: G.row_dimensions[r].height = 14 * (len(t) // 115 + 1)
    r += 1

# Default all fonts on data tabs already Arial. Order: put data tabs last (already). Open on Builder.
wb.active = 1
wb.save(OUT)

layout = dict(
 inputs=dict(ruleset="C5", name="C7", method="C20", bonusMode="C30", species="C35", background="C36", armor=f'C{ROW["Armor worn"]}', shield=f'C{ROW["Shield"]}', mageArmor=f'C{ROW["Mage Armor active"]}',
   fs=f'C{ROW["Fighting style"]}', hpMethod=f'C{ROW["Hit point method"]}', hpRolled=f'C{ROW["Rolled HP total (levels after 1st)"]}', acBonus=f'C{ROW["AC adjustment"]}', hpBonus=f'C{ROW["HP adjustment"]}', initBonus=f'C{ROW["Initiative adjustment"]}',
   speedBonus=f'C{ROW["Speed adjustment"]}', saveBonus=f'C{SAVE_ADJ}', targetAC=f"C{AT+1}", raging=f"C{AT+2}"),
 classRow=14, baseRow=22, pickRow=31, featRow=43, skillRow=SK0, attackRow=AT0, checkRow=CH0, checkCount=len(CHK))
json.dump(layout, open(os.path.join(HERE, "layout.json"), "w"))
print("saved", OUT, layout)
