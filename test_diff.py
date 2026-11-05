"""Differential test: random characters through engine.js (node) and the workbook (real Excel via COM). Any disagreement is printed."""
import json, os, random, shutil, subprocess, sys, tempfile
import win32com.client as com
import data as D

HERE = os.path.dirname(os.path.abspath(__file__))
N = int(sys.argv[1]) if len(sys.argv) > 1 else 40
random.seed(int(sys.argv[2]) if len(sys.argv) > 2 else 7)
L = json.load(open(os.path.join(HERE, "layout.json")))

def rand_state():
    rs = random.choice(D.RULESETS); era = rs["era"]
    names = [c["name"] for c in D.CLASSES if c["name"] != "Artificer" or rs["artificer"]]
    cn = random.sample(names, random.choice([1, 1, 2, 3]))
    left = 20; classes = []
    for i, c in enumerate(cn):
        lvl = random.randint(1, max(1, left - (len(cn) - i - 1))) if len(cn) > 1 else random.randint(1, 20)
        left -= lvl
        subs = [s["name"] for s in D.SUBCLASSES if s["era"] == era and s["cls"] == c and (not s["tasha"] or rs["tasha"])]
        classes.append(dict(cls=c, sub=random.choice(subs + [""]), lvl=lvl))
    while len(classes) < 3: classes.append(dict(cls="", sub="", lvl=0))
    sp = random.choice([s for s in D.SPECIES if s["era"] == era]); bg = random.choice([b for b in D.BACKGROUNDS if b["era"] == era])
    feats_pool = [x for x in D.FEATS if x["era"] == era and (not x["tasha"] or rs["tasha"])]
    feats = []
    for i in range(10):
        if i < random.choice([0, 1, 3, 6]):
            x = random.choice(feats_pool); o = x["opts"].split() or D.AB
            feats.append(dict(name=x["name"], a=random.choice(o) if x["points"] else "", b=random.choice(o) if x["points"] > 1 else ""))
        else: feats.append(dict(name="", a="", b=""))
    method = random.choice(["Standard Array", "Point Buy", "Rolled"])
    if method == "Standard Array":
        v = [15, 14, 13, 12, 10, 8]; random.shuffle(v)
    elif method == "Point Buy": v = [random.randint(8, 15) for _ in D.AB]
    else: v = [random.randint(3, 18) for _ in D.AB]
    skills = {s: random.choice(["P", "E"]) for s in random.sample(D.ALL_SKILLS, random.randint(0, 8))}
    attacks = []
    for i in range(5):
        if i < random.randint(0, 5):
            attacks.append(dict(weapon=random.choice(D.WEAPONS)["name"], ability=random.choice(["Auto", "Auto", "STR", "DEX", "CHA"]), prof=random.choice(["Auto", "Auto", "Yes", "No"]), magic=random.choice([0, 0, 1, 3]), dmg=random.choice([0, 0, 2]), twoHanded=random.random() < .4))
        else: attacks.append(dict(weapon="", ability="", prof="", magic=0, dmg=0, twoHanded=False))
    return dict(ruleset=rs["id"], classes=classes, hpMethod=random.choice(["Average", "Average", "Rolled"]), hpRolled=random.randint(0, 60), method=method, base=dict(zip(D.AB, v)),
        species=sp["name"], background=bg["name"], bonusMode=random.choice(["+2 / +1", "+1 / +1 / +1"]), picks=[random.choice(D.AB + [""]) for _ in range(3)], feats=feats,
        fightingStyle=random.choice(D.FIGHTING_STYLES), skills=skills, armor=random.choice(D.ARMOR)["name"], shield=random.random() < .3, mageArmor=random.random() < .2,
        acBonus=random.choice([0, 0, 1]), hpBonus=random.choice([0, 0, 5]), initBonus=random.choice([0, 0, 2]), speedBonus=random.choice([0, 0, 5]), saveBonus=random.choice([0, 0, 1]),
        attacks=attacks, targetAC=random.randint(8, 22), raging=random.random() < .3)

states = [rand_state() for _ in range(N)]
tmp = tempfile.mkdtemp()
sf = os.path.join(tmp, "s.json"); rf = os.path.join(tmp, "r.json")
json.dump(states, open(sf, "w"))
js = f"""const D=require({json.dumps(os.path.join(HERE,'data.json'))});const {{compute}}=require({json.dumps(os.path.join(HERE,'engine.js'))});const fs=require('fs');
const S=JSON.parse(fs.readFileSync({json.dumps(sf)}));fs.writeFileSync({json.dumps(rf)},JSON.stringify(S.map(s=>compute(s,D))));"""
subprocess.run(["node", "-e", js], check=True)
R = json.load(open(rf))

copy = os.path.join(tmp, "t.xlsx"); shutil.copy(os.path.join(HERE, "Hero-Ledger.xlsx"), copy)
xl = com.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
wb = xl.Workbooks.Open(copy)
B = wb.Worksheets("Builder"); S = wb.Worksheets("Sheet"); C = wb.Worksheets("Calc")
I = L["inputs"]
def put(a, v): B.Range(a).Value = (v if v not in ("", None) else None)
bad = 0
def eq(ctx, name, x, e):
    global bad
    if isinstance(e, float) or isinstance(x, float):
        try:
            if abs(float(x) - float(e)) < 1e-6: return
        except (TypeError, ValueError): pass
    elif str(x if x is not None else "") == str(e if e is not None else ""): return
    bad += 1
    if bad < 60: print(f"[{ctx}] {name}: excel={x!r} engine={e!r}")
try:
    for n, (s, r) in enumerate(zip(states, R)):
        xl.Calculation = -4135
        put(I["ruleset"], s["ruleset"]); put(I["method"], s["method"]); put(I["bonusMode"], s["bonusMode"]); put(I["species"], s["species"]); put(I["background"], s["background"])
        put(I["armor"], s["armor"]); put(I["shield"], "Yes" if s["shield"] else "No"); put(I["mageArmor"], "Yes" if s["mageArmor"] else "No"); put(I["fs"], s["fightingStyle"])
        put(I["hpMethod"], s["hpMethod"]); put(I["hpRolled"], s["hpRolled"])
        for k in ("acBonus", "hpBonus", "initBonus", "speedBonus", "saveBonus", "targetAC"): put(I[k], s[k] or None)
        put(I["raging"], "Yes" if s["raging"] else "No")
        for i, c in enumerate(s["classes"]):
            rr = L["classRow"] + i
            put(f"B{rr}", c["cls"]); put(f"C{rr}", c["sub"]); put(f"D{rr}", c["lvl"] or None)
        for i, a in enumerate(D.AB): put(f'C{L["baseRow"]+i}', s["base"][a])
        for i, p in enumerate(s["picks"]): put(f'C{L["pickRow"]+i}', p)
        for i, x in enumerate(s["feats"]):
            rr = L["featRow"] + i
            put(f"B{rr}", x["name"]); put(f"C{rr}", x["a"]); put(f"D{rr}", x["b"])
        for i, sk in enumerate(D.ALL_SKILLS): put(f'E{L["skillRow"]+i}', s["skills"].get(sk))
        for i, a in enumerate(s["attacks"]):
            rr = L["attackRow"] + i
            put(f"B{rr}", a["weapon"]); put(f"C{rr}", a["ability"]); put(f"D{rr}", a["prof"]); put(f"E{rr}", a["magic"] or None); put(f"F{rr}", a["dmg"] or None); put(f"G{rr}", "Yes" if a["twoHanded"] else "No")
        xl.Calculation = -4105; xl.Calculate()
        ctx = f'#{n} {s["ruleset"][:4]} {[ (c["cls"],c["lvl"]) for c in s["classes"] if c["cls"]]}'
        eq(ctx, "level", B.Range("D17").Value, r["level"]); eq(ctx, "pb", B.Range("F17").Value, r["pb"])
        for i, a in enumerate(D.AB):
            br = L["baseRow"] + i
            eq(ctx, a + " score", B.Range(f"H{br}").Value, float(r["ab"][a]["score"])); eq(ctx, a + " mod", B.Range(f"I{br}").Value, float(r["ab"][a]["mod"])); eq(ctx, a + " save", B.Range(f"J{br}").Value, float(r["saves"][a]["total"]))
        for i, sk in enumerate(D.ALL_SKILLS): eq(ctx, sk, B.Range(f'H{L["skillRow"]+i}').Value, float(r["skills"][sk]["total"]))
        eq(ctx, "passive", B.Range(f'I{L["skillRow"]+11}').Value, float(r["passive"]["Perception"]))
        eq(ctx, "hp", C.Range("B13").Value, float(r["hp"])); eq(ctx, "hitDice", C.Range("B14").Value, r["hitDice"])
        eq(ctx, "ac", C.Range("B56").Value, float(r["ac"])); eq(ctx, "speed", C.Range("B57").Value, float(r["speed"])); eq(ctx, "init", C.Range("B58").Value, float(r["init"])); eq(ctx, "carry", C.Range("B80").Value, float(r["carry"]))
        eq(ctx, "eff", C.Range("B17").Value, float(r["effLevel"])); eq(ctx, "attacks/act", C.Range("B34").Value, float(r["attacksPerAction"]))
        for i in range(9):
            col = "JKLM"[i % 4]; rr = 24 + (i // 4) * 2
            eq(ctx, f"slot{i+1}", S.Range(f"{col}{rr}").Value, float(r["slots"][i]))
        pact = f'{r["pact"]["n"]} at L{r["pact"]["level"]}' if r["pact"]["n"] else "none"
        eq(ctx, "pact", S.Range("K28").Value, pact)
        cs = [c for c in r["casters"]]; k = 0
        for i in range(3):
            if C.Range(f"H{4+i}").Value != "none":
                if k < len(cs):
                    eq(ctx, "dc", C.Range(f"N{4+i}").Value, float(cs[k]["dc"])); eq(ctx, "cantrips", C.Range(f"L{4+i}").Value, float(cs[k]["cantrips"])); eq(ctx, "prepared", C.Range(f"M{4+i}").Value, float(cs[k]["prepared"]))
                k += 1
        eq(ctx, "casterN", k, len(cs))
        for i, a in enumerate(r["attacks"]):
            rr = L["attackRow"] + i
            if a is None:
                eq(ctx, f"atk{i} blank", B.Range(f"H{rr}").Value, ""); continue
            eq(ctx, f"atk{i} hit", B.Range(f"H{rr}").Value, float(a["hit"])); eq(ctx, f"atk{i} avg", B.Range(f"J{rr}").Value, float(a["perHit"]))
            eq(ctx, f"atk{i} chance", B.Range(f"K{rr}").Value, float(a["chance"])); eq(ctx, f"atk{i} round", B.Range(f"L{rr}").Value, float(a["perRound"]))
            dtxt = f'{a["die"]}{"+" if a["flat"]>=0 else ""}{a["flat"]} {a["type"]}'
            eq(ctx, f"atk{i} text", B.Range(f"I{rr}").Value, dtxt); eq(ctx, f"atk{i} mastery", B.Range(f"M{rr}").Value, a["mastery"])
        for i, ck in enumerate(r["checks"]):
            v = B.Range(f'O{L["checkRow"]+i}').Value
            eq(ctx, "check: " + ck["text"], v, "OK" if ck["ok"] else "FIX")
        eq(ctx, "featRun", wb.Worksheets("Features").Range(f'F{len(D.FEATURES)+1}').Value, float(len(r["features"])))
        for i in range(16):
            rr = 32 + i
            if i < len(r["resources"]):
                x = r["resources"][i]
                eq(ctx, "res name", S.Range(f"B{rr}").Value, x["name"])
                val = S.Range(f"D{rr}").Value
                if isinstance(x["value"], str): eq(ctx, "res " + x["name"], val, x["value"])
                else: eq(ctx, "res " + x["name"], val, float(x["value"]))
            else: eq(ctx, "res blank", S.Range(f"B{rr}").Value, "")
        eq(ctx, "feat slots", C.Range("B25").Value, float(r["featSlots"])); eq(ctx, "skills allowed", C.Range("B31").Value, float(r["skillsAllowed"]))
finally:
    wb.Close(False); xl.Quit()
print(f"{N} characters compared, {bad} mismatches")
