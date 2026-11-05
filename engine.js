// Character rules engine. Pure function: compute(state, DATA) -> sheet. The Excel workbook implements the same rules in formulas.
(function (root) {
  const AB = ["STR", "DEX", "CON", "INT", "WIS", "CHA"];
  const mod = (s) => Math.floor((s - 10) / 2);
  const by = (arr, f) => arr.find(f);

  function lists(ruleset, D) {
    const rs = D.rulesets.find((r) => r.id === ruleset) || D.rulesets[0];
    return {
      rs,
      classes: D.classes.filter((c) => c.name !== "Artificer" || rs.artificer).map((c) => c.name),
      species: D.species.filter((s) => s.era === rs.era).map((s) => s.name),
      backgrounds: D.backgrounds.filter((b) => b.era === rs.era).map((b) => b.name),
      feats: D.feats.filter((f) => f.era === rs.era && (!f.tasha || rs.tasha)).map((f) => f.name),
      subclasses: (cls) => D.subclasses.filter((s) => s.era === rs.era && s.cls === cls && (!s.tasha || rs.tasha)).map((s) => s.name),
    };
  }

  function compute(st, D) {
    const L = lists(st.ruleset, D);
    const rs = L.rs, era = rs.era, e = String(era).slice(2);
    const rows = st.classes.map((c) => {
      const lvl = c.cls && +c.lvl >= 1 ? +c.lvl : 0;
      const def = lvl ? by(D.classes, (k) => k.name === c.cls) : null;
      if (!def) return { cls: "", sub: "", lvl: 0 };
      const sub = by(D.subclasses, (s) => s.era === era && s.cls === c.cls && s.name === c.sub);
      return { cls: c.cls, sub: c.sub || "", lvl, def, subDef: sub, unlocked: !!sub && lvl >= sub.level };
    });
    const active = rows.filter((r) => r.lvl);
    const level = active.reduce((a, r) => a + r.lvl, 0);
    const pb = level ? Math.ceil(level / 4) + 1 : 2;
    const lv = (name) => active.filter((r) => r.cls === name).reduce((a, r) => a + r.lvl, 0);
    const first = active[0];

    const species = by(D.species, (s) => s.era === era && s.name === st.species) || null;
    const bg = by(D.backgrounds, (b) => b.era === era && b.name === st.background) || null;
    const featDef = (n) => by(D.feats, (f) => f.era === era && f.name === n);
    const originFeat = era === 2024 && bg && bg.feat ? featDef(bg.feat) : null;
    const feats = st.feats.map((f) => ({ ...f, def: f.name ? featDef(f.name) : null }));
    const allFx = [...feats.map((f) => f.def), originFeat].filter(Boolean);
    const fxSum = (k) => allFx.reduce((a, f) => a + (f[k] || 0), 0);

    // abilities
    const useTasha = rs.tasha && species && species.tasha[0] > 0;
    const pickAmt = era === 2024
      ? (st.bonusMode === "+2 / +1" ? [2, 1, 0] : [1, 1, 1])
      : species ? (rs.tasha ? species.tasha : species.picks) : [0, 0, 0];
    const ab = {};
    AB.forEach((a) => {
      const fixed = species && !useTasha ? species.fixed[a] : 0;
      const picks = st.picks.reduce((t, p, i) => t + (p === a ? pickAmt[i] : 0), 0);
      const fb = feats.reduce((t, f) => {
        const pts = f.def ? f.def.points : 0;
        return t + (pts >= 1 && f.a === a ? 1 : 0) + (pts >= 2 && f.b === a ? 1 : 0);
      }, 0);
      ab[a] = { base: +st.base[a] || 0, fixed, picks, feats: fb };
    });
    const cap = fxSum("boon") > 0 ? 30 : 20;
    const barb20 = lv("Barbarian") >= 20, monk20 = era === 2024 && lv("Monk") >= 20;
    const capMax = era === 2024 ? 25 : 24;
    AB.forEach((a) => {
      const raw = Math.min(cap, ab[a].base + ab[a].fixed + ab[a].picks + ab[a].feats);
      const boost = (barb20 && (a === "STR" || a === "CON")) || (monk20 && (a === "DEX" || a === "WIS"));
      ab[a].score = boost ? Math.max(raw, Math.min(capMax, raw + 4)) : raw;
      ab[a].capstone = ab[a].score - raw;
      ab[a].mod = mod(ab[a].score);
    });
    const m = (a) => ab[a].mod;

    // saves
    const auraOn = lv("Paladin") >= 6;
    const aura = auraOn ? Math.max(1, m("CHA")) : 0;
    const saves = {};
    AB.forEach((a) => {
      const prof = (first && first.def.saves.includes(a)) ||
        feats.some((f) => f.name === "Resilient" && f.a === a) ||
        lv("Monk") >= 14 ||
        (lv("Rogue") >= 15 && (a === "WIS" || (era === 2024 && a === "CHA")));
      saves[a] = { prof: !!prof, total: m(a) + (prof ? pb : 0) + aura + (+st.saveBonus || 0) };
    });

    // skills
    const joat = lv("Bard") >= 2 ? Math.floor(pb / 2) : 0;
    const skills = {};
    let picks = 0, expertUsed = 0;
    D.skills.forEach((sk) => {
      const auto = (bg && bg.skills.includes(sk.name)) || (species && species.skill === sk.name);
      const pick = st.skills[sk.name] || "";
      const lvlp = pick === "E" ? 2 : pick === "P" || auto ? 1 : 0;
      if (pick && !auto) picks++;
      if (pick === "E") expertUsed++;
      const offList = pick && !auto;
      skills[sk.name] = { ability: sk.ability, auto: !!auto, pick, level: lvlp, offList,
        total: m(sk.ability) + (lvlp === 2 ? 2 * pb : lvlp === 1 ? pb : joat) + (+(st.skillBonus && st.skillBonus[sk.name]) || 0) };
    });
    const passiveFlat = fxSum("passive");
    const passive = {
      Perception: 10 + skills.Perception.total + passiveFlat,
      Investigation: 10 + skills.Investigation.total + passiveFlat,
      Insight: 10 + skills.Insight.total,
    };
    let skillsAllowed = first ? first.def.skillN : 0;
    active.slice(1).forEach((r) => { if (["Bard", "Ranger", "Rogue"].includes(r.cls)) skillsAllowed += 1; });
    skillsAllowed += (species ? species.extra : 0) + fxSum("skills");
    const expertAllowed =
      (era === 2024 ? (lv("Bard") >= 9 ? 4 : lv("Bard") >= 2 ? 2 : 0) : (lv("Bard") >= 10 ? 4 : lv("Bard") >= 3 ? 2 : 0)) +
      (lv("Rogue") >= 6 ? 4 : lv("Rogue") >= 1 ? 2 : 0) +
      (era === 2024 && lv("Ranger") >= 9 ? 2 : 0) + fxSum("expert");

    // hit points
    const dieAvg = (d) => Math.floor(d / 2) + 1;
    let hpDice = 0;
    if (first) {
      hpDice = first.def.die;
      if (st.hpMethod === "Rolled") hpDice += +st.hpRolled || 0;
      else active.forEach((r, i) => { hpDice += dieAvg(r.def.die) * (r.lvl - (i === 0 ? 1 : 0)); });
    }
    const subHp = active.reduce((a, r) => a + (r.unlocked ? r.subDef.hpPerLevel * r.lvl : 0), 0);
    const hp = Math.max(level, hpDice + m("CON") * level + (species ? species.hp : 0) * level + subHp + fxSum("hpPerLevel") * level + fxSum("hpFlat") + (+st.hpBonus || 0));
    const dieMap = {};
    active.forEach((r) => { dieMap[r.def.die] = (dieMap[r.def.die] || 0) + r.lvl; });
    const hitDice = Object.keys(dieMap).sort((a, b) => b - a).map((d) => dieMap[d] + "d" + d).join(" + ");

    // armor class
    const armor = by(D.armor, (a) => a.name === st.armor) || D.armor[0];
    const wearing = armor.kind !== "None";
    const dexCap = armor.kind === "Medium" ? 2 + fxSum("mam") : armor.dexCap;
    const fs = st.fightingStyle || "None";
    let baseAC;
    if (wearing) baseAC = armor.base + Math.min(dexCap, armor.kind === "Heavy" ? 0 : m("DEX"));
    else {
      baseAC = 10 + m("DEX");
      if (lv("Barbarian")) baseAC = Math.max(baseAC, 10 + m("DEX") + m("CON"));
      if (lv("Monk") && !st.shield) baseAC = Math.max(baseAC, 10 + m("DEX") + m("WIS"));
      active.forEach((r) => {
        if (r.unlocked && r.subDef.ac === "13+DEX") baseAC = Math.max(baseAC, 13 + m("DEX"));
        if (r.unlocked && r.subDef.ac === "10+DEX+CHA") baseAC = Math.max(baseAC, 10 + m("DEX") + m("CHA"));
      });
      if (st.mageArmor) baseAC = Math.max(baseAC, 13 + m("DEX"));
    }
    const ac = baseAC + (st.shield ? 2 : 0) + (wearing && fs === "Defense" ? 1 : 0) + (+st.acBonus || 0);

    // speed, initiative, carrying
    const heavy = armor.kind === "Heavy";
    const strShort = armor.str > 0 && ab.STR.score < armor.str && !(era === 2014 && species && species.name.startsWith("Dwarf"));
    let speed = species ? species.speed : 30;
    if (lv("Barbarian") >= 5 && !heavy) speed += 10;
    if (lv("Monk") >= 2 && !wearing && !st.shield) {
      const l = lv("Monk");
      speed += l >= 18 ? 30 : l >= 14 ? 25 : l >= 10 ? 20 : l >= 6 ? 15 : 10;
    }
    if (era === 2024 && lv("Ranger") >= 6 && !heavy) speed += 10;
    if (strShort) speed -= 10;
    speed += fxSum("speed") + (+st.speedBonus || 0);
    const init = m("DEX") + joat + fxSum("initPB") * pb + fxSum("initFlat") + (+st.initBonus || 0);
    const carryMult = species ? species.carry : 1;
    const carry = ab.STR.score * 15 * carryMult;

    // spellcasting
    const casters = [];
    active.forEach((r) => {
      let type = r.def.caster, ability = r.def.ability, key = r.cls;
      if (r.unlocked && r.subDef.third) { type = "third"; ability = "INT"; key = r.subDef.third; }
      if (type === "none") return;
      const row = by(D.spellcasting, (s) => s.era === era && s.name === key);
      const i = r.lvl - 1;
      let prepared = row.prepared[i];
      if (row.mode === "level") prepared = Math.max(1, r.lvl + m(ability));
      if (row.mode === "half") prepared = (r.cls === "Paladin" && r.lvl < 2) ? 0 : Math.max(1, Math.floor(r.lvl / 2) + m(ability));
      casters.push({ cls: r.cls, type, ability, lvl: r.lvl, dc: 8 + pb + m(ability), attack: pb + m(ability), cantrips: row.cantrips[i], prepared, mode: row.mode });
    });
    const slotCasters = casters.filter((c) => c.type !== "pact");
    let effLevel = 0;
    if (slotCasters.length === 1) {
      const c = slotCasters[0];
      if (c.type === "full") effLevel = c.lvl;
      if (c.type === "half") effLevel = era === 2014 && c.cls !== "Artificer" && c.lvl < 2 ? 0 : Math.ceil(c.lvl / 2);
      if (c.type === "third") effLevel = Math.ceil(c.lvl / 3);
    } else {
      slotCasters.forEach((c) => {
        if (c.type === "full") effLevel += c.lvl;
        if (c.type === "half") effLevel += era === 2024 || c.cls === "Artificer" ? Math.ceil(c.lvl / 2) : Math.floor(c.lvl / 2);
        if (c.type === "third") effLevel += Math.floor(c.lvl / 3);
      });
    }
    const slots = effLevel > 0 ? D.slots[Math.min(effLevel, 20) - 1] : [0, 0, 0, 0, 0, 0, 0, 0, 0];
    const pactLvl = lv("Warlock");
    const pact = pactLvl ? { n: D.pactN[pactLvl - 1], level: D.pactL[pactLvl - 1] } : { n: 0, level: 0 };

    // attacks
    const fighter = lv("Fighter");
    const attacksPerAction = fighter >= 20 ? 4 : fighter >= 11 ? 3 :
      fighter >= 5 || lv("Barbarian") >= 5 || lv("Monk") >= 5 || lv("Paladin") >= 5 || lv("Ranger") >= 5 ? 2 : 1;
    const rageDie = lv("Barbarian") >= 16 ? 4 : lv("Barbarian") >= 9 ? 3 : lv("Barbarian") >= 1 ? 2 : 0;
    const monkL = lv("Monk");
    const maAvg = monkL ? (era === 2024 ? (monkL >= 17 ? 6.5 : monkL >= 11 ? 5.5 : monkL >= 5 ? 4.5 : 3.5) : (monkL >= 17 ? 5.5 : monkL >= 11 ? 4.5 : monkL >= 5 ? 3.5 : 2.5)) : 0;
    const maTxt = monkL ? (era === 2024 ? (monkL >= 17 ? "1d12" : monkL >= 11 ? "1d10" : monkL >= 5 ? "1d8" : "1d6") : (monkL >= 17 ? "1d10" : monkL >= 11 ? "1d8" : monkL >= 5 ? "1d6" : "1d4")) : "";
    const target = +st.targetAC || 13;
    const attacks = st.attacks.map((a) => {
      const w = a.weapon ? by(D.weapons, (x) => x.name === a.weapon) : null;
      if (!w) return null;
      const melee = w.range === "Melee";
      const monkW = monkL > 0 && w["monk" + e] === 1;
      let abil = a.ability && a.ability !== "Auto" ? a.ability : null;
      if (!abil) {
        if (!melee) abil = "DEX";
        else if (w.finesse || monkW) abil = m("DEX") > m("STR") ? "DEX" : "STR";
        else abil = "STR";
      }
      const profList = w["prof" + e].split("|");
      const proficient = a.prof === "Yes" ? true : a.prof === "No" ? false : active.some((r) => profList.includes(r.cls));
      const hit = m(abil) + (proficient ? pb : 0) + (+a.magic || 0) + (fs === "Archery" && !melee ? 2 : 0);
      const two = !!a.twoHanded && w["ver" + e] !== "";
      const gwf = fs === "Great Weapon Fighting" && melee && (w.twoH || two);
      let die = gwf ? (two ? w["vgwf" + e] : w["gwf" + e]) : (two ? w["vavg" + e] : w["avg" + e]);
      let txt = two ? w["vtxt" + e] : w["txt" + e];
      if (monkW && !gwf && maAvg > die) { die = maAvg; txt = maTxt; }
      const duel = fs === "Dueling" && melee && !w.twoH && !two ? 2 : 0;
      const rage = st.raging && melee && abil === "STR" ? rageDie : 0;
      const flat = m(abil) + (+a.magic || 0) + (+a.dmg || 0) + duel + rage;
      const perHit = die + flat;
      const p = Math.max(0.05, Math.min(0.95, (21 + hit - target) / 20));
      const exp = (p - 0.05) * perHit + 0.05 * (perHit + die);
      return { weapon: w.name, ability: abil, proficient, hit, die: txt, flat, type: w.type, perHit, chance: p, expected: exp, perRound: exp * attacksPerAction, mastery: era === 2024 ? w.mastery : "", props: w.props };
    });

    // features and resources
    const features = [];
    D.features.forEach((f) => { if (f.era === era && lv(f.cls) >= f.level) features.push(f); });
    const resources = [];
    D.resources.forEach((r) => {
      if (r.era !== era) return;
      const l = lv(r.cls);
      if (!l) return;
      let v = r.v[l - 1];
      if (v === 0 || v === "") return;
      if (v === "CHA" || v === "INT" || v === "WIS") v = Math.max(1, m(v));
      if (v === "CHA+1") v = Math.max(1, m("CHA") + 1);
      resources.push({ cls: r.cls, name: r.name, rest: r.rest, value: v });
    });

    // advancement
    const asiCount = active.reduce((t, r) => {
      let n = [4, 8, 12, 16, 19].filter((x) => r.lvl >= x).length;
      if (r.cls === "Fighter") n += [6, 14].filter((x) => r.lvl >= x).length;
      if (r.cls === "Rogue") n += r.lvl >= 10 ? 1 : 0;
      return t + n;
    }, 0);
    const featSlots = asiCount + (species ? species.feat : 0);
    const featsUsed = feats.filter((f) => f.name).length;
    const nextXP = level >= 1 && level < 20 ? D.xp[level] : 0;
    const attune = lv("Artificer") >= 18 ? 6 : lv("Artificer") >= 14 ? 5 : lv("Artificer") >= 10 ? 4 : 3;

    // rules checks
    const checks = [];
    const add = (ok, text) => checks.push({ ok: !!ok, text });
    const baseVals = AB.map((a) => ab[a].base);
    const cost = baseVals.reduce((t, b) => t + (D.pointCost[b] === undefined ? 0 : D.pointCost[b]), 0);
    if (st.method === "Point Buy") add(baseVals.every((b) => b >= 8 && b <= 15) && cost <= 27, "Point buy: " + cost + " of 27 points, each score 8 to 15");
    else if (st.method === "Standard Array") add([15, 14, 13, 12, 10, 8].join() === [...baseVals].sort((a, b) => b - a).join(), "Standard array: scores must be 15, 14, 13, 12, 10, 8");
    else add(baseVals.every((b) => b >= 3 && b <= 18), "Rolled scores: each between 3 and 18");
    add(level >= 1 && level <= 20, "Total level is " + level + " (must be 1 to 20)");
    add(active.every((r) => L.classes.includes(r.cls)), "Classes are available in this rulebook");
    add(species && L.species.includes(species.name), "Species is available in this rulebook");
    add(bg && L.backgrounds.includes(bg.name), "Background is available in this rulebook");
    const usedPicks = st.picks.filter((p, i) => p && pickAmt[i] > 0);
    const needPicks = pickAmt.filter((x) => x > 0).length;
    let pickOk = usedPicks.length === needPicks && new Set(usedPicks).size === usedPicks.length;
    if (era === 2024 && bg) pickOk = pickOk && usedPicks.every((p) => bg.abilities.includes(p));
    if (era === 2014 && !rs.tasha && species && species.noPick) pickOk = pickOk && !usedPicks.includes(species.noPick);
    add(pickOk, era === 2024 ? "Background ability bonuses go to different allowed abilities" : "Species ability choices are complete and different");
    add(featsUsed <= featSlots, "Feats and ASIs: " + featsUsed + " used of " + featSlots + " available");
    add(feats.every((f) => {
      if (!f.def) return !f.name;
      const opt = f.def.opts.split(" ");
      return (f.def.points < 1 || (f.a && opt.includes(f.a))) && (f.def.points < 2 || (f.b && opt.includes(f.b)));
    }), "Each feat's ability increase is allowed and a feat from this rulebook");
    add(picks <= skillsAllowed, "Skills: " + picks + " chosen of " + skillsAllowed + " allowed");
    add(expertUsed <= expertAllowed, "Expertise: " + expertUsed + " used of " + expertAllowed + " allowed");
    const mcOk = active.length < 2 || active.every((r) => {
      const t = r.def.req.map((a) => ab[a].score >= 13);
      return r.def.reqMode === "ANY" ? t.some(Boolean) : t.every(Boolean);
    });
    add(mcOk, "Multiclass ability requirements (13+) are met");
    const trained = (k) => active.some((r) => r.def["armor" + e].includes(k));
    const armorK = wearing ? armor.kind[0] : "";
    add((!armorK || trained(armorK)) && (!st.shield || trained("S")), "Worn armor and shield match class training");
    add(!strShort, "Strength meets the worn armor's requirement");
    add(active.every((r) => !r.sub || (r.subDef && r.unlocked)), "Subclasses are unlocked at their class level");

    const skillList = {};
    if (first) D.skills.forEach((s) => { skillList[s.name] = first.def["skills" + e].includes(s.name); });

    return { era, rs, lists: L, level, pb, ab, saves, skills, skillList, passive, joat, aura, hp, hitDice, ac, baseAC, armor, stealthDis: !!armor.stealth, strShort, speed, init, carry, push: carry * 2,
      jumpLong: ab.STR.score, jumpHigh: Math.max(0, 3 + m("STR")), casters, effLevel, slots, pact, attacksPerAction, attacks, features, resources, featSlots, featsUsed, originFeat: originFeat ? originFeat.name : "",
      skillsAllowed, skillsUsed: picks, expertAllowed, expertUsed, pointCost: cost, nextXP, attune, checks, species, bg, resist: species ? species.resist : "", dv: species ? species.dv : 0, size: species ? species.size : "" };
  }
  root.HeroEngine = { compute, lists, AB };
  if (typeof module !== "undefined") module.exports = root.HeroEngine;
})(typeof window !== "undefined" ? window : globalThis);
