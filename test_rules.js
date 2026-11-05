// Known answers worked out by hand from the rulebooks, independent of the engine's logic.
const D = require("./data.json"); const { compute } = require("./engine.js");
const R24 = "2024 Player's Handbook", R24A = "2024 PHB + Forge of the Artificer", R14 = "2014 Player's Handbook", R14T = "2014 PHB + Tasha's Cauldron";
const base = (o = {}) => ({ ruleset: R24, classes: [], hpMethod: "Average", method: "Standard Array", base: { STR: 10, DEX: 10, CON: 10, INT: 10, WIS: 10, CHA: 10 }, species: "Human", background: "Sage", bonusMode: "+2 / +1", picks: ["", "", ""], feats: [], fightingStyle: "None", skills: {}, armor: "None", attacks: [], targetAC: 15, ...o });
const c = (...x) => { const a = []; for (let i = 0; i < x.length; i += 3) a.push({ cls: x[i], sub: x[i + 1], lvl: x[i + 2] }); while (a.length < 3) a.push({ cls: "", sub: "", lvl: 0 }); return a; };
const slots = (r) => r.slots.filter(Boolean).join();
let fail = 0, n = 0;
const t = (name, got, want) => { n++; if (String(got) !== String(want)) { fail++; console.log("FAIL", name, "got", got, "want", want); } };
const run = (o) => compute(base(o), D);

t("2024 Wizard 5 slots", slots(run({ classes: c("Wizard", "", 5) })), "4,3,2");
t("2024 Wizard 20 slots", slots(run({ classes: c("Wizard", "", 20) })), "4,3,3,3,3,2,2,1,1");
t("2024 Paladin 1 has slots", slots(run({ classes: c("Paladin", "", 1) })), "2");
t("2014 Paladin 1 no slots", slots(run({ ruleset: R14, species: "Human", background: "Sage", classes: c("Paladin", "", 1) })), "");
t("2014 Paladin 2", slots(run({ ruleset: R14, classes: c("Paladin", "", 2) })), "2");
t("2014 Paladin 3", slots(run({ ruleset: R14, classes: c("Paladin", "", 3) })), "3");
t("2014 Ranger 20", slots(run({ ruleset: R14, classes: c("Ranger", "", 20) })), "4,3,3,3,2");
t("2014 Pal3/Wiz1 rounds down", slots(run({ ruleset: R14, classes: c("Paladin", "", 3, "Wizard", "", 1) })), "3");
t("2024 Pal3/Wiz1 rounds up", slots(run({ classes: c("Paladin", "", 3, "Wizard", "", 1) })), "4,2");
t("2014 Artificer 1", slots(run({ ruleset: R14T, classes: c("Artificer", "", 1) })), "2");
t("Fighter 7 EK", slots(run({ classes: c("Fighter", "Eldritch Knight", 7) })), "4,2");
t("Fighter 7 no EK", slots(run({ classes: c("Fighter", "Champion", 7) })), "");
t("Rogue 3 AT", slots(run({ classes: c("Rogue", "Arcane Trickster", 3) })), "2");
t("Fighter 20 EK", slots(run({ classes: c("Fighter", "Eldritch Knight", 20) })), "4,3,3,1");
t("Warlock 5 pact", JSON.stringify(run({ classes: c("Warlock", "", 5) }).pact), '{"n":2,"level":3}');
t("Warlock 17 pact", JSON.stringify(run({ classes: c("Warlock", "", 17) }).pact), '{"n":4,"level":5}');
t("Warlock does not add slots", slots(run({ classes: c("Warlock", "", 5, "Wizard", "", 3) })), "4,2");
t("Prof bonus L1/5/9/13/17", [1, 5, 9, 13, 17].map((l) => run({ classes: c("Fighter", "", l) }).pb), "2,3,4,5,6");
// Hit points: Fighter 3 CON 14: 10+6+6 + 2*3 = 28
t("HP fighter 3", run({ classes: c("Fighter", "", 3), base: { STR: 10, DEX: 10, CON: 14, INT: 10, WIS: 10, CHA: 10 } }).hp, 28);
// Dwarf 2024 level 4 Wizard CON 10: 6+4*3 + 1*4 = 22
t("HP dwarf wizard 4", run({ species: "Dwarf", classes: c("Wizard", "", 4) }).hp, 22);
// Tough feat level 5 Rogue via Farmer background: 8+5*4 + 2*5 = 38
t("HP tough rogue 5", run({ background: "Farmer", classes: c("Rogue", "", 5) }).hp, 38);
// Multiclass: Barbarian 1 / Wizard 1 starts d12: 12 + 4 = 16
t("HP barb1 wiz1", run({ classes: c("Barbarian", "", 1, "Wizard", "", 1) }).hp, 16);
t("Hit dice text", run({ classes: c("Barbarian", "", 2, "Wizard", "", 3, "Fighter", "", 1) }).hitDice, "2d12 + 1d10 + 3d6");
// AC
const dex16 = { STR: 10, DEX: 16, CON: 14, INT: 10, WIS: 16, CHA: 10 };
t("AC plate", run({ classes: c("Fighter", "", 1), armor: "Plate", base: dex16 }).ac, 18);
t("AC plate defense shield", run({ classes: c("Fighter", "", 1), armor: "Plate", fightingStyle: "Defense", shield: true, base: dex16 }).ac, 21);
t("AC half plate caps dex", run({ classes: c("Fighter", "", 1), armor: "Half Plate", base: dex16 }).ac, 17);
t("AC medium armor master", run({ classes: c("Fighter", "", 4), armor: "Half Plate", base: dex16, feats: [{ name: "Medium Armor Master", a: "STR", b: "" }] }).ac, 18);
t("AC studded", run({ classes: c("Rogue", "", 1), armor: "Studded Leather", base: dex16 }).ac, 15);
t("AC barbarian unarmored", run({ classes: c("Barbarian", "", 1), base: dex16 }).ac, 15);
t("AC barbarian shield", run({ classes: c("Barbarian", "", 1), base: dex16, shield: true }).ac, 17);
t("AC monk unarmored", run({ classes: c("Monk", "", 1), base: dex16 }).ac, 16);
t("AC monk with shield loses it", run({ classes: c("Monk", "", 1), base: dex16, shield: true }).ac, 15);
t("AC mage armor", run({ classes: c("Wizard", "", 1), base: dex16, mageArmor: true }).ac, 16);
t("AC 2024 Draconic", run({ classes: c("Sorcerer", "Draconic Sorcery", 3), base: { ...dex16, CHA: 18 } }).ac, 10 + 3 + 4);
t("AC 2014 Draconic", run({ ruleset: R14, classes: c("Sorcerer", "Draconic Bloodline", 1), base: dex16 }).ac, 16);
// Ability math
const s2024 = run({ background: "Noble", picks: ["STR", "CHA", ""], base: { STR: 15, DEX: 10, CON: 10, INT: 10, WIS: 10, CHA: 14 }, classes: c("Paladin", "", 1) });
t("2024 bg +2/+1", [s2024.ab.STR.score, s2024.ab.CHA.score], "17,15");
const hill = run({ ruleset: R14, species: "Dwarf (Hill)", classes: c("Cleric", "", 1), base: { STR: 10, DEX: 10, CON: 14, INT: 10, WIS: 15, CHA: 10 } });
t("2014 hill dwarf", [hill.ab.CON.score, hill.ab.WIS.score, hill.hp], "16,16,12");
const tash = run({ ruleset: R14T, species: "Dwarf (Hill)", picks: ["INT", "DEX", ""], classes: c("Wizard", "", 1) });
t("Tasha floats", [tash.ab.CON.score, tash.ab.WIS.score, tash.ab.INT.score, tash.ab.DEX.score], "10,10,12,11");
t("Tasha human stays", run({ ruleset: R14T, species: "Human", classes: c("Wizard", "", 1) }).ab.STR.score, 11);
const barb = run({ classes: c("Barbarian", "", 20), base: { STR: 20, DEX: 10, CON: 20, INT: 10, WIS: 10, CHA: 10 } });
t("2024 Primal Champion", [barb.ab.STR.score, barb.ab.CON.score], "24,24");
const boon = run({ classes: c("Barbarian", "", 20), base: { STR: 20, DEX: 10, CON: 10, INT: 10, WIS: 10, CHA: 10 }, feats: [{ name: "Boon of Fate", a: "STR", b: "" }, { name: "Ability Score Improvement", a: "STR", b: "STR" }] });
t("2024 Primal Champion max 25", boon.ab.STR.score, 25);
t("Cap at 20", run({ classes: c("Fighter", "", 8), base: { STR: 20, DEX: 10, CON: 10, INT: 10, WIS: 10, CHA: 10 }, feats: [{ name: "Ability Score Improvement", a: "STR", b: "STR" }] }).ab.STR.score, 20);
// Saves and skills
const bard = run({ classes: c("Bard", "", 2), base: { STR: 10, DEX: 14, CON: 10, INT: 10, WIS: 10, CHA: 10 } });
t("Jack of All Trades on skill", bard.skills.Acrobatics.total, 2 + 1);
t("JoAT on initiative", bard.init, 3);
t("Expertise doubles", run({ classes: c("Rogue", "", 1), skills: { Stealth: "E" } }).skills.Stealth.total, 4);
t("Background skill auto", run({ background: "Sage", classes: c("Wizard", "", 1) }).skills.Arcana.total, 2);
t("Elf 2014 perception", run({ ruleset: R14, species: "Elf (Wood)", classes: c("Wizard", "", 1) }).skills.Perception.total, 2);
t("Aura of Protection", run({ classes: c("Paladin", "", 6), base: { STR: 10, DEX: 10, CON: 10, INT: 10, WIS: 10, CHA: 16 } }).saves.DEX.total, 3);
t("Saves from first class only", run({ classes: c("Wizard", "", 1, "Fighter", "", 1) }).saves.STR.prof, false);
t("Diamond Soul / Disciplined Survivor", run({ classes: c("Monk", "", 14) }).saves.CHA.prof, true);
t("2024 Slippery Mind has CHA", run({ classes: c("Rogue", "", 15) }).saves.CHA.prof, true);
t("2014 Slippery Mind no CHA", run({ ruleset: R14, classes: c("Rogue", "", 15) }).saves.CHA.prof, false);
t("Passive with Observant 2014", run({ ruleset: R14, classes: c("Wizard", "", 4), feats: [{ name: "Observant", a: "WIS", b: "" }] }).passive.Perception, 16);
// Speed and initiative
t("Goliath 2024 speed", run({ species: "Goliath", classes: c("Fighter", "", 1) }).speed, 35);
t("Barb 5 fast movement", run({ classes: c("Barbarian", "", 5) }).speed, 40);
t("Barb 5 in heavy armor", run({ classes: c("Barbarian", "", 5), armor: "Plate" }).speed, 20);
t("2014 dwarf heavy ok", run({ ruleset: R14, species: "Dwarf (Hill)", classes: c("Fighter", "", 1), armor: "Plate" }).speed, 25);
t("Monk 6 movement", run({ classes: c("Monk", "", 6) }).speed, 45);
t("2024 Alert uses prof", run({ background: "Criminal", classes: c("Rogue", "", 5) }).init, 3);
t("2014 Alert +5", run({ ruleset: R14, classes: c("Rogue", "", 4), feats: [{ name: "Alert", a: "", b: "" }] }).init, 5);
t("Goliath carry doubles", run({ species: "Goliath", classes: c("Fighter", "", 1) }).carry, 300);
// Attacks. Fighter 5, STR 16, Greatsword vs AC 15: +6, 2d6+3, 60%, two attacks
const atk = run({ classes: c("Fighter", "", 5), base: { STR: 16, DEX: 10, CON: 10, INT: 10, WIS: 10, CHA: 10 }, attacks: [{ weapon: "Greatsword", ability: "Auto", prof: "Auto" }] }).attacks[0];
t("Attack hit", atk.hit, 6); t("Attack per hit", atk.perHit, 10); t("Attack chance", atk.chance.toFixed(2), "0.60");
t("Expected per round", (atk.perRound).toFixed(2), ((0.55 * 10 + 0.05 * 17) * 2).toFixed(2));
t("Finesse uses DEX", run({ classes: c("Rogue", "", 1), base: dex16, attacks: [{ weapon: "Rapier", ability: "Auto", prof: "Auto" }] }).attacks[0].hit, 5);
t("Wizard not proficient greatsword", run({ classes: c("Wizard", "", 1), attacks: [{ weapon: "Greatsword", ability: "Auto", prof: "Auto" }] }).attacks[0].proficient, false);
t("2024 Rogue scimitar proficient", run({ classes: c("Rogue", "", 1), attacks: [{ weapon: "Scimitar", ability: "Auto", prof: "Auto" }] }).attacks[0].proficient, true);
t("2014 Rogue scimitar not proficient", run({ ruleset: R14, classes: c("Rogue", "", 1), attacks: [{ weapon: "Scimitar", ability: "Auto", prof: "Auto" }] }).attacks[0].proficient, false);
t("2024 mastery shown", run({ classes: c("Fighter", "", 1), attacks: [{ weapon: "Longsword", ability: "Auto", prof: "Auto" }] }).attacks[0].mastery, "Sap");
t("2014 no mastery", run({ ruleset: R14, classes: c("Fighter", "", 1), attacks: [{ weapon: "Longsword", ability: "Auto", prof: "Auto" }] }).attacks[0].mastery, "");
t("Monk 2024 d8 at 5", run({ classes: c("Monk", "", 5), attacks: [{ weapon: "Unarmed Strike", ability: "Auto", prof: "Auto" }] }).attacks[0].die, "1d8");
t("Monk 2014 d6 at 5", run({ ruleset: R14, classes: c("Monk", "", 5), attacks: [{ weapon: "Unarmed Strike", ability: "Auto", prof: "Auto" }] }).attacks[0].die, "1d6");
t("Rage +2", run({ classes: c("Barbarian", "", 1), raging: true, attacks: [{ weapon: "Greataxe", ability: "Auto", prof: "Auto" }] }).attacks[0].flat, 2);
t("Versatile 2H", run({ classes: c("Fighter", "", 1), attacks: [{ weapon: "Longsword", ability: "Auto", prof: "Auto", twoHanded: true }] }).attacks[0].die, "1d10");
t("Dueling +2", run({ classes: c("Fighter", "", 1), fightingStyle: "Dueling", attacks: [{ weapon: "Longsword", ability: "Auto", prof: "Auto" }] }).attacks[0].flat, 2);
t("Archery +2 hit", run({ classes: c("Fighter", "", 1), fightingStyle: "Archery", attacks: [{ weapon: "Longbow", ability: "Auto", prof: "Auto" }] }).attacks[0].hit, 4);
t("Attacks 3 at Fighter 11", run({ classes: c("Fighter", "", 11) }).attacksPerAction, 3);
// Spell numbers
t("Wizard 2014 prepared", run({ ruleset: R14, classes: c("Wizard", "", 5), base: { STR: 10, DEX: 10, CON: 10, INT: 16, WIS: 10, CHA: 10 } }).casters[0].prepared, 8);
t("Wizard 2024 prepared", run({ classes: c("Wizard", "", 5) }).casters[0].prepared, 9);
t("Spell DC", run({ classes: c("Cleric", "", 1), base: { STR: 10, DEX: 10, CON: 10, INT: 10, WIS: 16, CHA: 10 } }).casters[0].dc, 13);
t("2024 Artificer available", run({ ruleset: R24A, classes: c("Artificer", "", 3) }).checks[2].ok, true);
t("Artificer blocked in base PHB", run({ classes: c("Artificer", "", 3) }).checks[2].ok, false);
// Feat slots
t("Fighter 6 ASIs", run({ species: "Dwarf", classes: c("Fighter", "", 6) }).featSlots, 2);
t("Rogue 10 ASIs", run({ species: "Dwarf", classes: c("Rogue", "", 10) }).featSlots, 3);
t("Fighter 20 ASIs", run({ species: "Dwarf", classes: c("Fighter", "", 20) }).featSlots, 7);
t("Point cost", run({ method: "Point Buy", base: { STR: 15, DEX: 15, CON: 15, INT: 8, WIS: 8, CHA: 8 }, classes: c("Fighter", "", 1) }).pointCost, 27);
t("Multiclass req fails", run({ classes: c("Wizard", "", 1, "Fighter", "", 1), base: { STR: 13, DEX: 10, CON: 10, INT: 10, WIS: 10, CHA: 10 } }).checks[10].ok, false);
console.log(n + " known-answer checks, " + fail + " failed");
