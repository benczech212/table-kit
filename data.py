"""Rules data shared by the Excel workbook and the web page. Names and numbers only, no rulebook prose."""
import json, math, os

RULESETS = [
    {"id": "2024 Player's Handbook", "era": 2024, "artificer": 0, "tasha": 0},
    {"id": "2024 PHB + Forge of the Artificer", "era": 2024, "artificer": 1, "tasha": 0},
    {"id": "2014 Player's Handbook", "era": 2014, "artificer": 0, "tasha": 0},
    {"id": "2014 PHB + Tasha's Cauldron", "era": 2014, "artificer": 1, "tasha": 1},
]
AB = ["STR", "DEX", "CON", "INT", "WIS", "CHA"]
ABILITY_NAMES = dict(zip(AB, ["Strength", "Dexterity", "Constitution", "Intelligence", "Wisdom", "Charisma"]))
SKILLS = [("Acrobatics","DEX"),("Animal Handling","WIS"),("Arcana","INT"),("Athletics","STR"),("Deception","CHA"),("History","INT"),
          ("Insight","WIS"),("Intimidation","CHA"),("Investigation","INT"),("Medicine","WIS"),("Nature","INT"),("Perception","WIS"),
          ("Performance","CHA"),("Persuasion","CHA"),("Religion","INT"),("Sleight of Hand","DEX"),("Stealth","DEX"),("Survival","WIS")]
ALL_SKILLS = [s for s, _ in SKILLS]
XP = [0,300,900,2700,6500,14000,23000,34000,48000,64000,85000,100000,120000,140000,165000,195000,225000,265000,305000,355000]
POINT_COST = {8:0,9:1,10:2,11:3,12:4,13:5,14:7,15:9}
SLOTS = [[2],[3],[4,2],[4,3],[4,3,2],[4,3,3],[4,3,3,1],[4,3,3,2],[4,3,3,3,1],[4,3,3,3,2],[4,3,3,3,2,1],[4,3,3,3,2,1],[4,3,3,3,2,1,1],
         [4,3,3,3,2,1,1],[4,3,3,3,2,1,1,1],[4,3,3,3,2,1,1,1],[4,3,3,3,2,1,1,1,1],[4,3,3,3,3,1,1,1,1],[4,3,3,3,3,2,1,1,1],[4,3,3,3,3,2,2,1,1]]
SLOTS = [r + [0]*(9-len(r)) for r in SLOTS]
PACT_N = [1,2,2,2,2,2,2,2,2,2,3,3,3,3,3,3,4,4,4,4]
PACT_L = [1,1,2,2,3,3,4,4,5,5,5,5,5,5,5,5,5,5,5,5]

def steps(pairs, fill=0):
    """[(from_level, value)...] -> 20 values"""
    out = []
    cur = fill
    p = dict(pairs)
    for l in range(1, 21):
        cur = p.get(l, cur)
        out.append(cur)
    return out

def cls(name, die, saves, caster, sab, n, sk24, sk14, req, mode, armor24, armor14=None):
    return dict(name=name, die=die, saves=saves, caster=caster, ability=sab, skillN=n, skills24=sk24, skills14=sk14 or sk24,
                req=req, reqMode=mode, armor24=armor24, armor14=armor14 or armor24)

CLASSES = [
 cls("Barbarian",12,["STR","CON"],"none","",2,["Animal Handling","Athletics","Intimidation","Nature","Perception","Survival"],None,["STR"],"ALL","LMS"),
 cls("Bard",8,["DEX","CHA"],"full","CHA",3,ALL_SKILLS,None,["CHA"],"ALL","L"),
 cls("Cleric",8,["WIS","CHA"],"full","WIS",2,["History","Insight","Medicine","Persuasion","Religion"],None,["WIS"],"ALL","LMS"),
 cls("Druid",8,["INT","WIS"],"full","WIS",2,["Animal Handling","Arcana","Insight","Medicine","Nature","Perception","Religion","Survival"],None,["WIS"],"ALL","LS"),
 cls("Fighter",10,["STR","CON"],"none","",2,["Acrobatics","Animal Handling","Athletics","History","Insight","Intimidation","Perception","Persuasion","Survival"],
     ["Acrobatics","Animal Handling","Athletics","History","Insight","Intimidation","Perception","Survival"],["STR","DEX"],"ANY","LMHS"),
 cls("Monk",8,["STR","DEX"],"none","",2,["Acrobatics","Athletics","History","Insight","Religion","Stealth"],None,["DEX","WIS"],"ALL",""),
 cls("Paladin",10,["WIS","CHA"],"half","CHA",2,["Athletics","Insight","Intimidation","Medicine","Persuasion","Religion"],None,["STR","CHA"],"ALL","LMHS"),
 cls("Ranger",10,["STR","DEX"],"half","WIS",3,["Animal Handling","Athletics","Insight","Investigation","Nature","Perception","Stealth","Survival"],None,["DEX","WIS"],"ALL","LMS"),
 cls("Rogue",8,["DEX","INT"],"none","",4,["Acrobatics","Athletics","Deception","Insight","Intimidation","Investigation","Perception","Persuasion","Sleight of Hand","Stealth"],
     ["Acrobatics","Athletics","Deception","Insight","Intimidation","Investigation","Perception","Performance","Persuasion","Sleight of Hand","Stealth"],["DEX"],"ALL","L"),
 cls("Sorcerer",6,["CON","CHA"],"full","CHA",2,["Arcana","Deception","Insight","Intimidation","Persuasion","Religion"],None,["CHA"],"ALL",""),
 cls("Warlock",8,["WIS","CHA"],"pact","CHA",2,["Arcana","Deception","History","Intimidation","Investigation","Nature","Religion"],None,["CHA"],"ALL","L"),
 cls("Wizard",6,["INT","WIS"],"full","INT",2,["Arcana","History","Insight","Investigation","Medicine","Nature","Religion"],["Arcana","History","Insight","Investigation","Medicine","Religion"],["INT"],"ALL",""),
 cls("Artificer",8,["CON","INT"],"half","INT",2,["Arcana","History","Investigation","Medicine","Nature","Perception","Sleight of Hand"],None,["INT"],"ALL","LMS"),
]

FULL24 = [4,5,6,7,9,10,11,12,14,15,16,16,17,17,18,18,19,20,21,22]
HALF24 = [2,3,4,5,6,6,7,7,9,9,10,10,11,11,12,12,14,14,15,15]
THIRD = [0,0,3,4,4,4,5,6,6,7,8,8,9,10,10,11,11,11,12,13]
SPELLCASTING = {
 (2024,"Bard"): (steps([(1,2),(4,3),(10,4)]),"table",FULL24),
 (2024,"Cleric"): (steps([(1,3),(4,4),(10,5)]),"table",FULL24),
 (2024,"Druid"): (steps([(1,2),(4,3),(10,4)]),"table",FULL24),
 (2024,"Paladin"): ([0]*20,"table",HALF24),
 (2024,"Ranger"): ([0]*20,"table",HALF24),
 (2024,"Sorcerer"): (steps([(1,4),(4,5),(10,6)]),"table",[2,4,6,7,9,10,11,12,14,15,16,16,17,17,18,18,19,20,21,22]),
 (2024,"Warlock"): (steps([(1,2),(4,3),(10,4)]),"table",[2,3,4,5,6,7,8,9,10,10,11,11,12,12,13,13,14,14,15,15]),
 (2024,"Wizard"): (steps([(1,3),(4,4),(10,5)]),"table",[4,5,6,7,9,10,11,12,14,15,16,16,17,18,19,21,22,23,24,25]),
 (2024,"Artificer"): (steps([(1,2),(10,3),(14,4)]),"table",HALF24),
 (2024,"Eldritch Knight"): (steps([(3,2),(10,3)]),"table",THIRD),
 (2024,"Arcane Trickster"): (steps([(3,3),(10,4)]),"table",THIRD),
 (2014,"Bard"): (steps([(1,2),(4,3),(10,4)]),"table",[4,5,6,7,8,9,10,11,12,14,15,15,16,18,19,19,20,22,22,22]),
 (2014,"Cleric"): (steps([(1,3),(4,4),(10,5)]),"level",[0]*20),
 (2014,"Druid"): (steps([(1,2),(4,3),(10,4)]),"level",[0]*20),
 (2014,"Paladin"): ([0]*20,"half",[0]*20),
 (2014,"Ranger"): ([0]*20,"table",[0,2,3,3,4,4,5,5,6,6,7,7,8,8,9,9,10,10,11,11]),
 (2014,"Sorcerer"): (steps([(1,4),(4,5),(10,6)]),"table",[2,3,4,5,6,7,8,9,10,11,12,12,13,13,14,14,15,15,15,15]),
 (2014,"Warlock"): (steps([(1,2),(4,3),(10,4)]),"table",[2,3,4,5,6,7,8,9,10,10,11,11,12,12,13,13,14,14,15,15]),
 (2014,"Wizard"): (steps([(1,3),(4,4),(10,5)]),"level",[0]*20),
 (2014,"Artificer"): (steps([(1,2),(10,3),(14,4)]),"half",[0]*20),
 (2014,"Eldritch Knight"): (steps([(3,2),(10,3)]),"table",THIRD),
 (2014,"Arcane Trickster"): (steps([(3,3),(10,4)]),"table",THIRD),
}

SUB24 = {
 "Barbarian":["Path of the Berserker","Path of the Wild Heart","Path of the World Tree","Path of the Zealot"],
 "Bard":["College of Dance","College of Glamour","College of Lore","College of Valor"],
 "Cleric":["Life Domain","Light Domain","Trickery Domain","War Domain"],
 "Druid":["Circle of the Land","Circle of the Moon","Circle of the Sea","Circle of the Stars"],
 "Fighter":["Battle Master","Champion","Eldritch Knight","Psi Warrior"],
 "Monk":["Warrior of Mercy","Warrior of Shadow","Warrior of the Elements","Warrior of the Open Hand"],
 "Paladin":["Oath of Devotion","Oath of Glory","Oath of the Ancients","Oath of Vengeance"],
 "Ranger":["Beast Master","Fey Wanderer","Gloom Stalker","Hunter"],
 "Rogue":["Arcane Trickster","Assassin","Soulknife","Thief"],
 "Sorcerer":["Aberrant Sorcery","Clockwork Sorcery","Draconic Sorcery","Wild Magic Sorcery"],
 "Warlock":["Archfey Patron","Celestial Patron","Fiend Patron","Great Old One Patron"],
 "Wizard":["Abjurer","Diviner","Evoker","Illusionist"],
}
ART24 = ["Alchemist","Armorer","Artillerist","Battle Smith","Cartographer","Reanimator"]
SUB14 = {
 "Barbarian":["Path of the Berserker","Path of the Totem Warrior"],
 "Bard":["College of Lore","College of Valor"],
 "Cleric":["Knowledge Domain","Life Domain","Light Domain","Nature Domain","Tempest Domain","Trickery Domain","War Domain"],
 "Druid":["Circle of the Land","Circle of the Moon"],
 "Fighter":["Battle Master","Champion","Eldritch Knight"],
 "Monk":["Way of Shadow","Way of the Four Elements","Way of the Open Hand"],
 "Paladin":["Oath of Devotion","Oath of the Ancients","Oath of Vengeance"],
 "Ranger":["Beast Master","Hunter"],
 "Rogue":["Arcane Trickster","Assassin","Thief"],
 "Sorcerer":["Draconic Bloodline","Wild Magic"],
 "Warlock":["The Archfey","The Fiend","The Great Old One"],
 "Wizard":["School of Abjuration","School of Conjuration","School of Divination","School of Enchantment","School of Evocation","School of Illusion","School of Necromancy","School of Transmutation"],
}
SUBT = {
 "Barbarian":["Path of the Beast","Path of Wild Magic"],
 "Bard":["College of Creation","College of Eloquence"],
 "Cleric":["Order Domain","Peace Domain","Twilight Domain"],
 "Druid":["Circle of Spores","Circle of Stars","Circle of Wildfire"],
 "Fighter":["Psi Warrior","Rune Knight"],
 "Monk":["Way of Mercy","Way of the Astral Self"],
 "Paladin":["Oath of Glory","Oath of the Watchers"],
 "Ranger":["Fey Wanderer","Swarmkeeper"],
 "Rogue":["Phantom","Soulknife"],
 "Sorcerer":["Aberrant Mind","Clockwork Soul"],
 "Warlock":["The Fathomless","The Genie"],
 "Wizard":["Bladesinging","Order of Scribes"],
 "Artificer":["Alchemist","Armorer","Artillerist","Battle Smith"],
}
SUB_LEVEL_14 = {"Cleric":1,"Sorcerer":1,"Warlock":1,"Druid":2,"Wizard":2}
def subrow(era, c, n, tasha=0):
    r = dict(era=era, cls=c, name=n, level=3 if era==2024 else SUB_LEVEL_14.get(c,3), third="", hpPerLevel=0, ac="", tasha=tasha)
    if n in ("Eldritch Knight","Arcane Trickster"): r["third"]=n
    if n=="Draconic Sorcery": r.update(hpPerLevel=1, ac="10+DEX+CHA")
    if n=="Draconic Bloodline": r.update(hpPerLevel=1, ac="13+DEX")
    return r
SUBCLASSES = []
for c,l in SUB24.items():
    for n in l: SUBCLASSES.append(subrow(2024,c,n))
for n in ART24: SUBCLASSES.append(subrow(2024,"Artificer",n))
for c,l in SUB14.items():
    for n in l: SUBCLASSES.append(subrow(2014,c,n))
for c,l in SUBT.items():
    for n in l: SUBCLASSES.append(subrow(2014,c,n,1))

def features(txt):
    out=[]
    for part in txt.split(";"):
        lv,names=part.strip().split(":",1)
        for n in names.split(","):
            out.append((int(lv),n.strip()))
    return out
A="Ability Score Improvement"; B="Epic Boon"; S="Subclass feature"
FEATURES_RAW = {
 (2024,"Barbarian"):f"1:Rage,Unarmored Defense,Weapon Mastery;2:Danger Sense,Reckless Attack;3:Barbarian Subclass,Primal Knowledge;4:{A};5:Extra Attack,Fast Movement;6:{S};7:Feral Instinct,Instinctive Pounce;8:{A};9:Brutal Strike;10:{S};11:Relentless Rage;12:{A};13:Improved Brutal Strike;14:{S};15:Persistent Rage;16:{A};17:Improved Brutal Strike;18:Indomitable Might;19:{B};20:Primal Champion",
 (2024,"Bard"):f"1:Bardic Inspiration,Spellcasting;2:Expertise,Jack of All Trades;3:Bard Subclass;4:{A};5:Font of Inspiration;6:{S};7:Countercharm;8:{A};9:Expertise;10:Magical Secrets;12:{A};14:{S};16:{A};18:Superior Inspiration;19:{B};20:Words of Creation",
 (2024,"Cleric"):f"1:Divine Order,Spellcasting;2:Channel Divinity;3:Cleric Subclass;4:{A};5:Sear Undead;6:{S};7:Blessed Strikes;8:{A};10:Divine Intervention;12:{A};14:Improved Blessed Strikes;16:{A};17:{S};19:{B};20:Greater Divine Intervention",
 (2024,"Druid"):f"1:Druidic,Primal Order,Spellcasting;2:Wild Companion,Wild Shape;3:Druid Subclass;4:{A};5:Wild Resurgence;6:{S};7:Elemental Fury;8:{A};10:{S};12:{A};14:{S};15:Improved Elemental Fury;16:{A};18:Beast Spells;19:{B};20:Archdruid",
 (2024,"Fighter"):f"1:Fighting Style,Second Wind,Weapon Mastery;2:Action Surge,Tactical Mind;3:Fighter Subclass;4:{A};5:Extra Attack,Tactical Shift;6:{A};7:{S};8:{A};9:Indomitable,Tactical Master;10:{S};11:Two Extra Attacks;12:{A};13:Indomitable (two uses),Studied Attacks;14:{A};15:{S};16:{A};17:Action Surge (two uses),Indomitable (three uses);18:{S};19:{B};20:Three Extra Attacks",
 (2024,"Monk"):f"1:Martial Arts,Unarmored Defense;2:Monk's Focus,Unarmored Movement,Uncanny Metabolism;3:Deflect Attacks,Monk Subclass;4:{A},Slow Fall;5:Extra Attack,Stunning Strike;6:Empowered Strikes,{S};7:Evasion;8:{A};9:Acrobatic Movement;10:Heightened Focus,Self-Restoration;11:{S};12:{A};13:Deflect Energy;14:Disciplined Survivor;15:Perfect Focus;16:{A};17:{S};18:Superior Defense;19:{B};20:Body and Mind",
 (2024,"Paladin"):f"1:Lay on Hands,Spellcasting,Weapon Mastery;2:Fighting Style,Paladin's Smite;3:Channel Divinity,Paladin Subclass;4:{A};5:Extra Attack,Faithful Steed;6:Aura of Protection;7:{S};8:{A};9:Abjure Foes;10:Aura of Courage;11:Radiant Strikes;12:{A};14:Restoring Touch;15:{S};16:{A};18:Aura Expansion;19:{B};20:{S}",
 (2024,"Ranger"):f"1:Favored Enemy,Spellcasting,Weapon Mastery;2:Deft Explorer,Fighting Style;3:Ranger Subclass;4:{A};5:Extra Attack;6:Roving;7:{S};8:{A};9:Expertise;10:Tireless;11:{S};12:{A};13:Relentless Hunter;14:Nature's Veil;15:{S};16:{A};17:Precise Hunter;18:Feral Senses;19:{B};20:Foe Slayer",
 (2024,"Rogue"):f"1:Expertise,Sneak Attack,Thieves' Cant,Weapon Mastery;2:Cunning Action;3:Rogue Subclass,Steady Aim;4:{A};5:Cunning Strike,Uncanny Dodge;6:Expertise;7:Evasion,Reliable Talent;8:{A};9:{S};10:{A};11:Improved Cunning Strike;12:{A};13:{S};14:Devious Strikes;15:Slippery Mind;16:{A};17:{S};18:Elusive;19:{B};20:Stroke of Luck",
 (2024,"Sorcerer"):f"1:Innate Sorcery,Spellcasting;2:Font of Magic,Metamagic;3:Sorcerer Subclass;4:{A};5:Sorcerous Restoration;6:{S};7:Sorcery Incarnate;8:{A};10:Metamagic;12:{A};14:{S};16:{A};17:Metamagic;18:{S};19:{B};20:Arcane Apotheosis",
 (2024,"Warlock"):f"1:Eldritch Invocations,Pact Magic;2:Magical Cunning;3:Warlock Subclass;4:{A};6:{S};8:{A};9:Contact Patron;10:{S};11:Mystic Arcanum (level 6);12:{A};13:Mystic Arcanum (level 7);14:{S};15:Mystic Arcanum (level 8);16:{A};17:Mystic Arcanum (level 9);19:{B};20:Eldritch Master",
 (2024,"Wizard"):f"1:Arcane Recovery,Ritual Adept,Spellcasting;2:Scholar;3:Wizard Subclass;4:{A};5:Memorize Spell;6:{S};8:{A};10:{S};12:{A};14:{S};16:{A};18:Spell Mastery;19:{B};20:Signature Spells",
 (2024,"Artificer"):f"1:Spellcasting,Tinker's Magic;2:Replicate Magic Item;3:Artificer Subclass;4:{A};5:{S};6:Magic Item Tinker;7:Flash of Genius;8:{A};9:{S};10:Magic Item Adept;11:Spell-Storing Item;12:{A};14:Advanced Artifice;15:{S};16:{A};18:Magic Item Master;19:{B};20:Soul of Artifice",
 (2014,"Barbarian"):f"1:Rage,Unarmored Defense;2:Danger Sense,Reckless Attack;3:Primal Path;4:{A};5:Extra Attack,Fast Movement;6:{S};7:Feral Instinct;8:{A};9:Brutal Critical (1 die);10:{S};11:Relentless Rage;12:{A};13:Brutal Critical (2 dice);14:{S};15:Persistent Rage;16:{A};17:Brutal Critical (3 dice);18:Indomitable Might;19:{A};20:Primal Champion",
 (2014,"Bard"):f"1:Bardic Inspiration,Spellcasting;2:Jack of All Trades,Song of Rest;3:Bard College,Expertise;4:{A};5:Font of Inspiration;6:Countercharm,{S};8:{A};10:Expertise,Magical Secrets;12:{A};14:Magical Secrets,{S};16:{A};18:Magical Secrets;19:{A};20:Superior Inspiration",
 (2014,"Cleric"):f"1:Divine Domain,Spellcasting;2:Channel Divinity,{S};4:{A};5:Destroy Undead;6:{S};8:{A},{S};10:Divine Intervention;12:{A};16:{A};17:{S};19:{A};20:Divine Intervention Improvement",
 (2014,"Druid"):f"1:Druidic,Spellcasting;2:Druid Circle,Wild Shape;4:{A},Wild Shape Improvement;6:{S};8:{A},Wild Shape Improvement;10:{S};12:{A};14:{S};16:{A};18:Beast Spells,Timeless Body;19:{A};20:Archdruid",
 (2014,"Fighter"):f"1:Fighting Style,Second Wind;2:Action Surge;3:Martial Archetype;4:{A};5:Extra Attack;6:{A};7:{S};8:{A};9:Indomitable;10:{S};11:Extra Attack (2);12:{A};13:Indomitable (two uses);14:{A};15:{S};16:{A};17:Action Surge (two uses),Indomitable (three uses);18:{S};19:{A};20:Extra Attack (3)",
 (2014,"Monk"):f"1:Martial Arts,Unarmored Defense;2:Ki,Unarmored Movement;3:Deflect Missiles,Monastic Tradition;4:{A},Slow Fall;5:Extra Attack,Stunning Strike;6:Ki-Empowered Strikes,{S};7:Evasion,Stillness of Mind;8:{A};9:Unarmored Movement Improvement;10:Purity of Body;11:{S};12:{A};13:Tongue of the Sun and Moon;14:Diamond Soul;15:Timeless Body;16:{A};17:{S};18:Empty Body;19:{A};20:Perfect Self",
 (2014,"Paladin"):f"1:Divine Sense,Lay on Hands;2:Divine Smite,Fighting Style,Spellcasting;3:Divine Health,Sacred Oath;4:{A};5:Extra Attack;6:Aura of Protection;7:{S};8:{A};10:Aura of Courage;11:Improved Divine Smite;12:{A};14:Cleansing Touch;15:{S};16:{A};18:Aura Improvements;19:{A};20:{S}",
 (2014,"Ranger"):f"1:Favored Enemy,Natural Explorer;2:Fighting Style,Spellcasting;3:Primeval Awareness,Ranger Archetype;4:{A};5:Extra Attack;6:Favored Enemy Improvement,Natural Explorer Improvement;7:{S};8:{A},Land's Stride;10:Hide in Plain Sight,Natural Explorer Improvement;11:{S};12:{A};14:Favored Enemy Improvement,Vanish;15:{S};16:{A};18:Feral Senses;19:{A};20:Foe Slayer",
 (2014,"Rogue"):f"1:Expertise,Sneak Attack,Thieves' Cant;2:Cunning Action;3:Roguish Archetype;4:{A};5:Uncanny Dodge;6:Expertise;7:Evasion;8:{A};9:{S};10:{A};11:Reliable Talent;12:{A};13:{S};14:Blindsense;15:Slippery Mind;16:{A};17:{S};18:Elusive;19:{A};20:Stroke of Luck",
 (2014,"Sorcerer"):f"1:Sorcerous Origin,Spellcasting;2:Font of Magic;3:Metamagic;4:{A};6:{S};8:{A};10:Metamagic;12:{A};14:{S};16:{A};17:Metamagic;18:{S};19:{A};20:Sorcerous Restoration",
 (2014,"Warlock"):f"1:Otherworldly Patron,Pact Magic;2:Eldritch Invocations;3:Pact Boon;4:{A};6:{S};8:{A};10:{S};11:Mystic Arcanum (level 6);12:{A};13:Mystic Arcanum (level 7);14:{S};15:Mystic Arcanum (level 8);16:{A};17:Mystic Arcanum (level 9);19:{A};20:Eldritch Master",
 (2014,"Wizard"):f"1:Arcane Recovery,Spellcasting;2:Arcane Tradition;4:{A};6:{S};8:{A};10:{S};12:{A};14:{S};16:{A};18:Spell Mastery;19:{A};20:Signature Spells",
 (2014,"Artificer"):f"1:Magical Tinkering,Spellcasting;2:Infuse Item;3:Artificer Specialist,The Right Tool for the Job;4:{A};5:{S};6:Tool Expertise;7:Flash of Genius;8:{A};9:{S};10:Magic Item Adept;11:Spell-Storing Item;12:{A};14:Magic Item Savant;15:{S};16:{A};18:Magic Item Master;19:{A};20:Soul of Artifice",
}
FEATURES=[]
for (era,c),t in FEATURES_RAW.items():
    for lv,n in features(t):
        FEATURES.append(dict(era=era,cls=c,level=lv,name=n))

# Resources: value per class level (0 = none yet). Tokens: CHA/WIS/INT = ability mod (min 1), CHA+1 = mod + 1 (min 1)
LV=list(range(1,21))
RAGE=steps([(1,2),(3,3),(6,4),(12,5),(17,6)])
RD=steps([(1,"+2"),(9,"+3"),(16,"+4")])
BI=steps([(1,"d6"),(5,"d8"),(10,"d10"),(15,"d12")])
MOVE=steps([(2,"+10 ft"),(6,"+15 ft"),(10,"+20 ft"),(14,"+25 ft"),(18,"+30 ft")])
RES=[]
def res(era,c,n,rest,v): RES.append(dict(era=era,cls=c,name=n,rest=rest,v=list(v)))
sneak=[f"{math.ceil(l/2)}d6" for l in LV]
pts=[0]+LV[1:]
lay=[5*l for l in LV]
arec=[math.ceil(l/2) for l in LV]
book=[6+2*(l-1) for l in LV]
res(2024,"Barbarian","Rage uses","Long (1 on Short)",RAGE); res(2024,"Barbarian","Rage damage","-",RD)
res(2024,"Barbarian","Weapon masteries","-",steps([(1,2),(4,3),(10,4)]))
res(2024,"Bard","Bardic Inspiration uses","Long (Short from L5)",["CHA"]*20); res(2024,"Bard","Bardic Inspiration die","-",BI)
res(2024,"Cleric","Channel Divinity uses","Short",steps([(2,2),(6,3),(18,4)]))
res(2024,"Druid","Wild Shape uses","Short",steps([(2,2),(6,3),(17,4)])); res(2024,"Druid","Wild Shape max CR","-",steps([(2,"1/4"),(4,"1/2"),(8,"1")])); res(2024,"Druid","Wild Shape known forms","-",steps([(2,4),(4,6),(8,8)]))
res(2024,"Fighter","Second Wind uses","Short",steps([(1,2),(4,3),(10,4)])); res(2024,"Fighter","Action Surge uses","Short",steps([(2,1),(17,2)])); res(2024,"Fighter","Indomitable uses","Long",steps([(9,1),(13,2),(17,3)])); res(2024,"Fighter","Weapon masteries","-",steps([(1,3),(4,4),(10,5),(16,6)]))
res(2024,"Monk","Focus Points","Short",pts); res(2024,"Monk","Martial Arts die","-",steps([(1,"d6"),(5,"d8"),(11,"d10"),(17,"d12")])); res(2024,"Monk","Unarmored Movement","-",MOVE)
res(2024,"Paladin","Lay on Hands pool (HP)","Long",lay); res(2024,"Paladin","Channel Divinity uses","Short",steps([(3,2),(11,3)])); res(2024,"Paladin","Weapon masteries","-",[2]*20)
res(2024,"Ranger","Favored Enemy (free Hunter's Mark)","Long",steps([(1,2),(5,3),(9,4),(13,5),(17,6)])); res(2024,"Ranger","Weapon masteries","-",[2]*20)
res(2024,"Rogue","Sneak Attack","-",sneak); res(2024,"Rogue","Weapon masteries","-",[2]*20)
res(2024,"Sorcerer","Sorcery Points","Long",pts); res(2024,"Sorcerer","Metamagic options","-",steps([(2,2),(10,4),(17,6)])); res(2024,"Sorcerer","Innate Sorcery uses","Long",[2]*20)
res(2024,"Warlock","Eldritch Invocations","-",[1,3,3,3,5,5,6,6,7,7,7,8,8,8,9,9,9,10,10,10])
res(2024,"Wizard","Arcane Recovery (slot levels)","Long",arec); res(2024,"Wizard","Spellbook minimum (spells)","-",book)
res(2024,"Artificer","Plans known","-",steps([(2,4),(6,5),(10,6),(14,7),(18,8)])); res(2024,"Artificer","Magic items at once","-",steps([(2,2),(6,3),(10,4),(14,5),(18,6)])); res(2024,"Artificer","Flash of Genius uses","Long",[0]*6+["INT"]*14)
res(2014,"Barbarian","Rage uses","Long",RAGE[:19]+["Unlimited"]); res(2014,"Barbarian","Rage damage","-",RD); res(2014,"Barbarian","Brutal Critical dice","-",steps([(9,1),(13,2),(17,3)]))
res(2014,"Bard","Bardic Inspiration uses","Long (Short from L5)",["CHA"]*20); res(2014,"Bard","Bardic Inspiration die","-",BI); res(2014,"Bard","Song of Rest die","-",steps([(2,"d6"),(9,"d8"),(13,"d10"),(17,"d12")]))
res(2014,"Cleric","Channel Divinity uses","Short",steps([(2,1),(6,2),(18,3)])); res(2014,"Cleric","Destroy Undead CR","-",steps([(5,"1/2"),(8,"1"),(11,"2"),(14,"3"),(17,"4")]))
res(2014,"Druid","Wild Shape uses","Short",steps([(2,2)])); res(2014,"Druid","Wild Shape max CR","-",steps([(2,"1/4"),(4,"1/2"),(8,"1")]))
res(2014,"Fighter","Second Wind uses","Short",[1]*20); res(2014,"Fighter","Action Surge uses","Short",steps([(2,1),(17,2)])); res(2014,"Fighter","Indomitable uses","Long",steps([(9,1),(13,2),(17,3)]))
res(2014,"Monk","Ki Points","Short",pts); res(2014,"Monk","Martial Arts die","-",steps([(1,"d4"),(5,"d6"),(11,"d8"),(17,"d10")])); res(2014,"Monk","Unarmored Movement","-",MOVE)
res(2014,"Paladin","Lay on Hands pool (HP)","Long",lay); res(2014,"Paladin","Divine Sense uses","Long",["CHA+1"]*20); res(2014,"Paladin","Channel Divinity uses","Short",steps([(3,1)]))
res(2014,"Ranger","Favored enemy types","-",steps([(1,1),(6,2),(14,3)])); res(2014,"Ranger","Favored terrains","-",steps([(1,1),(6,2),(10,3)]))
res(2014,"Rogue","Sneak Attack","-",sneak)
res(2014,"Sorcerer","Sorcery Points","Long",pts); res(2014,"Sorcerer","Metamagic options","-",steps([(3,2),(10,3),(17,4)]))
res(2014,"Warlock","Eldritch Invocations","-",[0,2,2,2,3,3,4,4,5,5,5,6,6,6,7,7,7,8,8,8])
res(2014,"Wizard","Arcane Recovery (slot levels)","Long",arec); res(2014,"Wizard","Spellbook minimum (spells)","-",book)
res(2014,"Artificer","Infusions known","-",steps([(2,4),(6,6),(10,8),(14,10),(18,12)])); res(2014,"Artificer","Infused items","-",steps([(2,2),(6,3),(10,4),(14,5),(18,6)])); res(2014,"Artificer","Flash of Genius uses","Long",[0]*6+["INT"]*14)

def sp(era,name,size,speed,dv,resist,traits,fixed=None,picks=(0,0,0),tasha=(0,0,0),hp=0,carry=1,skill="",extra=0,feat=0,noPick=""):
    f=dict(zip(AB,[0]*6)); f.update(fixed or {})
    return dict(era=era,name=name,size=size,speed=speed,dv=dv,resist=resist,traits=traits,fixed=f,picks=list(picks),tasha=list(tasha),hp=hp,carry=carry,skill=skill,extra=extra,feat=feat,noPick=noPick)
SPECIES=[
 sp(2024,"Aasimar","Medium or Small",30,60,"Necrotic, Radiant","Celestial Revelation; Healing Hands; Light Bearer"),
 sp(2024,"Dragonborn","Medium",30,60,"By draconic ancestry","Breath Weapon; Draconic Flight (L5)"),
 sp(2024,"Dwarf","Medium",30,120,"Poison","Dwarven Resilience; Dwarven Toughness; Stonecunning",hp=1),
 sp(2024,"Elf (Drow)","Medium",30,120,"","Fey Ancestry; Keen Senses; Trance; Drow lineage spells",extra=1),
 sp(2024,"Elf (High)","Medium",30,60,"","Fey Ancestry; Keen Senses; Trance; High Elf lineage spells",extra=1),
 sp(2024,"Elf (Wood)","Medium",35,60,"","Fey Ancestry; Keen Senses; Trance; Wood Elf lineage spells",extra=1),
 sp(2024,"Gnome (Forest)","Small",30,60,"","Gnomish Cunning; Forest Gnome lineage"),
 sp(2024,"Gnome (Rock)","Small",30,60,"","Gnomish Cunning; Rock Gnome lineage"),
 sp(2024,"Goliath","Medium",35,0,"","Giant Ancestry; Large Form (L5); Powerful Build",carry=2),
 sp(2024,"Halfling","Small",30,0,"","Brave; Halfling Nimbleness; Luck; Naturally Stealthy"),
 sp(2024,"Human","Medium or Small",30,0,"","Resourceful; Skillful; Versatile (bonus Origin feat)",extra=1,feat=1),
 sp(2024,"Orc","Medium",30,120,"","Adrenaline Rush; Relentless Endurance"),
 sp(2024,"Tiefling (Abyssal)","Medium or Small",30,60,"Poison","Otherworldly Presence; Abyssal legacy spells"),
 sp(2024,"Tiefling (Chthonic)","Medium or Small",30,60,"Necrotic","Otherworldly Presence; Chthonic legacy spells"),
 sp(2024,"Tiefling (Infernal)","Medium or Small",30,60,"Fire","Otherworldly Presence; Infernal legacy spells"),
 sp(2014,"Dragonborn","Medium",30,0,"By draconic ancestry","Breath Weapon; Draconic Ancestry",fixed=dict(STR=2,CHA=1),tasha=(2,1,0)),
 sp(2014,"Dwarf (Hill)","Medium",25,60,"Poison","Dwarven Resilience; Dwarven Toughness; Stonecunning",fixed=dict(CON=2,WIS=1),tasha=(2,1,0),hp=1),
 sp(2014,"Dwarf (Mountain)","Medium",25,60,"Poison","Dwarven Resilience; Dwarven Armor Training; Stonecunning",fixed=dict(CON=2,STR=2),tasha=(2,2,0)),
 sp(2014,"Elf (Drow)","Medium",30,120,"","Fey Ancestry; Keen Senses; Trance; Drow Magic; Sunlight Sensitivity",fixed=dict(DEX=2,CHA=1),tasha=(2,1,0),skill="Perception"),
 sp(2014,"Elf (High)","Medium",30,60,"","Fey Ancestry; Keen Senses; Trance; Cantrip; Extra Language",fixed=dict(DEX=2,INT=1),tasha=(2,1,0),skill="Perception"),
 sp(2014,"Elf (Wood)","Medium",35,60,"","Fey Ancestry; Keen Senses; Trance; Mask of the Wild",fixed=dict(DEX=2,WIS=1),tasha=(2,1,0),skill="Perception"),
 sp(2014,"Gnome (Forest)","Small",25,60,"","Gnome Cunning; Natural Illusionist; Speak with Small Beasts",fixed=dict(INT=2,DEX=1),tasha=(2,1,0)),
 sp(2014,"Gnome (Rock)","Small",25,60,"","Gnome Cunning; Artificer's Lore; Tinker",fixed=dict(INT=2,CON=1),tasha=(2,1,0)),
 sp(2014,"Half-Elf","Medium",30,60,"","Fey Ancestry; Skill Versatility",fixed=dict(CHA=2),picks=(1,1,0),tasha=(2,1,1),extra=2,noPick="CHA"),
 sp(2014,"Half-Orc","Medium",30,60,"","Menacing; Relentless Endurance; Savage Attacks",fixed=dict(STR=2,CON=1),tasha=(2,1,0),skill="Intimidation"),
 sp(2014,"Halfling (Lightfoot)","Small",25,0,"","Lucky; Brave; Halfling Nimbleness; Naturally Stealthy",fixed=dict(DEX=2,CHA=1),tasha=(2,1,0)),
 sp(2014,"Halfling (Stout)","Small",25,0,"Poison","Lucky; Brave; Halfling Nimbleness; Stout Resilience",fixed=dict(DEX=2,CON=1),tasha=(2,1,0)),
 sp(2014,"Human","Medium",30,0,"","Extra Language",fixed=dict(STR=1,DEX=1,CON=1,INT=1,WIS=1,CHA=1)),
 sp(2014,"Human (Variant)","Medium",30,0,"","Skill; Feat",picks=(1,1,0),tasha=(1,1,0),extra=1,feat=1),
 sp(2014,"Tiefling","Medium",30,60,"Fire","Infernal Legacy",fixed=dict(CHA=2,INT=1),tasha=(2,1,0)),
]

def bg(era,n,ab,feat,sk,tool): return dict(era=era,name=n,abilities=ab,feat=feat,skills=sk,tool=tool)
BACKGROUNDS=[
 bg(2024,"Acolyte",["INT","WIS","CHA"],"Magic Initiate (Cleric)",["Insight","Religion"],"Calligrapher's Supplies"),
 bg(2024,"Artisan",["STR","DEX","INT"],"Crafter",["Investigation","Persuasion"],"One set of Artisan's Tools"),
 bg(2024,"Charlatan",["DEX","CON","CHA"],"Skilled",["Deception","Sleight of Hand"],"Forgery Kit"),
 bg(2024,"Criminal",["DEX","CON","INT"],"Alert",["Sleight of Hand","Stealth"],"Thieves' Tools"),
 bg(2024,"Entertainer",["STR","DEX","CHA"],"Musician",["Acrobatics","Performance"],"One Musical Instrument"),
 bg(2024,"Farmer",["STR","CON","WIS"],"Tough",["Animal Handling","Nature"],"Carpenter's Tools"),
 bg(2024,"Guard",["STR","INT","WIS"],"Alert",["Athletics","Perception"],"One Gaming Set"),
 bg(2024,"Guide",["DEX","CON","WIS"],"Magic Initiate (Druid)",["Stealth","Survival"],"Cartographer's Tools"),
 bg(2024,"Hermit",["CON","WIS","CHA"],"Healer",["Medicine","Religion"],"Herbalism Kit"),
 bg(2024,"Merchant",["CON","INT","CHA"],"Lucky",["Animal Handling","Persuasion"],"Navigator's Tools"),
 bg(2024,"Noble",["STR","INT","CHA"],"Skilled",["History","Persuasion"],"One Gaming Set"),
 bg(2024,"Sage",["CON","INT","WIS"],"Magic Initiate (Wizard)",["Arcana","History"],"Calligrapher's Supplies"),
 bg(2024,"Sailor",["STR","DEX","WIS"],"Tavern Brawler",["Acrobatics","Perception"],"Navigator's Tools"),
 bg(2024,"Scribe",["DEX","INT","WIS"],"Skilled",["Investigation","Perception"],"Calligrapher's Supplies"),
 bg(2024,"Soldier",["STR","DEX","CON"],"Savage Attacker",["Athletics","Intimidation"],"One Gaming Set"),
 bg(2024,"Wayfarer",["DEX","WIS","CHA"],"Lucky",["Insight","Stealth"],"Thieves' Tools"),
 bg(2014,"Acolyte",[],"",["Insight","Religion"],"Two languages"),
 bg(2014,"Charlatan",[],"",["Deception","Sleight of Hand"],"Disguise Kit, Forgery Kit"),
 bg(2014,"Criminal",[],"",["Deception","Stealth"],"One Gaming Set, Thieves' Tools"),
 bg(2014,"Entertainer",[],"",["Acrobatics","Performance"],"Disguise Kit, one Musical Instrument"),
 bg(2014,"Folk Hero",[],"",["Animal Handling","Survival"],"One set of Artisan's Tools, Vehicles (land)"),
 bg(2014,"Guild Artisan",[],"",["Insight","Persuasion"],"One set of Artisan's Tools, one language"),
 bg(2014,"Hermit",[],"",["Medicine","Religion"],"Herbalism Kit, one language"),
 bg(2014,"Noble",[],"",["History","Persuasion"],"One Gaming Set, one language"),
 bg(2014,"Outlander",[],"",["Athletics","Survival"],"One Musical Instrument, one language"),
 bg(2014,"Sage",[],"",["Arcana","History"],"Two languages"),
 bg(2014,"Sailor",[],"",["Athletics","Perception"],"Navigator's Tools, Vehicles (water)"),
 bg(2014,"Soldier",[],"",["Athletics","Intimidation"],"One Gaming Set, Vehicles (land)"),
 bg(2014,"Urchin",[],"",["Sleight of Hand","Stealth"],"Disguise Kit, Thieves' Tools"),
]

PHYS="STR DEX"; MENT="INT WIS CHA"; ANY="STR DEX CON INT WIS CHA"
def ft(era,name,cat,pts=0,opts="",tasha=0,**fx):
    d=dict(era=era,name=name,cat=cat,points=pts,opts=opts,tasha=tasha,hpPerLevel=0,hpFlat=0,speed=0,initPB=0,initFlat=0,passive=0,mam=0,skills=0,expert=0,boon=0)
    d.update(fx); return d
FEATS=[ft(2024,"Ability Score Improvement","ASI",2,ANY),
 ft(2024,"Alert","Origin",initPB=1),ft(2024,"Crafter","Origin"),ft(2024,"Healer","Origin"),ft(2024,"Lucky","Origin"),
 ft(2024,"Magic Initiate (Cleric)","Origin"),ft(2024,"Magic Initiate (Druid)","Origin"),ft(2024,"Magic Initiate (Wizard)","Origin"),
 ft(2024,"Musician","Origin"),ft(2024,"Savage Attacker","Origin"),ft(2024,"Skilled","Origin",skills=3),ft(2024,"Tavern Brawler","Origin"),ft(2024,"Tough","Origin",hpPerLevel=2)]
for n,o,fx in [("Actor","CHA",{}),("Athlete",PHYS,{}),("Charger",PHYS,{}),("Chef","CON WIS",{}),("Crossbow Expert","DEX",{}),("Crusher","STR CON",{}),
 ("Defensive Duelist","DEX",{}),("Dual Wielder",PHYS,{}),("Durable","CON",{}),("Elemental Adept",MENT,{}),("Fey Touched",MENT,{}),("Grappler",PHYS,{}),
 ("Great Weapon Master","STR",{}),("Heavily Armored","STR CON",{}),("Heavy Armor Master","STR CON",{}),("Inspiring Leader","WIS CHA",{}),("Keen Mind","INT",{}),
 ("Lightly Armored",PHYS,{}),("Mage Slayer",PHYS,{}),("Martial Weapon Training",PHYS,{}),("Medium Armor Master",PHYS,{"mam":1}),("Moderately Armored",PHYS,{}),
 ("Mounted Combatant","STR DEX WIS",{}),("Observant","INT WIS",{}),("Piercer",PHYS,{}),("Poisoner","DEX INT",{}),("Polearm Master",PHYS,{}),("Resilient",ANY,{}),
 ("Ritual Caster",MENT,{}),("Sentinel",PHYS,{}),("Shadow Touched",MENT,{}),("Sharpshooter","DEX",{}),("Shield Master","STR",{}),("Skill Expert",ANY,{"skills":1,"expert":1}),
 ("Skulker","DEX",{}),("Slasher",PHYS,{}),("Speedy","DEX CON",{"speed":10}),("Spell Sniper",MENT,{}),("Telekinetic",MENT,{}),("Telepathic",MENT,{}),("War Caster",MENT,{}),("Weapon Master",PHYS,{})]:
    FEATS.append(ft(2024,n,"General",1,o,**fx))
for n,fx in [("Combat Prowess",{}),("Dimensional Travel",{}),("Energy Resistance",{}),("Fate",{}),("Fortitude",{"hpFlat":40}),("Irresistible Offense",{}),("Recovery",{}),("Skill",{}),("Speed",{"speed":30}),("Spell Recall",{}),("the Night Spirit",{}),("Truesight",{})]:
    FEATS.append(ft(2024,"Boon of "+n,"Epic Boon",1,ANY,boon=1,**fx))
FEATS.append(ft(2014,"Ability Score Improvement","ASI",2,ANY))
for n,p,o,fx in [("Actor",1,"CHA",{}),("Alert",0,"",{"initFlat":5}),("Athlete",1,PHYS,{}),("Charger",0,"",{}),("Crossbow Expert",0,"",{}),("Defensive Duelist",0,"",{}),("Dual Wielder",0,"",{}),
 ("Dungeon Delver",0,"",{}),("Durable",1,"CON",{}),("Elemental Adept",0,"",{}),("Grappler",0,"",{}),("Great Weapon Master",0,"",{}),("Healer",0,"",{}),("Heavily Armored",1,"STR",{}),
 ("Heavy Armor Master",1,"STR",{}),("Inspiring Leader",0,"",{}),("Keen Mind",1,"INT",{}),("Lightly Armored",1,PHYS,{}),("Linguist",1,"INT",{}),("Lucky",0,"",{}),("Mage Slayer",0,"",{}),
 ("Magic Initiate",0,"",{}),("Martial Adept",0,"",{}),("Medium Armor Master",0,"",{"mam":1}),("Mobile",0,"",{"speed":10}),("Moderately Armored",1,PHYS,{}),("Mounted Combatant",0,"",{}),
 ("Observant",1,"INT WIS",{"passive":5}),("Polearm Master",0,"",{}),("Resilient",1,ANY,{}),("Ritual Caster",0,"",{}),("Savage Attacker",0,"",{}),("Sentinel",0,"",{}),("Sharpshooter",0,"",{}),
 ("Shield Master",0,"",{}),("Skilled",0,"",{"skills":3}),("Skulker",0,"",{}),("Spell Sniper",0,"",{}),("Tavern Brawler",1,"STR CON",{}),("Tough",0,"",{"hpPerLevel":2}),("War Caster",0,"",{}),("Weapon Master",1,PHYS,{})]:
    FEATS.append(ft(2014,n,"Feat",p,o,**fx))
for n,p,o,fx in [("Artificer Initiate",0,"",{}),("Chef",1,"CON WIS",{}),("Crusher",1,"STR CON",{}),("Eldritch Adept",0,"",{}),("Fey Touched",1,MENT,{}),("Fighting Initiate",0,"",{}),("Gunner",1,"DEX",{}),
 ("Metamagic Adept",0,"",{}),("Piercer",1,PHYS,{}),("Poisoner",0,"",{}),("Shadow Touched",1,MENT,{}),("Skill Expert",1,ANY,{"skills":1,"expert":1}),("Slasher",1,PHYS,{}),("Telekinetic",1,MENT,{}),("Telepathic",1,MENT,{})]:
    FEATS.append(ft(2014,n,"Feat (Tasha's)",p,o,tasha=1,**fx))

FIGHTING_STYLES=["None","Archery","Blind Fighting","Defense","Dueling","Great Weapon Fighting","Interception","Protection","Thrown Weapon Fighting","Two-Weapon Fighting","Unarmed Fighting"]

ARMOR=[dict(name=n,kind=k,base=b,dexCap=c,str=s,stealth=st,lb=w,gp=g) for n,k,b,c,s,st,w,g in [
 ("None","None",10,99,0,0,0,0),
 ("Padded","Light",11,99,0,1,8,5),("Leather","Light",11,99,0,0,10,10),("Studded Leather","Light",12,99,0,0,13,45),
 ("Hide","Medium",12,2,0,0,12,10),("Chain Shirt","Medium",13,2,0,0,20,50),("Scale Mail","Medium",14,2,0,1,45,50),("Breastplate","Medium",14,2,0,0,20,400),("Half Plate","Medium",15,2,0,1,40,750),
 ("Ring Mail","Heavy",14,0,0,1,40,30),("Chain Mail","Heavy",16,0,13,1,55,75),("Splint","Heavy",17,0,15,1,60,200),("Plate","Heavy",18,0,15,1,65,1500)]]

DIE={"1":[],"d4":[4],"d6":[6],"d8":[8],"d10":[10],"d12":[12],"2d6":[6,6]}
def avg(d): return sum((x+1)/2 for x in DIE[d]) if DIE[d] else 1
def gwf24(d):
    if not DIE[d]: return 1
    return sum(sum(max(v,3) for v in range(1,x+1))/x for x in DIE[d])
def gwf14(d):
    if not DIE[d]: return 1
    t=0
    for x in DIE[d]:
        m=(x+1)/2
        t+=sum((m if v<=2 else v) for v in range(1,x+1))/x
    return t
def w(name,cat,rng,dmg,typ,props,mastery,ver="",fin=0,thr=0,two=0,light=0,heavy=0,d14=None,v14=None,only24=0):
    return dict(name=name,cat=cat,range=rng,dmg24=dmg,ver24=ver,dmg14=d14 or dmg,ver14=ver if v14 is None else v14,type=typ,props=props,mastery=mastery,finesse=fin,thrown=thr,twoH=two,light=light,heavy=heavy,only24=only24)
WEAPONS=[
 w("Unarmed Strike","Simple","Melee","1","B","Always proficient",""),
 w("Club","Simple","Melee","d4","B","Light","Slow",light=1),
 w("Dagger","Simple","Melee","d4","P","Finesse, Light, Thrown 20/60","Nick",fin=1,thr=1,light=1),
 w("Greatclub","Simple","Melee","d8","B","Two-Handed","Push",two=1),
 w("Handaxe","Simple","Melee","d6","S","Light, Thrown 20/60","Vex",thr=1,light=1),
 w("Javelin","Simple","Melee","d6","P","Thrown 30/120","Slow",thr=1),
 w("Light Hammer","Simple","Melee","d4","B","Light, Thrown 20/60","Nick",thr=1,light=1),
 w("Mace","Simple","Melee","d6","B","","Sap"),
 w("Quarterstaff","Simple","Melee","d6","B","Versatile","Topple",ver="d8"),
 w("Sickle","Simple","Melee","d4","S","Light","Nick",light=1),
 w("Spear","Simple","Melee","d6","P","Thrown 20/60, Versatile","Sap",ver="d8",thr=1),
 w("Dart","Simple","Ranged","d4","P","Finesse, Thrown 20/60","Vex",fin=1,thr=1),
 w("Light Crossbow","Simple","Ranged","d8","P","Ammunition 80/320, Loading, Two-Handed","Slow",two=1),
 w("Shortbow","Simple","Ranged","d6","P","Ammunition 80/320, Two-Handed","Vex",two=1),
 w("Sling","Simple","Ranged","d4","B","Ammunition 30/120","Slow"),
 w("Battleaxe","Martial","Melee","d8","S","Versatile","Topple",ver="d10"),
 w("Flail","Martial","Melee","d8","B","","Sap"),
 w("Glaive","Martial","Melee","d10","S","Heavy, Reach, Two-Handed","Graze",two=1,heavy=1),
 w("Greataxe","Martial","Melee","d12","S","Heavy, Two-Handed","Cleave",two=1,heavy=1),
 w("Greatsword","Martial","Melee","2d6","S","Heavy, Two-Handed","Graze",two=1,heavy=1),
 w("Halberd","Martial","Melee","d10","S","Heavy, Reach, Two-Handed","Cleave",two=1,heavy=1),
 w("Lance","Martial","Melee","d10","P","Heavy, Reach, Two-Handed unless mounted","Topple",two=1,heavy=1,d14="d12"),
 w("Longsword","Martial","Melee","d8","S","Versatile","Sap",ver="d10"),
 w("Maul","Martial","Melee","2d6","B","Heavy, Two-Handed","Topple",two=1,heavy=1),
 w("Morningstar","Martial","Melee","d8","P","","Sap"),
 w("Pike","Martial","Melee","d10","P","Heavy, Reach, Two-Handed","Push",two=1,heavy=1),
 w("Rapier","Martial","Melee","d8","P","Finesse","Vex",fin=1),
 w("Scimitar","Martial","Melee","d6","S","Finesse, Light","Nick",fin=1,light=1),
 w("Shortsword","Martial","Melee","d6","P","Finesse, Light","Vex",fin=1,light=1),
 w("Trident","Martial","Melee","d8","P","Thrown 20/60, Versatile","Topple",ver="d10",thr=1,d14="d6",v14="d8"),
 w("War Pick","Martial","Melee","d8","P","Versatile","Sap",ver="d10",v14=""),
 w("Warhammer","Martial","Melee","d8","B","Versatile","Push",ver="d10"),
 w("Whip","Martial","Melee","d4","S","Finesse, Reach","Slow",fin=1),
 w("Blowgun","Martial","Ranged","1","P","Ammunition 25/100, Loading","Vex"),
 w("Hand Crossbow","Martial","Ranged","d6","P","Ammunition 30/120, Light, Loading","Vex",light=1),
 w("Heavy Crossbow","Martial","Ranged","d10","P","Ammunition 100/400, Heavy, Loading, Two-Handed","Push",two=1,heavy=1),
 w("Longbow","Martial","Ranged","d8","P","Ammunition 150/600, Heavy, Two-Handed","Slow",two=1,heavy=1),
 w("Musket","Martial","Ranged","d12","P","Ammunition 40/120, Loading, Two-Handed (2024 only)","Slow",two=1,only24=1),
 w("Pistol","Martial","Ranged","d10","P","Ammunition 30/90, Loading (2024 only)","Vex",only24=1),
]
EXTRA14={"Bard":["Hand Crossbow","Longsword","Rapier","Shortsword"],"Rogue":["Hand Crossbow","Longsword","Rapier","Shortsword"],"Monk":["Shortsword"]}
LIMIT14={"Sorcerer":["Dagger","Dart","Sling","Quarterstaff","Light Crossbow"],"Wizard":["Dagger","Dart","Sling","Quarterstaff","Light Crossbow"],
         "Druid":["Club","Dagger","Dart","Javelin","Mace","Quarterstaff","Scimitar","Sickle","Sling","Spear"]}
MARTIAL=("Barbarian","Fighter","Paladin","Ranger")
def fmt(d):
    if not d: return ""
    return ("1"+d) if d.startswith("d") else d
for x in WEAPONS:
    n=x["name"]; simple=x["cat"]=="Simple"; una=n=="Unarmed Strike"
    p24=[];p14=[]
    for c in [k["name"] for k in CLASSES]:
        ok24=una or simple or c in MARTIAL or (c=="Rogue" and (x["finesse"] or x["light"])) or (c=="Monk" and x["light"])
        if c in LIMIT14: ok14=una or n in LIMIT14[c]
        else: ok14=una or simple or c in MARTIAL or n in EXTRA14.get(c,[])
        if ok24:p24.append(c)
        if ok14:p14.append(c)
    x["prof24"]="|".join(p24); x["prof14"]="|".join(p14)
    x["monk24"]=int(una or (x["range"]=="Melee" and (simple or x["light"])))
    x["monk14"]=int(una or n=="Shortsword" or (x["range"]=="Melee" and simple and not x["twoH"] and not x["heavy"]))
    for e in ("24","14"):
        d=x["dmg"+e]; v=x["ver"+e]
        x["txt"+e]=fmt(d); x["vtxt"+e]=fmt(v) if v else fmt(d)
        x["avg"+e]=avg(d); x["vavg"+e]=avg(v) if v else avg(d)
        g=gwf24 if e=="24" else gwf14
        x["gwf"+e]=round(g(d),4); x["vgwf"+e]=round(g(v),4) if v else round(g(d),4)
MA_AVG={"d4":2.5,"d6":3.5,"d8":4.5,"d10":5.5,"d12":6.5}

spellrows=[]
for (era,n),(cn,mode,pr) in SPELLCASTING.items():
    spellrows.append(dict(era=era,name=n,cantrips=cn,mode=mode,prepared=pr))

DATA=dict(rulesets=RULESETS,abilities=AB,abilityNames=ABILITY_NAMES,skills=[dict(name=a,ability=b) for a,b in SKILLS],xp=XP,pointCost={str(k):v for k,v in POINT_COST.items()},
 slots=SLOTS,pactN=PACT_N,pactL=PACT_L,classes=CLASSES,spellcasting=spellrows,subclasses=SUBCLASSES,features=FEATURES,resources=RES,species=SPECIES,backgrounds=BACKGROUNDS,
 feats=FEATS,fightingStyles=FIGHTING_STYLES,armor=ARMOR,weapons=WEAPONS,maAvg=MA_AVG)

if __name__=="__main__":
    here=os.path.dirname(os.path.abspath(__file__))
    with open(os.path.join(here,"data.json"),"w",encoding="utf-8") as f:
        json.dump(DATA,f,ensure_ascii=False,indent=0)
    print({k:len(v) for k,v in DATA.items() if hasattr(v,"__len__")})
