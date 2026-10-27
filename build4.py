import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter
from collections import Counter

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MAX = 1000

NOMS = [
 "Essence Rose","Essence Orange","Essence Jade","Essence Violette","Essence Jaune",
 "Essence Mauve","Essence Verte","Essence Bleue","Essence Azur",
 "Broken Brawks","Vandale","Squeezie Le Toxique","Pas compris","Zacking Buff",
 "Ravitaillement","Mur Salvateur","Tout Commence Ici","BUY BUNDLE M8 !!","OCE/PCE",
 "Dégustation d'Adversaires","Trauma Dump III","Namifique","Marteen Supreme","Agent M.E.L",
 "Gentle Mates Minny","Pause Technique","Remontada","CHAMPIONNES D'EUROPE","Pollo",
 "Worlds Lyon","Moi je suis coach","Invitation","Visa Refusé","Vanyabubu","Pitchounes",
 "Maillot Champions 25","L'oeil de Nikof","Caprice de Gotaga","Zéro Soucis",
 "Les Rois du Saloon","Trauma Dump I","Reverse Squeeze","BLBLBLBL","Moroccan Monsters",
 "Zzzzzzz","Maillot Champions 24","CHAMPIONS DU MAJOR","CHAMPIONS DU MAJOR",
 "Le Roi des rois","La fée Rosalie","La Planète","Most Valuable Planet","Ravitaillement",
 "Trauma Dump II","Target Missed","Le Maillot Maudit","Siborg","Ghosty in Finals",
 "Baby Driver","Gotaga Pose","CHAMPIONS DU MAJOR","Just Passion","Calma","Bwarks",
 "Jus de bagarre","L'immense Draculan","1% chance","99% faith","L'Arme Nucléaire",
 "Worlds Japon","Versus Fighting","Leakof","L'Héritier","Grand Respect","Pétrodollars",
 "Ralliement","Brawks Déter","Le Chevalier Rose","Instalock","Meet & Greeter",
 "Mon Cogneur","Barakataga","Mercato","La Chevalière","Cooked","En slip sur la Faille",
 "Gros Bonnet","M4","La Luxueuse I","La Luxueuse II","C'est officiel !","Clim",
 "Alerte Leak","ONE MORE","Tunnel","Tweet la Défaite","Doig'Dance","Méta Coeurs",
 "Méta F*ck","Le Loup du Game","SLAY","Monster","Joueur de l'ombre","Lofi Girl","M16",
 "Trahison++","GP Explorer III","Keynote","Hatewatch","Leakof","La Gaming House",
 "M8 Résilience","Kampaï","SBG","CHAMPIONS DU MONDE !","The French Sorcerer","La Relève",
 "Gotaga Dance","Gotaga Circus","Brawks Saga","Le Marionnettiste","La Luxueuse III",
 "Le Quatrième du Nom","Brawks Trop Déter","6W26L","GP Explorer II","Paris Gentle Mates",
]

# (rareté, type, essence) relevés dans les noms de fichiers du site, cartes 001 à 122
META = [
 ("ES","Essence","Rose"),("ES","Essence","Orange"),("ES","Essence","Jade"),
 ("ES","Essence","Violette"),("ES","Essence","Jaune"),("ES","Essence","Mauve"),
 ("ES","Essence","Verte"),("ES","Essence","Bleue"),("ES","Essence","Azur"),
 ("C","Joueur","Rose"),("C","Joueur","Rose"),("PC","Joueur","Rose"),("PC","Joueur","Rose"),
 ("PC","Coach / Instant","Rose"),("C","Objet / Lien","Rose"),("PC","Moment / Instant","Rose"),
 ("R","Joueur","Rose"),("R","Coach / Instant","Rose"),("R","Moment / Instant","Rose"),
 ("R","Moment / Instant","Rose"),("R","Moment / Instant","Rose"),("E","Joueur","Rose"),
 ("E","Joueur","Rose"),("E","Joueur","Rose"),("E","Joueur","Rose"),
 ("E","Coach / Instant","Rose"),("E","Moment / Instant","Rose"),("E","Moment / Instant","Rose"),
 ("C","Joueur","Orange"),("PC","Coach / Instant","Orange"),("C","Coach / Instant","Orange"),
 ("PC","Objet / Lien","Orange"),("C","Moment / Instant","Orange"),("R","Joueur","Orange"),
 ("R","Joueur","Orange"),("R","Objet / Lien","Orange"),("R","Moment / Instant","Orange"),
 ("C","Moment / Instant","Jade"),("C","Joueur","Violette"),("PC","Joueur","Violette"),
 ("C","Moment / Instant","Violette"),("C","Moment / Instant","Violette"),
 ("C","Moment / Instant","Violette"),("R","Joueur","Violette"),("R","Joueur","Violette"),
 ("R","Objet / Lien","Violette"),("E","Moment / Instant","Violette"),
 ("E","Moment / Instant","Violette"),("E","Joueur","Jaune"),("PC","Joueur","Mauve"),
 ("PC","Joueur","Mauve"),("C","Joueur","Mauve"),("C","Objet / Lien","Mauve"),
 ("PC","Moment / Instant","Mauve"),("C","Moment / Instant","Mauve"),("R","Objet / Lien","Mauve"),
 ("E","Joueur","Mauve"),("E","Joueur","Mauve"),("E","Joueur","Mauve"),("E","Joueur","Mauve"),
 ("E","Moment / Instant","Mauve"),("C","Joueur","Verte"),("R","Joueur","Verte"),
 ("PC","Joueur","Bleue"),("C","Objet / Lien","Bleue"),("R","Moment / Instant","Bleue"),
 ("E","Joueur","Bleue"),("E","Joueur","Bleue"),("PC","Joueur","Azur"),
 ("PC","Coach / Instant","Azur"),("R","Joueur","Azur"),("C","Moment / Instant","Neutre"),
 ("C","Joueur","Neutre"),("C","Joueur","Neutre"),("C","Joueur","Neutre"),("C","Joueur","Neutre"),
 ("PC","Joueur","Neutre"),("PC","Joueur","Neutre"),("PC","Joueur","Neutre"),
 ("PC","Joueur","Neutre"),("PC","Joueur","Neutre"),("C","Coach / Instant","Neutre"),
 ("C","Coach / Instant","Neutre"),("PC","Objet / Lien","Neutre"),("C","Objet / Lien","Neutre"),
 ("C","Objet / Lien","Neutre"),("C","Objet / Lien","Neutre"),("C","Objet / Lien","Neutre"),
 ("PC","Objet / Lien","Neutre"),("PC","Objet / Lien","Neutre"),("C","Moment / Instant","Neutre"),
 ("C","Moment / Instant","Neutre"),("C","Moment / Instant","Neutre"),
 ("C","Moment / Instant","Neutre"),("C","Moment / Instant","Neutre"),
 ("C","Moment / Instant","Neutre"),("PC","Moment / Instant","Neutre"),
 ("PC","Moment / Instant","Neutre"),("PC","Moment / Instant","Neutre"),("R","Joueur","Neutre"),
 ("R","Joueur","Neutre"),("R","Joueur","Neutre"),("R","Coach / Instant","Neutre"),
 ("R","Coach / Instant","Neutre"),("R","Objet / Lien","Neutre"),("R","Objet / Lien","Neutre"),
 ("R","Moment / Instant","Neutre"),("R","Moment / Instant","Neutre"),
 ("R","Moment / Instant","Neutre"),("R","Moment / Instant","Neutre"),
 ("R","Moment / Instant","Neutre"),("E","Joueur","Neutre"),("L","Joueur","Rose"),
 ("L","Moment / Instant","Rose"),("L","Moment / Instant","Orange"),("L","Joueur","Mauve"),
 ("L","Joueur","Mauve"),("L","Joueur","Mauve"),("L","Joueur","Neutre"),("L","Joueur","Neutre"),
 ("L","Joueur","Neutre"),("L","Objet / Lien","Neutre"),
]

# Secrètes dévoilées sur le site (type, essence) — relevé du 06/10/2026 sur les noms de fichiers
SECRETES_REVELEES = {
    "123": ("Joueur", "Orange"),   # Le Quatrième du Nom
    "125": ("Joueur", "Mauve"),    # 6W26L
}

RARETE = {"ES": "Essence", "C": "Commune", "PC": "Peu commune", "R": "Rare",
          "E": "Épique", "L": "Légendaire"}

# Promos distribuées hors boosters, d'après la page officielle du TCG
PROMOS = [
    ("P01", "Major de Paris", "Vendue uniquement au Major de Paris et au Pop-Up. Visuel violet, Gotaga et Brawks devant le stade"),
    ("P02", "Major de Paris", "Vendue uniquement au Major de Paris et au Pop-Up. Visuel rose et bleu, chevalier devant la Tour Eiffel"),
    ("P03", "Squeezie le Toxique", "Carte promotionnelle du Starter « Section Toxique » — numéro confirmé par la checklist officielle du jeu"),
    ("P04", "Brawks Déter", "Carte promotionnelle du Starter « Section Déter » — numéro confirmé par la checklist officielle du jeu"),
    ("P05", "Tout Commence Ici", "Carte promotionnelle exclusive, en boîtier acrylique (Pack Fondateur)"),
    ("P06", "Le Marionnettiste - Running the World", "Carte collector des Packs Légendaire et Fondateur"),
]

assert len(NOMS) == 127, len(NOMS)
assert len(META) == 122, len(META)

counts = Counter(NOMS)
seen = Counter()

wb = Workbook()
head_fill = PatternFill("solid", fgColor="1F3864")
head_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")

# ---------------------------------------------------------------- Collection
ws = wb.active
ws.title = "Collection"

headers = ["Image", "N°", "Nom de la carte", "Extension", "Rareté", "Type", "Essence",
           "Reverse existe ?", "Qté normale", "Qté reverse", "Statut", "État",
           "Prix payé (€)", "Valeur estimée (€)", "Date d'acquisition", "Emplacement", "Notes"]
widths = [18, 8, 30, 20, 15, 18, 13, 15, 12, 12, 17, 14, 14, 18, 18, 22, 34]

for i, (h, w) in enumerate(zip(headers, widths), start=1):
    c = ws.cell(row=1, column=i, value=h)
    c.font = head_font
    c.fill = head_fill
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.column_dimensions[get_column_letter(i)].width = w
ws.row_dimensions[1].height = 38

for idx, nom in enumerate(NOMS):
    r = 2 + idx
    num = idx + 1
    ws.cell(row=r, column=2, value=f"{num:03d}")
    ws.cell(row=r, column=3, value=nom)
    ws.cell(row=r, column=4, value="Tout Commence Ici")
    notes = []
    if idx < len(META):
        code, typ, ess = META[idx]
        ws.cell(row=r, column=5, value=RARETE[code])
        ws.cell(row=r, column=6, value=typ)
        ws.cell(row=r, column=7, value=ess)
    else:
        ws.cell(r, 5, value="Secrète")
        if f"{num:03d}" in SECRETES_REVELEES:
            typ, ess = SECRETES_REVELEES[f"{num:03d}"]
            ws.cell(row=r, column=6, value=typ)
            ws.cell(row=r, column=7, value=ess)
        else:
            notes.append("Secrète : type et essence à confirmer, le site ne les détaille pas")
    if counts[nom] > 1:
        seen[nom] += 1
        notes.append(f"Nom présent {counts[nom]}× dans la cardlist (exemplaire {seen[nom]})")
    if notes:
        ws.cell(row=r, column=17, value=" — ".join(notes))

LAST = 1 + len(NOMS)

for k, (num, nom, note) in enumerate(PROMOS):
    r = LAST + 1 + k
    ws.cell(row=r, column=2, value=num)
    ws.cell(row=r, column=3, value=nom)
    ws.cell(row=r, column=4, value="Tout Commence Ici")
    ws.cell(row=r, column=5, value="Promo")
    ws.cell(row=r, column=17, value=note)

LAST += len(PROMOS)

for r in range(2, MAX + 1):
    ws.cell(row=r, column=11,
            value=(f'=IF($C{r}="","",'
                   f'IF(AND($I{r}>0,OR($H{r}<>"Oui",$J{r}>0)),"Complète",'
                   f'IF($I{r}+$J{r}>0,"Manque reverse","Manquante")))'))

table = Table(displayName="Collection", ref=f"A1:Q{LAST}")
table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True,
                                      showColumnStripes=False, showFirstColumn=False,
                                      showLastColumn=False)
ws.add_table(table)

centered = (2, 5, 6, 7, 8, 9, 10, 11, 12, 15)
for r in range(2, MAX + 1):
    ws.row_dimensions[r].height = 86
    for col in range(1, 18):
        c = ws.cell(row=r, column=col)
        c.font = Font(name="Arial", size=10)
        c.alignment = Alignment(horizontal="center" if col in centered else "general",
                                vertical="center", wrap_text=(col in (3, 16, 17)))
    ws.cell(row=r, column=9).number_format = "0"
    ws.cell(row=r, column=10).number_format = "0"
    ws.cell(row=r, column=13).number_format = '#,##0.00 €'
    ws.cell(row=r, column=14).number_format = '#,##0.00 €'
    ws.cell(row=r, column=15).number_format = "DD/MM/YYYY"

ws.freeze_panes = "C2"

# ---------------------------------------------------------------- Listes
lists = wb.create_sheet("Listes")
raretes = ["Commune", "Peu commune", "Rare", "Épique", "Légendaire", "Secrète", "Promo", "Essence"]
cols = {1: raretes,
        2: ["Joueur", "Coach / Instant", "Objet / Lien", "Moment / Instant", "Essence"],
        3: ["Mint", "Near Mint", "Excellent", "Bon", "Abîmé"],
        4: ["Oui", "Non"],
        5: ["Rose", "Orange", "Jade", "Violette", "Jaune", "Mauve", "Verte", "Bleue",
            "Azur", "Neutre"]}
for j, h in enumerate(["Raretés", "Types", "États", "Oui / Non", "Essences"], start=1):
    c = lists.cell(row=1, column=j, value=h)
    c.font = head_font
    c.fill = head_fill
    lists.column_dimensions[get_column_letter(j)].width = 18
for j, vals in cols.items():
    for i, v in enumerate(vals, start=2):
        lists.cell(row=i, column=j, value=v).font = Font(name="Arial", size=10)

n = lists.cell(row=13, column=1, value="Ces listes alimentent les menus déroulants de l'onglet Collection. Modifie-les librement.")
n.font = Font(name="Arial", size=10, italic=True, color="7F7F7F")

for col_letter, src in (("E", "A"), ("F", "B"), ("L", "C"), ("H", "D"), ("G", "E")):
    dv = DataValidation(type="list", formula1=f"=Listes!${src}$2:${src}$20", allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"{col_letter}2:{col_letter}{MAX}")

colors = {"Commune": ("EFEFEF", "404040"), "Peu commune": ("D6EAF8", "1B4F72"),
          "Rare": ("D5F5E3", "145A32"), "Épique": ("E8DAEF", "512E5F"),
          "Légendaire": ("FCF3CF", "7D6608"), "Secrète": ("E4D7F5", "4C1D95"),
          "Promo": ("D7F0F7", "0E6A7D"), "Essence": ("FDEDEC", "922B21")}
for val, (bg, fg) in colors.items():
    ws.conditional_formatting.add(f"E2:E{MAX}", CellIsRule(
        operator="equal", formula=[f'"{val}"'], fill=PatternFill("solid", fgColor=bg),
        font=Font(name="Arial", size=10, bold=True, color=fg)))

statuts = {"Complète": ("D5F5E3", "145A32"), "Manque reverse": ("FDEBD0", "9C4A00"),
           "Manquante": ("FADBD8", "922B21")}
for val, (bg, fg) in statuts.items():
    ws.conditional_formatting.add(f"K2:K{MAX}", CellIsRule(
        operator="equal", formula=[f'"{val}"'], fill=PatternFill("solid", fgColor=bg),
        font=Font(name="Arial", size=10, bold=True, color=fg)))

# ---------------------------------------------------------------- Récap
rc = wb.create_sheet("Récap")
for col, w in (("A", 36), ("B", 18), ("C", 20)):
    rc.column_dimensions[col].width = w
rc["A1"] = "Vue d'ensemble"
rc["A1"].font = Font(name="Arial", size=14, bold=True, color="1F3864")

rows = [
    ("Cartes de la cardlist", f"=COUNTA(Collection!$C$2:$C${MAX})", "0"),
    ("Cartes possédées (au moins 1)",
     f'=SUMPRODUCT((Collection!$C$2:$C${MAX}<>"")*((Collection!$I$2:$I${MAX}+Collection!$J$2:$J${MAX})>0))', "0"),
    ("Progression", "=IFERROR(B3/B2,0)", "0.0%"),
    ("Exemplaires normaux possédés", f"=SUM(Collection!$I$2:$I${MAX})", "0"),
    ("Exemplaires reverse possédés", f"=SUM(Collection!$J$2:$J${MAX})", "0"),
    ("Reverse existantes qu'il te manque",
     f'=SUMPRODUCT((Collection!$H$2:$H${MAX}="Oui")*(Collection!$J$2:$J${MAX}=0)*(Collection!$C$2:$C${MAX}<>""))', "0"),
    ("Total investi", f"=SUM(Collection!$M$2:$M${MAX})", '#,##0.00 €'),
    ("Valeur estimée", f"=SUM(Collection!$N$2:$N${MAX})", '#,##0.00 €'),
    ("Plus / moins-value", "=B9-B8", '#,##0.00 €'),
]
for i, (label, formula, fmt) in enumerate(rows, start=2):
    rc.cell(row=i, column=1, value=label).font = Font(name="Arial", size=11)
    b = rc.cell(row=i, column=2, value=formula)
    b.font = Font(name="Arial", size=11, bold=True)
    b.number_format = fmt
    b.alignment = Alignment(horizontal="right")


def bloc(title, start_row, source_col, values, list_col):
    rc.cell(row=start_row, column=1, value=title).font = Font(
        name="Arial", size=12, bold=True, color="1F3864")
    for j, h in enumerate(["Catégorie", "Possédées", "Total set"], start=1):
        c = rc.cell(row=start_row + 1, column=j, value=h)
        c.font = head_font
        c.fill = head_fill
        c.alignment = Alignment(horizontal="center")
    for k, val in enumerate(values):
        r = start_row + 2 + k
        rc.cell(row=r, column=1, value=f"=Listes!${list_col}${k + 2}").font = Font(name="Arial", size=11)
        b = rc.cell(row=r, column=2, value=(
            f'=SUMPRODUCT((Collection!${source_col}$2:${source_col}${MAX}=$A{r})'
            f'*((Collection!$I$2:$I${MAX}+Collection!$J$2:$J${MAX})>0))'))
        c = rc.cell(row=r, column=3, value=(
            f"=COUNTIFS(Collection!${source_col}$2:${source_col}${MAX},$A{r})"))
        for cell in (b, c):
            cell.font = Font(name="Arial", size=11)
            cell.alignment = Alignment(horizontal="center")


bloc("Par rareté", 12, "E", raretes, "A")
bloc("Par essence", 23, "G", cols[5], "E")

s = rc.cell(row=36, column=1, value="Colonne « Possédées » : cartes dont tu as au moins un exemplaire. Colonne « Total set » : nombre de cartes de cette catégorie dans l'extension.")
s.font = Font(name="Arial", size=10, italic=True, color="7F7F7F")

# ---------------------------------------------------------------- Mode d'emploi
me = wb.create_sheet("Mode d'emploi", 0)
me.column_dimensions["A"].width = 112

lines = [
    ("Suivi de collection — TCG Gentle Mates, « Tout Commence Ici »", "title"),
    ("", None),
    ("Ce qui est déjà rempli", "h"),
    ("Les 127 cartes du set, avec leur numéro officiel, leur nom, leur rareté, leur type et leur essence. Le tout relevé sur gentlemates.com/tcg-cards le 14/09/2026, une fois toutes les cartes révélées.", None),
    ("Les cartes 123 à 127 sont les secrètes, numérotées 123/122 à 127/122 : le set compte 122 cartes plus 5 secrètes. Numérotation et noms confirmés pour les 5, y compris la 123 (« Le Quatrième du Nom »), par la checklist officielle du jeu. Les secrètes 123 et 125 sont dévoilées (type et essence renseignés) ; les 124, 126 et 127 restent à compléter.", None),
    ("Après la cardlist, six promos numérotées P01 à P06 : elles ne sortent pas des boosters. P01 et P02 étaient vendues uniquement au Major de Paris, les autres viennent des starters et du Pack Fondateur. Ajoute une ligne si tu en trouves d'autres, le tableau s'étend tout seul.", None),
    ("", None),
    ("Ce qu'il te reste à renseigner", "h"),
    ("La colonne « Reverse existe ? », les quantités, et tes informations perso (état, prix payé, valeur estimée, emplacement). Le reste est déjà là. D'après la checklist officielle du jeu, une reverse existe pour les cartes Essence, Commune, Peu commune et Rare, mais pas pour les Épique, Légendaire, Secrète et Promo — sans exception constatée à ce jour.", None),
    ("", None),
    ("Une ligne = une carte, pas une version", "h"),
    ("Indique combien d'exemplaires tu as de chaque version dans Qté normale et Qté reverse. Inutile de dupliquer la ligne. Quelques noms reviennent plusieurs fois dans le set (CHAMPIONS DU MAJOR, Ravitaillement, Leakof) : ce sont des cartes distinctes, avec des numéros différents.", None),
    ("", None),
    ("La colonne Statut se calcule toute seule", "h"),
    ("Verte « Complète » : tu as la carte, et sa reverse si elle existe. Orange « Manque reverse ». Rouge « Manquante ». Ne tape rien dedans — filtre-la pour obtenir ta liste de recherche.", None),
    ("", None),
    ("Ajouter les visuels", "h"),
    ("Clique sur une cellule de la colonne A, puis Insertion → Images → Placer dans la cellule. L'image reste dans la cellule et suit les tris et les filtres ; la hauteur des lignes est déjà réglée pour ça.", None),
    ("", None),
    ("Où ne rien taper", "h"),
    ("Les onglets Récap (entièrement calculé) et Listes (qui alimente les menus déroulants). Toute la saisie se fait dans l'onglet Collection.", None),
    ("", None),
    ("Sur iPhone et iPad", "h"),
    ("Enregistre le fichier dans OneDrive : il apparaîtra dans l'appli Excel sur tes deux appareils Apple. Saisis par lots depuis le PC, garde le mobile pour la consultation.", None),
]

r = 1
for text, kind in lines:
    c = me.cell(row=r, column=1, value=text)
    if kind == "title":
        c.font = Font(name="Arial", size=16, bold=True, color="1F3864")
        me.row_dimensions[r].height = 26
    elif kind == "h":
        c.font = Font(name="Arial", size=12, bold=True, color="1F3864")
    else:
        c.font = Font(name="Arial", size=11)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        me.row_dimensions[r].height = 46
    r += 1

wb.save(os.path.join(BASE_DIR, "Suivi_collection_cartes.xlsx"))
print("ok", len(NOMS), "cartes, derniere ligne", LAST)
