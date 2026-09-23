# -*- coding: utf-8 -*-
"""
ArterPerBiotop.pyt

Sammanställer vilka arter som förekommer i varje naturvärdesbiotop och hur många
av dem som är invasiva, skyddade och rödlistade. Resultatet är en tabell med en
rad per biotop:

  Biotop (Kart-ID) | Artlista | Varav invasiva (lista) | Varav invasiva (antal) |
  Varav skyddade (lista) | Varav skyddade (antal) |
  Varav rödlistade (lista) | Varav rödlistade (antal)

Fältnamnen i tabellen är ASCII, rubrikerna ovan ligger som alias. En export till
Excel använder alias som kolumnrubriker.

Biotopen identifieras av ett valbart fält i polygonlagret, normalt A_kartID.
Polygoner med samma värde räknas som samma biotop.

Artpunkterna förväntas ha heltalsfält (1/0) som flaggar rödlistade, skyddade och
invasiva arter, och gärna ett källfält som skiljer egna fynd från uttag ur
Artdatabanken. Standardnamnen följer SIS-mallen för NVI: taxon_svensktNamn,
A_rodlist, A_skyddad, A_IV och kalla. Andra fält går att välja i dialogen. En art
som i en biotop bara är känd från andra källor än egen inventering får "(ADB)"
efter namnet.

Punkt-i-polygon avgörs med Spatial Join (en rad per punkt och polygon, så en punkt
på en gemensam gräns räknas i båda biotoperna). Det ger stöd för lagerurval,
definitionsfrågor, olika koordinatsystem och ett valfritt sökavstånd. Attributen
läses sedan direkt från indatalagren via TARGET_FID/JOIN_FID och grupperingen görs
i Python. Det undviker två begränsningar i systemverktygen (Pro 3.6): Sort på
flera fält kräver Advanced-licens, och Calculate Fields skapar nya fält som text,
som Summary Statistics sedan inte kan summera.

Krav: ArcGIS Pro 3.x (arcpy). Inga paket utöver Pythons standardbibliotek.
"""

import os
import re
import time
from xml.sax.saxutils import escape

import arcpy

# ── Konstanter ────────────────────────────────────────────────────────────────

ID_FIELD_DEFAULTS = ["A_kartID", "EkoID"]      # första som finns väljs
DEFAULT_NAME_FIELD = "taxon_svensktNamn"
DEFAULT_RED_FIELD = "A_rodlist"
DEFAULT_PROT_FIELD = "A_skyddad"
DEFAULT_INV_FIELD = "A_IV"
DEFAULT_SOURCE_FIELD = "kalla"
DEFAULT_OWN_VALUE = "Fältinventering"

ADB_MARK = " (ADB)"
SEPARATOR = ", "

SOURCE_OWN = "Egen inventering"
SOURCE_OTHER = "Annan källa"
SOURCE_BOTH = "Båda"

# Kategorier i utdataordning: (nyckel, fältprefix, rubrik)
CATEGORIES = [
    ("inv", "Invasiva", "invasiva"),
    ("prot", "Skyddade", "skyddade"),
    ("red", "Rodlistade", "rödlistade"),
]

# ArcGIS fälttyp (Field.type) -> AddField field_type.
FIELD_TYPE_MAP = {
    "SmallInteger": "SHORT",
    "Integer": "LONG",
    "BigInteger": "BIGINTEGER",
    "Single": "FLOAT",
    "Double": "DOUBLE",
    "String": "TEXT",
    "Date": "DATE",
    "DateOnly": "DATEONLY",
    "OID": "LONG",
}

# Svensk sorteringsordning: å, ä, ö efter z. Tecknen { | } kommer direkt efter z
# i teckentabellen, så en vanlig strängsortering ger rätt ordning.
_SV_SORT = str.maketrans({"å": "{", "ä": "|", "ö": "}", "é": "e", "è": "e", "ü": "y"})

TOOL_SUMMARY = (
    "Sammanställer vilka arter som förekommer i varje naturvärdesbiotop och hur många av "
    "dem som är invasiva, skyddade och rödlistade. Ger en tabell med en rad per biotop och "
    "valfritt en Excel-fil med samma innehåll."
)

# Verktygstips per parameter, visas i verktygsdialogen. Se _write_tool_metadata.
TOOLTIPS = {
    "in_points": (
        "Punktlager med artförekomster. Om lagret har ett urval eller en definitionsfråga "
        "används bara de punkterna, så ett urval kan till exempel begränsa sammanställningen "
        "till fynd efter ett visst år.\n"
        "Lagret läses bara, det ändras inte."
    ),
    "in_biotopes": (
        "Polygonlager med naturvärdesbiotoper. Urval och definitionsfråga respekteras, så bara "
        "de biotoper som syns i lagret kommer med.\n"
        "Lagret och artpunkterna får ha olika koordinatsystem."
    ),
    "id_field": (
        "Fält i biotoplagret som identifierar en biotop, normalt A_kartID (Kart-ID). Polygoner "
        "med samma värde räknas som samma biotop och blir en rad i tabellen. Polygoner utan "
        "värde hoppas över.\n"
        "Rubriken i tabellen blir \"Biotop (...)\" med fältets alias."
    ),
    "name_field": (
        "Fält i punktlagret med artens namn, normalt taxon_svensktNamn. Namnen används som de "
        "står, bortsett från mellanslag i början och slutet: stavningsvarianter ger separata "
        "arter, så rensa namnen först. Punkter utan namn hoppas över och räknas i en varning."
    ),
    "out_table": (
        "Tabell i en fil-geodatabas med en rad per biotop: artlista samt lista och antal för "
        "invasiva, skyddade och rödlistade arter. En art som är både skyddad och rödlistad står "
        "i båda kolumnerna. Listorna sorteras i svensk bokstavsordning.\n"
        "Finns tabellen redan skrivs den över om överskrivning är tillåten i Pro."
    ),
    "out_excel": (
        "Valfri Excel-fil (.xlsx) med samma tabell och rubrikerna som kolumnnamn. En befintlig "
        "fil skrivs över. Om filen är öppen i Excel misslyckas exporten med en varning, men "
        "tabellen i geodatabasen skapas ändå."
    ),
    "red_field": (
        "Heltalsfält i punktlagret där 1 betyder rödlistad art, normalt A_rodlist. Tomt värde "
        "räknas som 0. Lämna parametern tom om lagret saknar fältet: kolumnerna för rödlistade "
        "blir då tomma respektive 0."
    ),
    "prot_field": (
        "Heltalsfält i punktlagret där 1 betyder skyddad art (artskyddsförordningen), normalt "
        "A_skyddad. Tomt värde räknas som 0. Lämna tom om lagret saknar fältet."
    ),
    "inv_field": (
        "Heltalsfält i punktlagret där 1 betyder invasiv främmande art, normalt A_IV. Tomt värde "
        "räknas som 0. Lämna tom om lagret saknar fältet."
    ),
    "source_field": (
        "Textfält i punktlagret som anger fyndets källa, normalt kalla. Arter som i en biotop "
        "bara är kända från andra källor än egen inventering får \"(ADB)\" efter namnet i "
        "listorna. Antalen påverkas inte.\n"
        "Lämna tomt för att inte skilja på källor."
    ),
    "own_value": (
        "Värdet i källfältet som betyder eget fynd, normalt \"Fältinventering\". Alla andra "
        "värden, även tomma, räknas som annan källa. Skiftläge och mellanslag i början och "
        "slutet ignoreras."
    ),
    "out_long_table": (
        "Valfri tabell i en fil-geodatabas med en rad per biotop och art: flaggor, antal fynd, "
        "antal egna fynd och källa. Bra för kontroll, filtrering och pivotering.\n"
        "Lämna tomt för att hoppa över."
    ),
    "search_distance": (
        "Punkter inom detta avstånd från en biotop räknas som att de ligger i den, till exempel "
        "5 meter för GPS-fel vid kanten. En punkt kan då räknas i flera biotoper.\n"
        "Lämna tomt för att bara ta med punkter som ligger i eller på kanten av en polygon."
    ),
    "include_empty": (
        "Ta med biotoper utan några artfynd, med tomma listor och antal 0. Standard är att ta "
        "med dem, så att tabellen har en rad per biotop."
    ),
}


# ── Hjälpfunktioner ───────────────────────────────────────────────────────────

def _fmt_duration(seconds):
    if seconds < 60:
        return "{:.1f} s".format(seconds)
    return "{} min {:02d} s".format(int(seconds // 60), int(seconds % 60))


def _sort_key(name):
    return name.casefold().translate(_SV_SORT)


def _id_sort_key(value):
    """Numerisk ordning för ID som är tal eller siffersträngar ("2" före "10"),
    annars svensk bokstavsordning. Tal sorteras före text."""
    text = str(value).strip()
    if re.fullmatch(r"-?\d+(\.\d+)?", text):
        return (0, float(text), text)
    return (1, 0.0, _sort_key(text))


def _norm_id(value):
    if isinstance(value, str):
        value = value.strip()
        return value or None
    return value


def _flag(value):
    """1/0 från ett flaggvärde. Tomt, 0 och '0' ger 0."""
    if value is None:
        return 0
    if isinstance(value, str):
        value = value.strip()
        return 0 if value in ("", "0") else 1
    return 1 if value else 0


def _field_map(layer):
    return {f.name.lower(): f for f in arcpy.ListFields(layer)}


def _id_heading(field):
    """Rubrik för ID-kolumnen: 'Biotop (<alias>)', utan parentes i slutet av aliaset
    ('Kart-ID (arb)' blir 'Kart-ID')."""
    alias = (field.aliasName or field.name).strip()
    alias = re.sub(r"\s*\([^)]*\)\s*$", "", alias) or field.name
    return "Biotop ({})".format(alias)


def _in_gdb(path):
    parent = os.path.dirname(path)
    while parent:
        if parent.lower().endswith(".gdb"):
            return True
        new_parent = os.path.dirname(parent)
        if new_parent == parent:
            return False
        parent = new_parent
    return False


def _create_table(path, fields):
    """
    Skapa en tabell i en geodatabas. fields är en lista av
    (namn, AddField-typ, alias, längd eller None).
    """
    if arcpy.Exists(path):
        if not arcpy.env.overwriteOutput:
            raise ValueError(
                "Utdatatabellen {} finns redan. Tillåt överskrivning i Pro eller välj ett "
                "annat namn.".format(path)
            )
        arcpy.management.Delete(path)
    ws, name = os.path.split(path)
    arcpy.management.CreateTable(ws, name)
    for fname, ftype, alias, length in fields:
        arcpy.management.AddField(
            path, fname, ftype, field_alias=alias,
            field_length=length if ftype == "TEXT" else None,
        )


def _id_field_spec(field, alias):
    ftype = FIELD_TYPE_MAP.get(field.type, "TEXT")
    length = None
    if ftype == "TEXT":
        length = field.length if field.type == "String" else 255
    return ("Biotop", ftype, alias, length)


def _text_length(values, minimum=255):
    longest = max((len(v) for v in values if v), default=0)
    return max(minimum, longest + 10)


def _drop_oid_column(xlsx, oid_name):
    """TableToExcel tar alltid med OBJECTID som första kolumn. Ta bort den med
    openpyxl, som följer med Pro. Saknas paketet blir kolumnen kvar."""
    try:
        import openpyxl
    except ImportError:
        return
    wb = openpyxl.load_workbook(xlsx)
    ws = wb.active
    if ws.cell(row=1, column=1).value == oid_name:
        ws.delete_cols(1)
        wb.save(xlsx)


class _Steps:
    """Stegräknare: namnger varje steg i förloppsindikatorn och i meddelandena."""

    def __init__(self, total, messages):
        self.total = total
        self.messages = messages
        self.k = 0
        self.t_step = time.time()

    def start(self, text):
        self.k += 1
        self.t_step = time.time()
        label = "Steg {} av {}: {}".format(self.k, self.total, text)
        arcpy.SetProgressor("default", label)
        self.messages.addMessage(label)

    def done(self, text=None):
        msg = "    klart på {}".format(_fmt_duration(time.time() - self.t_step))
        if text:
            msg += ". " + text
        self.messages.addMessage(msg)


# ── Verktygstips ──────────────────────────────────────────────────────────────

def _write_tool_metadata(tool_cls, toolbox_alias):
    """
    Skriv verktygets metadatafil med parameterförklaringar från TOOLTIPS.

    Pro läser verktygstipsen i dialogen från <verktygslåda>.<verktyg>.pyt.xml
    (elementet dialogReference per parameter). Det finns inget attribut på
    arcpy.Parameter för detta. Filen skrivs bara om innehållet har ändrats.
    """
    here = os.path.dirname(os.path.abspath(__file__))
    toolbox = os.path.splitext(os.path.basename(__file__))[0]
    path = os.path.join(here, "{}.{}.pyt.xml".format(toolbox, tool_cls.__name__))

    def html(text):
        body = escape(text).replace("\n", "</SPAN></P><P><SPAN>")
        return escape('<DIV STYLE="text-align:Left;"><P><SPAN>{}</SPAN></P></DIV>'.format(body))

    tool = tool_cls()
    params = []
    for p in tool.getParameterInfo():
        tip = TOOLTIPS.get(p.name)
        if not tip:
            continue
        params.append(
            '<param name="{n}" displayname="{d}" type="{t}" direction="{r}">'
            "<dialogReference>{h}</dialogReference>"
            "<pythonReference>{h}</pythonReference></param>".format(
                n=p.name, d=escape(p.displayName, {'"': "&quot;"}),
                t=p.parameterType, r=p.direction, h=html(tip))
        )
    xml = (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<metadata xml:lang="sv"><Esri><ArcGISFormat>1.0</ArcGISFormat></Esri>'
        '<tool name="{name}" displayname="{label}" toolboxalias="{alias}" xmlns="">'
        "<parameters>{params}</parameters><summary>{summary}</summary></tool>"
        "<dataIdInfo><idCitation><resTitle>{label}</resTitle></idCitation>"
        "<idAbs>{summary}</idAbs></dataIdInfo></metadata>\n"
    ).format(name=tool_cls.__name__, label=escape(tool.label), alias=toolbox_alias,
             params="".join(params), summary=html(TOOL_SUMMARY))

    try:
        with open(path, encoding="utf-8") as fh:
            if fh.read() == xml:
                return
    except OSError:
        pass
    try:
        with open(path, "w", encoding="utf-8") as fh:
            fh.write(xml)
    except OSError:
        # Skrivskyddad plats: verktyget fungerar ändå, bara utan verktygstips.
        pass


# =============================================================================
# Toolbox
# =============================================================================

class Toolbox:
    def __init__(self):
        self.label = "Arter per biotop"
        self.alias = "arter_per_biotop"
        self.tools = [ArterPerBiotop]
        _write_tool_metadata(ArterPerBiotop, self.alias)


class ArterPerBiotop:
    def __init__(self):
        self.label = "Arter per naturvärdesbiotop"
        self.description = TOOL_SUMMARY
        self.canRunInBackground = False

    def getParameterInfo(self):
        p_points = arcpy.Parameter(
            displayName="Artförekomster (punkter)", name="in_points",
            datatype="GPFeatureLayer", parameterType="Required", direction="Input")
        p_points.filter.list = ["Point", "Multipoint"]

        p_bio = arcpy.Parameter(
            displayName="Naturvärdesbiotoper (polygoner)", name="in_biotopes",
            datatype="GPFeatureLayer", parameterType="Required", direction="Input")
        p_bio.filter.list = ["Polygon"]

        p_id = arcpy.Parameter(
            displayName="Biotopens ID-fält", name="id_field", datatype="Field",
            parameterType="Required", direction="Input")
        p_id.parameterDependencies = [p_bio.name]
        p_id.filter.list = ["Short", "Long", "Text", "Double", "Float", "OID"]

        p_name = arcpy.Parameter(
            displayName="Artnamnsfält", name="name_field", datatype="Field",
            parameterType="Required", direction="Input")
        p_name.parameterDependencies = [p_points.name]
        p_name.filter.list = ["Text"]

        p_out = arcpy.Parameter(
            displayName="Utdatatabell", name="out_table", datatype="DETable",
            parameterType="Required", direction="Output")

        p_xlsx = arcpy.Parameter(
            displayName="Excel-fil", name="out_excel", datatype="DEFile",
            parameterType="Optional", direction="Output")
        p_xlsx.filter.list = ["xlsx"]

        cat_fields = "Fält i artlagret"
        p_red = arcpy.Parameter(
            displayName="Rödlistad (1/0)", name="red_field", datatype="Field",
            parameterType="Optional", direction="Input", category=cat_fields)
        p_prot = arcpy.Parameter(
            displayName="Skyddad (1/0)", name="prot_field", datatype="Field",
            parameterType="Optional", direction="Input", category=cat_fields)
        p_inv = arcpy.Parameter(
            displayName="Invasiv (1/0)", name="inv_field", datatype="Field",
            parameterType="Optional", direction="Input", category=cat_fields)
        for p in (p_red, p_prot, p_inv):
            p.parameterDependencies = [p_points.name]
            p.filter.list = ["Short", "Long", "Text"]
        p_src = arcpy.Parameter(
            displayName="Källfält", name="source_field", datatype="Field",
            parameterType="Optional", direction="Input", category=cat_fields)
        p_src.parameterDependencies = [p_points.name]
        p_src.filter.list = ["Text"]
        p_own = arcpy.Parameter(
            displayName="Värde för egna fynd", name="own_value", datatype="GPString",
            parameterType="Optional", direction="Input", category=cat_fields)
        p_own.value = DEFAULT_OWN_VALUE

        cat_opt = "Alternativ"
        p_long = arcpy.Parameter(
            displayName="Kontrolltabell (en rad per biotop och art)", name="out_long_table",
            datatype="DETable", parameterType="Optional", direction="Output", category=cat_opt)
        p_dist = arcpy.Parameter(
            displayName="Sökavstånd", name="search_distance", datatype="GPLinearUnit",
            parameterType="Optional", direction="Input", category=cat_opt)
        p_empty = arcpy.Parameter(
            displayName="Ta med biotoper utan arter", name="include_empty",
            datatype="GPBoolean", parameterType="Optional", direction="Input", category=cat_opt)
        p_empty.value = True

        return [p_points, p_bio, p_id, p_name, p_out, p_xlsx,
                p_red, p_prot, p_inv, p_src, p_own,
                p_long, p_dist, p_empty]

    def isLicensed(self):
        return True

    def updateParameters(self, parameters):
        (p_points, p_bio, p_id, p_name, _p_out, _p_xlsx,
         p_red, p_prot, p_inv, p_src, _p_own,
         _p_long, _p_dist, _p_empty) = parameters

        # Fyll i standardfält när ett lager väljs och fältet finns. Bara när lagret
        # just har ändrats och bara i parametrar som användaren inte själv har rört.
        if p_points.value and p_points.altered and not p_points.hasBeenValidated:
            try:
                names = {f.name.lower(): f.name for f in arcpy.ListFields(p_points.value)}
            except Exception:
                names = {}
            for p, default in ((p_name, DEFAULT_NAME_FIELD), (p_red, DEFAULT_RED_FIELD),
                               (p_prot, DEFAULT_PROT_FIELD), (p_inv, DEFAULT_INV_FIELD),
                               (p_src, DEFAULT_SOURCE_FIELD)):
                if not p.altered and default.lower() in names:
                    p.value = names[default.lower()]

        if p_bio.value and p_bio.altered and not p_bio.hasBeenValidated and not p_id.altered:
            try:
                names = {f.name.lower(): f.name for f in arcpy.ListFields(p_bio.value)}
            except Exception:
                names = {}
            for default in ID_FIELD_DEFAULTS:
                if default.lower() in names:
                    p_id.value = names[default.lower()]
                    break

    def updateMessages(self, parameters):
        (_p_points, _p_bio, _p_id, _p_name, p_out, _p_xlsx,
         _p_red, _p_prot, _p_inv, p_src, p_own,
         p_long, p_dist, _p_empty) = parameters

        for p in (p_out, p_long):
            if p.valueAsText and not _in_gdb(p.valueAsText):
                p.setErrorMessage(
                    "Tabellen måste ligga i en fil-geodatabas (.gdb). Använd Excel-fil för att "
                    "få resultatet som kalkylblad.")

        if p_out.valueAsText and p_long.valueAsText and \
                os.path.normcase(p_out.valueAsText) == os.path.normcase(p_long.valueAsText):
            p_long.setErrorMessage("Kontrolltabellen måste ha ett annat namn än utdatatabellen.")

        if p_src.valueAsText and not (p_own.valueAsText or "").strip():
            p_own.setErrorMessage("Ange vilket värde i källfältet som betyder eget fynd.")

        if p_dist.valueAsText:
            try:
                if float(p_dist.valueAsText.split()[0].replace(",", ".")) < 0:
                    p_dist.setErrorMessage("Sökavståndet kan inte vara negativt.")
            except (ValueError, IndexError):
                pass

    def execute(self, parameters, messages):
        v = [p.valueAsText for p in parameters]
        try:
            _run(
                in_points=parameters[0].value, in_biotopes=parameters[1].value,
                id_field=v[2], name_field=v[3], out_table=v[4], out_excel=v[5],
                red_field=v[6], prot_field=v[7], inv_field=v[8],
                source_field=v[9], own_value=v[10],
                out_long_table=v[11], search_distance=v[12],
                include_empty=parameters[13].value is not False,
                messages=messages,
            )
        except ValueError as exc:
            messages.addErrorMessage(str(exc))
            raise arcpy.ExecuteError

    def postExecute(self, parameters):
        return


# =============================================================================
# Körningens innehåll (separat funktion, går att testa utanför Pro)
# =============================================================================

def _run(in_points, in_biotopes, id_field, name_field, out_table, out_excel=None,
         red_field=None, prot_field=None, inv_field=None, source_field=None,
         own_value=DEFAULT_OWN_VALUE, out_long_table=None, search_distance=None,
         include_empty=True, messages=None):
    own_key = (own_value or "").strip().casefold()
    steps = _Steps(4 + bool(out_excel), messages)
    join_fc = r"memory\arter_per_biotop_join"
    flag_fields = {"red": red_field, "prot": prot_field, "inv": inv_field}

    pfields = _field_map(in_points)
    bfields = _field_map(in_biotopes)
    for fld in (name_field, red_field, prot_field, inv_field, source_field):
        if fld and fld.lower() not in pfields:
            raise ValueError("Fältet {} finns inte i artlagret.".format(fld))
    if id_field.lower() not in bfields:
        raise ValueError("Fältet {} finns inte i biotoplagret.".format(id_field))
    id_fld = bfields[id_field.lower()]

    try:
        # ── 1 Punkt i polygon
        steps.start("Kopplar artpunkter till biotoper (Spatial Join)")
        if arcpy.Exists(join_fc):
            arcpy.management.Delete(join_fc)
        arcpy.analysis.SpatialJoin(
            in_points, in_biotopes, join_fc, "JOIN_ONE_TO_MANY", "KEEP_COMMON",
            match_option="INTERSECT", search_radius=search_distance or None)
        with arcpy.da.SearchCursor(join_fc, ["TARGET_FID", "JOIN_FID"]) as cur:
            pairs = [(t, j) for t, j in cur if t is not None and j is not None and j >= 0]
        steps.done("{} punkter ligger i minst en biotop{} ({} kopplingar).".format(
            len({t for t, _ in pairs}),
            " eller inom {}".format(search_distance) if search_distance else "", len(pairs)))

        # ── 2 Läs attribut
        steps.start("Läser attribut från artpunkter och biotoper")
        pcols = [arcpy.Describe(in_points).OIDFieldName, name_field] + \
            [f for f in (red_field, prot_field, inv_field, source_field) if f]
        idx = {f: i for i, f in enumerate(pcols)}
        with arcpy.da.SearchCursor(in_points, pcols) as cur:
            points = {row[0]: row for row in cur}
        biotopes = {}      # polygonens OID -> biotop-ID
        n_poly_no_id = 0
        with arcpy.da.SearchCursor(in_biotopes, ["OID@", id_field]) as cur:
            for oid, bid in cur:
                bid = _norm_id(bid)
                if bid is None:
                    n_poly_no_id += 1
                else:
                    biotopes[oid] = bid
        bio_ids = sorted(set(biotopes.values()), key=_id_sort_key)
        n_multi = len(biotopes) - len(bio_ids)
        steps.done("{} artpunkter, {} polygoner som bildar {} biotoper.".format(
            len(points), len(biotopes), len(bio_ids)))
        if n_multi:
            messages.addMessage(
                "    {} polygoner delar {} med en annan polygon och räknas ihop med den.".format(
                    n_multi, id_field))
        if n_poly_no_id:
            messages.addWarningMessage(
                "{} polygoner saknar värde i {} och hoppas över.".format(n_poly_no_id, id_field))

        # ── 3 Gruppera
        steps.start("Grupperar arter per biotop")
        # (biotop, art) -> {"n": fynd, "own": egna fynd, "red"/"prot"/"inv": 1/0}
        agg = {}
        seen = set()        # (punkt, biotop): en punkt räknas en gång per biotop
        n_no_name = 0
        for t, j in pairs:
            row = points.get(t)
            bid = biotopes.get(j)
            if row is None or bid is None or (t, bid) in seen:
                continue
            seen.add((t, bid))
            name = row[idx[name_field]]
            name = " ".join(name.split()) if isinstance(name, str) else None
            if not name:
                n_no_name += 1
                continue
            a = agg.setdefault((bid, name), {"n": 0, "own": 0, "red": 0, "prot": 0, "inv": 0})
            a["n"] += 1
            if source_field and (row[idx[source_field]] or "").strip().casefold() == own_key:
                a["own"] += 1
            for key, fld in flag_fields.items():
                if fld:
                    a[key] = max(a[key], _flag(row[idx[fld]]))
        if n_no_name:
            messages.addWarningMessage(
                "{} punkter i biotoper saknar artnamn och har hoppats över.".format(n_no_name))

        by_bio = {}
        for (bid, name), a in agg.items():
            by_bio.setdefault(bid, []).append((name, a))

        rows = []
        for bid in bio_ids:
            species = sorted(by_bio.get(bid, []), key=lambda s: _sort_key(s[0]))
            if not species and not include_empty:
                continue
            labels = [(n + (ADB_MARK if source_field and a["own"] == 0 else ""), a)
                      for n, a in species]
            row = [bid, SEPARATOR.join(lbl for lbl, _ in labels) or None]
            for key, _, _ in CATEGORIES:
                chosen = [lbl for lbl, a in labels if a[key]]
                row += [SEPARATOR.join(chosen) or None, len(chosen)]
            rows.append(row)
        n_species = len({n for _, n in agg})
        steps.done("{} arter i {} biotoper. {} biotoper utan fynd{}.".format(
            n_species, len(by_bio), len(bio_ids) - len(by_bio),
            "" if include_empty else " utelämnas"))
        for key, prefix, word in CATEGORIES:
            if not flag_fields[key]:
                messages.addWarningMessage(
                    "Inget fält angivet för {}: kolumnerna blir tomma.".format(word))

        # ── 4 Skriv tabeller
        steps.start("Skriver utdatatabell")
        heading = _id_heading(id_fld)
        cols = list(zip(*rows)) if rows else [[]] * (2 + 2 * len(CATEGORIES))
        spec = [_id_field_spec(id_fld, heading),
                ("Artlista", "TEXT", "Artlista", _text_length(cols[1]))]
        for i, (key, prefix, word) in enumerate(CATEGORIES):
            spec += [(prefix + "_lista", "TEXT", "Varav {} (lista)".format(word),
                      _text_length(cols[2 + 2 * i])),
                     (prefix + "_antal", "LONG", "Varav {} (antal)".format(word), None)]
        _create_table(out_table, spec)
        with arcpy.da.InsertCursor(out_table, [s[0] for s in spec]) as cur:
            for row in rows:
                cur.insertRow(row)
        messages.addMessage("    {} rader i {}".format(len(rows), out_table))

        if out_long_table:
            long_spec = [_id_field_spec(id_fld, heading),
                         ("Artnamn", "TEXT", "Artnamn", _text_length([n for _, n in agg], 100)),
                         ("Invasiv", "SHORT", "Invasiv", None),
                         ("Skyddad", "SHORT", "Skyddad", None),
                         ("Rodlistad", "SHORT", "Rödlistad", None),
                         ("Antal_fynd", "LONG", "Antal fynd", None)]
            if source_field:
                long_spec += [("Antal_egna", "LONG", "Antal egna fynd", None),
                              ("Kalla", "TEXT", "Källa", 30)]
            _create_table(out_long_table, long_spec)
            n_long = 0
            with arcpy.da.InsertCursor(out_long_table, [s[0] for s in long_spec]) as cur:
                for bid in bio_ids:
                    for name, a in sorted(by_bio.get(bid, []), key=lambda s: _sort_key(s[0])):
                        row = [bid, name, a["inv"], a["prot"], a["red"], a["n"]]
                        if source_field:
                            src = SOURCE_OTHER if a["own"] == 0 else (
                                SOURCE_OWN if a["own"] == a["n"] else SOURCE_BOTH)
                            row += [a["own"], src]
                        cur.insertRow(row)
                        n_long += 1
            messages.addMessage("    {} rader i {}".format(n_long, out_long_table))
        steps.done()

        # ── 5 Excel
        if out_excel:
            steps.start("Exporterar till Excel")
            try:
                if os.path.exists(out_excel):
                    os.remove(out_excel)
                arcpy.conversion.TableToExcel(out_table, out_excel, "ALIAS")
                _drop_oid_column(out_excel, arcpy.Describe(out_table).OIDFieldName)
                steps.done(out_excel)
            except (OSError, arcpy.ExecuteError) as exc:
                messages.addWarningMessage(
                    "Kunde inte skriva {} ({}). Är filen öppen i Excel? Tabellen i "
                    "geodatabasen är skapad.".format(
                        out_excel, (str(exc).strip().splitlines() or [""])[0]))

        if source_field:
            messages.addMessage(
                "Arter markerade {} är i den biotopen bara kända från andra källor än "
                "egen inventering.".format(ADB_MARK.strip()))
    finally:
        arcpy.ResetProgressor()
        if arcpy.Exists(join_fc):
            arcpy.management.Delete(join_fc)
