"""
field_map.py
============
Maps variable PDF field labels → canonical Excel column names.

Strategy (applied in order):
  1. Exact match (normalised)
  2. Alias lookup  (common PDF variants per column)
  3. Fuzzy token match using difflib (threshold-based)

All matching is case-insensitive and ignores: spaces, underscores,
hyphens, slashes, dots, and parentheses.
"""

import re
import difflib
from typing import Optional

# ── 1. Canonical Excel column names (all 129) ──────────────────────────────
EXCEL_COLUMNS = [
    "TagNo", "Item Name", "Quantity", "Project",
    "Plate MATERIAL", "Flange MATERIAL", "Flange Type", "Piece Type",
    "Weld Type", "Holder Type", "Line_Size", "Size in NPS OR DN",
    "Flange Schedule", "Pipe Wall Thk", "Rating", "RJ HOLDER MATERIAL",
    "Drain/Vent", "Smooth Finish", "Serration", "Stellite",
    "Fluid/Service_Name", "Service Discription", "Service Type",
    "Calculation Standard", "Flow Element Type", "Flow Rate Unit",
    "Flow Rate Minimum", "Flow Rate Maximum", "Flow Rate Normal",
    "Flow Rate at FullScale", "Tapping", "Multi Hole", "Pipe Material",
    "Pressure Unit", "Temp Unit", "Viscosity Unit", "Density Unit",
    "Density_Calc_Method", "MolecularWeight_Customer",
    "Compressibilty_atFlow_Customer", "Compressibilty_atBase_Customer",
    "Gas Name", "Base Pressure", "Upstream Pressure", "Base_Temperature",
    "Operating_Temperature", "Vapour_Pressure", "Density_at_Base_Customer",
    "Density_at_Flow_Customer", "Viscosity_Customer",
    "IsentropicExponenet_Customer", "DP Unit", "DP at Full Scale",
    "Design Pressure Discription", "Design Pressure",
    "Design Temp Discription", "Design Temp", "Flange Standard",
    "Plate Thk Customer", "Spare Plate Required", "Stud/Nut", "Gasket",
    "JackBolt", "Packing_Cost_Manual", "Accessories", "Accessories_Amt",
    "IBR", "IBR_Amt_Manual", "Nace", "Nace_Type", "Nace_Percent",
    "Calibration", "CalibrationAmt_Manual", "Freight_Required",
    "Freight_Amt_Manual", "Special_Requirement", "Special_Requirement_Amt",
    "OFA Tap Orientation", "FNA Tap Orientation", "Venturi Tap Orientation",
    "Plug_Material", "Pressure_Tap_Angle", "JackBolt_Position",
    "ItemCode_Heading", "ItemCode", "TestSchedule",
    "FNA Pipe Machining Required", "FNA Pipe Machined Cost",
    "FNA Total Assy. Length", "FNA Upstream Length",
    "FNA Pipe Length Show to Customer", "FNA Pipe Length Customer",
    "Chamfer", "ØD1", "Adapter Rating",
    "Flow Nozzle Holding Ring Material", "Flow Nozzle Material",
    "Nipple Material", "Nipple Size", "Nipple Schedule", "Nipple Quantity",
    "Venturi Throat Material", "Cyllinder Material", "Cone Material",
    "Piezometer Ring Material", "End Flange Standard", "End Flange Type",
    "End Flange Material", "End Flange Rating", "No of End Flange",
    "Adapter Material", "No. of Tapping 1", "Tap Size 1",
    "No. of Tapping 2", "Tap Size 2", "Companion Flange Required",
    "Tapping Flange", "Tapping Flange Size", "PitotTube Type",
    "Duct Inside Width", "Duct Outside Width", "Duct Inside Height",
    "Duct Outside Height", "PitotTube Probe Material", "Pitot Tube Type",
    "Pitot End Support", "Clamping Condition",
    "Pitot Tube End Connection Material", "Pitot Tube Sleeve Material",
]

# ── 2. Aliases: PDF label variants → canonical Excel column name ───────────
# Keys are normalised (lowercase, no special chars).
# Add more as you discover new PDF variants.
ALIASES: dict[str, str] = {
    # TagNo
    "tagno": "TagNo", "tag": "TagNo", "tagnumber": "TagNo",
    "tag no": "TagNo", "tag number": "TagNo", "instrumenttag": "TagNo",

    # Item Name
    "itemname": "Item Name", "item": "Item Name", "description": "Item Name",
    "itemdesc": "Item Name", "productname": "Item Name",

    # Quantity
    "qty": "Quantity", "nos": "Quantity", "quantity": "Quantity",
    "number": "Quantity", "count": "Quantity",

    # Project
    "projectno": "Project", "projectnumber": "Project", "jobno": "Project",

    # Plate MATERIAL
    "platematerial": "Plate MATERIAL", "plate material": "Plate MATERIAL",
    "plate mat": "Plate MATERIAL", "orifice material": "Plate MATERIAL",
    "disc material": "Plate MATERIAL",

    # Flange MATERIAL
    "flangematerial": "Flange MATERIAL", "flange material": "Flange MATERIAL",
    "flange mat": "Flange MATERIAL",

    # Flange Type
    "flangetype": "Flange Type", "flange type": "Flange Type",
    "facing": "Flange Type",

    # Line_Size
    "linesize": "Line_Size", "line size": "Line_Size",
    "pipe size": "Line_Size", "pipesize": "Line_Size",
    "nominal pipe size": "Line_Size", "nps": "Line_Size",
    "pipe diameter": "Line_Size",

    # Size in NPS OR DN
    "sizeindn": "Size in NPS OR DN", "size": "Size in NPS OR DN",
    "dn": "Size in NPS OR DN", "nominal diameter": "Size in NPS OR DN",

    # Flange Schedule
    "flangeschedule": "Flange Schedule", "schedule": "Flange Schedule",
    "sch": "Flange Schedule", "pipe schedule": "Flange Schedule",

    # Pipe Wall Thk
    "pipewallthk": "Pipe Wall Thk", "wall thickness": "Pipe Wall Thk",
    "wallthk": "Pipe Wall Thk", "pipe wall thickness": "Pipe Wall Thk",
    "thk": "Pipe Wall Thk",

    # Rating
    "rating": "Rating", "pressure rating": "Rating",
    "class": "Rating", "flange rating": "Rating",

    # Drain/Vent
    "drainvent": "Drain/Vent", "drain": "Drain/Vent",
    "vent": "Drain/Vent", "drain vent": "Drain/Vent",

    # Fluid/Service_Name
    "fluidservicename": "Fluid/Service_Name", "fluid": "Fluid/Service_Name",
    "service": "Fluid/Service_Name", "fluidname": "Fluid/Service_Name",
    "medium": "Fluid/Service_Name", "process fluid": "Fluid/Service_Name",

    # Service Discription
    "servicedescription": "Service Discription",
    "service description": "Service Discription",
    "servicediscription": "Service Discription",

    # Service Type
    "servicetype": "Service Type", "phase": "Service Type",
    "fluid phase": "Service Type", "state": "Service Type",

    # Calculation Standard
    "calculationstandard": "Calculation Standard",
    "calc standard": "Calculation Standard",
    "standard": "Calculation Standard", "code": "Calculation Standard",

    # Flow Element Type
    "flowelementtype": "Flow Element Type",
    "element type": "Flow Element Type",
    "meter type": "Flow Element Type", "type": "Flow Element Type",

    # Flow Rate Unit
    "flowrateunit": "Flow Rate Unit", "flow unit": "Flow Rate Unit",
    "units": "Flow Rate Unit",

    # Flow Rate Minimum
    "flowrateminimum": "Flow Rate Minimum", "min flow": "Flow Rate Minimum",
    "qmin": "Flow Rate Minimum", "minimum flow": "Flow Rate Minimum",

    # Flow Rate Maximum
    "flowratemaximum": "Flow Rate Maximum", "max flow": "Flow Rate Maximum",
    "qmax": "Flow Rate Maximum", "maximum flow": "Flow Rate Maximum",

    # Flow Rate Normal
    "flowratenormal": "Flow Rate Normal", "normal flow": "Flow Rate Normal",
    "qnormal": "Flow Rate Normal", "qn": "Flow Rate Normal",

    # Flow Rate at FullScale
    "flowrateatfullscale": "Flow Rate at FullScale",
    "full scale flow": "Flow Rate at FullScale",
    "qfullscale": "Flow Rate at FullScale",

    # Tapping
    "tapping": "Tapping", "tap type": "Tapping",
    "tapping type": "Tapping",

    # Pipe Material
    "pipematerial": "Pipe Material", "pipe mat": "Pipe Material",
    "run pipe material": "Pipe Material",

    # Pressure Unit
    "pressureunit": "Pressure Unit", "pressure units": "Pressure Unit",

    # Temp Unit
    "tempunit": "Temp Unit", "temperature unit": "Temp Unit",
    "temp units": "Temp Unit",

    # Upstream Pressure
    "upstreampressure": "Upstream Pressure",
    "operating pressure": "Upstream Pressure",
    "inlet pressure": "Upstream Pressure", "p1": "Upstream Pressure",

    # Operating_Temperature
    "operatingtemperature": "Operating_Temperature",
    "operating temp": "Operating_Temperature",
    "process temperature": "Operating_Temperature",
    "temperature": "Operating_Temperature", "temp": "Operating_Temperature",

    # Vapour_Pressure
    "vapourpressure": "Vapour_Pressure", "vapor pressure": "Vapour_Pressure",
    "vp": "Vapour_Pressure",

    # DP Unit
    "dpunit": "DP Unit", "differential pressure unit": "DP Unit",

    # DP at Full Scale
    "dpatfullscale": "DP at Full Scale",
    "differential pressure": "DP at Full Scale",
    "dp full scale": "DP at Full Scale", "dp": "DP at Full Scale",

    # Design Pressure
    "designpressure": "Design Pressure", "design press": "Design Pressure",
    "dp design": "Design Pressure",

    # Design Temp
    "designtemp": "Design Temp", "design temperature": "Design Temp",
    "design temp": "Design Temp",

    # Flange Standard
    "flangestandard": "Flange Standard", "flange std": "Flange Standard",
    "standard flange": "Flange Standard",

    # Gasket
    "gasket": "Gasket", "gasket type": "Gasket",
    "gasket material": "Gasket",

    # Stud/Nut
    "studnut": "Stud/Nut", "stud": "Stud/Nut", "nut": "Stud/Nut",
    "bolting": "Stud/Nut", "stud material": "Stud/Nut",

    # IBR
    "ibr": "IBR", "ibr required": "IBR",

    # Nace
    "nace": "Nace", "nace required": "Nace", "nace mr0175": "Nace",

    # Calibration
    "calibration": "Calibration", "calibration required": "Calibration",

    # Freight_Required
    "freightrequired": "Freight_Required",
    "freight": "Freight_Required",

    # OFA Tap Orientation
    "ofataporientation": "OFA Tap Orientation",
    "ofa tap": "OFA Tap Orientation",
    "tap orientation": "OFA Tap Orientation",

    # ItemCode
    "itemcode": "ItemCode", "item code": "ItemCode",
    "part number": "ItemCode", "partno": "ItemCode",

    # No. of Tapping 1 / 2
    "nooftapping1": "No. of Tapping 1", "tapping1": "No. of Tapping 1",
    "nooftapping2": "No. of Tapping 2", "tapping2": "No. of Tapping 2",

    # Tap Size 1 / 2
    "tapsize1": "Tap Size 1", "tap size 1": "Tap Size 1",
    "tapsize2": "Tap Size 2", "tap size 2": "Tap Size 2",

    # Nipple
    "nipplematerial": "Nipple Material",
    "nipplesize": "Nipple Size",
    "nippleschedule": "Nipple Schedule",
    "nipplequantity": "Nipple Quantity",

    # End Flange
    "endflangestandard": "End Flange Standard",
    "end flange standard": "End Flange Standard",
    "endflangematerial": "End Flange Material",
    "endflangetype": "End Flange Type",
    "endflangerating": "End Flange Rating",
    "noofendflange": "No of End Flange",

    # Pitot Tube
    "pitottubetype": "PitotTube Type",
    "pitot tube type": "Pitot Tube Type",
    "pitotprobematerial": "PitotTube Probe Material",
    "pitotendconnectionmaterial": "Pitot Tube End Connection Material",
    "pitottubesleevematerial": "Pitot Tube Sleeve Material",

    # Duct dimensions
    "ductinsidewidth": "Duct Inside Width",
    "ductoutsidewidth": "Duct Outside Width",
    "ductinsideheight": "Duct Inside Height",
    "ductoutsideheight": "Duct Outside Height",
}

# ── 3. Normalisation ───────────────────────────────────────────────────────
_STRIP_RE = re.compile(r"[\s_\-/\\.()]+")

def _norm(text: str) -> str:
    """Lowercase and strip all special/whitespace characters."""
    return _STRIP_RE.sub("", text.lower())


# Pre-build normalised lookup of canonical columns
_NORM_TO_CANONICAL: dict[str, str] = {_norm(c): c for c in EXCEL_COLUMNS}
_NORM_ALIASES: dict[str, str] = {_norm(k): v for k, v in ALIASES.items()}

# Sorted list of normalised canonical names for fuzzy matching
_NORM_CANONICALS = list(_NORM_TO_CANONICAL.keys())

# ── AUTO-ALIASES FROM COLUMN NAMES ─────────────────────────

for col in EXCEL_COLUMNS:
    norm_col = _norm(col)

    # Only add if not already manually defined
    if norm_col not in _NORM_ALIASES:
        _NORM_ALIASES[norm_col] = col

# ── 4. Public API ──────────────────────────────────────────────────────────
def resolve_column(pdf_label: str, fuzzy_threshold: float = 0.68) -> Optional[str]:
    """
    Given a raw label from a PDF, return the matching Excel column name.

    Resolution order:
      1. Exact normalised match against canonical columns
      2. Alias lookup
      3. Fuzzy match (difflib SequenceMatcher, threshold-controlled)
      4. None  →  field is dropped / logged as unmatched

    Args:
        pdf_label:        Raw string found in PDF (e.g. "Tag No.", "Pipe Sch.")
        fuzzy_threshold:  Minimum similarity ratio (0–1). 0.72 is recommended.
                          Lower = more permissive but more false matches.

    Returns:
        Canonical Excel column name, or None if no confident match found.
    """
    if not pdf_label or not isinstance(pdf_label, str):
        return None

    n = _norm(pdf_label)
    if not n:
        return None

    # 1. Exact normalised match
    if n in _NORM_TO_CANONICAL:
        return _NORM_TO_CANONICAL[n]

    # 2. Alias lookup
    if n in _NORM_ALIASES:
        return _NORM_ALIASES[n]

    # 3. Fuzzy match
    matches = difflib.get_close_matches(n, _NORM_CANONICALS, n=1, cutoff=fuzzy_threshold)
    if matches:
        return _NORM_TO_CANONICAL[matches[0]]

    return None

def _split_value_unit(value: str):
    """
    Splits value and unit if combined.
    Example:
        '12.5 kg/cm2' → ('12.5', 'kg/cm2')
    """
    import re

    if not isinstance(value, str):
        return value, None

    match = re.match(r"([\d.]+)\s*([a-zA-Z/°%]+)?", value)

    if match:
        val = match.group(1)
        unit = match.group(2)
        return val, unit

    return value, None

def map_pdf_record(raw: dict, fuzzy_threshold: float = 0.72) -> dict:
    """
    Convert a raw dict from the PDF extractor (label → value) into a row
    dict with canonical Excel column names as keys.

    Unmatched keys are stored under '__unmatched__' for logging.

    Example:
        raw = {"Tag No.": "303-FE-0701", "Pipe Sch.": "Standard", ...}
        → {"TagNo": "303-FE-0701", "Flange Schedule": "Standard", ...}
    """
    row: dict = {}
    unmatched: list = []

    for pdf_label, value in raw.items():
        col = resolve_column(pdf_label, fuzzy_threshold)

    if col:
        val, unit = _split_value_unit(value)

        row[col] = val

        # Auto-fill unit columns
        if "Pressure" in col and unit:
            row["Pressure Unit"] = unit

        elif "Temp" in col and unit:
            row["Temp Unit"] = unit

        elif "Flow Rate" in col and unit:
            row["Flow Rate Unit"] = unit

        elif "Density" in col and unit:
            row["Density Unit"] = unit

        elif "Viscosity" in col and unit:
            row["Viscosity Unit"] = unit
        
    if unmatched:
        row["__unmatched__"] = unmatched   # caller should log these

    return row


def build_excel_row(mapped: dict) -> dict:
    """
    Take a mapped dict (canonical keys) and return a full 129-column row dict
    with NaN for any column not present in the mapped data.
    Strips the '__unmatched__' key before returning.
    """
    import math
    mapped.pop("__unmatched__", None)
    return {col: mapped.get(col, float("nan")) for col in EXCEL_COLUMNS}