"""
Injects real Excel Slicer objects into a workbook's raw XML parts, bound
to the reporting_month column of each worksheet's Table.

openpyxl has no support for writing the Slicer/SlicerCache/Drawing XML
parts a slicer needs - slicers are an OOXML extension Microsoft added
after ECMA-376 was ratified, and openpyxl's maintainers have explicitly
declined to add write support for them (they're stripped from any file
openpyxl re-saves). See:
  https://github.com/chronossc/openpyxl/issues/17
  https://github.com/chronossc/openpyxl/issues/30
  https://openpyxl.readthedocs.io/en/stable/filters.html

Since the parts are just more XML in the .xlsx zip, this module builds
them by hand - following the same zipfile-patch pattern the pipeline
already uses for <ignoredErrors> - modeled directly on a real
Excel-authored workbook containing working slicers
(reference/Backdating POs 7.1.25-7.1.26.xlsx) so the output matches a
proven-valid structure rather than a guess at the schema.

Each worksheet passed in must already have an openpyxl Table
(openpyxl.worksheet.table.Table, added via ws.add_table()) and the
workbook must already be saved - this module only adds the slicer
machinery on top of an existing, already-saved Table.

Note: because this slicer cache is written directly rather than
computed by a live Excel session (which normally caches a table
slicer's member list the moment a person inserts it via the UI),
Excel shows an "Out of date" banner with an UPDATE button the first
time the file is opened. Clicking UPDATE once computes the cache and
the slicer is then fully functional - this is a one-time, harmless
side effect of hand-authoring the XML, not a functional defect.
"""
import re
import uuid
from dataclasses import dataclass

CONTENT_TYPES_PATH = "[Content_Types].xml"
WORKBOOK_PATH = "xl/workbook.xml"
WORKBOOK_RELS_PATH = "xl/_rels/workbook.xml.rels"

# Fixed OOXML extension URIs (not something we invent - these identify
# the extension type to Excel and must match exactly).
SLICER_LIST_EXT_URI = "{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}"
SLICER_CACHES_EXT_URI = "{46BE6895-7355-4a93-B00E-2C351335B9C9}"
TABLE_SLICER_CACHE_EXT_URI = "{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}"


@dataclass
class SheetSlicerTarget:
    sheet_path: str  # e.g. "xl/worksheets/sheet1.xml"
    table_id: int  # matches the id given to the sheet's openpyxl Table
    slicer_index: int  # 1-based, unique across the workbook
    reporting_month_column: int  # 1-based table column index of reporting_month
    num_columns: int  # total table columns; used to anchor the slicer widget


def _guid():
    return f"{{{str(uuid.uuid4()).upper()}}}"


def _rels_path_for(sheet_path):
    folder, filename = sheet_path.rsplit("/", 1)
    return f"{folder}/_rels/{filename}.rels"


def _next_rid(rels_xml):
    ids = [int(n) for n in re.findall(r'Id="rId(\d+)"', rels_xml)]
    return max(ids, default=0) + 1


def _add_relationship(rels_xml, rid, rel_type, target):
    tag = f'<Relationship Id="rId{rid}" Type="{rel_type}" Target="{target}"/>'
    return rels_xml.replace("</Relationships>", tag + "</Relationships>")


def _empty_rels_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        "</Relationships>"
    )


def _slicer_cache_xml(name, source_name, table_id, column):
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<slicerCacheDefinition '
        'xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" '
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x xr10" '
        'xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10" '
        f'name="{name}" xr10:uid="{_guid()}" sourceName="{source_name}">'
        f'<extLst><x:ext uri="{TABLE_SLICER_CACHE_EXT_URI}" '
        'xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">'
        f'<x15:tableSlicerCache tableId="{table_id}" column="{column}"/>'
        "</x:ext></extLst></slicerCacheDefinition>"
    )


def _slicer_xml(name, cache_name, caption):
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<slicers xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" '
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x xr10" '
        'xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10">'
        f'<slicer name="{name}" xr10:uid="{_guid()}" cache="{cache_name}" '
        f'caption="{caption}" rowHeight="241300"/>'
        "</slicers>"
    )


def _drawing_xml(slicer_name, anchor_col):
    from_col, to_col, to_col_off = anchor_col, anchor_col + 3, 85725
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<xdr:twoCellAnchor editAs="absolute">'
        f"<xdr:from><xdr:col>{from_col}</xdr:col><xdr:colOff>0</xdr:colOff>"
        "<xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"
        f"<xdr:to><xdr:col>{to_col}</xdr:col><xdr:colOff>{to_col_off}</xdr:colOff>"
        "<xdr:row>13</xdr:row><xdr:rowOff>47625</xdr:rowOff></xdr:to>"
        '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
        '<mc:Choice xmlns:sle15="http://schemas.microsoft.com/office/drawing/2012/slicer" Requires="sle15">'
        '<xdr:graphicFrame macro="">'
        f'<xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="{slicer_name}"/>'
        "<xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>"
        '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
        '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2010/slicer">'
        f'<sle:slicer xmlns:sle="http://schemas.microsoft.com/office/drawing/2010/slicer" name="{slicer_name}"/>'
        "</a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice>"
        '<mc:Fallback><xdr:sp macro="" textlink="">'
        '<xdr:nvSpPr><xdr:cNvPr id="0" name=""/><xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr></xdr:nvSpPr>'
        '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="2524125"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:solidFill><a:prstClr val="white"/></a:solidFill>'
        '<a:ln w="1"><a:solidFill><a:prstClr val="green"/></a:solidFill></a:ln></xdr:spPr>'
        '<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip"/><a:lstStyle/><a:p><a:r>'
        '<a:rPr lang="en-US" sz="1100"/>'
        "<a:t>This shape represents a table slicer. Table slicers are not supported in this "
        "version of Excel.</a:t>"
        "</a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback></mc:AlternateContent>"
        "<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"
    )


def add_reporting_month_slicers(contents, targets):
    """Return a copy of `contents` (dict[str, bytes] of zip part -> bytes,
    as already saved by openpyxl) with a reporting_month slicer added for
    every worksheet in `targets`."""
    contents = dict(contents)

    content_types_xml = contents[CONTENT_TYPES_PATH].decode("utf-8")
    workbook_xml = contents[WORKBOOK_PATH].decode("utf-8")
    workbook_rels_xml = contents[WORKBOOK_RELS_PATH].decode("utf-8")

    new_overrides = []
    slicer_cache_rids = []  # [(workbook.xml.rels rId, cache name), ...]

    for target in targets:
        n = target.slicer_index
        cache_name = "Slicer_Reporting_Month" if n == 1 else f"Slicer_Reporting_Month{n - 1}"
        slicer_name = "Reporting Month" if n == 1 else f"Reporting Month {n - 1}"

        cache_path = f"xl/slicerCaches/slicerCache{n}.xml"
        slicer_path = f"xl/slicers/slicer{n}.xml"
        drawing_path = f"xl/drawings/drawing{n}.xml"

        contents[cache_path] = _slicer_cache_xml(
            cache_name, "Reporting Month", target.table_id, target.reporting_month_column
        ).encode("utf-8")
        contents[slicer_path] = _slicer_xml(slicer_name, cache_name, "Reporting Month").encode("utf-8")
        contents[drawing_path] = _drawing_xml(slicer_name, target.num_columns).encode("utf-8")

        new_overrides.append((cache_path, "application/vnd.ms-excel.slicerCache+xml"))
        new_overrides.append((slicer_path, "application/vnd.ms-excel.slicer+xml"))
        new_overrides.append((drawing_path, "application/vnd.openxmlformats-officedocument.drawing+xml"))

        wb_rid = _next_rid(workbook_rels_xml)
        workbook_rels_xml = _add_relationship(
            workbook_rels_xml,
            wb_rid,
            "http://schemas.microsoft.com/office/2007/relationships/slicerCache",
            f"slicerCaches/slicerCache{n}.xml",
        )
        slicer_cache_rids.append((wb_rid, cache_name))

        rels_path = _rels_path_for(target.sheet_path)
        sheet_rels_xml = contents[rels_path].decode("utf-8") if rels_path in contents else _empty_rels_xml()
        drawing_rid = _next_rid(sheet_rels_xml)
        sheet_rels_xml = _add_relationship(
            sheet_rels_xml,
            drawing_rid,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
            f"../drawings/drawing{n}.xml",
        )
        slicer_rid = _next_rid(sheet_rels_xml)
        sheet_rels_xml = _add_relationship(
            sheet_rels_xml,
            slicer_rid,
            "http://schemas.microsoft.com/office/2007/relationships/slicer",
            f"../slicers/slicer{n}.xml",
        )
        contents[rels_path] = sheet_rels_xml.encode("utf-8")

        # Worksheet element order (ECMA-376): ignoredErrors, drawing,
        # tableParts, extLst - openpyxl already wrote <tableParts> in the
        # right place, so anchoring on it keeps <drawing> correctly ordered
        # before it and lets build_report.py's own ignoredErrors patch
        # (applied via the same anchor) land before <drawing> too.
        # openpyxl's own root <worksheet> element doesn't declare xmlns:r
        # (it's declared locally on individual elements like <tablePart>
        # instead), so every r:id use we add needs its own xmlns:r too.
        sheet_xml = contents[target.sheet_path].decode("utf-8")
        r_ns = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
        drawing_tag = f'<drawing {r_ns} r:id="rId{drawing_rid}"/>'
        slicer_list_ext = (
            f'<extLst><ext uri="{SLICER_LIST_EXT_URI}" '
            'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">'
            f'<x14:slicerList><x14:slicer {r_ns} r:id="rId{slicer_rid}"/></x14:slicerList>'
            "</ext></extLst>"
        )
        if "<tableParts" in sheet_xml:
            sheet_xml = sheet_xml.replace("<tableParts", drawing_tag + "<tableParts", 1)
        else:
            sheet_xml = sheet_xml.replace("</worksheet>", drawing_tag + "</worksheet>")
        sheet_xml = sheet_xml.replace("</worksheet>", slicer_list_ext + "</worksheet>")
        contents[target.sheet_path] = sheet_xml.encode("utf-8")

    overrides_xml = "".join(
        f'<Override PartName="/{path}" ContentType="{ctype}"/>' for path, ctype in new_overrides
    )
    content_types_xml = content_types_xml.replace("</Types>", overrides_xml + "</Types>")

    defined_names_xml = "".join(
        f'<definedName name="{name}">#N/A</definedName>' for _, name in slicer_cache_rids
    )
    if re.search(r"<definedNames\s*/>", workbook_xml):
        workbook_xml = re.sub(
            r"<definedNames\s*/>", f"<definedNames>{defined_names_xml}</definedNames>", workbook_xml
        )
    elif "<definedNames>" in workbook_xml:
        workbook_xml = workbook_xml.replace("</definedNames>", defined_names_xml + "</definedNames>")
    else:
        workbook_xml = workbook_xml.replace(
            "</sheets>", f"</sheets><definedNames>{defined_names_xml}</definedNames>"
        )

    slicer_caches_refs = "".join(f'<x14:slicerCache r:id="rId{rid}"/>' for rid, _ in slicer_cache_rids)
    ext_block = (
        f'<ext uri="{SLICER_CACHES_EXT_URI}" '
        'xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">'
        '<x15:slicerCaches xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">'
        f"{slicer_caches_refs}</x15:slicerCaches></ext>"
    )
    if "<extLst>" in workbook_xml:
        workbook_xml = workbook_xml.replace("</extLst>", ext_block + "</extLst>")
    else:
        workbook_xml = workbook_xml.replace("</workbook>", f"<extLst>{ext_block}</extLst></workbook>")

    contents[CONTENT_TYPES_PATH] = content_types_xml.encode("utf-8")
    contents[WORKBOOK_PATH] = workbook_xml.encode("utf-8")
    contents[WORKBOOK_RELS_PATH] = workbook_rels_xml.encode("utf-8")

    return contents
