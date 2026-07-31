"""Saves the workbook and patches in the two pieces of XML openpyxl can't
write itself: the "Number Stored as Text" ignoredErrors block, and the
reporting_month Slicer (see excel_slicers.py for why the latter needs
hand-authored XML).
"""
import logging
import re
import zipfile

from ..logging_config import log_stage
from .excel_slicers import SheetSlicerTarget, add_reporting_month_slicers

logger = logging.getLogger(__name__)


@log_stage(logger, "save workbook and patch XML (ignoredErrors, slicers)")
def save_workbook(wb, text_column_letters_by_sheet, slicer_info_by_sheet, dst):
    logger.debug("Writing workbook to %s (%d sheet(s))", dst, len(wb.sheetnames))
    wb.save(dst)

    with zipfile.ZipFile(dst, "r") as zin:
        names = zin.namelist()
        contents = {name: zin.read(name) for name in names}

    workbook_xml = contents["xl/workbook.xml"].decode("utf-8")
    rels_xml = contents["xl/_rels/workbook.xml.rels"].decode("utf-8")

    rid_to_target = {}
    for rel_tag in re.findall(r"<Relationship\b[^>]*/>", rels_xml):
        rid = re.search(r'Id="(rId\d+)"', rel_tag).group(1)
        target = re.search(r'Target="([^"]+)"', rel_tag).group(1)
        rid_to_target[rid] = target

    name_to_rid = {}
    for sheet_tag in re.findall(r"<sheet\b[^>]*/>", workbook_xml):
        name = re.search(r'name="([^"]+)"', sheet_tag).group(1)
        rid = re.search(r'r:id="(rId\d+)"', sheet_tag).group(1)
        name_to_rid[name] = rid

    sheet_path_by_name = {}
    for sheet_name in name_to_rid:
        target = rid_to_target[name_to_rid[sheet_name]].lstrip("/")
        sheet_path_by_name[sheet_name] = target if target.startswith("xl/") else f"xl/{target}"

    for sheet_name, (text_column_letters, last_row) in text_column_letters_by_sheet.items():
        sheet_path = sheet_path_by_name[sheet_name]

        ignored_errors_xml = "<ignoredErrors>" + "".join(
            f'<ignoredError sqref="{letter}2:{letter}{last_row}" numberStoredAsText="1"/>'
            for letter in text_column_letters
        ) + "</ignoredErrors>"

        # <ignoredErrors> must precede <drawing>/<tableParts> per the OOXML
        # worksheet element order, so anchor on <tableParts> (added by
        # ws.add_table() above) rather than blindly prepending </worksheet>.
        sheet_xml = contents[sheet_path].decode("utf-8")
        if "<tableParts" in sheet_xml:
            sheet_xml = sheet_xml.replace("<tableParts", ignored_errors_xml + "<tableParts", 1)
        else:
            sheet_xml = sheet_xml.replace("</worksheet>", ignored_errors_xml + "</worksheet>")
        contents[sheet_path] = sheet_xml.encode("utf-8")

    slicer_targets = [
        SheetSlicerTarget(
            sheet_path=sheet_path_by_name[sheet_name],
            table_id=table_id,
            slicer_index=table_id,
            reporting_month_column=reporting_month_column,
            num_columns=num_columns,
        )
        for sheet_name, (table_id, reporting_month_column, num_columns) in slicer_info_by_sheet.items()
    ]
    contents = add_reporting_month_slicers(contents, slicer_targets)

    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in contents:
            zout.writestr(name, contents[name])
