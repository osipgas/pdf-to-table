from .things import extract_first_page_text, get_departure_and_destination, convert_to_main_table, convert_to_sub_table, modify_excel, show_excel_table, insert_values_into_template, get_names, fill_template
from pathlib import Path
from .extract_table import ExtractTextTableInfoFromPDF
import zipfile
from copy import copy


def extract_tables(pdf_path, extract_to):
    output_path = "pdf_processing/extracted.zip"
    ExtractTextTableInfoFromPDF(pdf_path=pdf_path, output_path=output_path)
    with zipfile.ZipFile(output_path, 'r') as zip_ref:
        zip_ref.extractall(extract_to)


def pdf_to_excel(pdf_path, tables_path, save_path, template_path):
    # processing main table
    table_path = f"{tables_path}/fileoutpart1.xlsx"

    text = extract_first_page_text(pdf_path)
    departure, destination = get_departure_and_destination(text)

    main_df = convert_to_main_table(table_path, departure, destination)


    # processing sub table
    sub_table_path = f"{tables_path}/fileoutpart8.xlsx"
    if not Path(sub_table_path).exists():
        sub_table_path = f"{tables_path}/fileoutpart7.xlsx"

    sub_df = convert_to_sub_table(sub_table_path)

    tables = {"main": main_df, "sub": sub_df}







    # inserting data
    wb = modify_excel(template_path, len(tables["main"]) - 1, "result.xlsx")
    ws = wb.active
    insert_values_into_template(tables["main"], ws)





    departure_name, destination_name = None, None

    try:
        # extracting names
        maybe_departure, maybe_departure_name, maybe_destination, maybe_destination_name = get_names(pdf_path)

        if departure == maybe_departure:
            departure_name = maybe_departure_name

        if destination == maybe_destination:
            destination_name = maybe_destination_name
    except:
        pass




    # inserting special info

    departure_info = {
        "<Name>": departure_name, 
        "<Elevation>": tables["sub"].loc["DEP", "ELEV"], 
        "<ATIS>": tables["sub"].loc["DEP", "WX"],
        "<GND>": tables["sub"].loc["DEP", "GND"],
        "<TWR>": tables["sub"].loc["DEP", "TWR/CTAF"],
        "<RWY>": tables["sub"].loc["DEP", "LONGEST RWY ANGLE"]
    }



    destination_info = {
        "<Name>": destination_name, 
        "<Elevation>": tables["sub"].loc["DEST", "ELEV"], 
        "<ATIS>": tables["sub"].loc["DEST", "WX"],
        "<GND>": tables["sub"].loc["DEST", "GND"],
        "<TWR>": tables["sub"].loc["DEST", "TWR/CTAF"],
        "<RWY>": tables["sub"].loc["DEST", "LONGEST RWY ANGLE"]
    }


    text_template = """Departure (Name: <Name>, Code: ___, Elevation: <Elevation>, QFU°/QFU°: ___, DA: ___, Circuit pattern altitude: ___;
    ATIS: <ATIS>, GND: <GND>, TWR: <TWR>, A/A: _____ Approach: _____;
    RWY: <RWY>, Length: ____ m., Req. Dist.: ____ m., Surface: __________;
    Exp. Wind: ______, Exp. QNH: ____ hpa;
    RWY: ____, Wind: _______, QNH: _______, Squak: _______"""


    departure_text = fill_template(text_template, departure_info)
    destination_text = fill_template(text_template, destination_info)

    ws["C2"].value = departure_text
    ws[f"C{len(tables['main']) * 2}"].value = destination_text




    # formating table
    source_row = 2
    target_row = len(main_df) * 2

    source_height = ws.row_dimensions[source_row].height

    if source_height is not None:
        ws.row_dimensions[target_row].height = source_height

    ws[f"C{target_row}"].alignment = copy(
        ws[f"C{source_row}"].alignment
    )

    ws.unmerge_cells(f"C{len(tables['main']) * 2 + 1}:C{len(tables['main']) * 2 + 2}")
    ws.unmerge_cells(f"D{len(tables['main']) * 2 + 1}:D{len(tables['main']) * 2 + 2}")



    # saving
    wb.save(save_path)