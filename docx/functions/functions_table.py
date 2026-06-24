from docx.oxml import OxmlElement                   # this and next one for some magic table cells border manipulations
from docx.oxml.ns import qn                         # to add bottom border to cells.
from docx.shared import Pt                          # for setting font sizes
from docx.enum.text import WD_ALIGN_PARAGRAPH       # for setting alignments
from docx.shared import Cm                          # for changing paper size and modifying page margins

# Source - https://stackoverflow.com/a/49615968
# Posted by MadisonTrash, modified by community. See post 'Timeline' for change history
# Retrieved 2026-06-24, License - CC BY-SA 4.0
def set_cell_border(cell, **kwargs):
    """
    Set cell`s border
    Usage:

    set_cell_border(
        cell,
        top={"sz": 12, "val": "single", "color": "#FF0000", "space": "0"},
        bottom={"sz": 12, "color": "#00FF00", "val": "single"},
        start={"sz": 24, "val": "dashed", "shadow": "true"},
        end={"sz": 12, "val": "dashed"},
    )
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    # check for tag existnace, if none found, then create one
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)

    # list over all available tags #only interested in bottom borders
    edge = "bottom"
    edge_data = kwargs.get(edge)
    if edge_data:
        tag = 'w:{}'.format(edge)
        # check for tag existnace, if none found, then create one
        element = tcBorders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            tcBorders.append(element)
        # looks like order of attributes is important
        for key in ["sz", "val", "color", "space", "shadow"]:
            if key in edge_data:
                element.set(qn('w:{}'.format(key)), str(edge_data[key]))



def add_text_to_main_table(table_cell, text_inside_the_cell: str, text_alignment):
    ''' Formats given text and puts it into given table cell of main table. TextFormat: [Times New Roman, 12Pt, 1.15↕]. '''
    table_cell.text = ""
    table_cell.paragraphs[0].add_run(
        f"{text_inside_the_cell}",
        style = "TNR12PtStyle"
    )
    table_cell.paragraphs[0].alignment = text_alignment
    table_cell.paragraphs[0].paragraph_format.line_spacing = 1.15
    table_cell.paragraphs[0].paragraph_format.space_after = Pt(0)



def fill_in_first_row(table):
    ''' Fills in the first row of table and configures cells '''
    first_row = table.rows[0].cells
    add_text_to_main_table(first_row[0], "", WD_ALIGN_PARAGRAPH.LEFT)
    add_text_to_main_table(first_row[1], "", WD_ALIGN_PARAGRAPH.LEFT)
    add_text_to_main_table(first_row[2], "", WD_ALIGN_PARAGRAPH.LEFT)
    add_text_to_main_table(first_row[3], "", WD_ALIGN_PARAGRAPH.LEFT)
    add_text_to_main_table(first_row[4], "№", WD_ALIGN_PARAGRAPH.RIGHT)
    add_text_to_main_table(first_row[5], "б/н", WD_ALIGN_PARAGRAPH.LEFT)
    first_row[0].width = Cm(3.25)
    first_row[1].width = Cm(3.25)
    first_row[2].width = Cm(5.27)
    first_row[3].width = Cm(0.5)
    first_row[4].width = Cm(2.23)
    first_row[5].width = Cm(3.5)
    set_cell_border(first_row[0], bottom={"sz": 4, "color": "#000000", "val": "single"})
    set_cell_border(first_row[5], bottom={"sz": 4, "color": "#000000", "val": "single"})



def fill_in_second_row(table):
    ''' Fills in the second row of table and configures cells '''
    second_row = table.rows[1].cells
    add_text_to_main_table(second_row[0], "\nУважаемый (-ая) Елизавета Сергеевна!", WD_ALIGN_PARAGRAPH.CENTER)
    second_row_012345_merged = second_row[0].merge(second_row[5])
    second_row[0].paragraphs[0].paragraph_format.space_after = Pt(4)



def fill_in_third_row(table):
    ''' Fills in the third row of table and configures cells '''
    third_row = table.rows[2].cells
    add_text_to_main_table(
        third_row[0],
        "В соответствии со статьей 123 Трудового кодекса Российской Федерации, уведомляем Вас, " \
        "что по утвержденному графику отпусков ИППИ РАН на 2026 год Вам будет предоставлен " \
        "ежегодный основной оплачиваемый отпуск с",
        WD_ALIGN_PARAGRAPH.JUSTIFY
    )
    third_row[0].paragraphs[0].add_run(
        " 06.07.2026 ",
        style = "TNR12PtStyle"
    ).bold = True
    third_row[0].paragraphs[0].add_run(
        "года продолжительностью 10 календарных дней.",
        style = "TNR12PtStyle"
    )
    third_row_012345_merged = third_row[0].merge(third_row[5])
    third_row_012345_merged.paragraphs[0].paragraph_format.first_line_indent = Cm(1.3)
    third_row[0].paragraphs[0].paragraph_format.space_before = Pt(12)



def fill_in_fourth_row(table):
    ''' Fills in the fourth row of table and configures cells '''
    fourth_row = table.rows[3].cells
    add_text_to_main_table(
        fourth_row[0],
        "Просим Вас сообщить в отдел кадров о Вашем решении использовать отпуск, или " \
        "перенести его на другие даты в течение 3 рабочих дней со дня получения настоящего " \
        "уведомления. В соответствии со статьей 124 Трудового кодекса Российской Федерации, по " \
        "согласованию с работодателем ежегодный оплачиваемый отпуск может быть перенесен на другой " \
        "период по уважительной причине. В этом случае необходимо предоставить в отдел кадров " \
        "заявление, согласованное с непосредственным руководителем и директором (заместителем " \
        "директора) ИППИ РАН.",
        WD_ALIGN_PARAGRAPH.JUSTIFY
    )
    fourth_row[0].add_paragraph("")
    fourth_row[0].paragraphs[1].paragraph_format.space_after = Pt(0)
    fourth_row[0].paragraphs[1].alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    fourth_row[0].paragraphs[1].add_run(
        "Желаем Вам приятного отпуска!",
        style = "TNR12PtStyle"
    )
    fourth_row_012345_merged = fourth_row[0].merge(fourth_row[5])
    fourth_row_012345_merged.paragraphs[0].paragraph_format.first_line_indent = Cm(1.3)
    fourth_row_012345_merged.paragraphs[1].paragraph_format.first_line_indent = Cm(1.3)



def fill_in_fifth_row(table):
    ''' Fills in the fifth row of table and configures cells '''
    fifth_row = table.rows[4].cells
    fifth_row_012345_merged = fifth_row[0].merge(fifth_row[5])
    add_text_to_main_table(fifth_row[0], "", WD_ALIGN_PARAGRAPH.LEFT)



def fill_in_sixth_row(table):
    ''' Fills in the sixth row of table and configures cells '''
    sixth_row = table.rows[5].cells
    sixth_row[0].width = Cm(3.25)
    sixth_row[1].width = Cm(2.5)
    sixth_row[2].width = Cm(6)
    sixth_row[3].width = Cm(0.5)
    sixth_row[4].width = Cm(2.23)
    sixth_row[5].width = Cm(3.5)
    sixth_row_45_merged = sixth_row[4].merge(sixth_row[5])
    sixth_row_01_merged = sixth_row[0].merge(sixth_row[1])
    add_text_to_main_table(sixth_row[0], "Ведущий специалист отдела кадров", WD_ALIGN_PARAGRAPH.LEFT)
    add_text_to_main_table(sixth_row[1], "", WD_ALIGN_PARAGRAPH.LEFT)
    add_text_to_main_table(sixth_row[2], "", WD_ALIGN_PARAGRAPH.LEFT)
    add_text_to_main_table(sixth_row[3], "", WD_ALIGN_PARAGRAPH.LEFT)
    add_text_to_main_table(sixth_row[4], "Е.С. Борисова", WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_border(sixth_row[2], bottom={"sz": 4, "color": "#000000", "val": "single"})
    set_cell_border(sixth_row_45_merged, bottom={"sz": 4, "color": "#000000", "val": "single"})
    sixth_row[0].paragraphs[0].paragraph_format.space_before = Pt(0)
    sixth_row[0].paragraphs[0].paragraph_format.space_after = Pt(2)



def fill_in_seventh_row(table):
    ''' Fills in the seventh row of table and configures cells '''
    seventh_row = table.rows[6].cells
    seventh_row_012345_merged = seventh_row[0].merge(seventh_row[5])
    add_text_to_main_table(seventh_row[0], "", WD_ALIGN_PARAGRAPH.LEFT)



def fill_in_eigth_row(table):
    ''' Fills in the eigth row of table and configures cells '''
    eigth_row = table.rows[7].cells
    eigth_row[0].width = Cm(3.25)
    eigth_row[1].width = Cm(2.5)
    eigth_row[2].width = Cm(6)
    eigth_row[3].width = Cm(0.5)
    eigth_row[4].width = Cm(2.23)
    eigth_row[5].width = Cm(3.5)
    eigth_row_01_merged = eigth_row[0].merge(eigth_row[1])
    eigth_row_45_merged = eigth_row[4].merge(eigth_row[5])
    add_text_to_main_table(eigth_row[0], "С уведомлением ознакомлен ", WD_ALIGN_PARAGRAPH.LEFT)
    add_text_to_main_table(eigth_row[1], "", WD_ALIGN_PARAGRAPH.LEFT)
    add_text_to_main_table(eigth_row[2], "", WD_ALIGN_PARAGRAPH.LEFT)
    add_text_to_main_table(eigth_row[3], "", WD_ALIGN_PARAGRAPH.LEFT)
    add_text_to_main_table(eigth_row[4], "", WD_ALIGN_PARAGRAPH.LEFT)    
    set_cell_border(eigth_row[2], bottom={"sz": 4, "color": "#000000", "val": "single"})
    set_cell_border(eigth_row_45_merged, bottom={"sz": 4, "color": "#000000", "val": "single"})
    eigth_row[0].paragraphs[0].paragraph_format.space_after = Pt(0)




def fill_in_nineth_row(table):
    ''' Fills in the nineth row of table and configures cells '''
    nineth_row = table.rows[8].cells
    nineth_row[0].width = Cm(3.25)
    nineth_row[1].width = Cm(2.5)
    nineth_row[2].width = Cm(6)
    nineth_row[3].width = Cm(0.5)
    nineth_row[4].width = Cm(2.23)
    nineth_row[5].width = Cm(3.5)
    nineth_row_45_merged = nineth_row[4].merge(nineth_row[5])
    add_text_to_main_table(nineth_row[0], "", WD_ALIGN_PARAGRAPH.LEFT)
    add_text_to_main_table(nineth_row[1], "", WD_ALIGN_PARAGRAPH.LEFT)
    add_text_to_main_table(nineth_row[2], "", WD_ALIGN_PARAGRAPH.LEFT)
    add_text_to_main_table(nineth_row[3], "", WD_ALIGN_PARAGRAPH.LEFT)
    add_text_to_main_table(nineth_row[4], "\n«___»_____________202___г.", WD_ALIGN_PARAGRAPH.CENTER)