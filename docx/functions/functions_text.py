from docx.shared import Pt                          # for setting font sizes
from docx.enum.text import WD_ALIGN_PARAGRAPH       # for setting alignments
from docx.shared import Cm                          # for changing paper size and modifying page margins


def write_paragraph_organisation_name(document) -> None:
    ''' Writes first paragraph (heading containing info about organistaiotn name) into the given file <type: Document>  '''

    separator_paragraph = document.add_paragraph()
    separator_paragraph.paragraph_format.line_spacing = 1
    separator_paragraph.paragraph_format.space_after = Pt(1)
    separator_paragraph.add_run(
        " ",
        style = "TNR12PtStyle"
    ).bold = True

    # Source - https://stackoverflow.com/a/36409427
    # Posted by Andrei
    # Retrieved 2026-06-23, License - CC BY-SA 3.0
    table = document.add_table(rows=1, cols=1)
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER

    table_cell = table.rows[0].cells
    table_cell[0].text = ""
    table_cell[0].paragraphs[0].add_run(
        "Федеральное государственное бюджетное учреждение науки\nИнститут проблем передачи информации\nим. А.А. Харкевича \nРоссийской академии наук",
        style = "TNR14PtStyle"
    )
    table_cell[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    table_cell[0].paragraphs[0].paragraph_format.line_spacing = 1
    table_cell[0].paragraphs[0].paragraph_format.space_after = Pt(0)



def write_paragraph_recipient_person(document) -> None:
    ''' Writes second paragraph (about the recipient person) into the given file <type: Document>  '''

    separator_paragraph = document.add_paragraph()
    separator_paragraph.paragraph_format.line_spacing = 1
    separator_paragraph.paragraph_format.space_after = Pt(0)
    separator_paragraph.add_run(
        "\n",
        style = "TNR14PtStyle"
    )
    separator_paragraph.paragraph_format.space_after = Pt(0.5)
    separator_paragraph.add_run(
        "\n",
        style = "TNR12PtStyle"
    )
    
    recipient_person_paragraph = document.add_paragraph()
    recipient_person_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    recipient_person_paragraph.paragraph_format.space_after = Pt(0)
    recipient_person_paragraph.paragraph_format.left_indent = Cm(11.25)
    recipient_person_paragraph.add_run(
        "Ведущему специалисту\nОтдела кадров\n",
        style = "TNR14PtStyle"
    )
    recipient_person_paragraph.add_run(
        "Борисовой Е.С.",
        style = "TNR14PtStyle"
    ).bold = True



def write_bold_notification_text(document):
    ''' Writes third paragraph (bold notification title-text) into the given file <type: Document> '''

    separator_paragraph = document.add_paragraph()
    separator_paragraph.paragraph_format.line_spacing = 1.15
    separator_paragraph.paragraph_format.space_after = Pt(0.5)
    separator_paragraph.add_run(
        "\n\n‎",
        style = "TNR14PtStyle"
    )
    
    notification_bold_text_paragraph = document.add_paragraph()
    notification_bold_text_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    notification_bold_text_paragraph.paragraph_format.line_spacing = 1.15
    notification_bold_text_paragraph.paragraph_format.space_after = Pt(0)
    notification_bold_text_paragraph.add_run(
        "Уведомление\nо предоставлении ежегодного оплачиваемого отпуска\n‎",
        style = "TNR14PtStyle"
    ).bold = True