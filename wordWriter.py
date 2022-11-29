from docx import Document
from docx.shared import Inches

document = Document()

document.add_heading('preacher time table ', 0)


def output_timetable(records, streams):
    table = document.add_table(rows=len(records), cols=4)
    # table heading
    tableHeader = table.rows[0].cells
    tableHeader[0].text = "class"
    tableHeader[1].text = "tuesday"
    tableHeader[2].text = "wednesday"
    tableHeader[3].text = "thursday"
    # set all the classes
    for i in range(1, len(records)):
        cell = table.cell(i, 0)
        cell.text = streams[i]
    for i in range(1, len(records)):
        stream = table.rows[i].cells
        stream[1].text = records[i]["tuesday"]
        stream[2].text = records[i]["wednesday"]
        stream[3].text = records[i]["thursday"]
    table.style = document.styles['Table Grid']
    document.add_page_break()

    document.save('timetable.docx')
