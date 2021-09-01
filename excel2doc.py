from docxtpl import DocxTemplate;
import os
from xlrd import open_workbook
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from docx.shared import Cm, Inches, Pt


wb = open_workbook('format.xls')
for s in wb.sheets():
    docvalues = []
    for row in range(s.nrows):
        col_value = []
        for col in range(s.ncols):
            value  = (s.cell(row,col).value)
            try : value = str((value))
            except : pass
            # print(value)
            if len(value) > 1 and value[0] != 'Q' and value[2] == '0':
                value = value[0]
                col_value.append(value)
            else:
                col_value.append(value)
        docvalues.append(col_value)
# print(docvalues)

# context = {}
# for item in docvalues:
#     if item[0] != "table":
#         applicationName = item[2]
#         context["applicationName"] = applicationName
#         applicationPrefix = item[3]
#         context["applicationPrefix"] = applicationPrefix
#         documentId = item[4]
#         context["documentId"] = documentId
#         documentTitle = item[5]
#         context["documentTitle"] = documentTitle
#         documentRevision = item[6]
#         context["documentRevision"] = documentRevision
#     elif item[0] == "table":
#         functionId = item[1]
#         description = item[2]
#         priority = item[3]
#         criticality = item[4]
testData = [
    ['1', 'a', 'b' ,'c'],
    ['1', 'd','e','f'],
    ['2', 'h','i','j'],
    ['2', 'k','l','m'],
    ['3', 'O','P','Q'],
    ['3', 'r','s','t'],
    ['3', 'u','v','w'],
    ['4', 'x','y','z'],
]

# section = int(section)
length_of_excel = docvalues[len(docvalues) - 1][0]


global1 = []
section_dict = {}

section_array = []  
section = 1
i = 0;
for item in docvalues:
    # print(item)
    if item[0] == str(section):
        section_array.append(item)    
    else:
        section_dict[section] = section_array
        section_array = []
        section_array.append(item)
        section = section + 1
        section_dict[section] = section_array

print("dict", section_dict[4])
# global1.append(section_arry)

    # else:      

# print("yahoo", len(section_dict))
# for key in section_dict:  
# print("2",len(section_dict))
    # else:
        # print("pre clear ", section_array)
        # section_dict[section] = section_array
        # section_array.clear()
        # section = section + 1
            
    # print(section_array)    
# print("yahoo", global1)   
# dict = {}

# i = 0
# while i < len(docvalues):

#     entry = docvalues[i]
#     if entry[0] != "table":
#         dict[entry[0] + "-" + str(i)] = entry[1]
#         i = i + 1
#     elif entry[0] == "table":
#         dict[entry[1]] = entry[1:5]
#         i = i + 1

tpl = DocxTemplate("work_template/workTemp.docx")
entries = []
for key in section_dict:
    item = section_dict[key]
    # print("yes", item[0][2])
    tpl.add_heading(item[0][2], level=1)
    i = 0
    while i < len(item):
        if len(item[i][1]) == 3:
            tpl.add_heading(item[0][2], level=2)
        elif item[i][1] == "text":
            tpl.add_paragraph(item[i][2])
        elif item[i][1] == "table":
            # table = tpl.add_table(1,4)
            k = i
            j = len(item) - k
            print(j)
            while k < k + j: 
                if item[k][1] == "table":
                    print(item[k][1])
                # entries.append(item[i])
                k = k + 1
            # print(key ,item[i])
        i  = i + 1
    # print(key)

# print("ent", entries)

for entry in entries:
    print("ent", entry[2:6])
    table = tpl.add_table(1,4)
    new_row = table.add_row().cells
    # new_row[0].text = entry[0]
    # new_row[1].text = entry[1]
    # new_row[2].text = entry[2]
    # new_row[3].text = entry[3]
# tpl.save('workTemp_result.docx')

# os.startfile("C:/Users/Dan/Documents/projects/automated_word_document/workTemp_result.docx")
# entries = []

# for key in dict:
#     test = key.split("-")
#     # print(test)
#     if len(key) >= 3 and len(key) <= 4:
#         tpl.add_heading(dict[key], level=1)
#     elif len(key) >= 5 and len(key) <= 6:
#         tpl.add_heading(dict[key], level=2)
#     elif len(key) > 5 and key[0:3] != "FS-":
#         tpl.add_paragraph(dict[key])
#     elif key[0:3] == "FS-":
#         entries.append(dict[key]);
#         # tables_to_make[table_number] = dict[key]
#         # table_number = table_number + 1
#     else:
#         print("'Type' column has been inputted incorrectly")

# print(entries)




# table = tpl.add_table(1,4)
# table.style = "TableGrid"

# table.rows[0].cells[0].text = "Function ID"
# table.rows[0].cells[1].text = "Description"
# table.rows[0].cells[2].text = "Priority"
# table.rows[0].cells[3].text = "Criticality"

# row = table.rows[0]
# for cell in row.cells:
#     shading_elm_2 = parse_xml(r'<w:shd {} w:fill="0C3C60"/>'.format(nsdecls('w')))
#     cell._tc.get_or_add_tcPr().append(shading_elm_2)





# i = 0
# while i < 4:
#     for cell in table.columns[i].cells:
#         cell.font = Pt(10)
#         if i == 1:
#                 cell.width = Inches(4)
#         else:
#                 cell.width = Inches(1.2)
#     i = i + 1


# for row in table.rows:
#     for cell in row.cells:
#         paragraphs = cell.paragraphs
#         for paragraph in paragraphs:
#             for run in paragraph.runs:
#                 font = run.font
#             font.size= Pt(10)

# # tpl.render(context)
# tpl.save('workTemp_result.docx')

# os.startfile("C:/Users/Dan/Documents/projects/automated_word_document/workTemp_result.docx")