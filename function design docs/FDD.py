from docxtpl import DocxTemplate;
from xlrd import open_workbook

wb = open_workbook('FDD-format.xls')
for s in wb.sheets():
    docvalues = []
    for row in range(s.nrows):
        col_value = []
        for col in range(s.ncols):
            value  = (s.cell(row,col).value) 
            try : value = str((value))
            except : pass
            col_value.append(value)
        docvalues.append(col_value)

docvalues.pop(0)
context = {}
i = 0
for row in docvalues:
        function_reference = row[0]
        context["function_reference"] = function_reference
        component_name = row[1]
        context["component_name"] = component_name
        function_name = row[2]
        context["function_name"] = function_name
        function_description = row[3]
        context["function_description"] = function_description
        author = row[4]
        context["author"] = author
        additional_comments = row[5]
        context["additional_comments"] = additional_comments
        i = i + 1
        tpl = DocxTemplate("FDD-temp.docx")
        tpl.render(context)
        tpl.save('new_test'+str(i)+'.docx')