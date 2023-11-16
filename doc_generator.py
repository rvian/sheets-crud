from docx import Document
from docxtpl import DocxTemplate
from tkinter.filedialog import askdirectory


def GerarDoc(columnValues):

    # Importar documento modelo
    template = DocxTemplate('docs\Modelo.docx')
    
    # Solicita path onde será salvo o doc
    path = askdirectory()
    
    if path == '':
        return
    
    # Renderiza os valores nos placeholders e salva o arquivo no path definido
    template.render(columnValues)
    template.save(path.replace('/','\\')+'\Sheets_'+columnValues["CNPJ"].replace('/','')+'.docx')
    
    