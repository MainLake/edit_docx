from docx import Document
import os
from docx2pdf import convert

def fill_document(template_path, output_path, data):
    doc = Document(template_path)

    print('Dentro de fill_document')

    for paragraph in doc.paragraphs:
        for key, value in data.items():
            if key in paragraph.text:
                for run in paragraph.runs:
                    if key in run.text:
                        print(run.text)
                        # Reemplazar el texto dentro del run
                        new_text = run.text.replace(key, value)
                        run.text = new_text
                        
                        # Aplicar subrayado y negrita al nuevo texto
                        run.underline = True
                        run.bold = True

    doc.save(output_path)
    

if __name__ ==  '__main__':
    data = {
        "[p1dia]": "01",
        "[p1mes]": "01",
        "[p1anio]": "2021",
        "[nombre_alumno]": "Pablo Julian Garay de Leon",
        "[p1_alumno_matricula_1]": "223017",
        "[p1_tutor_rol_1]": "Padre",
        "[nombre_tutor]": "Julio Garay",
        "[p1_alumno_nombre_2]": "______Pablo Julian Garay de Leon______",
        "[p1_tutor_nombre_2]": "Julio Garay",
        "[p1_tutor_rol_2]": "Padre",
        "[p1_alumno_nombre_3]": "Pablo Julian Garay de Leon",
        "[p1_alumno_matricula_2]": "223017",
        "[p1_alumno_carrera]": "Ingenieria en Software",
        "[p1_tutor_nombre_3]": "______Julio Garay______",
        "[p2_dia]": "01",
        "[p2_mes]": "01",
        "[p2_anio]": "2021",
        "[p2_empresa_nombre]": "Empresa de Prueba______",
    }

    template_path = 'carta-de-exclusion-conocimientomedidas-de-proteccioncovid.docx'
    output_path = 'carta-de-exclusion-conocimientomedidas-de-proteccioncovid-filled.docx'

    if os.path.exists(output_path):
        os.remove(output_path)


    fill_document(template_path, output_path, data)

    convert(output_path, 'carta-de-exclusion-conocimientomedidas-de-proteccioncovid-filled.pdf')
    print('Documento creado con exito')

