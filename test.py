import docx
from docx.shared import RGBColor, Pt
from utils import get_formula_values
from config import input_word_file_name, output_word_file_name, values_to_replace, file_name_excel

formuls = get_formula_values(values_to_replace, file_name_excel)
# for i, j in formuls.items():
#     print(i,' - ', j)



doc = docx.Document(input_word_file_name)


for paragraph in doc.paragraphs:

    for key, value in formuls.items():

        if key in paragraph.text:

            for run in paragraph.runs:

                if key in run.text:
                    run.font.color.rgb = RGBColor(255, 0, 0)
                    run.font.name = 'Times New Roman'
                    run.font.size = docx.shared.Pt(14)
                    run.text = run.text.replace(key, value)

doc.save(output_word_file_name)
