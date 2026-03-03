from src.generator import DiplomaGenerator
import json
grades = {
    'он11бизнестiн': {'hours': '5', 'credits': '1'},
    'он12бизнестiн': {'hours': '10', 'credits': '2'},
}
student = {'name': 'Test', 'grades': grades}
d = DiplomaGenerator('templates/diploma_ru_template.xlsx', 'out.xlsx', {}, {})
d.workbook.worksheets = []  # Bypass UI
d.fill_student_data(student)
