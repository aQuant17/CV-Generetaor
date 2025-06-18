from flask import Flask, render_template, request, send_file
from docxtpl import DocxTemplate
import uuid

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def form():
    if request.method == 'POST':
        form = request.form
        data = form.to_dict()

        # EDUCATION
        edu_institution = form.getlist('edu_institution')
        edu_start = form.getlist('edu_start')
        edu_end = form.getlist('edu_end')
        edu_degree = form.getlist('edu_degree')
        data['education'] = []
        for i in range(len(edu_institution)):
            data['education'].append({
                'institution': edu_institution[i],
                'start': edu_start[i],
                'end': edu_end[i],
                'degree': edu_degree[i]
            })

        # EXPERIENCE
        exp_from = form.getlist('exp_from')
        exp_to = form.getlist('exp_to')
        exp_wds = form.getlist('exp_wds')
        exp_location = form.getlist('exp_location')
        exp_company = form.getlist('exp_company')
        exp_position = form.getlist('exp_position')
        exp_description = form.getlist('exp_description')
        data['experience'] = []
        for i in range(len(exp_from)):
            data['experience'].append({
                'from': exp_from[i],
                'to': exp_to[i],
                'wds': exp_wds[i],
                'location': exp_location[i],
                'company': exp_company[i],
                'position': exp_position[i],
                'description': exp_description[i]
            })

        # LANGUAGES
        lang_name = form.getlist('lang_name')
        lang_reading = form.getlist('lang_reading')
        lang_speaking = form.getlist('lang_speaking')
        lang_writing = form.getlist('lang_writing')
        data['languages'] = []
        for i in range(len(lang_name)):
            data['languages'].append({
                'name': lang_name[i],
                'reading': lang_reading[i],
                'speaking': lang_speaking[i],
                'writing': lang_writing[i]
            })

        # REGION EXPERIENCE
        country1 = form.getlist('country1')
        dates1 = form.getlist('dates1')
        country2 = form.getlist('country2')
        dates2 = form.getlist('dates2')
        data['region_experience'] = []
        for i in range(len(country1)):
            data['region_experience'].append({
                'country1': country1[i],
                'dates1': dates1[i],
                'country2': country2[i],
                'dates2': dates2[i]
            })

        # KEY QUALIFICATIONS
        key_qual = form.get('key_qualifications', '')
        data['key_qualifications'] = [x.strip() for x in key_qual.split(',') if x.strip()]

        # OTHER SIMPLE FIELDS
        data['memberships'] = form.get('memberships', '')
        other_info = form.get('other_relevant_info', '')
        data['other_relevant_info'] = [x.strip() for x in other_info.split(',') if x.strip()]
        data['category'] = form.get('category', '')
        data['staff_of'] = form.get('staff_of', '')
        data['civil_status'] = form.get('civil_status', '')
        data['present_position'] = form.get('present_position', '')
        data['years_in_firm'] = form.get('years_in_firm', '')
        data['other_skills'] = form.get('other_skills', '')
        data['first_names'] = form.get('first_names', '')
        data['family_name'] = form.get('family_name', '')
        data['dob'] = form.get('dob', '')
        data['nationality'] = form.get('nationality', '')

        # RENDER WORD TEMPLATE
        doc = DocxTemplate("EU_template.docx")
        doc.render(data)

        filename = f"output/CV_{uuid.uuid4()}.docx"
        doc.save(filename)

        return send_file(filename, as_attachment=True)

    return render_template("form.html")

if __name__ == '__main__':
    app.run(debug=True)
