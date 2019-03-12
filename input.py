# 디펜던시는 xlrd만 까셈
import re, xlrd


class Inputer:
    def __init__(self):
        self.original_html = None
        self.base_html = None
        self.output_html = ''
        self.output_javascript = ''

    def make_html(self, num, degree, course_name, course_type, year, semester, point, grade, retake):
        html = self.base_html
        # 행번호 정보 변경
        html = re.sub('<td id="majdetNo">\\d*', '<td id="majdetNo">' + num, html)
        html = re.sub('majtypecd_\\d*', 'majtypecd_' + num, html)
        html = re.sub('majtypenm_\\d*', 'majtypenm_' + num, html)
        html = re.sub('semstcd_\\d*', 'semstcd_' + num, html)
        html = re.sub('regyr_\\d*', 'regyr_' + num, html)
        html = re.sub('obtptcd_\\d*', 'obtptcd_' + num, html)
        html = re.sub('obtpovcd_\\d*', 'obtpovcd_' + num, html)
        html = re.sub('retakeyncd_\\d*', 'retakeyncd_' + num, html)
        # 학위정보 변경
        html = re.sub('학사', degree, html)
        # 강의명 변경
        html = re.sub('name="majcurrinm" value=".*?" title=".*?"', f'name="majcurrinm" value="{course_name}" title="{course_name}"', html)

        self.output_html += html

        # 아래는 자바스크립트로 처리합니다.
        # 강의종류 변경
        print(num)
        if '전공' in course_type:
            self.output_javascript += 'document.getElementById("majtypecd_' + num + '").value = "A";'
            self.output_javascript += 'document.getElementById("majtypenm_' + num + '").value = "전공";'
        else:
            self.output_javascript += 'document.getElementById("majtypecd_' + num + '").value = "C";'
            self.output_javascript += 'document.getElementById("majtypenm_' + num + '").value = "교양기타";'
        # 수강연도 변경
        self.output_javascript += 'document.getElementById("regyr_' + num + '").value = "' + year +'";'
        # 수강학기 변경
        self.output_javascript += 'document.getElementById("semstcd_' + num + '").value = "' + semester + '";'
        # 학점 정보 변경
        self.output_javascript += 'document.getElementById("obtptcd_' + num + '").value = "' + point + '";'
        # 평점 정보 변경 (A0는 A로 넣어야 합니다. pass는 PASS fail은 FAIL로 넣어야 합니다.)
        self.output_javascript += 'document.getElementById("obtpovcd_' + num + '").value = "' + grade + '";'
        # 재수강 정보 변경
        self.output_javascript += 'document.getElementById("retakeyncd_' + num + '").value = "' + retake + '";'

    def feed_html(self, html):
        self.original_html = html = re.sub('\n', '', html)
        self.base_html = re.search('<tr><td id="majdetNo">.*</tr>', html)[0]

    def main(self):
        excel_file_address = input('엑셀 파일 주소를 입력하세요.')
        excel_file = xlrd.open_workbook(excel_file_address)
        sheet = excel_file.sheet_by_index(0)  # 시트가 엑셀 파일의 첫 번째에 위치해야 한다.

        html = ''

        input_data = input('thred 태그 안의 data를 붙여넣은 후 엔터하세요.')
        while True:
            html += input_data
            input_data = input()

            if input_data == '':
                break

        self.feed_html(html)

        for number in range(1, sheet.nrows):  # 첫번째 row는 뺀다. (col 이름이 들어간 row)
            row = sheet.row(number)

            degree = row[0].value
            course_name = row[1].value
            course_type = row[2].value
            year_and_semester = row[3].value
            point = str(int(row[4].value))
            grade = row[5].value
            grade = re.sub('(?P<grade>.)0', '\\g<grade>', grade)
            grade = re.sub('P', 'PASS', grade)
            grade = re.sub('F', 'FAIL', grade)
            retake = row[6].value

            string = re.search('.*(?P<year>\\d{4})년-.*-(?P<semester>.{1,4})학기', year_and_semester)

            year = string.group(1)
            semester = string.group(2)

            self.make_html(str(number), degree, course_name, course_type, year, semester, point, grade, retake)

        result = re.sub('(?P<head>.*)<tr><td id="majdetNo">.*</tr>(?P<tail>.*)', f'\\g<head>{self.output_html}\\g<tail>', html)

        print('아래를 개발자 도구로 붙여넣기 하세요.')
        print(result)

        #자바스크립트 버튼 생성
        buttorn_html = '<button id="java_button" type="button" onclick="rewrite()">이 버튼을 눌러서 교정</button>'
        print(buttorn_html)

        #자바스크립트 추가
        print('아래는 body 바로 아래에 넣으십시오.')
        javascript = '<script>' \
                 'function rewrite() {' \
                 + self.output_javascript + \
                 'document.getElementById("java_button").remove();' \
                 '}' \
                 '</script>' \

        print(javascript)



inputer = Inputer()
inputer.main()