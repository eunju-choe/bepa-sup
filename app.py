from flask import Flask, render_template, request, send_file, redirect, url_for
import pandas as pd
import os
import zipfile
import re

app = Flask(__name__)

# 업로드 및 처리 폴더 설정
UPLOAD_FOLDER = 'uploads'
PROCESSED_FOLDER = 'processed'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['PROCESSED_FOLDER'] = PROCESSED_FOLDER

@app.route('/')
def index():
    return render_template('index.html')

"""
=============== 교육 담당자 서포터 기능 ===============
"""
@app.route('/edu')
def edu_index():
    return render_template('edu_index.html')

@app.route('/edu/upload', methods=['POST'])
def upload_edu_file():
    if 'file' not in request.files:
        return "파일이 없습니다."
    
    file = request.files['file']
    if file.filename == '':
        return "파일 이름이 없습니다."
    
    # CSV 파일 저장
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(file_path)
    
    # CSV 파일 처리
    try:
        df = pd.read_csv(file_path, encoding='CP949', skiprows=1)
    except Exception as e:
        return f"CSV 파일을 처리하는 중 오류가 발생했습니다: {str(e)}"
    
    df = df.dropna(subset=['연번', '이름']).drop(columns=['비고1', '비고2', 'Unnamed: 14'])
    df = df.rename(columns={'교육\n일시': '교육일시', '교육\n시간': '교육시간'})
    df[['연번', '교육시간']] = df[['연번', '교육시간']].astype('int')
    
    # 체크박스 선택 여부 확인
    include_date = request.form.get('include_date') == 'yes'

    # 중복된 이름이 있는 데이터 추출 (체크박스 선택 시 '교육일시' 포함)
    subset_columns = ['이름', '구분1(외부/내부)', '구분2(법정의무/직무역량)', '과정구분3', '과정명']
    if include_date:
        subset_columns.append('교육일시')

    duplicated_names = df[df.duplicated(subset=subset_columns, keep=False)]
    duplicated_names = duplicated_names.sort_values(by=['이름', '과정명'])

    # 이름 개수 불일치 확인
    name_counts = df.groupby('과정명')['이름'].agg(고유개수='nunique', 이름개수='count').reset_index()
    name_mismatch = name_counts[name_counts['이름개수'] != name_counts['고유개수']]

    # 공백과 괄호 제거 전 후 비교
    space_yes = df.groupby('과정명').size().reset_index(name='띄어쓰기 제거 전')
    space_yes['과정명'] = space_yes['과정명'].str.replace(' ', '', regex=False)

    bracket_yes = space_yes.groupby('과정명').sum().reset_index()
    bracket_yes.columns = ['과정명', '괄호 제거 전']
    bracket_yes['과정명'] = bracket_yes['과정명'].apply(lambda x:re.sub(r'\(\d+(차시|시간)\)', '', x))

    space_no = df.copy()
    space_no['과정명'] = space_no['과정명'].str.replace(' ', '', regex=False)
    space_no = space_no.groupby('과정명').size().reset_index(name='띄어쓰기 제거 후')

    bracket_no = space_no.copy()
    bracket_no['과정명'] = bracket_no['과정명'].apply(lambda x: re.sub(r'\(\d+(차시|시간)\)', '', x))
    bracket_no = bracket_no.groupby('과정명').sum().reset_index()
    bracket_no.columns = ['과정명', '괄호 제거 후']

    space_df = pd.merge(space_yes, space_no, on='과정명', how='outer')
    space_df['일치 여부'] = space_df['띄어쓰기 제거 전'] == space_df['띄어쓰기 제거 후']
    space_df = space_df[space_df['일치 여부'] == False]
    space_df = space_df.groupby('과정명').agg({
        "띄어쓰기 제거 전" : lambda x: ", ".join(map(str, x)),
        "띄어쓰기 제거 후" : "first"
    }).reset_index()
    space_df.columns = ['과정명', '구분 별 개수', '전체 개수']

    bracket_df = pd.merge(bracket_yes, bracket_no, on='과정명', how='outer')
    bracket_df['일치 여부'] = bracket_df['괄호 제거 전'] == bracket_df['괄호 제거 후']
    bracket_df = bracket_df[bracket_df['일치 여부'] == False]
    bracket_df = bracket_df.groupby('과정명').agg({
        "괄호 제거 전" : lambda x: ", ".join(map(str, x)),
        "괄호 제거 후" : "first"
    }).reset_index()
    bracket_df.columns = ['과정명', '구분 별 개수', '전체 개수']

    # 파일 저장 경로 설정
    duplicated_file = os.path.join(app.config['PROCESSED_FOLDER'], 'duplicated_names.csv')
    mismatch_file = os.path.join(app.config['PROCESSED_FOLDER'], 'name_mismatch.csv')
    space_comparison_file = os.path.join(app.config['PROCESSED_FOLDER'], 'space_comparison.csv')
    bracket_comparison_file = os.path.join(app.config['PROCESSED_FOLDER'], 'bracket_comparison.csv')
    
    # 파일 경로 출력 ## 테스트 코드
    print(f"Duplicated file path: {duplicated_file}")
    print(f"Mismatch file path: {mismatch_file}")
    print(f"Space comparison file path: {space_comparison_file}")
    print(f"Bracket comparison file path: {bracket_comparison_file}")
    
    # 처리된 데이터 저장
    duplicated_names.to_csv(duplicated_file, index=False, encoding='CP949')
    name_mismatch[['과정명', '고유개수', '이름개수']].to_csv(mismatch_file, index=False, encoding='CP949')
    space_df.to_csv(space_comparison_file, index=False, encoding='CP949')
    bracket_df.to_csv(bracket_comparison_file, index=False, encoding='CP949')

    # 처리된 데이터프레임을 HTML로 변환하여 보여주기
    return render_template('edu_result.html', 
                           duplicated_names=duplicated_names.to_html(index=False, escape=False),
                           name_mismatch=name_mismatch[['과정명', '고유개수', '이름개수']].to_html(index=False, escape=False),
                           space_comparison=space_df.to_html(index=False, escape=False),
                           bracket_comparison=bracket_df.to_html(index=False, escape=False),
                           duplicated_file='duplicated_names.csv',
                           mismatch_file='name_mismatch.csv',
                           space_comparison_file='space_comparison.csv',
                           bracket_comparison_file='bracket_comparison.csv')

# 파일 다운로드 처리
@app.route('/download/<filename>')
def download_file(filename):
    file_path = os.path.join(app.config['PROCESSED_FOLDER'], filename)
    return send_file(file_path, as_attachment=True)

"""
=============== 관내여비 담당자 서포터 기능 ===============
"""
@app.route('/trip')
def trip_index():
    return render_template('trip_index.html')

@app.route('/trip/upload', methods=['POST'])
def upload_trip_file():
    trip_file = request.files.get('trip_file')
    tag_file = request.files.get('tag_file')
    
    if not trip_file or not tag_file or trip_file.filename == '' or tag_file.filename == '':
        return '파일이 없습니다.'
    
    trip_path = os.path.join(UPLOAD_FOLDER, 'trip_all.xlsx')
    tag_path = os.path.join(UPLOAD_FOLDER, 'tag_all.xlsx')
    
    trip_file.save(trip_path)
    tag_file.save(tag_path)
    
    # 엑셀 파일 처리
    df_trip = pd.read_excel(trip_path, header=1)
    df_trip.columns.values[:8] = pd.read_excel(trip_path, nrows=0).columns[:8]
    df_trip.drop(['No', '근태분류', '첨부파일', '신청서', '문서제목', '문서삭제사유', '결재상태'], axis=1, inplace=True, errors='ignore')
    df_trip = df_trip[df_trip['근태항목'] == '관내출장']
    
    # 여비 산출 로직
    df_trip['여비'] = 10000
    df_trip.loc[df_trip['출장시간'] >= 240, '여비'] = 20000
    df_trip.loc[df_trip['교통수단'] == '관용차량', '여비'] -= 10000
    df_trip.loc[df_trip['여비'] < 0, '여비'] = 0
    
    # 부서별 엑셀 파일 저장 함수
    def save_department_files(df_trip, folder_path):
        department_files = []
        for dept, group in df_trip.groupby('부서'):
            file_path = os.path.join(folder_path, f'{dept}_trip.xlsx')
            group.to_excel(file_path, index=False)
            department_files.append(file_path)
        return department_files
    
    # 부서별 엑셀 파일 저장
    department_files = save_department_files(df_trip, app.config['PROCESSED_FOLDER'])
    
    # 부서별 엑셀 파일 압축본 생성
    def create_zip_file(file_paths, zip_path):
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for file in file_paths:
                zipf.write(file, os.path.basename(file))
        return zip_path
    
    zip_path = os.path.join(PROCESSED_FOLDER, 'department_files.zip')
    zip_file_path = create_zip_file(department_files, zip_path)
    
    # 최종 처리된 파일 저장
    output_path = os.path.join(UPLOAD_FOLDER, 'processed_trip_all.xlsx')
    df_trip.to_excel(output_path, index=False)
    
    return render_template('trip_result.html', output_path='processed_trip_all.xlsx', zip_file_path='department_files.zip')

if __name__ == '__main__':
    app.run(debug=True)