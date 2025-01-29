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
    file = request.files.get('file')
    if not file or file.filename == '':
        return "파일이 없습니다."
    
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(file_path)
    
    try:
        df = pd.read_csv(file_path, encoding='CP949', skiprows=1)
    except Exception as e:
        return f"CSV 파일을 처리하는 중 오류가 발생했습니다: {str(e)}"
    
    df = df.dropna(subset=['연번', '이름']).drop(columns=['비고1', '비고2', 'Unnamed: 14'], errors='ignore')
    df = df.rename(columns={'교육\n일시': '교육일시', '교육\n시간': '교육시간'})
    df[['연번', '교육시간']] = df[['연번', '교육시간']].astype('int')
    
    include_date = request.form.get('include_date') == 'yes'
    subset_columns = ['이름', '구분1(외부/내부)', '구분2(법정의무/직무역량)', '과정구분3', '과정명']
    if include_date:
        subset_columns.append('교육일시')
    
    duplicated_names = df[df.duplicated(subset=subset_columns, keep=False)]
    duplicated_file = os.path.join(PROCESSED_FOLDER, 'duplicated_names.csv')
    duplicated_names.to_csv(duplicated_file, index=False, encoding='CP949')
    
    return render_template('edu_result.html', duplicated_file='duplicated_names.csv')

@app.route('/download/<filename>')
def download_file(filename):
    file_path = os.path.join(PROCESSED_FOLDER, filename)
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
    
    df_trip = pd.read_excel(trip_path, header=1)
    df_trip.columns.values[:8] = pd.read_excel(trip_path, nrows=0).columns[:8]
    df_trip.drop(['No', '근태분류', '첨부파일', '신청서', '문서제목', '문서삭제사유', '결재상태'], axis=1, inplace=True, errors='ignore')
    df_trip = df_trip[df_trip['근태항목'] == '관내출장']
    
    df_trip['여비'] = 10000
    df_trip.loc[df_trip['출장시간'] >= 240, '여비'] = 20000
    df_trip.loc[df_trip['교통수단'] == '관용차량', '여비'] -= 10000
    df_trip.loc[df_trip['여비'] < 0, '여비'] = 0
    
    output_path = os.path.join(UPLOAD_FOLDER, 'processed_trip_all.xlsx')
    df_trip.to_excel(output_path, index=False)
    
    return render_template('trip_result.html', output_path='processed_trip_all.xlsx')

if __name__ == '__main__':
    app.run(debug=True)
