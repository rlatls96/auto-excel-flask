from flask import Flask, request, send_file, render_template_string
import pandas as pd
from datetime import datetime
import os
import io

app = Flask(__name__)

# HTML Template
UPLOAD_PAGE = '''
<!doctype html>
<title>Excel Auto Processor</title>
<h1>엑셀 파일 업로드</h1>
<form method=post enctype=multipart/form-data>
  <input type=file name=file>
  <input type=submit value='변환 시작'>
</form>
'''

# Flask Route
@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            return '파일이 없습니다.'
        file = request.files['file']
        if file.filename == '':
            return '파일명이 없습니다.'

        # 엑셀 파일 처리
        df = pd.read_excel(file, sheet_name='Sheet1')

        columns_needed = [
            'MBL No', 'HBL No', 'POD', 'FDEST', 'Sub Loc Cd', 'GWT', 'VOL WT',
            'Qty', 'QTY Unit Cd', 'Item Name', 'CNEE Address', 'SHPR Address', 'POL ATD'
        ]

        df_extracted = df[columns_needed].copy()

        # 선사 코드 매핑
        prefix_to_ssl = {
            'MAEU': 'Maersk', 'HDMU': 'HMM', 'SMLM': 'SM Line', 'WWSU': 'SWIRE',
            'YMJA': 'YangMing', 'ONEY': 'ONE', 'CMDU': 'CMA-CGM', 'ZIMU': 'ZIM',
            'EGLV': 'CMA-CGM', 'MEDU': 'MSC', 'HLCU': 'Hapag-Lloyd',
            'COSU': 'COSCO', 'OOLU': 'OOCL', 'SSBF': 'SWIRE'
        }

        # FDEST 문자 -> 숫자 코드 매핑
        fdest_text_to_code = {
            'TOR': '495', 'NYK': '495', 'WDR': '495',
            'VAN': '809', 'BUB': '809', 'FSD': '809',
            'CAL': '701', 'EDM': '702',
            'MTR': '395', 'JON': '395', 'VVL': '395', 'PXT': '395',
            'WNP': '504'
        }

        # 도시+선사 → 서브로케이션 매핑
        preferred_subloc = {
            ('495', 'Maersk'): '3046 / 3037', ('495', 'HMM'): '3037', ('495', 'SM Line'): '3037',
            ('495', 'SWIRE'): '3037', ('495', 'YangMing'): '3046', ('495', 'ONE'): '3046 / 3037',
            ('495', 'CMA-CGM'): '3046 / 3037', ('495', 'ZIM'): '3037', ('495', 'Evergreen'): '3037',
            ('495', 'MSC'): '3037', ('495', 'Hapag-Lloyd'): '3046', ('495', 'OOCL'): '3037',
            ('495', 'COSCO'): '3037',
            ('809', 'Maersk'): '3380', ('809', 'HMM'): '3380 / 3891', ('809', 'SM Line'): '3401',
            ('809', 'SWIRE'): '3401', ('809', 'YangMing'): '3380 / 3891', ('809', 'ONE'): '3380 / 3891',
            ('809', 'CMA-CGM'): '3395 / 3891', ('809', 'ZIM'): '3891', ('809', 'Evergreen'): '3395',
            ('809', 'MSC'): '3380 / 3891', ('809', 'Hapag-Lloyd'): '3891', ('809', 'OOCL'): '3891',
            ('809', 'COSCO'): '3395',
            ('395', 'Maersk'): '2423', ('395', 'HMM'): '2414', ('395', 'SM Line'): '2414',
            ('395', 'SWIRE'): '2414', ('395', 'YangMing'): '2423', ('395', 'ONE'): '2414',
            ('395', 'CMA-CGM'): '2423', ('395', 'ZIM'): '2414', ('395', 'Evergreen'): '2414 / 2423',
            ('395', 'MSC'): '2414', ('395', 'Hapag-Lloyd'): '2423',
            ('504', 'Maersk'): '3147 / 3150', ('504', 'HMM'): '3147', ('504', 'SM Line'): '3147',
            ('504', 'SWIRE'): '3147', ('504', 'YangMing'): '3150', ('504', 'ONE'): '3147',
            ('504', 'CMA-CGM'): '', ('504', 'ZIM'): '3147', ('504', 'Evergreen'): '3147',
            ('504', 'MSC'): '3147', ('504', 'Hapag-Lloyd'): '',
            ('701', 'Maersk'): '5426', ('701', 'HMM'): '5426', ('701', 'SM Line'): '5426',
            ('701', 'SWIRE'): '5426', ('701', 'YangMing'): '3237', ('701', 'ONE'): '5426',
            ('701', 'CMA-CGM'): '', ('701', 'ZIM'): '5426', ('701', 'Evergreen'): '3237',
            ('701', 'MSC'): '5426', ('701', 'Hapag-Lloyd'): '',
            ('702', 'Maersk'): '3297', ('702', 'HMM'): '4492', ('702', 'SM Line'): '4492',
            ('702', 'SWIRE'): '4492', ('702', 'YangMing'): '3297', ('702', 'ONE'): '4492',
            ('702', 'CMA-CGM'): '', ('702', 'ZIM'): '4492', ('702', 'Evergreen'): '4492',
            ('702', 'MSC'): '4492', ('702', 'Hapag-Lloyd'): ''
        }

        # 변환된 MBL No 만들기
        carrier_mapping = {
            'MAEU': '9381', 'HDMU': '9463', 'SMLM': '918P', 'WWSU': '9311',
            'SSBF': '9311', 'YMJA': '91NG', 'ONEY': '919J', 'CMDU': '9558',
            'ZIMU': '9312ZIMU', 'EGLV': '9476', 'MEDU': '9066', 'HLCU': '9529',
            'COSU': '9502COSU', 'OOLU': '9082'
        }

        mbl_series = df['MBL No']
        scac_series = df['SCAC']
        converted_mbl = []

        for mbl, scac in zip(mbl_series, scac_series):
            if isinstance(mbl, str) and len(mbl) >= 4:
                prefix = mbl[:4]
                if prefix in carrier_mapping:
                    new_prefix = carrier_mapping[prefix]
                    converted_mbl.append(new_prefix + mbl[4:])
                elif isinstance(scac, str) and scac in carrier_mapping:
                    new_prefix = carrier_mapping[scac]
                    converted_mbl.append(new_prefix + mbl)
                else:
                    converted_mbl.append(mbl)
            else:
                converted_mbl.append(mbl)

        df_extracted.insert(1, 'MBL No (Carrier Code Changed)', converted_mbl)

        today_str = datetime.now().strftime("%Y-%m-%d")
        df_extracted['Today'] = today_str

        new_subloc = []
        for subloc, mbl, scac, fdest in zip(df_extracted['Sub Loc Cd'], df['MBL No'], df['SCAC'], df['FDEST']):
            if pd.isna(subloc):
                ssl = ''
                if isinstance(scac, str) and scac.strip() in prefix_to_ssl:
                    ssl = prefix_to_ssl[scac.strip()]
                elif isinstance(mbl, str) and len(mbl.strip()) >= 4:
                    prefix = mbl.strip()[:4]
                    ssl = prefix_to_ssl.get(prefix, '')

                fdest_clean = fdest.strip().upper() if isinstance(fdest, str) else ''
                fdest_str = fdest_text_to_code.get(fdest_clean, '')

                sub_loc = preferred_subloc.get((fdest_str, ssl), '')

                new_subloc.append(sub_loc)
            else:
                new_subloc.append(subloc)

        df_extracted['Sub Loc Cd'] = new_subloc

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_extracted.to_excel(writer, index=False)
        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name="Processed_Excel.xlsx",
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    return render_template_string(UPLOAD_PAGE)

if __name__ == '__main__':
    app.run(debug=True)
