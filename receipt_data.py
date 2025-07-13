import requests
import uuid
import time
import json
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
from openai import OpenAI

# CLOVA OCR API 호출
api_url = 'YOUR_API_URL'
secret_key = 'YOUR_SECRET_KEY'
file_name = "영수증5"
store_id = "맑은농장"
image_file = fr"YOUR_FILE_FATH\jpg\{store_id}\{file_name}.jpg"

request_json = {
    'images': [
        {
            'format': 'jpg',
            'name': 'demo'
        }
    ],
    'requestId': str(uuid.uuid4()),
    'version': 'V2',
    'timestamp': int(round(time.time() * 1000))
}

payload = {'message': json.dumps(request_json).encode('UTF-8')}

# image file 처리
files = [
  ('file', open(image_file,'rb'))
]
headers = {
  'X-OCR-SECRET': secret_key
}

response = requests.request("POST", api_url, headers=headers, data = payload, files = files)
json_data = f"YOUR_FILE_FATH\{file_name}.json"

try:
    with open(json_data, 'r', encoding='utf-8') as f:
        data = json.load(f)
        fields= data['images'][0]['fields']
        print(f"{file_name}.json 로드 완료")
except FileNotFoundError:
    with open(json_data, "w", encoding="utf-8") as f:
        json.dump(response.json(), f, ensure_ascii=False, indent=2)
        fields = response.json()['images'][0]['fields']
        print(f"{file_name}.json 생성 완료")
        
# 응답 내용을 영수증 형태로 다시 변환
string_result = ''
for i in fields:
    if i['lineBreak'] == True:
        linebreak = '\n'
    else:
        linebreak = ' '
    string_result = string_result + i['inferText'] + linebreak
    
print(string_result)

client = OpenAI(
  api_key="YOUR_API_KEY"
)

completion = client.chat.completions.create(
  model="gpt-3.5-turbo-1106",
  messages=[
    {"role": "system", "content": """You can extract the 날짜, 업체명, 품목, 단가, 수량 and 금액 from a receipt and convert them into a JSON file. 
    The keys must be strictly fixed as follows: 날짜, 업체명 and 상품목록. Do not change these names."""},
    {"role": "user", "content": f"""Analyze {string_result} and extract only the information related to the 날짜, 업체명, 품목, 단가, 수량 and 금액. 
    If the 상품명 consists of only numbers, exclude it."""}
  ]
)
message = completion.choices[0].message.content

print(message)

data = json.loads(message)

# 날짜와 업체명 추출
sales_date = data["날짜"]
store_name = data["업체명"]

# 파일 경로 및 오픈할 sheet 이름
file_path = "YOUR_FILE_FATH\csv\샘플 데이터.xlsx"
sheet_name = "지출내역"

#데이터 프레임으로 전환 및 생성
receipt_data = pd.DataFrame(data['상품목록'])

# 날짜, 업체명을 컬럼으로 앞에 삽입
receipt_data.insert(0, "날짜", sales_date)
# 간혹 날짜 형식이 yy-mm-dd 되어 있는 경우가 있으므로 날짜 형식 변환: yy-mm-dd → yyyy-mm-dd
receipt_data["날짜"] = pd.to_datetime(receipt_data["날짜"], format="%y-%m-%d").dt.strftime("%Y-%m-%d")

receipt_data.insert(1, "업체명", store_name)

columns_to_convert = ["단가", "수량", "금액"]

# 쉼표나 공백 등 제거하고 숫자로 변환
for col in columns_to_convert:
    receipt_data[col] = receipt_data[col].astype(str).str.replace(",", "").str.strip().fillna("0").astype(int)

# 엑셀 열기
wb = load_workbook(file_path)
ws = wb[sheet_name]

# 테이블 정보 가져오기
table_name = list(ws.tables.keys())[0]
table = ws.tables[table_name]
start_cell, end_cell = table.ref.split(":")
start_col_letter = ''.join(filter(str.isalpha, start_cell))
start_row = int(''.join(filter(str.isdigit, start_cell)))
end_col_letter = ''.join(filter(str.isalpha, end_cell))

start_col_index = column_index_from_string(start_col_letter)
end_col_index = column_index_from_string(end_col_letter)

# 실제 데이터가 있는 마지막 행 찾기
def get_last_data_row(ws, start_row, col_index):
    row = start_row
    while ws.cell(row=row, column=col_index).value:
        row += 1
    return row - 1

actual_last_row = get_last_data_row(ws, start_row + 1, start_col_index)

# 새 데이터 삽입 (실제 마지막 데이터 아래부터)
for r_idx, row in enumerate(receipt_data.values.tolist(), start=actual_last_row + 1):
    for c_idx, value in enumerate(row):
        ws.cell(row=r_idx, column=start_col_index + c_idx, value=value)

# 테이블 범위 재설정
new_end_row = actual_last_row + len(receipt_data)
new_ref = f"{start_col_letter}{start_row}:{end_col_letter}{new_end_row}"

# 지정한 파일 경로에 저장 및 확인
wb.save(file_path)
print(f"{file_path} 파일에 저장되었습니다.")