import os
import xlwings as xw
import requests
import xml.etree.ElementTree as ET
import re

# 원본 엑셀 파일 경로
origin_excel_path = r"C:\cji_d\apiTest\pat"
target_param = "출원번호"
key_path=r"C:\cji_d\06. Python\05 program\2_rejectApi"

# access_key.txt에서 API 키 읽기
def load_access_key():
    key_file_path = os.path.join(key_path, "access_key.txt")
    try:
        with open(key_file_path, "r", encoding="utf-8") as f:
            return f.read().strip()
    except FileNotFoundError:
        print("access_key.txt 파일을 찾을 수 없습니다.")
        return None

access_key = load_access_key()
if not access_key:
    exit("프로그램을 종료합니다. API 키가 필요합니다.")
print(f"access_key: {access_key}")
# 엑셀 파일 불러오기 (xls, xlsx 지원)
def load_excel(file_path):
    app = xw.App(visible=False)
    wb = app.books.open(file_path)
    return app, wb

# API 요청 및 응답 처리 함수
def request_rejection_api(application_number):
    url = f"http://plus.kipris.or.kr/openapi/rest/IntermediateDocumentOPService/rejectDecisionInfo?applicationNumber={application_number}&accessKey={access_key}"
    response = requests.get(url, headers={"Accept": "application/xml"})
    print(f"url: {url}")

    if response.status_code == 200:
        print("200 OK 응답 받음")

        # 응답 메시지 확인
        msg_or_none = extract_result_msg(response.text)
        if msg_or_none:
            return msg_or_none  # 오류 메시지 또는 거절 사유 없음 메시지 반환

        # 정상적인 경우만 XML 파싱 진행
        return request_rejection_reason_parsing(response.text)

    return "API 요청 실패"


# API 응답 메시지 확인 함수
def extract_result_msg(xml_text):
    """
    API 응답에서 resultMsg 태그의 메시지를 추출하고,
    <body> 태그 내 컨텐츠가 비어 있는지 확인하여 적절한 메시지를 반환.
    """
    try:
        root = ET.fromstring(xml_text)
        
        # resultMsg 태그 확인
        result_msg = root.find(".//resultMsg")
        if result_msg is not None and result_msg.text:
            print(f"API 오류 메시지: {result_msg.text}")
            return result_msg.text  # 오류 메시지 반환

        # body 태그 내부 확인
        body = root.find(".//body")
        if body is not None and not any(body):
            print("해당 출원번호에 대한 API 응답에 거절이유 내용이 비어있습니다.")
            return "해당 출원번호에 대한 API 응답에 거절이유 내용이 비어있습니다."

        return None  # 정상 응답일 경우 None 반환
    
    except ET.ParseError:
        print(" XML 파싱 오류")
        return "XML 파싱 오류 발생"
    




# XML 파싱 및 거절 사유 추출 함수
def request_rejection_reason_parsing(xml_text):
    root = ET.fromstring(xml_text)
    rejection_info_list = root.findall(".//rejectDecisionInfo")
    print(f"rejection_info_list: {len(rejection_info_list)}")  # 거절 사유 개수 출력
    
    if not rejection_info_list:
        return None

    reason_texts = []
    total_rejections = 0  # 거절 사유의 총 개수 카운트

    for idx, reject_info in enumerate(rejection_info_list, 1):  # 번호 매기기
        attachmentfile = reject_info.find("attachmentfileContent").text if reject_info.find("attachmentfileContent") is not None else ""
        details = reject_info.findall("rejectionContentDetail")  # 거절 사유가 여러 개일 수 있음
        attachmentfile = html_modify(attachmentfile)
        # 상세 내용이 여러 개일 경우 반복문으로 처리
        for detail in details:
            detail_text = detail.text if detail is not None else ""
            
            html_text = html_modify(detail_text)
            print("html_text: " + html_text)
            

            # 제목과 상세 내용 형식으로 저장
            reason_texts.append(f"< {idx} 번째 의견통지서 >\n{attachmentfile}\n{html_text}" if attachmentfile else f"< 거 절 이 유 {idx} >\n{html_text}")

            total_rejections += 1  # 거절 사유 개수 증가

    print(f"Total rejection reasons: {total_rejections}")  # 거절 사유 총 개수 출력
    return "\n\n".join(reason_texts)  # 여러 개의 거절 사유를 개행 문자로 구분하여 반환


# 엑셀 데이터 처리 함수
def process_excel(file_path):
    app, wb = load_excel(file_path)
    ws = wb.sheets[0]  # 첫 번째 시트 선택

    # A8 셀부터 마지막 행과 열을 찾는 부분
    
    first_row = 8
    first_col = 1

    # 마지막 열 찾기 (A8부터 시작)
    last_col = ws.range(first_row, first_col).end('right').column
    
    # 마지막 행 찾기 (A8부터 시작)
    last_row = ws.range(first_row, first_col).end('down').row

    target_col = None

    # target_param과 일치하는 출원번호 열 찾기 (A8부터 A:last_col까지)
    for col in range(first_col, last_col + 1):
        if ws.cells(first_row, col).value == target_param:
            target_col = col
            break

    if target_col is None:
        print("출원번호가 없습니다.")
        wb.close()
        app.quit()
        return

    result_col = last_col + 1  # API 응답이 들어갈 새로운 열
    ws.cells(first_row, result_col).value = "거절이유"  # 새 필드명 추가

    print(f"target_col: {target_col}, last_row: {last_row} , last_column: {last_col}, result_col: {result_col}")

    # 데이터 순회
    for row in range(first_row + 1, last_row + 1):
        app_number = ws.cells(row, target_col).value
        if app_number:
            # '-'를 제거하고 숫자만 남김
            app_number = str(app_number).replace("-", "")
            rejection_text = request_rejection_api(app_number)

            ws.cells(row, result_col).value = rejection_text if rejection_text else ""  # 거절 사유 입력

    # 새 파일로 저장
    new_file_path = os.path.join(origin_excel_path, "processed_" + os.path.basename(file_path))
    wb.save(new_file_path)
    wb.close()
    app.quit()
    print(f"✅ 처리 완료: {new_file_path}")

# HTML 수정
def html_modify(html_contents):
    html_txt = html_contents    
    # HTML 엔터티 및 개행 문자 변환
    html_txt = re.sub(r"<BR>|BR&gt;", "\n", html_txt, flags=re.IGNORECASE)  # <BR> → 개행 문자로 변환
    html_txt = re.sub(r"&lt;", "<", html_txt)  # HTML 엔터티 변환
    html_txt = re.sub(r"&gt;", ">", html_txt)
    html_txt = re.sub(r"</?p>", " ", html_txt, flags=re.IGNORECASE)  # <p> 또는 </p>를 공백으로 변경
    html_txt = html_txt.strip()        
    return html_txt


# 폴더 내 엑셀 파일 처리
files = [f for f in os.listdir(origin_excel_path) if f.endswith((".xls", ".xlsx"))]
if files:
    process_excel(os.path.join(origin_excel_path, files[0]))
else:
    print("엑셀 파일이 없습니다.")
