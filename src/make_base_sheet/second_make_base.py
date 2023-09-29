from openpyxl import Workbook

#  두번째 요구사항의 시트의 배이스 엑셀을 만든다

# 시트와 원하는 열의 넓이를 적어라


def makeThirdBase(write_wb, width):
    # 첫 번째 구조 데이터 입력
    header_data1 = ["순번", "이름", "연락처", "학번",  "상담구분", "사업참여유형", "상담방법", "상담유형", "상담유형 기타", "상담사명",
                    "회기", "상담일자", "상담시", "상담분", "상담시간(분)", "상담내용", "차기상담신청유형", "차기상담일자", "차기상담시", "차기상담분"]

    for row_data in header_data1:
        write_wb.append(row_data)

    for column in range(ord('A'), ord('Z')):  # B부터 D까지의 열에 대해 반복
        column_letter = chr(column)
        write_wb.column_dimensions[column_letter].width = width
