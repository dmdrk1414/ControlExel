from openpyxl import Workbook

#  두번째 요구사항의 시트의 배이스 엑셀을 만든다


def make_second_base(write_wb, width):
    #  구조 데이터 입력
    header_data1 = ["순번", "이름", "연락처", "학번",  "상담구분", "사업참여유형", "상담방법", "상담유형", "상담유형 기타 사유", "상담사명",
                    "회기", "상담일자", "상담시", "상담분", "상담시간(분)", "상담내용", "차기상담신청유형", "차기상담일자", "차기상담시", "차기상담분"]

    write_wb.append(header_data1)

    for column in range(ord('A'), ord('Z')):  # B부터 Z까지의 열에 대해 반복
        column_letter = chr(column)
        write_wb.column_dimensions[column_letter].width = width
