from openpyxl import Workbook

#  두번째 요구사항의 시트의 배이스 엑셀을 만든다


def make_third_base(write_wb, width, select_write_sheet):
    #  구조 데이터 입력
    header_data1 = [
        [select_write_sheet + "프로그램 참여목록"],
        ["No", "신청경로", "학번", "성명",  "연락처", "학년", "성별", "소속학과", "이수여부", "빌드업신청",
         "상담사", "신청일자", "신청내용"]
    ]

    for row_data in header_data1:
        write_wb.append(row_data)

    for column in range(ord('A'), ord('Z')):  # B부터 Z까지의 열에 대해 반복
        column_letter = chr(column)
        write_wb.column_dimensions[column_letter].width = width
