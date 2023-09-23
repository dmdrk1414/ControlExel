from openpyxl import Workbook

#  두번째 요구사항의 시트의 배이스 엑셀을 만든다

# 시트와 원하는 열의 넓이를 적어라


def makeThirdBase(write_wb, width):
    # 첫 번째 구조 데이터 입력
    header_data1 = [
        [""],
        ["전산망과 빌드업 ", "참여자 db을 확인하여", "같은 학번이 있으면", "그 학번의 정산망자료를", "넣는다."],
        ["구분", "번호", "사업연도", "신청구분", "참여자명", "생년월일",
                     "참여유형", "참여분류", "학번", "전공", "연락처", "신청일자",
                     "처리상태", "처리일자", "담당자명", "대학명"]
    ]

    for row_data in header_data1:
        write_wb.append(row_data)

    for column in range(ord('A'), ord('Z')):  # B부터 D까지의 열에 대해 반복
        column_letter = chr(column)
        write_wb.column_dimensions[column_letter].width = width
