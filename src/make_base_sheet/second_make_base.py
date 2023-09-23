from openpyxl import Workbook

#  두번째 요구사항의 시트의 배이스 엑셀을 만든다

# 시트와 원하는 열의 넓이를 적어라


def makeSecondBase(write_wb, width):
    # 첫 번째 구조 데이터 입력
    header_data1 = [
        ["", "", "", "", "재맞고 신청", "", "", "", "", "", "1차 상담", "",
            "", "", "2차 상담", "", "", "", "포트폴리오", "", "", "", ""],
        ["구분", "", "", "", "교과", "비교과", "", "", "", "", "교과", "", "비교과",
            "교과", "비교과", "교과", "비교과", "", "", "", "", "", ""],
        ["개발시트", "이름", "학번", "인정여부", "직업탐색과미래설계", "신입생세미나",
            "포트폴리오워크숍", "단기직무체험", "자기주도", "진로집단상담", "직업탐색과미래설계", "신입생세미나", "자기주도", "진로집단상담", "직업탐색과미래설계", "신입생세미나",
            "자기주도", "진로집단상담", "직업탐색과미래설계", "포트폴리오워크숍", "자기주도", "진로집단상담"]
    ]

    for row_data in header_data1:
        write_wb.append(row_data)

    for column in range(ord('A'), ord('Z')):  # B부터 D까지의 열에 대해 반복
        column_letter = chr(column)
        write_wb.column_dimensions[column_letter].width = width
