from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
from openpyxl import Workbook

직업탐색과미래설계 = "직업탐색과미래설계"
신입생세미나 = "신입생세미나"
포트폴리오워크숍 = "포트폴리오워크숍"
포트폴리오컨설팅 = "포트폴리오컨설팅"
자기주도 = "자기주도"
진로집단상담 = "진로집단상담"
단기직무체험 = "단기직무체험"


신청경로 = "신청경로"
one상담 = "1차상담"
two상담 = "2차상담"
three상담 = "3차상담"
포트폴리오 = "포트폴리오"


def makeSecondExcel(workReadSheet, writeSheet, readExcelFile):
    # 5번 행의 전체 열을 검사하여 "학번"이 있는 열 찾기
    build_up_row = workReadSheet[2]  # 5번 행 선택
    column_with_student_id = None
    column_with_student_name = None
    column_with_student_acknowledgment = None
    column_with_student_application_category = None

    # build_up 의  시트의  학생들이  있는  기본적인 row
    build_up_base_row = 3
    build_up_max_row = workReadSheet.max_row

    # writ의 시트의 기본 학생 행싶 (쓰고싶은)
    write_sheet_base_row = 4

    for cell in build_up_row:
        if cell.value == "학번":
            # 5번 행의 전체 열중에 "학번"이있는 열: k
            column_with_student_id = cell.column_letter
        if cell.value == "참여자명":
            column_with_student_name = cell.column_letter
        if cell.value == "처리상태":
            column_with_student_acknowledgment = cell.column_letter
        if cell.value == "신청구분":
            column_with_student_application_category = cell.column_letter

    # 전산원 시트의 모든 학생 탐색
    for build_up_row in range(build_up_base_row, build_up_max_row + 1):
        # student id
        value_in_id = workReadSheet[str(
            column_with_student_id) + str(build_up_row)].value

        # student name
        value_in_name = workReadSheet[str(
            column_with_student_name) + str(build_up_row)].value

        # 처리상태
        value_in_acknowledgment = workReadSheet[str(
            column_with_student_acknowledgment) + str(build_up_row)].value

        # 신청구분
        value_in_application_category = workReadSheet[str(
            column_with_student_application_category) + str(build_up_row)].value

        # 실제 액셀에 쓰기
        writeSheet['B' + str(build_up_row + 1)] = value_in_id
        writeSheet['C' + str(build_up_row + 1)] = value_in_name
        writeSheet['D' + str(build_up_row + 1)] = value_in_acknowledgment
        writeSheet['A' + str(build_up_row + 1)] = value_in_application_category

        # build_up 시트의3학생시작은3이고 실제4시작은4이다 그래서 +1.을해야한다.
        developer_sheet_base_row_wand_write = build_up_row + 1

        # ====================================================
        # 직업탐색과미래설계
        search_want_sheet(readExcelFile, name_sheet=직업탐색과미래설계, num_row_search_student_id=2,
                          num_row_sheet_base=3, student_id_computer_sheet=value_in_id, writeSheet=writeSheet, write_sheet_row=developer_sheet_base_row_wand_write, str_day_want=신청경로, writeSheet_column="E")

        search_want_sheet(readExcelFile, name_sheet=직업탐색과미래설계, num_row_search_student_id=2,
                          num_row_sheet_base=3, student_id_computer_sheet=value_in_id, writeSheet=writeSheet, write_sheet_row=developer_sheet_base_row_wand_write, str_day_want=one상담, writeSheet_column="K")

        search_want_sheet(readExcelFile, name_sheet=직업탐색과미래설계, num_row_search_student_id=2,
                          num_row_sheet_base=3, student_id_computer_sheet=value_in_id, writeSheet=writeSheet, write_sheet_row=developer_sheet_base_row_wand_write, str_day_want=two상담, writeSheet_column="P")

        search_want_sheet(readExcelFile, name_sheet=직업탐색과미래설계, num_row_search_student_id=2,
                          num_row_sheet_base=3, student_id_computer_sheet=value_in_id, writeSheet=writeSheet, write_sheet_row=developer_sheet_base_row_wand_write, str_day_want=포트폴리오, writeSheet_column="V")

        # ====================================================
        # 신입생세미나
        search_want_sheet(readExcelFile, name_sheet=신입생세미나, num_row_search_student_id=2,
                          num_row_sheet_base=3, student_id_computer_sheet=value_in_id, writeSheet=writeSheet, write_sheet_row=developer_sheet_base_row_wand_write, str_day_want=신청경로, writeSheet_column="F")

        search_want_sheet(readExcelFile, name_sheet=신입생세미나, num_row_search_student_id=2,
                          num_row_sheet_base=3, student_id_computer_sheet=value_in_id, writeSheet=writeSheet, write_sheet_row=developer_sheet_base_row_wand_write, str_day_want=one상담, writeSheet_column="K")
        # ====================================================
        # 포폴워크숍
        search_want_sheet(readExcelFile, name_sheet=포트폴리오워크숍, num_row_search_student_id=2,
                          num_row_sheet_base=3, student_id_computer_sheet=value_in_id, writeSheet=writeSheet, write_sheet_row=developer_sheet_base_row_wand_write, str_day_want=신청경로, writeSheet_column="G")

        search_want_sheet(readExcelFile, name_sheet=포트폴리오워크숍, num_row_search_student_id=2,
                          num_row_sheet_base=3, student_id_computer_sheet=value_in_id, writeSheet=writeSheet, write_sheet_row=developer_sheet_base_row_wand_write, str_day_want=two상담, writeSheet_column="Q")

        search_want_sheet(readExcelFile, name_sheet=포트폴리오워크숍, num_row_search_student_id=2,
                          num_row_sheet_base=3, student_id_computer_sheet=value_in_id, writeSheet=writeSheet, write_sheet_row=developer_sheet_base_row_wand_write, str_day_want=포트폴리오, writeSheet_column="W")

        # ====================================================
        # 자기주도
        search_want_sheet(readExcelFile, name_sheet=자기주도, num_row_search_student_id=2,
                          num_row_sheet_base=3, student_id_computer_sheet=value_in_id, writeSheet=writeSheet, write_sheet_row=developer_sheet_base_row_wand_write, str_day_want=신청경로, writeSheet_column="H")

        search_want_sheet(readExcelFile, name_sheet=자기주도, num_row_search_student_id=2,
                          num_row_sheet_base=3, student_id_computer_sheet=value_in_id, writeSheet=writeSheet, write_sheet_row=developer_sheet_base_row_wand_write, str_day_want=one상담, writeSheet_column="M")

        search_want_sheet(readExcelFile, name_sheet=자기주도, num_row_search_student_id=2,
                          num_row_sheet_base=3, student_id_computer_sheet=value_in_id, writeSheet=writeSheet, write_sheet_row=developer_sheet_base_row_wand_write, str_day_want=two상담, writeSheet_column="R")

        search_want_sheet(readExcelFile, name_sheet=자기주도, num_row_search_student_id=2,
                          num_row_sheet_base=3, student_id_computer_sheet=value_in_id, writeSheet=writeSheet, write_sheet_row=developer_sheet_base_row_wand_write, str_day_want=포트폴리오, writeSheet_column="X")

        # ====================================================
        # 단기직무체험
        search_want_sheet(readExcelFile, name_sheet=단기직무체험, num_row_search_student_id=2,
                          num_row_sheet_base=3, student_id_computer_sheet=value_in_id, writeSheet=writeSheet, write_sheet_row=developer_sheet_base_row_wand_write, str_day_want=신청경로, writeSheet_column="I")

        search_want_sheet(readExcelFile, name_sheet=단기직무체험, num_row_search_student_id=2,
                          num_row_sheet_base=3, student_id_computer_sheet=value_in_id, writeSheet=writeSheet, write_sheet_row=developer_sheet_base_row_wand_write, str_day_want=one상담, writeSheet_column="N")

        search_want_sheet(readExcelFile, name_sheet=단기직무체험, num_row_search_student_id=2,
                          num_row_sheet_base=3, student_id_computer_sheet=value_in_id, writeSheet=writeSheet, write_sheet_row=developer_sheet_base_row_wand_write, str_day_want=two상담, writeSheet_column="S")

        search_want_sheet(readExcelFile, name_sheet=단기직무체험, num_row_search_student_id=2,
                          num_row_sheet_base=3, student_id_computer_sheet=value_in_id, writeSheet=writeSheet, write_sheet_row=developer_sheet_base_row_wand_write, str_day_want=포트폴리오, writeSheet_column="Y")

        # ====================================================
        # 진로집단상담
        search_want_sheet(readExcelFile, name_sheet=진로집단상담, num_row_search_student_id=2,
                          num_row_sheet_base=3, student_id_computer_sheet=value_in_id, writeSheet=writeSheet, write_sheet_row=developer_sheet_base_row_wand_write, str_day_want=신청경로, writeSheet_column="J")

        search_want_sheet(readExcelFile, name_sheet=진로집단상담, num_row_search_student_id=2,
                          num_row_sheet_base=3, student_id_computer_sheet=value_in_id, writeSheet=writeSheet, write_sheet_row=developer_sheet_base_row_wand_write, str_day_want=one상담, writeSheet_column="O")

        search_want_sheet(readExcelFile, name_sheet=진로집단상담, num_row_search_student_id=2,
                          num_row_sheet_base=3, student_id_computer_sheet=value_in_id, writeSheet=writeSheet, write_sheet_row=developer_sheet_base_row_wand_write, str_day_want=two상담, writeSheet_column="T")

        search_want_sheet(readExcelFile, name_sheet=진로집단상담, num_row_search_student_id=2,
                          num_row_sheet_base=3, student_id_computer_sheet=value_in_id, writeSheet=writeSheet, write_sheet_row=developer_sheet_base_row_wand_write, str_day_want=포트폴리오, writeSheet_column="Z")

        # ====================================================
        # 포트폴리오컨설팅
        search_want_sheet(readExcelFile, name_sheet=포트폴리오컨설팅, num_row_search_student_id=2,
                          num_row_sheet_base=3, student_id_computer_sheet=value_in_id, writeSheet=writeSheet, write_sheet_row=developer_sheet_base_row_wand_write, str_day_want=three상담, writeSheet_column="U")


def search_want_sheet(readExcelFile, name_sheet, num_row_search_student_id, num_row_sheet_base, student_id_computer_sheet, writeSheet, write_sheet_row, str_day_want, writeSheet_column):
    """_summary_

    Args:
        readExcelFile (Excel):읽고 싶은 엑셀파일
        name_sheet (string): 원하는 시트 이름
        num_row_search_student_id (num):그 시트에 원하는 열값을 찾고 싶은 행의 숫자('학번' 찾고 싶은 행을 입력 포트폴리오는 4번이다)
        num_row_sheet_base (num): 해당하는 시트의 학생들이 나오는 행번호이다(포트폴리오는 5 )
        student_id_computer_sheet (num):전산망 시트의 찾고싶은 해당 학생의 학번(value_in_id)
        writeSheet (string): 2번째 요구상황이 있는 엑셀 시트 (writeSheet)
        write_sheet_base_row (num): 전산망의 찾고싶은 학생이 있는 row 번호(write_sheet_base_row)
        str_day_want (string): 각 시트의 원하는 날짜가있는 문자열 (포트폴리오는 "상담일자/시간" 이다.)
        writeSheet_column (string):   쓰기를  원하는  시트의  열을 .적는것이다. ( 포트폴리오는 G.이다.)
    """

    sheet = readExcelFile[name_sheet]

    # sheet열의 개수 확인
    num_max_row = sheet.max_row

    # num_row_search_student_id번 행의 전체 열을 검사하여 "학번"이 있는 열 찾기
    # num_row_search_student_id번 행 선택
    row_sheet_search = sheet[num_row_search_student_id]
    sheet_column_with_student_id = None
    sheet_column_with_student_day = None
    sheet_student_day = None
    base_row = num_row_sheet_base

    for cell in row_sheet_search:
        if cell.value == "학번":
            # base_row 행의 전체 열중에 "학번"이있는 열
            sheet_column_with_student_id = cell.column_letter
        if cell.value == str_day_want:
            # base_row 행의 전체 열중에 "원하는문자열"이있는 열
            sheet_column_with_student_day = cell.column_letter

    for portfolio_base in range(base_row, num_max_row + 1):
        #  해당 시트에 있는 학생들의 학번
        portfolio_student_id = sheet[str(
            sheet_column_with_student_id) + str(portfolio_base)].value

        # 만약에 해당 시트에 있는 학생들의 학번과, 전상망 찾고싶은 학생이 있는 학번의 학번이 같으면
        if portfolio_student_id == student_id_computer_sheet:
            # 해당 시트의 일자
            sheet_student_day = sheet[str(
                sheet_column_with_student_day) + str(portfolio_base)].value

            # 2번째 요구사항 시트에 해당 시트 날짜 넣기
            #  해당  학번이  있는  열에  데이터 삽입
            writeSheet[writeSheet_column +
                       str(write_sheet_row)] = sheet_student_day
