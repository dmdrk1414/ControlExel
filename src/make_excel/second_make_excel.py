from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
from openpyxl import Workbook


def makeSecondExcel(workReadSheet, writeSheet, readExcelFile):
    # 전산망의 기본적은 row 갯수
    numBaseRow = 5

    # E열의 개수 확인
    numOfRow = workReadSheet.max_row - numBaseRow

    # 5번 행의 전체 열을 검사하여 "학번"이 있는 열 찾기
    row_5 = workReadSheet[5]  # 5번 행 선택
    column_with_student_id = None
    column_with_student_name = None
    column_with_student_acknowledgment = None
    computer_base_row = 6
    computer_max_row = workReadSheet.max_row

    for cell in row_5:
        if cell.value == "학번":
            # 5번 행의 전체 열중에 "학번"이있는 열: k
            column_with_student_id = cell.column_letter
        if cell.value == "참여자명":
            column_with_student_name = cell.column_letter
        if cell.value == "처리상태":
            column_with_student_acknowledgment = cell.column_letter

    # 전산원 시트의 모든 학생 탐색 base_row = 6
    for computer_row in range(computer_base_row, computer_max_row + 1):
        value_in_id = workReadSheet[str(
            column_with_student_id) + str(computer_row)].value
        value_in_name = workReadSheet[str(
            column_with_student_name) + str(computer_row)].value
        value_in_acknowledgment = workReadSheet[str(
            column_with_student_acknowledgment) + str(computer_row)].value

        writeSheet['C' + str(computer_row - 2)] = value_in_id
        writeSheet['B' + str(computer_row - 2)] = value_in_name
        writeSheet['D' + str(computer_row - 2)] = value_in_acknowledgment

        # ====================================================
        # search_want_sheet(readExcelFile, name_sheet='(교과)포트폴리오DB', num_row_search_student_id=4,
        #                   num_row_sheet_base=5, student_id_computer_sheet=value_in_id, writeSheet=writeSheet, computer_row=computer_row, str_day_want="상담일자/시간", writeSheet_column="G")

        search_want_sheet(readExcelFile, name_sheet='신입생세미나', num_row_search_student_id=1,
                          num_row_sheet_base=2, student_id_computer_sheet=value_in_id, writeSheet=writeSheet, computer_row=computer_row, str_day_want="참여신청일자", writeSheet_column="F")

        search_want_sheet(readExcelFile, name_sheet='포폴워크숍', num_row_search_student_id=2,
                          num_row_sheet_base=3, student_id_computer_sheet=value_in_id, writeSheet=writeSheet, computer_row=computer_row, str_day_want="참여신청일자", writeSheet_column="G")

        search_want_sheet(readExcelFile, name_sheet='자기주도', num_row_search_student_id=2,
                          num_row_sheet_base=3, student_id_computer_sheet=value_in_id, writeSheet=writeSheet, computer_row=computer_row, str_day_want="상담일자", writeSheet_column="I")

        search_want_sheet(readExcelFile, name_sheet='단기직무체험', num_row_search_student_id=2,
                          num_row_sheet_base=3, student_id_computer_sheet=value_in_id, writeSheet=writeSheet, computer_row=computer_row, str_day_want="상담일자", writeSheet_column="H")

        search_want_sheet(readExcelFile, name_sheet='진로집단상담', num_row_search_student_id=2,
                          num_row_sheet_base=3, student_id_computer_sheet=value_in_id, writeSheet=writeSheet, computer_row=computer_row, str_day_want="상담일자", writeSheet_column="J")


def search_want_sheet(readExcelFile, name_sheet, num_row_search_student_id, num_row_sheet_base, student_id_computer_sheet, writeSheet, computer_row, str_day_want, writeSheet_column):
    """_summary_

    Args:
        readExcelFile (Excel):읽고 싶은 엑셀파일
        name_sheet (string): 원하는 시트 이름
        num_row_search_student_id (num):그 시트에 원하는 열값을 찾고 싶은 행의 숫자('학번' 찾고 싶은 행을 입력 포트폴리오는 4번이다)
        num_row_sheet_base (num): 해당하는 시트의 학생들이 나오는 행번호이다(포트폴리오는 5 )
        student_id_computer_sheet (num):전산망 시트의 찾고싶은 해당 학생의 학번(value_in_id)
        writeSheet (string): 2번째 요구상황이 있는 엑셀 시트 (writeSheet)
        computer_row (num): 전산망의 찾고싶은 학생이 있는 row 번호(computer_row)
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
            # base_row 행의 전체 열중에 "상담일자/시간"이있는 열
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
            writeSheet[writeSheet_column +
                       str(computer_row - 2)] = sheet_student_day
