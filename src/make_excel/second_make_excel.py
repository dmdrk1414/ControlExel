from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
from openpyxl import Workbook


def makeThirdExcel(workReadSheet, writeSheet, readExcelFile):
    """_summary_

    Args:
        workReadSheet (Excel sheet): 전산원 시트
        writeSheet (Ecxel sheet):  쓰고싶은  엑셀의 시트
        readExcelFile (Excel): 읽고 싶은 엑셀 
    """
    # 전산망의 기본적은 row 갯수
    numBaseRow = 5

    # E열의 개수 확인
    numOfRow = workReadSheet.max_row - numBaseRow

    build_sheet_str = "빌드업신청"
    build_sheet = readExcelFile[build_sheet_str]
    num_max_row_build_sheet = build_sheet.max_row
    bild_base_row = 3
    build_sheet_row_2 = build_sheet[2]  # 2번 행 선택

    # 5번 행의 전체 열을 검사하여 "학번"이 있는 열 찾기
    row_5 = workReadSheet[5]  # 5번 행 선택
    column_with_student_id = None
    column_with_student_name = None
    column_with_student_acknowledgment = None
    computer_base_row = 6
    computer_max_row = workReadSheet.max_row

    # write sheet 의  기본 row
    write_sheet_base_row_start_write = 4

    #  전산원에서  학변이있는  열을 찾는방법
    for cell in row_5:
        if cell.value == "학번":
            # 5번 행의 전체 열중에 "학번"이있는 열:
            computer_column_with_student_id = cell.column_letter

    #   빌드업에서 학변이있는  열을 찾는방법
    for cell in build_sheet_row_2:
        if cell.value == "학번":
            # 5번 행의 전체 열중에 "학번"이있는 열:
            build_column_with_student_id = cell.column_letter

    # 전산원 시트의 모든 학생 student id 탐색
    for computer_row in range(computer_base_row, computer_max_row + 1):
        #   전산원 학생의  한줄 학번
        computer_value_in_id = workReadSheet[str(
            computer_column_with_student_id) + str(computer_row)].value

        for build_base in range(bild_base_row, num_max_row_build_sheet + 1):
            #  build 시트에 있는 학생들의 학번
            build_student_id = build_sheet[str(
                build_column_with_student_id) + str(build_base)].value

            if (computer_value_in_id == build_student_id):
                #  전산망_sheet 특정 행의 모든 정보를 가져오기
                computer_row_data = []
                for cell in workReadSheet[computer_row]:
                    computer_row_data.append(cell.value)

            # 배열의 각 원소를  특정 행의 각 셀에 입력
            for i, value in enumerate(computer_row_data, start=1):
                cell = writeSheet.cell(
                    row=write_sheet_base_row_start_write, column=i)
                cell.value = value

                #  전상망의  데이터를 3 번째  시트에  적고  그다음줄에  적을  준비를 .한다.
                write_sheet_base_row_start_write = write_sheet_base_row_start_write + 1

        # 만약에 해당 시트에 있는 학생들의 학번과, 전상망 찾고싶은 학생이 있는 학번의 학번이 같으면
        # if build_student_id == value_in_id:
        #  해당  전산원의  모든  자료를 .쓰기.
