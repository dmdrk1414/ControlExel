# 모듈을 현재 디렉토리에서 import합니다.
from src.call_request_excel.call_first_request_excel import make_first_request_excel
from src.call_request_excel.call_second_request_excel import make_second_request_excel


def run():
    while (True):
        print("프로그램 시작!!!")
        print("0: 프로그램 종료(exit)")
        print("1: 첫번째 요구사항 (데이터 결합)")
        print("2: 두번째 요구사항 (상담입력요청)")
        print("")
        num = input("번호을 입력해 주세요 (0, 1, 2, 3): ")

        # 요구사항을 실행한다
        if (num == "0"):
            print("프로그램 종료")
            exit()
        elif (num == "1"):
            print("첫번째 시작하기")
            make_first_request_excel()
            print("첫번째 끝!!\n")
        elif (num == "2"):
            print("두번째 시작하기")
            make_second_request_excel()
            print("두번째 끝!!\n")
        else:
            print("다시 입력해주세요 (0, 1, 2, 3)")


if __name__ == "__main__":
    run()
