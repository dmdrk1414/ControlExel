# 모듈을 현재 디렉토리에서 import합니다.
from src.call_request_excel.call_first_request_excel import make_first_request_excel
from src.call_request_excel.call_second_request_excel import make_second_request_excel


def run():
    # 요구사항을 실행한다
    # print("첫번째 시작하기")
    # make_first_request_excel()
    # print("첫번째 끝!!\n")

    print("두번째 시작하기")
    make_second_request_excel()
    print("두번째 끝!!\n")


if __name__ == "__main__":
    run()
