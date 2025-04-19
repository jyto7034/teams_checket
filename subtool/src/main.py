# notifier.py
import sys
import os # os 모듈 임포트
import argparse
from plyer import notification
import traceback

def get_base_dir():
    """ 실행 파일/스크립트가 위치한 디렉토리를 반환합니다. """
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(sys.executable if getattr(sys, 'frozen', False) else __file__))


def show_notification(title, message):
    """지정된 제목과 메시지로 윈도우 토스트 알림을 표시합니다."""
    try:
        base_dir = get_base_dir()
        icon_filename = 'icon.ico'
        icon_path = os.path.join(base_dir, icon_filename)

        # 알림 호출 인자를 담을 딕셔너리
        notify_kwargs = {
            'title': title,
            'message': message,
            'app_name': "Checker 알림",
            'timeout': 10,
        }

        # 아이콘 파일이 실제로 존재하는지 확인
        if os.path.exists(icon_path):
            notify_kwargs['app_icon'] = icon_path
            print(f"아이콘 사용: {icon_path}") # 로그 추가 (디버깅용)
        else:
            print(f"아이콘 파일 '{icon_filename}'을(를) 찾을 수 없습니다: {icon_path}. 기본 아이콘 사용됨.") # 로그 추가

        # 딕셔너리를 사용하여 notify 함수 호출
        notification.notify(**notify_kwargs)

        print(f"알림 표시 성공: 제목='{title}'") # 성공 시 표준 출력으로 로그
        return True
    except Exception as e:
        # 오류 발생 시 표준 에러로 상세 정보 출력
        print(f"알림 표시 중 오류 발생: {e}", file=sys.stderr)
        traceback.print_exc(file=sys.stderr)
        return False

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="윈도우 토스트 알림을 표시합니다.")
    parser.add_argument("--title", required=True, help="알림 제목")
    parser.add_argument("--message", required=True, help="알림 본문 메시지")

    args = parser.parse_args()

    if show_notification(args.title, args.message):
        sys.exit(0) # 성공 시 종료 코드 0
    else:
        sys.exit(1) # 실패 시 종료 코드 1