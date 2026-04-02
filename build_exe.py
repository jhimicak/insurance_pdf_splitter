import os
import subprocess
import sys
import customtkinter

def build():
    # 1. Get customtkinter path
    ctk_path = os.path.dirname(customtkinter.__file__)
    
    # 2. Command construction
    # --noconsole: Hide terminal window
    # --onefile: Bundle everything into a single .exe
    # --add-data: Include customtkinter theme/icon files
    # separator is ; for Windows
    cmd = [
        "pyinstaller",
        "--noconsole",        # 다시 콘솔 창을 숨깁니다 (실행 속도에 영향)
        "--onedir",           # --onefile 대신 --onedir를 사용하여 압축 해제 과정을 없앱니다 (압도적으로 빠름)
        f"--add-data={ctk_path};customtkinter/",
        "--name=보험료_연말정산_시스템",
        "--clean",
        "app.py"
    ]
    
    print("--------------------------------------------------")
    print("EXE 변환을 시작합니다. (PyInstaller)")
    print(f"CustomTkinter 경로: {ctk_path}")
    print("잠시만 기다려 주세요... 이 작업은 1~2분 정도 소요될 수 있습니다.")
    print("--------------------------------------------------")
    
    try:
        # 기존 파일이 열려있는지 확인 (PermissionError 방지)
        dist_file = os.path.join("dist", "보험료_연말정산_시스템.exe")
        if os.path.exists(dist_file):
            try:
                with open(dist_file, 'a'): pass
            except PermissionError:
                print("\n❌ 오류: '보험료_연말정산_시스템.exe'가 현재 실행 중입니다.")
                print("작업 관리자에서 해당 프로그램을 종료하거나 실행 중인 창을 닫고 다시 시도해 주세요.")
                return

        subprocess.check_call(cmd)
        print("\n" + "="*50)
        print("✅ 변환이 성공적으로 완료되었습니다!")
        print("📁 'dist/보험료_연말정산_시스템' 폴더 안의 '.exe' 파일을 실행하세요.")
        print("\n💡 [참고] --onedir 방식을 사용하여 이제 더블 클릭 시 즉시 실행됩니다.")
        print("🚀 배포 시에는 'dist' 폴더 안의 해당 폴더 전체를 압축해서 전달하세요.")
        print("="*50)
    except Exception as e:
        print(f"\n❌ 변환 중 오류가 발생했습니다: {str(e)}")
        print("백신 프로그램이 파일을 검사 중이거나, 권한 문제일 수 있습니다.")

if __name__ == "__main__":
    build()
