import os
import winreg
import win32com.client

def register_hwp_module():
    try:
        try: 
            hnc_path = r"SOFTWARE\HNC"
        except FileNotFoundError:
            try:
                hnc_path = r"SOFTWARE\Hnc"
            except FileNotFoundError:
                print("한글 레지스트리에 없음")
        # 레지스트리 경로 설정
        reg_path = hnc_path + r"\HwpAutomation\Modules"
        # 레지스트리 키 열기 (읽기/쓰기)
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_WRITE)
        except FileNotFoundError:
            # 레지스트리 키가 없다면 새로 생성
            key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, reg_path)

        # "FilePathCheckerModuleExample" 값 확인 후 설정
        try:
            value = winreg.QueryValueEx(key, "FilePathCheckerModuleExample")
        except FileNotFoundError:
            # 레지스트리에 값이 없으면 새로 추가
            module_path = os.path.join(os.getcwd(), "\files\FilePathCheckerModuleExample.dll")
            winreg.SetValueEx(key, "FilePathCheckerModuleExample", 0, winreg.REG_SZ, module_path)

        # 한글 오토메이션 모듈 등록
        hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")

        print("모듈이 성공적으로 등록되었습니다.")
    
    except Exception as e:
        print(f"오류 발생: {e}")

if __name__ == "__main__":
    register_hwp_module()
