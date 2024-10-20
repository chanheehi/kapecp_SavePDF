import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import Select
import os
import pandas as pd
print('\n')

# 파일 다운로드 기다림 (다운로드 경로에서 파일 존재 확인)
def Wait_for_download(download_dir, timeout=60):
    end_time = time.time() + timeout
    while time.time() < end_time:
        # 다운로드 경로에 .crdownload 확장자가 없으면 다운로드가 완료된 것으로 간주
        if any(file.endswith('.crdownload') for file in os.listdir(download_dir)):
            time.sleep(1)  # 다운로드 중인 파일이 존재하면 잠시 대기
        else:
            # 다운로드 완료됨
            return True
    return False

if __name__ == "__main__":

    # xlsx 파일을 불러오기
    df = pd.read_excel('./이력번호.xlsx')
    data_list = df['이력번호'].tolist()

    # 다운로드 설정
    download_dir = "./"  # 원하는 다운로드 경로로 변경
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": download_dir,  # 다운로드 경로 지정
        "download.prompt_for_download": False,       # 다운로드 확인창 비활성화
        "directory_upgrade": True,
        "safebrowsing.enabled": True
    })
    # WebDriver 설정
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)

    # 창을 최대화
    driver.maximize_window()

    # 웹사이트 열기
    driver.get('https://www.ekape.or.kr/kapecp/ui/kapecp/index.html')

    # 3초 대기 후 현재 창 핸들 확인
    time.sleep(1)
    main_window = driver.current_window_handle

    # 모든 창 핸들 가져오기
    all_windows = driver.window_handles

    # 팝업 창이 있는지 확인하고, 팝업 창이 있으면 닫기
    for window in all_windows:
        if window != main_window:
            driver.switch_to.window(window)
            driver.close()

    # 메인 창으로 돌아오기
    driver.switch_to.window(main_window)

    # 입력창이 상호작용 가능할 때까지 최대 10초 대기
    try:
        # class가 nexainput인 input 태그를 찾음
        search_input = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "input.nexainput"))
        )
        
        # 마우스로 클릭하기 위한 ActionChains 설정
        actions = webdriver.ActionChains(driver)
        
        # 요소 위치에 마우스를 클릭
        actions.move_to_element(search_input).click().perform()
        search_input.clear()
        search_input.send_keys(f'00{data_list[0]}')

        # 엔터 키 입력으로 검색 실행
        search_input.send_keys(Keys.RETURN)

        # 새로운 창이 뜰 때까지 대기
        WebDriverWait(driver, 60).until(EC.new_window_is_opened)

        num = 0
        for index, data in enumerate(data_list):
            if index == 0:
                # 새로운 창 핸들로 전환
                new_window = [window for window in driver.window_handles if window != main_window][0]
                driver.switch_to.window(new_window)

                # 'WaitControl' 클래스가 있는 div 태그의 visibility 속성이 hidden이 될 때까지 대기
                WebDriverWait(driver, 60).until(
                    EC.visibility_of_element_located((By.CSS_SELECTOR, "div.WaitControl"))
                )

                # 'WaitControl'의 visibility가 hidden이 될 때까지 대기
                WebDriverWait(driver, 60).until(
                    lambda driver: driver.execute_script(
                        "return window.getComputedStyle(document.querySelector('div.WaitControl')).visibility"
                    ) == 'hidden'
                )

                time.sleep(2)

                tabpageElements = driver.find_elements(By.CLASS_NAME, 'TabpageControl')  # 'TabpageControl' 클래스를 가진 모든 요소 찾기
                num = len(tabpageElements)


                # =======================================================도축검사정보======================================================= #
                # id가 'mainframe.VFrameSet.frameLogin.popup.form.divContent.form.tab00.tabbutton_1'인 모든 요소 찾기
                tab_buttons = driver.find_elements(By.ID, "mainframe.VFrameSet.frameLogin.popup.form.divContent.form.tab00.tabbutton_1")
                # 첫 번째 요소 클릭
                tab_buttons[0].click()
                
                # id가 'mainframe.VFrameSet.frameLogin.popup.form.divContent.form.tab00.도축검사정보.form.divResult.form.divTitle.form.btn00'인 모든 요소 찾기
                btn_elements = driver.find_elements(By.ID, "mainframe.VFrameSet.frameLogin.popup.form.divContent.form.tab00.도축검사정보.form.divResult.form.divTitle.form.btn00")
                btn_elements[0].click()

                # 새로운 창이 뜰 때까지 대기 (최대 10초 대기)
                WebDriverWait(driver, 60).until(EC.new_window_is_opened)
                all_windows = driver.window_handles
                driver.switch_to.window(all_windows[-1])
                time.sleep(5)
            
                # 'iframe'으로 전환
                WebDriverWait(driver, 60).until(
                    EC.frame_to_be_available_and_switch_to_it((By.ID, "mainframe.VFrameSet.frameLogin.popup.markAnyPopup.form.ClipReport_WebBrowser"))
                )

                # 'WaitControl' 클래스의 div 요소가 hidden이 될 때까지 기다림
                WebDriverWait(driver, 60).until(
                    lambda driver: driver.execute_script(
                        "return window.getComputedStyle(document.querySelector('.report_menu_progress')).display"
                    ) == 'none'
                )
                
                # 버튼 요소를 찾고 'title' 속성 값 가져오기
                pdf_button = WebDriverWait(driver, 60).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "button[title='저장']"))
                )

                # ActionChains를 사용하여 해당 버튼을 클릭
                actions = webdriver.ActionChains(driver)
                actions.move_to_element(pdf_button).click().perform()

                # Select 요소를 찾음
                select_element = WebDriverWait(driver, 60).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "select[title='파일형식']"))
                )
                
                # 'PDF 저장(*.pdf)' 옵션을 선택 (visible text로 선택)
                select = Select(select_element)
                select.select_by_visible_text("PDF 저장(*.pdf)")

                # 입력창에 이력번호 입력
                input_element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "input[title='파일명']"))
                )
                input_element.clear()
                input_element.send_keys(f'00{data}a')
                
                # '저장' 버튼 클릭
                save_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'report_save_view_position report_view_box')]//button[@title='저장']"))
                )
                save_button.click()

                # 다운로드가 완료될 때까지 기다림
                if Wait_for_download(download_dir, timeout=120):
                    print("파일이 성공적으로 다운로드되었습니다.")
                else:
                    print("다운로드 실패 또는 시간 초과")

                time.sleep(3)
                # 마지막으로 열린 창을 닫고, 첫 번째 열린 창으로 제어를 다시 넘김
                driver.close()
                driver.switch_to.window(new_window)

                # =======================================================등급판정정보======================================================= #
                
                # id가 'mainframe.VFrameSet.frameLogin.popup.form.divContent.form.tab00.tabbutton_2'인 모든 요소 찾기
                tab_buttons = driver.find_elements(By.ID, "mainframe.VFrameSet.frameLogin.popup.form.divContent.form.tab00.tabbutton_2")
                # 첫 번째 요소 클릭
                tab_buttons[0].click()
                # id가 'mainframe.VFrameSet.frameLogin.popup.form.divContent.form.tab00.도축검사정보.form.divResult.form.divTitle.form.btn00'인 모든 요소 찾기
                btn_elements = driver.find_elements(By.ID, "mainframe.VFrameSet.frameLogin.popup.form.divContent.form.tab00.등급판정정보.form.divResult.form.divTitle.form.btn00")
                btn_elements[0].click()

                # 새로운 창이 뜰 때까지 대기 (최대 10초 대기)
                WebDriverWait(driver, 60).until(EC.new_window_is_opened)
                all_windows = driver.window_handles
                driver.switch_to.window(all_windows[-1])

            
                # 'iframe'으로 전환
                WebDriverWait(driver, 10).until(
                    EC.frame_to_be_available_and_switch_to_it((By.ID, "mainframe.VFrameSet.frameLogin.popup.markAnyPopup.form.ClipReport_WebBrowser"))
                )

                # 'WaitControl' 클래스의 div 요소가 hidden이 될 때까지 기다림
                WebDriverWait(driver, 60).until(
                    EC.invisibility_of_element_located((By.CLASS_NAME, "report_menu_progress"))
                )

                # 버튼 요소를 찾고 'title' 속성 값 가져오기
                pdf_button = WebDriverWait(driver, 60).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "button[title='저장']"))
                )

                # ActionChains를 사용하여 해당 버튼을 클릭
                actions = webdriver.ActionChains(driver)
                actions.move_to_element(pdf_button).click().perform()

                # Select 요소를 찾음
                select_element = WebDriverWait(driver, 60).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "select[title='파일형식']"))
                )
                
                # 'PDF 저장(*.pdf)' 옵션을 선택 (visible text로 선택)
                select = Select(select_element)
                select.select_by_visible_text("PDF 저장(*.pdf)")



                # 입력창에 이력번호 입력
                input_element = WebDriverWait(driver, 60).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "input[title='파일명']"))
                )
                input_element.clear()
                input_element.send_keys(f'00{data}b')
                
                # '저장' 버튼 클릭
                save_button = WebDriverWait(driver, 60).until(
                    EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'report_save_view_position report_view_box')]//button[@title='저장']"))
                )
                save_button.click()
                # 다운로드가 완료될 때까지 기다림
                if Wait_for_download(download_dir, timeout=120):
                    print("파일이 성공적으로 다운로드되었습니다.")
                else:
                    print("다운로드 실패 또는 시간 초과")

                time.sleep(3)
                # 마지막으로 열린 창을 닫고, 첫 번째 열린 창으로 제어를 다시 넘김
                driver.close()
            else:
                # input 요소가 상호작용 가능할 때까지 대기
                all_windows = driver.window_handles
                driver.switch_to.window(all_windows[-1])
                # 검색어 입력
                search_input = WebDriverWait(driver, 60).until(
                    EC.element_to_be_clickable((By.ID, "mainframe.VFrameSet.frameLogin.popup.form.divContent.form.edtSearchKeyword:input"))
                )
                actions = webdriver.ActionChains(driver)
                actions.move_to_element(search_input).click().perform()
                search_input.clear()
                search_input.send_keys(f'00{data}')
                search_input.send_keys(Keys.RETURN)


                # 'WaitControl' 클래스가 있는 div 태그의 visibility 속성이 hidden이 될 때까지 대기
                WebDriverWait(driver, 60).until(
                    EC.visibility_of_element_located((By.CSS_SELECTOR, "div.WaitControl"))
                )

                # 'WaitControl'의 visibility가 hidden이 될 때까지 대기
                WebDriverWait(driver, 60).until(
                    lambda driver: driver.execute_script(
                        "return window.getComputedStyle(document.querySelector('div.WaitControl')).visibility"
                    ) == 'hidden'
                )
                tabpageElements = driver.find_elements(By.CLASS_NAME, 'TabpageControl')  # 'TabpageControl' 클래스를 가진 모든 요소 찾기
                tempNum = num + 1
                num = num + len(tabpageElements)
                # id가 'mainframe.VFrameSet.frameLogin.popup.form.divContent.form.tab00.tabbutton_1'인 모든 요소 찾기
                tab_buttons = driver.find_elements(By.ID, f"mainframe.VFrameSet.frameLogin.popup.form.divContent.form.tab00.tabbutton_{tempNum}")
                # 첫 번째 요소 클릭
                tab_buttons[0].click()
                # id가 'mainframe.VFrameSet.frameLogin.popup.form.divContent.form.tab00.도축검사정보.form.divResult.form.divTitle.form.btn00'인 모든 요소 찾기
                btn_elements = driver.find_elements(By.ID, "mainframe.VFrameSet.frameLogin.popup.form.divContent.form.tab00.도축검사정보.form.divResult.form.divTitle.form.btn00")
                btn_elements[0].click()

                # 새로운 창이 뜰 때까지 대기 (최대 10초 대기)
                WebDriverWait(driver, 60).until(EC.new_window_is_opened)
                all_windows = driver.window_handles
                driver.switch_to.window(all_windows[-1])
                time.sleep(5)
            
                # 'iframe'으로 전환
                WebDriverWait(driver, 60).until(
                    EC.frame_to_be_available_and_switch_to_it((By.ID, "mainframe.VFrameSet.frameLogin.popup.markAnyPopup.form.ClipReport_WebBrowser"))
                )

                # 'WaitControl' 클래스의 div 요소가 hidden이 될 때까지 기다림
                WebDriverWait(driver, 60).until(
                    lambda driver: driver.execute_script(
                        "return window.getComputedStyle(document.querySelector('.report_menu_progress')).display"
                    ) == 'none'
                )
                
                # 버튼 요소를 찾고 'title' 속성 값 가져오기
                pdf_button = WebDriverWait(driver, 60).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "button[title='저장']"))
                )

                # ActionChains를 사용하여 해당 버튼을 클릭
                actions = webdriver.ActionChains(driver)
                actions.move_to_element(pdf_button).click().perform()

                # Select 요소를 찾음
                select_element = WebDriverWait(driver, 60).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "select[title='파일형식']"))
                )
                
                # 'PDF 저장(*.pdf)' 옵션을 선택 (visible text로 선택)
                select = Select(select_element)
                select.select_by_visible_text("PDF 저장(*.pdf)")

                # 입력창에 이력번호 입력
                input_element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "input[title='파일명']"))
                )
                input_element.clear()
                input_element.send_keys(f'00{data}a')
                
                # '저장' 버튼 클릭
                save_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'report_save_view_position report_view_box')]//button[@title='저장']"))
                )
                save_button.click()

                # 다운로드가 완료될 때까지 기다림
                if Wait_for_download(download_dir, timeout=120):
                    print("파일이 성공적으로 다운로드되었습니다.")
                else:
                    print("다운로드 실패 또는 시간 초과")

                time.sleep(3)
                # 마지막으로 열린 창을 닫고, 첫 번째 열린 창으로 제어를 다시 넘김
                driver.close()
                driver.switch_to.window(new_window)

                # =======================================================등급판정정보======================================================= #
                
                # id가 'mainframe.VFrameSet.frameLogin.popup.form.divContent.form.tab00.tabbutton_2'인 모든 요소 찾기
                tempNum = tempNum+1
                tab_buttons = driver.find_elements(By.ID, f"mainframe.VFrameSet.frameLogin.popup.form.divContent.form.tab00.tabbutton_{tempNum}")
                # 첫 번째 요소 클릭
                tab_buttons[0].click()
                # id가 'mainframe.VFrameSet.frameLogin.popup.form.divContent.form.tab00.도축검사정보.form.divResult.form.divTitle.form.btn00'인 모든 요소 찾기
                btn_elements = driver.find_elements(By.ID, "mainframe.VFrameSet.frameLogin.popup.form.divContent.form.tab00.등급판정정보.form.divResult.form.divTitle.form.btn00")
                btn_elements[0].click()

                # 새로운 창이 뜰 때까지 대기 (최대 10초 대기)
                WebDriverWait(driver, 60).until(EC.new_window_is_opened)
                all_windows = driver.window_handles
                driver.switch_to.window(all_windows[-1])

            
                # 'iframe'으로 전환
                WebDriverWait(driver, 10).until(
                    EC.frame_to_be_available_and_switch_to_it((By.ID, "mainframe.VFrameSet.frameLogin.popup.markAnyPopup.form.ClipReport_WebBrowser"))
                )

                # 'WaitControl' 클래스의 div 요소가 hidden이 될 때까지 기다림
                WebDriverWait(driver, 60).until(
                    EC.invisibility_of_element_located((By.CLASS_NAME, "report_menu_progress"))
                )

                # 버튼 요소를 찾고 'title' 속성 값 가져오기
                pdf_button = WebDriverWait(driver, 60).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "button[title='저장']"))
                )

                # ActionChains를 사용하여 해당 버튼을 클릭
                actions = webdriver.ActionChains(driver)
                actions.move_to_element(pdf_button).click().perform()

                # Select 요소를 찾음
                select_element = WebDriverWait(driver, 60).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "select[title='파일형식']"))
                )
                
                # 'PDF 저장(*.pdf)' 옵션을 선택 (visible text로 선택)
                select = Select(select_element)
                select.select_by_visible_text("PDF 저장(*.pdf)")



                # 입력창에 이력번호 입력
                input_element = WebDriverWait(driver, 60).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "input[title='파일명']"))
                )
                input_element.clear()
                input_element.send_keys(f'00{data}b')
                
                # '저장' 버튼 클릭
                save_button = WebDriverWait(driver, 60).until(
                    EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'report_save_view_position report_view_box')]//button[@title='저장']"))
                )
                save_button.click()
                # 다운로드가 완료될 때까지 기다림
                if Wait_for_download(download_dir, timeout=120):
                    print("파일이 성공적으로 다운로드되었습니다.")
                else:
                    print("다운로드 실패 또는 시간 초과")

                time.sleep(3)
                # 마지막으로 열린 창을 닫고, 첫 번째 열린 창으로 제어를 다시 넘김
                driver.close()


    except Exception as e:
        print(f"오류 발생: {e}")

    finally:
        # 10초 후 브라우저 닫기
        time.sleep(300)


