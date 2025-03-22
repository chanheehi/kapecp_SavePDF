import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import Select
import os
import pandas as pd
import tkinter as tk
from tkinter import messagebox

print('\n')

def Show_message(message):
  root = tk.Tk()
  root.withdraw()  # 메인 윈도우 숨기기
  root.attributes("-topmost", True)
  messagebox.showwarning("message", message)
  root.destroy()

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
    # WebDriver 설정
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service)

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

        for index, data in enumerate(data_list):
            if index == 0:
                # 새로운 창 핸들로 전환
                new_window = [window for window in driver.window_handles if window != main_window][0]
                driver.switch_to.window(new_window)

                time.sleep(8)

                # '도축검사정보' Btn 클릭
                target_tab = driver.find_elements(By.XPATH, "//*[contains(@id, '_tab2')]")[0]
                target_tab.click()

                # '열람용 도축검사증명서 보기' Btn 클릭
                target_tab1 = driver.find_elements(By.XPATH, "//*[contains(@id, 'tab2')]")[1]
                showBtn = target_tab1.find_element(By.CLASS_NAME, 'print_btn')
                showBtn.click()

                # 새로운 창이 뜰 때까지 대기 (최대 10초 대기)
                WebDriverWait(driver, 60).until(EC.new_window_is_opened)
                WebDriverWait(driver, 60).until(
                    lambda d: d.execute_script('return document.readyState') == 'complete'
                )
                all_windows = driver.window_handles
                driver.switch_to.window(all_windows[-1])
                
                time.sleep(5)

                menuTable_div = driver.find_element(By.CLASS_NAME, 'report_menu_table_td_div')
                menu_btn = menuTable_div.find_elements(By.TAG_NAME, "button")

                # ActionChains를 사용하여 해당 버튼을 클릭
                actions = webdriver.ActionChains(driver)
                actions.move_to_element(menu_btn[0]).click().perform()

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

                time.sleep(2)
                # 마지막으로 열린 창을 닫고, 첫 번째 열린 창으로 제어를 다시 넘김
                driver.close()
                driver.switch_to.window(new_window)

                # =======================================================등급판정정보======================================================= #
                
                # '등급판정정보' Btn 클릭
                target_tab = driver.find_elements(By.XPATH, "//*[contains(@id, '_tab3')]")[0]
                target_tab.click()

                # '열람용 등급판정확인서 보기' Btn 클릭 해야함
                target_tab1 = driver.find_elements(By.XPATH, "//*[contains(@id, 'tab3')]")[1]
                showBtn = target_tab1.find_element(By.CLASS_NAME, 'print_btn')
                showBtn.click()

                # 새로운 창이 뜰 때까지 대기 (최대 10초 대기)
                WebDriverWait(driver, 60).until(EC.new_window_is_opened)
                WebDriverWait(driver, 60).until(
                    lambda d: d.execute_script('return document.readyState') == 'complete'
                )
                all_windows = driver.window_handles
                driver.switch_to.window(all_windows[-1])

            
                time.sleep(5)

                menuTable_div = driver.find_element(By.CLASS_NAME, 'report_menu_table_td_div')
                menu_btn = menuTable_div.find_elements(By.TAG_NAME, "button")

                # ActionChains를 사용하여 해당 버튼을 클릭
                actions = webdriver.ActionChains(driver)
                actions.move_to_element(menu_btn[0]).click().perform()


                
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

                time.sleep(2)
                # 마지막으로 열린 창을 닫고, 첫 번째 열린 창으로 제어를 다시 넘김
                driver.close()
            else:
                # input 요소가 상호작용 가능할 때까지 대기
                all_windows = driver.window_handles
                driver.switch_to.window(all_windows[-1])
                # 검색어 입력
                time.sleep(1)
                search_input = driver.find_element(By.CSS_SELECTOR, ".search_input")

                actions = webdriver.ActionChains(driver)
                actions.move_to_element(search_input).click().perform()
                search_input.clear()
                search_input.send_keys(f'00{data}')
                search_input.send_keys(Keys.RETURN)

                time.sleep(8)

                # '도축검사정보' Btn 클릭
                target_tab = driver.find_elements(By.XPATH, "//*[contains(@id, '_tab2')]")[0]
                target_tab.click()

                # '열람용 도축검사증명서 보기' Btn 클릭
                target_tab1 = driver.find_elements(By.XPATH, "//*[contains(@id, 'tab2')]")[1]
                showBtn = target_tab1.find_element(By.CLASS_NAME, 'print_btn')
                showBtn.click()

                # 새로운 창이 뜰 때까지 대기 (최대 10초 대기)
                WebDriverWait(driver, 60).until(EC.new_window_is_opened)
                WebDriverWait(driver, 60).until(
                    lambda d: d.execute_script('return document.readyState') == 'complete'
                )
                all_windows = driver.window_handles
                driver.switch_to.window(all_windows[-1])
                
                
                time.sleep(5)

                menuTable_div = driver.find_element(By.CLASS_NAME, 'report_menu_table_td_div')
                menu_btn = menuTable_div.find_elements(By.TAG_NAME, "button")

                # ActionChains를 사용하여 해당 버튼을 클릭
                actions = webdriver.ActionChains(driver)
                actions.move_to_element(menu_btn[0]).click().perform()

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

                time.sleep(2)
                # 마지막으로 열린 창을 닫고, 첫 번째 열린 창으로 제어를 다시 넘김
                driver.close()
                driver.switch_to.window(new_window)

                # =======================================================등급판정정보======================================================= #
                
                # '등급판정정보' Btn 클릭
                target_tab = driver.find_elements(By.XPATH, "//*[contains(@id, '_tab3')]")[0]
                target_tab.click()

                # '열람용 등급판정확인서 보기' Btn 클릭 해야함
                target_tab1 = driver.find_elements(By.XPATH, "//*[contains(@id, 'tab3')]")[1]
                showBtn = target_tab1.find_element(By.CLASS_NAME, 'print_btn')
                showBtn.click()

                # 새로운 창이 뜰 때까지 대기 (최대 10초 대기)
                WebDriverWait(driver, 60).until(EC.new_window_is_opened)
                WebDriverWait(driver, 60).until(
                    lambda d: d.execute_script('return document.readyState') == 'complete'
                )
                all_windows = driver.window_handles
                driver.switch_to.window(all_windows[-1])

            
                time.sleep(5)

                menuTable_div = driver.find_element(By.CLASS_NAME, 'report_menu_table_td_div')
                menu_btn = menuTable_div.find_elements(By.TAG_NAME, "button")

                # ActionChains를 사용하여 해당 버튼을 클릭
                actions = webdriver.ActionChains(driver)
                actions.move_to_element(menu_btn[0]).click().perform()


                
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

                time.sleep(2)
                # 마지막으로 열린 창을 닫고, 첫 번째 열린 창으로 제어를 다시 넘김
                driver.close()
    except Exception as e:
        print(f"오류 발생: {e}")

    finally:
        time.sleep(3)
        driver.quit()
    
    Show_message('작업 완료')

