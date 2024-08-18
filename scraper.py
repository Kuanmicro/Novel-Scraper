import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pyautogui
import pyperclip
import time
import os
import keyboard
import winsound
import zipfile
from threading import Thread, Lock
import logging
from datetime import datetime
import re
import json
import queue


# 작업 목록 파일 경로 설정
task_list_path = os.path.join(os.path.expanduser("~"), "Desktop", "Scraper", "task_list.json")

# 로그 파일 저장 경로 설정
log_dir = os.path.join(os.path.expanduser("~"), "Desktop", "Scraper", "Log")
os.makedirs(log_dir, exist_ok=True)

# 로그 파일 이름 설정
log_filename = os.path.join(log_dir, f'scraper_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log')

# 로깅 설정
logging.basicConfig(filename=log_filename, level=logging.DEBUG, format='%(asctime)s %(message)s')
logging.info("로그 파일이 성공적으로 설정되었습니다.")

# PyAutoGUI fail-safe 비활성화
pyautogui.FAILSAFE = False

# 기본 좌표와 딜레이 설정
start_x, start_y = 37, 464
end_x, end_y = 1300, 1150
next_x, next_y = 1680, 765
initial_x, initial_y = 60, 20  # 초기 설정 좌표
back_x, back_y = 20, 65  # 뒤로가기 버튼 좌표
delay_start = 2
delay_end = 2
delay_next = 2
initial_delay = 0.5  # 초기 커맨드 딜레이
coordinate_delay = 3  # 좌표 설정 대기 시간 기본값
toggle_key = 'f4'  # 시작/중단 키

# 초기 저장 경로 (기본값)
base_save_path = r'C:\Users\Home\Desktop\소설'

# 전역 변수
stop_flag = False
scraper_thread = None
lock = Lock()

def update_log_safe(message):
    """ 메인 스레드에서 안전하게 로그를 업데이트하기 위한 함수 """
    gui_queue.put(message)

def process_queue():
    """ 큐에서 메세지를 가져와서 로그를 업데이트 """
    try:
        while True:
            message = gui_queue.get_nowait()
            update_log(message)
    except queue.Empty:
        pass
    root.after(100, process_queue)  # 100ms 후 다시 큐를 처리

def check_stop():
    """ 중단 요청이 있는지 확인하고, 중단이 요청되었으면 스크립트를 안전하게 종료 """
    global stop_flag, scraper_thread
    if stop_flag:
        logging.info("중단 요청이 감지되었습니다. 스크립트를 중단합니다.")
        update_log("중단 요청이 감지되었습니다. 스크립트를 중단합니다.")
        stop_flag = False
        scraper_thread = None
        raise KeyboardInterrupt("스크립트가 사용자에 의해 중단되었습니다.")

# 작업 목록 로드 함수
def load_task_list():
    if os.path.exists(task_list_path):
        with open(task_list_path, 'r', encoding='utf-8') as file:
            return json.load(file)
    return []

# 작업 목록 저장 함수
def save_task_list(task_list):
    with open(task_list_path, 'w', encoding='utf-8') as file:
        json.dump(task_list, file, ensure_ascii=False, indent=4)

# 작업 목록을 업데이트하고 GUI를 갱신하는 함수
def update_task_listbox(task_listbox, task_list):
    task_listbox.delete(0, tk.END)
    for task in task_list:
        task_listbox.insert(tk.END, task['title'])

# 작업 목록 GUI 열기 함수
def open_task_list_gui():
    task_list = load_task_list()

    task_list_window = tk.Toplevel(root)
    task_list_window.title("작업 목록 관리")

    # 리스트박스 및 스크롤바 설정
    scrollbar = tk.Scrollbar(task_list_window)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    task_listbox = tk.Listbox(task_list_window, selectmode=tk.SINGLE, yscrollcommand=scrollbar.set, height=15)
    task_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.config(command=task_listbox.yview)

    update_task_listbox(task_listbox, task_list)

    # 소설 제목 입력창
    entry_task_title = tk.Entry(task_list_window, width=30)
    entry_task_title.pack(fill=tk.X, padx=5, pady=5)

    # 선택된 항목을 얻는 함수
    def get_selected_task():
        try:
            index = task_listbox.curselection()[0]
            return task_list[index], index
        except IndexError:
            return None, None

    # 작업 추가 함수
    def add_task():
        title = entry_task_title.get().strip()
        if title:
            task_list.append({'title': title})
            save_task_list(task_list)
            update_task_listbox(task_listbox, task_list)
            entry_task_title.delete(0, tk.END)

    # 작업 삭제 함수
    def delete_task():
        task, index = get_selected_task()
        if task:
            del task_list[index]
            save_task_list(task_list)
            update_task_listbox(task_listbox, task_list)

    # 작업 순서 위로 함수
    def move_up_task():
        task, index = get_selected_task()
        if task and index > 0:
            task_list.insert(index - 1, task_list.pop(index))
            save_task_list(task_list)
            update_task_listbox(task_listbox, task_list)

    # 작업 순서 아래로 함수
    def move_down_task():
        task, index = get_selected_task()
        if task and index < len(task_list) - 1:
            task_list.insert(index + 1, task_list.pop(index))
            save_task_list(task_list)
            update_task_listbox(task_listbox, task_list)

    # 작업 목록 GUI에 버튼 추가
    tk.Button(task_list_window, text="추가", command=add_task).pack(fill=tk.X)
    tk.Button(task_list_window, text="삭제", command=delete_task).pack(fill=tk.X)
    tk.Button(task_list_window, text="△", command=move_up_task).pack(fill=tk.X)
    tk.Button(task_list_window, text="▽", command=move_down_task).pack(fill=tk.X)

# 다음 소설을 작업 목록에서 가져오는 함수
def get_next_task():
    task_list = load_task_list()
    if task_list:
        return task_list[0]  # 첫 번째 작업을 가져옴
    return None

# 탭 닫기 및 다음 작업으로 이동하는 함수
def close_tab_and_proceed():
    check_stop()  # 중단 요청 확인
    pyautogui.click((initial_x, initial_y))
    time.sleep(initial_delay)
    check_stop()  # 중단 요청 확인

    # Ctrl + W로 탭 닫기
    pyautogui.hotkey('ctrl', 'w')
    time.sleep(1)  # 탭이 닫히는 데 시간이 걸릴 수 있으므로 잠시 대기
    check_stop()  # 중단 요청 확인

    # 다음 작업으로 이동
    next_task = get_next_task()
    if next_task:
        combobox_title.set(next_task['title'])
        check_stop()  # 중단 요청 확인
        on_start()
    else:
        logging.info("모든 작업이 완료되었습니다.")


# 로그를 업데이트하는 함수
def update_log(message):
    log_text.config(state=tk.NORMAL)
    log_text.insert(tk.END, message + '\n')
    log_text.see(tk.END)
    log_text.config(state=tk.DISABLED)

# 초기 설정 함수
def initial_setup():
    try:
        check_stop()  # 중단 요청 확인
        pyautogui.click((initial_x, initial_y))
        time.sleep(initial_delay)
        check_stop()  # 중단 요청 확인

        pyautogui.hotkey('ctrl', '1')
        time.sleep(initial_delay)
        check_stop()  # 중단 요청 확인

        pyautogui.press('home')
        time.sleep(initial_delay)
        check_stop()  # 중단 요청 확인

    except Exception as e:
        logging.error(f"초기 설정 오류: {e}")
        update_log(f"초기 설정 오류: {e}")


# 파일 압축 함수
def zip_files(save_path, title, total_chapters):
    try:
        zip_filename = os.path.join(save_path, f"{title}.zip")
        with zipfile.ZipFile(zip_filename, 'w') as zipf:
            for chapter in range(1, total_chapters + 1):
                check_stop()  # 중단 요청 확인
                file_name = f'{title} {chapter}화.txt'
                file_path = os.path.join(save_path, file_name)
                zipf.write(file_path, os.path.basename(file_path))

        logging.info(f"모든 파일이 압축되었습니다: {zip_filename}")
        update_log(f"모든 파일이 압축되었습니다: {zip_filename}")
    except Exception as e:
        logging.error(f"압축 오류: {e}")
        update_log(f"압축 오류: {e}")


# 화수 추출 함수 (세 번째 줄에서 추출)
def extract_chapter_number_and_total(text):
    # 텍스트의 줄들을 가져옵니다
    lines = text.splitlines()
    
    current_chapter = None
    total_chapters = None

    # 두 번째 줄과 세 번째 줄에서 "(현재 화/총 화)" 형식을 추출
    for i in [1, 2]:  # 두 번째 줄과 세 번째 줄 (0-indexed)
        if len(lines) > i:
            line = lines[i].strip()
            match = re.search(r'\((\d+)/(\d+)\)', line)
            if match:
                current_chapter = int(match.group(1))
                total_chapters = int(match.group(2))
                break  # 첫 번째로 매칭되는 것을 찾으면 중지
    
    return current_chapter, total_chapters


# 빠진 화수 찾기 함수
def find_missing_chapters(chapters_saved, total_chapters):
    all_chapters = set(range(1, total_chapters + 1))
    saved_chapters = set(chapters_saved)
    return sorted(all_chapters - saved_chapters)

# 작은 파일 찾기 함수
def find_small_files(title, save_path, total_chapters):
    small_files = []
    try:
        for chapter in range(1, total_chapters + 1):
            file_name = f'{title} {chapter}화.txt'
            file_path = os.path.join(save_path, file_name)
            if os.path.exists(file_path) and os.path.getsize(file_path) <= 1024:
                small_files.append(chapter)
        
        if small_files:
            logging.warning(f"크기가 작은 파일 발견: {small_files}")
            update_log(f"크기가 작은 파일 발견: {small_files}")
        else:
            logging.info("모든 파일의 크기가 정상입니다.")
            update_log("모든 파일의 크기가 정상입니다.")
    except Exception as e:
        logging.error(f"1KB 파일 검사 오류: {e}")
        update_log(f"1KB 파일 검사 오류: {e}")
    return small_files

# 파일 검사 및 알림 함수
def check_files_and_notify():
    title = combobox_title.get().strip()
    total_chapters = entry_total_chapters.get().strip()

    if not title:
        messagebox.showerror("입력 오류", "제목을 입력하세요.")
        return

    save_path = os.path.join(base_save_path, title)

    # 총 화수가 입력되지 않은 경우 마지막 화에서 자동 추출
    if not total_chapters.isdigit():
        files = os.listdir(save_path)
        chapter_numbers = []
        
        # 파일 이름에서 화수를 추출
        for file in files:
            match = re.match(rf'{re.escape(title)} (\d+)화\.txt', file)
            if match:
                chapter_numbers.append(int(match.group(1)))
        
        if chapter_numbers:
            last_chapter = max(chapter_numbers)
            last_chapter_file = os.path.join(save_path, f'{title} {last_chapter}화.txt')
            
            with open(last_chapter_file, 'r', encoding='utf-8') as file:
                text = file.read()
                _, total_chapters = extract_chapter_number_and_total(text)
            
            if not total_chapters:
                messagebox.showerror("추출 오류", "총 화수를 추출할 수 없습니다.")
                return
        else:
            messagebox.showerror("입력 오류", "저장된 파일에서 화수를 찾을 수 없습니다.")
            return

    total_chapters = int(total_chapters)

    find_small_files(title, save_path, total_chapters)

    chapters_saved = []
    for chapter in range(1, total_chapters + 1):
        file_name = f'{title} {chapter}화.txt'
        if os.path.exists(os.path.join(save_path, file_name)):
            chapters_saved.append(chapter)

    missing_chapters = find_missing_chapters(chapters_saved, total_chapters)
    if missing_chapters:
        logging.warning(f"빠진 화수: {missing_chapters}")
        update_log(f"빠진 화수: {missing_chapters}")
    else:
        logging.info("모든 화수가 정상적으로 저장되었습니다.")
        update_log("모든 화수가 정상적으로 저장되었습니다.")

# 스크래핑 시작 함수
def start_scraper(title, start_chapter, num_chapters_to_process, total_chapters, save_path, settings, one=False):
    global stop_flag, scraper_thread
    start_x, start_y, end_x, end_y, next_x, next_y, delay_start, delay_end, delay_next, coordinate_delay, toggle_key, initial_x, initial_y, back_x, back_y = settings
    
    # start_chapter를 정수로 변환 (변환이 불가능하면 기본값 설정)
    try:
        start_chapter = int(start_chapter) if start_chapter else None
    except ValueError:
        logging.error(f"시작 화수 값이 유효하지 않습니다: {start_chapter}")
        update_log(f"시작 화수 값이 유효하지 않습니다: {start_chapter}")
        return

    save_path = os.path.join(save_path, title)
    os.makedirs(save_path, exist_ok=True)
    
    files = os.listdir(save_path)
    chapter_numbers = [int(re.match(rf'{re.escape(title)} (\d+)화\.txt', file).group(1)) for file in files if re.match(rf'{re.escape(title)} (\d+)화\.txt', file)]

    if chapter_numbers:
        previous_chapter = max(chapter_numbers)
    else:
        previous_chapter = start_chapter - 1 if start_chapter is not None else 0

    start_chapter = int(start_chapter) if start_chapter else previous_chapter + 1
    total_chapters = int(total_chapters) if total_chapters else None

    chapters_saved = []
    previous_text = None

    try:
        initial_setup()
        check_stop()  # 중단 요청 확인

        for chapter in range(start_chapter, start_chapter + (num_chapters_to_process if num_chapters_to_process else (total_chapters if total_chapters else 99999))):
            check_stop()  # 중단 요청 확인

            retry_count = 0
            max_retries = 3
            successful_copy = False

            while retry_count < max_retries and not successful_copy:
                try:
                    check_stop()  # 중단 요청 확인

                    # 페이지 상단으로 이동
                    pyautogui.press('home')
                    time.sleep(2)
                    check_stop()  # 중단 요청 확인

                    pyautogui.click((start_x, start_y))
                    time.sleep(delay_start)
                    check_stop()  # 중단 요청 확인
                    
                    pyautogui.press('end')
                    time.sleep(delay_end)
                    check_stop()  # 중단 요청 확인

                    pyautogui.keyDown('shift')
                    pyautogui.click((end_x, end_y))
                    pyautogui.keyUp('shift')
                    time.sleep(1)
                    check_stop()  # 중단 요청 확인

                    pyautogui.hotkey('ctrl', 'c')
                    time.sleep(1)
                    check_stop()  # 중단 요청 확인

                    text = pyperclip.paste()
                    current_chapter, extracted_total_chapters = extract_chapter_number_and_total(text)

                    if one:
                        file_name = f'{title} {current_chapter}화.txt'
                        with open(os.path.join(save_path, file_name), 'w', encoding='utf-8') as file:
                            file.write(text)

                        logging.info(f"현재 진행 상황: {current_chapter} 화 저장 완료 (One 모드)")
                        update_log(f"현재 진행 상황: {current_chapter} 화 저장 완료 (One 모드)")
                        successful_copy = True
                        break

                    # 화수가 추출되지 않으면 새로고침 후 재시도
                    if current_chapter is None or len(text) < 100 or text == previous_text:
                        logging.error(f"복사 오류 또는 화수 추출 실패: {text[:30]}")
                        update_log(f"복사 오류 또는 화수 추출 실패: {text[:30]}")
                        retry_count += 1

                        if retry_count < max_retries:
                            if current_chapter is None:
                                logging.info("화수가 추출되지 않았습니다. 페이지를 새로고침합니다.")
                                update_log("화수가 추출되지 않았습니다. 페이지를 새로고침합니다.")
                                pyautogui.press('f5')
                                time.sleep(7)
                                check_stop()  # 중단 요청 확인

                                continue  # 재시도

                            if len(text) < 100 or current_chapter == previous_chapter:
                                pyautogui.press('f5')
                                logging.info(f"페이지 새로고침 시도 {retry_count}회 - 현재 화수: {current_chapter}")
                                update_log(f"페이지 새로고침 시도 {retry_count}회 - 현재 화수: {current_chapter}")
                                time.sleep(7)
                                check_stop()  # 중단 요청 확인

                                pyautogui.press('home')
                                time.sleep(2)
                                check_stop()  # 중단 요청 확인

                                pyautogui.click((start_x, start_y))
                                time.sleep(delay_start)
                                check_stop()  # 중단 요청 확인

                                pyautogui.press('end')
                                time.sleep(delay_end)
                                check_stop()  # 중단 요청 확인

                                pyautogui.keyDown('shift')
                                pyautogui.click((end_x, end_y))
                                pyautogui.keyUp('shift')
                                time.sleep(1)
                                check_stop()  # 중단 요청 확인

                                pyautogui.hotkey('ctrl', 'c')
                                time.sleep(1)
                                check_stop()  # 중단 요청 확인

                                text = pyperclip.paste()
                                current_chapter, extracted_total_chapters = extract_chapter_number_and_total(text)

                                if text == previous_text:
                                    logging.info("동일한 텍스트가 감지되어 다음 화로 이동합니다.")
                                    update_log("동일한 텍스트가 감지되어 다음 화로 이동합니다.")
                                
                                    pyautogui.click((next_x, next_y))
                                    time.sleep(delay_next)
                                    check_stop()  # 중단 요청 확인

                                    continue

                            elif current_chapter is not None and current_chapter > previous_chapter + 1:
                                pyautogui.click((back_x, back_y))
                                logging.info(f"뒤로 가기 실행 - 현재 화수: {current_chapter}, 이전 화수: {previous_chapter}")
                                update_log(f"뒤로 가기 실행 - 현재 화수: {current_chapter}, 이전 화수: {previous_chapter}")
                                time.sleep(5)
                                check_stop()  # 중단 요청 확인
                            
                            pyautogui.click((initial_x, initial_y))
                            logging.info("페이지 특정 위치를 클릭하여 포커스를 설정했습니다.")
                            update_log("페이지 특정 위치를 클릭하여 포커스를 설정했습니다.")
                            time.sleep(1)
                            check_stop()  # 중단 요청 확인

                            pyautogui.press('home')
                            logging.info("홈 키를 눌러 페이지 상단으로 이동했습니다.")
                            update_log("홈 키를 눌러 페이지 상단으로 이동했습니다.")
                            time.sleep(3)
                            check_stop()  # 중단 요청 확인

                        else:
                            logging.error(f"복사 및 추출 실패로 챕터 {current_chapter}를 건너뜁니다.")
                            update_log(f"복사 및 추출 실패로 챕터 {current_chapter}를 건너뜁니다.")
                            break
                    else:
                        successful_copy = True
                        previous_text = text
                        previous_chapter = current_chapter
                        chapters_saved.append(current_chapter)
                        file_name = f'{title} {current_chapter}화.txt'
                        with open(os.path.join(save_path, file_name), 'w', encoding='utf-8') as file:
                            file.write(text)

                        logging.info(f"현재 진행 상황: {current_chapter}/{extracted_total_chapters} 화 저장 완료")
                        update_log(f"현재 진행 상황: {current_chapter}/{extracted_total_chapters} 화 저장 완료")

                        if not total_chapters or total_chapters != extracted_total_chapters:
                            total_chapters = extracted_total_chapters
                            logging.info(f"총 화수가 {total_chapters}로 업데이트되었습니다.")
                            update_log(f"총 화수가 {total_chapters}로 업데이트되었습니다.")

                        if current_chapter == total_chapters:
                            logging.info("마지막 화에 도달했습니다. 스크립트를 중단하고 파일을 압축합니다.")
                            update_log("마지막 화에 도달했습니다. 스크립트를 중단하고 파일을 압축합니다.")
                            zip_files(save_path, title, total_chapters)

                            try:
                                logging.info("알람을 재생합니다...")
                                winsound.Beep(1000, 700)  # 1000Hz, 0.7초 동안 소리 재생
                                logging.info("알람 소리 재생 완료.")
                                update_log("알람 소리 재생 완료.")
                            except RuntimeError as e:
                                logging.error(f"알람 소리 재생 오류 발생: {e}")
                                update_log(f"알람 소리 재생 오류 발생: {e}")

                            scraper_thread = None
                            return

                    pyautogui.click((next_x, next_y))
                    time.sleep(delay_next)
                    check_stop()  # 중단 요청 확인

                except Exception as e:
                    logging.error(f"챕터 {current_chapter}에서 오류 발생: {e}")
                    update_log(f"챕터 {current_chapter}에서 오류 발생: {e}")
                    retry_count += 1

            if not successful_copy:
                pyautogui.click((next_x, next_y))
                time.sleep(delay_next)
                logging.error(f"다음 화로 이동했습니다.")
                update_log(f"다음 화로 이동했습니다.")
                check_stop()  # 중단 요청 확인
                continue

        if current_chapter == total_chapters:
            missing_chapters = find_missing_chapters(chapters_saved, total_chapters)
            small_files = find_small_files(title, save_path, total_chapters)
            
            if missing_chapters or small_files:
                logging.warning(f"작업 완료 후 오류 발견. 누락된 화수: {missing_chapters}, 작은 파일: {small_files}")
                update_log(f"작업 완료 후 오류 발견. 누락된 화수: {missing_chapters}, 작은 파일: {small_files}")
            else:
                logging.info("모든 화수가 정상적으로 저장되었습니다. 파일 압축을 시작합니다.")
                update_log("모든 화수가 정상적으로 저장되었습니다. 파일 압축을 시작합니다.")
                zip_files(save_path, title, total_chapters)

                try:
                    logging.info("알람을 재생합니다...")
                    winsound.Beep(1000, 700)  # 1000Hz, 0.7초 동안 소리 재생
                    logging.info("알람 소리 재생 완료.")
                    update_log("알람 소리 재생 완료.")
                except RuntimeError as e:
                    logging.error(f"알람 소리 재생 오류 발생: {e}")
                    update_log(f"알람 소리 재생 오류 발생: {e}")

    except KeyboardInterrupt:
        logging.info(f"KeyboardInterrupt: 스크립트가 사용자에 의해 중단되었습니다. 마지막 장: {current_chapter}")
        update_log(f"KeyboardInterrupt: 스크립트가 사용자에 의해 중단되었습니다. 마지막 장: {current_chapter}")
    except Exception as e:
        logging.error(f"오류 발생: {e}")
        update_log(f"오류 발생: {e}")
    finally:
        scraper_thread = None
        stop_flag = False

# 좌표 설정 함수
def get_coordinates(x_var, y_var, label, delay):
    messagebox.showinfo("좌표 선택", f"원하는 위치에 마우스를 놓고 {delay}초 기다리세요.")
    time.sleep(delay)
    x, y = pyautogui.position()
    x_var.set(x)
    y_var.set(y)
    label.config(text=f"({x}, {y})")
    messagebox.showinfo("좌표 기록됨", f"좌표가 기록되었습니다: ({x}, {y})")

# 설정 창 열기
def open_settings():
    settings_window = tk.Toplevel(root)
    settings_window.title("설정")

    # 라벨과 입력 필드 간격을 줄이기 위해 padx와 pady 값을 줄임
    tk.Label(settings_window, text="시작 부분 좌표").grid(row=0, column=0, padx=(2, 5), pady=5, sticky="e")
    tk.Label(settings_window, text="끝 부분 좌표").grid(row=1, column=0, padx=(2, 5), pady=5, sticky="e")
    tk.Label(settings_window, text="다음 화 버튼 좌표").grid(row=2, column=0, padx=(2, 5), pady=5, sticky="e")
    tk.Label(settings_window, text="초기 설정 좌표").grid(row=3, column=0, padx=(2, 5), pady=5, sticky="e")
    tk.Label(settings_window, text="뒤로가기 버튼 좌표").grid(row=4, column=0, padx=(2, 5), pady=5, sticky="e")
    tk.Label(settings_window, text="시작 딜레이 (초)").grid(row=5, column=0, padx=(2, 5), pady=5, sticky="e")
    tk.Label(settings_window, text="끝 딜레이 (초)").grid(row=6, column=0, padx=(2, 5), pady=5, sticky="e")
    tk.Label(settings_window, text="다음 화 딜레이 (초)").grid(row=7, column=0, padx=(2, 5), pady=5, sticky="e")
    tk.Label(settings_window, text="좌표 설정 대기 시간 (초)").grid(row=8, column=0, padx=(2, 5), pady=5, sticky="e")
    tk.Label(settings_window, text="시작/중단 키").grid(row=9, column=0, padx=(2, 5), pady=5, sticky="e")
    
    start_x_var = tk.IntVar(value=start_x)
    start_y_var = tk.IntVar(value=start_y)
    end_x_var = tk.IntVar(value=end_x)
    end_y_var = tk.IntVar(value=end_y)
    next_x_var = tk.IntVar(value=next_x)
    next_y_var = tk.IntVar(value=next_y)
    initial_x_var = tk.IntVar(value=initial_x)
    initial_y_var = tk.IntVar(value=initial_y)
    back_x_var = tk.IntVar(value=back_x)
    back_y_var = tk.IntVar(value=back_y)
    delay_start_var = tk.DoubleVar(value=delay_start)
    delay_end_var = tk.DoubleVar(value=delay_end)
    delay_next_var = tk.DoubleVar(value=delay_next)
    coordinate_delay_var = tk.DoubleVar(value=coordinate_delay)

    tk.Entry(settings_window, textvariable=start_x_var, width=5).grid(row=0, column=1, padx=(2, 1), pady=5, sticky="e")
    tk.Entry(settings_window, textvariable=start_y_var, width=5).grid(row=0, column=2, padx=(1, 5), pady=5, sticky="w")
    tk.Entry(settings_window, textvariable=end_x_var, width=5).grid(row=1, column=1, padx=(2, 1), pady=5, sticky="e")
    tk.Entry(settings_window, textvariable=end_y_var, width=5).grid(row=1, column=2, padx=(1, 5), pady=5, sticky="w")
    tk.Entry(settings_window, textvariable=next_x_var, width=5).grid(row=2, column=1, padx=(2, 1), pady=5, sticky="e")
    tk.Entry(settings_window, textvariable=next_y_var, width=5).grid(row=2, column=2, padx=(1, 5), pady=5, sticky="w")
    tk.Entry(settings_window, textvariable=initial_x_var, width=5).grid(row=3, column=1, padx=(2, 1), pady=5, sticky="e")
    tk.Entry(settings_window, textvariable=initial_y_var, width=5).grid(row=3, column=2, padx=(1, 5), pady=5, sticky="w")
    tk.Entry(settings_window, textvariable=back_x_var, width=5).grid(row=4, column=1, padx=(2, 1), pady=5, sticky="e")
    tk.Entry(settings_window, textvariable=back_y_var, width=5).grid(row=4, column=2, padx=(1, 5), pady=5, sticky="w")
    tk.Entry(settings_window, textvariable=delay_start_var, width=10).grid(row=5, column=1, padx=(2, 5), pady=5, sticky="w")
    tk.Entry(settings_window, textvariable=delay_end_var, width=10).grid(row=6, column=1, padx=(2, 5), pady=5, sticky="w")
    tk.Entry(settings_window, textvariable=delay_next_var, width=10).grid(row=7, column=1, padx=(2, 5), pady=5, sticky="w")
    tk.Entry(settings_window, textvariable=coordinate_delay_var, width=10).grid(row=8, column=1, padx=(2, 5), pady=5, sticky="w")

    toggle_key_entry = tk.Entry(settings_window, width=10)
    toggle_key_entry.insert(0, toggle_key)
    toggle_key_entry.grid(row=9, column=1, padx=(2, 5), pady=5, sticky="w")

    save_path_label = tk.Label(settings_window, text=base_save_path)
    save_path_label.grid(row=10, column=1, padx=(2, 5), pady=5, sticky="w")

    tk.Button(settings_window, text='시작 부분 좌표 설정', command=lambda: get_coordinates(start_x_var, start_y_var, None, coordinate_delay_var.get())).grid(row=0, column=3, padx=5, pady=5, sticky="w")
    tk.Button(settings_window, text='끝 부분 좌표 설정', command=lambda: get_coordinates(end_x_var, end_y_var, None, coordinate_delay_var.get())).grid(row=1, column=3, padx=5, pady=5, sticky="w")
    tk.Button(settings_window, text='다음 화 버튼 좌표 설정', command=lambda: get_coordinates(next_x_var, next_y_var, None, coordinate_delay_var.get())).grid(row=2, column=3, padx=5, pady=5, sticky="w")
    tk.Button(settings_window, text='초기 설정 좌표 설정', command=lambda: get_coordinates(initial_x_var, initial_y_var, None, coordinate_delay_var.get())).grid(row=3, column=3, padx=5, pady=5, sticky="w")
    tk.Button(settings_window, text='뒤로가기 버튼 좌표 설정', command=lambda: get_coordinates(back_x_var, back_y_var, None, coordinate_delay_var.get())).grid(row=4, column=3, padx=5, pady=5, sticky="w")

    def choose_save_path():
        global base_save_path
        directory = filedialog.askdirectory()
        if directory:
            base_save_path = directory
            save_path_label.config(text=directory)

    tk.Button(settings_window, text='저장 경로 선택', command=choose_save_path).grid(row=10, column=2, padx=5, pady=5, sticky="w")

    def save_settings():
        global start_x, start_y, end_x, end_y, next_x, next_y, delay_start, delay_end, delay_next, coordinate_delay, toggle_key, initial_x, initial_y, back_x, back_y
        start_x, start_y = start_x_var.get(), start_y_var.get()
        end_x, end_y = end_x_var.get(), end_y_var.get()
        next_x, next_y = next_x_var.get(), next_y_var.get()
        initial_x, initial_y = initial_x_var.get(), initial_y_var.get()
        back_x, back_y = back_x_var.get(), back_y_var.get()
        delay_start, delay_end, delay_next = delay_start_var.get(), delay_end_var.get(), delay_next_var.get()
        coordinate_delay = coordinate_delay_var.get()
        toggle_key = toggle_key_entry.get()
        settings_window.destroy()

        register_key_events()

    tk.Button(settings_window, text="저장", command=save_settings).grid(row=11, column=1, padx=5, pady=5, sticky="w")

# 키 이벤트 등록
def register_key_events():
    keyboard.unhook_all()
    keyboard.on_press_key(toggle_key, on_toggle_key_press)
    keyboard.on_press_key('f2', lambda e: run_single_chapter())  # F2 키로 1화 추출 실행

# 시작 버튼 함수
# 기존 스크래퍼 시작 함수 수정
def on_start():
    global scraper_thread, stop_flag
    stop_flag = False

    title = combobox_title.get().strip()
    start_chapter = entry_start_chapter.get().strip()
    num_chapters_to_process = entry_num_chapters_to_process.get().strip()
    total_chapters = entry_total_chapters.get().strip()

    try:
        if not title:
            # 제목이 입력되지 않았으면 작업 목록에서 가져옴
            next_task = get_next_task()
            if next_task:
                title = next_task['title']
                combobox_title.set(title)
            else:
                messagebox.showerror("입력 오류", "제목을 입력하거나 작업 목록에 작업이 없습니다.")
                return

        # 이후의 기존 스크립트 로직을 유지
        total_chapters = None if not total_chapters.isdigit() else int(total_chapters)
        start_chapter = None if not start_chapter else int(start_chapter)
        num_chapters_to_process = None if not num_chapters_to_process else int(num_chapters_to_process)
        
        settings = (start_x, start_y, end_x, end_y, next_x, next_y, delay_start, delay_end, delay_next, coordinate_delay, toggle_key, initial_x, initial_y, back_x, back_y)
        start_scraper(title, start_chapter, num_chapters_to_process, total_chapters, base_save_path, settings)

        # 작업 완료 후 탭 닫기 및 다음 작업으로 이동
        task_list = load_task_list()
        if task_list and title == task_list[0]['title']:
            del task_list[0]  # 첫 번째 작업을 제거
            save_task_list(task_list)
            close_tab_and_proceed()  # 탭을 닫고 다음 작업으로 이동

    except Exception as e:
        messagebox.showerror("오류 발생", f"오류가 발생했습니다: {e}")
        logging.error(f"스크립트 오류: {e}")
        update_log(f"스크립트 오류: {e}")
    finally:
        scraper_thread = None
        stop_flag = False
        register_key_events()
# 드롭다운 메뉴 업데이트
def update_dropdown():
    folders = [folder for folder in next(os.walk(base_save_path))[1]]
    combobox_title['values'] = folders

# 한 화만 진행하는 버튼 처리 함수
def run_single_chapter():
    global scraper_thread, stop_flag
    stop_flag = False

    title = combobox_title.get().strip()
    start_chapter = entry_start_chapter.get().strip()

    try:
        if not title:
            messagebox.showerror("입력 오류", "제목을 입력하세요.")
            return

        # One 기능을 위한 플래그 설정
        is_one_mode = True
        
        settings = (start_x, start_y, end_x, end_y, next_x, next_y, delay_start, delay_end, delay_next, coordinate_delay, toggle_key, initial_x, initial_y, back_x, back_y)
        
        # 1화만 실행하도록 설정
        start_scraper(title, start_chapter, 1, None, base_save_path, settings, is_one_mode)
    except Exception as e:
        messagebox.showerror("오류 발생", f"오류가 발생했습니다: {e}")
        logging.error(f"스크립트 오류: {e}")
        update_log(f"스크립트 오류: {e}")
    finally:
        scraper_thread = None
        stop_flag = False
        register_key_events()


# 중단 요청을 처리하는 함수 수정
def on_toggle_key_press(event):
    global stop_flag, scraper_thread
    if scraper_thread is None:  # 스크립트가 실행 중이 아닐 때
        stop_flag = False  # 중단 플래그 초기화
        scraper_thread = Thread(target=on_start)
        scraper_thread.start()
    else:  # 스크립트가 실행 중일 때
        stop_flag = True  # 중단 요청
        logging.info("중단 요청이 감지되었습니다. 현재 작업을 중단합니다.")
        update_log_safe("중단 요청이 감지되었습니다. 현재 작업을 중단합니다.")
        scraper_thread.join()  # 스레드가 종료될 때까지 대기
        scraper_thread = None  # 스레드가 종료된 후에 초기화


# GUI 시작
def start_gui():
    global root, combobox_title, entry_start_chapter, entry_num_chapters_to_process, entry_total_chapters, scraper_thread, stop_flag, log_text
    global gui_queue  # 큐를 전역으로 선언하여 다른 함수에서도 접근할 수 있게 함

    stop_flag = False
    scraper_thread = None

    root = tk.Tk()
    root.title("Novel Scraper")

    tk.Label(root, text="소설 제목").grid(row=0, column=0, padx=5, pady=5, sticky="e")
    combobox_title = ttk.Combobox(root)
    combobox_title.grid(row=0, column=1, padx=5, pady=5, sticky="w")

    tk.Label(root, text="시작 화수 (선택사항)").grid(row=1, column=0, padx=5, pady=5, sticky="e")
    entry_start_chapter = tk.Entry(root)
    entry_start_chapter.grid(row=1, column=1, padx=5, pady=5, sticky="w")

    tk.Label(root, text="진행할 화수 (선택사항)").grid(row=2, column=0, padx=5, pady=5, sticky="e")
    entry_num_chapters_to_process = tk.Entry(root)
    entry_num_chapters_to_process.grid(row=2, column=1, padx=5, pady=5, sticky="w")

    tk.Button(root, text="One (F2)", command=run_single_chapter).grid(row=2, column=2, padx=1, pady=5, sticky="w")

    tk.Label(root, text="소설 총 화수 (최초 입력)").grid(row=3, column=0, padx=5, pady=5, sticky="e")
    entry_total_chapters = tk.Entry(root)
    entry_total_chapters.grid(row=3, column=1, padx=5, pady=5, sticky="w")

    tk.Button(root, text='설정', command=open_settings).grid(row=4, column=0, padx=1, pady=5)
    tk.Button(root, text='시작/중단 (F4)', command=lambda: on_toggle_key_press(None)).grid(row=4, column=1, padx=1, pady=5)
    tk.Button(root, text='저장 검사', command=check_files_and_notify).grid(row=4, column=2, padx=1, pady=5, sticky="w")

    tk.Button(root, text='작업 목록', command=open_task_list_gui).grid(row=4, column=3, padx=1, pady=5, sticky="w")

    tk.Button(root, text='새로고침', command=update_dropdown).grid(row=0, column=2, padx=1, pady=5, sticky="w")

    log_text = tk.Text(root, state=tk.DISABLED, height=15, width=50)
    log_text.grid(row=5, column=0, columnspan=4, padx=10, pady=10)

    # GUI 큐 초기화 및 처리 시작
    gui_queue = queue.Queue()  # 큐 초기화
    root.after(100, process_queue)  # 100ms 후에 큐를 처리하는 함수 실행

    root.after(100, update_dropdown)
    root.after(100, register_key_events)

    root.mainloop()

if __name__ == "__main__":
    start_gui()
