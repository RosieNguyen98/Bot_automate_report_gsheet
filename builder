import requests
import io
import os
import json
import pandas as pd

import time
from datetime import date, datetime, timedelta

import pywinauto
from pywinauto import Application
import pyautogui

import tempfile
import pyperclip
from PIL import Image, ImageGrab

import gspread
from google.oauth2 import service_account
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

import sys
import concurrent.futures
from tabulate import tabulate

from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive

webhook_err = "https://chat.googleapis.com/v1/spaces/AAAAoNBFaxk/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=KhEud8UmgmJdlrACsmiHJ2cidZzTLkJrqRFKtoL4B4U"
sheet_id = "112tkWx3HP_O-_jxs_HcdCnKEEGyccSrJR-blM2-ONgQ"

service_path = r"H:\My Drive\Tung NJV\Python_source"
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
creds = service_account.Credentials.from_service_account_file(rf"{service_path}\task_service.json", scopes=scopes)

client = gspread.authorize(creds)
drive_service = build('drive', 'v3', credentials=creds)
sheet_service = build('sheets', 'v4', credentials=creds)
# pbi_title = "Dataset WH Tasks - Power BI Desktop"
pbi_title = "Dataset WH Tasks"
# save_path = r"H:\My Drive\HaNoiWarehousePowerBI\HaNoiWarehouseDataSource\SourceHaNoiTasks"
save_path = r"H:\My Drive\HaNoiWarehousePowerBI\Source Data\SourceHaNoiTasks"


# client = gspread.service_account_from_dict(service_key, scope)
with open(os.path.join(service_path, "task_service.json"), "r") as key_file:
    service_key = json.load(key_file)
_gsetting = {
    'client_config_backend': 'service',
    'service_config': {
        'client_json_dict': service_key
    }
}
gauth = GoogleAuth(settings=_gsetting)
gauth.ServiceAuth()


def restart_computer():
    requests.post(webhook_err, json={'text': f'Computer is going to restart in 10s'})
    time.sleep(10)
    if os.name == "posix":  # Unix/Linux/MacOS
        os.system("sudo reboot")
    elif os.name == "nt":  # Windows
        os.system("shutdown /r /t 1")
    else:
        print("Unsupported operating system.")


browser_list = [r'Microsoft\u200b Edge', r'Google Chrome']
def open_browser():
    global browser_name
    # Browser always opens
    for name in browser_list:
        try:
            try:
                browser = Application(backend="uia").connect(title_re=f".*{name}.*", timeout=5, found_index=0).window(title_re=f".*{name}.*", found_index=0)
                browser.restore().maximize()
                browser_name = name
                break
            except:
                # os.startfile(r"C:\Program Files\Google\Chrome\Application\chrome.exe")
                os.startfile(r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe")
                time.sleep(10)
                # w.open('https://mail.google.com/chat/u/0/#chat/space/AAAAN9_PCV0')
                browser = Application(backend="uia").connect(title_re=f".*{name}.*", timeout=5, found_index=0).window(title_re=f".*{name}.*", found_index=0)
                browser.restore().maximize()
                browser_name = name
                break
        except pywinauto.findwindows.ElementNotFoundError:
            lst = get_running_windows()
            for _ in lst:
                if name in _:
                    browser = Application(backend="uia").connect(title=f"{_}", timeout=5, found_index=0).window(title=f"{_}", found_index=0)
                    browser.restore().maximize()
                    browser_name = name
                    break
    return browser

def ref_pbi_web():
    browser = open_browser()
    
    browser.maximize()
    time.sleep(1)
    pyautogui.hotkey('ctrl', '1')  # Tránh 1 lỗi
    time.sleep(1)

    list_links = [{"name": "Daily Task", "url": "https://app.powerbi.com/datahub/datasets/0e5822e6-2458-4da8-ae6b-0e0820621c7d?experience=power-bi"}, #dailytask
                  {"name": "Dashboard", "url": "https://app.powerbi.com/datahub/datasets/9f13ee34-e5a5-4cb2-9592-a72c9937846b?experience=power-bi"}  #dashboard
                ]
    for link in list_links:
        try:
            # browser.child_window(title="New Tab", control_type="Button").click_input()  # Open new tab
            browser.NewTabButton.click_input()
            time.sleep(0.5)
            pyperclip.copy(link['url'])
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.5)
            pyautogui.hotkey('enter')
            time.sleep(20)  # Wait tab
            

            browser.child_window(title="Refresh", control_type="Button", found_index=1).click_input()
            time.sleep(3)
            browser.child_window(title="Refresh now").click_input()
            time.sleep(3)
        except Exception as e:
            print("Refresh {} failed: {}".format(link['name'], str(e)))
            continue

    browser.close()


def get_pbi_window(pbi_title):
    # app = Application(backend="uia").connect(title=pbi_title)
    # pbi = app.window(title=pbi_title)
    # pbi.maximize()
    # pbi.set_focus()
    # return pbi
    app = Application(backend="uia").connect(title_re=".*" + pbi_title + ".*")
    pbi = app.window(title_re=".*" + pbi_title + ".*")
    pbi.maximize()
    pbi.set_focus()
    return pbi

def refresh_pbi(pbi):
    home = pbi.child_window(title = "Home", control_type="TabItem", found_index=0).click_input() # found 2 take first
    time.sleep(1)

    refresh = pbi.child_window(title = "Refresh", control_type="Button", found_index=0).click_input() # found 2 take first first
    time.sleep(1)

    start_time = datetime.now()
    while True:
        try:
            time.sleep(5)
            pbi.child_window(title="Close", control_type="Button", found_index=1).click_input()  # Refresh succeed and need click close => found 2 take second
            time.sleep(3)
            return True
        except:
            cur_time = datetime.now()
            duration = (cur_time - start_time).total_seconds()
            if duration <= 600:
                continue
            elif duration > 600:
                cancel = pbi.child_window(title="Cancel", control_type="Button", found_index=1)
                if not cancel:
                    return True
                else:
                    cancel.click_input()
                    return False
            else:
                close = pbi.child_window(title="Close", control_type="Button")
                if close.exists() and close.is_visible():
                    close.click_input()
                    return False
                else:
                    return True

def export_pbi(pbi, table_title, file_path, wait_time):
    pbi.child_window(title=table_title, control_type='Group', found_index=0).click_input()
    time.sleep(1)

    pbi.child_window(title="More options", control_type="Button", found_index=0).click_input()
    time.sleep(1)
    
    pbi.child_window(title="Export data", control_type="MenuItem", found_index=0).click_input()
    if wait_time:
        time.sleep(int(wait_time))
    else:
        time.sleep(3)

    pyautogui.write(file_path)
    time.sleep(1)
    pyautogui.hotkey('Enter')
    time.sleep(2)

    try:
        pbi.child_window(title="Confirm Save As", control_type="Window").child_window(title="Yes", auto_id="CommandButton_6", control_type="Button").click_input()
    except:
        pass
    time.sleep(1)

def combinetosend(folder_path, prefix):
    def file_content(folder_path, prefix, file):
        try:
            if (file.endswith(".csv")) and (prefix in file):
                df_file = pd.read_csv(os.path.join(folder_path, file))
                print(f"{file} - {len(df_file)}")
            else:
                df_file = pd.DataFrame()
        except:
            df_file = pd.DataFrame()
        return df_file

    try:
        with concurrent.futures.ThreadPoolExecutor() as executor:
            df_file_list = list(executor.map(lambda file: file_content(folder_path, prefix, file), os.listdir(folder_path)))
        df_combined = pd.concat(df_file_list, ignore_index=True)
        combined_status = True
    except:
        df_combined = pd.DataFrame()
        combined_status = False

    return df_combined, combined_status

def sheet_content(sheet_id, range_name, header_row_num):
    try:
        # if datetime.now().hour < 5:
        #     range_name = datetime.now().strftime('%Y-%m-%d')
        # else:
        #     range_name = (datetime.today() + timedelta(days=1)).strftime('%Y-%m-%d')
        # print(range_name)

        data = sheet_service.spreadsheets().values().get(spreadsheetId=sheet_id, range=range_name).execute()
        values = data.get('values', [])

        datarows_num = header_row_num + 1
        df = pd.DataFrame(values[datarows_num:], columns=values[header_row_num])
        return df
    except Exception as e:
        requests.post(webhook_err, json={'text': f'Can not get content of {range_name}'})
        return pd.DataFrame()


def capture_area(left, top, right, bottom):
    image = ImageGrab.grab(bbox=(left, top, right, bottom))
    return image

def send_image(image, row):
    temp_path = tempfile.gettempdir()
    image_path = os.path.join(temp_path, row['export_name'])
    image.save(image_path, format='PNG')

    drive = GoogleDrive(gauth)

    metadata = {
        'parents': [{'id': row['folder']}],
        'title': row['export_name'],
        'mimeType': 'image/png'
    }
    
    permission = {
        'type': 'anyone',
        'role': 'reader',
    }

    image_file = drive.CreateFile(metadata=metadata)
    image_file.SetContentFile(image_path)
    image_file.Upload()
    try:
        image_file.InsertPermission(permission)
    except:
        pass

    image_link = image_file['alternateLink']
    # image_link = image_file['embedLink']
    # payload = {'text': '<users/all> {}\n {}'.format(row['mess_text'], image_link)}
    payload = {'text': '{}\n {}'.format(row['mess_text'], image_link)}
    response = requests.post(row['webhook'], json=payload)

def send_report(pbi_title, space_sheet, list_task):
    pbi = get_pbi_window(pbi_title)
    df_send = sheet_content(sheet_id=sheet_id, range_name=space_sheet, header_row_num=0)
    # df_send = df_send[df_send['task_name'].isin(list_task) & (df_send['send_flag'] == 'y')].to_dict('records')
    df_send = df_send[df_send['task_name'].isin(list_task) & (df_send['proceed_flag'] == 'y')].to_dict('records')

    for row in df_send:
        print('Sending {}'.format(row['task_name']), row['send_type'])

        if row['send_type'] == "image" and row['send_flag'] == 'y':
            pbi_page = pbi.child_window(title = row['page_name'], control_type="Group", found_index=0).click_input()
            if row['page_sleep']:
                time.sleep(int(row['page_sleep']))
            else:
                time.sleep(5)
            image = capture_area(int(row['left']), int(row['top']), int(row['right']), int(row['bottom']))
            send_image(image=image, row=row)

        elif row['send_type'] == "file":
            pbi_page = pbi.child_window(title = row['page_name'], control_type="Group", found_index=0).click_input()
            if row['page_sleep']:
                time.sleep(int(row['page_sleep']))
            else:
                time.sleep(3)
            export_pbi(pbi=pbi, table_title=row['table_title'], file_path=os.path.join(row['folder'], row['export_name']), wait_time=row['wait_time'])
            time.sleep(1)
            if row['send_flag'] == 'y':
                # payload = {'text': '<users/all> {}\n {}'.format(row['mess_text'], row['mess_link'])}
                payload = {'text': '{}\n {}'.format(row['mess_text'], row['mess_link'])}
                response = requests.post(row['webhook'], json=payload)
        elif row['send_type'] == "combinetosend":
            df_combined, combined_status = combinetosend(folder_path=row['table_title'], prefix=row['prefix'])
            if combined_status:
                df_combined.to_csv(os.path.join(row['folder'], row['export_name']), index=False)
                # payload = {'text': '<users/all> {}\n {}'.format(row['mess_text'], row['mess_link'])}
                payload = {'text': '{}\n {}'.format(row['mess_text'], row['mess_link'])}
                response = requests.post(row['webhook'], json=payload)


def get_job_id(domain, headers, query_id, query_name, params):
    try:
        # print(params)
        response = requests.post(f'{domain}/api/queries/{query_id}/refresh', headers=headers, params=params)
        # print(response.text)
        if response.status_code != 200:
            error = response.json()['message']
            requests.post(webhook_err, json={'text': f'Query {query_id} - {query_name} failed!: {error}'})
        else:
            job_id = response.json()['job']['id']
    except Exception as e:
        requests.post(webhook_err, json={'text': f'Query {query_id} - {query_name} failed!: {(str(e))}'})
    return job_id

def get_job_status(domain, headers, query_id, query_name, job_id):
    while True:
        try:
            response = requests.get(f'{domain}/api/jobs/{job_id}', headers=headers)
            job_status = response.json()['job']['status']
            if job_status == 3:
                print(f'Succeed Query {query_id} - {query_name}')
                # return True
                return response.json()['job']['query_result_id']

            elif job_status == 4:
                error = response.json()['job']['error']
                requests.post(webhook_err, json={'text': f'Query {query_id} - {query_name} failed!: {error}'})
                return False

        except requests.exceptions.RequestException as e:
            print(f'Error Query {query_id} - {query_name}: {str(e)}')
            continue

        except Exception as e:
            requests.post(webhook_err, json={'text': f'Query {query_id} - {query_name} failed!: {(str(e))}'})
            return False

def run_redash(row):
    headers = {"Authorization": "Key {}".format(row['api_key'])}
    if row['params']:
        params = json.loads(row['params'].replace("'", "\""))
    else:
        params = {}

    for attempt in range(1, 6):
        try:
            job_id = get_job_id(domain=row['domain'], headers=headers, query_id=row['query_id'], query_name=row['query_name'], params=params)
            if job_id:
                query_result_id = get_job_status(domain=row['domain'], headers=headers, query_id=row['query_id'], query_name=row['query_name'], job_id=job_id)
                if query_result_id:
                    excuted_result = 'Succeed'

                    response = requests.get('{}/api/query_results/{}.json'.format(row['domain'], query_result_id), headers=headers)
                    data = response.json()
                    df_result = pd.DataFrame(data['query_result']['data']['rows'], columns=[col['name'] for col in data['query_result']['data']['columns']])
                    # print(df_result)
                    df_result.to_csv(os.path.join(save_path, "{}.csv".format(row['query_name'])), index=False, encoding='utf-8')

                    redash_rows_cnt = len(data['query_result']['data']['rows'])
                    runtime = data['query_result']['runtime']

                    break
        except Exception as e:
            requests.post(webhook_err, json={'text': 'Attempt {} failed: Query {} - {} !: {}'.format(attempt, row['query_id'], row['query_name'], str(e))})
            time.sleep(10)
    else:
        excuted_result = 'Failed'
        redash_rows_cnt = 0
        runtime = 0
        query_result_id = 0
        print('Query {} - {} failed after 5 attempt!'.format(row['query_id'], row['query_name']))
    return {'query_id': row['query_id'], 'query_name': row['query_short_name'], 'excuted':  excuted_result, 'rows_cnt': redash_rows_cnt, 'runtime': "{:.2f}".format(runtime), 'query_result_id': query_result_id}


def run_redash_tasks(content_sheet_name, list_task):
    df_task = sheet_content(sheet_id=sheet_id, range_name=content_sheet_name, header_row_num=0)
    df_task = df_task[df_task['task_name'].isin(list_task) & (df_task['active_flag'] == 'y') & (df_task['run_flag'] == 'y')]
    df_task = df_task.drop_duplicates(subset=['query_id', 'params'], keep='first')

    tasks = ",".join(task for task in list_task)
    print(f'Running {tasks} queries:\n{df_task}')

    with concurrent.futures.ThreadPoolExecutor() as executor:
        results = list(executor.map(lambda row: run_redash(row=row), df_task.to_dict('records')))

    requests.post(webhook_err, json={'text': "<users/all> Run Redash Results:\n```{}```".format(tabulate(results, headers='keys'))})

    len_results = len(results)
    len_fail = len([row for row in results if row['excuted'] == 'Failed'])

    if len_results == len_fail:
        requests.post(webhook_err, json={'text': "<users/all> All queries failed, not send report!"})
        return False
    elif len_results > len_fail:
        return True



def run_task(run_type, list_task):
    try:
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        print('\nRun schedule at {}'.format(now))
        
        if run_type == "web_service":
            ref_pbi_web()
        else:

            if run_type == "once":
                redash_ref_result = run_redash_tasks(content_sheet_name="WH - Task", list_task=list_task)

                if redash_ref_result == True:
                    pbi = get_pbi_window(pbi_title=pbi_title)
                    pbi_ref_result = refresh_pbi(pbi=pbi)
                else:
                    pbi_ref_result = False

            elif run_type == "quick":
                pbi_ref_result = True

            if pbi_ref_result == True:
                print("Refresh PBI succeed, prepare to send")
                send_report(pbi_title=pbi_title, space_sheet="WH - Mess", list_task=list_task)
            elif pbi_ref_result == False:
                print("Refresh PBI failed, not send report!")
                requests.post(webhook_err, json={'text': f'<users/all> {pbi_title} refresh failed, not send report!'})
                # restart_computer()
    except Exception as e:
        print(str(e))
        try:
            requests.post(webhook_err, json={'text': '<users/all> Send failed: {}\n{}!'.format(",".join(task for task in list_task), str(e))})
        except:
            pass
    print('End task')

# run_task(run_type="once", list_task=["invent_b2b_hn"])
