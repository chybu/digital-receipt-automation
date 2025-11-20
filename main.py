from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException, StaleElementReferenceException
from time import sleep
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from datetime import datetime,timedelta
import pyautogui, os, pickle, httpx, math, asyncio, json, sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import Calendar
from calendar import monthrange
import pandas as pd
import pygetwindow as gw
import subprocess
from getAuth import getAuth
from traceback import print_exc

if hasattr(sys, '_MEIPASS'):
    # Running in the temp directory (PyInstaller extracted files)
    script_dir = os.path.dirname(sys.executable)
else:
    # Running in the normal directory (not packaged)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
path_file = os.path.join(script_dir, 'path.json')
with open(path_file, 'r') as file:
    path_dic = json.load(file)
exe_path = path_dic["exe_path"]
driver_path = path_dic["driver_path"]
batch_script = path_dic["batch_script"]

tinh_chat_dic = {1:"Hàng hóa, dịch vụ",2:"Hàng khuyến mại",3:"Chiết khấu thương mại",4:"Ghi chú, diễn giải", 5:"Hàng hoá đặc trưng"}

async def returnNone():
    return None

async def get_receipt_by_API(auth:str, ma_so_thue:str, ky_hieu_hoa_don:str, so_hoa_don:str, ky_hieu_mau_so:str, index:int):
    if index<2: url = "https://hoadondientu.gdt.gov.vn:30000/query/invoices/detail"
    else: url = "https://hoadondientu.gdt.gov.vn:30000/sco-query/invoices/detail"

    querystring = {"nbmst":ma_so_thue,"khhdon":ky_hieu_hoa_don,"shdon":so_hoa_don,"khmshdon":ky_hieu_mau_so}
    
    headers = {
        "Accept": "application/json",
        "Accept-Language": "vi",
        "Authorization": auth,
        "Connection": "keep-alive",
        "End-Point": "/tra-cuu/tra-cuu-hoa-don",
        "Origin": "https://hoadondientu.gdt.gov.vn",
        "Referer": "https://hoadondientu.gdt.gov.vn/",
        "Sec-Fetch-Dest": "empty",
        "Sec-Fetch-Mode": "cors",
        "Sec-Fetch-Site": "same-site",
        "sec-ch-ua-mobile": "?0",
    }
    too_many_request_counter = 0
    while True:
        try:
            async with httpx.AsyncClient(verify=False) as client:
                response = await client.get(url, headers=headers, params=querystring, timeout=6)
                if response.status_code==429:
                    if too_many_request_counter>2: return ("need recover",)
                    else: 
                        too_many_request_counter+=1
                        await asyncio.sleep(3)
                else: break
        except (httpx.TimeoutException):""
    # print(response.status_code, index)
    if response.status_code!=200: 
        return ("bad auth",)
    try:
        res = response.json()
        return ("good",res)
    except Exception:
        return ("need recover",)

def format_empty_str(text:str):
    if not text.strip(): return "n/a"
    else: return text.strip()

def open_web(driver:webdriver, url:str):
    driver.get(url)
    # Wait for the page to finish loading
    sleep(5)
    while driver.execute_script("return document.readyState") != "complete":
        sleep(1)
    
def enter_ban_ra(driver:webdriver, start:datetime, end:datetime, auth:str, using_username:str, pw:str, safemode:bool):
    invalid_hoa_don_list = []
    invalid_dich_vu_list = []
    
    url = "https://hoadondientu.gdt.gov.vn/tra-cuu/tra-cuu-hoa-don"
    if driver.current_url!=url: open_web(driver, url)

    ban_ra_tab = driver.find_element(by=By.CSS_SELECTOR, value="#__next > section > section > main > div > div > div > div > div.ant-tabs-bar.ant-tabs-top-bar > div > div > div > div > div:nth-child(1) > div:nth-child(1)")
    ban_ra_tab.click()

    start_input = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#tngay > div > input")))
    start_calendar_icon = driver.find_element(By.CSS_SELECTOR, "#tngay > div > i.anticon.anticon-calendar.ant-calendar-picker-icon")
    end_input = driver.find_element(by=By.CSS_SELECTOR, value="#dngay > div > input")
    end_calendar_icon = driver.find_element(By.CSS_SELECTOR, "#dngay > div > i.anticon.anticon-calendar.ant-calendar-picker-icon")
    search_button = driver.find_element(by=By.CSS_SELECTOR, value="#__next > section > section > main > div > div > div > div > div.ant-tabs-content.ant-tabs-content-animated.ant-tabs-top-content > div.ant-tabs-tabpane.ant-tabs-tabpane-active > div.ant-row > div:nth-child(1) > div > div > form > div.ant-row-flex.ant-row-flex-center > div:nth-child(1) > button")
                 
    def enter_date():
        # clear input
        ActionChains(driver).move_to_element(end_calendar_icon).perform()
        end_close_icon = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#dngay > div > i.anticon.anticon-close-circle.ant-calendar-picker-clear")))
        end_close_icon.click()

        ActionChains(driver).move_to_element(start_calendar_icon).perform()
        start_close_icon = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#tngay > div > i.anticon.anticon-close-circle.ant-calendar-picker-clear")))
        start_close_icon.click()
        # change end date
        ActionChains(driver).move_to_element(end_input).click().perform()
        ActionChains(driver).move_to_element(end_input).click().send_keys(end.strftime("%d/%m/%Y")).perform()
        end_selected_day = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "td.ant-calendar-cell.ant-calendar-selected-day")))
        # end_selected_day = driver.find_element(By.CSS_SELECTOR, 'td.ant-calendar-cell.ant-calendar-selected-day')
        end_selected_day.click()
        # change start day
        ActionChains(driver).move_to_element(start_input).click().perform()
        ActionChains(driver).move_to_element(start_input).click().send_keys(start.strftime("%d/%m/%Y")).perform()
        start_selected_day = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "td.ant-calendar-cell.ant-calendar-selected-day")))
        # start_selected_day = driver.find_element(By.CSS_SELECTOR, 'td.ant-calendar-cell.ant-calendar-selected-day')
        start_selected_day.click()
        # searching
        search_button.click()
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#__next > div.Spin__SpinWrapper-sc-1px1x1e-0.enNzTo > div > div")))
        WebDriverWait(driver, 60).until(EC.invisibility_of_element((By.CSS_SELECTOR, "#__next > div.Spin__SpinWrapper-sc-1px1x1e-0.enNzTo > div > div")))
        res = driver.find_element(by=By.CSS_SELECTOR, value="#__next > section > section > main > div > div > div > div > div.ant-tabs-content.ant-tabs-content-animated.ant-tabs-top-content > div.ant-tabs-tabpane.ant-tabs-tabpane-active > div.ant-row > div:nth-child(2) > div.ant-row-flex.ant-row-flex-space-between.ant-row-flex-middle > div:nth-child(1) > div > span")
        # res.text example: Có 3 kết quả
        return int(res.text.split(" ")[1])
    
    def recover(target):
        print("recover")
        sleep(10)
        while True:
            try:
                open_web(driver, url)
                ban_ra_tab = driver.find_element(by=By.CSS_SELECTOR, value="#__next > section > section > main > div > div > div > div > div.ant-tabs-bar.ant-tabs-top-bar > div > div > div > div > div:nth-child(1) > div:nth-child(1)")
                ban_ra_tab.click()
                break
            except Exception:
                print("deep sleep")
                sleep(10*60)
                log_in(driver, using_username, pw)
        # refind all needed elements
        nonlocal start_input
        nonlocal start_calendar_icon
        nonlocal end_input
        nonlocal end_calendar_icon
        nonlocal search_button
        start_input = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#tngay > div > input")))
        start_calendar_icon = driver.find_element(By.CSS_SELECTOR, "#tngay > div > i.anticon.anticon-calendar.ant-calendar-picker-icon")
        end_input = driver.find_element(by=By.CSS_SELECTOR, value="#dngay > div > input")
        end_calendar_icon = driver.find_element(By.CSS_SELECTOR, "#dngay > div > i.anticon.anticon-calendar.ant-calendar-picker-icon")
        search_button = driver.find_element(by=By.CSS_SELECTOR, value="#__next > section > section > main > div > div > div > div > div.ant-tabs-content.ant-tabs-content-animated.ant-tabs-top-content > div.ant-tabs-tabpane.ant-tabs-tabpane-active > div.ant-row > div:nth-child(1) > div > div > form > div.ant-row-flex.ant-row-flex-center > div:nth-child(1) > button")
    
        enter_date()
        # 1 / 38
        page_text = driver.find_element(by=By.XPATH, value="/html/body/div[1]/section/section/main/div/div/div/div/div[3]/div[1]/div[3]/div[2]/div[1]/div[2]/div/div[2]/div").text
        page_number = int(page_text.split('/')[0].strip())
        while page_number!=target:
            next_page_button = driver.find_element(by=By.CSS_SELECTOR, value="#__next > section > section > main > div > div > div > div > div.ant-tabs-content.ant-tabs-content-animated.ant-tabs-top-content > div.ant-tabs-tabpane.ant-tabs-tabpane-active > div.ant-row > div:nth-child(2) > div.ant-row-flex.ant-row-flex-space-between.ant-row-flex-middle > div:nth-child(2) > div > div:nth-child(3) > button")
            next_page_button.click()
            # wait for the page to change
            WebDriverWait(driver, 5*60).until(lambda d: d.find_element(by=By.XPATH, value="/html/body/div[1]/section/section/main/div/div/div/div/div[3]/div[1]/div[3]/div[2]/div[1]/div[2]/div/div[2]/div").text!=page_text)
            page_text = driver.find_element(by=By.XPATH, value="/html/body/div[1]/section/section/main/div/div/div/div/div[3]/div[1]/div[3]/div[2]/div[1]/div[2]/div/div[2]/div").text
            page_number = int(page_text.split('/')[0].strip())
    
    final_khms = []
    final_ky_hieu = []
    final_so_hoa_don = []
    final_kyhieu_sohoadon = []
    final_ngaylap = []
    final_mccqt = []
    final_MST_nguoimua = []
    final_ten_nguoimua = []
    final_dia_chi_nguoimua = []
    final_MST_nguoiban = []
    final_ten_nguoiban = []
    final_dia_chi_nguoiban = []
    final_tien_chua_thue = []
    final_tien_thue = []
    final_chiet_khau_thuong_mai = []
    final_tien_phi = []
    final_tien_thanh_toan = []
    final_don_vi_tien = []
    final_hinh_thuc_thanh_toan = []
    final_thue_suat = []
    final_tong_tien_bang_so = []
    final_tong_tien_bang_chu = []
    final_trang_thai = []
    final_ket_qua = []

    final_khshd_list = []
    final_tinh_chat_list = []
    final_ten_list = []
    final_don_vi_list = []
    final_so_luong_list = []
    final_don_gia_list = []
    final_chiet_khau_list = []
    final_thue_suat_list = []
    final_thanh_tien_list = []
    final_tien_thue_list = []
    final_tong_tien_list = []
    final_ten_target_list = []
    final_mst_target_list = []
    final_trang_thai_hoa_don_list = []
    
    number = enter_date()
    page_now = 1
    if number==0: return None, None, auth, invalid_hoa_don_list, invalid_dich_vu_list
    counter = 0
    finish = True
    while counter!=number:
        temp_khms = []
        temp_ky_hieu = []
        temp_so_hoa_don = []
        temp_kyhieu_sohoadon = []
        temp_ngaylap = []
        temp_mccqt = []
        temp_MST_nguoimua = []
        temp_ten_nguoimua = []
        temp_dia_chi_nguoimua = []
        temp_MST_nguoiban = []
        temp_ten_nguoiban = []
        temp_dia_chi_nguoiban = []
        temp_tien_chua_thue = []
        temp_tien_thue = []
        temp_chiet_khau_thuong_mai = []
        temp_tien_phi = []
        temp_tien_thanh_toan = []
        temp_don_vi_tien = []
        temp_hinh_thuc_thanh_toan = []
        temp_thue_suat = []
        temp_tong_tien_bang_so = []
        temp_tong_tien_bang_chu = []
        temp_trang_thai = []
        temp_ket_qua = []

        temp_khshd_list = []
        temp_tinh_chat_list = []
        temp_ten_list = []
        temp_don_vi_list = []
        temp_so_luong_list = []
        temp_don_gia_list = []
        temp_chiet_khau_list = []
        temp_thue_suat_list = []
        temp_thanh_tien_list = []
        temp_tien_thue_list = []
        temp_tong_tien_list = []
        temp_ten_target_list = []
        temp_mst_target_list = []
        temp_trang_thai_hoa_don_list = []
                
        stt_list = driver.find_elements(By.CSS_SELECTOR, "tr.ant-table-row.ant-table-row-level-0 > td:nth-child(1) > span")
        khms_list = driver.find_elements(By.CSS_SELECTOR, "tr.ant-table-row.ant-table-row-level-0 > td:nth-child(3)")
        ky_hieu_list = driver.find_elements(By.CSS_SELECTOR, "tr.ant-table-row.ant-table-row-level-0 > td:nth-child(4) > span")
        so_hoa_don_list = driver.find_elements(By.CSS_SELECTOR, "tr.ant-table-row.ant-table-row-level-0 > td:nth-child(5)")
        trang_thai_list = driver.find_elements(By.CSS_SELECTOR, "tr.ant-table-row.ant-table-row-level-0 > td:nth-child(14) > span")
        ket_qua_list = driver.find_elements(By.CSS_SELECTOR, "tr.ant-table-row.ant-table-row-level-0 > td:nth-child(15) > span")
        element_list = driver.find_elements(by=By.CSS_SELECTOR, value="tr.ant-table-row.ant-table-row-level-0")
            
        def network_test():
            try:
                if page_now<math.ceil(number/15): target = 15
                else: target = number-counter
                network_counter = 0
                for element, stt in zip(element_list, stt_list):
                    if not stt.text.strip(): continue
                    network_counter+=1
                    element.click()
                    sleep(0.1)
                return network_counter==target
            except Exception: return False
        
        if not network_test():
            recover(page_now)
            continue
        
        if finish:
            success_list = [False]*len(stt_list)
            res_list = [None]*len(stt_list)
            finish = False
            
        async def getData():
            tasks = []  
            for stt, khms, kh, so_hoa_don, success\
            in zip(stt_list, khms_list, ky_hieu_list, so_hoa_don_list, success_list):
                if not stt.text.strip(): tasks.append(asyncio.create_task(returnNone()))
                elif success: tasks.append(asyncio.create_task(returnNone()))
                else: tasks.append(asyncio.create_task(get_receipt_by_API(auth=auth, ma_so_thue=using_username, ky_hieu_hoa_don=kh.text, so_hoa_don=so_hoa_don.text, ky_hieu_mau_so=khms.text, index=0)))
            data_list = await asyncio.gather(*tasks)
            return data_list

        data_list = asyncio.run(getData())
        need_recover = False
        need_new_auth = False
        
        for i in range(len(stt_list)):
            success = success_list[i]
            if not success:
                stt = stt_list[i]
                if not stt.text.strip():
                    success_list[i] = True
                else:
                    data = data_list[i]
                    api_res = data[0]
                    if api_res=="good":
                        success_list[i] = True
                        res_list[i] = data[1]
                    elif api_res=="bad auth":
                        need_new_auth = True
                    else:
                        need_recover = True
        if need_recover:
            recover(page_now)
        elif need_new_auth:
            try:
                auth = getAuth(banra=True, driver_path=driver_path)
            except NoSuchElementException: recover(page_now)
        else:
            finish = True
            for stt, khms, kh, so_hoa_don, trang_thai, ket_qua, API_res\
            in zip(stt_list, khms_list, ky_hieu_list, so_hoa_don_list,trang_thai_list, ket_qua_list, res_list):
                if not stt.text.strip(): continue
                counter+=1
                
                invalid_dich_vu_ct = 0
                
                def format_none(target):
                    if target is None: return "n/a"
                    else:
                        target = str(target)
                        return target.strip()
                    
                def format_money(target):
                    if target is None: return "n/a"
                    return str(int(target))
                
                try:
                    current_khms = format_empty_str(khms.text)
                    current_ky_hieu = format_empty_str(kh.text)
                    current_so_hoa_don = format_empty_str(so_hoa_don.text)
                    current_trang_thai_hoa_don = format_empty_str(trang_thai.text)
                    current_ket_qua = format_empty_str(ket_qua.text)
                
                    current_mst_nguoimua = format_none(API_res["nmmst"])
                    current_kyhieu_sohoadon = f"{current_ky_hieu}_{current_so_hoa_don}_{current_mst_nguoimua}"

                    current_ngaylap = format_none(API_res["tdlap"][:10])
                    current_mccqt = format_none(API_res["mhdon"])
                    current_ten_nguoiban = format_none(API_res["nbten"])
                    current_mst_nguoiban = format_none(API_res["nbmst"])
                    current_dia_chi_nguoiban = format_none(API_res["nbdchi"])
                    current_ten_nguoimua = format_none(API_res["nmten"])
                    current_dia_chi_nguoimua = format_none(API_res["nmdchi"])
                    
                    if "thtttoan" in API_res:
                        current_hinh_thuc_thanh_toan = format_none(API_res["thtttoan"])
                    else:
                        current_hinh_thuc_thanh_toan = "n/a"
                        
                    if "tgtphi" in API_res:
                        current_tien_phi = format_money(API_res["tgtphi"])
                        
                    if "dvtte" in API_res:
                        current_dvtte = format_none(API_res["dvtte"])
                    else:
                        current_dvtte = "n/a"
                        
                    if len(API_res["thttltsuat"])!=0 and "tsuat" in API_res["thttltsuat"][0]:
                        thue_suat_dic = API_res["thttltsuat"][0]
                        current_thue_suat = format_none(thue_suat_dic["tsuat"])
                    else: 
                        current_thue_suat = "n/a"

                    current_tong_tien_bang_so = format_money(API_res["tgtttbso"])
                    current_tong_tien_bang_chu = format_none(API_res["tgtttbchu"])
                    current_tien_chua_thue = format_money(API_res["tgtcthue"])
                    current_tien_thue = format_money(API_res["tgtthue"])
                    current_chiet_khau_thuong_mai = format_money(API_res["ttcktmai"])
                    current_tien_thanh_toan = format_money(API_res["tgtttbso"])

                    temp_khms.append(current_khms)
                    temp_ky_hieu.append(current_ky_hieu)
                    temp_so_hoa_don.append(current_so_hoa_don)
                    temp_trang_thai.append(current_trang_thai_hoa_don)
                    temp_ket_qua.append(current_ket_qua)
                    temp_kyhieu_sohoadon.append(current_kyhieu_sohoadon)
                    temp_ngaylap.append(current_ngaylap)
                    temp_mccqt.append(current_mccqt)
                    temp_ten_nguoiban.append(current_ten_nguoiban)
                    temp_MST_nguoiban.append(current_mst_nguoiban)
                    temp_dia_chi_nguoiban.append(current_dia_chi_nguoiban)
                    temp_ten_nguoimua.append(current_ten_nguoimua)
                    temp_MST_nguoimua.append(current_mst_nguoimua)
                    temp_dia_chi_nguoimua.append(current_dia_chi_nguoimua)
                    temp_hinh_thuc_thanh_toan.append(current_hinh_thuc_thanh_toan)
                    temp_tien_phi.append(current_tien_phi)
                    temp_don_vi_tien.append(current_dvtte)
                    temp_thue_suat.append(current_thue_suat)
                    temp_tong_tien_bang_so.append(current_tong_tien_bang_so)
                    temp_tong_tien_bang_chu.append(current_tong_tien_bang_chu)
                    temp_tien_chua_thue.append(current_tien_chua_thue)
                    temp_tien_thue.append(current_tien_thue)
                    temp_chiet_khau_thuong_mai.append(current_chiet_khau_thuong_mai)
                    temp_tien_thanh_toan.append(current_tien_thanh_toan)
                except Exception:
                    # neu hoa don bi loi thi se ko parse hang hoa dich vu trong hoa don
                    error_string = f"LOI FORMAT HOA DON BAN RA. SO HOA DON: {so_hoa_don.text}"
                    invalid_hoa_don_list.append(error_string)
                    
                    print("-------------------------------------------------")
                    print(error_string)
                    print_exc()
                    print("-------------------------------------------------")

                    continue
                
                
                hang_hoa_dich_vu_list = API_res["hdhhdvu"]
                
                def format_thue_suat(thue_suat):
                    thue = round(thue_suat*100,2)
                    string = f"{thue}%"
                    string = string.replace(".",",")
                    return string
                def format_tien_thue(thue):
                    string = str(thue)
                    string = string.replace(".", ",")
                    return string
                def get_thue_suat(target):
                    if target is None: return 0
                    return float(target)
                def format_tinh_chat(target):
                    if target is None: return "n/a"
                    return tinh_chat_dic[target]
                def get_thanh_tien(target):
                    if target is None: return 0
                    return int(target)
                
                for hhdv_dic in hang_hoa_dich_vu_list:
                    try:
                        current_tinh_chat = format_tinh_chat(hhdv_dic["tchat"])
                        current_ten = format_none(hhdv_dic["ten"])
                        current_don_vi = format_none(hhdv_dic["dvtinh"])
                        current_so_luong = format_none(hhdv_dic["sluong"]).replace(".",",")
                        current_don_gia = format_none(hhdv_dic["dgia"]).replace(".",",")
                        current_chiet_khau = format_none(hhdv_dic["stckhau"]).replace(".",",")
                        
                        current_thanh_tien = get_thanh_tien(hhdv_dic["thtien"])
                        if current_thanh_tien==0:
                            current_thue_suat = "n/a"
                            current_tien_thue = "n/a"
                            current_tong_tien = "n/a"
                        else:
                            current_thue_suat = get_thue_suat(hhdv_dic["tsuat"])
                            current_tien_thue = round(current_thanh_tien*current_thue_suat,2)
                            current_tong_tien = current_thanh_tien+current_tien_thue
                            
                            # format thue suat from float to x,x% format
                            current_thue_suat = format_thue_suat(current_thue_suat)
                            # format tien thue from float to x,x format
                            current_tien_thue = format_tien_thue(current_tien_thue)
                            # format tong tien from float to x,x format
                            current_tong_tien = format_tien_thue(current_tong_tien)
                    
                        temp_tinh_chat_list.append(current_tinh_chat)
                        temp_ten_list.append(current_ten)
                        temp_don_vi_list.append(current_don_vi)
                        temp_so_luong_list.append(current_so_luong)
                        temp_don_gia_list.append(current_don_gia)
                        temp_chiet_khau_list.append(current_chiet_khau)
                        temp_thue_suat_list.append(current_thue_suat)
                        temp_thanh_tien_list.append(current_thanh_tien)
                        temp_tien_thue_list.append(current_tien_thue)
                        temp_tong_tien_list.append(current_tong_tien)
                        temp_khshd_list.append(current_kyhieu_sohoadon)
                        temp_ten_target_list.append(current_ten_nguoimua)
                        temp_mst_target_list.append(current_mst_nguoimua)
                        temp_trang_thai_hoa_don_list.append(current_trang_thai_hoa_don)
                        
                    except Exception:
                        invalid_dich_vu_ct+=1
                        print("-------------------------------------------------")
                        print(f"LOI FORMAT DICH VU TRONG HOA DON BAN RA. SO HOA DON: {so_hoa_don.text}")
                        print_exc() 
                        print("-------------------------------------------------")

                if invalid_dich_vu_ct!=0:
                    error_string = f"{invalid_dich_vu_ct} LOI FORMAT DICH VU TRONG HOA DON BAN RA. SO HOA DON: {so_hoa_don.text}"
                    invalid_dich_vu_list.append(error_string)
            
        if finish:   
            final_khms.extend(temp_khms)
            final_ky_hieu.extend(temp_ky_hieu)
            final_so_hoa_don.extend(temp_so_hoa_don)
            final_kyhieu_sohoadon.extend(temp_kyhieu_sohoadon)
            final_ngaylap.extend(temp_ngaylap)
            final_mccqt.extend(temp_mccqt)
            final_MST_nguoimua.extend(temp_MST_nguoimua)
            final_ten_nguoimua.extend(temp_ten_nguoimua)
            final_dia_chi_nguoimua.extend(temp_dia_chi_nguoimua)
            final_MST_nguoiban.extend(temp_MST_nguoiban)
            final_ten_nguoiban.extend(temp_ten_nguoiban)
            final_dia_chi_nguoiban.extend(temp_dia_chi_nguoiban)
            final_tien_chua_thue.extend(temp_tien_chua_thue)
            final_tien_thue.extend(temp_tien_thue)
            final_chiet_khau_thuong_mai.extend(temp_chiet_khau_thuong_mai)
            final_tien_phi.extend(temp_tien_phi)
            final_tien_thanh_toan.extend(temp_tien_thanh_toan)
            final_don_vi_tien.extend(temp_don_vi_tien)
            final_hinh_thuc_thanh_toan.extend(temp_hinh_thuc_thanh_toan)
            final_thue_suat.extend(temp_thue_suat)
            final_tong_tien_bang_so.extend(temp_tong_tien_bang_so)
            final_tong_tien_bang_chu.extend(temp_tong_tien_bang_chu)
            final_trang_thai.extend(temp_trang_thai)
            final_ket_qua.extend(temp_ket_qua)

            final_khshd_list.extend(temp_khshd_list)
            final_tinh_chat_list.extend(temp_tinh_chat_list)
            final_ten_list.extend(temp_ten_list)
            final_don_vi_list.extend(temp_don_vi_list)
            final_so_luong_list.extend(temp_so_luong_list)
            final_don_gia_list.extend(temp_don_gia_list)
            final_chiet_khau_list.extend(temp_chiet_khau_list)
            final_thue_suat_list.extend(temp_thue_suat_list)
            final_thanh_tien_list.extend(temp_thanh_tien_list)
            final_tien_thue_list.extend(temp_tien_thue_list)
            final_tong_tien_list.extend(temp_tong_tien_list)
            final_ten_target_list.extend(temp_ten_target_list)
            final_mst_target_list.extend(temp_mst_target_list)
            final_trang_thai_hoa_don_list.extend(temp_trang_thai_hoa_don_list)
        
        # if there is more, turn to the next page
        if counter!=number and finish:
            initial_page = driver.find_element(by=By.CSS_SELECTOR, value="#__next > section > section > main > div > div > div > div > div.ant-tabs-content.ant-tabs-content-animated.ant-tabs-top-content > div.ant-tabs-tabpane.ant-tabs-tabpane-active > div.ant-row > div:nth-child(2) > div.ant-row-flex.ant-row-flex-space-between.ant-row-flex-middle > div:nth-child(2) > div > div:nth-child(2) > div").text
            next_page_button = driver.find_element(by=By.CSS_SELECTOR, value="#__next > section > section > main > div > div > div > div > div.ant-tabs-content.ant-tabs-content-animated.ant-tabs-top-content > div.ant-tabs-tabpane.ant-tabs-tabpane-active > div.ant-row > div:nth-child(2) > div.ant-row-flex.ant-row-flex-space-between.ant-row-flex-middle > div:nth-child(2) > div > div:nth-child(3) > button")
            next_page_button.click()
            # wait for the page to change
            WebDriverWait(driver, 5*60).until(lambda d: d.find_element(by=By.XPATH, value="/html/body/div[1]/section/section/main/div/div/div/div/div[3]/div[1]/div[3]/div[2]/div[1]/div[2]/div/div[2]/div").text!=initial_page)
            page_now+=1
            if safemode: sleep(30)
            else: sleep(15)
    if safemode: sleep(30)
    else: sleep(15)
    
    df = pd.DataFrame({
        'final_khms': final_khms,
        'final_ky_hieu': final_ky_hieu,
        'final_so_hoa_don': final_so_hoa_don,
        'final_kyhieu_sohoadon': final_kyhieu_sohoadon,
        'final_ngaylap': final_ngaylap,
        'final_mccqt': final_mccqt,
        'final_MST_nguoimua': final_MST_nguoimua,
        'final_ten_nguoimua': final_ten_nguoimua,
        'final_dia_chi_nguoimua': final_dia_chi_nguoimua,
        'final_MST_nguoiban': final_MST_nguoiban,
        'final_ten_nguoiban': final_ten_nguoiban,
        'final_dia_chi_nguoiban': final_dia_chi_nguoiban,
        'final_tien_chua_thue': final_tien_chua_thue,
        'final_tien_thue': final_tien_thue,
        'final_chiet_khau_thuong_mai': final_chiet_khau_thuong_mai,
        'final_tien_phi': final_tien_phi,
        'final_tien_thanh_toan': final_tien_thanh_toan,
        'final_don_vi_tien': final_don_vi_tien,
        'final_hinh_thuc_thanh_toan': final_hinh_thuc_thanh_toan,
        'final_thue_suat': final_thue_suat,
        'final_tong_tien_bang_so': final_tong_tien_bang_so,
        'final_tong_tien_bang_chu': final_tong_tien_bang_chu,
        'final_trang_thai': final_trang_thai,
        'final_ket_qua': final_ket_qua,
    })

    hanghoa_df = pd.DataFrame({
        "kyhieu_sohoadon": final_khshd_list,
        "tinh_chat": final_tinh_chat_list, 
        "ten": final_ten_list, 
        "don_vi": final_don_vi_list,
        "so_luong": final_so_luong_list, 
        "don_gia": final_don_gia_list, 
        "chiet_khau": final_chiet_khau_list, 
        "thue_suat": final_thue_suat_list, 
        "thanh_tien": final_thanh_tien_list,
        "tien_thue": final_tien_thue_list,
        "tong_tien": final_tong_tien_list,
        "ten_target": final_ten_target_list,
        "mst_target": final_mst_target_list,
        "trang_thai_hoa_don": final_trang_thai_hoa_don_list
    })
    return df, hanghoa_df, auth, invalid_hoa_don_list, invalid_dich_vu_list

def enter_mua_vao(driver:webdriver, start:datetime, end:datetime, auth:str, using_username:str, pw:str, safemode:bool):
    invalid_hoa_don_list = []
    invalid_dich_vu_list = []
    
    url = "https://hoadondientu.gdt.gov.vn/tra-cuu/tra-cuu-hoa-don"
    if driver.current_url!=url: open_web(driver, url)

    mua_vao_tab = driver.find_element(by=By.CSS_SELECTOR, value="#__next > section > section > main > div > div > div > div > div.ant-tabs-bar.ant-tabs-top-bar > div > div > div > div > div:nth-child(1) > div:nth-child(2)")
    mua_vao_tab.click()

    start_input = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH, "/html/body/div/section/section/main/div/div/div/div/div[3]/div[2]/div[3]/div[1]/div/div/form/div[1]/div[4]/div/div[2]/div/div[2]/div/div/div/span/span/div/input")))
    start_calendar_icon = driver.find_element(By.XPATH, "/html/body/div/section/section/main/div/div/div/div/div[3]/div[2]/div[3]/div[1]/div/div/form/div[1]/div[4]/div/div[2]/div/div[2]/div/div/div/span/span/div/i[2]")
    end_input = driver.find_element(by=By.XPATH, value="/html/body/div/section/section/main/div/div/div/div/div[3]/div[2]/div[3]/div[1]/div/div/form/div[1]/div[4]/div/div[3]/div/div[2]/div/div/div/span/span/div/input")
    end_calendar_icon = driver.find_element(By.XPATH, "/html/body/div/section/section/main/div/div/div/div/div[3]/div[2]/div[3]/div[1]/div/div/form/div[1]/div[4]/div/div[3]/div/div[2]/div/div/div/span/span/div/i[2]")
    search_button = driver.find_element(by=By.XPATH, value="/html/body/div/section/section/main/div/div/div/div/div[3]/div[2]/div[3]/div[1]/div/div/form/div[3]/div[1]/button")
    ket_qua_kiem_tra_dropdown = driver.find_element(by=By.XPATH, value="/html/body/div[1]/section/section/main/div/div/div/div/div[3]/div[2]/div[3]/div[1]/div/div/form/div[1]/div[5]/div/div[2]/div/div/div/span/div/div")
    ket_qua_kiem_tra_dropdown.click()
    sleep(0.25)
    ket_qua_kiem_tra_dropdown.click()
    ket_qua_kiem_tra_list = driver.find_elements(by=By.CSS_SELECTOR, value="ul.ant-select-dropdown-menu.ant-select-dropdown-menu-root.ant-select-dropdown-menu-vertical li")

    def enter_date():
        # clear input
        ActionChains(driver).move_to_element(end_calendar_icon).perform()
        try:
            end_close_icon = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH, "/html/body/div/section/section/main/div/div/div/div/div[3]/div[2]/div[3]/div[1]/div/div/form/div[1]/div[4]/div/div[3]/div/div[2]/div/div/div/span/span/div/i[1]")))
            # end_close_icon = driver.find_element(By.CSS_SELECTOR, "#dngay > div > i.anticon.anticon-close-circle.ant-calendar-picker-clear")
            end_close_icon.click()
        except NoSuchElementException: print("end day is already clear")
            
        ActionChains(driver).move_to_element(start_calendar_icon).perform()
        try:
            start_close_icon = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH, "/html/body/div/section/section/main/div/div/div/div/div[3]/div[2]/div[3]/div[1]/div/div/form/div[1]/div[4]/div/div[2]/div/div[2]/div/div/div/span/span/div/i[1]")))
            # start_close_icon = driver.find_element(By.CSS_SELECTOR, "#tngay > div > i.anticon.anticon-close-circle.ant-calendar-picker-clear")
            start_close_icon.click()
        except NoSuchElementException: print("start day is already clear")
        # change end date
        ActionChains(driver).move_to_element(end_input).click().perform()
        ActionChains(driver).move_to_element(end_input).click().send_keys(end.strftime("%d/%m/%Y")).perform()
        end_selected_day = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "td.ant-calendar-cell.ant-calendar-selected-day")))
        # end_selected_day = driver.find_element(By.CSS_SELECTOR, 'td.ant-calendar-cell.ant-calendar-selected-day')
        end_selected_day.click()
        # change start day
        ActionChains(driver).move_to_element(start_input).click().perform()
        ActionChains(driver).move_to_element(start_input).click().send_keys(start.strftime("%d/%m/%Y")).perform()
        start_selected_day = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "td.ant-calendar-cell.ant-calendar-selected-day")))
        # start_selected_day = driver.find_element(By.CSS_SELECTOR, 'td.ant-calendar-cell.ant-calendar-selected-day')
        start_selected_day.click()
        
    def search():
        search_button.click()
        WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#__next > div.Spin__SpinWrapper-sc-1px1x1e-0.enNzTo > div > div")))
        WebDriverWait(driver, 60).until(EC.invisibility_of_element((By.CSS_SELECTOR, "#__next > div.Spin__SpinWrapper-sc-1px1x1e-0.enNzTo > div > div")))
        res = driver.find_element(by=By.CSS_SELECTOR, value="#__next > section > section > main > div > div > div > div > div.ant-tabs-content.ant-tabs-content-animated.ant-tabs-top-content > div.ant-tabs-tabpane.ant-tabs-tabpane-active > div.ant-row > div:nth-child(2) > div.ant-row-flex.ant-row-flex-space-between.ant-row-flex-middle > div:nth-child(1) > div > span")
        # res.text example: Có 3 kết quả
        return int(res.text.split(" ")[1])
   
    def recover(target, index):
        print("recover")
        sleep(10)
        while True:
            try:
                open_web(driver, url)
                mua_vao_tab = driver.find_element(by=By.CSS_SELECTOR, value="#__next > section > section > main > div > div > div > div > div.ant-tabs-bar.ant-tabs-top-bar > div > div > div > div > div:nth-child(1) > div:nth-child(2)")
                mua_vao_tab.click()
                break
            except Exception: # in case being block because of too many request, then need to wait and relogin
                print("deep sleep")
                sleep(10*60)
                log_in(driver, using_username, pw)
            
        # refind all needed element
        nonlocal start_input
        nonlocal start_calendar_icon
        nonlocal end_input
        nonlocal end_calendar_icon
        nonlocal search_button
        nonlocal ket_qua_kiem_tra_dropdown
        nonlocal ket_qua_kiem_tra_list
        
        start_input = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH, "/html/body/div/section/section/main/div/div/div/div/div[3]/div[2]/div[3]/div[1]/div/div/form/div[1]/div[4]/div/div[2]/div/div[2]/div/div/div/span/span/div/input")))
        start_calendar_icon = driver.find_element(By.XPATH, "/html/body/div/section/section/main/div/div/div/div/div[3]/div[2]/div[3]/div[1]/div/div/form/div[1]/div[4]/div/div[2]/div/div[2]/div/div/div/span/span/div/i[2]")
        end_input = driver.find_element(by=By.XPATH, value="/html/body/div/section/section/main/div/div/div/div/div[3]/div[2]/div[3]/div[1]/div/div/form/div[1]/div[4]/div/div[3]/div/div[2]/div/div/div/span/span/div/input")
        end_calendar_icon = driver.find_element(By.XPATH, "/html/body/div/section/section/main/div/div/div/div/div[3]/div[2]/div[3]/div[1]/div/div/form/div[1]/div[4]/div/div[3]/div/div[2]/div/div/div/span/span/div/i[2]")
        search_button = driver.find_element(by=By.XPATH, value="/html/body/div/section/section/main/div/div/div/div/div[3]/div[2]/div[3]/div[1]/div/div/form/div[3]/div[1]/button")
        ket_qua_kiem_tra_dropdown = driver.find_element(by=By.XPATH, value="/html/body/div[1]/section/section/main/div/div/div/div/div[3]/div[2]/div[3]/div[1]/div/div/form/div[1]/div[5]/div/div[2]/div/div/div/span/div/div")
        ket_qua_kiem_tra_dropdown.click()
        sleep(0.25)
        ket_qua_kiem_tra_dropdown.click()
        ket_qua_kiem_tra_list = driver.find_elements(by=By.CSS_SELECTOR, value="ul.ant-select-dropdown-menu.ant-select-dropdown-menu-root.ant-select-dropdown-menu-vertical li")

        enter_date()
        page_text = driver.find_element(by=By.XPATH, value="/html/body/div[1]/section/section/main/div/div/div/div/div[3]/div[2]/div[3]/div[2]/div[1]/div[2]/div/div[2]/div").text
        page_number = int(page_text.split('/')[0].strip())
        
        ket_qua_kiem_tra_dropdown.click()
        ket_qua_kiem_tra = ket_qua_kiem_tra_list[index]
        ket_qua_kiem_tra.click()
        search()
        
        while page_number!=target:
            next_page_button = driver.find_element(by=By.XPATH, value="/html/body/div[1]/section/section/main/div/div/div/div/div[3]/div[2]/div[3]/div[2]/div[1]/div[2]/div/div[3]/button")
            next_page_button.click()
            # wait for the page to change
            WebDriverWait(driver, 5*60).until(lambda d: d.find_element(by=By.XPATH, value="/html/body/div[1]/section/section/main/div/div/div/div/div[3]/div[2]/div[3]/div[2]/div[1]/div[2]/div/div[2]/div").text!=page_text)
            page_text = driver.find_element(by=By.XPATH, value="/html/body/div[1]/section/section/main/div/div/div/div/div[3]/div[2]/div[3]/div[2]/div[1]/div[2]/div/div[2]/div").text
            page_number = int(page_text.split('/')[0].strip())
 
    final_khms = []
    final_ky_hieu = []
    final_so_hoa_don = []
    final_kyhieu_sohoadon = []
    final_ngaylap = []
    final_mccqt = []
    final_MST_nguoimua = []
    final_ten_nguoimua = []
    final_dia_chi_nguoimua = []
    final_MST_nguoiban = []
    final_ten_nguoiban = []
    final_dia_chi_nguoiban = []
    final_tien_chua_thue = []
    final_tien_thue = []
    final_chiet_khau_thuong_mai = []
    final_tien_phi = []
    final_tien_thanh_toan = []
    final_don_vi_tien = []
    final_hinh_thuc_thanh_toan = []
    final_thue_suat = []
    final_tong_tien_bang_so = []
    final_tong_tien_bang_chu = []
    final_trang_thai = []
    final_ket_qua = []

    final_khshd_list = []
    final_tinh_chat_list = []
    final_ten_list = []
    final_don_vi_list = []
    final_so_luong_list = []
    final_don_gia_list = []
    final_chiet_khau_list = []
    final_thue_suat_list = []
    final_thanh_tien_list = []
    final_tien_thue_list = []
    final_tong_tien_list = []
    final_ten_target_list = []
    final_mst_target_list = []
    final_trang_thai_hoa_don_list = []
    
    enter_date()
    total = 0
    ket_qua_kiem_tra_index = 0
    finish = True
    
    while ket_qua_kiem_tra_index<len(ket_qua_kiem_tra_list):
        ket_qua_kiem_tra = ket_qua_kiem_tra_list[ket_qua_kiem_tra_index]
        if ket_qua_kiem_tra_index<2:
            hoa_don_tab = driver.find_element(By.XPATH, "/html/body/div[1]/section/section/main/div/div/div/div/div[3]/div[2]/div[3]/div[2]/div[2]/div[1]/div/div/div/div/div[1]/div[1]")
        else:
            hoa_don_tab = driver.find_element(By.XPATH,"/html/body/div[1]/section/section/main/div/div/div/div/div[3]/div[2]/div[3]/div[2]/div[2]/div[1]/div/div/div/div/div[1]/div[2]")
        hoa_don_tab.click()
        sleep(10)
        ket_qua_kiem_tra_dropdown.click()
        ket_qua_kiem_tra.click()
        number = search()
        if number==0:
            ket_qua_kiem_tra_index+=1
            continue
        counter = 0
        total+=number
        page_now = 1
        while counter!=number:
            temp_khms = []
            temp_ky_hieu = []
            temp_so_hoa_don = []
            temp_kyhieu_sohoadon = []
            temp_ngaylap = []
            temp_mccqt = []
            temp_MST_nguoimua = []
            temp_ten_nguoimua = []
            temp_dia_chi_nguoimua = []
            temp_MST_nguoiban = []
            temp_ten_nguoiban = []
            temp_dia_chi_nguoiban = []
            temp_tien_chua_thue = []
            temp_tien_thue = []
            temp_chiet_khau_thuong_mai = []
            temp_tien_phi = []
            temp_tien_thanh_toan = []
            temp_don_vi_tien = []
            temp_hinh_thuc_thanh_toan = []
            temp_thue_suat = []
            temp_tong_tien_bang_so = []
            temp_tong_tien_bang_chu = []
            temp_trang_thai = []
            temp_ket_qua = []

            temp_khshd_list = []
            temp_tinh_chat_list = []
            temp_ten_list = []
            temp_don_vi_list = []
            temp_so_luong_list = []
            temp_don_gia_list = []
            temp_chiet_khau_list = []
            temp_thue_suat_list = []
            temp_thanh_tien_list = []
            temp_tien_thue_list = []
            temp_tong_tien_list = []
            temp_ten_target_list = []
            temp_mst_target_list = []
            temp_trang_thai_hoa_don_list = []
            
            stt_list = driver.find_elements(By.CSS_SELECTOR, "tr.ant-table-row.ant-table-row-level-0 > td:nth-child(1) > span")
            khms_list = driver.find_elements(By.CSS_SELECTOR, "tr.ant-table-row.ant-table-row-level-0 > td:nth-child(3)")
            ky_hieu_list = driver.find_elements(By.CSS_SELECTOR, "tr.ant-table-row.ant-table-row-level-0 > td:nth-child(4) > span")
            so_hoa_don_list = driver.find_elements(By.CSS_SELECTOR, "tr.ant-table-row.ant-table-row-level-0 > td:nth-child(5)")
            thong_tin_list = driver.find_elements(By.CSS_SELECTOR, "tr.ant-table-row.ant-table-row-level-0 > td:nth-child(7) > div > div")
            if ket_qua_kiem_tra_index<2:
                trang_thai_list = driver.find_elements(By.CSS_SELECTOR, "tr.ant-table-row.ant-table-row-level-0 > td:nth-child(14) > span")
                ket_qua_list = driver.find_elements(By.CSS_SELECTOR, "tr.ant-table-row.ant-table-row-level-0 > td:nth-child(15) > span")
            else:
                trang_thai_list = driver.find_elements(By.CSS_SELECTOR, "tr.ant-table-row.ant-table-row-level-0 > td:nth-child(12) > span")
                ket_qua_list = driver.find_elements(By.CSS_SELECTOR, "tr.ant-table-row.ant-table-row-level-0 > td:nth-child(13) > span")
            element_list = driver.find_elements(by=By.CSS_SELECTOR, value="tr.ant-table-row.ant-table-row-level-0")
            
            def network_test():
                try:
                    if page_now<math.ceil(number/15): target = 15
                    else: target = number-counter
                    network_counter = 0
                    for element, stt in zip(element_list, stt_list):
                        if not stt.text.strip(): continue
                        network_counter+=1
                        element.click()
                        sleep(0.1)
                    return network_counter==target
                except Exception:
                    return False
            
            if not network_test():
                recover(page_now, ket_qua_kiem_tra_index)
                continue
            
            if finish:
                success_list = [False]*len(stt_list)
                res_list = [None]*len(stt_list)
                finish = False
            # print("sst", len(stt_list))
            # print("khms", len(khms_list))
            # print("ky hieu", len(ky_hieu_list))
            # print("so hoa don", len(so_hoa_don_list))
            # print("thong tin", len(thong_tin_list))
            # print("success", len(success_list))
            async def getData():
                tasks = []  
                for stt, khms, kh, so_hoa_don, thong_tin, success\
                in zip(stt_list, khms_list, ky_hieu_list, so_hoa_don_list, thong_tin_list, success_list):
                    if not stt.text.strip(): tasks.append(asyncio.create_task(returnNone()))
                    elif success: tasks.append(asyncio.create_task(returnNone()))
                    else: 
                        if "MST" in thong_tin.text:
                            mst = thong_tin.text.split("\n")[0].split(" ")[-1]
                        else:
                            mst = using_username
                        tasks.append(asyncio.create_task(get_receipt_by_API(auth=auth, ma_so_thue=mst, ky_hieu_hoa_don=kh.text, so_hoa_don=so_hoa_don.text, ky_hieu_mau_so=khms.text, index=ket_qua_kiem_tra_index)))
                data_list = await asyncio.gather(*tasks)
                return data_list
            
            data_list = asyncio.run(getData())
            need_recover = False
            need_new_auth = False
            
            for i in range(len(stt_list)):
                success = success_list[i]
                if not success:
                    stt = stt_list[i]
                    if not stt.text.strip():
                        success_list[i] = True
                    else:
                        data = data_list[i]
                        api_res = data[0]
                        if api_res=="good":
                            success_list[i] = True
                            res_list[i] = data[1]
                        elif api_res=="bad auth":
                            need_new_auth = True
                        else:
                            need_recover = True
            
            
            if need_recover:
                recover(page_now, ket_qua_kiem_tra_index)
            elif need_new_auth:
                try:
                    auth = getAuth(banra=False, driver_path=driver_path)
                except NoSuchElementException: recover(page_now, ket_qua_kiem_tra_index)
            else:
                finish = True
                for stt, khms, kh, so_hoa_don, trang_thai, ket_qua, API_res\
                in zip(stt_list, khms_list, ky_hieu_list, so_hoa_don_list,trang_thai_list, ket_qua_list, res_list):
                    if not stt.text.strip(): continue
                    counter+=1
                    
                    invalid_dich_vu_ct = 0
                    
                    def format_none(target):
                        if target is None: return "n/a"
                        else:
                            target = str(target)
                            return target.strip()
                        
                    def format_money(target):
                        if target is None: return "n/a"
                        return str(int(target))
                    
                    try:
                        current_khms = format_empty_str(khms.text)
                        current_ky_hieu = format_empty_str(kh.text)
                        current_so_hoa_don = format_empty_str(so_hoa_don.text)
                        current_trang_thai_hoa_don = format_empty_str(trang_thai.text)
                        current_ket_qua = format_empty_str(ket_qua.text)
                        
                        current_mst_nguoiban = format_none(API_res["nbmst"])
                        current_kyhieu_sohoadon = f"{current_ky_hieu}_{current_so_hoa_don}_{current_mst_nguoiban}"
                        
                        current_ngaylap = format_none(API_res["tdlap"][:10])
                        current_mccqt = format_none(API_res["mhdon"])
                        current_ten_nguoiban = format_none(API_res["nbten"])
                        current_dia_chi_nguoiban = format_none(API_res["nbdchi"])
                        current_ten_nguoimua = format_none(API_res["nmten"])
                        current_mst_nguoimua = format_none(API_res["nmmst"])
                        current_dia_chi_nguoimua = format_none(API_res["nmdchi"])
                        
                        if "thtttoan" in API_res:
                            current_hinh_thuc_thanh_toan = format_none(API_res["thtttoan"])
                        else:
                            current_hinh_thuc_thanh_toan = "n/a"
                        
                        if "tgtphi" in API_res:
                            current_tien_phi = format_money(API_res["tgtphi"])
                        else:
                            current_tien_phi = "n/a"
                            
                        if "dvtte" in API_res:
                            current_dvtte = format_none(API_res["dvtte"])
                        else:
                            current_dvtte = "n/a"
                            
                        if len(API_res["thttltsuat"])!=0 and "tsuat" in API_res["thttltsuat"][0]:
                            thue_suat_dic = API_res["thttltsuat"][0]
                            current_thue_suat = format_none(thue_suat_dic["tsuat"])
                        else: 
                            current_thue_suat = "n/a"
                        
                        current_tong_tien_bang_so = format_money(API_res["tgtttbso"])
                        current_tong_tien_bang_chu = format_none(API_res["tgtttbchu"])
                        current_tien_chua_thue = format_money(API_res["tgtcthue"])
                        current_tien_thue = format_money(API_res["tgtthue"])
                        current_chiet_khau_thuong_mai = format_money(API_res["ttcktmai"])
                        current_tien_thanh_toan = format_money(API_res["tgtttbso"])
                        
                        temp_khms.append(current_khms)
                        temp_ky_hieu.append(current_ky_hieu)
                        temp_so_hoa_don.append(current_so_hoa_don)
                        temp_trang_thai.append(current_trang_thai_hoa_don)
                        temp_ket_qua.append(current_ket_qua)
                        temp_kyhieu_sohoadon.append(current_kyhieu_sohoadon)
                        temp_ngaylap.append(current_ngaylap)
                        temp_mccqt.append(current_mccqt)
                        temp_ten_nguoiban.append(current_ten_nguoiban)
                        temp_MST_nguoiban.append(current_mst_nguoiban)
                        temp_dia_chi_nguoiban.append(current_dia_chi_nguoiban)
                        temp_ten_nguoimua.append(current_ten_nguoimua)
                        temp_MST_nguoimua.append(current_mst_nguoimua)
                        temp_dia_chi_nguoimua.append(current_dia_chi_nguoimua)
                        temp_hinh_thuc_thanh_toan.append(current_hinh_thuc_thanh_toan)
                        temp_tien_phi.append(current_tien_phi)
                        temp_don_vi_tien.append(current_dvtte)
                        temp_thue_suat.append(current_thue_suat)
                        temp_tong_tien_bang_so.append(current_tong_tien_bang_so)
                        temp_tong_tien_bang_chu.append(current_tong_tien_bang_chu)
                        temp_tien_chua_thue.append(current_tien_chua_thue)
                        temp_tien_thue.append(current_tien_thue)
                        temp_chiet_khau_thuong_mai.append(current_chiet_khau_thuong_mai)
                        temp_tien_thanh_toan.append(current_tien_thanh_toan)
                    except Exception:
                        # neu hoa don bi loi thi se ko parse hang hoa dich vu trong hoa don
                        error_string = f"LOI FORMAT HOA DON MUA VAO. SO HOA DON: {so_hoa_don.text}"
                        invalid_hoa_don_list.append(error_string)
                    
                        print("-------------------------------------------------")
                        print(error_string)
                        print_exc()
                        print("-------------------------------------------------")

                        continue
                    
                    

                    hang_hoa_dich_vu_list = API_res["hdhhdvu"]
                    
                    def format_thue_suat(thue_suat):
                        thue = round(thue_suat*100,2)
                        string = f"{thue}%"
                        string = string.replace(".",",")
                        return string
                    def format_tien_thue(thue):
                        string = str(thue)
                        string = string.replace(".", ",")
                        return string
                    def get_thue_suat(target):
                        if target is None: return 0
                        return float(target)
                    def format_tinh_chat(target):
                            if target is None: return "n/a"
                            return tinh_chat_dic[target]
                    def get_thanh_tien(target):
                        if target is None: return 0
                        return int(target)
                    
                    
                    for hhdv_dic in hang_hoa_dich_vu_list:
                        try:
                            current_tinh_chat = format_tinh_chat(hhdv_dic["tchat"])
                            current_ten = format_none(hhdv_dic["ten"])
                            current_don_vi = format_none(hhdv_dic["dvtinh"])
                            current_so_luong = format_none(hhdv_dic["sluong"]).replace(".",",")
                            current_don_gia = format_none(hhdv_dic["dgia"]).replace(".",",")
                            current_chiet_khau = format_none(hhdv_dic["stckhau"]).replace(".",",")
                            
                            current_thanh_tien = get_thanh_tien(hhdv_dic["thtien"])
                            if current_thanh_tien==0:
                                current_thue_suat = "n/a"
                                current_tien_thue = "n/a"
                                current_tong_tien = "n/a"
                            else:
                                current_thue_suat = get_thue_suat(hhdv_dic["tsuat"])
                                current_tien_thue = round(current_thanh_tien*current_thue_suat,2)
                                current_tong_tien = current_thanh_tien+current_tien_thue
                                
                                # format thue suat from float to x,x% format
                                current_thue_suat = format_thue_suat(current_thue_suat)
                                # format tien thue from float to x,x format
                                current_tien_thue = format_tien_thue(current_tien_thue)
                                # format tong tien from float to x,x format
                                current_tong_tien = format_tien_thue(current_tong_tien)
                                
                            temp_tinh_chat_list.append(current_tinh_chat)
                            temp_ten_list.append(current_ten)
                            temp_don_vi_list.append(current_don_vi)
                            temp_so_luong_list.append(current_so_luong)
                            temp_don_gia_list.append(current_don_gia)
                            temp_chiet_khau_list.append(current_chiet_khau)
                            temp_thue_suat_list.append(current_thue_suat)
                            temp_thanh_tien_list.append(current_thanh_tien)
                            temp_tien_thue_list.append(current_tien_thue)
                            temp_tong_tien_list.append(current_tong_tien)
                            temp_khshd_list.append(current_kyhieu_sohoadon)
                            temp_ten_target_list.append(current_ten_nguoiban)
                            temp_mst_target_list.append(current_mst_nguoiban)
                            temp_trang_thai_hoa_don_list.append(current_trang_thai_hoa_don)
                        
                        except Exception:
                            invalid_dich_vu_ct+=1
                            print("-------------------------------------------------")
                            print(f"LOI FORMAT DICH VU TRONG HOA DON MUA VAO. SO HOA DON: {so_hoa_don.text}")
                            print_exc() 
                            print("-------------------------------------------------")  
                            
                    if invalid_dich_vu_ct!=0:
                        error_string = f"{invalid_dich_vu_ct} LOI FORMAT DICH VU TRONG HOA DON MUA VAO. SO HOA DON: {so_hoa_don.text}"
                        invalid_dich_vu_list.append(error_string)              
                    
            if finish:
                final_khms.extend(temp_khms)
                final_ky_hieu.extend(temp_ky_hieu)
                final_so_hoa_don.extend(temp_so_hoa_don)
                final_kyhieu_sohoadon.extend(temp_kyhieu_sohoadon)
                final_ngaylap.extend(temp_ngaylap)
                final_mccqt.extend(temp_mccqt)
                final_MST_nguoimua.extend(temp_MST_nguoimua)
                final_ten_nguoimua.extend(temp_ten_nguoimua)
                final_dia_chi_nguoimua.extend(temp_dia_chi_nguoimua)
                final_MST_nguoiban.extend(temp_MST_nguoiban)
                final_ten_nguoiban.extend(temp_ten_nguoiban)
                final_dia_chi_nguoiban.extend(temp_dia_chi_nguoiban)
                final_tien_chua_thue.extend(temp_tien_chua_thue)
                final_tien_thue.extend(temp_tien_thue)
                final_chiet_khau_thuong_mai.extend(temp_chiet_khau_thuong_mai)
                final_tien_phi.extend(temp_tien_phi)
                final_tien_thanh_toan.extend(temp_tien_thanh_toan)
                final_don_vi_tien.extend(temp_don_vi_tien)
                final_hinh_thuc_thanh_toan.extend(temp_hinh_thuc_thanh_toan)
                final_thue_suat.extend(temp_thue_suat)
                final_tong_tien_bang_so.extend(temp_tong_tien_bang_so)
                final_tong_tien_bang_chu.extend(temp_tong_tien_bang_chu)
                final_trang_thai.extend(temp_trang_thai)
                final_ket_qua.extend(temp_ket_qua)

                final_khshd_list.extend(temp_khshd_list)
                final_tinh_chat_list.extend(temp_tinh_chat_list)
                final_ten_list.extend(temp_ten_list)
                final_don_vi_list.extend(temp_don_vi_list)
                final_so_luong_list.extend(temp_so_luong_list)
                final_don_gia_list.extend(temp_don_gia_list)
                final_chiet_khau_list.extend(temp_chiet_khau_list)
                final_thue_suat_list.extend(temp_thue_suat_list)
                final_thanh_tien_list.extend(temp_thanh_tien_list)
                final_tien_thue_list.extend(temp_tien_thue_list)
                final_tong_tien_list.extend(temp_tong_tien_list)
                final_ten_target_list.extend(temp_ten_target_list)
                final_mst_target_list.extend(temp_mst_target_list)
                final_trang_thai_hoa_don_list.extend(temp_trang_thai_hoa_don_list)
 
            # if there is more, turn to the next page
            if counter!=number and finish:
                initial_page = driver.find_element(by=By.XPATH, value="/html/body/div[1]/section/section/main/div/div/div/div/div[3]/div[2]/div[3]/div[2]/div[1]/div[2]/div/div[2]/div").text
                next_page_button = driver.find_element(by=By.XPATH, value="/html/body/div[1]/section/section/main/div/div/div/div/div[3]/div[2]/div[3]/div[2]/div[1]/div[2]/div/div[3]/button")
                next_page_button.click()
                # wait for the page to change
                WebDriverWait(driver, 5*60).until(lambda d: d.find_element(by=By.XPATH, value="/html/body/div[1]/section/section/main/div/div/div/div/div[3]/div[2]/div[3]/div[2]/div[1]/div[2]/div/div[2]/div").text!=initial_page)
                page_now+=1
                if safemode: sleep(30)
                else: sleep(15)
            
        if safemode: sleep(30)
        else: sleep(15)
        ket_qua_kiem_tra_index+=1
        
    if total==0: return None, None, auth, invalid_hoa_don_list, invalid_dich_vu_list
    
    df = pd.DataFrame({
        'final_khms': final_khms,
        'final_ky_hieu': final_ky_hieu,
        'final_so_hoa_don': final_so_hoa_don,
        'final_kyhieu_sohoadon': final_kyhieu_sohoadon,
        'final_ngaylap': final_ngaylap,
        'final_mccqt': final_mccqt,
        'final_MST_nguoimua': final_MST_nguoimua,
        'final_ten_nguoimua': final_ten_nguoimua,
        'final_dia_chi_nguoimua': final_dia_chi_nguoimua,
        'final_MST_nguoiban': final_MST_nguoiban,
        'final_ten_nguoiban': final_ten_nguoiban,
        'final_dia_chi_nguoiban': final_dia_chi_nguoiban,
        'final_tien_chua_thue': final_tien_chua_thue,
        'final_tien_thue': final_tien_thue,
        'final_chiet_khau_thuong_mai': final_chiet_khau_thuong_mai,
        'final_tien_phi': final_tien_phi,
        'final_tien_thanh_toan': final_tien_thanh_toan,
        'final_don_vi_tien': final_don_vi_tien,
        'final_hinh_thuc_thanh_toan': final_hinh_thuc_thanh_toan,
        'final_thue_suat': final_thue_suat,
        'final_tong_tien_bang_so': final_tong_tien_bang_so,
        'final_tong_tien_bang_chu': final_tong_tien_bang_chu,
        'final_trang_thai': final_trang_thai,
        'final_ket_qua': final_ket_qua,
    })

    hanghoa_df = pd.DataFrame({
        "kyhieu_sohoadon": final_khshd_list,
        "tinh_chat": final_tinh_chat_list, 
        "ten": final_ten_list, 
        "don_vi": final_don_vi_list,
        "so_luong": final_so_luong_list, 
        "don_gia": final_don_gia_list, 
        "chiet_khau": final_chiet_khau_list, 
        "thue_suat": final_thue_suat_list, 
        "thanh_tien": final_thanh_tien_list,
        "tien_thue": final_tien_thue_list,
        "tong_tien": final_tong_tien_list,
        "ten_target": final_ten_target_list,
        "mst_target": final_mst_target_list,
        "trang_thai_hoa_don": final_trang_thai_hoa_don_list
    })
    return df, hanghoa_df, auth, invalid_hoa_don_list, invalid_dich_vu_list

def log_in(driver:webdriver, user:str, pw:str):
    sleep(1.5)
    pyautogui.click()
    sleep(1.5)
    url = "https://hoadondientu.gdt.gov.vn/"
    try: 
        driver.maximize_window()
    except WebDriverException: ""
    open_web(driver, url)
    focused_window = gw.getActiveWindow()
    window_title = focused_window.title
    while "google chrome" not in window_title.lower():
        pyautogui.click()
        sleep(1)
        focused_window = gw.getActiveWindow()
        if focused_window is not None: window_title = str(focused_window.title).strip()
        messagebox.showinfo("Chuyển Tab", "Vui Lòng Chuyển Tab Qua Chrome!")
        sleep(2)
    # close the warning button if exist
    try:
        warning_button = driver.find_element(by=By.CSS_SELECTOR, value="body > div:nth-child(4) > div > div.ant-modal-wrap > div > div.ant-modal-content > button")
        warning_button.click()
    except NoSuchElementException: print("No warning")
    # press the log in button and enter username, password
    log_in_button = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#__next > section > header > div.home-header-menu > div > div:nth-child(6) > span")))
    log_in_button.click()
    # enter username and password
    user_input = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#username")))
    # sleep(uniform(3,5))# simulate human typing time
    user_input.send_keys(user)
    pw_input = driver.find_element(by=By.CSS_SELECTOR, value="#password")
    # sleep(uniform(3,5))# simulate human typing time
    pw_input.send_keys(pw)
    # after enter capcha, log out button would appear. Wait for the logout button appear
    try:
        # wait for finish entering capcha
        avatar = WebDriverWait(driver, 24*60*60).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#__next > section > header > div.home-header-buttons > button:nth-child(2)")))
        WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#__next > div.Spin__SpinWrapper-sc-1px1x1e-0.enNzTo > div > div")))
        WebDriverWait(driver, 60).until(EC.invisibility_of_element((By.CSS_SELECTOR, "#__next > div.Spin__SpinWrapper-sc-1px1x1e-0.enNzTo > div > div")))
        sleep(3)
        avatar = WebDriverWait(driver, 24*60*60).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#__next > section > header > div.home-header-buttons > button:nth-child(2)")))
        avatar.click()
    except (StaleElementReferenceException,TimeoutException): 
        print("cannot log in")
        return False    
    # Wait for the page to finish logging in
    return True
    
def log_out(driver:webdriver):
    log_out_butt = driver.find_element(by=By.CSS_SELECTOR, value="#__next > section > header > div.home-header-buttons > button:nth-child(3)")
    log_out_butt.click()
    sleep(3)

def open_chrome():    
    subprocess.Popen([batch_script], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    sleep(5)  # wait for chrome to open
    
    for i in range(10):
        check_process = subprocess.run(["netstat", "-an"], capture_output=True, text=True)
        if "9222" in check_process.stdout and "LISTENING" in check_process.stdout:
            return True
        # print("sleeping")
        sleep(3)
    return False

def main():
    if not open_chrome(): 
        print("cannot open chrome")
        return
    options = webdriver.ChromeOptions()
    options.debugger_address = "127.0.0.1:9222"  

    service = Service(driver_path)

    driver = webdriver.Chrome(service=service, options=options)
    
    if os.path.exists(f"{exe_path}/users/usernameDic.pkl"):
        with open(f"{exe_path}/users/usernameDic.pkl", "rb") as file:
            usernameDic = pickle.load(file)
            companyList = list(usernameDic.keys())
        with open(f"{exe_path}/users/pwDic.pkl", "rb") as file:
            pwDic = pickle.load(file)
        with open(f"{exe_path}/users/companyDic.pkl", "rb") as file:
            companyDic = pickle.load(file)

    else: 
        usernameDic = {" ":" "} # key is company name, value is username
        companyList = list(usernameDic.keys())
        pwDic = {} # key is username, value is password
        companyDic = {" ":"Không Có"} # key is username, value is company name
        
    
    def ban_ra(start:datetime, end:datetime, text:str):
        nonlocal ban_ra_date_set
        nonlocal auth
        nonlocal using_username
        nonlocal pwDic
        nonlocal safemode
        date_range = get_range(start, end, ban_ra_date_set)
        if date_range:
            final_df = pd.DataFrame(columns=df_columns)
            final_hanghoa_df = pd.DataFrame(columns=hanghoa_df_columns)
            for start, end in date_range:
                df, hanghoa_df, auth, invalid_hang_hoa_list, invalid_dich_vu_list = enter_ban_ra(driver, start, end, auth, using_username, pwDic[using_username], safemode.get())
                if df is not None: 
                    final_df = pd.concat([final_df, df], ignore_index=True)
                    final_hanghoa_df = pd.concat([final_hanghoa_df, hanghoa_df], ignore_index=True)
            nonlocal idx_counter
            scrape_dic[text] = idx_counter
            final_df['counter'] = idx_counter
            final_hanghoa_df['counter'] = idx_counter
            idx_counter+=1
            return final_df, final_hanghoa_df, invalid_hang_hoa_list, invalid_dich_vu_list
        else: return None, None, None, None

    def mua_vao(start:datetime, end:datetime, text:str):
        nonlocal mua_vao_date_set
        nonlocal auth
        nonlocal using_username
        nonlocal pwDic
        nonlocal safemode
        date_range = get_range(start, end, mua_vao_date_set)
        if date_range:
            final_df = pd.DataFrame(columns=df_columns)
            final_hanghoa_df = pd.DataFrame(columns=hanghoa_df_columns)
            for start, end in date_range:
                df, hanghoa_df, auth, invalid_hang_hoa_list, invalid_dich_vu_list = enter_mua_vao(driver, start, end, auth, using_username, pwDic[using_username], safemode.get())
                if df is not None: 
                    final_df = pd.concat([final_df, df], ignore_index=True)
                    final_hanghoa_df = pd.concat([final_hanghoa_df, hanghoa_df], ignore_index=True)
            nonlocal idx_counter
            scrape_dic[text] = idx_counter
            final_df['counter'] = idx_counter
            final_hanghoa_df['counter'] = idx_counter
            if "cả" not in text: idx_counter+=1 # if ca hai, then add up the counter in the ban_ra
            return final_df, final_hanghoa_df, invalid_hang_hoa_list, invalid_dich_vu_list
        else: return None, None, None, None
        
    def get_range(start:datetime, end:datetime, having_set:set):
        """
        find the missing dates to avoid duplicate scrape
        """
        all_date = set()
        cur_date = start
        while cur_date<=end:
            all_date.add(cur_date.date())
            cur_date+=timedelta(days=1)

        missing_date = sorted(all_date - having_set)
        if missing_date:
            missing_range = []
            iterator = iter(missing_date)
            start_range = next(iterator)
            max_range = monthrange(start_range.year,start_range.month)[1] # the website only accept a specific range, depends on the start date
            having_set.add(start_range) # add missing date into the having_set
            prev = start_range
            while True:
                try:
                    next_date = next(iterator)
                    having_set.add(next_date) # add missing date into the having_set
                    if next_date!=prev + timedelta(days=1) or (next_date-start_range).days+1>max_range:
                        missing_range.append((start_range, prev))
                        start_range = next_date
                        prev = next_date
                        max_range = monthrange(start_range.year,start_range.month)[1]
                    else: prev = next_date
                except StopIteration: break
            missing_range.append((start_range, prev))
            return missing_range

    # create root
    root = tk.Tk()
    s = ttk.Style(root)
    s.theme_use('clam')
    root.state("zoomed")
    root.title("Hoa Don Dien Tu")

    # variables
    logged_in = False
    using_username = " "
    auth = ""
    hoa_don = "bán ra"
    history = [" "]
    ban_ra_date_set = set()
    mua_vao_date_set = set()
    idx_counter = 0
    df_columns = [
        'final_khms', 'final_ky_hieu', 'final_so_hoa_don', 'final_kyhieu_sohoadon','final_ngaylap', 'final_mccqt',
        'final_MST_nguoimua', 'final_ten_nguoimua', 'final_dia_chi_nguoimua',
        'final_MST_nguoiban', 'final_ten_nguoiban', 'final_dia_chi_nguoiban',
        'final_tien_chua_thue', 'final_tien_thue','final_chiet_khau_thuong_mai', 'final_tien_phi', 'final_tien_thanh_toan',
        'final_don_vi_tien', 'final_hinh_thuc_thanh_toan', 'final_thue_suat', 'final_tong_tien_bang_so', 'final_tong_tien_bang_chu',
        'final_trang_thai', 'final_ket_qua'
    ]
    hanghoa_df_columns = [
        "kyhieu_sohoadon", "tinh_chat", "ten", "don_vi", "so_luong", "don_gia", "chiet_khau", "thue_suat",
        "thanh_tien", "tien_thue", "tong_tien", "ten_target", "mst_target", "trang_thai_hoa_don"
    ]
    scrape_dic = dict()
    ban_ra_df = pd.DataFrame(columns= df_columns)
    ban_ra_df['counter'] = 0
    ban_ra_hanghoa_df = pd.DataFrame(columns=hanghoa_df_columns)
    ban_ra_hanghoa_df['counter'] = 0
    mua_vao_df = pd.DataFrame(columns=df_columns)
    mua_vao_df['counter'] = 0
    mua_vao_hanghoa_df = pd.DataFrame(columns=hanghoa_df_columns)
    mua_vao_hanghoa_df['counter'] = 0
    
    # create frames. There are 5 frames: 2 calender frames stay side by side. 1 frame for loai hoa don button. 1 frame for submit button and dropdown. 1 frame for loading and result.
    max_date = datetime.now() # maxdate for calender widget
    above_frame = tk.Frame(root)
    above_frame.pack(pady=(20), anchor='w')

    first_cal_frame = tk.Frame(above_frame)
    first_cal_frame.pack(side='left', anchor='n', padx=(50,0))

    second_cal_frame = tk.Frame(above_frame)
    second_cal_frame.pack(side='left', anchor='n', padx=50)
    
    user_frame = tk.Frame(above_frame)
    user_frame.pack(side='left', anchor='n', padx=(0,50))

    button_frame = tk.Frame(root)
    button_frame.pack(anchor='w', padx=50)

    submit_frame = tk.Frame(root)
    submit_frame.pack(anchor='w', padx=50)

    export_frame = tk.Frame(root)
    export_frame.pack(anchor='w', padx=50)

    # calendar 1
    cal1 = Calendar(first_cal_frame, font="Arial 15", selectmode='day', year=2024, month=1, day=1, maxdate=max_date)
    cal1.pack()

    title_label1 = tk.Label(first_cal_frame, text='Từ Ngày', font=("Arial", 15), relief="solid", bd=2, padx=10, pady=5)
    title_label1.pack(pady=10, padx=10)
    # calender 2
    cal2 = Calendar(second_cal_frame, font="Arial 15", selectmode='day', year=2024, month=1, day=1, maxdate=max_date)
    cal2.pack()

    title_label2 = tk.Label(second_cal_frame, text='Đến Ngày', font=("Arial", 15), relief="solid", bd=2, padx=10, pady=5)
    title_label2.pack(pady=10, padx=10)
    
    # user manager
    def open_registration_window():
        root.iconify()
        user_window = tk.Toplevel(root)
        user_window.title("Tạo Tài Khoản")
        user_window.geometry("500x500")

        title_label = tk.Label(user_window, text="Tạo Tài Khoản", font=("Arial", 15),relief="solid", bd=2,padx=10, pady=5)
        title_label.pack(pady=20)
        
        # Company Name Field
        company_label = tk.Label(user_window, text="Tên Công Ty:",font=("Arial", 15))
        company_label.pack()
        company_entry = tk.Entry(user_window, width=30)
        company_entry.pack()

        # Username Field
        username_label = tk.Label(user_window, text="Username:",font=("Arial", 15))
        username_label.pack()
        username_entry = tk.Entry(user_window, width=30)
        username_entry.pack()

        # Password Field
        password_label = tk.Label(user_window, text="Mật Khẩu:",font=("Arial", 15))
        password_label.pack()
        password_entry = tk.Entry(user_window, width=30, show="*")
        password_entry.pack()

        # Confirm Password Field
        confirm_password_label = tk.Label(user_window, text="Xác Nhận Mật Khẩu:",font=("Arial", 15))
        confirm_password_label.pack()
        confirm_password_entry = tk.Entry(user_window, width=30, show="*")
        confirm_password_entry.pack()
        
        def register():
            company = company_entry.get().strip()
            username = username_entry.get().strip()
            password = password_entry.get().strip()
            confirm_password = confirm_password_entry.get().strip()

            if not username or not company or not password or not confirm_password:
                messagebox.showerror("Error", "Nhập Tất Cả Các Ô!")
                return
            if username in pwDic:
                messagebox.showerror("Error", "Tài Khoản Đã Tồn Tại!")
                user_window.destroy()  
                root.deiconify()
                return
            if password != confirm_password:
                messagebox.showerror("Error", "Mật Khẩu Không Trùng!")
                return

            messagebox.showinfo("Thành Công", "Tạo Tài Khoản Thành Công")
            usernameDic[company] = username
            pwDic[username] = password
            companyList.append(company)
            companyDic[username] = company
            nonlocal user_dropdown
            nonlocal username_StringVar
            user_dropdown.destroy()
            username_StringVar.set(companyList[-1])
            user_dropdown = tk.OptionMenu(user_frame, username_StringVar, *companyList)
            user_dropdown.config(font=("Arial", 15), bg='yellow', fg='black')
            user_dropdown.pack()
            user_window.destroy()
            root.deiconify()

        register_button = tk.Button(user_window, text="Tạo", command=register, bg="blue", fg="white",font=("Arial", 15))
        register_button.pack(pady=20)
        
    def delete_user():
        nonlocal username_StringVar
        target = username_StringVar.get()
        if not target.strip(): return
        response = messagebox.askyesno("Xóa Tài Khoản", "Xác nhận xoá?")
        if response:
            companyList.remove(target)
            target_username = usernameDic[target]
            del companyDic[target_username]
            del usernameDic[target]
            del pwDic[target_username]
            nonlocal user_dropdown
            user_dropdown.destroy()
            username_StringVar.set(companyList[-1])
            user_dropdown = tk.OptionMenu(user_frame, username_StringVar, *companyList)
            user_dropdown.config(font=("Arial", 15), bg='yellow', fg='black')
            user_dropdown.pack()
    
    def log_in_user():
        nonlocal username_StringVar
        nonlocal using_username
        nonlocal logged_in
        nonlocal hoa_don
        nonlocal history
        nonlocal ban_ra_date_set
        nonlocal mua_vao_date_set
        nonlocal idx_counter
        nonlocal scrape_dic
        nonlocal ban_ra_df
        nonlocal ban_ra_hanghoa_df
        nonlocal mua_vao_df
        nonlocal mua_vao_hanghoa_df
        nonlocal auth
        
        target = username_StringVar.get().strip()
        if target and usernameDic[target]!=using_username:
            root.iconify()
            if logged_in:
                log_out(driver)
            if not log_in(driver,usernameDic[target], pwDic[usernameDic[target]]):
                messagebox.showwarning("Đăng Nhập Không Thành Công", "Đăng Nhập Không Thành Công")
                root.deiconify()
                return
            using_username = usernameDic[target]
            # reset variables
            logged_in = True
            auth = ""
            hoa_don = "bán ra"
            history = [" "]
            ban_ra_date_set = set()
            mua_vao_date_set = set()
            idx_counter = 0
            scrape_dic = dict()
            ban_ra_df = pd.DataFrame(columns= df_columns)
            ban_ra_df['counter'] = 0
            ban_ra_hanghoa_df = pd.DataFrame(columns=hanghoa_df_columns)
            ban_ra_hanghoa_df['counter'] = 0
            mua_vao_df = pd.DataFrame(columns=df_columns)
            mua_vao_df['counter'] = 0
            mua_vao_hanghoa_df = pd.DataFrame(columns=hanghoa_df_columns)
            mua_vao_hanghoa_df['counter'] = 0
            # reset dropdown
            nonlocal dropdown
            nonlocal selected_search
            dropdown.destroy()
            dropdown = tk.OptionMenu(submit_frame, selected_search, *history, command=lambda event: dropdown_func())
            dropdown.config(font=("Arial", 25), bg='yellow', fg='black')
            dropdown.pack(side='left')
            # reset res label and user_label
            nonlocal res_label
            res_label.config(text="")
            nonlocal user_label
            user_label.config(text=f"Tài Khoản Hiện Tại:\n{companyDic[using_username]}")
            
            root.deiconify()
            messagebox.showwarning("Đăng Nhập Thành Công", "Đăng Nhập Thành Công")
                    
    user_label = tk.Label(user_frame, text=f"Tài Khoản Hiện Tại:\n{companyDic[using_username]}", font=("Arial", 15), relief="solid", bd=2, padx=10, pady=5)
    user_label.pack()
    
    create_button = tk.Button(user_frame, text="Tạo Tài Khoản",font=("Arial", 15), command=open_registration_window)
    create_button.pack(padx=10)
    
    del_button = tk.Button(user_frame, text="Xóa Tài Khoản",font=("Arial", 15), command=delete_user)
    del_button.pack(padx=10)
    
    log_in_button = tk.Button(user_frame, text="Đăng Nhập",font=("Arial", 15), command=log_in_user)
    log_in_button.pack(padx=10)
    
    # this is username but it is company name 
    username_StringVar = tk.StringVar()
    username_StringVar.set(companyList[0])
    user_dropdown = tk.OptionMenu(user_frame, username_StringVar, *companyList)
    user_dropdown.config(font=("Arial", 15), bg='yellow', fg='black')
    user_dropdown.pack()

    # loai hoa don label
    button_text_label = tk.Label(button_frame, text="Loại Hoá Đơn:", font=("Arial", 15))
    button_text_label.pack(side="left")

    def update_label(loai_hoa_don):
        nonlocal hoa_don
        hoa_don= loai_hoa_don
        update_search_bar()
    # Button 1
    button_1 = tk.Button(button_frame, text="Bán Ra",font=("Arial", 15), command=lambda: update_label("bán ra"))
    button_1.pack(side="left", padx=10)

    # Button 2
    button_2 = tk.Button(button_frame, text="Mua Vào",font=("Arial", 15), command=lambda: update_label("mua vào"))
    button_2.pack(side="left", padx=10)

    # Button 3
    button_3 = tk.Button(button_frame, text="Cả Hai",font=("Arial", 15), command=lambda: update_label("cả hai"))
    button_3.pack(side="left", padx=10)
    def submit_func():
        nonlocal logged_in
        if not logged_in: 
            messagebox.showwarning("Đăng Nhập", "Đăng Nhập Vào Tài Khoản Để Tiếp Tục")
            return
        # search_text = f"Hoá đơn {hoa_don} từ {cal1.get_date()} đến {cal2.get_date()}"
        search_text = selected_search.get()
        if not search_text.strip(): return
        nonlocal ban_ra_df
        nonlocal ban_ra_hanghoa_df
        nonlocal mua_vao_df
        nonlocal mua_vao_hanghoa_df
        nonlocal companyDic
        nonlocal using_username
        root.iconify()
        sleep(1.5)
        pyautogui.click()
        sleep(1.5)
        focused_window = gw.getActiveWindow()
        window_title = focused_window.title
        while "google chrome" not in window_title.lower():
            pyautogui.click()
            sleep(1)
            focused_window = gw.getActiveWindow()
            if focused_window is not None: window_title = str(focused_window.title).strip()
            messagebox.showinfo("Chuyển Tab", "Vui Lòng Chuyển Tab Qua Chrome!")
            sleep(2)
        valid = False
        parts = selected_search.get().split(" ")
        start_date = datetime.strptime(parts[5], "%d/%m/%Y")
        end_date = datetime.strptime(parts[-1], "%d/%m/%Y")
        hoa_don_type = parts[2]
        if hoa_don_type=="mua": 
            df, hanghoa_df, invalid_hoa_don_list, invalid_dich_vu_list = mua_vao(start_date, end_date, search_text)
            if df is not None:
                valid = True
                mua_vao_df = pd.concat([mua_vao_df, df], ignore_index=True)
                mua_vao_hanghoa_df = pd.concat([mua_vao_hanghoa_df, hanghoa_df], ignore_index=True)
                
                output_string = f"{df.shape[0]} hoá đơn mua vào"
                if len(invalid_hoa_don_list)!=0:
                    output_string = output_string + "\n" + "\n".join(invalid_hoa_don_list)
                if len(invalid_dich_vu_list)!=0:
                    output_string = output_string + "\n" + "\n".join(invalid_dich_vu_list)
                    
                res_label.config(text=output_string)
            else: res_label.config(text="0 hoá đơn mua vào")
        elif hoa_don_type=="bán": 
            df, hanghoa_df, invalid_hoa_don_list, invalid_dich_vu_list = ban_ra(start_date, end_date, search_text)
            if df is not None:
                valid = True
                ban_ra_df = pd.concat([ban_ra_df, df], ignore_index=True)
                ban_ra_hanghoa_df = pd.concat([ban_ra_hanghoa_df, hanghoa_df], ignore_index=True)
                
                output_string = f"{df.shape[0]} hoá đơn bán ra"
                if len(invalid_hoa_don_list)!=0:
                    output_string = output_string + "\n" + "\n".join(invalid_hoa_don_list)
                if len(invalid_dich_vu_list)!=0:
                    output_string = output_string + "\n" + "\n".join(invalid_dich_vu_list)
                
                res_label.config(text=output_string)
            else: res_label.config(text="0 hoá đơn bán ra")
        else: 
            muavao_num = 0
            banra_num = 0
            df, hanghoa_df, mua_invalid_hoa_don_list, mua_invalid_dich_vu_list = mua_vao(start_date, end_date, search_text)
            if df is not None:
                valid = True
                mua_vao_df = pd.concat([mua_vao_df, df], ignore_index=True)
                mua_vao_hanghoa_df = pd.concat([mua_vao_hanghoa_df, hanghoa_df], ignore_index=True)
                muavao_num = df.shape[0]
                
                mua_invalid_string = ""
                if len(mua_invalid_hoa_don_list)!=0:
                    mua_invalid_string = mua_invalid_string + "\n" + "\n".join(mua_invalid_hoa_don_list)
                if len(mua_invalid_dich_vu_list)!=0:
                    mua_invalid_string = mua_invalid_string + "\n" + "\n".join(mua_invalid_dich_vu_list)
                
            df, hanghoa_df, ban_invalid_hoa_don_list, ban_invalid_dich_vu_list = ban_ra(start_date, end_date, search_text)
            if df is not None:
                valid = True
                ban_ra_df = pd.concat([ban_ra_df, df], ignore_index=True)
                ban_ra_hanghoa_df = pd.concat([ban_ra_hanghoa_df, hanghoa_df], ignore_index=True)
                banra_num = df.shape[0]
                
                ban_invalid_string = ""
                if len(ban_invalid_hoa_don_list)!=0:
                    ban_invalid_string = ban_invalid_string + "\n" + "\n".join(ban_invalid_hoa_don_list)
                if len(ban_invalid_dich_vu_list)!=0:
                    ban_invalid_string = ban_invalid_string + "\n" + "\n".join(ban_invalid_dich_vu_list)
                    
            output_string = f"{muavao_num} hoá đơn mua vào\n{banra_num} hoá đơn bán ra"
            if len(mua_invalid_string)!=0:
                output_string = output_string + "\n" + mua_invalid_string
            if len(ban_invalid_string)!=0:
                output_string = output_string + "\n" + ban_invalid_string
                
            res_label.config(text=output_string)

        if valid and search_text not in history:
            history.insert(0,search_text)
            nonlocal dropdown
            dropdown.destroy()
            dropdown = tk.OptionMenu(submit_frame, selected_search, *history, command=lambda event: dropdown_func())
            dropdown.config(font=("Arial", 25), bg='yellow', fg='black')
            dropdown.pack(side='left')
        root.deiconify()
        messagebox.showwarning("Hoàn Thành", "Hoàn Thành")
        
        
    def removeRange(start:datetime, end:datetime, havingSet:set):
        remove_set = set()
        cur_date = start
        while cur_date<=end:
            remove_set.add(cur_date.date())
            cur_date+=timedelta(days=1)
        return havingSet - remove_set

    def remove_func():
        search_text = selected_search.get()
        if not search_text.strip(): return # empty input
        if search_text not in history: return # search not submmited yet
        response = messagebox.askyesno("Xóa Dữ Liệu", "Xác nhận xoá?")
        if response:  # If the user clicked "Yes"
            nonlocal ban_ra_df
            nonlocal ban_ra_hanghoa_df
            nonlocal mua_vao_df
            nonlocal mua_vao_hanghoa_df
            nonlocal mua_vao_date_set
            nonlocal ban_ra_date_set
            counter_number = scrape_dic[search_text]
            parts = selected_search.get().split(" ")
            start_date = datetime.strptime(parts[5], "%d/%m/%Y")
            end_date = datetime.strptime(parts[-1], "%d/%m/%Y")
            hoadon_type = parts[2]
            if hoadon_type=="bán":
                ban_ra_df = ban_ra_df[ban_ra_df["counter"]!=counter_number]
                ban_ra_hanghoa_df = ban_ra_hanghoa_df[ban_ra_hanghoa_df["counter"]!=counter_number]
                ban_ra_date_set = removeRange(start_date, end_date, ban_ra_date_set)
            elif hoadon_type=="mua":
                mua_vao_df = mua_vao_df[mua_vao_df["counter"]!=counter_number]
                mua_vao_hanghoa_df = mua_vao_hanghoa_df[mua_vao_hanghoa_df["counter"]!=counter_number]
                mua_vao_date_set = removeRange(start_date, end_date, mua_vao_date_set)
            else:
                ban_ra_df = ban_ra_df[ban_ra_df["counter"]!=counter_number]
                ban_ra_hanghoa_df = ban_ra_hanghoa_df[ban_ra_hanghoa_df["counter"]!=counter_number]
                mua_vao_df = mua_vao_df[mua_vao_df["counter"]!=counter_number]
                mua_vao_hanghoa_df = mua_vao_hanghoa_df[mua_vao_hanghoa_df["counter"]!=counter_number]
                ban_ra_date_set = removeRange(start_date, end_date, ban_ra_date_set)
                mua_vao_date_set = removeRange(start_date, end_date, mua_vao_date_set)
            history.remove(search_text)
            selected_search.set(history[-1])
            res_label.config(text=" ")
            nonlocal dropdown
            dropdown.destroy()
            dropdown = tk.OptionMenu(submit_frame, selected_search, *history, command=lambda event: dropdown_func())
            dropdown.config(font=("Arial", 25), bg='yellow', fg='black')
            dropdown.pack(side='left')

    def format_calender(calender_input:str):
        dtobj = datetime.strptime(calender_input,"%m/%d/%y")
        return dtobj.strftime("%d/%m/%Y")

    # Submit button
    submit_button = tk.Button(submit_frame, text="Submit", pady=20,font=("Arial", 15), command= lambda: submit_func())
    submit_button.pack(pady=20, side='left')
    # Remove button
    remove_button = tk.Button(submit_frame, text="Xoá", pady=20,font=("Arial", 15), command= lambda: remove_func())
    remove_button.pack(pady=20, side='left', padx=10)

    search_text = f"Hoá đơn {hoa_don} từ {format_calender(cal1.get_date())} đến {format_calender(cal2.get_date())}"

    # result label
    res_label = tk.Label(submit_frame, text=" ", font=("Arial", 15), bg='yellow', fg='black', pady=20, padx=20)
    res_label.pack(side="left", padx=10)

    def dropdown_func():
        """
        sync the search bar with the calendar
        """
        if not selected_search.get().strip(): return # empty input
        parts = selected_search.get().split(" ")
        dropdown_start_date, dropdown_end_date = parts[5], parts[-1]
        calendar_start_date = format_calender(cal1.get_date())
        calendar_end_date = format_calender(cal2.get_date())
        if dropdown_start_date!=calendar_start_date or dropdown_end_date!=calendar_end_date:
            correct_start_date = datetime.strptime(dropdown_start_date, "%d/%m/%Y")
            cal1.selection_set(correct_start_date)
            correct_end_date = datetime.strptime(dropdown_end_date, "%d/%m/%Y")
            cal2.selection_set(correct_end_date)
            cal2.config(mindate=correct_start_date)

    # Dropdown
    selected_search = tk.StringVar()
    selected_search.set(search_text)
    dropdown = tk.OptionMenu(submit_frame, selected_search, *history, command=lambda event: dropdown_func())
    dropdown.config(font=("Arial", 25), bg='yellow', fg='black')
    dropdown.pack(side='left')

    def export_func():
        nonlocal mua_vao_df
        nonlocal ban_ra_df
        nonlocal using_username
        nonlocal companyDic
        if ban_ra_df.empty and  mua_vao_df.empty: return
        download_path = filedialog.askdirectory()
        if download_path: 
            output_path = f"{download_path}/hoa_don_dien_tu_{companyDic[using_username]}_{datetime.now().strftime('%d_%m_%Y_%H_%M_%S')}.xlsx"
            with pd.ExcelWriter(output_path) as writer:
                if not ban_ra_df.empty: 
                    copydf = ban_ra_df.copy()
                    copydf.drop(columns=['counter'], inplace=True)
                    copydf.columns = [ 'Ký hiệu mẫu số', 'Ký hiệu hóa đơn', 'Số hóa đơn', 'kyhieu_sohoadon_mst', 'Ngày Lập', 'MCCQT',
                                      'Mã Số Thuế Người Mua', 'Tên Người Mua', 'Địa Chỉ Người Mua', 
                                      'Mã Số Thuế Người Bán', 'Tên Người Bán', 'Địa Chỉ Người Bán',
                                      'Tổng Tiền Chưa Thuế', 'Tổng Tiền Thuế', 'Tổng Tiền Chiết Khấu Thương Mại', 'Tổng Tiền Phí', 'Tổng Tiền Thanh Toán',
                                      'Đơn Vị Tiền Tệ', 'Hình Thức Thanh Toán', 'Thuế Suất', 'Tổng Tiền Bằng Số', 'Tồng Tiền Bằng Chữ',
                                      'Trạng Thái Hoá Đơn', 'Kết Quả Kiểm Tra Hoá Đơn'
                                      ]
                    copydf.to_excel(writer,sheet_name="Hoá Đơn Điện Tử Bán Ra", index=False)

                    copydf = ban_ra_hanghoa_df.copy()
                    copydf.drop(columns=['counter'], inplace=True)
                    copydf.columns = [
                        'kyhieu_sohoadon_mst', 'Tính Chất', 'Tên Hàng Hoá, Dịch Vụ', 'Đơn Vị Tính', 'Số Lượng', 'Đơn Giá', 'Chiết Khấu', 'Thuế Xuất', 'Thành tiền chưa có thuế GTGT', 'Thuế GTGT', 'Thành tiền', 'Tên Người Mua', 'MST Người Mua', 'Trạng Thái Hoá Đơn'
                    ]
                    copydf.to_excel(writer,sheet_name="Hàng Hoá, Dịch Vụ Bán Ra", index=False)
                if not mua_vao_df.empty: 
                    copydf = mua_vao_df.copy()
                    copydf.drop(columns=['counter'], inplace=True)
                    copydf.columns = [ 'Ký hiệu mẫu số', 'Ký hiệu hóa đơn', 'Số hóa đơn', 'kyhieu_sohoadon_mst', 'Ngày Lập', 'MCCQT',
                                      'Mã Số Thuế Người Mua', 'Tên Người Mua', 'Địa Chỉ Người Mua', 
                                      'Mã Số Thuế Người Bán', 'Tên Người Bán', 'Địa Chỉ Người Bán',
                                      'Tổng Tiền Chưa Thuế', 'Tổng Tiền Thuế', 'Tổng Tiền Chiết Khấu Thương Mại', 'Tổng Tiền Phí', 'Tổng Tiền Thanh Toán',
                                      'Đơn Vị Tiền Tệ', 'Hình Thức Thanh Toán', 'Thuế Suất', 'Tổng Tiền Bằng Số', 'Tồng Tiền Bằng Chữ',
                                      'Trạng Thái Hoá Đơn', 'Kết Quả Kiểm Tra Hoá Đơn'
                                      ]
                    copydf.to_excel(writer,sheet_name="Hoá Đơn Điện Tử Mua Vào", index=False)

                    copydf = mua_vao_hanghoa_df.copy()
                    copydf.drop(columns=['counter'], inplace=True)
                    copydf.columns = [
                        'kyhieu_sohoadon_mst', 'Tính Chất', 'Tên Hàng Hoá, Dịch Vụ', 'Đơn Vị Tính', 'Số Lượng', 'Đơn Giá', 'Chiết Khấu', 'Thuế Xuất', 'Thành tiền chưa có thuế GTGT', 'Thuế GTGT', 'Thành tiền', 'Tên Người Bán', 'MST Người Bán', 'Trạng Thái Hoá Đơn'
                    ]
                    copydf.to_excel(writer,sheet_name="Hàng Hoá, Dịch Vụ Mua Vào", index=False)
                    messagebox.showinfo("Xuất File Thành Công", "Xuất File Thành Công!")

    # Export Button
    export_button = tk.Button(export_frame, text="Xuất file", pady=20, font=("Arial", 15), command= lambda: export_func())
    export_button.pack(side="left", padx=(0,20))
    safemode = tk.BooleanVar()
    checkbox = tk.Checkbutton(export_frame, text="An Toàn", variable=safemode, font=("Arial", 15))
    checkbox.pack(side="left")
    
    def update_search_bar():
        search_text = f"Hoá đơn {hoa_don} từ {format_calender(cal1.get_date())} đến {format_calender(cal2.get_date())}"
        selected_search.set(search_text)
        res_label.config(text=" ")

    def update_cal2_mindate():
        min = datetime.strptime(cal1.get_date(), "%m/%d/%y")
        cal2_date = datetime.strptime(cal2.get_date(), "%m/%d/%y")
        if cal2_date < min: cal2.selection_set(min)

    def cal1_func():
        update_cal2_mindate()
        update_search_bar()
        
    def update_cal1_maxdate():
        max = datetime.strptime(cal2.get_date(), "%m/%d/%y")
        cal1_date = datetime.strptime(cal1.get_date(), "%m/%d/%y")
        if cal1_date > max: cal1.selection_set(max)
        
    def cal2_func():
        update_cal1_maxdate()
        update_search_bar()

    cal1.bind("<<CalendarSelected>>", lambda event: cal1_func())
    cal2.bind("<<CalendarSelected>>", lambda event: cal2_func())

    try:
        root.mainloop()
    finally:
        driver.close()
        user_dir = os.path.join(exe_path, "users")
        os.makedirs(user_dir, exist_ok=True)
        with open(f"{exe_path}/users/usernameDic.pkl", "wb") as file:
            pickle.dump(usernameDic, file)
        with open(f"{exe_path}/users/pwDic.pkl", "wb") as file:
            pickle.dump(pwDic, file)
        with open(f"{exe_path}/users/companyDic.pkl", "wb") as file:
            pickle.dump(companyDic, file)
            
if __name__=="__main__":
    import multiprocessing
    multiprocessing.freeze_support()
    main()

