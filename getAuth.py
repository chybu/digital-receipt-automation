import websocket, json, httpx, traceback
from multiprocessing import Process, Queue, Event
from websocket import WebSocketConnectionClosedException
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
from selenium.common.exceptions import NoSuchElementException

def get_websocket_url():
        response = httpx.get("http://localhost:9222/json")
        tabs = response.json()
        # Use the first tab's WebSocket URL (modify if needed)
        return tabs[0]["webSocketDebuggerUrl"]

def getAuthorization(que:Queue, ready_event, error):
    while True:
        try:
            # Connect to the WebSocket URL of the pre-opened Chrome instance
            CHROME_DEVTOOLS_URL =  get_websocket_url()
            # Initialize WebSocket
            ws = websocket.WebSocket()
            ws.connect(CHROME_DEVTOOLS_URL)

            # Enable Network Monitoring
            ws.send(json.dumps({"id": 1, "method": "Network.enable"}))
            
            # Signal that the WebSocket is ready
            ready_event.set()
            # Listen for Network Events
            while not error.is_set():
                response = ws.recv()
                message = json.loads(response)
                if message.get("method") == "Network.requestWillBeSent":
                    request = message["params"]["request"]
                    if message["params"]["type"] == "XHR":
                        if 'Authorization' in request["headers"] and request["headers"]['Authorization']:
                            ws.close()
                            auth = request["headers"]['Authorization']
                            que.put(auth)
                            return
            return
        except (ConnectionError, WebSocketConnectionClosedException): print("cannot connect in getAuth")
                
def xem_hoa_don(ready_event, banra, error, driver_path):
    try:
        # Wait until the WebSocket is ready
        ready_event.wait()
        options = webdriver.ChromeOptions()
        options.debugger_address = "127.0.0.1:9222"  

        service = Service(driver_path)

        driver = webdriver.Chrome(service=service, options=options)
        
        stt_list = driver.find_elements(By.CSS_SELECTOR, "tr.ant-table-row.ant-table-row-level-0 > td:nth-child(1) > span")
        element_list = driver.find_elements(by=By.CSS_SELECTOR, value="tr.ant-table-row.ant-table-row-level-0")
        for element, stt in zip(element_list, stt_list):
            if not stt.text.strip(): continue
            element.click()
            break
        if not banra:
            xem_hoa_don_button = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/section/section/main/div/div/div/div/div[3]/div[2]/div[3]/div[2]/div[1]/div[2]/div/div[5]/button")))

        else:
            # xem_hoa_don_button = driver.find_element(by=By.XPATH, value="/html/body/div[1]/section/section/main/div/div/div/div/div[3]/div[1]/div[3]/div[2]/div[1]/div[2]/div/div[5]/button")
            xem_hoa_don_button = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/section/section/main/div/div/div/div/div[3]/div[1]/div[3]/div[2]/div[1]/div[2]/div/div[5]/button")))
        xem_hoa_don_button.click()
        # wait for loading the receipt
        WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#__next > div.Spin__SpinWrapper-sc-1px1x1e-0.enNzTo > div > div")))
        WebDriverWait(driver, 60).until(EC.invisibility_of_element((By.CSS_SELECTOR, "#__next > div.Spin__SpinWrapper-sc-1px1x1e-0.enNzTo > div > div")))
        sleep(1)
        close_receipt_button = driver.find_element(by=By.CSS_SELECTOR, value="div.ant-modal-content i.anticon.anticon-close.ant-modal-close-icon")
        close_receipt_button.click()
    except Exception as e: 
        traceback.print_exc()
        print("error in xem_hoa_don")
        error.set()
    
def getAuth(banra, driver_path):
    output_que = Queue()
    ready_event = Event()
    error = Event()
    selenium_process = Process(target=xem_hoa_don, args=(ready_event, banra, error, driver_path))
    network_process = Process(target=getAuthorization, args=(output_que,ready_event, error))

    selenium_process.start()
    network_process.start()
    
    selenium_process.join()
    network_process.join()
    
    if error.is_set(): raise NoSuchElementException
    
    return output_que.get()
    
if __name__ == "__main__":
    getAuth()
