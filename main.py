from pathlib import Path
from datetime import datetime
import requests
import pprint
import xlwings as xw
import wmi
import math
import pyttsx3

conn = wmi.WMI()



 
url = "https://xltraders.online/api/v1/validate"
 
 
def Text2Speech(Text):
    #print(Text)
    
    global engine
    try:
        engine = pyttsx3.init()
        voices = engine.getProperty('voices')
        try:
           engine.setProperty('voice', voices[0].id)
        except Exception as e:
            pass
        engine.setProperty('rate', 130)
        engine.say(Text)
        engine.runAndWait()
        del engine
    except Exception as e:
        #print(f"Issue in voice module : {e}")
        pass
    
# Text2Speech("Welcome, this is a check 1, 2, 3, 4, 5")
 

# file_path = Path(__file__).parent / 'hello.txt'

def sysinfo():
    sysdrive = conn.Win32_OperatingSystem()[0].SystemDrive
    freespace = next((f"{100 * int(i.FreeSpace) / int(i.Size):.2f}" for i in conn.Win32_LogicalDisk() if i.DeviceID == sysdrive), None)
    
    osdrivefreecheck = None
    
    try:
        osdrivefreecheck = float(freespace)>50.00
    except:
        osdrivefreecheck = None
        
    return {
            "sysID": conn.Win32_ComputerSystemProduct()[0].UUID,
            "vendor" : conn.Win32_ComputerSystemProduct()[0].vendor,
            "proc" : conn.Win32_Processor()[0].Name,
            "tmem" : ("{:.2f}".format(math.ceil(int(conn.Win32_Computersystem()[0].TotalPhysicalMemory)/1024/1024/1024))),
            "username" : conn.Win32_Computersystem()[0].UserName,
            "partofdomain" : conn.Win32_Computersystem()[0].PartOfDomain,
            "workgroup" : conn.Win32_Computersystem()[0].Workgroup,
            "freemem" : ("{:.2f}".format(int(conn.Win32_OperatingSystem()[0].FreePhysicalMemory)/1024/1024)),
            "osarch" : conn.Win32_OperatingSystem()[0].OSArchitecture,
            "sysdrive" : sysdrive,
            "serialnum" : conn.Win32_Baseboard()[0].SerialNumber,
            "manufacturer" : conn.Win32_Baseboard()[0].Manufacturer,
            "product" : conn.Win32_Baseboard()[0].Product,
            "macs" : [mac.MACAddress for mac in conn.Win32_NetworkAdapterConfiguration() if mac.MACAddress is not None],
            "freespace" : f'{freespace}%',
            "osdrivefreecheck" : osdrivefreecheck,
            "os": conn.Win32_OperatingSystem()[0].Caption
            }
    
import threading
import time

def writevalues(sheet, os, osd, mem, proc):
    sheet.range('B13').value = os
    sheet.range('B14').value = osd
    sheet.range('B15').value = mem
    sheet.range('B16').value = proc
    # time.sleep(1.5)
    
def readmsg(osmsg, osd, mem):
    Text2Speech(osmsg)
    Text2Speech(osd)
    Text2Speech(mem)
    
def syscheck():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    
    data = sysinfo()
    
    os = data["os"]
    proc = data["proc"]
    osmsg = f'Operating System : {os}'
    mem = f'Total Memory : {data["tmem"]} GB, Free Memory {data["freemem"]} GB'
    osd = f'OS Drive is {data["sysdrive"]}, {data["freespace"]} free'

    t1 = threading.Thread(target=writevalues(sheet, os, osd, mem, proc))
    t2 = threading.Thread(target=readmsg(osmsg, osd, mem))
    t1.start()
    t2.start()
    # if (writevalues(sheet, os, osd, mem, proc)):
    #     for m in [osd, mem, osmsg]:
    #         Text2Speech(m)
    # else:
    #     pass
    

msgs = []

def main():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    sheet.range('E4').clear_contents()
    sheet.range('E5').clear_contents()
    sheet.range('E5').clear()
    sheet.range('A1').clear_contents()
    
    if sheet.range('B5').value == None:
        sheet.range('C5').value = "Please input user id"
        sheet.range('C5').color = '#F97068'
        sheet.range('C5').font.color = '#140201'
        msgs.append("Please input user id")
    else:
        sheet.range('C5').value = "OK"
        sheet.range('C5').color = '#6BF178'

    if sheet.range('B6').value == None:
        sheet.range('C6').value = 'No activation key found, subscribe any plan first to get activation key'
        sheet.range('C6').color = '#F97068'
        sheet.range('C6').font.color = '#140201'
        msgs.append('No activation key found, subscribe any plan first to get activation key')
    else:
        sheet.range('C6').value = "OK"
        sheet.range('C6').color = '#6BF178'

    if sheet.range('B5').value != None and sheet.range('B6').value != None:
        userid = sheet.range('B5').value
        token = sheet.range('B6').value
        data = {"userid": userid, "token": token }
        response = requests.post(url, json=data)

        if response.status_code == 200:
            Text2Speech("All data is correct, you are good to go")
            date_string = response.json()["validity"]
            date_format = '%m-%d-%Y'
            desired_date = datetime.strptime(date_string, date_format)
            current_date = datetime.now()
            time_difference = desired_date - current_date
            days_difference = time_difference.days
            Text2Speech(f'Your license validity is up to {desired_date}')
            sheet.range("B9").value = days_difference
            sheet.range("B8").value = date_string
            sheet.range('B8').color = '#FFFF82'
            sheet.range("E5").value = ""
        else:
            sheet.range('E4').value = "Error Message"
            sheet.range("E5").value = response.json()["Value"]
            sheet.range('E5').color = '#F97068'
            sheet.range('E5').font.color = '#140201'
            Text2Speech(response.json()["Value"])
            sheet.range("B9").value = ""
            sheet.range("B8").value = ""
    else:
        sheet.range('A1').value = "Correct the errors please"
        sheet.range('A1').color = '#F97068'
        sheet.range('A1').font.color = '#140201'
        msgs.append("Correct the errors please")
        for m in msgs:
            Text2Speech(m)
        
# def final_data():
#     data = sysinfo()
#     data["userid"] = userid
#     data["token"] = token
#     return data

# pprint.pprint(final_data())


@xw.func
def hello(name):
    return f"Hello {name}!"


if __name__ == "__main__":
    xw.Book("main.xlsm").set_mock_caller()
    main()
