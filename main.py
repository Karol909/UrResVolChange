import urReceive
import pandas as pd
import time
import math
import random
import openpyxl
from datetime import datetime
import urDashboard
import excelLog
import cv2 as cv
import os


urDashboard = urDashboard.URDashboard('192.168.0.64', 29999)
urReceive = urReceive.URReceive('192.168.0.100', 5555)

urDashboard.load("20251030_researchplus_vol")
urDashboard.play()

excelLog = excelLog.ExcelLog()
data_number = 0
data_string = None
connected = False
volume_number = 0
max_volume_number = 300
running = True
pic_taken = False

table_row = [0,0,0,0,0,0,0,0,0,0]

pictures_folder = "Pics_Q37764J_1000uL"
os.makedirs(pictures_folder, exist_ok=True)

cam = cv.VideoCapture(1, cv.CAP_DSHOW)
cam.set(cv.CAP_PROP_AUTOFOCUS, 0)
cam.set(cv.CAP_PROP_FOCUS, 100)


while running:
    if not connected:
        connected = urReceive.listen()
    else:
        data_string = urReceive.get_data_UR()
        if data_string: 
            for line in data_string.strip().splitlines():
                parts = line.strip().split()
                if len(parts) == 2:
                    try:
                        data_number += 1
                        data_type = int(parts[0])
                        data = float(parts[1])
                        print("Data to excel: ", data_type, data)
                        excelLog.save(data_number, data_type, data)

                        if data_type <= len(table_row):
                            table_row[data_type-1] = data
                            
                        

                        if data_type == 99:
                            table_row[-1] = data
                            volume_number += 1
                            excelLog.save_row(table_row)
                            print(volume_number)
                        if data_type == 999:
                            if data == 0:
                                urDashboard.stop()
                                connected = False
                                pic_taken = False
                                if volume_number >= max_volume_number:
                                    excelLog.unload()
                                    urDashboard.stop()
                                    urDashboard.disconnect()
                                    urReceive.disconnect()
                                    running = False
                                    break
                                else:
                                    time.sleep(5)
                                    urDashboard.play()
                            elif data == 1:
                                for _ in range(30):
                                    cam.read()
                                    time.sleep(0.03)
                                while True:
                                    ret, frame = cam.read()
                                    print("Photo")
                                    if ret and not pic_taken:
                                        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                                        filename = os.path.join(pictures_folder, f"photo_{table_row[0]}_{timestamp}.jpg")
                                        cv.imwrite(filename, frame)
                                        pic_taken = True
                                        break
                            
                    except ValueError:
                        print("Error")
                        pass
            

cam.release()
cv.destroyAllWindows()