import cv2
import os
import pandas as pd
import numpy as np
from ultralytics import YOLO
from tracker import *
import time
import openpyxl
# from openpyxl.drawing.image import Image
from math import dist
model = YOLO('yolov8s.pt')
countc = 0

def RGB(event, x, y, flags, param):
    if event == cv2.EVENT_MOUSEMOVE:
        colorsBGR = [x, y]
        # print(colorsBGR)


def excel_data(filename_ex="excel_data.xlsx", sheetname="vehicle_info", data=None):
    try:
        # Load the workbook
        # workbook = openpyxl.load_workbook(filename_ex)

        # # Get the worksheet by name
        # worksheet = workbook[sheetname]

        # Find the next available row
        next_row = worksheet.max_row + 1
        # worksheet.row_dimensions[next_row].height = 200

        # Insert serial number in the first column
        worksheet.cell(row=next_row, column=1, value=next_row - 1)

        # Insert data in subsequent columns
        for col, value in enumerate(data, start=2):
            worksheet.cell(row=next_row, column=col, value=value)

        # Save the changes
        print("New row added successfully.")

    except FileNotFoundError:
        print(f"File '{filename_ex}' not found.")


with open('stored.txt', 'w') as file:
    pass


def append_text(filename, text, speed):
    with open(filename, 'a') as file:
        if speed > 30:
            file.write(text + ' (overspeeding)' + '\n')
            vio = "overspeeding"
        elif speed < 20:
            file.write(text + ' (underspeeding)' + '\n')
            vio = "underspeeding"
        else:
            file.write(text + '\n')
            vio = "no violation"
    return vio


filename_ex = "excel_data.xlsx"
sheetname = "vehicle_info"
workbook = openpyxl.load_workbook(filename_ex)
worksheet = workbook[sheetname]
Headers = ['#', 'Direction', 'Sl_no.', 'Speed', 'Violation']

try:

    # Clear all cells in the worksheet
    for row in worksheet.iter_rows():
        for cell in row:
            cell.value = None
            workbook.save(filename_ex)
            workbook.close()
    # workbook.close()
    for col, value in enumerate(Headers, start=1):
        worksheet.cell(row=1, column=col, value=value)

    # Save the changes

    workbook.save(filename_ex)
    workbook.close()
except FileNotFoundError:
    pass


cv2.namedWindow('RGB')
cv2.setMouseCallback('RGB', RGB)

cap = cv2.VideoCapture('veh2.mp4')


my_file = open("coco.txt", "r")
data = my_file.read()
class_list = data.split("\n")
# print(class_list)

output_dir = "violator_vehicles"
for item in os.listdir(output_dir):
    item_path = os.path.join(output_dir, item)
    if os.path.isfile(item_path):
        # Delete files
        os.remove(item_path)
os.makedirs(output_dir, exist_ok=True)

count = 0

tracker = Tracker()

cy1 = 322
cy2 = 368

offset = 6

vh_down = {}
counter = []


vh_up = {}
counter1 = []

while True:
    ret, frame = cap.read()
    if not ret:
        break
    count += 1
    if count % 3 != 0:
        continue
    frame = cv2.resize(frame, (1020, 500))

    results = model.predict(frame)
 #   print(results)
    a = results[0].boxes.data
    px = pd.DataFrame(a).astype("float")
#    print(px)
    list = []

    for index, row in px.iterrows():
        #        print(row)

        x1 = int(row[0])
        y1 = int(row[1])
        x2 = int(row[2])
        y2 = int(row[3])
        d = int(row[5])
        c = class_list[d]
        if 'car' in c:
            list.append([x1, y1, x2, y2])
    bbox_id = tracker.update(list)
    for bbox in bbox_id:
        x3, y3, x4, y4, id = bbox
        cx = int(x3+x4)//2
        cy = int(y3+y4)//2

        cv2.rectangle(frame, (x3, y3), (x4, y4), (0, 0, 255), 2)

        if cy1 < (cy+offset) and cy1 > (cy-offset):
            vh_down[id] = time.time()
        if id in vh_down:

            if cy2 < (cy+offset) and cy2 > (cy-offset):
                elapsed_time = time.time() - vh_down[id]
                if counter.count(id) == 0:
                    counter.append(id)
                    distance = 10  # meters
                    a_speed_ms = distance / elapsed_time
                    a_speed_kh = a_speed_ms * 3.6
                    cv2.circle(frame, (cx, cy), 4, (0, 0, 255), -1)
                    cv2.putText(frame, str(id), (x3, y3),
                                cv2.FONT_HERSHEY_COMPLEX, 0.6, (255, 255, 255), 1)
                    cv2.putText(frame, str(int(a_speed_kh))+'Km/h', (x4, y4),
                                cv2.FONT_HERSHEY_COMPLEX, 0.8, (0, 255, 255), 2)
                    image_name = os.path.join(
                        output_dir, f"snapshot_{countc}.jpg")
                    cv2.imwrite(image_name, frame)
                    # print(f"Snapshot saved: {image_name}")
                    countc += 1
                    text = f"goingdown : {len(counter)} Speed : {a_speed_kh} km/hr"
                    vio = append_text('stored.txt', text, a_speed_kh)
                    data = ['Going Down', len(counter), str(
                        round(a_speed_kh, 2))+'km/hr', vio]
                    excel_data(data=data)

        ##### going UP#####
        if cy2 < (cy+offset) and cy2 > (cy-offset):
            vh_up[id] = time.time()
        if id in vh_up:

            if cy1 < (cy+offset) and cy1 > (cy-offset):
                elapsed1_time = time.time() - vh_up[id]

                if counter1.count(id) == 0:
                    counter1.append(id)
                    distance1 = 10  # meters
                    a_speed_ms1 = distance1 / elapsed1_time
                    a_speed_kh1 = a_speed_ms1 * 3.6
                    cv2.circle(frame, (cx, cy), 4, (0, 0, 255), -1)
                    cv2.putText(frame, str(id), (x3, y3),
                                cv2.FONT_HERSHEY_COMPLEX, 0.6, (255, 255, 255), 1)
                    cv2.putText(frame, str(int(a_speed_kh1))+'Km/h', (x4, y4),
                                cv2.FONT_HERSHEY_COMPLEX, 0.8, (0, 255, 255), 2)
                    image_name = os.path.join(
                        output_dir, f"snapshot_{countc}.jpg")
                    cv2.imwrite(image_name, frame)
                    print(f"Snapshot saved: {image_name}")
                    countc += 1
                    text = f"goingup   : {len(counter1)} Speed : {a_speed_kh1} km/hr"
                    vio = append_text('stored.txt', text, a_speed_kh1)
                    data = ['Going Up', len(counter1), str(
                        round(a_speed_kh1, 2)) + "km/hr", vio]
                    excel_data(data=data)

    cv2.line(frame, (274, cy1), (814, cy1), (255, 255, 255), 1)

    cv2.putText(frame, ('L1'), (277, 320),
                cv2.FONT_HERSHEY_COMPLEX, 0.8, (0, 255, 255), 2)

    cv2.line(frame, (177, cy2), (927, cy2), (255, 255, 255), 1)

    cv2.putText(frame, ('L2'), (182, 367),
                cv2.FONT_HERSHEY_COMPLEX, 0.8, (0, 255, 255), 2)
    d = (len(counter))
    u = (len(counter1))
    cv2.putText(frame, ('goingdown:-')+str(d), (60, 90),
                cv2.FONT_HERSHEY_COMPLEX, 0.8, (0, 255, 255), 2)

    cv2.putText(frame, ('goingup:-')+str(u), (60, 130),
                cv2.FONT_HERSHEY_COMPLEX, 0.8, (0, 255, 255), 2)
    cv2.imshow("RGB", frame)
    if cv2.waitKey(1) & 0xFF == 27:
        break

workbook.save(filename_ex)
workbook.close()
cap.release()
cv2.destroyAllWindows()
