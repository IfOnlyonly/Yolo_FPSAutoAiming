import torch
from mss import mss
from PIL import Image
import win32api, win32con
import math

# Load model
model = torch.hub.load('ultralytics/yolov5', 'yolov5s')
model.cuda()
# Detect only human
model.classes = [0]

# Screen Resolution
x_screen = 2560  # Input your monitor x pixels number
y_screen = 1440  # Input your monitor y pixels number

# Set the region size for detection
x_window = int(x_screen / 8)
y_window = int(y_screen / 4)

# Set the Move factor for crosshair movement (higher value means faster)
MoveFactor = 0.5
# Set the aiming factor, ranging in [0 , 1](1 = top of the detection block, 0 = bottom of the detection block )
AimFactor = 0.80

# Set boundary of the detecting zone
d_left = int(x_screen / 2 - x_window / 2)
d_top = int(y_screen / 2 - y_window / 2)
mon = {'left': d_left, 'top': d_top, 'width': x_window, 'height': y_window}

# Main Loop
with mss() as sct:
    while True:
        d_xc = 5000
        d_yc = 5000
        d_2c = 5000
        screenShot = sct.grab(mon)

        img = Image.frombytes(
            'RGB',
            (screenShot.width, screenShot.height),
            screenShot.rgb,
        )

        # Make detections
        results = model(img)
        # Retrieve result date
        result_panda = results.pandas().xyxy[0]
        # Calculate objects number
        n_object = int(results.pandas().xyxy[0].size / 7)

        # Find the object closet to the crosshair
        for i in range(n_object):
            try:
                d_xmin = result_panda.at[0, "xmin"] + int(x_screen / 2 - x_window / 2)
                d_ymin = result_panda.at[0, "ymin"] + int(y_screen / 2 - y_window / 2)
                d_xmax = result_panda.at[0, "xmax"] + int(x_screen / 2 - x_window / 2)
                d_ymax = result_panda.at[0, "ymax"] + int(y_screen / 2 - y_window / 2)
            except:
                d_xmin = 0
                d_ymin = 0
                d_xmax = 0
                d_ymax = 0
            d_xc_temp = int((d_xmin + d_xmax) / 2 - x_screen / 2)
            d_yc_temp = int((d_ymin * AimFactor + d_ymax * (1 - AimFactor)) - y_screen / 2)
            d_2c_temp = math.sqrt(d_xc_temp * d_xc_temp + d_yc_temp * d_yc_temp)
            if d_2c_temp <= d_2c:
                d_xc = d_xc_temp
                d_yc = d_yc_temp
                d_2c = d_2c_temp

        if d_xc == 5000:
            d_xc = 0
        if d_yc == 5000:
            d_yc = 0

        # Move Crosshair
        if d_xc != 0 and d_yc != 0 and (win32api.GetKeyState(0x10) < -100):
            win32api.mouse_event(win32con.MOUSEEVENTF_MOVE, int(d_xc * MoveFactor), int(d_yc * MoveFactor), 0, 0)
            # print("Move")
