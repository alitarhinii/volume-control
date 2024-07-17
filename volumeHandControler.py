import cv2 as cv
import mediapipe as mp
from cvzone import HandTrackingModule as htm
import numpy as np
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from comtypes import CLSCTX_ALL

wCam, hCam = 2000, 2000
cap = cv.VideoCapture(0)
cap.set(3, wCam)
cap.set(4, hCam)

detector = htm.HandDetector(detectionCon=0.7)
devices = AudioUtilities.GetSpeakers()
# This line uses the AudioUtilities class from the pycaw library to get a reference to
# the audio endpoint device representing the speakers
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
#  Here, we activate an interface on the IMMDevice object using the Activate method. This method takes
# three arguments:
# IAudioEndpointVolume._iid_: This is a globally unique identifier (GUID) representing the interface 
# of the object we want to activate. In this case, it's the interface for controlling the audio
# endpoint volume.
# CLSCTX_ALL: This is a constant defined in the comtypes module that specifies the execution context 
# for the object. CLSCTX_ALL indicates that the object can be executed in any context.
# None: This parameter is reserved for future use and should be set to None.
volume = interface.QueryInterface(IAudioEndpointVolume)
# This interface allows us to control the volume of the audio endpoint device.
volumeRange = volume.GetVolumeRange()

minVol = volumeRange[0]
maxVol = volumeRange[1]

while True:
    suc, img = cap.read()
    hands, img = detector.findHands(img)
    if hands:
        for hand in hands:
            lmList = hand["lmList"]
            if len(lmList) != 0:
                x1, y1 = lmList[4][1], lmList[4][2]
                x2, y2 = lmList[8][1], lmList[8][2]
                cx, cy = (x1 + x2) // 2, (y1 + y2) // 2
                cv.circle(img, (x1, y1), 15, (0, 0, 255), cv.FILLED)
                cv.circle(img, (x2, y2), 15, (0, 0, 255), cv.FILLED)
                cv.circle(img, (cx, cy), 15, (255, 255, 0), cv.FILLED)
                cv.line(img, (x1, y1), (x2, y2), (255, 0, 255), thickness=1, lineType=cv.LINE_AA)
                length = np.hypot(x2 - x1, y2 - y1)
                vol = np.interp(length, [50, 440], [minVol, maxVol])
                volBar = np.interp(length, [50, 440], [400, 150])
                percentage = np.interp(length, [50, 440], [0, 100])
                volume.SetMasterVolumeLevel(vol, None)
                if length < 50:
                    cv.circle(img, (cx, cy), 15, (255, 255, 255), cv.FILLED)
                    
                cv.rectangle(img, (50, int(volBar)), (85, 400), (255, 88, 99), cv.FILLED)
                cv.rectangle(img, (50, 150), (85, 400), (255, 88, 99), 3)
                cv.putText(img, f'Volume {int(percentage)}%', (50, 450), cv.FONT_HERSHEY_PLAIN, 1.5, (0, 255, 0), 3)

    cv.imshow("image", img)
    if cv.waitKey(1) & 0xFF == ord('q'):
        break

cap.release()
cv.destroyAllWindows()
