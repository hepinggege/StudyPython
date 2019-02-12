import cv2
import numpy as np

cap = cv2.VideoCapture(r'C:\Users\PC\Desktop\IMG_9652_Trim.MOV')

def filter_out_blue(src_frame):
    if src_frame is not None:
        hsv = cv2.cvtColor(src_frame, cv2.COLOR_BGR2HSV)
        lower_blue = np.array([78,43,46])
        upper_blue = np.array([110,255,255])
        # inRange()方法返回的矩阵只包含0,255 (CV_8U) 0表示不在区间内
        mask = cv2.inRange(hsv, lower_blue, upper_blue)
        return cv2.bitwise_and(src_frame, src_frame, mask=mask)


while(cap.isOpened()):
    ret,frame = cap.read()
    if ret:
        res = filter_out_blue(frame)
        cv2.imshow("result",res)
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break
    else:
        break

cap.release()
cv2.destroyAllWindows()
