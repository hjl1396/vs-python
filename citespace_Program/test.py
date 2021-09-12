import pyautogui
import pyperclip  
import time

matrix1=[[825,527],[708,599]]
matrix2=[[945,598],[1078,596]]
matrix3=[[1358,780],[1072,964],[1076,1058]]
matrix4=[[1332,646],[1465,646]]



time.sleep(10)

num=1740

for k in range(17):

    for i in range(len(matrix1)):
        pyautogui.click(x=matrix1[i][0],y=matrix1[i][1],button='left',duration=0.25)
        time.sleep(1)
    i=0
    for i in range(len(matrix2)):
        pyautogui.click(x=matrix2[i][0],y=matrix2[i][1],button='left',duration=0.25)
        pyautogui.typewrite(['backspace','backspace','backspace','backspace','backspace','backspace','backspace'])
        pyperclip.copy(num+k*500+i*499)  # 先复制
        pyautogui.hotkey('ctrl','v')# 再粘贴
        #pyautogui.typewrite(str(num+k*500+i*499))
        time.sleep(1)
    pyautogui.click(x=matrix4[i][0],y=matrix4[i][1],button='left',duration=0.25)
    i=0
    for i in range(len(matrix3)):
        pyautogui.click(x=matrix3[i][0],y=matrix3[i][1],button='left',duration=0.25)
        time.sleep(1)
    time.sleep(20)