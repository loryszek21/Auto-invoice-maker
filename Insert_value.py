import pyautogui as gui, time
screenWidth, screenHeight = gui.size();
# gui.moveTo(5,screenHeight-5)
# gui.click()

# gui.typewrite('Notatnik')
# 

# time.sleep(2) #damy 2 sekundy aplikacji notepad na poprawne uruchomienie
# gui.keyDown('winleft')
# gui.press('up')
# gui.keyUp('winleft')



# def identifyloc():
#     while True:
#         currentMouseX, currentMouseY = gui.position()
#         print(currentMouseX,currentMouseY)
#         time.sleep(3)
# identifyloc()
# time.sleep(3)
res = gui.locateOnScreen('Add.png', confidence=0.7)
print(res)
gui.click(res)
gui.typewrite("mare")
time.sleep(1)
gui.press('enter')
gui.press('tab')
gui.typewrite("filet")
time.sleep(1)
for i in range(9):
    gui.press('down')
gui.press('enter')
gui.typewrite("20")
gui.press('tab')
gui.press('right')
gui.typewrite("35")


# for res in gui.locateOnScreen("Add.png"):
#     print(res)
# gui.click(159, 71)

