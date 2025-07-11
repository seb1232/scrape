import pyautogui
from pynput import mouse

def on_click(x, y, button, pressed):
    if pressed:
        print(f"Mouse clicked at ({x}, {y})")

with mouse.Listener(on_click=on_click) as listener:
    print("Click anywhere to print the (x, y) coordinates. Press Ctrl+C to exit.")
    listener.join()