import pyautogui
import time

def display_mouse_position():
    """Display real-time mouse coordinates"""
    print("Move your mouse to the desired button and note the coordinates.")
    print("Press Ctrl+C to stop.")
    
    try:
        while True:
            x, y = pyautogui.position()
            print(f"Mouse position: x={x}, y={y}", end='\r')
            time.sleep(0.1)
    except KeyboardInterrupt:
        print("\nCoordinate detection stopped.")

if __name__ == "__main__":
    display_mouse_position()