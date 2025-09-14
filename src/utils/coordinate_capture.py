import pyautogui
import time

def capture_coordinates():
    """Capture coordinates by clicking"""
    coordinates = []
    labels = ["Port dropdown", "COM3 option", "Baud Rate dropdown", "115200 option", "Open button"]
    
    print("Click on each button when prompted. You have 5 seconds between each click.")
    
    for i, label in enumerate(labels):
        print(f"\n{i+1}. Position mouse over '{label}' and wait...")
        time.sleep(5)
        x, y = pyautogui.position()
        coordinates.append((x, y))
        print(f"Captured: {label} at ({x}, {y})")
    
    print("\nCaptured coordinates:")
    for i, (label, (x, y)) in enumerate(zip(labels, coordinates)):
        print(f"{label}: x={x}, y={y}")
    
    return coordinates

if __name__ == "__main__":
    capture_coordinates()