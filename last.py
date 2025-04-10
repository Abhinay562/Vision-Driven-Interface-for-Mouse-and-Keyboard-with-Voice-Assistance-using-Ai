import cv2
import numpy as np
import mediapipe as mp
import pyautogui
import math
import time
from pynput.keyboard import Controller, Key
from enum import Enum, auto
import speech_recognition as sr
import queue
import threading
import subprocess
import screen_brightness_control as sbc
import pycaw.pycaw as pycaw
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
import os
import glob

# Initialize mediapipe
mp_hands = mp.solutions.hands
mp_drawing = mp.solutions.drawing_utils

class HandType(Enum):
    LEFT = auto()
    RIGHT = auto()

class SystemController:
    def __init__(self):
        devices = pycaw.AudioUtilities.GetSpeakers()
        interface = devices.Activate(pycaw.IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        self.volume = cast(interface, POINTER(pycaw.IAudioEndpointVolume))
        self.recent_app = None
        
    def set_volume(self, level):
        self.volume.SetMasterVolumeLevelScalar(level / 100, None)
        
    def increase_volume(self, amount=10):
        current = self.volume.GetMasterVolumeLevelScalar() * 100
        self.set_volume(min(100, current + amount))
        print(f"Volume increased to {min(100, current + amount)}%")
        
    def decrease_volume(self, amount=10):
        current = self.volume.GetMasterVolumeLevelScalar() * 100
        self.set_volume(max(0, current - amount))
        print(f"Volume decreased to {max(0, current - amount)}%")
        
    def max_volume(self):
        self.set_volume(100)
        print("Volume set to maximum (100%)")
        
    def min_volume(self):
        self.set_volume(0)
        print("Volume set to minimum (0%)")
        
    def mute(self):
        self.set_volume(0)
        print("Volume muted")
        
    def set_brightness(self, level):
        sbc.set_brightness(level)
        
    def increase_brightness(self, amount=10):
        current = sbc.get_brightness()[0]
        self.set_brightness(min(100, current + amount))
        print(f"Brightness increased to {min(100, current + amount)}%")
        
    def decrease_brightness(self, amount=10):
        current = sbc.get_brightness()[0]
        self.set_brightness(max(0, current - amount))
        print(f"Brightness decreased to {max(0, current - amount)}%")
        
    def take_screenshot(self):
        screenshot = pyautogui.screenshot()
        screenshot.save(f"screenshot_{int(time.time())}.png")
        print("Screenshot taken")
        
    def open_app(self, app_name):
        app_mappings = {
            "whatsapp": (r"C:\Program Files\WindowsApps\5319275A.WhatsAppDesktop_2*\WhatsApp.exe", "WhatsApp.exe"),
            "notepad": (r"C:\Windows\notepad.exe", "notepad.exe"),
            "calculator": (r"C:\Windows\System32\calc.exe", "calc.exe"),
            "file explorer": (r"C:\Windows\explorer.exe", "explorer.exe"),
            "chrome": (r"C:\Program Files\Google\Chrome\Application\chrome.exe", "chrome.exe"),
            "firefox": (r"C:\Program Files\Mozilla Firefox\firefox.exe", "firefox.exe"),
            "edge": (r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe", "msedge.exe"),
            "microsoft edge": (r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe", "msedge.exe"),
            "photos": (r"C:\Program Files\WindowsApps\Microsoft.Windows.Photos_*\Microsoft.Photos.exe", "Microsoft.Photos.exe"),
            "youtube": (r"C:\Program Files\Google\Chrome\Application\chrome.exe --app=https://www.youtube.com", "chrome.exe"),
            "copilot": (r"C:\Windows\System32\cmd.exe", "cmd.exe"),
            "paint": (r"C:\Windows\System32\mspaint.exe", "mspaint.exe")
        }
        
        app_name = app_name.lower().strip()
        if app_name in app_mappings:
            path, exe_name = app_mappings[app_name]
            try:
                if '*' in path:
                    possible_files = glob.glob(path)
                    if possible_files:
                        process = subprocess.Popen(possible_files[0])
                        self.recent_app = (exe_name, process.pid)
                        print(f"Opened {app_name} at {possible_files[0]}")
                        return True
                else:
                    process = subprocess.Popen(path)
                    self.recent_app = (exe_name, process.pid)
                    print(f"Opened {app_name} at {path}")
                    return True
            except Exception as e:
                print(f"Failed to open {app_name} from path: {e}")
        
        try:
            process = subprocess.Popen(app_name + ".exe")
            self.recent_app = (app_name + ".exe", process.pid)
            print(f"Opened {app_name} (fallback)")
            return True
        except Exception as e:
            print(f"Failed to open {app_name} as executable: {e}")
            return False
            
    def close_recent_app(self):
        if self.recent_app:
            exe_name, pid = self.recent_app
            try:
                subprocess.run(['taskkill', '/PID', str(pid), '/F'])
                print(f"Closed recent app: {exe_name}")
                self.recent_app = None
                return True
            except Exception as e:
                print(f"Failed to close {exe_name}: {e}")
                try:
                    subprocess.run(['taskkill', '/IM', exe_name, '/F'])
                    print(f"Closed {exe_name} by name")
                    self.recent_app = None
                    return True
                except:
                    return False
        else:
            print("No recent app to close")
            return False

class VirtualMouse:
    def __init__(self, mp_hands, hands, mp_draw, window_width, window_height):
        self.mp_hands = mp_hands
        self.hands = hands
        self.mp_draw = mp_draw
        self.screen_width, self.screen_height = pyautogui.size()
        self.frame_reduction = 50
        self.smoothening = 10
        self.prev_x, self.prev_y = 0, 0
        self.window_width = window_width
        self.window_height = window_height
        self.last_click_time = 0
        self.double_click_threshold = 0.3
        self.click_distance_threshold = 0.04
        self.drag_distance_threshold = 0.06
        self.is_dragging = False
        self.is_scrolling = False
        self.last_scroll_time = 0
        self.scroll_cooldown = 0.1
        pyautogui.FAILSAFE = False
        pyautogui.PAUSE = 0.01

    def calculate_distance(self, p1, p2):
        return math.sqrt((p2[0] - p1[0])**2 + (p2[1] - p1[1])**2)
    
    def get_finger_positions(self, hand_landmarks, img_shape):
        h, w, _ = img_shape
        landmarks = {
            'wrist': (int(hand_landmarks.landmark[0].x * w), int(hand_landmarks.landmark[0].y * h)),
            'thumb_tip': (int(hand_landmarks.landmark[4].x * w), int(hand_landmarks.landmark[4].y * h)),
            'thumb_mcp': (int(hand_landmarks.landmark[2].x * w), int(hand_landmarks.landmark[2].y * h)),
            'index_tip': (int(hand_landmarks.landmark[8].x * w), int(hand_landmarks.landmark[8].y * h)),
            'middle_tip': (int(hand_landmarks.landmark[12].x * w), int(hand_landmarks.landmark[12].y * h)),
            'ring_tip': (int(hand_landmarks.landmark[16].x * w), int(hand_landmarks.landmark[16].y * h)),
            'pinky_tip': (int(hand_landmarks.landmark[20].x * w), int(hand_landmarks.landmark[20].y * h)),
            'index_mcp': (int(hand_landmarks.landmark[5].x * w), int(hand_landmarks.landmark[5].y * h)),
            'middle_mcp': (int(hand_landmarks.landmark[9].x * w), int(hand_landmarks.landmark[9].y * h)),
            'ring_mcp': (int(hand_landmarks.landmark[13].x * w), int(hand_landmarks.landmark[13].y * h)),
            'pinky_mcp': (int(hand_landmarks.landmark[17].x * w), int(hand_landmarks.landmark[17].y * h))
        }
        is_finger_up = {
            'thumb': landmarks['thumb_tip'][1] < landmarks['thumb_mcp'][1],
            'index': landmarks['index_tip'][1] < landmarks['index_mcp'][1],
            'middle': landmarks['middle_tip'][1] < landmarks['middle_mcp'][1],
            'ring': landmarks['ring_tip'][1] < landmarks['ring_mcp'][1],
            'pinky': landmarks['pinky_tip'][1] < landmarks['pinky_mcp'][1]
        }
        return landmarks, is_finger_up
    
    def detect_gestures(self, hand_landmarks, img_shape):
        landmarks, is_finger_up = self.get_finger_positions(hand_landmarks, img_shape)
        thumb_index_distance = self.calculate_distance(landmarks['thumb_tip'], landmarks['index_tip'])
        thumb_middle_distance = self.calculate_distance(landmarks['thumb_tip'], landmarks['middle_tip'])
        left_click = thumb_index_distance < self.click_distance_threshold * img_shape[1]
        right_click = thumb_middle_distance < self.click_distance_threshold * img_shape[1]
        scrolling = is_finger_up['pinky'] and not is_finger_up['ring']
        all_fingers_touching = all(
            self.calculate_distance(landmarks['thumb_tip'], landmarks[f'{finger}_tip']) < self.drag_distance_threshold * img_shape[1]
            for finger in ['index', 'middle', 'ring', 'pinky']
        )
        dragging = all_fingers_touching
        return left_click, right_click, scrolling, dragging
    
    def move_mouse(self, finger_pos):
        frame_x = np.interp(finger_pos[0], 
                          (self.frame_reduction, self.window_width - self.frame_reduction), 
                          (0, self.screen_width))
        frame_y = np.interp(finger_pos[1], 
                          (self.frame_reduction, self.window_height - self.frame_reduction), 
                          (0, self.screen_height))
        current_x = self.prev_x + (frame_x - self.prev_x) / self.smoothening
        current_y = self.prev_y + (frame_y - self.prev_y) / self.smoothening
        pyautogui.moveTo(current_x, current_y)
        self.prev_x, self.prev_y = current_x, current_y
    
    def handle_hand_gestures(self, results, hand_index, img):
        if not results or not results.multi_hand_landmarks or not img.size:
            if self.is_dragging:
                pyautogui.mouseUp()
                self.is_dragging = False
            return

        if hand_index is None:
            cv2.putText(img, "No Right Hand", (70, 30), 
                       cv2.FONT_HERSHEY_PLAIN, 1, (0,0,255), 1)
            if self.is_dragging:
                pyautogui.mouseUp()
                self.is_dragging = False
            return
        
        hand_landmarks = results.multi_hand_landmarks[hand_index]
        landmarks, _ = self.get_finger_positions(hand_landmarks, img.shape)
        left_click, right_click, scrolling, dragging = self.detect_gestures(hand_landmarks, img.shape)
        
        self.move_mouse(landmarks['index_tip'])
        
        if left_click:
            current_time = time.time()
            if current_time - self.last_click_time < self.double_click_threshold:
                pyautogui.doubleClick()
                cv2.circle(img, (self.window_width//4, self.window_height//2), 
                          15, (0,255,255), -1)
            else:
                pyautogui.click()
                cv2.circle(img, (self.window_width//4, self.window_height//2), 
                          10, (0,255,0), -1)
            self.last_click_time = current_time
        
        if right_click:
            pyautogui.rightClick()
            cv2.circle(img, (3*self.window_width//4, self.window_height//2), 
                      10, (255,0,0), -1)
            time.sleep(0.2)
        
        if dragging:
            if not self.is_dragging:
                pyautogui.mouseDown()
                self.is_dragging = True
                cv2.putText(img, "DRAG MODE", (self.window_width//2 - 50, 50), 
                           cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0,0,255), 2)
        elif self.is_dragging:
            pyautogui.mouseUp()
            self.is_dragging = False
            time.sleep(0.1)
        
        current_time = time.time()
        if scrolling and current_time - self.last_scroll_time > self.scroll_cooldown:
            scroll_amount = -15 if landmarks['pinky_tip'][1] < landmarks['wrist'][1] else 15
            pyautogui.scroll(scroll_amount)
            self.last_scroll_time = current_time
            cv2.putText(img, "SCROLLING", (self.window_width//2 - 50, 30), 
                       cv2.FONT_HERSHEY_PLAIN, 1, (255,255,0), 1)
        
        status_text = "Mouse: " + ("Left Click" if left_click else "Right Click" if right_click else 
                                  "Scrolling" if scrolling else "Dragging" if self.is_dragging else "Moving")
        cv2.putText(img, status_text, (10, 30), cv2.FONT_HERSHEY_PLAIN, 1, (0,255,0), 1)

class VirtualKeyboard:
    def __init__(self, mp_hands, hands, mp_draw, window_width, window_height):
        self.mp_hands = mp_hands
        self.hands = hands
        self.mp_draw = mp_draw
        self.window_width = window_width
        self.window_height = window_height
        self.keyboard = Controller()
        
        self.keyboard_layout = {
            'normal': [
                ['Esc', '1', '2', '3', '4', '5', '6', '7', '8', '9', '0', '-', '=', '⌫'],
                ['Tab', 'q', 'w', 'e', 'r', 't', 'y', 'u', 'i', 'o', 'p', '[', ']', '\\'],
                ['Caps', 'a', 's', 'd', 'f', 'g', 'h', 'j', 'k', 'l', ';', "'", 'Enter'],
                ['Shift', 'z', 'x', 'c', 'v', 'b', 'n', 'm', ',', '.', '/', 'Shift'],
                ['Ctrl', 'Win', 'Alt', 'Space', 'Alt', 'Ctrl']
            ],
            'shift': [
                ['Esc', '!', '@', '#', '$', '%', '^', '&', '*', '(', ')', '_', '+', '⌫'],
                ['Tab', 'Q', 'W', 'E', 'R', 'T', 'Y', 'U', 'I', 'O', 'P', '{', '}', '|'],
                ['Caps', 'A', 'S', 'D', 'F', 'G', 'H', 'J', 'K', 'L', ':', '"', 'Enter'],
                ['Shift', 'Z', 'X', 'C', 'V', 'B', 'N', 'M', '<', '>', '?', 'Shift'],
                ['Ctrl', 'Win', 'Alt', 'Space', 'Alt', 'Ctrl']
            ]
        }
        
        self.special_keys = {
            '⌫': Key.backspace, 'Tab': Key.tab, 'Enter': Key.enter, 'Caps': 'Caps',
            'Shift': 'Shift', 'Space': Key.space, 'Ctrl': Key.ctrl_l, 'Alt': Key.alt_l,
            'Win': Key.cmd, 'Esc': Key.esc
        }
        
        self.shift_pressed = False
        self.caps_lock = False
        self.ctrl_pressed = False
        self.alt_pressed = False
        self.win_pressed = False
        
        self.button_height = 50
        self.button_margin = 6
        self.key_widths = {
            '⌫': 70, 'Tab': 60, 'Caps': 70, 'Enter': 80, 'Shift': 90,
            'Ctrl': 70, 'Alt': 70, 'Win': 70, 'Space': 220, 'Esc': 60
        }
        self.default_key_width = 50
        
        self.key_color = (50, 50, 50)
        self.text_color = (255, 255, 255)
        self.active_key_color = (0, 150, 255)
        self.special_key_color = (70, 70, 70)
        self.click_cooldown = 0.2
        self.last_click_time = 0
        self.prev_clicked = False
        self.prev_finger_x, self.prev_finger_y = 0, 0
        self.smoothening = 15
        self.keyboard_start_x = 20
        self.keyboard_start_y = window_height - 5 * (self.button_height + self.button_margin) - 20

    def get_key_width(self, key):
        return self.key_widths.get(key, self.default_key_width)
    
    def draw_keyboard(self, img):
        current_y = self.keyboard_start_y
        layout = 'shift' if self.shift_pressed else 'normal'
        
        for row_idx, row in enumerate(self.keyboard_layout[layout]):
            current_x = self.keyboard_start_x
            if row_idx == 1:
                current_x += 20
            elif row_idx == 2:
                current_x += 40
            elif row_idx == 3:
                current_x += 60
            elif row_idx == 4:
                current_x += 80
                
            for key in row:
                width = self.get_key_width(key)
                key_col = (self.active_key_color if key in ['Shift', 'Caps', 'Ctrl', 'Alt', 'Win'] and 
                          ((key == 'Shift' and self.shift_pressed) or (key == 'Caps' and self.caps_lock) or
                           (key == 'Ctrl' and self.ctrl_pressed) or (key == 'Alt' and self.alt_pressed) or
                           (key == 'Win' and self.win_pressed)) 
                          else self.special_key_color if key in self.special_keys else self.key_color)
                
                self.draw_rounded_rect(img, (current_x, current_y), 
                                     (current_x + width, current_y + self.button_height), 
                                     key_col, radius=6)
                
                text_size = cv2.getTextSize(key, cv2.FONT_HERSHEY_SIMPLEX, 0.8, 2)[0]
                text_x = current_x + (width - text_size[0]) // 2
                text_y = current_y + (self.button_height + text_size[1]) // 2
                cv2.putText(img, key, (text_x, text_y), 
                           cv2.FONT_HERSHEY_SIMPLEX, 0.8, self.text_color, 2)
                
                current_x += width + self.button_margin
            current_y += self.button_height + self.button_margin

    def draw_rounded_rect(self, img, pt1, pt2, color, radius=6):
        x1, y1 = pt1
        x2, y2 = pt2
        cv2.rectangle(img, (x1 + radius, y1), (x2 - radius, y2), color, -1)
        cv2.rectangle(img, (x1, y1 + radius), (x2, y2 - radius), color, -1)
        cv2.circle(img, (x1 + radius, y1 + radius), radius, color, -1)
        cv2.circle(img, (x2 - radius, y1 + radius), radius, color, -1)
        cv2.circle(img, (x1 + radius, y2 - radius), radius, color, -1)
        cv2.circle(img, (x2 - radius, y2 - radius), radius, color, -1)

    def get_clicked_key(self, finger_pos):
        current_y = self.keyboard_start_y
        layout = 'shift' if self.shift_pressed else 'normal'
        x, y = finger_pos
        
        for row_idx, row in enumerate(self.keyboard_layout[layout]):
            current_x = self.keyboard_start_x
            if row_idx == 1:
                current_x += 20
            elif row_idx == 2:
                current_x += 40
            elif row_idx == 3:
                current_x += 60
            elif row_idx == 4:
                current_x += 80
                
            for key in row:
                width = self.get_key_width(key)
                if (current_x < x < current_x + width and 
                    current_y < y < current_y + self.button_height):
                    return key
                current_x += width + self.button_margin
            current_y += self.button_height + self.button_margin
        return None
    
    def handle_key_press(self, key):
        if not key:
            return
        if key == 'Shift':
            self.shift_pressed = not self.shift_pressed
        elif key == 'Caps':
            self.caps_lock = not self.caps_lock
        elif key == 'Ctrl':
            self.ctrl_pressed = not self.ctrl_pressed
        elif key == 'Alt':
            self.alt_pressed = not self.alt_pressed
        elif key == 'Win':
            self.win_pressed = not self.win_pressed
        elif key in self.special_keys:
            special_key = self.special_keys[key]
            if not isinstance(special_key, str):
                self.keyboard.press(special_key)
                self.keyboard.release(special_key)
        else:
            char = key.upper() if self.caps_lock != self.shift_pressed else key.lower()
            if self.ctrl_pressed and key.lower() in 'cvxz':
                pyautogui.hotkey('ctrl', key.lower())
            elif self.win_pressed and key.lower() == 'd':
                pyautogui.hotkey('win', 'd')
            else:
                self.keyboard.press(char)
                self.keyboard.release(char)
            if self.shift_pressed and key != 'Shift':
                self.shift_pressed = False
    
    def detect_click(self, hand_landmarks):
        index_tip = hand_landmarks.landmark[8]
        thumb_tip = hand_landmarks.landmark[4]
        distance = math.sqrt((thumb_tip.x - index_tip.x) ** 2 + (thumb_tip.y - index_tip.y) ** 2)
        return distance < 0.05
    
    def handle_hand_gestures(self, results, hand_index, img):
        if not results or not results.multi_hand_landmarks or not img.size:
            return
        if hand_index is None:
            cv2.putText(img, "No Left Hand", (70, 30), cv2.FONT_HERSHEY_PLAIN, 1, (0, 0, 255), 1)
            return
        
        hand_landmarks = results.multi_hand_landmarks[hand_index]
        index_tip = hand_landmarks.landmark[8]
        finger_x = int(index_tip.x * self.window_width)
        finger_y = int(index_tip.y * self.window_height)
        finger_x = max(0, min(finger_x, self.window_width-1))
        finger_y = max(0, min(finger_y, self.window_height-1))
        
        current_x = self.prev_finger_x + (finger_x - self.prev_finger_x) / self.smoothening
        current_y = self.prev_finger_y + (finger_y - self.prev_finger_y) / self.smoothening
        self.prev_finger_x, self.prev_finger_y = current_x, current_y
        
        cv2.circle(img, (int(current_x), int(current_y)), 8, (0, 255, 0), cv2.FILLED)
        cv2.circle(img, (int(current_x), int(current_y)), 10, (0, 200, 0), 2)
        
        is_clicked = self.detect_click(hand_landmarks)
        current_time = time.time()
        
        if is_clicked and not self.prev_clicked and current_time - self.last_click_time > self.click_cooldown:
            clicked_key = self.get_clicked_key((current_x, current_y))
            if clicked_key:
                self.handle_key_press(clicked_key)
                self.last_click_time = current_time
                cv2.circle(img, (int(current_x), int(current_y)), 12, (0, 255, 255), -1)
        
        mode = 'Shift' if self.shift_pressed else 'Caps' if self.caps_lock else 'normal'
        status_text = f"Keyboard: {mode}" + (" | Pressing" if is_clicked else "")
        cv2.putText(img, status_text, (10, self.keyboard_start_y - 10), 
                   cv2.FONT_HERSHEY_PLAIN, 1.2, (255, 255, 255), 2)
        self.prev_clicked = is_clicked

class VoiceCommandHandler:
    def __init__(self, system_controller):
        self.recognizer = sr.Recognizer()
        self.microphone = sr.Microphone()
        self.command_queue = queue.Queue()
        self.listening = False
        self.system_controller = system_controller
        
        self.commands = {
            "open": self.handle_open_app,
            "increase volume": lambda: self.system_controller.increase_volume(10),
            "decrease volume": lambda: self.system_controller.decrease_volume(10),
            "max volume": lambda: self.system_controller.max_volume(),
            "min volume": lambda: self.system_controller.min_volume(),
            "mute": lambda: self.system_controller.mute(),
            "increase brightness": lambda: self.system_controller.increase_brightness(10),
            "decrease brightness": lambda: self.system_controller.decrease_brightness(10),
            "screenshot": lambda: self.system_controller.take_screenshot(),
            "close": lambda: pyautogui.hotkey('alt', 'f4'),
            "close recent": lambda: self.system_controller.close_recent_app(),
            "scroll up": lambda: pyautogui.scroll(100),
            "scroll down": lambda: pyautogui.scroll(-100),
            "left click": lambda: pyautogui.click(),
            "right click": lambda: pyautogui.rightClick(),
            "double click": lambda: pyautogui.doubleClick(),
            "drag": lambda: pyautogui.mouseDown(),
            "drop": lambda: pyautogui.mouseUp(),
            "minimize": lambda: pyautogui.hotkey('win', 'down'),
            "maximize": lambda: pyautogui.hotkey('win', 'up'),
            "desktop": lambda: pyautogui.hotkey('win', 'd'),
            "task manager": lambda: pyautogui.hotkey('ctrl', 'shift', 'esc'),
            "help": self.show_help
        }
        
    def handle_open_app(self, text):
        app_name = text.replace("open", "").strip()
        if app_name:
            self.system_controller.open_app(app_name)
        
    def show_help(self):
        print("Available commands: open [app name], increase/decrease volume, max/min volume, mute, "
              "increase/decrease brightness, screenshot, close, close recent, scroll up/down, "
              "left/right/double click, drag/drop, minimize, maximize, desktop, task manager, help")
        
    def listen_in_background(self):
        def callback(recognizer, audio):
            try:
                text = recognizer.recognize_google(audio).lower()
                print(f"Recognized command: '{text}'")
                
                if "open" in text:
                    self.command_queue.put(lambda: self.commands["open"](text))
                    return
                
                for command, action in self.commands.items():
                    if command in text:
                        self.command_queue.put(action if command != "open" else lambda: action(text))
                        return
                
                print(f"Unrecognized command: '{text}'")
            except sr.UnknownValueError:
                print("Could not understand audio")
            except sr.RequestError as e:
                print(f"Voice recognition error: {e}")
            except Exception as e:
                print(f"Unexpected error in voice recognition: {e}")
        
        self.listening = True
        try:
            with self.microphone as source:
                self.recognizer.adjust_for_ambient_noise(source, duration=1)  # Reduced duration for faster startup
                self.recognizer.energy_threshold = 300
                self.recognizer.dynamic_energy_threshold = True  # Adapt to noise levels
                print("Adjusted for ambient noise")
            stop_listening = self.recognizer.listen_in_background(self.microphone, callback, phrase_time_limit=5)
            print("Started listening in background")
            return stop_listening
        except Exception as e:
            print(f"Error starting voice listener: {e}")
            return None
        
    def process_commands(self):
        while not self.command_queue.empty():
            command_action = self.command_queue.get()
            try:
                command_action()
            except Exception as e:
                print(f"Error executing command: {e}")

class MouseAndKeyboard:
    def __init__(self):
        self.mp_hands = mp.solutions.hands
        self.hands = self.mp_hands.Hands(
            max_num_hands=2,
            min_detection_confidence=0.8,
            min_tracking_confidence=0.8
        )
        self.mp_draw = mp.solutions.drawing_utils
        self.window_width = 960
        self.window_height = 540
        self.system_controller = SystemController()
        self.mouse = VirtualMouse(self.mp_hands, self.hands, self.mp_draw, 
                                self.window_width, self.window_height)
        self.keyboard = VirtualKeyboard(self.mp_hands, self.hands, self.mp_draw, 
                                      self.window_width, self.window_height)
        self.voice_handler = VoiceCommandHandler(self.system_controller)
        self.stop_listening = self.voice_handler.listen_in_background()

    def start(self):
        cap = cv2.VideoCapture(0)
        if not cap.isOpened():
            print("Error: Could not open webcam")
            return
        cap.set(cv2.CAP_PROP_FRAME_WIDTH, self.window_width)
        cap.set(cv2.CAP_PROP_FRAME_HEIGHT, self.window_height)
        cap.set(cv2.CAP_PROP_FPS, 30)
        
        screen_width, screen_height = pyautogui.size()
        window_x = 0
        window_y = screen_height - self.window_height - 40
        
        WINDOW_NAME = "Virtual Mouse and Keyboard with Voice Control"
        cv2.namedWindow(WINDOW_NAME, cv2.WINDOW_NORMAL)
        cv2.setWindowProperty(WINDOW_NAME, cv2.WND_PROP_TOPMOST, 1)
        cv2.moveWindow(WINDOW_NAME, window_x, window_y)
        cv2.resizeWindow(WINDOW_NAME, self.window_width, self.window_height)

        while True:
            success, camera_img = cap.read()
            if not success:
                print("Failed to capture image")
                continue

            camera_img = cv2.flip(camera_img, 1)
            camera_img = cv2.convertScaleAbs(camera_img, alpha=1.3, beta=30)
            rgb_camera_img = cv2.cvtColor(camera_img, cv2.COLOR_BGR2RGB)
            results = self.hands.process(rgb_camera_img)
            
            img = np.zeros((self.window_height, self.window_width, 3), dtype=np.uint8)
            self.keyboard.draw_keyboard(img)

            right_hand_index = None
            left_hand_index = None

            if results.multi_hand_landmarks:
                for idx, hand_landmarks in enumerate(results.multi_hand_landmarks):
                    self.mp_draw.draw_landmarks(
                        camera_img, hand_landmarks, self.mp_hands.HAND_CONNECTIONS,
                        self.mp_draw.DrawingSpec(color=(255,0,255), thickness=2, circle_radius=4),
                        self.mp_draw.DrawingSpec(color=(0,255,0), thickness=2)
                    )
                    hand_type = results.multi_handedness[idx].classification[0].label
                    if hand_type == "Right":
                        right_hand_index = idx
                    elif hand_type == "Left":
                        left_hand_index = idx

            if right_hand_index is not None:
                self.mouse.handle_hand_gestures(results, right_hand_index, camera_img)
            if left_hand_index is not None:
                self.keyboard.handle_hand_gestures(results, left_hand_index, img)
                
            self.voice_handler.process_commands()

            camera_img = cv2.resize(camera_img, 
                                  (int(self.window_height * camera_img.shape[1] / camera_img.shape[0]), 
                                  self.window_height))
            combined_img = np.zeros((self.window_height, self.window_width + camera_img.shape[1], 3), 
                                  dtype=np.uint8)
            combined_img[:, :camera_img.shape[1]] = camera_img
            combined_img[:, camera_img.shape[1]:] = img
            
            cv2.putText(combined_img, "Voice: Listening | Say 'open [app]', 'increase volume', or 'help'", 
                       (10, 20), cv2.FONT_HERSHEY_PLAIN, 1.5, (0, 255, 255), 2)

            cv2.imshow(WINDOW_NAME, combined_img)
            if cv2.waitKey(1) == ord('q') or cv2.getWindowProperty(WINDOW_NAME, cv2.WND_PROP_VISIBLE) < 1:
                break
                
        cap.release()
        if self.stop_listening:
            self.stop_listening()
        cv2.destroyAllWindows()

if __name__ == "__main__":
    try:
        MouseAndKeyboard().start()
    except Exception as e:
        print(f"Error: {e}")