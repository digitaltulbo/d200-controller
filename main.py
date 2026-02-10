#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Ulanzi D200 Controller for Studio Birthday
Headless Mode - Hotkey Driven
"""

import os
import sys
import json
import time
import threading
import subprocess
from pathlib import Path
from datetime import datetime

# Check for required libraries
try:
    import keyboard
    import psutil
    # Check for windows specific modules if on Windows, but since we are on macOS for development
    # we might need to mock or handle this. 
    # The user says "My operating system is: darwin", but the target environment seems to be Windows 
    # (based on paths like "C:\Program Files..." and usage of win32gui in original code).
    # I will implement the logic assuming the code RUNS on Windows, but since I can't test win32gui here,
    # I will include the imports inside try/except blocks or check functionality.
    # However, the user specifically asked to replicate the logic.
    # IMPORTANT: The user is on macOS but the project is for a Windows environment (Lightroom path, .exe, etc).
    # I will write the code to be compatible with Windows, even if I can't run it fully here.
except ImportError as e:
    print(f"Missing required library: {e}")
    sys.exit(1)

# Windows libraries
try:
    import win32gui
    import win32con
    import win32api
    WINDOWS_AVAILABLE = True
except ImportError:
    # On macOS this will fail, which is expected during this dev phase
    WINDOWS_AVAILABLE = False

# =============================================================================
# Constants & Config
# =============================================================================

BASE_DIR = Path(__file__).parent
CONFIG_FILE = BASE_DIR / "config.json"
SOUNDS_DIR = BASE_DIR / "Sounds"

# =============================================================================
# Sound Player
# =============================================================================

class SoundPlayer:
    """MP3 Sound Player using pygame"""
    
    SOUND_FILES = {
        'start': 'Start_shoot.mp3',
        'end_15min': 'end_15min.mp3',
        'end_5min': 'end_5min.mp3',
        'end': 'The_end.mp3'
    }
    
    @staticmethod
    def play(sound_type: str):
        sound_file = SoundPlayer.SOUND_FILES.get(sound_type)
        if not sound_file:
            return
        
        sound_path = SOUNDS_DIR / sound_file
        if not sound_path.exists():
            print(f"[Sound] File not found: {sound_path}")
            return
            
        def _play_thread():
            try:
                import pygame
                if not pygame.mixer.get_init():
                    pygame.mixer.init()
                pygame.mixer.music.load(str(sound_path))
                pygame.mixer.music.play()
                while pygame.mixer.music.get_busy():
                    time.sleep(0.1)
            except Exception as e:
                print(f"[Sound] Error: {e}")

        threading.Thread(target=_play_thread, daemon=True).start()

# =============================================================================
# Utilities
# =============================================================================

class ConfigManager:
    def __init__(self):
        self.config = self._load()
        
    def _load(self):
        if CONFIG_FILE.exists():
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                print(f"[Config] Load error: {e}")
        return {}
    
    def get(self, key, default=None):
        keys = key.split('.')
        val = self.config
        try:
            for k in keys: val = val[k]
            return val
        except: return default

class WindowsController:
    def __init__(self, config):
        self.config = config

    def is_process_running(self, process_name: str) -> bool:
        if not WINDOWS_AVAILABLE: 
            # Mock for non-Windows dev environment
            return False 
        for proc in psutil.process_iter(['name']):
            try:
                if proc.info['name'] and process_name.lower() in proc.info['name'].lower():
                    return True
            except: pass
        return False

    def find_window_by_title(self, title_contains: str):
        if not WINDOWS_AVAILABLE: return None
        result = []
        def enum_callback(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if title_contains.lower() in title.lower():
                    result.append(hwnd)
            return True
        win32gui.EnumWindows(enum_callback, None)
        return result[0] if result else None

    def activate_window(self, hwnd: int) -> bool:
        if not WINDOWS_AVAILABLE or not hwnd: return False
        try:
            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
            wait_time = self.config.get('delays.window_activation_wait_ms', 500) / 1000.0
            time.sleep(wait_time)
            return True
        except: return False

    def ensure_lightroom_running(self) -> bool:
        process_name = self.config.get('lightroom_process_name', 'Lightroom.exe')
        
        if self.is_process_running(process_name):
            print("[System] Lightroom is already running.")
            return True
            
        lr_path = self.config.get('lightroom_path')
        if not lr_path or not os.path.exists(lr_path):
            print(f"[System] Lightroom path invalid: {lr_path}")
            return False
            
        print("[System] Launching Lightroom...")
        try:
            subprocess.Popen([lr_path])
            title_contains = self.config.get('lightroom_window_title_contains', 'Lightroom')
            # Wait for window to appear
            for _ in range(20):
                time.sleep(1.5)
                if self.find_window_by_title(title_contains):
                    time.sleep(5) # Extra wait for initialization
                    return True
        except Exception as e:
            print(f"[System] Failed to launch Lightroom: {e}")
            return False
        return False

    def wait_for_lightroom_focus(self, max_retries: int = 10) -> bool:
        if not WINDOWS_AVAILABLE:
            print("[System] Not on Windows, skipping focus check.")
            return True
            
        title_contains = self.config.get('lightroom_window_title_contains', 'Lightroom')
        print(f"[System] Waiting for focus: {title_contains}")
        
        for _ in range(max_retries):
            hwnd = self.find_window_by_title(title_contains)
            if not hwnd:
                time.sleep(1.5)
                continue
            
            # Try to bring to front
            self.activate_window(hwnd)
            
            if win32gui.GetForegroundWindow() == hwnd:
                return True
            time.sleep(0.8)
            
        print("[System] Failed to acquire focus.")
        return False

# =============================================================================
# Session Timer
# =============================================================================

class SessionTimer:
    def __init__(self, duration_minutes: int, on_remind=None, on_end=None):
        self.duration_minutes = duration_minutes
        self.total_seconds = duration_minutes * 60
        self.remaining_seconds = self.total_seconds
        self.is_running = False
        self._thread = None
        self.on_remind = on_remind
        self.on_end = on_end
        self.reminder_points = {15: 'end_15min', 5: 'end_5min'}
        self.reminded = set()
    
    def start(self):
        if self.is_running: return
        self.is_running = True
        self.reminded.clear()
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()
        print(f"[Timer] Started for {self.duration_minutes} minutes.")
    
    def stop(self):
        if self.is_running:
            self.is_running = False
            print("[Timer] Stopped.")
    
    def _run(self):
        while self.remaining_seconds > 0 and self.is_running:
            time.sleep(1)
            self.remaining_seconds -= 1
            
            remaining_min = self.remaining_seconds // 60
            if remaining_min in self.reminder_points and remaining_min not in self.reminded:
                # Check if we just crossed the minute boundary exactly or close to it
                # Logic from original: integer check
                self.reminded.add(remaining_min)
                sound_key = self.reminder_points[remaining_min]
                print(f"[Timer] Reminder: {remaining_min} min left. Playing {sound_key}")
                SoundPlayer.play(sound_key)
                
        if self.is_running: # If finished naturally
            print("[Timer] Time is up!")
            SoundPlayer.play('end')
            if self.on_end: self.on_end()
        self.is_running = False

# =============================================================================
# Main Controller
# =============================================================================

class D200Controller:
    def __init__(self):
        self.config = ConfigManager()
        self.win = WindowsController(self.config)
        self.timer = None
        self.is_session_active = False
        self._lock = threading.Lock()

    def start_session(self, minutes: int):
        with self._lock:
            if self.is_session_active:
                print("[Action] Session already active! Ignoring start command.")
                # Optional: Play error sound or "already running" sound
                return

            print(f"
>>> Starting {minutes} min Session <<<")
            self.is_session_active = True
        
        # Run sequence in separate thread to not block keyboard listener
        threading.Thread(target=self._run_start_sequence, args=(minutes,), daemon=True).start()

    def end_session(self):
        with self._lock:
            if not self.is_session_active:
                print("[Action] No active session to end.")
                return
            
            print("
>>> Ending Session <<<")
            if self.timer:
                self.timer.stop()
                self.timer = None
            
            self.is_session_active = False
            SoundPlayer.play('end')
            print("[Action] Session ended manually.")

    def _run_start_sequence(self, minutes):
        try:
            # 1. Ensure Lightroom
            if not self.win.ensure_lightroom_running():
                print("[Error] Could not start Lightroom.")
                self.is_session_active = False
                return

            # 2. Focus
            if not self.win.wait_for_lightroom_focus():
                print("[Error] Could not focus Lightroom.")
                self.is_session_active = False
                return

            # 3. Tethering Macro Sequence
            # Hardcoded logic based on Dashboard.py analysis
            print("[Macro] Sending Tether Start keys...")
            time.sleep(1.5)
            
            # File Menu
            keyboard.send('alt+f')
            time.sleep(0.5)
            
            # Navigate to Tether
            for _ in range(8):
                keyboard.send('down')
                time.sleep(0.1)
            
            keyboard.send('right')
            time.sleep(0.3)
            keyboard.send('enter')
            time.sleep(0.8)
            
            # Session Name
            session_name = datetime.now().strftime("%Y-%m-%d_%H-%M")
            keyboard.write(session_name)
            print(f"[Macro] Set session name: {session_name}")
            
            # Tabs
            for _ in range(4):
                time.sleep(0.2)
                keyboard.send('tab')
            
            # File Numbering Start
            time.sleep(0.2)
            keyboard.write('1')
            time.sleep(0.3)
            keyboard.send('enter')
            
            # Post-setup View
            time.sleep(2.5)
            keyboard.send('ctrl+alt+1')
            time.sleep(0.8)
            keyboard.send('e')
            
            print("[Macro] Sequence complete.")
            
            # 4. Start Timer & Sound
            SoundPlayer.play('start')
            self.timer = SessionTimer(minutes, on_end=self._on_timer_end)
            self.timer.start()
            
        except Exception as e:
            print(f"[Error] Exception during sequence: {e}")
            self.is_session_active = False

    def _on_timer_end(self):
        # Timer finished naturally
        with self._lock:
            self.is_session_active = False
        print("[Session] Time expired.")


# =============================================================================
# Entry Point
# =============================================================================

def main():
    print("==========================================")
    print("   Studio Birthday - D200 Controller      ")
    print("   Headless Mode v1.0                     ")
    print("==========================================")
    
    controller = D200Controller()
    
    # Define Hotkeys
    # Use suppress=True to block the key from going to other apps
    
    # 1. Basic (30 min) - Ctrl+Alt+Shift+F1
    hk_basic = 'ctrl+alt+shift+f1'
    keyboard.add_hotkey(hk_basic, lambda: controller.start_session(30), suppress=True)
    print(f"[Ready] Hotkey registered: {hk_basic} -> Basic Session (30m)")
    
    # 2. Premium (55 min) - Ctrl+Alt+Shift+F2
    hk_premium = 'ctrl+alt+shift+f2'
    keyboard.add_hotkey(hk_premium, lambda: controller.start_session(55), suppress=True)
    print(f"[Ready] Hotkey registered: {hk_premium} -> Premium Session (55m)")
    
    # 3. End Session - Ctrl+Alt+Shift+F3
    hk_end = 'ctrl+alt+shift+f3'
    keyboard.add_hotkey(hk_end, lambda: controller.end_session(), suppress=True)
    print(f"[Ready] Hotkey registered: {hk_end} -> Force End Session")
    
    print("
Waiting for input... (Press Ctrl+C to quit program)")
    
    try:
        # Keep main thread alive
        keyboard.wait()
    except KeyboardInterrupt:
        print("
[System] Shutting down...")
    except Exception as e:
        print(f"
[Error] Unexpected error: {e}")

if __name__ == "__main__":
    main()
