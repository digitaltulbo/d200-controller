#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Studio Birthday - D200 Macro Controller (Headless)
Platform: Windows 11
"""

import os
import sys
import json
import time
import subprocess
import threading
from pathlib import Path
from datetime import datetime

# Library imports with error handling for non-Windows dev environment
try:
    import keyboard
    import psutil
    import pygame
except ImportError as e:
    print(f"Error: Missing required library: {e}")
    print("Please install requirements: pip install keyboard psutil pygame pywin32")
    # For dev purposes, we don't exit immediately to allow file creation, 
    # but runtime will fail if libs are missing.

WINDOWS_AVAILABLE = False
try:
    import win32gui
    import win32con
    WINDOWS_AVAILABLE = True
except ImportError:
    pass

# =============================================================================
# Configuration & Constants
# =============================================================================

BASE_DIR = Path(__file__).parent
CONFIG_FILE = BASE_DIR / "config.json"
SOUNDS_DIR = BASE_DIR / "Sounds"

class Config:
    def __init__(self):
        self.data = self._load()
    
    def _load(self):
        if CONFIG_FILE.exists():
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                print(f"[설정 오류] 파일을 읽을 수 없습니다: {e}")
        return {}

    def get(self, key, default=None):
        return self.data.get(key, default)

# =============================================================================
# Sound Player
# =============================================================================

class SoundPlayer:
    @staticmethod
    def play(filename):
        path = SOUNDS_DIR / filename
        if not path.exists():
            print(f"[사운드 경고] 파일을 찾을 수 없음: {filename}")
            return

        def _worker():
            try:
                if not pygame.mixer.get_init():
                    pygame.mixer.init()
                pygame.mixer.music.load(str(path))
                pygame.mixer.music.play()
                while pygame.mixer.music.get_busy():
                    time.sleep(0.1)
            except Exception as e:
                print(f"[사운드 오류] 재생 실패: {e}")
        
        threading.Thread(target=_worker, daemon=True).start()

# =============================================================================
# Lightroom Controller
# =============================================================================

class LightroomController:
    def __init__(self, config):
        self.config = config
        self.is_running_macro = False
        self._lock = threading.Lock()

    def _find_window(self, title_text):
        if not WINDOWS_AVAILABLE: return None
        result = []
        def callback(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                t = win32gui.GetWindowText(hwnd)
                if title_text.lower() in t.lower():
                    result.append(hwnd)
        try:
            win32gui.EnumWindows(callback, None)
        except: pass
        return result[0] if result else None

    def _activate_window(self, hwnd):
        if not WINDOWS_AVAILABLE: return
        try:
            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
            return True
        except: return False

    def launch_and_focus(self):
        # 1. Check/Launch
        lr_path = self.config.get("lightroom_path")
        process_name = self.config.get("lightroom_process_name", "Lightroom.exe")
        
        is_running = False
        for proc in psutil.process_iter(['name']):
            if proc.info['name'] and process_name.lower() in proc.info['name'].lower():
                is_running = True
                break
        
        if not is_running:
            if not lr_path or not os.path.exists(lr_path):
                print(f"[오류] 라이트룸 경로가 잘못되었습니다: {lr_path}")
                return False
            print("[시스템] 라이트룸 실행 중...")
            subprocess.Popen([lr_path])
            time.sleep(5) # Initial wait
        
        # 2. Focus
        title_keyword = self.config.get("lightroom_window_title_contains", "Lightroom")
        print(f"[시스템] 라이트룸 창 활성화 대기 중 ({title_keyword})...")
        
        for i in range(20):
            hwnd = self._find_window(title_keyword)
            if hwnd:
                self._activate_window(hwnd)
                time.sleep(0.5)
                # Double check focus
                if WINDOWS_AVAILABLE and win32gui.GetForegroundWindow() == hwnd:
                    print("[시스템] 라이트룸 포커스 확보 완료")
                    return True
            time.sleep(1.0)
            
        print("[오류] 라이트룸 창을 찾을 수 없거나 활성화할 수 없습니다.")
        return False

    def kill_process(self):
        process_name = self.config.get("lightroom_process_name", "Lightroom.exe")
        killed = False
        for proc in psutil.process_iter(['name']):
            try:
                if proc.info['name'] and process_name.lower() in proc.info['name'].lower():
                    proc.terminate()
                    killed = True
            except: pass
        
        if killed:
            print("[시스템] 라이트룸 프로세스가 종료되었습니다.")
        else:
            print("[시스템] 실행 중인 라이트룸 프로세스가 없습니다.")

    def run_tether_sequence(self):
        if self.is_running_macro:
            print("[경고] 이미 매크로가 실행 중입니다.")
            return

        with self._lock:
            self.is_running_macro = True

        try:
            if not self.launch_and_focus():
                return

            print("[매크로] 테더링 시작 시퀀스 진행...")
            time.sleep(1.0) # Stability wait

            # a. Alt+F
            keyboard.send('alt+f')
            time.sleep(0.4)

            # b. Down x 8
            for _ in range(8):
                keyboard.send('down')
                time.sleep(0.1)
            
            # c. Right
            keyboard.send('right')
            time.sleep(0.3)

            # d. Enter
            keyboard.send('enter')
            time.sleep(0.8)

            # e. Session Name
            session_name = datetime.now().strftime("%Y-%m-%d_%H-%M")
            keyboard.write(session_name)
            
            # f. Tab x 4
            for _ in range(4):
                time.sleep(0.2)
                keyboard.send('tab')
            
            # g. Type "1"
            time.sleep(0.2)
            keyboard.write('1')
            time.sleep(0.3)

            # h. Enter
            keyboard.send('enter')
            time.sleep(2.5) # Wait for tether bar creation

            # i. Ctrl+Alt+1 (Preset)
            keyboard.send('ctrl+alt+1')
            time.sleep(0.8)

            # j. E (Library View)
            keyboard.send('e')

            print(f"[완료] 테더링 시작됨: {session_name}")
            SoundPlayer.play("Start_shoot.mp3")

        except Exception as e:
            print(f"[오류] 매크로 실행 중 에러 발생: {e}")
        finally:
            self.is_running_macro = False

# =============================================================================
# Main
# =============================================================================

def main():
    print("\n" + "="*50)
    print("   D200 Studio Controller (Headless)")
    print("   - Ctrl+Alt+Shift+F1 : 촬영 시작 (테더링)")
    print("   - Ctrl+Alt+Shift+F3 : 세션 종료 (강제종료)")
    print("="*50 + "\n")

    config = Config()
    lr_controller = LightroomController(config)

    def on_start():
        print("\n>>> [명령] 촬영 시작 요청")
        threading.Thread(target=lr_controller.run_tether_sequence, daemon=True).start()

    def on_end():
        print("\n>>> [명령] 세션 종료 요청")
        lr_controller.kill_process()
        SoundPlayer.play("The_end.mp3")
        print("[정보] 세션이 종료되었습니다.")

    # Hotkey Registration
    try:
        # suppress=True prevents the key from reaching the active window
        keyboard.add_hotkey('ctrl+alt+shift+f1', on_start, suppress=True)
        keyboard.add_hotkey('ctrl+alt+shift+f3', on_end, suppress=True)
        print("[대기] 핫키 입력 대기 중... (종료하려면 터미널 닫기)")
        keyboard.wait()
    except KeyboardInterrupt:
        print("\n[시스템] 프로그램을 종료합니다.")
    except Exception as e:
        print(f"\n[치명적 오류] {e}")

if __name__ == "__main__":
    main()