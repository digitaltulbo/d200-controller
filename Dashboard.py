#!/usr/bin/env pythonw
# -*- coding: utf-8 -*-
"""
=============================================================================
Studio Birthday - Lightroom Macro Panel (Hybrid Portable Edition)
=============================================================================
핵심 로직은 Dashboard.exe로 실행되며, UI와 사운드, 설정은 외부 폴더에서 관리됩니다.
"""

import os
import sys
import json
import time
import ctypes
import threading
import subprocess
import glob
import shutil
import zipfile
from pathlib import Path
from datetime import datetime

import webview
import psutil

# 실행 경로 설정 (EXE 실행 시와 스크립트 실행 시 대응)
if getattr(sys, 'frozen', False):
    # EXE로 실행 중일 때 (PyInstaller)
    BASE_DIR = Path(sys.executable).parent
else:
    # 일반 파이썬 스크립트로 실행 중일 때
    BASE_DIR = Path(__file__).parent

# Windows 전용 모듈
try:
    import win32gui
    import win32con
    import win32api
    import keyboard
    WINDOWS_AVAILABLE = True
except ImportError:
    WINDOWS_AVAILABLE = False

# =============================================================================
# 설정
# =============================================================================

APP_NAME = "Lightroom Macro Panel"
APP_VERSION = "4.1.0 (Portable)"
CONFIG_FILE = BASE_DIR / "config.json"

DEFAULT_CONFIG = {
    "export_target_folder": "Desktop\\내보내기",
    "session_duration_basic": 30,      # 베이직 패키지 (분)
    "session_duration_premium": 55,    # 프리미엄 패키지 (분)
    "gui_settings": {
        "monitor_index": 1,
        "fullscreen": True,
        "width": 1920,
        "height": 1280
    }
}


# =============================================================================
# Sound Player (pygame)
# =============================================================================

class SoundPlayer:
    """MP3 사운드 파일 재생 (pygame 사용)"""
    
    SOUND_FILES = {
        'start': 'Start_shoot.mp3',
        'end_15min': 'end_15min.mp3',
        'end_5min': 'end_5min.mp3',
        'end': 'The_end.mp3'
    }
    
    _initialized = False
    
    @classmethod
    def get_sounds_dir(cls):
        return BASE_DIR / "Sounds"
    
    @classmethod
    def play(cls, sound_type: str):
        """사운드 재생 (비동기)"""
        sound_file = cls.SOUND_FILES.get(sound_type)
        if not sound_file:
            return
        
        sound_path = cls.get_sounds_dir() / sound_file
        
        if not sound_path.exists():
            print(f"Sound file not found: {sound_path}")
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
                print(f"Sound playback error: {e}")
        
        threading.Thread(target=_play_thread, daemon=True).start()


# =============================================================================
# Session Timer
# =============================================================================

class SessionTimer:
    """촬영 세션 타이머 + 사운드 알림"""
    
    def __init__(self, duration_minutes: int, on_tick=None, on_remind=None, on_end=None):
        self.duration_minutes = duration_minutes
        self.total_seconds = duration_minutes * 60
        self.remaining_seconds = self.total_seconds
        self.is_running = False
        self._thread = None
        self.on_tick = on_tick
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
        SoundPlayer.play('start')
    
    def stop(self):
        self.is_running = False
    
    def _run(self):
        while self.remaining_seconds > 0 and self.is_running:
            time.sleep(1)
            self.remaining_seconds -= 1
            if self.on_tick: self.on_tick(self.remaining_seconds)
            remaining_min = self.remaining_seconds // 60
            if remaining_min in self.reminder_points and remaining_min not in self.reminded:
                self.reminded.add(remaining_min)
                SoundPlayer.play(self.reminder_points[remaining_min])
                if self.on_remind: self.on_remind(f"{remaining_min}분 남았습니다!")
        if self.is_running:
            SoundPlayer.play('end')
            if self.on_end: self.on_end()
        self.is_running = False


# =============================================================================
# Config Manager
# =============================================================================

class ConfigManager:
    def __init__(self, config_path: Path = CONFIG_FILE):
        self.config_path = config_path
        self.config = self._load()
    
    def _load(self):
        if self.config_path.exists():
            try:
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except: pass
        return DEFAULT_CONFIG.copy()
    
    def get(self, key, default=None):
        keys = key.split('.')
        val = self.config
        try:
            for k in keys: val = val[k]
            return val
        except: return default


# =============================================================================
# Windows Controller
# =============================================================================

class WindowsController:
    def __init__(self, config: ConfigManager):
        self.config = config
    
    def is_process_running(self, process_name: str) -> bool:
        if not WINDOWS_AVAILABLE: return False
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
            time.sleep(self.config.get('delays.window_activation_wait_ms', 500) / 1000.0)
            return True
        except: return False

    def ensure_lightroom_running(self) -> bool:
        process_name = self.config.get('lightroom_process_name', 'Lightroom.exe')
        if self.is_process_running(process_name): return True
        lr_path = self.config.get('lightroom_path')
        if not lr_path or not os.path.exists(lr_path): return False
        try:
            subprocess.Popen([lr_path])
            title_contains = self.config.get('lightroom_window_title_contains', 'Lightroom')
            for _ in range(20):
                time.sleep(1.5)
                if self.find_window_by_title(title_contains):
                    time.sleep(5)
                    return True
        except: return False
        return False

    def wait_for_lightroom_focus(self, max_retries: int = 10) -> bool:
        if not WINDOWS_AVAILABLE: return True
        title_contains = self.config.get('lightroom_window_title_contains', 'Lightroom')
        for _ in range(max_retries):
            hwnd = self.find_window_by_title(title_contains)
            if not hwnd:
                time.sleep(1.5)
                continue
            if win32gui.GetForegroundWindow() == hwnd: return True
            self.activate_window(hwnd)
            time.sleep(0.8)
            if win32gui.GetForegroundWindow() == hwnd: return True
        return False


# =============================================================================
# Macro Actions
# =============================================================================

class MacroActions:
    def __init__(self, config: ConfigManager):
        self.config = config
        self.win = WindowsController(config)
    
    def start_tether(self):
        if not self.win.ensure_lightroom_running(): return "라이트룸 실행 실패"
        if not self.win.wait_for_lightroom_focus(): return "라이트룸 포커스 실패"
        time.sleep(1.5)
        keyboard.send('alt+f')
        time.sleep(0.5)
        for _ in range(8):
            keyboard.send('down')
            time.sleep(0.1)
        keyboard.send('right')
        time.sleep(0.3)
        keyboard.send('enter')
        time.sleep(0.8)
        session_name = datetime.now().strftime("%Y-%m-%d_%H-%M")
        keyboard.write(session_name)
        for _ in range(4):
            time.sleep(0.2)
            keyboard.send('tab')
        time.sleep(0.2)
        keyboard.write('1')
        time.sleep(0.3)
        keyboard.send('enter')
        time.sleep(2.5)
        keyboard.send('ctrl+alt+1')
        time.sleep(0.8)
        keyboard.send('e')
        return f"테더링 시작: {session_name}"
    
    def export_all(self):
        """라이트룸에서 전체 사진 내보내기 (촬영 시작과 동일한 방식)"""
        # 1. 라이트룸이 실행 중인지 확인 + 포커스 확보
        if not self.win.ensure_lightroom_running():
            return "라이트룸이 실행 중이지 않습니다."
        if not self.win.wait_for_lightroom_focus():
            return "라이트룸 창을 활성화할 수 없습니다."
        
        # 2. 라이트룸 UI가 준비될 때까지 대기 (촬영 시작과 동일)
        time.sleep(1.5)
        
        # 3. 전체 선택 (Ctrl+A) - keyboard.send 사용
        keyboard.send('ctrl+a')
        time.sleep(0.5)
        
        # 4. 내보내기 단축키 (Ctrl+Alt+Shift+E) - keyboard.send 사용
        keyboard.send('ctrl+alt+shift+e')
        time.sleep(0.3)
        
        return "라이트룸에 내보내기 명령을 전달했습니다."

    
    def compress_folder(self):
        """내보내기 폴더를 ZIP으로 압축"""
        # 1. 경로 설정 (바탕화면 내 '내보내기' 폴더 고정)
        desktop_path = Path.home() / "Desktop"
        source_path = desktop_path / "내보내기"
        
        # 2. 폴더가 없으면 생성 (에러 방지)
        if not source_path.exists():
            source_path.mkdir(parents=True, exist_ok=True)
            return None, "바탕화면에 '내보내기' 폴더가 생성되었습니다. 사진을 넣어주세요."
            
        # 3. 파일 존재 확인
        files = [f for f in source_path.iterdir() if f.is_file()]
        if not files:
            return None, "폴더에 사진이 없습니다. 내보내기를 먼저 해주세요."
        
        # 4. 압축 진행
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        zip_filename = f"사진_{timestamp}.zip"
        zip_path = desktop_path / zip_filename
        
        try:
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                for file in source_path.rglob("*"):
                    if file.is_file():
                        zf.write(file, file.relative_to(source_path))
            return str(zip_path), f"압축 완료: {zip_filename}"
        except Exception as e:
            return None, f"압축 오류: {str(e)}"
    
    def end_session(self):
        """촬영 종료 - 폴더 비우기 + 라이트룸 종료"""
        desktop_path = Path.home() / "Desktop"
        source_path = desktop_path / "내보내기"
        
        if source_path.exists():
            for item in source_path.iterdir():
                try:
                    if item.is_file(): item.unlink()
                    elif item.is_dir(): shutil.rmtree(item)
                except: pass
        
        for zip_file in desktop_path.glob("사진_*.zip"):
            try: zip_file.unlink()
            except: pass
            
        process_name = self.config.get('lightroom_process_name', 'Lightroom.exe')
        for proc in psutil.process_iter(['name']):
            try:
                if proc.info['name'] and process_name.lower() in proc.info['name'].lower():
                    proc.terminate()
            except: pass
        return "세션 종료 완료"


# =============================================================================
# JavaScript API
# =============================================================================

class Api:
    def __init__(self, actions: MacroActions, config: ConfigManager):
        self.actions = actions
        self.config = config
        self.window = None
        self.timer = None
        self.local_version = "4.1.0"
        self._load_local_version()

    def _load_local_version(self):
        v_path = BASE_DIR / "version.json"
        if v_path.exists():
            try:
                with open(v_path, "r", encoding="utf-8") as f:
                    self.local_version = json.load(f).get("version", self.local_version)
            except: pass

    def check_update(self):
        """본부(구글 드라이브)의 버전과 비교 (지연 방지를 위해 별도 스레드 권장)"""
        # 드라이브 체크를 매우 가볍게 시도
        sources = [
            Path("H:/내 드라이브/01.Studio-Improvement/lightroom_macro_panel-v3_portable/version.json"),
            Path("G:/내 드라이브/01.Studio-Improvement/lightroom_macro_panel-v3_portable/version.json")
        ]
        
        for v_path in sources:
            try:
                # 존재 여부 확인 시 타임아웃을 유발할 수 있는 OS 호출 최소화
                if v_path.is_file(): 
                    with open(v_path, "r", encoding="utf-8") as f:
                        data = json.load(f)
                        remote_version = data.get("version", "")
                        if remote_version and remote_version > self.local_version:
                            return {
                                "available": True,
                                "version": remote_version,
                                "message": data.get("message", "새로운 버전이 준비되었습니다.")
                            }
            except: 
                continue # 드라이브가 없거나 권한 에러 시 즉시 다음으로
        return {"available": False}

    def apply_update(self):
        """업데이트 스크립트 실행 후 프로그램 종료"""
        def do_update():
            sync_script = Path("H:/내 드라이브/01.Studio-Improvement/SYNC.bat")
            if not sync_script.exists():
                sync_script = Path("G:/내 드라이브/01.Studio-Improvement/SYNC.bat")

            if sync_script.exists():
                subprocess.Popen(['cmd', '/c', 'start', '', str(sync_script)], shell=True)
                time.sleep(0.5)
                os._exit(0)
        
        threading.Thread(target=do_update, daemon=True).start()
        return "업데이트 시작..."

    def set_window(self, window): self.window = window

    def execute_action(self, action: str) -> str:
        """액션을 완전히 별도의 스레드에서 실행하여 UI 프리징 방지"""
        def run_in_bg():
            result = ""
            try:
                if action == "export":
                    result = self.actions.export_all()
                elif action == "compress":
                    _, result = self.actions.compress_folder()
                elif action == "end":
                    self._stop_timer()
                    result = self.actions.end_session()
                else:
                    result = f"알 수 없는 액션: {action}"
            except Exception as e:
                result = f"오류: {str(e)}"
            
            # 메인 UI가 응답할 수 있도록 결과 전달 시점을 약간 조절
            if self.window:
                # 안전한 호출을 위해 evaluate_js 전송 전 짧은 휴식
                time.sleep(0.2)
                self.window.evaluate_js(f'hideLoading(); updateStatus("{result}")')
        
        thread = threading.Thread(target=run_in_bg, daemon=True)
        thread.start()
        return "명령 전달됨"
    
    def start_session(self, minutes: int) -> str:
        """세션 시작을 백그라운드에서 실행"""
        def run_in_bg():
            result = ""
            try:
                # 타이머는 UI 스레드에서 즉시 반응하도록 수정 가능하지만 가급적 여기서 처리
                result = self.actions.start_tether()
                self._start_timer(minutes)
            except Exception as e:
                result = f"오류: {str(e)}"
            
            if self.window:
                time.sleep(0.2)
                self.window.evaluate_js(f'hideLoading(); updateStatus("{result}")')
        
        threading.Thread(target=run_in_bg, daemon=True).start()
        return "준비 중..."



    def _start_timer(self, minutes: int):
        if self.timer and self.timer.is_running: self.timer.stop()
        def on_tick(remaining):
            if self.window:
                m, s = divmod(remaining, 60)
                self.window.evaluate_js(f'updateTimer("{m:02d}:{s:02d}")')
        def on_remind(msg):
            if self.window: self.window.evaluate_js(f'showReminder("{msg}")')
        def on_end():
            if self.window: self.window.evaluate_js('showReminder("재촬영 시간이 종료되었습니다.")')
        self.timer = SessionTimer(minutes, on_tick, on_remind, on_end)
        self.timer.start()

    def _stop_timer(self):
        if self.timer: self.timer.stop(); self.timer = None


# =============================================================================
# Main
# =============================================================================

def main():
    config = ConfigManager()
    actions = MacroActions(config)
    api = Api(actions, config)
    
    ui_path = BASE_DIR / "ui" / "index.html"
    if not ui_path.exists():
        sys.exit(1)
    
    gui = config.get("gui_settings", {})
    screens = webview.screens
    target_screen = screens[0]
    if len(screens) > 1: target_screen = screens[1]
    
    window = webview.create_window(
        title=APP_NAME,
        url=str(ui_path),
        js_api=api,
        width=gui.get("width", 1920),
        height=gui.get("height", 1280),
        screen=target_screen,
        fullscreen=gui.get("fullscreen", True),
        frameless=True,
        background_color='#0a0a0f'
    )
    
    api.set_window(window)
    webview.start(debug=False)

if __name__ == "__main__":
    main()
