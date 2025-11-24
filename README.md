# close_popup_window - use python to close specific pop up window and create the log
"""
Auto Close Error Dialog Program
Specifically designed for Windows Script Host and system error dialogs
"""

import pygetwindow as gw
import pyautogui
import time
from datetime import datetime
import os
import win32gui
import win32con

# Disable PyAutoGUI fail-safe (recommended for automated scripts)
pyautogui.FAILSAFE = False

# ============================================
# 📝 Only change these settings!
# ============================================
# Window names to monitor (can add multiple)
window_names = [
    "Microsoft Dynamics NAV Classic"
]

# Keywords to search for (if ANY keyword is found, close the window)
keyword_list = [
    "deadlocked",
    "locked",
    "Error",
    "interrupted",
    "timeout"
]

check_interval_seconds = 60        # Check every 60 seconds (fast response)

# Also check window content (not just title)
check_window_content = True
# ============================================

# Create log folder
if not os.path.exists("nav_error_logs"):
    os.makedirs("nav_error_logs")

def write_log(message):
    """Write message to log file"""
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_content = f"[{current_time}] {message}\n"
    
    # Display on screen
    print(log_content.strip())
    
    # Write to file
    today_date = datetime.now().strftime("%Y-%m-%d")
    log_filename = f"nav_error_logs/log_{today_date}.txt"
    
    with open(log_filename, "a", encoding="utf-8") as f:
        f.write(log_content)

def get_window_text_win32(hwnd):
    """Get window text using Win32 API"""
    try:
        length = win32gui.GetWindowTextLength(hwnd)
        if length > 0:
            return win32gui.GetWindowText(hwnd)
    except:
        pass
    return ""

def enum_child_windows(hwnd):
    """Get all text from child windows (for dialog boxes)"""
    texts = []
    
    def callback(child_hwnd, data):
        text = get_window_text_win32(child_hwnd)
        if text and text.strip() and len(text) > 1:
            # Filter out button text if too short
            if text not in ["OK", "Cancel", "Yes", "No", "確定", "取消"]:
                data.append(text)
        return True
    
    try:
        win32gui.EnumChildWindows(hwnd, callback, texts)
    except:
        pass
    
    return texts

def get_all_window_text(hwnd):
    """Get all text from window and children"""
    all_texts = []
    
    # Get main window text
    main_text = get_window_text_win32(hwnd)
    if main_text:
        all_texts.append(main_text)
    
    # Get all child window texts (this works great for dialog boxes!)
    child_texts = enum_child_windows(hwnd)
    all_texts.extend(child_texts)
    
    return all_texts

def check_and_close_window():
    """Check and close window if keyword found"""
    try:
        # Get all visible windows
        def enum_windows_callback(hwnd, results):
            if win32gui.IsWindowVisible(hwnd):
                window_title = get_window_text_win32(hwnd)
                if window_title:
                    results.append((hwnd, window_title))
            return True
        
        windows = []
        win32gui.EnumWindows(enum_windows_callback, windows)
        
        for hwnd, window_title in windows:
            # Check if window title matches any of our target window names
            is_target_window = False
            matched_window_name = None
            
            for target_name in window_names:
                if target_name.lower() in window_title.lower():
                    is_target_window = True
                    matched_window_name = target_name
                    break
            
            if not is_target_window:
                continue
            
            # Get all text from the window
            all_texts = get_all_window_text(hwnd)
            combined_text = " ".join(all_texts)
            
            # Check if any keyword exists
            found_keyword = None
            search_location = "title"
            
            # Check in title first
            for keyword in keyword_list:
                if keyword.lower() in window_title.lower():
                    found_keyword = keyword
                    break
            
            # Check in content
            if not found_keyword and check_window_content:
                for keyword in keyword_list:
                    if keyword.lower() in combined_text.lower():
                        found_keyword = keyword
                        search_location = "content"
                        break
            
            # If keyword found, close the window and LOG it
            if found_keyword:
                # Write to log (ONLY when closing)
                write_log("=" * 60)
                write_log(f"CLOSED WINDOW")
                write_log(f"Window Title: {window_title}")
                write_log(f"Matched Window Name: {matched_window_name}")
                write_log(f"Triggered Keyword: '{found_keyword}' (found in {search_location})")
                write_log(f"Window Content: {combined_text[:500]}")
                
                # Bring window to front
                try:
                    win32gui.SetForegroundWindow(hwnd)
                    time.sleep(0.3)
                except:
                    pass
                
                # Press Enter to close (OK button)
                pyautogui.press('enter')
                write_log(f"✓ Pressed Enter to close window")
                write_log("=" * 60)
                write_log("")
                
                time.sleep(1)
                return True
            # If NO keyword found, do NOT log (skip silently)
        
        return False
        
    except Exception as e:
        write_log(f"Error occurred: {e}")
        import traceback
        write_log(f"Details: {traceback.format_exc()}")
        return False

def main():
    """Run 24/7"""
    write_log("=" * 60)
    write_log("Auto Close Error Dialog Program Started!")
    write_log(f"Monitoring windows: {', '.join(window_names)}")
    write_log(f"Monitoring keywords: {', '.join(keyword_list)}")
    write_log(f"Check interval: {check_interval_seconds} seconds")
    write_log(f"Check window content: {check_window_content}")
    write_log("=" * 60)
    write_log("Program is running. Press Ctrl+C to stop.")
    write_log("")
    
    check_count = 0
    close_count = 0
    
    try:
        while True:
            check_count += 1
            
            # Report status every 100 checks
            if check_count % 100 == 0:
                write_log(f"Status: Checked {check_count} times, Closed {close_count} windows")
            
            # Check window
            if check_and_close_window():
                close_count += 1
            
            # Sleep to save CPU
            time.sleep(check_interval_seconds)
            
    except KeyboardInterrupt:
        write_log("")
        write_log("=" * 60)
        write_log("Program stopped by user")
        write_log(f"Total checks: {check_count}")
        write_log(f"Total windows closed: {close_count}")
        write_log("=" * 60)

if __name__ == "__main__":
    main()
