"""
System Utilities for Sleep Prevention and Process Monitoring
Handles PC sleep prevention and process monitoring
"""

import os
import sys
import time
import ctypes
import platform
from pathlib import Path
import threading


class SystemSleepPrevention:
    """Prevents system sleep during long-running processes"""
    
    def __init__(self):
        self.is_preventing_sleep = False
        self._prevention_thread = None
        self._stop_prevention = threading.Event()
        
    def start_preventing_sleep(self, reason: str = "Processing Excel files"):
        """Start preventing system sleep"""
        if self.is_preventing_sleep:
            return
            
        self.is_preventing_sleep = True
        self._stop_prevention.clear()
        
        if platform.system() == "Windows":
            self._prevention_thread = threading.Thread(
                target=self._windows_sleep_prevention,
                args=(reason,),
                daemon=True
            )
        else:
            # For Linux/Mac, we'll use a simple keep-alive approach
            self._prevention_thread = threading.Thread(
                target=self._generic_sleep_prevention,
                daemon=True
            )
        
        self._prevention_thread.start()
    
    def stop_preventing_sleep(self):
        """Stop preventing system sleep"""
        if not self.is_preventing_sleep:
            return
            
        self.is_preventing_sleep = False
        self._stop_prevention.set()
        
        if self._prevention_thread and self._prevention_thread.is_alive():
            self._prevention_thread.join(timeout=2.0)
    
    def _windows_sleep_prevention(self, reason: str):
        """Windows-specific sleep prevention using SetThreadExecutionState"""
        try:
            # ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_AWAYMODE_REQUIRED
            ES_CONTINUOUS = 0x80000000
            ES_SYSTEM_REQUIRED = 0x00000001
            ES_AWAYMODE_REQUIRED = 0x00000040
            
            flags = ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_AWAYMODE_REQUIRED
            
            while not self._stop_prevention.is_set():
                ctypes.windll.kernel32.SetThreadExecutionState(flags)
                time.sleep(30)  # Refresh every 30 seconds
                
        except Exception as e:
            print(f"Warning: Could not prevent system sleep: {e}")
        finally:
            # Restore normal power management
            try:
                ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)
            except:
                pass
    
    def _generic_sleep_prevention(self):
        """Generic sleep prevention for Linux/Mac"""
        try:
            while not self._stop_prevention.is_set():
                # Simple keep-alive by touching a temporary file
                temp_file = Path.home() / ".excel_matcher_keepalive"
                temp_file.touch()
                time.sleep(30)
        except Exception as e:
            print(f"Warning: Could not prevent system sleep: {e}")
        finally:
            # Clean up temporary file
            try:
                temp_file = Path.home() / ".excel_matcher_keepalive"
                if temp_file.exists():
                    temp_file.unlink()
            except:
                pass


class ProcessMonitor:
    """Monitors process health and detects interruptions"""
    
    def __init__(self):
        self.start_time = None
        self.last_activity_time = None
        self.is_monitoring = False
        self._monitor_thread = None
        self._stop_monitoring = threading.Event()
        
    def start_monitoring(self):
        """Start monitoring process health"""
        if self.is_monitoring:
            return
            
        self.is_monitoring = True
        self.start_time = time.time()
        self.last_activity_time = time.time()
        self._stop_monitoring.clear()
        
        self._monitor_thread = threading.Thread(
            target=self._monitor_loop,
            daemon=True
        )
        self._monitor_thread.start()
    
    def stop_monitoring(self):
        """Stop monitoring process health"""
        if not self.is_monitoring:
            return
            
        self.is_monitoring = False
        self._stop_monitoring.set()
        
        if self._monitor_thread and self._monitor_thread.is_alive():
            self._monitor_thread.join(timeout=2.0)
    
    def update_activity(self):
        """Update last activity time"""
        self.last_activity_time = time.time()
    
    def _monitor_loop(self):
        """Main monitoring loop"""
        while not self._stop_monitoring.is_set():
            try:
                current_time = time.time()
                
                # Check for long periods of inactivity (possible sleep)
                if self.last_activity_time:
                    inactive_time = current_time - self.last_activity_time
                    if inactive_time > 300:  # 5 minutes of inactivity
                        print(f"Warning: Process inactive for {inactive_time:.0f} seconds")
                
                time.sleep(10)  # Check every 10 seconds
                
            except Exception as e:
                print(f"Warning: Process monitoring error: {e}")
                time.sleep(30)
    
    def get_uptime(self) -> float:
        """Get process uptime in seconds"""
        if self.start_time:
            return time.time() - self.start_time
        return 0.0
    
    def get_inactive_time(self) -> float:
        """Get time since last activity in seconds"""
        if self.last_activity_time:
            return time.time() - self.last_activity_time
        return 0.0
