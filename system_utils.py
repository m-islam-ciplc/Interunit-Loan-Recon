"""
System Utilities for Sleep Prevention and Process Monitoring
Handles PC sleep prevention, progress checkpointing, and auto-resume functionality
"""

import os
import sys
import json
import time
import ctypes
import platform
from pathlib import Path
from datetime import datetime
from typing import Dict, Any, Optional, Tuple
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


class ProgressCheckpointer:
    """Handles progress checkpointing and auto-resume functionality"""
    
    def __init__(self, checkpoint_dir: str = "checkpoints"):
        self.checkpoint_dir = Path(checkpoint_dir)
        self.checkpoint_dir.mkdir(exist_ok=True)
        self.current_checkpoint = None
        self.checkpoint_interval = 30  # Save checkpoint every 30 seconds
        
    def create_checkpoint(self, 
                         step: str, 
                         progress: int, 
                         matches_found: int,
                         file1_path: str,
                         file2_path: str,
                         matches: list = None,
                         statistics: dict = None) -> str:
        """Create a progress checkpoint"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        checkpoint_id = f"checkpoint_{timestamp}_{step.lower().replace(' ', '_')}"
        
        checkpoint_data = {
            "checkpoint_id": checkpoint_id,
            "timestamp": timestamp,
            "step": step,
            "progress": progress,
            "matches_found": matches_found,
            "file1_path": file1_path,
            "file2_path": file2_path,
            "matches": matches or [],
            "statistics": statistics or {},
            "system_info": {
                "platform": platform.system(),
                "python_version": sys.version,
                "process_id": os.getpid()
            }
        }
        
        checkpoint_file = self.checkpoint_dir / f"{checkpoint_id}.json"
        
        try:
            with open(checkpoint_file, 'w', encoding='utf-8') as f:
                json.dump(checkpoint_data, f, indent=2, default=str)
            
            self.current_checkpoint = checkpoint_id
            return checkpoint_id
            
        except Exception as e:
            print(f"Warning: Could not create checkpoint: {e}")
            return None
    
    def get_latest_checkpoint(self) -> Optional[Dict[str, Any]]:
        """Get the most recent checkpoint"""
        try:
            checkpoint_files = list(self.checkpoint_dir.glob("checkpoint_*.json"))
            if not checkpoint_files:
                return None
            
            # Sort by modification time, get the latest
            latest_file = max(checkpoint_files, key=lambda f: f.stat().st_mtime)
            
            with open(latest_file, 'r', encoding='utf-8') as f:
                return json.load(f)
                
        except Exception as e:
            print(f"Warning: Could not read checkpoint: {e}")
            return None
    
    def cleanup_old_checkpoints(self, keep_last: int = 3):
        """Clean up old checkpoints, keeping only the most recent ones"""
        try:
            checkpoint_files = list(self.checkpoint_dir.glob("checkpoint_*.json"))
            if len(checkpoint_files) <= keep_last:
                return
            
            # Sort by modification time, keep the most recent
            checkpoint_files.sort(key=lambda f: f.stat().st_mtime, reverse=True)
            
            for old_file in checkpoint_files[keep_last:]:
                old_file.unlink()
                
        except Exception as e:
            print(f"Warning: Could not cleanup old checkpoints: {e}")
    
    def delete_checkpoint(self, checkpoint_id: str):
        """Delete a specific checkpoint"""
        try:
            checkpoint_file = self.checkpoint_dir / f"{checkpoint_id}.json"
            if checkpoint_file.exists():
                checkpoint_file.unlink()
        except Exception as e:
            print(f"Warning: Could not delete checkpoint: {e}")


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


class AutoResumeManager:
    """Manages auto-resume functionality"""
    
    def __init__(self, checkpointer: ProgressCheckpointer):
        self.checkpointer = checkpointer
        self.resume_data = None
    
    def check_for_resume(self) -> Tuple[bool, Optional[Dict[str, Any]]]:
        """Check if there's a checkpoint to resume from"""
        checkpoint = self.checkpointer.get_latest_checkpoint()
        
        if not checkpoint:
            return False, None
        
        # Check if checkpoint is recent (within last 24 hours)
        checkpoint_time = datetime.strptime(checkpoint['timestamp'], "%Y%m%d_%H%M%S")
        time_diff = datetime.now() - checkpoint_time
        
        if time_diff.total_seconds() > 86400:  # 24 hours
            print("Checkpoint is too old, not resuming")
            return False, None
        
        self.resume_data = checkpoint
        return True, checkpoint
    
    def get_resume_info(self) -> Optional[Dict[str, Any]]:
        """Get information about the checkpoint to resume from"""
        if not self.resume_data:
            return None
        
        return {
            "step": self.resume_data['step'],
            "progress": self.resume_data['progress'],
            "matches_found": self.resume_data['matches_found'],
            "timestamp": self.resume_data['timestamp'],
            "file1": os.path.basename(self.resume_data['file1_path']),
            "file2": os.path.basename(self.resume_data['file2_path'])
        }
    
    def clear_resume_data(self):
        """Clear resume data after successful completion"""
        if self.resume_data:
            self.checkpointer.delete_checkpoint(self.resume_data['checkpoint_id'])
            self.resume_data = None
