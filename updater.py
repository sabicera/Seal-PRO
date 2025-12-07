import requests
import json
import os
import sys
import subprocess
from pathlib import Path
from packaging import version


class AutoUpdater:
    def __init__(self, current_version, repo_owner, repo_name):
        self.current_version = current_version
        self.repo_owner = repo_owner
        self.repo_name = repo_name
        self.github_api_url = f"https://api.github.com/repos/{repo_owner}/{repo_name}/releases/latest"
        
    def check_for_updates(self):
        try:
            response = requests.get(self.github_api_url, timeout=5)
            if response.status_code == 200:
                release_data = response.json()
                latest_version = release_data['tag_name'].lstrip('v')
                
                # Compare versions
                if version.parse(latest_version) > version.parse(self.current_version):
                    # Find the exe asset
                    download_url = None
                    for asset in release_data.get('assets', []):
                        if asset['name'].endswith('.exe'):
                            download_url = asset['browser_download_url']
                            break
                    
                    release_notes = release_data.get('body', 'No release notes available')
                    
                    return True, latest_version, download_url, release_notes
                else:
                    return False, latest_version, None, None
            else:
                return False, None, None, None
                
        except Exception as e:
            print(f"Update check failed: {e}")
            return False, None, None, None
    
    def download_update(self, download_url, progress_callback=None):
        try:
            if getattr(sys, 'frozen', False):
                # Running as exe
                app_dir = Path(sys._MEIPASS).parent
            else:
                # Running as script
                app_dir = Path(__file__).parent
            
            download_path = app_dir / "SealCheckConverter_new.exe"
            
            response = requests.get(download_url, stream=True, timeout=30)
            total_size = int(response.headers.get('content-length', 0))
            downloaded = 0
            
            with open(download_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
                        downloaded += len(chunk)
                        if progress_callback and total_size > 0:
                            progress = int((downloaded / total_size) * 100)
                            progress_callback(progress)
            
            return download_path
            
        except Exception as e:
            print(f"Download failed: {e}")
            return None
    
    def apply_update(self, new_exe_path):
        try:
            if getattr(sys, 'frozen', False):
                current_exe = Path(sys.executable)
                old_exe = current_exe.parent / f"{current_exe.stem}_old.exe"
                
                # Create update script
                update_script = current_exe.parent / "update.bat"
                
                script_content = f"""@echo off
timeout /t 2 /nobreak > nul
del /f /q "{old_exe}"
move /y "{current_exe}" "{old_exe}"
move /y "{new_exe_path}" "{current_exe}"
start "" "{current_exe}"
del "%~f0"
"""
                
                with open(update_script, 'w') as f:
                    f.write(script_content)
                
                # Run update script and exit
                subprocess.Popen([str(update_script)], shell=True)
                sys.exit(0)
                
        except Exception as e:
            print(f"Update failed: {e}")
            return False