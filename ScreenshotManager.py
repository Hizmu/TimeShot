# ScreenshotManager.py

from datetime import datetime as dt
import os
import shutil
import zipfile
import pyautogui


class ScreenshotManager:
    def __init__(self, screenshort_root, archive_root, is_archive=False, is_delete=False):
        self.screenshort_root = screenshort_root
        self.archive_root = archive_root
        self.is_archive = is_archive
        self.is_delete = is_delete
        self.autostart = False
        self._lastfolders = []
        self.last_date = dt.now().date()
        self.update_paths()

    def update_paths(self):
        self.last_date = dt.now().date()
        name_dir = self.last_date.strftime('%Y-%m-%d')
        self.screenshot_path = os.path.join(self.screenshort_root, name_dir)

    def make_screenshot(self):
        screenshot = pyautogui.screenshot()
        current_date = dt.now()
        self._make_dir(self.screenshot_path)

        screenshot.save(os.path.join(self.screenshot_path, f'{current_date.time().strftime("%H-%M-%S")}.png'),
                        bitmap_format='png', optimize=True)

    def archive_and_delete(self):
        if self.is_archive:
            self._make_dir(self.archive_root)
            with zipfile.ZipFile(os.path.join(self.archive_path, f'{self.last_date.strftime("%Y-%m-%d")}.zip'), 'w',
                                 zipfile.ZIP_BZIP2) as zipf:
                self._zip_directory(self.screenshot_path, zipf)

        if self.is_delete:
            test = self._lastfolders.pop(0)
            shutil.rmtree(test)

    def _make_dir(self, directory):
        if not (directory in self._lastfolders):
            self._lastfolders.append(directory)
        if not os.path.exists(directory):
            os.makedirs(directory)

    def _zip_directory(self, path, zipf):
        for root, dirs, files in os.walk(path):
            for file in files:
                zipf.write(os.path.join(root, file),
                           os.path.relpath(os.path.join(root, file),
                                           os.path.join(path, '..')))
    
        
