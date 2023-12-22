import tkinter as tk
from tkinter import filedialog
from ScreenshotManager import ScreenshotManager
from datetime import datetime as dtdt
from dateutil.relativedelta import relativedelta
import os
from ConfigManager import ConfigManager
from win32com.client import Dispatch
from datetime import datetime as dt
import sys
class TimeShot:
    def __init__(self):
        self.__config = ConfigManager()

        self.win = tk.Tk(baseName=sys.argv[0])
        self._last_date = dt.now().date()
        self.vcmd = (self.win.register(self.validate))
        self.setup_initial_ui()
        
    def validate(self, P):

        if P.isdigit():
            return True
        elif P == "":
            return False
        else:
            return False
        
    def setup_initial_ui(self):
        self.win.title('TimeShot')
        self.win.minsize(400, 200)
        self.win.maxsize(400, 200)  
        screenshot_root = self.__config['Dirs']['screenshot']
        archive_root = self.__config['Dirs']['arhive']

        self._screenshot_root_var = tk.StringVar(value=screenshot_root)
        self._archive_root_var = tk.StringVar(value=archive_root)

        self._screenshot_manager = ScreenshotManager(screenshot_root, archive_root)

        self.setup_variables()
        self.create_ui_elements()

    def setup_variables(self):
        self.is_archive_var = tk.BooleanVar(value=self.__config['DATA']['arhive'])
        self.is_delete_var = tk.BooleanVar(value=self.__config['DATA']['delete'])
        self.sc_btw_sc_var = tk.IntVar(value=self.__config['IntervalBetweenScreenshot']['seconds'])
        self.mn_btw_sc_var = tk.IntVar(value=self.__config['IntervalBetweenScreenshot']['minutes'])
        self.hr_btw_sc_var = tk.IntVar(value=self.__config['IntervalBetweenScreenshot']['hours'])
        self.days_to_del_var = tk.IntVar(value=self.__config['TimeToDelete']['days'])
        self.mth_to_del_var = tk.IntVar(value=self.__config['TimeToDelete']['months'])
        self.autostart_var = tk.BooleanVar(value=self.__config['DATA']['autostart'])
        self.tm_btw_screenshot = 5000
        self.dt_to_del = dt.now()
        self.path = self.__config['Dirs']['exe']
        self.is_started_var = tk.IntVar(value=0)

    def create_ui_elements(self):
        tk.Checkbutton(self.win, text='Archive', variable=self.is_archive_var, onvalue=1, offvalue=0,
                        command=self.is_checked_archive).place(x=5, y=37)
        tk.Checkbutton(self.win, text='Delete', variable=self.is_delete_var, onvalue=1, offvalue=0,
                        command=self.is_checked_delete).place(x=5, y=90)
        self.create_entry(self._screenshot_root_var, 5, 5, height=25, button_text='Screenshot Path',
                          button_command=self.read_screenshot_path)
        self.create_time_entry(self.hr_btw_sc_var, 'H', 150, 37)
        self.create_time_entry(self.mn_btw_sc_var, 'M', 180, 37)
        self.create_time_entry(self.sc_btw_sc_var, 's', 220, 37)
        self.create_entry(self._archive_root_var, 5, 120, height=25, button_text='Archive Path',
                          button_command=self.read_archive_path)
        self.create_time_entry(self.mth_to_del_var, 'M', 150, 92)
        self.create_time_entry(self.days_to_del_var, 'D', 180, 92)

        tk.Button(self.win, text='Save', command=self.save).place(x=300, y=116)

        tk.Checkbutton(self.win, text='Auto Start', variable=self.autostart_var, command=self.autostart).place(x=300, y=0)

        self.create_run_stop_buttons()

        if self.autostart_var.get():
            self.win.wm_state('iconic')
            self.run()

        self.win.mainloop()

    def create_entry(self, text_variable, x, y, height, button_text, button_command):
        entry = tk.Entry(self.win, textvariable=text_variable, state='readonly')
        entry.place(x=x, y=y, height=height)
        tk.Button(self.win, text=button_text, command=button_command).place(x=x + 145, y=y - 2)
        return entry

    def create_time_entry(self, text_variable, label_text, x, y):
        entry = tk.Entry(self.win, textvariable=text_variable, width=3,validate='all',validatecommand=(self.vcmd,'%P'))
        entry.place(x=x, y=y)
        tk.Label(self.win, text=label_text).place(x=x + 20, y=y - 2, width=10)
        return entry

    def create_run_stop_buttons(self):
        self.bt_run = tk.Button(self.win, text='Run', command=self.run)
        self.bt_run.place(x=5, y=160)
        self.bt_stop = tk.Button(self.win, text='Stop', command=self.stop, state=tk.DISABLED)
        self.bt_stop.place(x=150, y=160)

    def is_checked_archive(self):
        self._screenshot_manager.is_archive = self.is_archive_var.get()
        self.__config.set('DATA', 'arhive', self._screenshot_manager.is_archive)

    def is_checked_delete(self):
        self._screenshot_manager.is_delete = self.is_delete_var.get()
        self.__config.set('DATA', 'delete', self._screenshot_manager.is_delete)

    def autostart(self):
        self.__config.set('DATA', 'AutoStart', self.autostart_var.get())
        target = self.__config['Dirs']['exe'] + '\TimeShot.exe'
        path = os.path.expanduser('~') + R'\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\TimeShot.lnk'

        if self.autostart_var.get():
            shell = Dispatch('WScript.Shell')
            shortcut = shell.CreateShortCut(path)
            shortcut.Targetpath = target
            shortcut.save()
        else:
            os.remove(path)

    def save(self):
        self.__config.set('IntervalBetweenScreenshot', 'seconds', self.sc_btw_sc_var.get())
        self.__config.set('IntervalBetweenScreenshot', 'hours', self.hr_btw_sc_var.get())
        self.__config.set('IntervalBetweenScreenshot', 'minutes', self.mn_btw_sc_var.get())
        self.__config.set('TimeToDelete', 'days', self.days_to_del_var.get())
        self.__config.set('TimeToDelete', 'months', self.mth_to_del_var.get())
        self.__config.set('Dirs', 'screenshot', self._screenshot_manager.screenshort_root)
        self.__config.set('Dirs', 'arhive', self._screenshot_manager.archive_root)

    def read_screenshot_path(self):
        directory = filedialog.askdirectory()
        if directory:
            self._screenshot_root_var.set(directory)
            self._screenshot_manager.screenshort_root = directory

    def read_archive_path(self):
        directory = filedialog.askdirectory()
        if directory:
            self._archive_root_var.set(directory)
            self._screenshot_manager.archive_root = directory

    def run(self):
        self.stopped = False
        self._screenshot_manager.update_paths()
        self.tm_btw_screenshot = self.sc_btw_sc_var.get() * 1000 + self.mn_btw_sc_var.get() * 60 * 1000 + self.hr_btw_sc_var.get() * 3600 * 1000

        self.calc_delete_date()

        self.bt_run['state'] = tk.DISABLED
        self.bt_stop['state'] = tk.NORMAL
        self.loop()

    def calc_delete_date(self):
        date = dt.now()
        date = date + relativedelta(months=+self.mth_to_del_var.get(), days=+self.days_to_del_var.get())
        self.dt_to_del = date.date()

    def stop(self):
        self.stopped = True
        self.is_hide = False
        self.bt_run['state'] = tk.NORMAL
        self.bt_stop['state'] = tk.DISABLED

    def loop(self):
        if self.stopped:
            self.win.after_cancel(self.loop)
            return

        self._screenshot_manager.make_screenshot()
        if self._last_date != dt.now().date():
            self._screenshot_manager.update_paths()
            self._last_date = dt.now().date()

        if self._last_date == self.dt_to_del:
            self._screenshot_manager.archive_and_delete()
            self.calc_delete_date()

        self.win.after(self.tm_btw_screenshot, self.loop)

if __name__ == '__main__':
    app = TimeShot()
