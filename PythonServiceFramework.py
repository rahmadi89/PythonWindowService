"""
run cmd as administrator and then run these commands:

mkdir C:\pyvenv\

cd c:\pyvenv\

python -m venv .

Scripts\activate.bat

pip install pywin32 pyinstaller 

pyinstaller --onefile --hidden-import win32timezone PythonServiceFramework.py

dist\PythonServiceFramework.exe install

dist\PythonServiceFramework.exe start

dist\PythonServiceFramework.exe stop

"""
import win32serviceutil
import win32service
import win32event
import servicemanager
import sys
from abc import ABC, abstractmethod

class PythonServiceFramework(win32serviceutil.ServiceFramework, ABC):
    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        # Create an event which we will use to wait on.
        # The "service stop" request will set this event.
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)

    @abstractmethod
    def initialize(self):
        pass

    @abstractmethod
    def stop(self):
        pass

    @abstractmethod
    def main(self):
        pass
    
    def SvcStop(self):
        """ Called when the service asked to stop"""
        # Do something just before stop
        self.stop()
        # Let the SCM be aware that the service is stopping
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        # Set Event to wait stop
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        """ Called when the service ask to run."""
        # Let the SCM be aware that the service is starting
        self.ReportServiceStatus(win32service.SERVICE_START_PENDING)
        # Log to Event Viewer (Necessary)
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,servicemanager.PYS_SERVICE_STARTED,(self._svc_name_,''))
        # Do Something at service initialize
        self.initialize()
        # Let the SCM be aware that the service is started and running
        self.ReportServiceStatus(win32service.SERVICE_RUNNING)
        # Do the main work
        self.main()
        
        
class TestService(PythonServiceFramework):
    # Define the class variables
    _svc_name_ = "TimesheetService"
    _svc_display_name_ = "Timesheet Service"
    _svc_description_ = "Manage Timesheet. Register Tasks."
    
    def __init__(self,args):
        super().__init__(args)
        self.isRunning = None

    def initialize(self):
        self.isRunning = True
    
    def stop(self):
        self.isRunning = False
        
    def main(self): 
        while self.isRunning:
            f = open('D:\\test.txt', 'a')
            rc = None
            while rc != win32event.WAIT_OBJECT_0:
                f.write('Service Started  \n')
                f.flush()
                #block for 24*60*60 seconds and wait for a stop event
                #it is used for a one-day loop
                rc = win32event.WaitForSingleObject(self.hWaitStop, 24*60*60*1000)
            f.write('Service Stopped \n')
            f.close()
    
if __name__ == '__main__':
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(TestService)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(TestService)