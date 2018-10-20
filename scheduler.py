
# To register the addin:
#   scheduler.py
# This will install the COM server, and write the necessary
# AddIn key to Outlook
#
# To unregister completely:
#   scheduler.py --unregister
#
# To debug, execute:
#   scheduler.py --debug
#
# Then open Pythonwin, and select "Tools->Trace Collector Debugging Tool"
# Restart Outlook, and you should see some output generated.
#
# NOTE: If the AddIn fails with an error, Outlook will re-register
# the addin to not automatically load next time Outlook starts. To
# correct this, simply re-register the addin (see above)


from win32com import universal
from win32com.server.exception import COMException
from win32com.client import gencache, DispatchWithEvents
import winerror
import pythoncom
from win32com.client import constants, Dispatch
from win32com.server.util import wrap, unwrap
import win32ui
import win32con
import sys
import os
import gui
import outlook


gencache.EnsureModule('{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}', 0, 2, 8, bForDemand=True) # office 16
gencache.EnsureModule('{00062FFF-0000-0000-C000-000000000046}', 0, 9, 6) # outlook 16

universal.RegisterInterfaces('{AC0714F2-3D04-11D1-AE7D-00A0C90F26F4}', 0, 1, 0, ["_IDTExtensibility2"])
try:
    universal.RegisterInterfaces('{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}', 0, 2, 8, ["IRibbonExtensibility", "IRibbonControl"])
except:
    pass

class OutlookAddin:
    _com_interfaces_ = ["{B65AD801-ABAF-11D0-BB8B-00A0C90F2744}"]
    _com_interfaces_.append("{000C0396-0000-0000-C000-000000000046}")
    _public_methods_ = ['SchedulerSettings', 'SchedulerCall']
    # pythoncom.CreateGuid()
    _reg_clsid_ = '{9F4796B5-536F-47A4-9CB7-B6AFE5C10B2A}'
    _reg_progid_ = "Python.Scheduler"


    def __init__(self):
        self.appHostApp = None     

    def OnConnection(self, application, connectMode, addin, custom):
        try:
            self.appHostApp = application
        except Exception as e:
            win32ui.MessageBox (str (e), 'Python error', win32con. MB_OKCANCEL)

    def OnDisconnection(self, mode, custom):
        print ("OnDisconnection")
        self.appHostApp=None

    def OnAddInsUpdate(self, custom):
        print ("OnAddInsUpdate", custom)

    def OnStartupComplete(self, custom):
        print ("OnStartupComplete", custom)

    def OnBeginShutdown(self, custom):
        print ("OnBeginShutdown", custom)


    # xml for adding ribbon panels
    def GetCustomUI (self, arg):
        xml = """
            <customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
                <ribbon>
                    <tabs>
                        <tab idMso="TabMail">
                            <group id="Scheduler" label="Scheduler">
                                <button id="SchedulerSettings" label="Settings" onAction="SchedulerSettings" imageMso="CurrentViewSettings" tag="SchedulerSettings" size="large" enabled="true"/>
                                <button id="SchedulerCall" label="Schedule" onAction="SchedulerCall" imageMso="ArrangeByDate" tag="SchedulerCall" size="large" enabled="true"/>
                            </group>
                        </tab>
                    </tabs>
                </ribbon>
            </customUI>
        """
        #print xml
        return xml

    def SchedulerSettings(self, a1):
        try:
            gui.show_window()
        except pythoncom.com_error (hr, msg):
            print ("The Scheduler call failed with code %d: %s" % (hr, msg))

    def SchedulerCall(self, a1):
        try:
            outlook.get_availability()
        except pythoncom.com_error (hr, msg):
            print ("The Scheduler call failed with code %d: %s" % (hr, msg))

def RegisterAddin(klass):
    import winreg
    key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, "Software\\Microsoft\\Office\\Outlook\\Addins")
    subkey = winreg.CreateKey(key, klass._reg_progid_)
    winreg.SetValueEx(subkey, "CommandLineSafe", 0, winreg.REG_DWORD, 0)
    winreg.SetValueEx(subkey, "LoadBehavior", 0, winreg.REG_DWORD, 3)
    winreg.SetValueEx(subkey, "Description", 0, winreg.REG_SZ, "Scheduler Outlook Addin")
    winreg.SetValueEx(subkey, "FriendlyName", 0, winreg.REG_SZ, "Scheduler Outlook Addin")

def UnregisterAddin(klass):
    import winreg
    try:
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, "Software\\Microsoft\\Office\\Outlook\\Addins\\" + klass._reg_progid_)
    except WindowsError:
        pass

if __name__ == '__main__':
    import win32com.server.register
    win32com.server.register.UseCommandLine(OutlookAddin)
    if "--unregister" in sys.argv:
        UnregisterAddin(OutlookAddin)
    else:
        RegisterAddin(OutlookAddin)
        

