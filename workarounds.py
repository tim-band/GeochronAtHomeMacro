# Library of workarounds to Zen macro difficulties
# Usage:
#
# import workarounds
# reload(workarounds)
#
# zen = workarounds.Workarounds(Zen, Zeiss)
#

import os
import random

class UserCancelledException(Exception):
    pass

class Workarounds:
    def __init__(self, zen, zeiss):
        self.zen = zen
        self.zeiss = zeiss
        self.backlash = -5

    def move_stage(self, x=None, y=None, z=None):
        """
        Move the stage and/or focus with backlash correction
        """
        if x != None or y != None:
            if x == None:
                x = self.zen.Devices.Stage.ActualPositionX
                xb = x
            else:
                xb = x + self.backlash
            if y == None:
                y = self.zen.Devices.Stage.ActualPositionY
                yb = y
            else:
                yb = y + self.backlash
            self.zen.Devices.Stage.MoveTo(xb, yb)
            self.zen.Devices.Stage.MoveTo(x, y)
        if z != None:
            self.zen.Devices.Focus.MoveTo(z + self.backlash)
            self.zen.Devices.Focus.MoveTo(z)

    def set_focus(self, z):
        self.move_stage(z=z)

    def show_window(self, window):
        """
        Show the window and throw an exception if the user cancels
        """
        r = window.Show()
        if r.HasCanceled:
            raise UserCancelledException("Operation cancelled")
        return r

    def switch_to_locate_tab(self):
        self.zen.Application.LeftToolArea.SwitchToTab(
            self.zeiss.Micro.Scripting.Research.ZenToolTab.LocateTab
        )

    def set_hardware(self, hardware_setting=None, camera_setting=None):
        """
        Set hardware and/or camera setting
        """
        self.switch_to_locate_tab()
        if type(hardware_setting) is str:
            hs = self.zen.Application.HardwareSettings.GetByName(hardware_setting)
            if hs == None:
                raise Exception('No such harware setting {0}'.format(hardware_setting))
            hardware_setting = hs
        if hardware_setting != None:
            self.zen.Devices.ApplyHardwareSetting(hardware_setting)
        if type(camera_setting) == str:
            cs = self.zen.Acquisition.CameraSettings.GetByName(camera_setting)
            if cs == None:
                raise Exception('No such camera setting {0}'.format(camera_setting))
            camera_setting = cs
        if camera_setting != None:
            self.zen.Acquisition.ActiveCamera.ApplyCameraSetting(camera_setting)

    def show_live(self):
        self.switch_to_locate_tab()
        live = self.zen.Acquisition.StartLive()
        self.zen.Application.Documents.ActiveDocument = live
        return live

    def autofocus(self, experiment, timeoutSeconds=30):
        """
        Autofocus using the settings in the experiment passed
        """
        self.switch_to_locate_tab()
        if type(experiment) is str:
            e = self.zeiss.Micro.Scripting.ZenExperiment(experiment)
            if e == None:
                raise Exception('No such acquisition experiment {0}'.format(experiment))
            experiment = e
        self.show_live()
        result = self.zen.Acquisition.FindAutofocus(experiment, timeoutSeconds=timeoutSeconds)

    def discard_changes(self, e):
        """
        Discard a document without getting input from the user
        """
        dummy_name = 'dummy{0:04d}'.format(random.randint(0,9999))
        e.SaveAs(dummy_name)
        e.Delete()
        e.Close()

    def save_as(self, zen_image, path, overwrite=False):
        """
        Saves the image with the given file path.
        Returns True on success, False if the file already exists and
        overwrite is False, throws an exception on any other error
        """
        if os.path.exists(path):
            if not overwrite:
                return False
            os.unlink(path)
        if not zen_image.Save():
            return False
        src = zen_image.FileName
        zen_image.Close()
        os.rename(src, path)
        return True
