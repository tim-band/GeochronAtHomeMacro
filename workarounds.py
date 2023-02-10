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
        self.duck_z = 500
        self.max_move_without_ducking = 500

    def move_stage(self, x=None, y=None, z=None):
        """
        Move the stage and/or focus with backlash correction,
        and trying not to crash sideways into things
        """
        current_x = self.zen.Devices.Stage.ActualPositionX
        current_y = self.zen.Devices.Stage.ActualPositionY
        current_z = self.zen.Devices.Focus.ActualPosition
        if z == None:
            z = current_z
            zb = z
            minz = z
        else:
            minz = min(current_z, z)
            zb = minz - abs(self.backlash)
        if x == None:
            x = current_x
            xb = x
        else:
            xb = x + self.backlash
            if self.max_move_without_ducking < abs(x - current_x):
                zb = minz - self.duck_z
        if y == None:
            y = current_y
            yb = y
        else:
            yb = y + self.backlash
            if self.max_move_without_ducking < abs(y - current_y):
                zb = minz - self.duck_z
        self.zen.Devices.Focus.MoveTo(zb)
        self.zen.Devices.Stage.MoveTo(xb, yb)
        self.zen.Devices.Stage.MoveTo(x, y)
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

    def switch_to_acquisition_tab(self):
        self.zen.Application.LeftToolArea.SwitchToTab(
            self.zeiss.Micro.Scripting.Research.ZenToolTab.AcquisitionTab
        )

    def set_hardware(self, hardware_setting=None, camera_setting=None):
        """
        Set hardware and/or camera setting
        """
        self.switch_to_locate_tab()
        if type(hardware_setting) is str:
            hs = self.zen.Application.HardwareSettings.GetByName(hardware_setting)
            if hs == None:
                print dir(self.zeiss.Micro.Scripting.ZenHardwareSetting)
                print hardware_setting
                hs = self.zeiss.Micro.Scripting.ZenHardwareSetting.__class__(hardware_setting)
                if hs == None:
                    raise Exception('No such harware setting {0}'.format(hardware_setting))
            hardware_setting = hs
        if hardware_setting != None:
            self.zen.Devices.ApplyHardwareSetting(hardware_setting)
        if type(camera_setting) == str:
            cs = self.zen.Acquisition.CameraSettings.GetByName(camera_setting)
            if cs == None:
                cs = self.zeiss.Micro.Scripting.ZenCameraSetting.__class__(camera_setting)
                if cs == None:
                    raise Exception('No such camera setting {0}'.format(camera_setting))
            camera_setting = cs
        if camera_setting != None:
            self.zen.Acquisition.ActiveCamera.ApplyCameraSetting(camera_setting)

    def get_experiment(self, experiment):
        if type(experiment) is str:
            e = self.zeiss.Micro.Scripting.ZenExperiment(experiment)
            if e == None:
                raise Exception('No such acquisition experiment {0}'.format(experiment))
            return e
        return experiment

    def show_live(self, experiment=None):
        self.switch_to_locate_tab()
        live = self.zen.Acquisition.StartLive()
        if experiment is not None:
            self.switch_to_acquisition_tab()
            live = self.zen.Acquisition.StartLive(
                self.get_experiment(experiment)
            )
        self.zen.Application.Documents.ActiveDocument = live
        return live

    def autofocus(self, experiment, timeoutSeconds=30):
        """
        Autofocus using the settings in the experiment passed
        """
        e = self.get_experiment(experiment)
        self.show_live(e)
        self.zen.Acquisition.FindAutofocus(e, timeoutSeconds=timeoutSeconds)

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
