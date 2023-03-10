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
import xml.etree.ElementTree as ET

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

def add_support_points(exp_file_name, points, interpolation_degree=2, region_name='TR1'):
    """
    exp_file_name is the full file name of an experiment, which should
    be closed.
    points is a sequence of tuples (x, y, z) in stage co-ordinates.
    interpolation_degree is 0 for horizontal, 1 for flat, 2, 3 or 4 for
    'parabolic'. Not sure Zeiss are correct about that.
    """
    tree = ET.parse(exp_file_name)
    root = tree.getroot()
    acq_block = root.find('./ExperimentBlocks/AcquisitionBlock[@IsActivated="true"]')
    sample_holder = acq_block.find('./SubDimensionSetups/RegionsSetup[@IsActivated="true"]/SampleHolder')
    global_interpolation_degree = sample_holder.find('./GlobalInterpolationExpansionDegree')
    local_interpolation_degree = sample_holder.find('./LocalInterpolationExpansionDegree')
    tile_region = sample_holder.find('./TileRegions/TileRegion[@Name="{0}"]'.format(region_name))
    focus_strategy_mode = acq_block.find('./HelperSetups/FocusSetup[@IsActivated="true"]/FocusStrategy[@IsActivated="true"]/StrategyMode')

    focus_strategy_mode.text = 'LocalFocusSurface'
    global_interpolation_degree.text = str(interpolation_degree)
    local_interpolation_degree.text = str(interpolation_degree)
    
    old_sps = tile_region.find('SupportPoints')
    if old_sps is not None:
        tile_region.remove(old_sps)
    support_points = ET.SubElement(tile_region, 'SupportPoints')

    [cx, cy] = map(float, tile_region.find('CenterPosition').text.split(','))
    [w, h] = map(float, tile_region.find('ContourSize').text.split(','))
    
    id_prefix = '{0:08d}'.format(random.randint(0,99999999))
    n = 0
    for (x, y, z) in points:
        sp = ET.SubElement(support_points, 'SupportPoint', {
            'Name': 'SP',
            'Id': '{0}{1:04d}'.format(id_prefix, n)
        })
        ET.SubElement(sp, 'X').text = str(x)
        ET.SubElement(sp, 'Y').text = str(y)
        ET.SubElement(sp, 'Z').text = str(z)
        n += 1
    tree.write(exp_file_name)
