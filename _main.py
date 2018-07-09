"""
    Vectorworks plugin for export patch list to Jands Vista Console
"""

__author__ = "Robert Janiak"
__contact__ = "robert@stagevision.pl"

import vs
import os
from dlg import *

POPUP_LIST = [kFixtureNamePopup, kFixtureNumberPopup, kDmxUniversePopup, kDmxAddressPopup]
DEFAULT_FIELD = ['Purpose', 'Channel', 'Universe', 'Dimmer']
DEFAULT_DEVICE_TYPE = ['Light', 'Moving Light']
PLUGIN_NAME = 'Lighting Device'

SETUP_DIALOG_C = 12255

# ----------- MAIN ----------- #
def execute():

    global file, dialog_id, plugin_handle

    # check if Lighting Device plugin existing
    plugin_handle = vs.GetObject(PLUGIN_NAME)

    device_list = []
    vs.ForEachObject(lambda h: device_list.append(h), "(PON='Lighting Device')")

    if device_list:
        # create dialog from external file
        dialog_id = CreateDialog()
        if vs.RunLayoutDialog(dialog_id, dialog_handler) == kOK:
            # get platform type for difference file directory style

            err, dir_path = vs.GetFolder('Select a Folder')
            file_name, _ = os.path.splitext(vs.GetFName())
            file_name = file_name + '.csv'

            if err == 0:
                # create file
                with open(dir_path + file_name, 'w') as file:
                    # main procedure for write data
                    for criteria in cCriteria:
                        vs.ForEachObject(do_it, criteria)

                vs.AlrtDialog('File: \n' + file_name + ' \n\nCreated in:\n' + dir_path)
    else:
        vs.AlrtDialog('Lighting Device not found')


def dialog_handler(item, _):
    # set data during initialize dialog
    if item == SETUP_DIALOG_C:

        global check_image

        param_list = []
        # create Lighting Device POI fields list
        for field in range(1, vs.NumFields(plugin_handle)+1):
            param = vs.GetFldName(plugin_handle, field)
            is_param, name = vs.GetLocalizedPluginParameter(PLUGIN_NAME, param)
            if is_param:
                if name[0] != '_':
                    param_list.append([param, name])

        # fill popup menus
        for p, popup in enumerate(POPUP_LIST):
            for f, field in enumerate(param_list):
                vs.AddChoice(dialog_id, popup, field[1], f)
            field_index = vs.GetChoiceIndex(dialog_id, popup, DEFAULT_FIELD[p])
            vs.SelectChoice(dialog_id, popup, field_index, True)

        # create device type list
        device_list = []
        device_id = 1
        while True:
            is_field, device_type = vs.GetLocalizedPluginChoice(PLUGIN_NAME, 'Device Type', device_id)
            if not is_field:
                break
            device_list.append(device_type)
            device_id += 1

        # insert list browser column
        vs.InsertLBColumn(dialog_id, kLightingTypesLB, 0, 'Device Types', 150)
        vs.InsertLBColumn(dialog_id, kLightingTypesLB, 1, '#', 50)

        # set list browser column type
        vs.SetLBControlType(dialog_id, kLightingTypesLB, 1, 3)
        vs.SetLBItemDisplayType(dialog_id, kLightingTypesLB, 1,1)

        # add list browser images
        check_image_path = 'Vectorworks/Standard Images/Checkmark.png'
        check_image = vs.AddListBrowserImage(dialog_id, kLightingTypesLB, check_image_path)
        blank_image_path = 'Vectorworks/Standard Images/Blank.png'
        blank_image = vs.AddListBrowserImage(dialog_id, kLightingTypesLB, blank_image_path)

        # insert list browser column data
        vs.InsertLBColumnDataItem(dialog_id, kLightingTypesLB, 1, 'YES', check_image, -1, 0)
        vs.InsertLBColumnDataItem(dialog_id, kLightingTypesLB, 1, 'NO', blank_image, -1, 0)

        # enable list browser lines
        vs.EnableLBColumnLines(dialog_id, kLightingTypesLB, True)

        # insert lighting device types to list browser and check defaults
        for i, device_type in enumerate(device_list):
            vs.InsertLBItem(dialog_id, kLightingTypesLB, i, device_type)
            if device_type in DEFAULT_DEVICE_TYPE:
                vs.SetLBItemUsingColumnDataItem(dialog_id, kLightingTypesLB, i, 1, check_image)

    # get data after OK click
    if item == kOK:

        global name_field, number_field, universe_field, address_field, cCriteria

        # get name, number, universe, address name field
        name_field = vs.GetItemText(dialog_id, kFixtureNamePopup)
        number_field = vs.GetItemText(dialog_id, kFixtureNumberPopup)
        universe_field = vs.GetItemText(dialog_id, kDmxUniversePopup)
        address_field = vs.GetItemText(dialog_id, kDmxAddressPopup)

        # get list browser data and create criteria list
        cCriteria = []
        for row in range(vs.GetNumLBItems(dialog_id, kLightingTypesLB)):
            _, _, image = vs.GetLBItemInfo(dialog_id, kLightingTypesLB, row, 1)
            if image == check_image:
                _, data, _ = vs.GetLBItemInfo(dialog_id, kLightingTypesLB, row, 0)
                cCriteria.append("('Lighting Device'.'Device Type'='" + data + "' )")

    return item

# check if value is integer number
def is_number(s):
    try:
        int(s)
        return True
    except ValueError:
        return False


# callback for vs.ForEachObject() procedure
def do_it(h): 
    # get fixture record data
    inst_type = vs.GetRField(h, 'Lighting Device', 'Inst Type')
    fixture_mode = vs.GetRField(h, 'Lighting Device', 'Fixture Mode')
    channel = vs.GetRField(h, 'Lighting Device', number_field)
    universe = vs.GetRField(h, 'Lighting Device', universe_field)
    address = vs.GetRField(h, 'Lighting Device', address_field)
    purpose = vs.GetRField(h, 'Lighting Device', name_field)
    uid = vs.GetRField(h, 'Lighting Device', 'UID')

    # prepare patch list data for write
    fixture_type = fixture_mode.split('.')[0] if fixture_mode else inst_type
    user_id = channel if channel else uid.split('.')[0]
    dmx_universe = universe if is_number(universe) else ''
    dmx_address = address if (is_number(universe) and is_number(address)) else ''
    name = purpose if purpose else uid.split('.')[0]

    # write patch list data
    file.write(user_id+','+dmx_universe+':'+dmx_address+','+fixture_type+','+name+'\n\r')
