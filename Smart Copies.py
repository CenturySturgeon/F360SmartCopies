# Author- Juan Gras
from . import commands
import adsk.core, adsk.fusion, adsk.cam, traceback
import ctypes
from .lib import fusion360utils as futil
handlers=[]

CF_TEXT = 1
kernel32 = ctypes.windll.kernel32
kernel32.GlobalLock.argtypes = [ctypes.c_void_p]
kernel32.GlobalLock.restype = ctypes.c_void_p
kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]
user32 = ctypes.windll.user32
user32.GetClipboardData.restype = ctypes.c_void_p

def get_clipboard_text():
    user32.OpenClipboard(0)
    try:
        if user32.IsClipboardFormatAvailable(CF_TEXT):
            data = user32.GetClipboardData(CF_TEXT)
            data_locked = kernel32.GlobalLock(data)
            text = ctypes.c_char_p(data_locked)
            value = text.value
            kernel32.GlobalUnlock(data_locked)
            return value
    finally:
        user32.CloseClipboard()

class SC_CreateButtonPressedEventHandler(adsk.core.CommandCreatedEventHandler):
    def __init__(self):
        super().__init__()
    def notify(self,args):
        app=adsk.core.Application.get()
        ui=app.userInterface
        design = app.activeProduct

        dataProject = app.data.dataProjects[0]
        rootFolder = dataProject.rootFolder
        doc = app.activeDocument
        uParams = design.userParameters

        # Gets excel table copied on clipboard
        excel_clipboard = get_clipboard_text()
        # Decodes excel table and converts it to string
        excel_clipboard = excel_clipboard.decode('UTF-8')
        excel_clipboard = str(excel_clipboard)
        # Removes row delimiter \r and splits the table rows into lists
        excel_clipboard = excel_clipboard.replace("\r","")
        excel_clipboard = excel_clipboard.split("\n")
        # Each column will be separated by special delimiter \t, so it will be used to split into individual words
        for i in range(0,len(excel_clipboard)-1):
            excel_clipboard[i] = excel_clipboard[i].split("\t")

        # Sets column headers from the first list of excel_clipboard and removes it to avoid conflict
        columnHeaders = excel_clipboard[0]
        del excel_clipboard[0]
        file_names = ""
        # Iterates through data rows only
        for i in range(len(excel_clipboard)-1):
            row = excel_clipboard[i]
            nameoffile = "SmartCopy_" + str(i+1)
            # Iterates through each column of the row
            for j in range(len(row)):
                word = row[j]        
                colheader = columnHeaders[j]
                # Checks if the column header refers to the file's name instead of a user parameter and sets it if that's the case
                if colheader.replace(" ","").replace("_","").lower() == "nameoffile":
                    nameoffile = row[j]
                else:
                    paramvalue = row[j]
                    paramname = colheader
                    try:
                        # Update the user parameter by expression to avoid issues with units
                        uParams.itemByName(paramname).expression = paramvalue
                    except:
                        # Exception so execution doesn't stop on mistake (for better or worse)
                        ui.messageBox("Error while trying to update user parameter '", paramname, "' with expresion '", paramvalue, "'")
            try:
                doc.saveAs(nameoffile, rootFolder, '', '')
            except:
                ui.messageBox("Error while trying to save file '", nameoffile, "' under folder '", dataProject.name)
            # Appends all the file names into a single message
            file_names = file_names + nameoffile + ", "
        msg = f'All documents were saved to\n project: "{dataProject.name}" with the file names: "{file_names}"'
        ui.messageBox(msg)

def run(context):
    try:
        app = adsk.core.Application.get()
        ui  = app.userInterface
        cmdDefs = ui.commandDefinitions
        workSpace = ui.workspaces.itemById('FusionSolidEnvironment')
        tbPanels = workSpace.toolbarPanels
        
        #region Delete previous run created button and panel
        global tbPanel
        tbPanel = tbPanels.itemById('SmartCopiesPanel')
        if tbPanel:
            tbPanel.deleteMe()
        global SC_CreateButton
        SC_CreateButton = ui.commandDefinitions.itemById('SC_CreateButton')
        if SC_CreateButton:
            SC_CreateButton.deleteMe()
        #endregion

        # Create a toolbar panel next to the Select Panel
        tbPanel = tbPanels.add('SmartCopiesPanel', 'Smart Copies', 'SelectPanel', False)
        # Create a button command definition.
        SC_CreateButton = cmdDefs.addButtonDefinition('SC_CreateButton', 'Create Copies', 'Creates copies of current file using excel data stored on clipboard','Resources/Printer')
        # Create the button control
        buttonControl = tbPanel.controls.addCommand(SC_CreateButton)

        # Event Handlers
        SC_CreateButtonPressed=SC_CreateButtonPressedEventHandler()
        SC_CreateButton.commandCreated.add(SC_CreateButtonPressed)
        handlers.append(SC_CreateButtonPressed)
        
        # Make the button available in the panel at the top.
        buttonControl.isPromotedByDefault = True
        buttonControl.isPromoted = True
    except:
        futil.handle_error('run')

def stop(context):
    try:
        # Remove all of the event handlers
        futil.clear_handlers()
        # Remove the toolbar panel and its button
        if tbPanel:
            tbPanel.deleteMe()
        if SC_CreateButton:
            SC_CreateButton.deleteMe()
        commands.stop()
    except:
        futil.handle_error('stop')