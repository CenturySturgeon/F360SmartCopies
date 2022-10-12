# Author- Juan Gras
from email import message
from xml.dom.minidom import Document
from . import commands
import adsk.core, adsk.fusion, adsk.cam, traceback
from .lib import fusion360utils as futil
import ctypes

import time
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
        ui = None
        try:
            ### Make it so that users can choose the project by puting it in a dropdown menu
            ### Add the user parameters names are case sensitive option
            ### Add the "impatient mode" where users will give a maximum amount of time for their files to be saved correctly to the cloud before moving to the next one

            app=adsk.core.Application.get()
            ui=app.userInterface
            design = app.activeProduct
            dataProject = app.data.dataProjects[2]
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
            # Iterates through data rows only (nameoffile and uparams names)
            for i in range(len(excel_clipboard)-1):
                row = excel_clipboard[i]
                nameoffile = "SmartCopy_" + str(i+1)
                saveworthy = True
                # Iterates through each column of the row (values for the filenames [if provided] and user parameters)
                for j in range(len(row)):
                    word = row[j]        
                    colheader = columnHeaders[j]
                    # Checks if the column header refers to the file's name instead of a user parameter and sets it if that's the case
                    if colheader.replace(" ","").replace("_","").lower() == "nameoffile":
                        nameoffile = row[j]
                    else:
                        paramvalue = row[j].replace(" ", "")
                        paramname = colheader
                        try:
                            ### Make it so there's an option that case sentiveness doesn't ruin everything by first iterating through all uparams and finds the one that matches
                            # Update the user parameter by expression to avoid issues with units
                            uParams.itemByName(paramname).expression = paramvalue
                        except:
                            # Exception so execution doesn't stop on mistake (for better or worse)
                            ui.messageBox("Error while trying to update user parameter '" + paramname + "' with expresion '" + paramvalue + "'. File was not saved.")
                            nameoffile = ""
                            saveworthy = False
                # Determines if the file is worthy of saving due to bad parameters and if it is, it saves it
                try:
                    if saveworthy:
                        # Since the files are saved in the cloud, a syncronization filter is added so it can verify that the save is also completed in the cloud (if it isn't errors will occur, mainly the new file will be saved as a version of the previous) for more info check the isComplete property of the dataFile
                        doc.saveAs(nameoffile, rootFolder, '', '')
                        # donesaving determines when the version of the file in the cloud has already been saved
                        donesaving = False
                        starttime = time.time()
                        tooktoolong = False
                        messagecounter = 0
                        #This block checks that each save doesnt take more than 20s, if it does it is forced to stop at 20
                        #It also checks if the file is already saved to the cloud so it can move on with the next one
                        while donesaving == False and tooktoolong == False:
                            endtime = time.time()
                            if endtime-starttime > 20 and messagecounter == 0:
                                ui.messageBox("Since Fusion saves its files on the cloud and the internet Connection is slow or the files are too heavy; errors might occur. Wait time for each file is 20s")
                                tooktoolong = True
                                messagecounter = 1
                            for i in range(dataProject.rootFolder.dataFiles.count):
                                cloudfile = dataProject.rootFolder.dataFiles.item(i)
                                if cloudfile.name == nameoffile:
                                    if cloudfile.isComplete:
                                        donesaving = True
                except:
                    ui.messageBox("Error while trying to save file '" + nameoffile + "' under project folder '" + dataProject.name)
                    nameoffile = ""
                # Appends all the file names into a single message
                coma_or_not = ", "
                if i == len(excel_clipboard)-2:
                    coma_or_not = ""
                file_names = file_names + nameoffile + coma_or_not

            msg = f'Files ["{file_names}"] were all saved to \n project: "{dataProject.name}"'
            ui.messageBox(msg)
        except:
            if ui:
                ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

def run(context):
    ui = None
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
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

def stop(context):
    ui = None
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
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))