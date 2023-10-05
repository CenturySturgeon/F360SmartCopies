# Author- Juan Gras
from . import commands
import adsk.core, adsk.fusion, adsk.cam, traceback
from .lib import fusion360utils as futil
import ctypes

import time
handlers=[]

# Excel data extraction and analysis
CF_TEXT = 1
kernel32 = ctypes.windll.kernel32
kernel32.GlobalLock.argtypes = [ctypes.c_void_p]
kernel32.GlobalLock.restype = ctypes.c_void_p
kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]
user32 = ctypes.windll.user32
user32.GetClipboardData.restype = ctypes.c_void_p

def get_clipboard_text():
    """Function that gets the text copied on the clipboard."""
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

def get_excelclipboard():
    """Function that gets the excel table copied on clipboard and formats it accordingly"""
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
    return excel_clipboard

def check_uparams_exist(excel_clipboard,columnHeaders,caseSensitive):
    """Function that confirms if all the user parameters provided by the excel clipboard exist in the active document."""
    app = adsk.core.Application.get()
    ui = app.userInterface
    design = app.activeProduct
    uParams = design.userParameters
    allParamsExist = True

    for i in range(len(excel_clipboard)-1):
        row = excel_clipboard[i]
        # Iterates through each column of the row (values for the filenames [if provided] and user parameters)
        for j in range(len(row)):
            colheader = columnHeaders[j]
            # Checks if the column header refers to the file's name instead of a user parameter and sets it if that's the case
            if colheader.replace(" ","").replace("_","").lower() == "nameoffile":
                pass
            else:
                paramname = colheader
                paramexists = False
                for k in range(uParams.count):
                    if not caseSensitive:
                        if uParams.item(k).name.upper() == paramname.upper():
                            paramexists = True   
                    elif uParams.item(k).name == paramname:
                        paramexists = True
                        break
                if paramexists == False:
                    allParamsExist = False
    return allParamsExist

#region Clock Saving Time dialog box and execution event handler
class SC_SavingTimeButton_PressedEventHandler(adsk.core.CommandCreatedEventHandler):
    def __init__(self):
        super().__init__()
    def notify(self,args):
        ui = None
        try:
            app=adsk.core.Application.get()
            ui=app.userInterface
            doc = app.activeDocument
            dataProject = app.data.dataProjects[1] 
            rootFolder = dataProject.rootFolder

            starttime = time.time()
            nameoffile = "SmartCopies_ClockTestFile"
            doc.saveAs(nameoffile, rootFolder, '', '')

            donesaving = False
            while not donesaving:
                endtime = time.time()
                for i in range(dataProject.rootFolder.dataFiles.count):
                    cloudfile = dataProject.rootFolder.dataFiles.item(i)
                    if cloudfile.name == nameoffile:
                        if cloudfile.isComplete:
                            donesaving = True
            
            endtime = time.time()
            totalSavingTime = endtime-starttime
            msg = f'File ["{nameoffile}"] was saved to project: "{dataProject.name}".\nTotal Saving Time: "{totalSavingTime}" s'

            ui.messageBox(msg)
            
        except:
            if ui:
                ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
#endregion

#region Create copies dialog box and execution event handler
# Here is the dialog box and its commandInputs (buttons and stuff)
class SC_CreateButton_PressedEventHandler(adsk.core.CommandCreatedEventHandler):
    def __init__(self):
        super().__init__()
    def notify(self,args):
        ui = None
        try:
            app=adsk.core.Application.get()
            ui=app.userInterface

            # Remember, once you resize a Dialog box it's default size will be forever modified to that. To fix it go and delete the file NULastDisplayedLayout.xml (if you recently used fusion check the date in case there are 2 or more in different folders [conflicts with inventor]) -> C:\Users\USER_NAME\AppData\Roaming\Autodesk\Neutron Platform\Options\VARIABLE_FOLDER\NULastDisplayedLayout.xml
            # It's ok to delete the entire file and Fusion will just revert everything back to the default state. 
            # To read more https://forums.autodesk.com/t5/fusion-360-api-and-scripts/how-do-you-change-the-size-of-a-command-dialog-box/td-p/6231098
            
            cmd = args.command
            cmd.setDialogInitialSize(360, 175)
            inputs=cmd.commandInputs
            
            DProjects_DropDown = inputs.addDropDownCommandInput('DataProjectsDropdown_CommandInput','Project Folder Name:',1)
            DProjects_List=DProjects_DropDown.listItems
            for i in range(0,app.data.dataProjects.count):
                isSelected = False
                if i == 2:
                    isSelected = True
                DProjects_List.add(app.data.dataProjects.item(i).name,isSelected)

            ValueInput1 = adsk.core.ValueInput.createByReal(10)
            inputs.addValueInput('MaxWaitTime_ValueInput', 'Maximum File Waiting Time:', 's', ValueInput1)
            
            inputs.addBoolValueInput('CaseSensitive_CheckboxInput', 'Case Sensitive Parameter Names:', True, '', True)
            inputs.addBoolValueInput('StopSaving_CheckboxInput', 'Stop Queue If Wait Time is Exceeded:', True, '', False)
            
            onExecute = cmdDefOKButtonPressedEventHandler()
            cmd.execute.add(onExecute)
            handlers.append(onExecute)
            
        except:
            if ui:
                ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

# Here is what happens when the OK button of the dialog box is executed
class cmdDefOKButtonPressedEventHandler(adsk.core.CommandEventHandler):
    def __init__(self):
        super().__init__()
    def notify(self,args):
        try:
            eventArgs = adsk.core.CommandEventArgs.cast(args)
            app = adsk.core.Application.get()
            ui = app.userInterface
            design = app.activeProduct
            dialogBoxInputs = eventArgs.command.commandInputs

            dataProjectName = dialogBoxInputs.itemById('DataProjectsDropdown_CommandInput').selectedItem.name
            caseSensitive = dialogBoxInputs.itemById('CaseSensitive_CheckboxInput').value
            excesiveTimeBreak = dialogBoxInputs.itemById('StopSaving_CheckboxInput').value
            waitTime=dialogBoxInputs.itemById('MaxWaitTime_ValueInput').value

            # Sets a minimum and maximum waiting time value to avoid issues like not enough or excesive file saving time
            if waitTime < 1:
                waitTime = 5
            elif waitTime > 300:
                waitTime = 300
                
            for i in range(app.data.dataProjects.count):
                if app.data.dataProjects.item(i).name == dataProjectName:
                    dataProject = app.data.dataProjects[i]
            
            rootFolder = dataProject.rootFolder
            doc = app.activeDocument
            uParams = design.userParameters

            excel_clipboard = get_excelclipboard()
            # Sets column headers from the first row of excel_clipboard and removes it to avoid conflicts down the road
            columnHeaders = excel_clipboard[0]
            del excel_clipboard[0]
            
            # Checks all user parameters exist (caseSensitive flag criteria is used) and sets initial values for the variables
            AllParamsExists = check_uparams_exist(excel_clipboard, columnHeaders, caseSensitive)
            file_names = ""
            tooktoolong = False
            invalidMessageCounter = 0
            invalidFileList = []
            if not AllParamsExists:
                ui.messageBox("Not all user parameters were found, please review them or try turning on the case sensitive flag.")
            else:
                # Iterates through data rows only (nameoffile and uparams names)
                for i in range(len(excel_clipboard)-1):
                    # Checks if the save has taken more time than what was set as maximum and the Stop saving flag is enabled, if this condition occurs the process will stop
                    if tooktoolong==True and excesiveTimeBreak:
                        pass
                    else:
                        row = excel_clipboard[i]
                        nameoffile = "SmartCopy_" + str(i+1)
                        saveworthy = True
                        # Iterates through each column of the row (values for the filenames [if provided] and user parameters)
                        for j in range(len(row)):
                            colheader = columnHeaders[j]
                            # Checks if the column header refers to the file's name instead of a user parameter and sets it if that's the case
                            if colheader.replace(" ","").replace("_","").lower() == "nameoffile" or colheader.replace(" ","").replace("_","").lower() == "nameoffiles":
                                nameoffile = row[j]
                            else:
                                if saveworthy: 
                                    paramvalue = row[j].replace(" ", "")
                                    paramname = colheader
                                    try:
                                        #  Searches for and updates the user parameter by expression to avoid issues with units
                                        if caseSensitive:
                                            uParams.itemByName(paramname).expression = paramvalue
                                        else:
                                            for k in range(uParams.count):
                                                if uParams.item(k).name.upper() == paramname.upper():
                                                    uParams.item(k).expression = paramvalue
                                    # Exception so execution doesn't stop on mistake (for better or worse)
                                    except:
                                        if invalidMessageCounter == 0:
                                            ui.messageBox("Error while trying to update user parameter '" + paramname + "' with expresion '" + paramvalue + "': File was not saved.\n\nThis message is only displayed once, if more files fail like this at the end of the process you'll see the error summary.")
                                        invalidFileList.append(nameoffile)
                                        nameoffile = ""
                                        saveworthy = False
                                        invalidMessageCounter += 1
                            # Depending on how people order the columns the name displayed on the saved files or error summary might be wrong. 
                            # This is a last meassure to ensure the information provided by the error messages is correct.
                            if not saveworthy and nameoffile != "SmartCopy_" + str(i+1):
                                del invalidFileList[-1]
                                invalidFileList.append(nameoffile)
                                nameoffile = ""
                        # Determines if the file is worthy of saving due to bad parameters
                        try:
                            if saveworthy:
                                # Since the files are saved in the cloud, a syncronization filter is added so it can verify that the save is also completed in the cloud
                                # If it isn't errors will occur, mainly the new file will be saved as a version of the previous.
                                # For more info check the isComplete property of the dataFile...
                                doc.saveAs(nameoffile, rootFolder, '', '')
                                # donesaving determines when the version of the file in the cloud has already been saved
                                donesaving = False
                                starttime = time.time()
                                tooktoolong = False
                                #This block checks that each save doesnt take more than 'waitTime' seconds, if it does it is forced to stop at 'waitTime' seconds
                                #It also checks if the file is already saved to the cloud so it can move on with the next one
                                while donesaving == False and tooktoolong == False:
                                    endtime = time.time()
                                    if endtime-starttime > waitTime:
                                        tooktoolong = True
                                        if excesiveTimeBreak:
                                            ui.messageBox("'Stop queue' flag was enabled and time for saving the file was exceeded, no more files will be saved to avoid overwriting issues.")
                                    for i in range(dataProject.rootFolder.dataFiles.count):
                                        cloudfile = dataProject.rootFolder.dataFiles.item(i)
                                        if cloudfile.name == nameoffile:
                                            if cloudfile.isComplete:
                                                donesaving = True
                        except:
                            ui.messageBox("Error while trying to save file '" + nameoffile + "' under project folder '" + dataProject.name)
                            nameoffile = ""
                        # To have a clean "Files saved" message it's not worthwhile to have multiple commas between file names
                        if nameoffile != "":
                            file_names = file_names + nameoffile + ", "
                
                # Delete the last comma of the message
                file_names = file_names.rstrip(', ')
                msg = f'Files ["{file_names}"] were all saved to project: "{dataProject.name}"'
                
                # If no files were saved, then no "Files saved" message should show up
                if file_names.replace(", ", "") !="":
                    ui.messageBox(msg)
                
                # Displays the error summary message
                if invalidMessageCounter > 1:
                    errormsg = f'Files "{str(invalidFileList)}" were not saved due to user parameter error(s). Please check the units or validate the expressions employed are correct.'
                    ui.messageBox(errormsg)
                #app.activeDocument.close(False)
        except:
            if ui:
                ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
#endregion

def run(context):
    ui = None
    try:
        app = adsk.core.Application.get()
        ui  = app.userInterface
        cmdDefs = ui.commandDefinitions
        workSpace = ui.workspaces.itemById('FusionSolidEnvironment')
        tbPanels = workSpace.toolbarPanels
        
        #region Delete previous run created button and toolbar panel
        global tbPanel
        tbPanel = tbPanels.itemById('SmartCopiesPanel')
        if tbPanel:
            tbPanel.deleteMe()
        global SC_CreateButton
        global SC_SavingTimeButton
        SC_CreateButton = ui.commandDefinitions.itemById('SC_CreateButton')
        SC_SavingTimeButton  = ui.commandDefinitions.itemById('SC_SavingTimeButton')
        if SC_CreateButton:
            SC_CreateButton.deleteMe()
        if SC_SavingTimeButton:
            SC_SavingTimeButton.deleteMe()
        #endregion

        # Create a toolbar panel next to the Select Panel
        tbPanel = tbPanels.add('SmartCopiesPanel', 'Smart Copies', 'SelectPanel', False)
        
        #region Create a button command definition.
        SC_CreateButton = cmdDefs.addButtonDefinition('SC_CreateButton', 'Create Copies', 'Creates copies of current file using excel data stored on clipboard.','Resources/Printer')
        SC_SavingTimeButton = cmdDefs.addButtonDefinition('SC_SavingTimeButton', 'Clock Saving Time', 'Displays the time it takes to save the active document to the fussion cloud.','Resources/Clock')
        #endregion

        #region Create the button controls
        SC_CreateButton_Control = tbPanel.controls.addCommand(SC_CreateButton)
        SC_SavingTimeButton_Control = tbPanel.controls.addCommand(SC_SavingTimeButton)
        #endregion

        #region Event Handlers
        SC_CreateButton_Pressed = SC_CreateButton_PressedEventHandler()
        SC_CreateButton.commandCreated.add(SC_CreateButton_Pressed)
        handlers.append(SC_CreateButton_Pressed)

        SC_SavingTimeButton_Pressed = SC_SavingTimeButton_PressedEventHandler()
        SC_SavingTimeButton.commandCreated.add(SC_SavingTimeButton_Pressed)
        handlers.append(SC_SavingTimeButton_Pressed)
        #endregion
        
        # Make the button available in the panel at the top.
        SC_CreateButton_Control.isPromotedByDefault = True
        SC_CreateButton_Control.isPromoted = True
    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

def stop(context):
    ui = None
    try:
        # Remove all of the event handlers
        futil.clear_handlers()
        # Remove the toolbar panel and its buttons
        if tbPanel:
            tbPanel.deleteMe()
        if SC_CreateButton:
            SC_CreateButton.deleteMe()
        if SC_SavingTimeButton:
            SC_SavingTimeButton.deleteMe()
        commands.stop()
    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))