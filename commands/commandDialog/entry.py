import adsk.core
import adsk.fusion
import os
import sys
import importlib.util
from ...lib import fusionAddInUtils as futil
from ... import config

# Add the script directory to Python path to import the KCLExporter
script_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'fusion-kcl-export-script')
if script_dir not in sys.path:
    sys.path.append(script_dir)
# Import the script module (Python treats hyphens in filenames as modules)
spec = importlib.util.spec_from_file_location("fusion_kcl_export", os.path.join(script_dir, "fusion-kcl-export.py"))
fusion_kcl_export = importlib.util.module_from_spec(spec)
spec.loader.exec_module(fusion_kcl_export)
KCLExporter = fusion_kcl_export.KCLExporter
app = adsk.core.Application.get()
ui = app.userInterface


# TODO *** Specify the command identity information. ***
CMD_ID = f'{config.COMPANY_NAME}_{config.ADDIN_NAME}_kclExport'
CMD_NAME = 'Export to KCL'
CMD_Description = 'Export current design to KCL format'

# Specify that the command will be promoted to the panel.
IS_PROMOTED = True

# TODO *** Define the location where the command button will be created. ***
# This is done by specifying the workspace, the tab, and the panel, and the 
# command it will be inserted beside. Not providing the command to position it
# will insert it at the end.
WORKSPACE_ID = 'FusionSolidEnvironment'
PANEL_ID = 'SolidScriptsAddinsPanel'
COMMAND_BESIDE_ID = 'ScriptsManagerCommand'

# Resource location for command icons, here we assume a sub folder in this directory named "resources".
ICON_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'resources', '')

# Local list of event handlers used to maintain a reference so
# they are not released and garbage collected.
local_handlers = []


# Executed when add-in is run.
def start():
    # Create a command Definition.
    cmd_def = ui.commandDefinitions.addButtonDefinition(CMD_ID, CMD_NAME, CMD_Description, ICON_FOLDER)

    # Define an event handler for the command created event. It will be called when the button is clicked.
    futil.add_handler(cmd_def.commandCreated, command_created)

    # ******** Add a button into the UI so the user can run the command. ********
    # Get the target workspace the button will be created in.
    workspace = ui.workspaces.itemById(WORKSPACE_ID)

    # Get the panel the button will be created in.
    panel = workspace.toolbarPanels.itemById(PANEL_ID)

    # Create the button command control in the UI after the specified existing command.
    control = panel.controls.addCommand(cmd_def, COMMAND_BESIDE_ID, False)

    # Specify if the command is promoted to the main toolbar. 
    control.isPromoted = IS_PROMOTED


# Executed when add-in is stopped.
def stop():
    # Get the various UI elements for this command
    workspace = ui.workspaces.itemById(WORKSPACE_ID)
    panel = workspace.toolbarPanels.itemById(PANEL_ID)
    command_control = panel.controls.itemById(CMD_ID)
    command_definition = ui.commandDefinitions.itemById(CMD_ID)

    # Delete the button command control
    if command_control:
        command_control.deleteMe()

    # Delete the command definition
    if command_definition:
        command_definition.deleteMe()


# Function that is called when a user clicks the corresponding button in the UI.
# This defines the contents of the command dialog and connects to the command related events.
def command_created(args: adsk.core.CommandCreatedEventArgs):
    # General logging for debug.
    futil.log(f'{CMD_NAME} Command Created Event')

    # https://help.autodesk.com/view/fusion360/ENU/?contextId=CommandInputs
    inputs = args.command.commandInputs

    # TODO Define the dialog for your command by adding different inputs to the command.

    # Create file path input for output location
    inputs.addStringValueInput('output_path', 'Output Path', '')
    
    # Create browse button for file selection
    inputs.addBoolValueInput('browse_file', 'Browse...', False, '', True)

    # TODO Connect to the events that are needed by this command.
    futil.add_handler(args.command.execute, command_execute, local_handlers=local_handlers)
    futil.add_handler(args.command.inputChanged, command_input_changed, local_handlers=local_handlers)
    futil.add_handler(args.command.executePreview, command_preview, local_handlers=local_handlers)
    futil.add_handler(args.command.validateInputs, command_validate_input, local_handlers=local_handlers)
    futil.add_handler(args.command.destroy, command_destroy, local_handlers=local_handlers)


# This event handler is called when the user clicks the OK button in the command dialog or 
# is immediately called after the created event not command inputs were created for the dialog.
def command_execute(args: adsk.core.CommandEventArgs):
    # General logging for debug.
    futil.log(f'{CMD_NAME} Command Execute Event')

    # TODO ******************************** Your code here ********************************

    # Get a reference to your command's inputs.
    inputs = args.command.commandInputs
    output_path: adsk.core.StringValueCommandInput = inputs.itemById('output_path')

    # Get the active design
    design = app.activeProduct
    if not design:
        ui.messageBox('No active design found.')
        return

    # Export to KCL using the real exporter
    try:
        # Create the exporter with clean output (no debug)
        exporter = KCLExporter(debug_planes=False)
        
        # Export the design to KCL
        kcl_content = exporter.export_design(design)
        
        # Ensure the output path has .kcl extension
        if not output_path.value.lower().endswith('.kcl'):
            output_path.value += '.kcl'
        
        # Write the KCL file
        with open(output_path.value, 'w', encoding='utf-8') as f:
            f.write(kcl_content)
        
        ui.messageBox(f'Successfully exported to KCL: {output_path.value}')
    except Exception as e:
        ui.messageBox(f'Error exporting to KCL: {str(e)}')


# This event handler is called when the command needs to compute a new preview in the graphics window.
def command_preview(args: adsk.core.CommandEventArgs):
    # General logging for debug.
    futil.log(f'{CMD_NAME} Command Preview Event')
    inputs = args.command.commandInputs


# This event handler is called when the user changes anything in the command dialog
# allowing you to modify values of other inputs based on that change.
def command_input_changed(args: adsk.core.InputChangedEventArgs):
    changed_input = args.input
    inputs = args.inputs

    # General logging for debug.
    futil.log(f'{CMD_NAME} Input Changed Event fired from a change to {changed_input.id}')
    
    # Handle browse button click
    if changed_input.id == 'browse_file':
        if changed_input.value:
            # Show file save dialog
            fileDialog = ui.createFileDialog()
            fileDialog.isMultiSelectEnabled = False
            fileDialog.title = "Save KCL File"
            fileDialog.filter = 'KCL files (*.kcl);;All files (*.*)'
            fileDialog.filterIndex = 0
            dialogResult = fileDialog.showSave()
            if dialogResult == adsk.core.DialogResults.DialogOK:
                filename = fileDialog.filename
                output_path_input = inputs.itemById('output_path')
                output_path_input.value = filename
            # Reset the button
            changed_input.value = False


# This event handler is called when the user interacts with any of the inputs in the dialog
# which allows you to verify that all of the inputs are valid and enables the OK button.
def command_validate_input(args: adsk.core.ValidateInputsEventArgs):
    # General logging for debug.
    futil.log(f'{CMD_NAME} Validate Input Event')

    inputs = args.inputs
    
    # Verify the validity of the input values. This controls if the OK button is enabled or not.
    output_path_input = inputs.itemById('output_path')
    if output_path_input.value and output_path_input.value.strip():
        args.areInputsValid = True
    else:
        args.areInputsValid = False
        

# This event handler is called when the command terminates.
def command_destroy(args: adsk.core.CommandEventArgs):
    # General logging for debug.
    futil.log(f'{CMD_NAME} Command Destroy Event')

    global local_handlers
    local_handlers = []
