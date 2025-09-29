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
CMD_ID = f'{config.COMPANY_NAME}_{config.ADDIN_NAME}_batchProcess'
CMD_NAME = 'Batch Export to KCL'
CMD_Description = 'Batch export entire project folders to KCL format'

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

    # Add explanation text
    explanation_text = inputs.addTextBoxCommandInput('explanation', '', 
        'BATCH EXPORT FROM PROJECT FOLDER:\n\n'
        'This tool exports all .f3d files from the same project folder as your currently open design.\n\n'
        'HOW IT WORKS:\n'
        '1. Uses your currently active design to find its project folder\n'
        '2. Finds all other .f3d files in that same folder\n'
        '3. Opens each file and exports it to KCL format\n'
        '4. Skips your currently active design (to avoid conflicts)\n\n'
        'REQUIREMENTS:\n'
        '• Have a design file open in Fusion 360\n'
        '• Design must be saved to a project folder\n'
        '• Other .f3d files should be in the same project folder', 8, True)
    explanation_text.isFullWidth = True
    
    # Create output folder input
    inputs.addStringValueInput('output_folder', 'Output Folder', '')
    
    # Create browse button for output folder selection
    inputs.addBoolValueInput('browse_output', 'Browse Output...', False, '', True)
    
    # Create progress text box (read-only)
    inputs.addTextBoxCommandInput('progress_text', 'Progress', 'Ready to process...', 3, True)

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
    output_folder: adsk.core.StringValueCommandInput = inputs.itemById('output_folder')
    progress_text: adsk.core.TextBoxCommandInput = inputs.itemById('progress_text')

    # Batch process from Fusion 360 project (no local folder needed)
    try:
        successful_count = batch_export_to_kcl(
            None,  # No project folder - uses active Fusion 360 project
            output_folder.value, 
            progress_text
        )
        if successful_count > 0:
            ui.messageBox(f'Successfully exported {successful_count} files to: {output_folder.value}')
        else:
            # Get more detailed error info
            error_details = get_batch_error_summary()
            ui.messageBox(f'No files were successfully exported.\n\n{error_details}')
    except Exception as e:
        ui.messageBox(f'Error during batch export: {str(e)}')


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
    
    # Handle browse output button click
    if changed_input.id == 'browse_output':
        if changed_input.value:
            # Show folder selection dialog for output
            folderDialog = ui.createFolderDialog()
            folderDialog.title = "Select Output Folder"
            dialogResult = folderDialog.showDialog()
            if dialogResult == adsk.core.DialogResults.DialogOK:
                folder_path = folderDialog.folder
                output_folder_input = inputs.itemById('output_folder')
                output_folder_input.value = folder_path
            # Reset the button
            changed_input.value = False


# This event handler is called when the user interacts with any of the inputs in the dialog
# which allows you to verify that all of the inputs are valid and enables the OK button.
def command_validate_input(args: adsk.core.ValidateInputsEventArgs):
    # General logging for debug.
    futil.log(f'{CMD_NAME} Validate Input Event')

    inputs = args.inputs
    
    # Verify the validity of the input values. This controls if the OK button is enabled or not.
    output_folder_input = inputs.itemById('output_folder')
    
    if output_folder_input.value and output_folder_input.value.strip():
        args.areInputsValid = True
    else:
        args.areInputsValid = False
        

# This event handler is called when the command terminates.
def command_destroy(args: adsk.core.CommandEventArgs):
    # General logging for debug.
    futil.log(f'{CMD_NAME} Command Destroy Event')

    global local_handlers
    local_handlers = []


def get_batch_error_summary():
    """Generate a helpful error summary for batch export failures"""
    issues = []
    
    # Add the last error if available
    if hasattr(batch_export_to_kcl, 'last_error'):
        if batch_export_to_kcl.last_error:
            issues.append(f"💥 Error: {batch_export_to_kcl.last_error}")
    else:
        issues.append("❌ No specific error details available")
    
    # Add helpful instructions
    issues.append("")
    issues.append("📋 TROUBLESHOOTING:")
    issues.append("1. Ensure you have a design file open in Fusion 360")
    issues.append("2. Make sure the design is saved to a Fusion 360 project folder")
    issues.append("3. Verify there are other .f3d files in the same project folder")
    issues.append("4. Check that you have read access to the project folder")
    issues.append("")
    issues.append("💡 How batch export works:")
    issues.append("   • Uses your currently open design to find its project folder")
    issues.append("   • Exports all other .f3d files from that same folder")
    issues.append("   • Skips the currently active design (to avoid conflicts)")
    
    return "\n".join(issues)


def batch_export_to_kcl(project_folder, output_folder, progress_text=None):
    """NEW batch export approach - use active design's project folder"""
    
    # Clear any previous error
    if hasattr(batch_export_to_kcl, 'last_error'):
        delattr(batch_export_to_kcl, 'last_error')
    
    futil.log(f"=== BATCH EXPORT STARTED (ACTIVE DESIGN APPROACH) ===")
    
    try:
        # Get the currently active design
        design = app.activeProduct
        if not design:
            raise Exception("No active design found. Please open a design file first.")
        
        if design.objectType != adsk.fusion.Design.classType():
            raise Exception("Active product is not a Fusion 360 design.")
        
        # Get the active design's data file from the parent document
        active_document = design.parentDocument
        if not active_document:
            raise Exception("Active design does not have an associated document.")
        
        active_data_file = active_document.dataFile
        if not active_data_file:
            raise Exception("Active design document does not have an associated data file. Please save the design first.")
        
        futil.log(f"Active design: {active_data_file.name}")
        futil.log(f"Active design ID: {active_data_file.id}")
        
        # Get the parent folder of the active design
        project_folder_obj = active_data_file.parentFolder
        if not project_folder_obj:
            raise Exception("Could not find parent folder for active design.")
        
        futil.log(f"Project folder: {project_folder_obj.name}")
        futil.log(f"Project folder ID: {project_folder_obj.id}")
        
        # Collect all design files in the project folder
        design_files = []
        file_count = project_folder_obj.dataFiles.count
        futil.log(f"Found {file_count} files in project folder")
        
        for i in range(file_count):
            try:
                data_file = project_folder_obj.dataFiles.item(i)
                
                # Only include Fusion 360 design files (.f3d) that are not the currently active design
                if (hasattr(data_file, 'fileExtension') and 
                    data_file.fileExtension == 'f3d' and 
                    data_file.id != active_data_file.id):
                    
                    design_files.append(data_file)
                    futil.log(f"  📄 Will export: {data_file.name}")
                else:
                    futil.log(f"  ⏭️  Skipping: {data_file.name} (not .f3d or is active design)")
                    
            except Exception as file_error:
                futil.log(f"Error accessing file {i}: {str(file_error)}")
        
        if not design_files:
            raise Exception("No other design files found in the same project folder as the active design.")
        
        futil.log(f"=== FOUND {len(design_files)} DESIGN FILES TO EXPORT ===")
        
        # Create output folder if it doesn't exist
        os.makedirs(output_folder, exist_ok=True)
        
        total_files = len(design_files)
        processed = 0
        successful = 0
        
        if progress_text:
            progress_text.text = f"Found {total_files} design files to export..."
        
        # Process each design file
        for data_file in design_files:
            document = None
            try:
                processed += 1
                if progress_text:
                    progress_text.text = f"Processing {processed}/{total_files}: {data_file.name}"
                
                futil.log(f"Opening design file: {data_file.name}")
                
                # Open the design file
                document = app.documents.open(data_file)
                if not document:
                    raise Exception(f"Failed to open design file: {data_file.name}")
                
                futil.log(f"Document opened: {document.name}")
                
                # Activate the document
                document.activate()
                futil.log(f"Document activated")
                
                # Get the design from the opened document
                opened_design = app.activeProduct
                if not opened_design:
                    raise Exception("No active product after opening design file")
                
                if opened_design.objectType != adsk.fusion.Design.classType():
                    raise Exception(f"Opened file is not a design: {opened_design.objectType}")
                
                futil.log(f"Opened design: {opened_design.parentDocument.name}")
                futil.log(f"Content: {opened_design.rootComponent.sketches.count} sketches, {opened_design.rootComponent.features.count} features")
                
                # Generate output filename
                base_name = os.path.splitext(data_file.name)[0]
                output_file = os.path.join(output_folder, f"{base_name}.kcl")
                
                # Export using the script's KCLExporter
                futil.log(f"Starting KCL export...")
                
                # Create the exporter (use the same settings as the script)
                exporter = KCLExporter(debug_planes=True)
                
                # Export the design
                kcl_content = exporter.export_design(opened_design)
                
                # Write the KCL file
                with open(output_file, 'w', encoding='utf-8') as f:
                    f.write(kcl_content)
                
                successful += 1
                futil.log(f"SUCCESS: Exported {data_file.name} -> {output_file}")
                
                if progress_text:
                    progress_text.text = f"Exported {successful}/{total_files}: {data_file.name}"
                
            except Exception as file_error:
                error_msg = str(file_error)
                futil.log(f"ERROR processing {data_file.name}: {error_msg}")
                if progress_text:
                    progress_text.text = f"Error: {data_file.name} - {error_msg}"
                
                # Store the last error
                if not hasattr(batch_export_to_kcl, 'last_error'):
                    batch_export_to_kcl.last_error = error_msg
                    
            finally:
                # Always close the document
                if document:
                    try:
                        document.close(False)  # Close without saving
                        futil.log(f"Document closed: {document.name}")
                    except Exception as close_error:
                        futil.log(f"Error closing document: {str(close_error)}")
        
        # Reactivate the original design
        try:
            original_document = app.documents.open(active_data_file)
            if original_document:
                original_document.activate()
                futil.log(f"Reactivated original design: {active_data_file.name}")
        except Exception as reactivate_error:
            futil.log(f"Warning: Could not reactivate original design: {str(reactivate_error)}")
        
        if progress_text:
            progress_text.text = f"Completed! Successfully processed {successful}/{total_files} files."
        
        futil.log(f"=== BATCH EXPORT COMPLETED ===")
        futil.log(f"Total files: {total_files}, Successful: {successful}")
        
        return successful
        
    except Exception as error:
        error_msg = f"Batch export error: {str(error)}"
        futil.log(f"ERROR: {error_msg}")
        if progress_text:
            progress_text.text = f"Error: {str(error)}"
        
        # Store the error
        if not hasattr(batch_export_to_kcl, 'last_error'):
            batch_export_to_kcl.last_error = error_msg
        
        return 0

