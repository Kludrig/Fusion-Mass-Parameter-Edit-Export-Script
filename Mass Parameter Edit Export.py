import adsk.core, adsk.fusion, traceback
import os
import re

# Global list to keep handlers alive
handlers = []

def run(context):
    ui = None
    try:
        app = adsk.core.Application.get()
        ui = app.userInterface
        
        # Create command
        cmdDef = ui.commandDefinitions.itemById('BatchExportCmd')
        if cmdDef:
            cmdDef.deleteMe()
            
        cmdDef = ui.commandDefinitions.addButtonDefinition(
            'BatchExportCmd',
            'Batch Parameter Export',
            'Export multiple files with parameter variations'
        )
        
        # Connect to command created event
        onCommandCreated = MyCommandCreatedHandler()
        cmdDef.commandCreated.add(onCommandCreated)
        handlers.append(onCommandCreated)
        
        # Execute command
        cmdDef.execute()
        
        # Prevent auto-terminate
        adsk.autoTerminate(False)
        
    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


def sanitize_id(name):
    """Convert name to valid ID by removing invalid characters"""
    safe_id = re.sub(r'[^a-zA-Z0-9_]', '_', name)
    if not safe_id[0].isalpha():
        safe_id = 'obj_' + safe_id
    return safe_id


class MyCommandCreatedHandler(adsk.core.CommandCreatedEventHandler):
    def __init__(self):
        super().__init__()
        
    def notify(self, args):
        try:
            cmd = args.command
            cmd.isExecutedWhenPreEmpted = False
            
            onExecute = MyCommandExecuteHandler()
            cmd.execute.add(onExecute)
            handlers.append(onExecute)
            
            onDestroy = MyCommandDestroyHandler()
            cmd.destroy.add(onDestroy)
            handlers.append(onDestroy)
            
            # Get design
            app = adsk.core.Application.get()
            design = adsk.fusion.Design.cast(app.activeProduct)
            
            if not design:
                return
            
            inputs = cmd.commandInputs
            
            # Get data
            userParams = design.userParameters
            rootComp = design.rootComponent
            
            bodies = []
            for body in rootComp.bRepBodies:
                bodies.append(body.name)
            
            components = []
            for occ in rootComp.allOccurrences:
                components.append(occ.name)
            
            # 1. Parameter dropdown
            paramDropdown = inputs.addDropDownCommandInput(
                'paramSelect',
                'User Parameter',
                adsk.core.DropDownStyles.LabeledIconDropDownStyle
            )
            paramDropdown.tooltip = 'Select the parameter to modify for each variation'
            for i in range(userParams.count):
                param = userParams.item(i)
                paramDropdown.listItems.add(f'{param.name} = {param.expression}', False)
            if paramDropdown.listItems.count > 0:
                paramDropdown.listItems.item(0).isSelected = True
            
            # 2. Parameter type
            typeGroup = inputs.addRadioButtonGroupCommandInput('paramType', 'Parameter Type')
            typeGroup.listItems.add('Text (single quotes)', True)
            typeGroup.listItems.add('Numeric', False)
            
            # 3. Variations text box
            variationsInput = inputs.addTextBoxCommandInput(
                'variations',
                'Variations (comma-separated)',
                'A, B, C',
                8,
                False
            )
            variationsInput.tooltip = 'Enter all variations separated by commas. Example: A, B, C, D'
            
            # Add spacer
            inputs.addTextBoxCommandInput('spacer1', '', '', 1, True)
            
            # 4. FILE NAMING OPTIONS GROUP
            namingGroup = inputs.addGroupCommandInput('namingGroup', 'File Naming Options')
            namingGroup.isExpanded = True
            namingGroup.isEnabledCheckBoxDisplayed = False
            namingInputs = namingGroup.children
            
            # Prefix
            prefixInput = namingInputs.addStringValueInput('filePrefix', 'Prefix', '')
            prefixInput.tooltip = 'Optional prefix for all filenames'
            
            # Suffix
            suffixInput = namingInputs.addStringValueInput('fileSuffix', 'Suffix', '')
            suffixInput.tooltip = 'Optional suffix for all filenames'
            
            # Numbering
            namingInputs.addBoolValueInput('addNumbering', 'Add numbering (001, 002, ...)', True, '', False)
            
            # Include parameter name
            namingInputs.addBoolValueInput('includeParamName', 'Include parameter name in filename', True, '', False)
            
            # 5. EXPORT FORMAT GROUP
            formatGroup = inputs.addGroupCommandInput('formatGroup', 'Export Format')
            formatGroup.isExpanded = True
            formatGroup.isEnabledCheckBoxDisplayed = False
            formatInputs = formatGroup.children
            
            # Format dropdown
            formatDropdown = formatInputs.addDropDownCommandInput(
                'exportFormat',
                'File Format',
                adsk.core.DropDownStyles.LabeledIconDropDownStyle
            )
            formatDropdown.listItems.add('STL (Binary)', True)
            formatDropdown.listItems.add('STL (ASCII)', False)
            formatDropdown.listItems.add('3MF (with color)', False)
            formatDropdown.listItems.add('OBJ (with color)', False)
            formatDropdown.listItems.add('STEP', False)
            formatDropdown.listItems.add('F3D (Archive)', False)
            formatDropdown.tooltip = 'Select export file format'
            
            # Unit dropdown
            unitDropdown = formatInputs.addDropDownCommandInput(
                'exportUnit',
                'Export Unit',
                adsk.core.DropDownStyles.LabeledIconDropDownStyle
            )
            unitDropdown.listItems.add('Millimeters', True)
            unitDropdown.listItems.add('Centimeters', False)
            unitDropdown.listItems.add('Meters', False)
            unitDropdown.listItems.add('Inches', False)
            unitDropdown.listItems.add('Feet', False)
            unitDropdown.tooltip = 'Unit for exported files'
            
            # Mesh options
            formatInputs.addTextBoxCommandInput('stlLabel', 'Mesh Options', '---', 1, True)
            
            meshRefinement = formatInputs.addDropDownCommandInput(
                'meshRefinement',
                'Mesh Refinement',
                adsk.core.DropDownStyles.LabeledIconDropDownStyle
            )
            meshRefinement.listItems.add('Low', False)
            meshRefinement.listItems.add('Medium', True)
            meshRefinement.listItems.add('High', False)
            meshRefinement.tooltip = 'Higher refinement = smoother curves but larger files'
            meshRefinement.isVisible = True
            
            # Add spacer
            inputs.addTextBoxCommandInput('spacer2', '', '', 1, True)
            
            # 6. Group for bodies
            if bodies:
                bodyGroup = inputs.addGroupCommandInput('bodyGroup', 'Bodies')
                bodyGroup.isExpanded = True
                bodyGroup.isEnabledCheckBoxDisplayed = False
                bodyInputs = bodyGroup.children
                
                for body in bodies:
                    safe_id = sanitize_id(body)
                    checkbox = bodyInputs.addBoolValueInput(f'body_{safe_id}', body, True, '', False)
                    checkbox.tooltip = f'Export body: {body}'
            
            # 7. Group for components
            if components:
                compGroup = inputs.addGroupCommandInput('compGroup', 'Components')
                compGroup.isExpanded = True
                compGroup.isEnabledCheckBoxDisplayed = False
                compInputs = compGroup.children
                
                for comp in components:
                    safe_id = sanitize_id(comp)
                    checkbox = compInputs.addBoolValueInput(f'comp_{safe_id}', comp, True, '', False)
                    checkbox.tooltip = f'Export component: {comp}'
            
            # Add spacer
            inputs.addTextBoxCommandInput('spacer3', '', '', 1, True)
            
            # 8. Output folder
            folderGroup = inputs.addGroupCommandInput('folderGroup', 'Output Folder')
            folderGroup.isExpanded = True
            folderGroup.isEnabledCheckBoxDisplayed = False
            folderInputs = folderGroup.children
            
            folderInputs.addTextBoxCommandInput('outputFolder', 'Folder Path', '', 2, True)
            folderInputs.addBoolValueInput('browseBtn', 'Browse...', False, '', False)
            
            # Connect input changed handler
            onInputChanged = MyInputChangedHandler()
            cmd.inputChanged.add(onInputChanged)
            handlers.append(onInputChanged)
            
        except:
            app = adsk.core.Application.get()
            ui = app.userInterface
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


class MyInputChangedHandler(adsk.core.InputChangedEventHandler):
    def __init__(self):
        super().__init__()
        self.isResetting = False
        
    def notify(self, args):
        try:
            input = args.input
            cmd = args.firingEvent.sender
            inputs = cmd.commandInputs
            
            # Check if browse button was clicked
            if input.id == 'browseBtn':
                if self.isResetting:
                    self.isResetting = False
                    return
                
                if input.value:
                    app = adsk.core.Application.get()
                    ui = app.userInterface
                    
                    folderDlg = ui.createFolderDialog()
                    folderDlg.title = 'Select Export Folder'
                    
                    if folderDlg.showDialog() == adsk.core.DialogResults.DialogOK:
                        folderGroup = inputs.itemById('folderGroup')
                        if folderGroup:
                            outputFolder = folderGroup.children.itemById('outputFolder')
                            if outputFolder:
                                outputFolder.text = folderDlg.folder
                    
                    self.isResetting = True
                    input.value = False
            
            # Show/hide mesh options based on format selection
            elif input.id == 'exportFormat':
                formatGroup = inputs.itemById('formatGroup')
                if formatGroup:
                    stlLabel = formatGroup.children.itemById('stlLabel')
                    meshRefinement = formatGroup.children.itemById('meshRefinement')
                    unitDropdown = formatGroup.children.itemById('exportUnit')
                    formatDropdown = formatGroup.children.itemById('exportFormat')
                    
                    if formatDropdown:
                        selectedFormat = formatDropdown.selectedItem.name
                        showMeshOptions = 'STL' in selectedFormat or '3MF' in selectedFormat or 'OBJ' in selectedFormat
                        showUnitOptions = 'STL' in selectedFormat or 'OBJ' in selectedFormat or '3MF' in selectedFormat
                        
                        if stlLabel:
                            stlLabel.isVisible = showMeshOptions
                        if meshRefinement:
                            meshRefinement.isVisible = showMeshOptions
                        if unitDropdown:
                            unitDropdown.isVisible = showUnitOptions
                    
        except:
            pass


class MyCommandExecuteHandler(adsk.core.CommandEventHandler):
    def __init__(self):
        super().__init__()
        
    def notify(self, args):
        try:
            app = adsk.core.Application.get()
            ui = app.userInterface
            design = adsk.fusion.Design.cast(app.activeProduct)
            
            inputs = args.command.commandInputs
            
            # Get selected parameter
            paramDropdown = inputs.itemById('paramSelect')
            selectedParamText = paramDropdown.selectedItem.name
            paramName = selectedParamText.split(' = ')[0]
            
            userParam = design.userParameters.itemByName(paramName)
            if not userParam:
                ui.messageBox('Parameter not found')
                return
            
            originalExpression = userParam.expression
            
            # Get parameter type
            paramType = inputs.itemById('paramType')
            isTextParam = paramType.selectedItem.index == 0
            
            # Get variations
            variationsInput = inputs.itemById('variations')
            variations = [v.strip() for v in variationsInput.text.split(',') if v.strip()]
            
            if not variations:
                ui.messageBox('No variations entered')
                return
            
            # Get naming options
            namingGroup = inputs.itemById('namingGroup')
            prefix = namingGroup.children.itemById('filePrefix').value if namingGroup else ''
            suffix = namingGroup.children.itemById('fileSuffix').value if namingGroup else ''
            addNumbering = namingGroup.children.itemById('addNumbering').value if namingGroup else False
            includeParamName = namingGroup.children.itemById('includeParamName').value if namingGroup else False
            
            # Get export format
            formatGroup = inputs.itemById('formatGroup')
            formatDropdown = formatGroup.children.itemById('exportFormat')
            selectedFormat = formatDropdown.selectedItem.name
            
            # Get export unit
            unitDropdown = formatGroup.children.itemById('exportUnit')
            selectedUnit = unitDropdown.selectedItem.name
            
            # Map unit selection to MeshUnits enum (FIXED)
            unitMap = {
                'Millimeters': adsk.fusion.MeshUnits.MillimeterMeshUnit,
                'Centimeters': adsk.fusion.MeshUnits.CentimeterMeshUnit,
                'Meters': adsk.fusion.MeshUnits.MeterMeshUnit,
                'Inches': adsk.fusion.MeshUnits.InchMeshUnit,
                'Feet': adsk.fusion.MeshUnits.FootMeshUnit
            }
            exportUnit = unitMap.get(selectedUnit, adsk.fusion.MeshUnits.MillimeterMeshUnit)
            
            # Get mesh refinement options if applicable
            meshRefinement = adsk.fusion.MeshRefinementSettings.MeshRefinementMedium
            if 'STL' in selectedFormat or '3MF' in selectedFormat or 'OBJ' in selectedFormat:
                meshDropdown = formatGroup.children.itemById('meshRefinement')
                if meshDropdown:
                    meshSelection = meshDropdown.selectedItem.name
                    if meshSelection == 'Low':
                        meshRefinement = adsk.fusion.MeshRefinementSettings.MeshRefinementLow
                    elif meshSelection == 'High':
                        meshRefinement = adsk.fusion.MeshRefinementSettings.MeshRefinementHigh
            
            # Determine file extension
            if 'STL' in selectedFormat:
                fileExt = '.stl'
            elif '3MF' in selectedFormat:
                fileExt = '.3mf'
            elif 'OBJ' in selectedFormat:
                fileExt = '.obj'
            elif 'STEP' in selectedFormat:
                fileExt = '.step'
            elif 'F3D' in selectedFormat:
                fileExt = '.f3d'
            else:
                fileExt = '.stl'
            
            # Get selected export objects
            selectedObjects = []
            rootComp = design.rootComponent
            
            bodyGroup = inputs.itemById('bodyGroup')
            if bodyGroup:
                for body in rootComp.bRepBodies:
                    safe_id = sanitize_id(body.name)
                    boolInput = bodyGroup.children.itemById(f'body_{safe_id}')
                    if boolInput and boolInput.value:
                        selectedObjects.append((body.name, body))
            
            compGroup = inputs.itemById('compGroup')
            if compGroup:
                for occ in rootComp.allOccurrences:
                    safe_id = sanitize_id(occ.name)
                    boolInput = compGroup.children.itemById(f'comp_{safe_id}')
                    if boolInput and boolInput.value:
                        selectedObjects.append((occ.name, occ))
            
            if not selectedObjects:
                ui.messageBox('No objects selected for export')
                return
            
            # Get output folder
            folderGroup = inputs.itemById('folderGroup')
            if not folderGroup:
                ui.messageBox('Could not find folder group')
                return
            
            outputFolderInput = folderGroup.children.itemById('outputFolder')
            if not outputFolderInput:
                ui.messageBox('Could not find output folder input')
                return
            
            outputFolder = outputFolderInput.text
            if not outputFolder or not os.path.exists(outputFolder):
                ui.messageBox('Please select a valid output folder')
                return
            
            # Confirm
            confirmMsg = f'Ready to export:\n\n'
            confirmMsg += f'Parameter: {paramName}\n'
            confirmMsg += f'Type: {"Text" if isTextParam else "Numeric"}\n'
            confirmMsg += f'Format: {selectedFormat}\n'
            confirmMsg += f'Unit: {selectedUnit}\n'
            confirmMsg += f'Variations: {len(variations)}\n'
            confirmMsg += f'Objects: {len(selectedObjects)}\n'
            confirmMsg += f'Total files: {len(variations) * len(selectedObjects)}\n\n'
            confirmMsg += 'Continue?'
            
            if ui.messageBox(confirmMsg, 'Confirm', adsk.core.MessageBoxButtonTypes.YesNoButtonType) != adsk.core.DialogResults.DialogYes:
                return
            
            # Progress dialog
            totalOperations = len(variations) * len(selectedObjects)
            progressDialog = ui.createProgressDialog()
            progressDialog.cancelButtonText = 'Cancel'
            progressDialog.isBackgroundTranslucent = False
            progressDialog.isCancelButtonShown = True
            progressDialog.show('Batch Export', 'Exporting %v of %m', 0, totalOperations)
            
            # Export
            exportMgr = design.exportManager
            successCount = 0
            currentProgress = 0
            
            for variantIdx, variant in enumerate(variations):
                if progressDialog.wasCancelled:
                    break
                
                try:
                    # Update parameter
                    if isTextParam:
                        newExpression = "'{}'".format(variant.replace("'", "\\'"))
                    else:
                        newExpression = str(variant)
                    
                    userParam.expression = newExpression
                    
                    # OPTIMIZED: Simplified regeneration
                    try:
                        design.computeAll()
                    except:
                        pass
                    
                    # Wait for compute
                    for _ in range(100):
                        adsk.doEvents()
                    
                    # Single viewport refresh
                    app.activeViewport.refresh()
                    
                    # Export each object
                    for objName, objEntity in selectedObjects:
                        if progressDialog.wasCancelled:
                            break
                        
                        currentProgress += 1
                        progressDialog.progressValue = currentProgress
                        progressDialog.message = f'Exporting: {variant} - {objName}'
                        
                        try:
                            # Build filename
                            safeVariant = "".join(c for c in variant if c.isalnum() or c in (' ', '-', '_', '.')).strip() or f'variant_{variantIdx+1}'
                            safeObj = "".join(c for c in objName if c.isalnum() or c in (' ', '-', '_', '.')).strip()
                            
                            filenameParts = []
                            
                            if prefix:
                                filenameParts.append(prefix)
                            
                            if addNumbering:
                                filenameParts.append(f'{variantIdx+1:03d}')
                            
                            if includeParamName:
                                filenameParts.append(paramName)
                            
                            filenameParts.append(safeVariant)
                            
                            if len(selectedObjects) > 1:
                                filenameParts.append(safeObj)
                            
                            if suffix:
                                filenameParts.append(suffix)
                            
                            filename = '_'.join(filenameParts) + fileExt
                            fullPath = os.path.join(outputFolder, filename)
                            
                            # Export based on format with selected unit
                            if 'STL' in selectedFormat:
                                stlOpts = exportMgr.createSTLExportOptions(objEntity, fullPath)
                                stlOpts.isBinaryFormat = 'Binary' in selectedFormat
                                stlOpts.meshRefinement = meshRefinement
                                stlOpts.sendToPrintUtility = False
                                stlOpts.unit = exportUnit
                                exportMgr.execute(stlOpts)
                                
                            elif '3MF' in selectedFormat:
                                mfOpts = exportMgr.createC3MFExportOptions(objEntity, fullPath)
                                mfOpts.meshRefinement = meshRefinement
                                # 3MF uses millimeters by default
                                exportMgr.execute(mfOpts)
                                
                            elif 'OBJ' in selectedFormat:
                                objOpts = exportMgr.createOBJExportOptions(objEntity, fullPath)
                                objOpts.meshRefinement = meshRefinement
                                objOpts.unit = exportUnit
                                exportMgr.execute(objOpts)
                                
                            elif 'STEP' in selectedFormat:
                                stepOpts = exportMgr.createSTEPExportOptions(fullPath, objEntity)
                                exportMgr.execute(stepOpts)
                                
                            elif 'F3D' in selectedFormat:
                                f3dOpts = exportMgr.createFusionArchiveExportOptions(fullPath)
                                exportMgr.execute(f3dOpts)
                            
                            successCount += 1
                            
                        except Exception as e:
                            pass
                            
                except:
                    currentProgress += len(selectedObjects)
            
            # Restore parameter
            try:
                userParam.expression = originalExpression
                try:
                    design.computeAll()
                except:
                    pass
                adsk.doEvents()
            except:
                pass
            
            progressDialog.hide()
            
            ui.messageBox(f'Done!\n\nExported {successCount} of {totalOperations} files')
            
        except:
            app = adsk.core.Application.get()
            ui = app.userInterface
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


class MyCommandDestroyHandler(adsk.core.CommandEventHandler):
    def __init__(self):
        super().__init__()
        
    def notify(self, args):
        adsk.terminate()
