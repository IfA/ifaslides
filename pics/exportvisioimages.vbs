'################################################
'
' This script can be used to export Visio files to PNG/JPG/SVG/PDF files.
'
' If a path to a visio file or a path to a folder containing visio files is given as command line argument,
' the list of visio files to be exported is determined by this argument instead of by the configuration below.
'
' The (global) export options can be configured in the Configuration Section below.
'
' Custom (per-page) export options can be specified by extending the page name. Options need to be placed in square
' brackets and separated by a ';'. Custom export options will always override the global export options. 

' Currently, the following custom options are supported:
' * no-export: The page will be ignored by this script.
' * png-export/jpg-export/pdf-export/svg-export: The page will be exported as png/jpg/pdf/svg (the global 'exportType' setting is ignored).
' * If the name of a sheet contains the string '[png-export]', this page will always be exported as png (the global 'exportType')
'	setting is ignored. This works the analogously with other supported export types (e.g. [pdf-export])
' * export-layers: Instead of generating a single image file for this page, each layer will be exported individually to
'   its own file (the layer name is appended to the page name). This is especially useful for generating
'   'animation'-like series of images for beamer presentations.
'	Note: The 'export-layers' option does not work if the 'pdf' export type is used!
' 
' A valid page name would thus e.g. be 'my-page [png-export;export-layers]'.
' 
' If the name of a sheet represents a path section (e.g. 'name\of\the\sheet'), it will be exported to a sub-folder (e.g.
' the file 'sheet' in the sub-folder 'name\of\the'). The sub-folder will be located below the export folder determined
' by the settings in this file (see below).
'
'################################################
Option Explicit 
Dim relativePathToFolderContainingVisioFiles, exportType, rasterExportResolution, relativePathToExportFolder, createSubFolderPerVisioFile, visioFileList, excludeList, customPathMappings, convertSVGToPDF, convertSVGToEPS, deleteTemporarySVGs, inkscapeShell, Visioapp


'################## Begin Configuration Section ##################

' the relative path (from the path this script is located in) to the folder containing the visio files to be exported
relativePathToFolderContainingVisioFiles = "."

' the list of visio files for which images shall be exported; if this is empty, images will be exported for all visio files in the folder
Set visioFileList = CreateObject("Scripting.Dictionary")
' example: Only export images from the file 'my-file.vsdx'
' visioFileList.Add "my-file.vsdx", ""

' the relative path (from the path this script is located in) to the folder into which the image files will be exported
relativePathToExportFolder = "."

' if the images shall be exported in a sub-folder 'relativePathToExportFolder/[visio-file-name]'
createSubFolderPerVisioFile = False 

' the type of the images to be exported; this needs to be one of 'png', 'jpg','pdf', or 'svg'
exportType = "pdf"

' If 'svg' is chosen as exportType and this is set to True, the exported SVG will also be converted to a PDF file;
' This might deliver better results than the PDF exporter built into Visio.
' Note: This uses Inkscape to convert the SVG to PDF so Inkscape must be present on the Path.
convertSVGToPDF = False

' If 'svg' is chosen as exportType and this is set to True, the exported SVG will also be converted to an EPS file;
' Note: This uses Inkscape to convert the SVG to EPS so Inkscape must be present on the Path.
convertSVGToEPS = False

' If 'svg' is chosen as exportType, convertSVGToPDF/EPS is set to True and this is set to True, the exported SVG will be
' deleted after being converted to a PDF file.
deleteTemporarySVGs = False

' if a raster format (png or jpg) is specified as export format, this specifies the export resolution in DPI
rasterExportResolution = "600"

' a list of path sections that will be included from the automatic export
Set excludeList = CreateObject("Scripting.Dictionary")
' example: All images whose path contains "path\to\exclude" will not be exported.
' excludeList.Add "path\to\exclude", ""

' custom path mappings can be used to export a specific (set of) image(s) to a custom location by replacing a section of its/their path(s)
Set customPathMappings = CreateObject("Scripting.Dictionary")
' example: The "original\path\section" will be replaced by the "custom\path\section" before exporting the image
' customPathMappings.Add "original\path\section", "custom\path\section"
' customPathMappings.Add "your\custom", "mappings"
' customPathMappings.Add "pamtram-metamodel\", "appendix\pamtram-metamodel\"
' customPathMappings.Add "pamtram-metamodel-structure\", "appendix\pamtram-metamodel\"
' customPathMappings.Add "pamtram-metamodel-mapping\", "appendix\pamtram-metamodel\"
' customPathMappings.Add "pamtram-metamodel-condition\", "appendix\pamtram-metamodel\"
' customPathMappings.Add "mapping-model-structure\", "mapping-model\"
' customPathMappings.Add "mapping-model-mapping\", "mapping-model\"
' customPathMappings.Add "mapping-model-condition\", "mapping-model\"
' customPathMappings.Add "mapping-model-constraint\", "mapping-model\"
' customPathMappings.Add "genlib\", "validation\genlib\"

'################## End Configuration Section ##################


Sub main()
	
	Dim VisioFiles, objshell, objFile, VisioFolder, FileNumber, VisioFile, flag

	' Check if a valid export type was specified
	If (StrComp(exportType, "png") <> 0 AND StrComp(exportType, "jpg") <> 0 AND StrComp(exportType, "pdf") <> 0 AND StrComp(exportType, "svg") <> 0) Then
		Wscript.Echo "Unknown export type '" & exportType & "'. Only 'png', 'jpg', 'pdf', and 'svg' are currently supported!"
		Wscript.Quit
	End If
	
	Set objshell = CreateObject("scripting.filesystemobject")
	Set objFile = objshell.GetFile(Wscript.ScriptFullName)

	Dim ArgCount
	ArgCount = WScript.Arguments.Count
	Select Case ArgCount
		' If no argument was specified, we browse the visio source folder resulting from the settings above...
		Case 0
			' Check if a valid source path (containing the visio files) was specified
			If Not objshell.FolderExists(objshell.GetParentFolderName(objFile) & "/" & relativePathToFolderContainingVisioFiles) Then
				Wscript.Echo "Folder '" & objshell.GetParentFolderName(objFile) & "/" & relativePathToFolderContainingVisioFiles & "' does not exist!"
				Wscript.Quit
			End If

			' Loop the files in the folder and export the image files for each visio file
			Set VisioFolder = objshell.GetFolder(objshell.GetParentFolderName(objFile) & "/" & relativePathToFolderContainingVisioFiles)
			Set VisioFiles = VisioFolder.Files
			
		' If one argument was specified, we use it as source visio file/folder ...
		Case 1 
			Dim VisioPaths
			VisioPaths = WScript.Arguments(0)

			' Check if the object is a folder
			If objshell.FolderExists(VisioPaths) Then
				Set VisioFolder = objshell.GetFolder(VisioPaths)
				Set VisioFiles = VisioFolder.Files	
			' Otherwise we assume that the object is a file
			Else 
				VisioFiles = Array(objshell.GetFile(VisioPaths))
			End If 
		Case Else 
	 		WScript.Echo "Unsupported number of arguments (only 0 or 1 allowed)!"
	 		WScript.Quit
	End Select

	On Error Resume Next

	' Create a global visio instance. Not creating a new instance for every visio file to be exported should save some resources.
	Set Visioapp = CreateObject("Visio.Application")

	If Err.Number <> 0 Then
		Wscript.Echo "Unable to open Visio application! (Fehlercode: " & Err.Number & ", Quelle: " & Err.Source & ", Beschreibung: " & Err.Description & ")"
		Wscript.Quit
	End If

	' Some general settings for the visio instance
	Visioapp.Visible = False
	Visioapp.AlertResponse = 7 ' Prevent the 'Do you want to save changes'-dialog by automatically responding 'no'
	Visioapp.Settings.SetRasterExportResolution 3, rasterExportResolution, rasterExportResolution, 0 ' Set the export resolution for PNG/JPG exports
	Visioapp.Settings.RasterExportQuality = 100 ' Set the export quality for JPG exports
	Visioapp.Settings.SVGExportFormat = 0 ' Do not include Visio data in SVG exports

	' Create a global inkscape shell if necessary. Not creating a new instance every time we convert a file
	' should save some time and resources
	If (StrComp(exportType, "svg") = 0 AND (convertSVGToPDF = True OR convertSVGToEPS = True)) Then
		Set inkscapeShell = WScript.CreateObject("WSCript.shell")
	End If

	' Loop the files and export every one separately
	For Each VisioFile In VisioFiles
		FileNumber = FileNumber + 1 
		VisioFile = VisioFile.Path
		If GetVisioFile(VisioFile) Then  'if the file is Visio file, then convert it
			If visioFileList.Count = 0 OR visioFileList.Exists(objshell.GetFileName(VisioFile)) = True Then 
				ExportVisioFile VisioFile, objshell.GetParentFolderName(objFile)
				flag = flag + 1
			End If
		End If 	
	Next

	' Quit the visio instance before finishing the script
	Visioapp.Quit

	On Error Goto 0
	
	Wscript.Echo "Image files successfully exported!"

End Sub 

' This function is to export a Visio file to PNG/PDF file.
' Copied from the original export script from https://gallery.technet.microsoft.com/office/How-to-export-multiple-6a80db79
' with custom changes and additions
'
Function ExportVisioFile(VisioFile, ParentFolder)  
	Dim objshell, BaseName, Visio, Pages

	' Open the file docked and read-only (necessary so that the script works even if the user has already opened the file)
	Set Visio = Visioapp.Documents.OpenEx(VisioFile, 3)
	Set Pages = Visioapp.ActiveDocument.Pages

	If Err.Number <> 0 Then
		Wscript.Echo "Unable to open Visio file '" & VisioFile & "'! Make sure that the file exists."
		Wscript.Quit
	End If

	' Enable diagram services
	'Visioapp.ActiveDocument.DiagramServicesEnabled = visServiceVersion140 + visServiceVersion150
	Visioapp.ActiveDocument.DiagramServicesEnabled = 7 + 8
	
	Set objshell = CreateObject("scripting.filesystemobject")
	BaseName = objshell.GetBaseName(VisioFile) 'Get the file name
	
	Dim Folder, intCounter, i, exclude, exportOptionsRegEx, customExportOptions

	' Determine the folder where the files will be exported to
	Folder = ParentFolder & "\" & relativePathToExportFolder
	If createSubFolderPerVisioFile = True Then
		Folder = Folder & "\" & GetFilenameWithoutExtension(objshell.GetFileName(VisioFile))
	End If

	' A regular expression that we will use to parse potential export options encoded in the page name (everything that
	' is placed in square brackets)
	Set exportOptionsRegEx = New RegExp
	exportOptionsRegEx.Pattern = "\[(.*)\]"

	' Export each page to its own image file
	For intCounter = 1 To Pages.Count
		Dim PageName, Page, ExportPath, ExportFolder, localExportType, layerIndex, vsoLayer
		Set Page = Pages.Item(intCounter)
		PageName = Replace(Page.Name, " ", "")

		' Collect the custom export options passed by the user in the page name (everything that
		' is placed in square brackets)
		Set customExportOptions = CreateObject("System.Collections.ArrayList")
		Dim match, exportOption
		For Each match in exportOptionsRegEx.Execute(PageName)
			For Each exportOption in Split(match.SubMatches(0), ";")
				customExportOptions.Add exportOption
			Next
		Next

		' Stript the PageName from the custom options
		'
		PageName = exportOptionsRegEx.Replace(PageName, "")

		' check if the user specified a custom export type that will overwrite the global export type
		localExportType = exportType
		If customExportOptions.Contains("png-export") Then
			localExportType = "png"
		ElseIf customExportOptions.Contains("jpg-export") Then
			localExportType = "jpg"
		ElseIf customExportOptions.Contains("pdf-export") Then
			localExportType = "pdf"
		ElseIf customExportOptions.Contains("svg-export") Then
			localExportType = "svg"
		End If

		ExportPath = Folder & "\" & PageName & "." & localExportType
	
		' Handle the list of excluded images
		exclude = False
		If customExportOptions.Contains("no-export") Then
			exclude = True
		Else
			For i = 0 To excludeList.Count - 1
				If (InStr(ExportPath, excludeList.Keys()(i))) Then
					exclude = True
					Exit For
				End If
			Next
		End If

		If exclude = False Then

			' Handle custom path mappings
			For i = 0 To customPathMappings.Count - 1
				ExportPath = Replace(ExportPath, customPathMappings.Keys()(i), customPathMappings.Items()(i))
			Next
			ExportFolder = objshell.GetParentFolderName(ExportPath)

			' Create folders if necessary
			If Not objshell.FolderExists(ExportFolder) Then
				GeneratePath(ExportFolder)
			End If

			' remove print margins and fit to drawing
			Page.PageSheet.OpenSheetWindow
			' Visioapp.ActiveWindow.Shape.CellsSRC(visSectionObject, visRowPrintProperties, visPrintPropertiesLeftMargin).FormulaU = "0"
			Visioapp.ActiveWindow.Shape.CellsSRC(1, 25, 0).FormulaU = "0"
			' Visioapp.ActiveWindow.Shape.CellsSRC(visSectionObject, visRowPrintProperties, visPrintPropertiesRightMargin).FormulaU = "0"
			Visioapp.ActiveWindow.Shape.CellsSRC(1, 25, 1).FormulaU = "0"
			' Visioapp.ActiveWindow.Shape.CellsSRC(visSectionObject, visRowPrintProperties, visPrintPropertiesTopMargin).FormulaU = "0"
			Visioapp.ActiveWindow.Shape.CellsSRC(1, 25, 3).FormulaU = "0"
			' Visioapp.ActiveWindow.Shape.CellsSRC(visSectionObject, visRowPrintProperties, visPrintPropertiesBottomMargin).FormulaU = "0"
			Visioapp.ActiveWindow.Shape.CellsSRC(1, 25, 2).FormulaU = "0"
			' Visioapp.ActiveWindow.Close
			Visioapp.ActivePage.ResizeToFitContents

			If customExportOptions.Contains("export-layers") AND Page.Layers.Count > 0 Then
				' Export each layer individually to its own file (the layer name is appended to the page name)

				' First, we hide all layers
				For Each vsoLayer In Page.Layers
					' vsoLayer.CellsC(visLayerVisible).FormulaU = "0"
					vsoLayer.CellsC(4).FormulaU = "0"
				Next

				' Then, we export each layer individually by temporarily making it visible
				For Each vsoLayer In Page.Layers
					' vsoLayer.CellsC(visLayerVisible).FormulaU = "1"
					vsoLayer.CellsC(4).FormulaU = "1"

					ExportPage Visioapp, Page, Replace(ExportPath, "." & localExportType, "-" & vsoLayer.Name & "." & localExportType), localExportType, objshell, inkscapeShell

					' vsoLayer.CellsC(visLayerVisible).FormulaU = "0"
					vsoLayer.CellsC(4).FormulaU = "0"
				Next
			Else
				' Export the page as single file
				ExportPage Visioapp, Page, ExportPath, localExportType, objshell, inkscapeShell
			End If

		End If
	Next

	If Err.Number <> 0 Then
		Wscript.Echo "Internal Error during export! (Fehlercode: " & Err.Number & ", Quelle: " & Err.Source & ", Beschreibung: " & Err.Description & ")"
		Visioapp.AlertResponse = 7 ' Prevent the 'Do you want to save changes'-dialog by automatically responding 'no'
		Visio.Close
		Visioapp.Quit
		Wscript.Quit
	End If

	' Close the file
	Visio.Close

	Set objshell = Nothing 
End Function 

' Exports the given Page of a Visio file to the given ExportPath using the given 'localExportType'
'
Function ExportPage(Visioapp, Page, ExportPath, localExportType, objshell, inkscapeShell)
	' Export the page to the specified format
	If (StrComp(localExportType, "png") = 0 OR StrComp(localExportType, "jpg") = 0 OR StrComp(localExportType, "svg") = 0) Then

		Page.Export(ExportPath)

	ElseIf (StrComp(localExportType, "pdf") = 0) Then
		'see https://msdn.microsoft.com/de-de/vba/visio-vba/articles/document-exportasfixedformat-method-visio
		'parameters not specified, are left out
		'the document tags for accessibility are disabled in order to remove black borders in exported version
		Visioapp.ActiveDocument.ExportAsFixedFormat 1, ExportPath, 1, 2,,,,,,False

	End If

	' Convert the exported svg to pdf if necessary
	If (StrComp(localExportType, "svg") = 0 AND (convertSVGToPDF = True OR convertSVGToEPS = True)) Then
		' Run the conversion synchronously (last parameter ist 'True'). Otherwise, the potential deletion in the next 
		' step might delete the SVG before it has been converted to PDF
		If (convertSVGToPDF = True) Then
			inkscapeShell.run "inkscape -D -f " & ExportPath & " -A " & Replace(ExportPath, ".svg", ".pdf"), 1, True
		End If	
		If (convertSVGToEPS = True) Then
			inkscapeShell.run "inkscape -D -f " & ExportPath & " -E " & Replace(ExportPath, ".svg", ".eps"), 1, True
		End If

		If (deleteTemporarySVGs = True) Then
			objshell.DeleteFile(ExportPath)
		End If
	End If
End Function

' Copied from the original export script from https://gallery.technet.microsoft.com/office/How-to-export-multiple-6a80db79
'
Function GetVisioFile(VisioFile) 'This function is to check if the file is a Visio file
	Dim objshell
	Set objshell= CreateObject("scripting.filesystemobject")
	Dim Arrs ,Arr
	Arrs = Array("vsdx","vssx","vstx","vxdm","vssm","vstm","vsd","vdw","vss","vst")
	
	Dim blnIsVisioFile,FileExtension
	blnIsVisioFile = False 
	FileExtension = objshell.GetExtensionName(VisioFile)  'Get the file extension
	For Each Arr In Arrs
		If InStr(UCase(FileExtension),UCase(Arr)) <> 0 Then 
			blnIsVisioFile= True
			Exit For 
		End If 
	Next 
	If InStr(FileExtension, "~") Then ' Disregard temporary files
		blnIsVisioFile = False
	End If
	GetVisioFile = blnIsVisioFile
	Set objshell = Nothing 

End Function 


' ---------------------------------------------------------------------
' Copied from https://www.pcreview.co.uk/threads/create-folder-in-vbscript.1469322/
'
'* @info Generate a folder tree from the path
'*
'* @param (String) Path
'* @return (Boolean) Folder Exists: Recursion continues (Y/N)
' ---------------------------------------------------------------------
Function GeneratePath(pFolderPath)
Dim objFSO
Set objFSO = CreateObject("scripting.filesystemobject")
GeneratePath = False
If Not objFSO.FolderExists(pFolderPath) Then
If GeneratePath(objFSO.GetParentFolderName(pFolderPath)) Then
GeneratePath = True
Call objFSO.CreateFolder(pFolderPath)
End If
Else
GeneratePath = True
End If
End Function

' ---------------------------------------------------------------------
' Copied from https://social.technet.microsoft.com/Forums/en-US/ebe19301-541a-412b-8e89-08c4263cc60b/get-filename-without-extension?forum=ITCG
' ---------------------------------------------------------------------
Function GetFilenameWithoutExtension(ByVal FileName)
  Dim Result, i
  Result = FileName
  i = InStrRev(FileName, ".")
  If ( i > 0 ) Then
    Result = Mid(FileName, 1, i - 1)
  End If
  GetFilenameWithoutExtension = Result
End Function

Call main 