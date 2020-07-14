if wscript.arguments.count < 2 then
wscript.echo "Missing arguments"
wscript.quit
end if 

Dim appRef
Set appRef = CreateObject( "Photoshop.Application" )
Dim Source
Source = Wscript.Arguments(0)
Dim Destination
Destination = Wscript.Arguments(1)
Dim baseImage
Dim sourceImage
 

Set baseImage = appRef.Open("{path to your source psd with actions in}.psd")
Set sourceImage = appRef.Open(Source)
' these are example calls to the named actions in the psd
Call appRef.DoAction("ChromaKeyAction","FolderName")
' change actibe doc
appRef.ActiveDocument = baseImage
' more actions
Call appRef.DoAction("Position","Folder2Name")
Call appRef.DoAction("Move","Folder2Name")
' export as jpeg
Set jpgSaveOptions = CreateObject("Photoshop.JPEGSaveOptions")
jpgSaveOptions.EmbedColorProfile = True
jpgSaveOptions.FormatOptions = 1 'for psStandardBaseline
jpgSaveOptions.Matte = 1 'for psNoMatte
'jpgSaveOptions.Quality = 1 goes too 12 higher better?
appRef.ActiveDocument.SaveAs Destination & "/" & sourceImage.Name,jpgSaveOptions, True, 2 'for psLowercase


sourceImage.Close(2) 'Close without saving
baseImage.Close(2) 'Close without saving



