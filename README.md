# VBA Script for Exporting PowerPoint Speaker Notes

Here's a VBA script which exports the speaker notes of each slide in an active PowerPoint presentation to individual text files.

```vba
Sub ExportSpeakerNotes_SlideNo()

    Dim ppt As Presentation
    Set ppt = ActivePresentation
    
    Dim sld As Slide
    Dim slideNumber As Integer
    Dim outputPath As String
    Dim stream As Object

    'Ask user where to save the txt files
    outputPath = InputBox("Enter the path where you want to store the text files (e.g. C:\MyTextFiles\)", "Output Path")
    
    'Add a backslash at the end of the path if it's not there
    If Right(outputPath, 1) <> "\" Then
        outputPath = outputPath & "\"
    End If
    
    'Create a new ADODB.Stream object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 'Specify stream type - we want To save text/string data.
    stream.Charset = "utf-8" 'Specify charset For the source text data.

    'Loop through each slide
    For Each sld In ppt.Slides
        slideNumber = sld.slideIndex
        
        'Save the speaker notes to a txt file
        stream.Open
        stream.WriteText sld.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.text
        stream.SaveToFile outputPath & "Slide" & slideNumber & ".txt", 2 '2 = adSaveCreateOverWrite
        stream.Close
    Next sld

    Set stream = Nothing 'Clean up

End Sub
