'''''''''''''''''''''''''''''''''''''''''''''''''''''
' Save Object, Chart, Or Shape As Image To Desktop  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'recieves input of objectWorksheet (ex. "Sheet 1") as string, objectName (ex. "Shape 1") as string, and imageFileName (ex. "foobar") as string
'outputs a .png image to the desktop of the specified shape, chart, or other object

Sub saveChartOrShapeAsImageToDesktop(objectWorksheet As String, objectName As String, imageFileName As String)

        'diminsion variables

        'find desktop
        Dim oWSHShell As Object: Set oWSHShell = CreateObject("WScript.Shell")
        Dim desktop as String: desktop = oWSHShell.SpecialFolders("Desktop")
        Set oWSHShell = Nothing

        'find worksheet
        Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("objectWorksheet")

        'find object
        Dim tempObject As Shape: ws.Shapes(objectName)

        'find file path
        Dim filePath As String: filePath = desktop & "\" & imageFileName & ".png"

        Dim temporaryChart As ChartObject
        
        'prevent screen updating
        Application.ScreenUpdating = False

        'copy object to memory
        tempObject.CopyPicture xlScreen, xlPicture
        
        'convert object to chart
        Set temporaryChart = ws.ChartObjects.Add(0, 0, tempObject.Width + 1, tempObject.Height + 1)

        'export chart
        With temporaryChart
                .Activate 'Required, otherwise image is blank with Excel 2016 or fast CPU (?)
                .Border.LineStyle = xlLineStyleNone 'No border
                .Chart.Paste
                .Chart.Export filePath
                .Delete
        End With
        
        'turn on screen updating
        Application.ScreenUpdating = True

        'garbage collection
        Set temporaryChart = Nothing
End Sub