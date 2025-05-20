// VBA script created to convert a Excel table into a Visio flowchart

Sub Teste()
    Dim visApp As Object
    Dim visDoc As Object
    Dim visPage As Visio.Page
    Dim visStencil As Object
    Dim visStencil1 As Object
    Dim shapeDict As Object
    Dim shapeDictIn As Object
    Dim stepID As String, stepText As String, nextStep As String
    Dim lastRow As Long
    Dim iterator As Double, var As Double
    Dim proximaEtapa As String
    Dim resultado As Variant
    Dim stencilPath As String
    Dim sheetName As String
    Dim i As Integer
    Dim iter As Integer
    Dim sheetCount As Integer
    Dim p As Integer
    Dim inicio As Object
    Dim visValidationRuleSet As Visio.ValidationRuleSet
    Dim DiagramServices As Integer
    
    ' Counting the amount of tabs on the sheet
    sheetCount = ThisWorkbook.Sheets.Count
    
    For p = 1 To sheetCount
        ' Obtaining the name of the tab
        sheetName = ThisWorkbook.Sheets(p).Name
    
        ' Starting the Visio application and documents
        Set visApp = CreateObject("Visio.Application")
        Set visDoc = visApp.Documents.Add("")
        Set visPage = visDoc.Pages(1)
        
       ' Accessing favorite shapes
        stencilPath = "C:\Path\Favoritos (1).vssx"
        If stencilPath = "False" Then
            MsgBox "Operation canceled."
            Exit Sub
        End If
    
    
    
        ' Opening stencils
        Set visStencil = visApp.Documents.OpenEx("BASFLO_U.VSSX", 64)
        Set visStencil1 = visApp.Documents.OpenEx(stencilPath, visOpenRO)
    
        ' Counting the number of rows on column b, because it will always be filled
        lastRow = ThisWorkbook.Sheets(sheetName).Cells(Rows.Count, "B").End(xlUp).Row
    
        ' Creating a dictionary
        Set shapeDict = CreateObject("Scripting.Dictionary")
        
        Set shapeDictIn = CreateObject("Scripting.Dictionary")
        
        
        ' The main loop of the interaction
        var = 1.5
        iter = 2
        ' Colocar a forma de início
            iterator = 5

            
            
        For i = 2 To lastRow
            
            stepID = ThisWorkbook.Sheets(sheetName).Cells(i, 1).Value
            stepText = ThisWorkbook.Sheets(sheetName).Cells(i, 2).Value
            proximaEtapa = ThisWorkbook.Sheets(sheetName).Cells(i, 6).Value
    
            ' Creating a figure and adding it into the dictionary
            Call figures(visApp, stepText, i, iter, iterator, var, sheetName, visStencil1, visPage, visStencil, shapeDict, stepID, lastRow, inicio, proximaEtapa)
            
            Call comentarios(sheetName, i, iter, var, visStencil1, visPage)
            
            ' Checking the next step
            resultado = verificarProx(proximaEtapa, visStencil1, iter, iterator, var, visPage, stepID)
            iterator = resultado(0)
            var = resultado(2)
            iter = resultado(1)
            iter = iter + 1
                
            
            
        Next i
        


visPage.AutoSizeDrawing

        Dim visPage1 As Visio.Page

        Dim visMaster As Object
        
        Dim shape2 As Visio.Shape
        Dim sheetWidth As Double
        Dim sheetHeight As Double
        Dim originX As Double
        Dim originY As Double
        Dim visDoc1 As Visio.Document
        Dim visLayer As Visio.Layer
        
        Dim visStencilBPMN As Visio.Document
        Dim entireProcess As Visio.Shape
        Dim vsoLayers As Visio.Layers
        
        Set visStencilBPMN = visApp.Documents.OpenEx("BPMN_M.vssx", visOpenDocked)
        
        originX = visPage.PageSheet.CellsU("PageWidth").ResultIU
        originY = visPage.PageSheet.CellsU("PageHeight").ResultIU
        
        Set entireProcess = visPage.Drop(visStencilBPMN.Masters.ItemU("Pool / Lane"), 15, 17)
        
        entireProcess.Delete
        
        Set vsoLayers = visPage.Layers
        
        Set visLayer = vsoLayers.Add("TargetLayer")
        Set visDoc1 = visApp.ActiveDocument
        Set shape2 = visPage.Drop(visDoc1.Masters.ItemU("Pool / Lane"), 5, 17)
        
        visLayer.Add shape2, 1

        sheetWidth = visPage.PageSheet.CellsU("PageWidth").ResultIU
        sheetHeight = visPage.PageSheet.CellsU("PageHeight").ResultIU
        
        shape2.Cells("LockTextEdit").FormulaU = "FALSE"
        shape2.Cells("LockWidth").FormulaU = "FALSE"
        shape2.Cells("LockDelete").FormulaU = "FALSE"
        Dim smallShape As Visio.Shape
        Dim j As Integer
        j = 0
        
        For Each smallShape In shape2.Shapes
            
            smallShape.CellsU("Height").FormulaForceU = ""
                smallShape.CellsU("Angle").FormulaForceU = ""
                smallShape.CellsU("PinX").FormulaForceU = ""
                smallShape.CellsU("PinY").FormulaForceU = ""
                smallShape.CellsU("LocPinX").FormulaForceU = ""
                smallShape.CellsU("LocPinY").FormulaForceU = ""
                smallShape.CellsU("FlipX").FormulaForceU = ""
                smallShape.CellsU("FlipY").FormulaForceU = ""
                smallShape.CellsU("ResizeMode").FormulaForceU = ""
                smallShape.CellsU("LockWidth").FormulaForceU = "0"
                smallShape.CellsU("LockGroup").FormulaForceU = "0"
                smallShape.CellsU("LockAspect").FormulaForceU = "0"
                smallShape.CellsU("LockMoveX").FormulaForceU = "0"
                smallShape.CellsU("LockMoveY").FormulaForceU = "0"
                smallShape.CellsU("Width").FormulaForceU = ""
  
        Next smallShape
        
        shape2.CellsU("Height").FormulaForceU = ""
        shape2.CellsU("Angle").FormulaForceU = ""
        shape2.CellsU("LocPinX").FormulaForceU = ""
        shape2.CellsU("LocPinY").FormulaForceU = ""
        shape2.CellsU("FlipX").FormulaForceU = ""
        shape2.CellsU("FlipY").FormulaForceU = ""
        shape2.CellsU("ResizeMode").FormulaForceU = ""
        shape2.CellsU("LockWidth").FormulaForceU = "0"
        shape2.CellsU("LockGroup").FormulaForceU = "0"
        shape2.CellsU("LockAspect").FormulaForceU = "0"
        shape2.CellsU("Width").FormulaForceU = ""
        shape2.CellsU("Height").FormulaForceU = ""
        
        shape2.CellsU("LockMoveX").FormulaForceU = "0"
        shape2.CellsU("LockMoveY").FormulaForceU = "0"
        
        ' Name of the tab
        
        shape2.Text = sheetName
        
        
        shape2.Cells("LockAspect").FormulaU = "FALSE"
        shape2.Cells("LockReplace").FormulaU = "FALSE"
        shape2.Cells("LockRotate").FormulaU = "FALSE"
        shape2.Cells("LockGroup").FormulaU = "FALSE"
        shape2.Cells("LockCalcWH").FormulaU = "FALSE"
        shape2.Cells("ReplaceLockShapeData").FormulaU = "TRUE"
        
        shape2.CellsU("Height").ResultIU = (sheetHeight + 6) / 2
        shape2.CellsU("Width").ResultIU = 3
        
        
        ' Cleaning and finishing the code
    visStencil.Close
    visStencil1.Close
    visPage.ResizeToFitContents

    Set visStencil = Nothing
    Set visStencil1 = Nothing
    Set visPage = Nothing
    Set visDoc = Nothing
    Set visApp = Nothing
        
  Next p
        
        

   

    MsgBox "Process finished successfully!"
End Sub

Function verificarProx(proximaEtapa As String, visStencil1 As Object, iter As Integer, iterator As Double, var As Double, visPage As Object, stepID As String) As Variant

    If InStr(proximaEtapa, "FIM") > 0 Then
        iterator = iterator - 7
        iter = 1
        var = -5.5
        verificarProx = Array(iterator, iter, var)
    End If
    verificarProx = Array(iterator, iter, var)
End Function

Sub figures(visApp As Object, stepText As String, i As Integer, iter As Integer, iterator As Double, var As Double, sheetName As String, visStencil1 As Object, visPage As Object, visStencil As Object, shapeDict As Object, stepID As String, lastRow As Long, inicio As Object, proximaEtapa As String)
    Dim visShape As Object
         Dim vInicio As Object
         Dim fim As Object
         Dim visStencilBPMN As Visio.Document
        Dim entireProcess As Visio.Shape
        Dim shapeBPMN As Visio.Shape
        Dim l As Integer
        
        If InStr(stepText, "?") > 0 Then
        Set visShape = visPage.Drop(visStencil.Masters.ItemU("Decision"), 2 * iter, iterator)
            visShape.Text = stepText

        Else
            Set visShape = visPage.Drop(visStencil1.Masters.ItemU("Atividade"), 2 * iter, iterator)
            For Each formas In visShape.Shapes
                If formas.Text = "Função" Then
                    formas.Text = ""
                    formas.Text = ThisWorkbook.Sheets(sheetName).Cells(i, 5).Value
                    dummy = ThisWorkbook.Sheets(sheetName).Cells(i, 5).Value
                ElseIf formas.Text = "atividade" Then
                formas.Text = ""
                    formas.Text = stepText
                    dummy = stepText
                Else
                    formas.Text = ""
                    formas.Text = ThisWorkbook.Sheets(sheetName).Cells(i, 3).Value
                    dummy = ThisWorkbook.Sheets(sheetName).Cells(i, 3).Value
                End If
            Next formas

        
End If
visShape.Cells("Width").Formula = "GUARD(""Width"" + 0.1)"
        visShape.Cells("Height").Formula = "GUARD(""Height"" + 0.1)"
        
        shapeDict.Add stepID, visShape
        If i = lastRow Then
        
            Call conectores(sheetName, shapeDict, visStencil, visPage, lastRow, inicio, visApp)
            End If
        
        If stepID Like "*[A-Za-z]*" Or stepID = "1" Then
        Set vInicio = visPage.Drop(visStencil1.Masters.ItemU("Início"), 0, iterator)
        Set visStencilBPMN = visApp.Documents.OpenEx("BPMN_M.vssx", visOpenDocked)
        Set visConnector = visPage.Drop(visStencilBPMN.Masters.ItemU("Sequence Flow"), 0, 0)
        visConnector.CellsU("BeginX").GlueTo vInicio.CellsU("PinX")
        visConnector.CellsU("EndX").GlueTo visShape.CellsU("PinX")
        Set visStencilBPMN = Nothing
    End If
    
    If InStr(proximaEtapa, "FIM") > 0 Then
    Set visStencilBPMN = Nothing
        Set fim = visPage.Drop(visStencil1.Masters.ItemU("Fim"), 2 * (iter + 1), iterator)
        Set visStencilBPMN = visApp.Documents.OpenEx("BPMN_M.vssx", visOpenDocked)
        visStencilBPMN.Protection = False
        
        Set visConnector = visPage.Drop(visStencilBPMN.Masters.ItemU("Sequence Flow"), 0, 0)
        visConnector.CellsU("BeginX").GlueTo visShape.CellsU("PinX")
        visConnector.CellsU("EndX").GlueTo fim.CellsU("PinX")
 
        
    End If
    

End Sub

Sub comentarios(sheetName As String, i As Integer, iter As Integer, var As Double, visStencil1 As Object, visPage As Object)
    Dim comments As String
    Dim visTextBox As Object
    comments = ThisWorkbook.Sheets(sheetName).Cells(i, 7).Value

    If comments <> "" Then
        Set visTextBox = visPage.Drop(visStencil1.Masters.ItemU("Comentários"), 2 * iter, var)
        visTextBox.Text = comments
        visTextBox.Cells("Width").FormulaU = "1.5 in"
        visTextBox.Cells("Height").FormulaU = "5 in"
        
            
    visTextBox.Cells("LockWidth").FormulaU = "0"
    visTextBox.Cells("LockHeight").FormulaU = "0"
    Else
            Set visTextBox = Nothing
    End If
    
End Sub

Sub conectores(sheetName As String, shapeDict As Object, visStencil As Object, visPage As Object, lastRow As Long, inicio As Object, visApp As Object)
    Dim k As Integer
    Dim prevShape As Object
    Dim visShape1 As Object
    Dim visConnector As Object
    Dim firstPart As String
    Dim secondPart As String
    Dim stepID As String
    Dim nextStep As String
    Dim stepText As String
    Dim visStencilBPMN As Object
    Dim lastSpacePos As Long
    Dim position As Long

    ' Creating connectors
    For k = 2 To lastRow
        stepID = ThisWorkbook.Sheets(sheetName).Cells(k, 1).Value
        nextStep = ThisWorkbook.Sheets(sheetName).Cells(k, 6).Value
        stepText = ThisWorkbook.Sheets(sheetName).Cells(k, 2).Value

        
        If shapeDict.Exists(stepID) And shapeDict.Exists(nextStep) Then
                Set prevShape = shapeDict(stepID)
                Set visShape1 = shapeDict(nextStep)
                
                Set visStencilBPMN = visApp.Documents.OpenEx("BPMN_M.vssx", visOpenDocked)
                Set visConnector = visPage.Drop(visStencilBPMN.Masters.ItemU("Sequence Flow"), 0, 0)
                visConnector.CellsU("BeginX").GlueTo prevShape.CellsU("PinX")
                visConnector.CellsU("EndX").GlueTo visShape1.CellsU("PinX")
                ' visConnector.Cells("EndArrow").FormulaU = "13"
                ' visConnector.Cells("EndArrowSize").FormulaU = "2"

            ElseIf InStr(nextStep, " ou ") > 0 Then
            lastSpacePos = InStr(1, nextStep, "ou ", vbTextCompare)
            position = InStr(1, nextStep, " ou ", vbTextCompare)
                firstPart = Left(nextStep, position - 1)
                
                
                Set prevShape = shapeDict(stepID)
                Set visShape1 = shapeDict(firstPart)
    
                Set visConnector = visPage.Drop(visStencil.Masters.ItemU("Dynamic Connector"), 0, 0)
                visConnector.CellsU("BeginX").GlueTo prevShape.CellsU("PinX")
                visConnector.CellsU("EndX").GlueTo visShape1.CellsU("PinX")
    
    
                visConnector.Cells("EndArrow").FormulaU = "13"
                visConnector.Cells("EndArrowSize").FormulaU = "2"
                
                
    
            secondPart = Mid(nextStep, lastSpacePos + Len("ou "))
                
                
                Set prevShape = shapeDict(stepID)
                Set visShape1 = shapeDict(secondPart)
    
                Set visConnector = visPage.Drop(visStencil.Masters.ItemU("Dynamic Connector"), 0, 0)
                visConnector.CellsU("BeginX").GlueTo prevShape.CellsU("PinX")
                visConnector.CellsU("EndX").GlueTo visShape1.CellsU("PinX")
    
    
                visConnector.Cells("EndArrow").FormulaU = "13"
                visConnector.Cells("EndArrowSize").FormulaU = "2"
        End If
        
        

    Next k

End Sub
