Imports System.IO
Imports System.Xml
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml.Packaging
Public Class Form1
    Private Sub btnConvert_Click(sender As Object, e As EventArgs) Handles btnConvert.Click

        Dim source As String
        Dim dest As String
        Try
            source = txtSource.Text
            dest = txtDestination.Text
            If checkFiles() Then

                Dim xlsData = readXML(source)

                creaFileExcel(dest, xlsData)
                MsgBox("Done! Exported " & xlsData.Count & " rows.")
            End If
        Catch ex As Exception
            MsgBox("Exception: " & ex.Message)
        End Try



    End Sub

    Private Function checkFiles() As Boolean
        Dim filesOk As Boolean = True
        If txtSource.Text.Trim.Length = 0 Then
            filesOk = False
            MsgBox("Select Source File")
        End If
        If txtDestination.Text.Trim.Length = 0 Then
            filesOk = False
            MsgBox("Select Destination File")
        End If


        Return filesOk
    End Function
    Private Function readXML(ByVal source As String) As List(Of strXlsRow)

        Dim result As New List(Of strXlsRow)
        Dim xmldoc As New XmlDataDocument()
        Dim xmlnode As XmlNodeList
        Dim i As Integer
        Dim str As String
        Dim fs As New FileStream(source, FileMode.Open, FileAccess.Read)
        xmldoc.Load(fs)
        xmlnode = xmldoc.GetElementsByTagName("App")

        For Each element As XmlNode In xmlnode
            Dim newXlsRow As New strXlsRow

            newXlsRow.appAction = element.Attributes("action").Value
            newXlsRow.appId = element.Attributes("id").Value


            For Each child As XmlNode In element.ChildNodes
                Select Case child.Name
                    Case "BaseVehicle"
                        newXlsRow.BaseVehicleId = child.Attributes("id").Value
                    Case "Note"
                        newXlsRow.Note = child.InnerText
                        Dim temp() As String = newXlsRow.Note.Split(";")
                        For Each item In temp
                            If item.Contains("OEM") Then
                                Dim tempOEM() As String = item.Split(":")
                                newXlsRow.OEM = tempOEM(1)
                            End If
                        Next
                    Case "Qty"
                        newXlsRow.Qty = child.InnerText
                    Case "PartType"
                        newXlsRow.PartTypeId = child.Attributes("id").Value
                    Case "Position"
                        newXlsRow.PositionId = child.Attributes("id").Value
                    Case "Part"
                        newXlsRow.Part = child.InnerText
                End Select
            Next
            result.Add(newXlsRow)
        Next
        Return result

        'For i = 0 To xmlnode.Count - 1
        '    xmlnode(i).ChildNodes.Item(0).InnerText.Trim()
        '    str = xmlnode(i).ChildNodes.Item(0).InnerText.Trim() & "  " & xmlnode(i).ChildNodes.Item(1).InnerText.Trim() & "  " & xmlnode(i).ChildNodes.Item(2).InnerText.Trim()
        '    MsgBox(str)
        'Next
    End Function




    Structure strXlsRow
        Dim appId As Int32
        Dim appAction As String
        Dim BaseVehicleId As Int32
        Dim Note As String
        Dim OEM As String
        Dim Qty As Int32
        Dim PartTypeId As Int32
        Dim PositionId As Int32
        Dim Part As String

    End Structure




    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub creaFileExcel(ByVal destination As String, ByVal xmldata As List(Of strXlsRow))
        Dim ultimaFotografia As Long = 0
        Dim fileName As String = ""


        ' fileName = Application.StartupPath() & "\" & ultimaFotografia.ToString & ".xlsx"
        fileName = destination

        Me.CreateBasicWorkbook(fileName, True, xmldata)

    End Sub


    ''' <summary>
    ''' Creates a basic workbook
    ''' </summary>
    ''' <param name="workbookName">Name of the workbook</param>
    ''' <param name="createStylesInCode">Create the styles in code?</param>
    Private Sub CreateBasicWorkbook(workbookName As String, createStylesInCode As Boolean, ByVal data As List(Of strXlsRow))
        Dim spreadsheet As DocumentFormat.OpenXml.Packaging.SpreadsheetDocument
        Dim worksheet As DocumentFormat.OpenXml.Spreadsheet.Worksheet
        Dim styleXml As String

        spreadsheet = Excel.CreateWorkbook(workbookName)
        If (spreadsheet Is Nothing) Then
            Return
        End If

        If (createStylesInCode) Then
            Excel.AddBasicStyles(spreadsheet)
        Else
            Using styleXmlReader As System.IO.StreamReader = New System.IO.StreamReader("PredefinedStyles.xml")
                styleXml = styleXmlReader.ReadToEnd()
                Excel.AddPredefinedStyles(spreadsheet, styleXml)
            End Using
        End If

        Excel.AddWorksheet(spreadsheet, "Sheet 1")

        worksheet = spreadsheet.WorkbookPart.WorksheetParts.First().Worksheet

        addHeader(spreadsheet, worksheet)
        For i As Int16 = 0 To data.Count - 1
            addExcelRow(spreadsheet, worksheet, i + 2, data(i))

        Next





        worksheet.Save()
        spreadsheet.Close()

    End Sub


    Private Sub addHeader(ByVal spreadsheet As DocumentFormat.OpenXml.Packaging.SpreadsheetDocument, worksheet As DocumentFormat.OpenXml.Spreadsheet.Worksheet)
        Excel.SetStringCellValue(spreadsheet, worksheet, 1, 1, "Action", False, False)
        Excel.SetStringCellValue(spreadsheet, worksheet, 2, 1, "App ID", False, False)
        Excel.SetStringCellValue(spreadsheet, worksheet, 3, 1, "BaseVehicle", False, False)
        Excel.SetStringCellValue(spreadsheet, worksheet, 4, 1, "Note", False, False)
        Excel.SetStringCellValue(spreadsheet, worksheet, 5, 1, "OEM", False, False)
        Excel.SetStringCellValue(spreadsheet, worksheet, 6, 1, "Qty", False, False)
        Excel.SetStringCellValue(spreadsheet, worksheet, 7, 1, "Parttype ID", False, False)
        Excel.SetStringCellValue(spreadsheet, worksheet, 8, 1, "Position ID", False, False)
        Excel.SetStringCellValue(spreadsheet, worksheet, 9, 1, "Part", False, False)

    End Sub

    Private Sub addExcelRow(ByVal spreadsheet As DocumentFormat.OpenXml.Packaging.SpreadsheetDocument, worksheet As DocumentFormat.OpenXml.Spreadsheet.Worksheet, rowNumber As Int16, row As strXlsRow)


        Excel.SetStringCellValue(spreadsheet, worksheet, 1, rowNumber, row.appAction, Nothing, False)
        Excel.SetDoubleCellValue(spreadsheet, worksheet, 2, rowNumber, row.appId, Nothing, False)
        Excel.SetDoubleCellValue(spreadsheet, worksheet, 3, rowNumber, row.BaseVehicleId, False, False)
        Excel.SetStringCellValue(spreadsheet, worksheet, 4, rowNumber, row.Note, Nothing, False)
        Excel.SetStringCellValue(spreadsheet, worksheet, 5, rowNumber, row.OEM, Nothing, False)
        Excel.SetDoubleCellValue(spreadsheet, worksheet, 6, rowNumber, row.Qty, False, False)
        Excel.SetDoubleCellValue(spreadsheet, worksheet, 7, rowNumber, row.PartTypeId, Nothing, False)
        Excel.SetDoubleCellValue(spreadsheet, worksheet, 8, rowNumber, row.PositionId, Nothing, False)
        Excel.SetStringCellValue(spreadsheet, worksheet, 9, rowNumber, row.Part, Nothing, False)




    End Sub

    Private Sub btnSelectSource_Click(sender As Object, e As EventArgs) Handles btnSelectSource.Click

        OpenFileDialog1.ShowDialog()
        txtSource.Text = OpenFileDialog1.FileName

    End Sub

    Private Sub btnSelectDestination_Click(sender As Object, e As EventArgs) Handles btnSelectDestination.Click
        SaveFileDialog1.ShowDialog()
        txtDestination.Text = SaveFileDialog1.FileName
    End Sub
End Class
