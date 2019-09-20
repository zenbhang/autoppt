Imports Microsoft.Office.Interop
Imports Microsoft.Office.Core
Class MainWindow
    Private Sub BtnGenerate_Click(sender As Object, e As RoutedEventArgs) Handles btnGenerate.Click
        Dim dtDate As Date = dtDatePicker.DisplayDate
        Dim iMonth As Integer = dtDate.Date.Month
        Dim sFilePath As String = ""
        Dim filename As String = ""

        Dim oApp As PowerPoint.Application
        Dim oPres As PowerPoint.Presentation
        oApp = New PowerPoint.Application

        Dim iNum As Integer = 1
        Dim iRun As Integer = 1

        Try
            oPres = oApp.Presentations.Open(sFilePath)
            Do Until iNum > oPres.Slides.Count
                oPres.Slides(iNum).Select()

                Do Until iRun > oPres.Slides(iNum).Shapes.Count
                    If oPres.Slides(iNum).Shapes(iRun).HasTextFrame Then
                        If oPres.Slides(iNum).Shapes(iRun).TextFrame.HasText Then
                            Try
                                Dim sSearch = oPres.Slides(iNum).Shapes(iRun).TextFrame.TextRange.Find("[DATE]").Text.ToString()
                                If sSearch = "[DATE]" Then
                                    Dim iStart = oPres.Slides(iNum).Shapes(iRun).TextFrame.TextRange.Find("[DATE]").Characters.Start
                                    oPres.Slides(iNum).Shapes(iRun).TextFrame.TextRange.Characters(iStart).InsertBefore(Format(dtDate, "MM/dd"))
                                    oPres.Slides(iNum).Shapes(iRun).TextFrame.TextRange.Find("[DATE]").Delete()
                                End If
                            Catch exception As NullReferenceException

                            End Try
                        End If
                    End If
                    iRun += 1
                Loop
                iRun = 1
                iNum += 1
            Loop


            oPres.SaveAs(filename)
            oApp.Quit()
        Catch ex As Exception
        End Try
    End Sub
End Class
