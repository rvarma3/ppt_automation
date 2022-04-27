Sub exltoppt_v2_1()


    Dim rng  As Range
    
    
    Dim PPApp As PowerPoint.Application
    Dim PPPres As PowerPoint.Presentation
    Dim PPactiveSlide As PowerPoint.Slide
    Dim PPnewSlide As PowerPoint.Slide
    'Dim PPSlide2 As PowerPoint.Slide
    Dim myShape As Object
    Dim shpCurrShape1 As Object
    Dim shpCurrShape2 As Object
    Dim MyPic As Shape
    Dim i As Integer
    'Dim testing As Range
    'Dim fileNameString As String
    
    'Application.ScreenUpdating = False
    
    ' fileNameString = Application.ActiveWorkbook.Path & _
    ' Format(Date, "YYYYMMDD") & " Skywards Forward Ticketed Report.pdf"
    
    
'    fileNameString = "C:\Users\s456781\OneDrive - Emirates Group\Documents\Coro Testing\" & _
 '   Format(Date, "YYYYMMDD") & " Skywards Forward Ticketed Report.pdf"
    
    fileNameString = "C:\Users\s456781\OneDrive - Emirates Group\CORO Automation\Solution" & _
    "CORO" & Format(Date, "YYYYMMDD") & " Skywards Forward Ticketed Report.pdf"

    
    fileNameString2 = Application.ActiveWorkbook.Path & _
    "\CORO" & Format(Date, "YYYYMMDD") & " Skywards Forward Ticketed Report.pptx"

    
    'copy the ranges
    
    
        ' Set rng1 = ThisWorkbook.Sheets(1).Range("b5:ax40")
        ' Set rng2 = ThisWorkbook.Sheets(2).Range("b5:ax40")
        ' Set rng3 = ThisWorkbook.Sheets(3).Range("b5:dx40")
        ' Set rng4 = ThisWorkbook.Sheets(4).Range("b5:dw35")
        ' Set rng5 = ThisWorkbook.Sheets(5).Range("b5:dw29")
        ' Set rng6 = ThisWorkbook.Sheets(6).Range("b5:ax40")
        ' Set rng7 = ThisWorkbook.Sheets(7).Range("b5:ax40")
        ' Set rng8 = ThisWorkbook.Sheets(8).Range("a5:dv35")
        ' Set rng9 = ThisWorkbook.Sheets(9).Range("a5:dv29")
        
        
        
               
        
        
    ' create a new pp application if not created already
        
        On Error Resume Next
            Set PPApp = GetObject(class:="PowerPoint.Application")
        On Error GoTo 0

   
        If PPApp Is Nothing Then
        Set PPApp = CreateObject(class:="PowerPoint.Application")
        End If
        
   PPApp.Visible = True
'PPApp.Activate
        
    'add a presentation to the application
    
    'Make a presentation in PowerPoint if does not already exist
        
        If PPApp.Presentations.Count = 0 Then
            Set PPPres = PPApp.Presentations.Add
        End If
        
        Set PPnewSlide = PPApp.ActivePresentation.Slides.Add(1, ppLayoutBlank)
        Set PPactiveSlide = PPApp.ActivePresentation.Slides(1)
        Set MyPic = ThisWorkbook.Sheets(19).Shapes("Picture 9")
        
        MyPic.Copy
        
        PPactiveSlide.Shapes.Paste
        
        Set myShape = PPactiveSlide.Shapes(PPactiveSlide.Shapes.Count)
        
        With myShape

                'size:
                ''1 inch = 72 points
                '.Height = 72 * 3.39
                '.Width = 72 * 6.67

                .ScaleHeight 0.98, msoTrue
                .ScaleWidth 0.98, msoTrue

                .LockAspectRatio = msoTrue


                'position:
                .Rotation = 0

                .Left = 1.5
                .Top = 0

                'Relative to original picture size = true

        End With
        
        
        With PPactiveSlide
        
                If Not .Shapes.HasTitle Then
                    Set shpCurrShape1 = .Shapes.AddTextbox(1, 312, 250, 600, 29)
                Else
                    Set shpCurrShape1 = .Shapes.Title
                End If
                
                End With
                
               
                        With shpCurrShape1
                        With .TextFrame.TextRange
                            '~~> Set text here
                            .Text = "SKYWARDS | WEEKLY" _
                            & vbNewLine & "" _
                            & vbNewLine & "TICKETED INSIGHTS" _
                            & vbNewLine & "" _
                            & vbNewLine & "EBI SNAP: " & _
                            Format(Date - 2, "DD-MMM-YYYY")
                            '~~> Alignment
                            .ParagraphFormat.Alignment = 3
                           '~~> Working with font
                           With .Font
                              .Italic = msoTrue
                              .Bold = msoFalse
                              .Name = "Heroic Condensed Medium"
                              .Size = 60
                              .Color = RGB(89, 89, 89)
                           End With
                        
                           With .Lines(2).Font
                              .Italic = msoTrue
                              .Bold = msoTrue
                              .Name = "Heroic Condensed Medium"
                              .Size = 14
                              .Color = RGB(89, 89, 89)
                           End With
                           
                           With .Lines(3).Font
                              .Italic = msoTrue
                              .Bold = msoFalse
                              .Name = "Heroic Condensed Medium"
                              .Size = 60
                              .Color = RGB(89, 89, 89)
                           End With
                           
                           With .Lines(4).Font
                              .Italic = msoTrue
                              .Bold = msoTrue
                              .Name = "Heroic Condensed Medium"
                              .Size = 14
                              .Color = RGB(89, 89, 89)
                           End With
                           
                            With .Lines(5).Font
                              .Italic = msoTrue
                              .Bold = msoFalse
                              .Name = "Heroic Condensed Medium"
                              .Size = 36
                              .Color = RGB(89, 89, 89)
                           End With
                           
                        End With
                    End With
       
        
        
        
         For i = 1 To 11
        
                If i = 1 Then
                    Set rng = ThisWorkbook.Sheets(1).Range("b5:ax40")
                ElseIf i = 2 Then
                    Set rng = ThisWorkbook.Sheets(2).Range("b5:ax40")
                ElseIf i = 3 Then
                    Set rng = ThisWorkbook.Sheets(3).Range("b5:dx40")
                ElseIf i = 4 Then
                    Set rng = ThisWorkbook.Sheets(4).Range("b5:dw35")
                ElseIf i = 5 Then
                    Set rng = ThisWorkbook.Sheets(5).Range("b5:dw29")
                ElseIf i = 6 Then
                    Set rng = ThisWorkbook.Sheets(6).Range("b5:ax40")
                ElseIf i = 7 Then
                    Set rng = ThisWorkbook.Sheets(7).Range("b5:ax40")
                ElseIf i = 8 Then
                    Set rng = ThisWorkbook.Sheets(8).Range("a5:dv35")
                ElseIf i = 9 Then
                    Set rng = ThisWorkbook.Sheets(9).Range("a5:dv29")
                ElseIf i = 10 Then
                    Set rng = ThisWorkbook.Sheets(10).Range("b5:dq22")
                Else
                    Set rng = ThisWorkbook.Sheets(11).Range("b5:ax40")
                End If
        
        
        ' add a new slide to the presentation
        
        Set PPnewSlide = PPApp.ActivePresentation.Slides.Add(i + 1, ppLayoutBlank)
        'PPApp.ActiveWindow.View.GotoSlide PPApp.ActivePresentation.Slides.Count
        'Set PPSlide = PPApp.ActivePresentation.Slides(PPApp.ActivePresentation.Slides.Count)
        Set PPactiveSlide = PPApp.ActivePresentation.Slides(i + 1)
        

        
        
        
                
                'add the title slide
                
                With PPactiveSlide
        
                If Not .Shapes.HasTitle Then
                    Set shpCurrShape1 = .Shapes.AddTextbox(1, 104, 24, 800, 29)
                    Set shpCurrShape2 = .Shapes.AddTextbox(1, 104, 72, 800, 29)
                Else
                    Set shpCurrShape1 = .Shapes.Title
                    Set shpCurrShape2 = .Shapes.Title
                End If
                
                End With
                
  ' copy this part over
  
                If i = 1 Then
                
                    With shpCurrShape1
                        With .TextFrame.TextRange
                            '~~> Set text here
                            .Text = "SKYWARDS OUTLOOK – REGION | POS | FLOWS u TICKETED OD PAX"
                            '~~> Alignment
                            .ParagraphFormat.Alignment = 1
                           '~~> Working with font
                           With .Font
                              .Bold = msoFalse
                              .Underline = msoTrue
                              .Name = "Heroic Condensed Medium"
                              .Size = 40
                              .Color = RGB(192, 0, 0)
                           End With
                           With .Find(findwhat:="REGION | POS | FLOWS").Font
                              .Bold = msoFalse
                              .Name = "Heroic Condensed Medium"
                              .Size = 28
                              .Color = RGB(192, 0, 0)
                           End With
                           
                           With .Find(findwhat:=" u", after:=.Start + 30).Font
                              .Bold = msoFalse
                              .Underline = msoFalse
                              .Name = "Wingdings 3"
                              .Size = 18
                              .Color = RGB(166, 137, 93)
                           End With
                           With .Find(findwhat:=" TICKETED OD PAX").Font
                              .Bold = msoFalse
                              .Underline = msoFalse
                              .Name = "Heroic Condensed Medium"
                              .Size = 26
                              .Color = RGB(166, 137, 93)
                           End With
                        End With
                    End With
                    
                        With shpCurrShape2
                        With .TextFrame.TextRange
                            '~~> Set text here
                            .Text = "Travel " & Format(Date, "MMMYY") & "-" & Format(Date + 150, "MMMYY") & _
                            " (current + 5 months) | EBI Snap " & Format(Date - 2, "DD-MMM") & _
                            " (sorted descending Pax TY)"
                            '~~> Alignment
                            .ParagraphFormat.Alignment = 1
                           '~~> Working with font
                           With .Font
                              .Bold = msoTrue
                              .Name = "Emirates Light"
                              .Size = 18
                              .Color = RGB(0, 0, 0)
                           End With
                           With .Find(findwhat:="(current + 5 months)").Font
                              .Bold = msoFalse
                              .Name = "Emirates Light"
                              .Size = 14
                              .Color = RGB(0, 0, 0)
                           End With
                           With .Find(findwhat:="(sorted descending Pax TY)").Font
                              .Bold = msoFalse
                              .Name = "Emirates Light"
                              .Size = 14
                              .Color = RGB(0, 0, 0)
                           End With
                           
                           
                           
                          End With
                    End With
                    
                
                ElseIf i = 2 Then
                
                        With shpCurrShape1
                        With .TextFrame.TextRange
                            '~~> Set text here
                            .Text = "SKYWARDS OUTLOOK – REGION | POS | FLOWS u TICKETED OD PAX (FJ)"
                            '~~> Alignment
                            .ParagraphFormat.Alignment = 1
                           '~~> Working with font
                             With .Font
                              .Bold = msoFalse
                              .Underline = msoTrue
                              .Name = "Heroic Condensed Medium"
                              .Size = 40
                              .Color = RGB(192, 0, 0)
                           End With
                           With .Find(findwhat:="REGION | POS | FLOWS").Font
                              .Bold = msoFalse
                              .Name = "Heroic Condensed Medium"
                              .Size = 28
                              .Color = RGB(192, 0, 0)
                           End With
                           
                           With .Find(findwhat:=" u", after:=.Start + 30).Font
                              .Bold = msoFalse
                              .Underline = msoFalse
                              .Name = "Wingdings 3"
                              .Size = 18
                              .Color = RGB(166, 137, 93)
                           End With
                           With .Find(findwhat:=" TICKETED OD PAX (FJ)").Font
                              .Bold = msoFalse
                              .Underline = msoFalse
                              .Name = "Heroic Condensed Medium"
                              .Size = 26
                              .Color = RGB(166, 137, 93)
                           End With
                        End With
                    End With
                    
                        With shpCurrShape2
                        With .TextFrame.TextRange
                            '~~> Set text here
                            .Text = "Travel " & Format(Date, "MMMYY") & "-" & Format(Date + 150, "MMMYY") & _
                            " (current + 5 months) | EBI Snap " & Format(Date - 2, "DD-MMM") & _
                            " (sorted descending Pax TY)"
                            '~~> Alignment
                            .ParagraphFormat.Alignment = 1
                           '~~> Working with font
                           With .Font
                              .Bold = msoTrue
                              .Name = "Emirates Light"
                              .Size = 18
                              .Color = RGB(0, 0, 0)
                           End With
                           With .Find(findwhat:="(current + 5 months)").Font
                              .Bold = msoFalse
                              .Name = "Emirates Light"
                              .Size = 14
                              .Color = RGB(0, 0, 0)
                           End With
                           With .Find(findwhat:="(sorted descending Pax TY)").Font
                              .Bold = msoFalse
                              .Name = "Emirates Light"
                              .Size = 14
                              .Color = RGB(0, 0, 0)
                           End With
                        End With
                    End With
                    
                    
                ElseIf i = 3 Then
                
                        With shpCurrShape1
                        With .TextFrame.TextRange
                            '~~> Set text here
                            .Text = "SKYWARDS OUTLOOK – CHANNEL | EOL u TICKETED OD PAX"
                            '~~> Alignment
                            .ParagraphFormat.Alignment = 1
                           '~~> Working with font
                           With .Font
                              .Bold = msoFalse
                              .Underline = msoTrue
                              .Name = "Heroic Condensed Medium"
                              .Size = 40
                              .Color = RGB(192, 0, 0)
                           End With
                           With .Find(findwhat:="CHANNEL | EOL").Font
                              .Bold = msoFalse
                              .Name = "Heroic Condensed Medium"
                              .Size = 28
                              .Color = RGB(192, 0, 0)
                           End With
                           
                           With .Find(findwhat:=" u", after:=.Start + 30).Font
                              .Bold = msoFalse
                              .Underline = msoFalse
                              .Name = "Wingdings 3"
                              .Size = 18
                              .Color = RGB(166, 137, 93)
                           End With
                           With .Find(findwhat:=" TICKETED OD PAX").Font
                              .Bold = msoFalse
                              .Underline = msoFalse
                              .Name = "Heroic Condensed Medium"
                              .Size = 26
                              .Color = RGB(166, 137, 93)
                           End With
                           
                           
                        End With
                    End With
                    
                        With shpCurrShape2
                        With .TextFrame.TextRange
                            '~~> Set text here
                            .Text = "Travel " & Format(Date, "MMMYY") & "-" & Format(Date + 150, "MMMYY") & _
                            " (current + 5 months) | EBI Snap " & Format(Date - 2, "DD-MMM") & _
                            " (sorted descending Pax TY)"
                            '~~> Alignment
                            .ParagraphFormat.Alignment = 1
                           '~~> Working with font
                           With .Font
                              .Bold = msoTrue
                              .Name = "Emirates Light"
                              .Size = 18
                              .Color = RGB(0, 0, 0)
                           End With
                           With .Find(findwhat:="(current + 5 months)").Font
                              .Bold = msoFalse
                              .Name = "Emirates Light"
                              .Size = 14
                              .Color = RGB(0, 0, 0)
                           End With
                           With .Find(findwhat:="(sorted descending Pax TY)").Font
                              .Bold = msoFalse
                              .Name = "Emirates Light"
                              .Size = 14
                              .Color = RGB(0, 0, 0)
                           End With
                        End With
                    End With
                    

                ElseIf i = 4 Then
                
                        With shpCurrShape1
                        With .TextFrame.TextRange
                            '~~> Set text here
                            .Text = "SKYWARDS OUTLOOK – TRAVEL MONTH | CABIN | FARE BRAND u TICKETED OD PAX"
                            '~~> Alignment
                            .ParagraphFormat.Alignment = 1
                           '~~> Working with font
                           With .Font
                              .Bold = msoFalse
                              .Underline = msoTrue
                              .Name = "Heroic Condensed Medium"
                              .Size = 40
                              .Color = RGB(192, 0, 0)
                           End With
                           With .Find(findwhat:="TRAVEL MONTH | CABIN | FARE BRAND").Font
                              .Bold = msoFalse
                              .Name = "Heroic Condensed Medium"
                              .Size = 28
                              .Color = RGB(192, 0, 0)
                           End With
                           
                           With .Find(findwhat:=" u", after:=.Start + 30).Font
                              .Bold = msoFalse
                              .Underline = msoFalse
                              .Name = "Wingdings 3"
                              .Size = 18
                              .Color = RGB(166, 137, 93)
                           End With
                           With .Find(findwhat:=" TICKETED OD PAX").Font
                              .Bold = msoFalse
                              .Underline = msoFalse
                              .Name = "Heroic Condensed Medium"
                              .Size = 26
                              .Color = RGB(166, 137, 93)
                           End With
                           
                        End With
                    End With
                    
                        With shpCurrShape2
                        With .TextFrame.TextRange
                            '~~> Set text here
                            .Text = "Travel " & Format(Date, "MMMYY") & "-" & Format(Date + 150, "MMMYY") & _
                            " (current + 5 months) | EBI Snap " & Format(Date - 2, "DD-MMM") & _
                            " (sorted descending Pax TY)"
                            '~~> Alignment
                            .ParagraphFormat.Alignment = 1
                           '~~> Working with font
                           With .Font
                              .Bold = msoTrue
                              .Name = "Emirates Light"
                              .Size = 18
                              .Color = RGB(0, 0, 0)
                           End With
                           With .Find(findwhat:="(current + 5 months)").Font
                              .Bold = msoFalse
                              .Name = "Emirates Light"
                              .Size = 14
                              .Color = RGB(0, 0, 0)
                           End With
                           With .Find(findwhat:="(sorted descending Pax TY)").Font
                              .Bold = msoFalse
                              .Name = "Emirates Light"
                              .Size = 14
                              .Color = RGB(0, 0, 0)
                           End With
                        End With
                    End With
                    
                
                ElseIf i = 5 Then
                
                        With shpCurrShape1
                        With .TextFrame.TextRange
                            '~~> Set text here
                            .Text = "SKYWARDS OUTLOOK – FREEDOM | JOURNEY | PAX in PNR u TICKETED OD PAX"
                            '~~> Alignment
                            .ParagraphFormat.Alignment = 1
                           '~~> Working with font
                           With .Font
                              .Bold = msoFalse
                              .Underline = msoTrue
                              .Name = "Heroic Condensed Medium"
                              .Size = 40
                              .Color = RGB(192, 0, 0)
                           End With
                           With .Find(findwhat:="FREEDOM | JOURNEY | PAX in PNR").Font
                              .Bold = msoFalse
                              .Name = "Heroic Condensed Medium"
                              .Size = 28
                              .Color = RGB(192, 0, 0)
                           End With
                           
                           With .Find(findwhat:=" u", after:=.Start + 30).Font
                              .Bold = msoFalse
                              .Underline = msoFalse
                              .Name = "Wingdings 3"
                              .Size = 18
                              .Color = RGB(166, 137, 93)
                           End With
                           With .Find(findwhat:=" TICKETED OD PAX").Font
                              .Bold = msoFalse
                              .Underline = msoFalse
                              .Name = "Heroic Condensed Medium"
                              .Size = 26
                              .Color = RGB(166, 137, 93)
                           End With
                           
                        End With
                    End With
                    
                        With shpCurrShape2
                        With .TextFrame.TextRange
                            '~~> Set text here
                            .Text = "Travel " & Format(Date, "MMMYY") & "-" & Format(Date + 150, "MMMYY") & _
                            " (current + 5 months) | EBI Snap " & Format(Date - 2, "DD-MMM") & _
                            " (sorted descending Pax TY)"
                            '~~> Alignment
                            .ParagraphFormat.Alignment = 1
                           '~~> Working with font
                           With .Font
                              .Bold = msoTrue
                              .Name = "Emirates Light"
                              .Size = 18
                              .Color = RGB(0, 0, 0)
                           End With
                           With .Find(findwhat:="(current + 5 months)").Font
                              .Bold = msoFalse
                              .Name = "Emirates Light"
                              .Size = 14
                              .Color = RGB(0, 0, 0)
                           End With
                           With .Find(findwhat:="(sorted descending Pax TY)").Font
                              .Bold = msoFalse
                              .Name = "Emirates Light"
                              .Size = 14
                              .Color = RGB(0, 0, 0)
                           End With
                        End With
                    End With
                    
                    
                    
                ElseIf i = 6 Then
                
                        With shpCurrShape1
                        With .TextFrame.TextRange
                            '~~> Set text here
                            .Text = "SKYWARDS OUTLOOK – REGION | POS | FLOWS u TICKETED REVENUE (AED)"
                            '~~> Alignment
                            .ParagraphFormat.Alignment = 1
                           '~~> Working with font
                           With .Font
                              .Bold = msoFalse
                              .Underline = msoTrue
                              .Name = "Heroic Condensed Medium"
                              .Size = 40
                              .Color = RGB(192, 0, 0)
                           End With
                           With .Find(findwhat:="REGION | POS | FLOWS").Font
                              .Bold = msoFalse
                              .Name = "Heroic Condensed Medium"
                              .Size = 28
                              .Color = RGB(192, 0, 0)
                           End With
                           
                           With .Find(findwhat:=" u", after:=.Start + 30).Font
                              .Bold = msoFalse
                              .Underline = msoFalse
                              .Name = "Wingdings 3"
                              .Size = 18
                              .Color = RGB(166, 137, 93)
                           End With
                           With .Find(findwhat:=" TICKETED REVENUE (AED)").Font
                              .Bold = msoFalse
                              .Underline = msoFalse
                              .Name = "Heroic Condensed Medium"
                              .Size = 26
                              .Color = RGB(166, 137, 93)
                           End With
                           
                           
                        End With
                    End With
                    
                        With shpCurrShape2
                        With .TextFrame.TextRange
                            '~~> Set text here
                            .Text = "Travel " & Format(Date, "MMMYY") & "-" & Format(Date + 150, "MMMYY") & _
                            " (current + 5 months) | EBI Snap " & Format(Date - 2, "DD-MMM") & _
                            " (sorted descending Rev TY)"
                            '~~> Alignment
                            .ParagraphFormat.Alignment = 1
                           '~~> Working with font
                           With .Font
                              .Bold = msoTrue
                              .Name = "Emirates Light"
                              .Size = 18
                              .Color = RGB(0, 0, 0)
                           End With
                           With .Find(findwhat:="(current + 5 months)").Font
                              .Bold = msoFalse
                              .Name = "Emirates Light"
                              .Size = 14
                              .Color = RGB(0, 0, 0)
                           End With
                           With .Find(findwhat:="(sorted descending Rev TY)").Font
                              .Bold = msoFalse
                              .Name = "Emirates Light"
                              .Size = 14
                              .Color = RGB(0, 0, 0)
                           End With
                        End With
                    End With
                    
                    
                ElseIf i = 7 Then
                
                        With shpCurrShape1
                        With .TextFrame.TextRange
                            '~~> Set text here
                            .Text = "SKYWARDS OUTLOOK – REGION | POS | FLOWS u TICKETED REVENUE (AED) (FJ)"
                            '~~> Alignment
                            .ParagraphFormat.Alignment = 1
                           '~~> Working with font
                           With .Font
                              .Bold = msoFalse
                              .Underline = msoTrue
                              .Name = "Heroic Condensed Medium"
                              .Size = 40
                              .Color = RGB(192, 0, 0)
                           End With
                           With .Find(findwhat:="REGION | POS | FLOWS").Font
                              .Bold = msoFalse
                              .Name = "Heroic Condensed Medium"
                              .Size = 28
                              .Color = RGB(192, 0, 0)
                           End With
                           
                           With .Find(findwhat:=" u", after:=.Start + 30).Font
                              .Bold = msoFalse
                              .Underline = msoFalse
                              .Name = "Wingdings 3"
                              .Size = 18
                              .Color = RGB(166, 137, 93)
                           End With
                           With .Find(findwhat:=" TICKETED REVENUE (AED) (FJ)").Font
                              .Bold = msoFalse
                              .Underline = msoFalse
                              .Name = "Heroic Condensed Medium"
                              .Size = 26
                              .Color = RGB(166, 137, 93)
                           End With
                           
                           
                        End With
                    End With
                    
                        With shpCurrShape2
                        With .TextFrame.TextRange
                            '~~> Set text here
                            .Text = "Travel " & Format(Date, "MMMYY") & "-" & Format(Date + 150, "MMMYY") & _
                            " (current + 5 months) | EBI Snap " & Format(Date - 2, "DD-MMM") & _
                            " (sorted descending Rev TY)"
                            '~~> Alignment
                            .ParagraphFormat.Alignment = 1
                           '~~> Working with font
                           With .Font
                              .Bold = msoTrue
                              .Name = "Emirates Light"
                              .Size = 18
                              .Color = RGB(0, 0, 0)
                           End With
                           With .Find(findwhat:="(current + 5 months)").Font
                              .Bold = msoFalse
                              .Name = "Emirates Light"
                              .Size = 14
                              .Color = RGB(0, 0, 0)
                           End With
                           With .Find(findwhat:="(sorted descending Rev TY)").Font
                              .Bold = msoFalse
                              .Name = "Emirates Light"
                              .Size = 14
                              .Color = RGB(0, 0, 0)
                           End With
                        End With
                    End With
                    
                    
                ElseIf i = 8 Then
                
                        With shpCurrShape1
                        With .TextFrame.TextRange
                            '~~> Set text here
                            .Text = "SKYWARDS OUTLOOK – TRAVEL MONTH | CABIN | FARE BRAND u TICKETED REVENUE (AED)"
                            '~~> Alignment
                            .ParagraphFormat.Alignment = 1
                           '~~> Working with font
                           With .Font
                              .Bold = msoFalse
                              .Underline = msoTrue
                              .Name = "Heroic Condensed Medium"
                              .Size = 40
                              .Color = RGB(192, 0, 0)
                           End With
                           With .Find(findwhat:="TRAVEL MONTH | CABIN | FARE BRAND").Font
                              .Bold = msoFalse
                              .Name = "Heroic Condensed Medium"
                              .Size = 28
                              .Color = RGB(192, 0, 0)
                           End With
                           
                           With .Find(findwhat:=" u", after:=.Start + 30).Font
                              .Bold = msoFalse
                              .Underline = msoFalse
                              .Name = "Wingdings 3"
                              .Size = 18
                              .Color = RGB(166, 137, 93)
                           End With
                           With .Find(findwhat:=" TICKETED REVENUE (AED)").Font
                              .Bold = msoFalse
                              .Underline = msoFalse
                              .Name = "Heroic Condensed Medium"
                              .Size = 26
                              .Color = RGB(166, 137, 93)
                           End With
                           
                           
                        End With
                    End With
                    
                        With shpCurrShape2
                        With .TextFrame.TextRange
                            '~~> Set text here
                            .Text = "Travel " & Format(Date, "MMMYY") & "-" & Format(Date + 150, "MMMYY") & _
                            " (current + 5 months) | EBI Snap " & Format(Date - 2, "DD-MMM")
                            '~~> Alignment
                            .ParagraphFormat.Alignment = 1
                           '~~> Working with font
                           With .Font
                              .Bold = msoTrue
                              .Name = "Emirates Light"
                              .Size = 18
                              .Color = RGB(0, 0, 0)
                           End With
                           With .Find(findwhat:="(current + 5 months)").Font
                              .Bold = msoFalse
                              .Name = "Emirates Light"
                              .Size = 14
                              .Color = RGB(0, 0, 0)
                           End With

                        End With
                    End With
                    

                ElseIf i = 9 Then
                
                        With shpCurrShape1
                        With .TextFrame.TextRange
                            '~~> Set text here
                            .Text = "SKYWARDS OUTLOOK – FREEDOM | JOURNEY | PAX in PNR u TICKETED REVENUE (AED)"
                            '~~> Alignment
                            .ParagraphFormat.Alignment = 1
                           '~~> Working with font
                           With .Font
                              .Bold = msoFalse
                              .Underline = msoTrue
                              .Name = "Heroic Condensed Medium"
                              .Size = 40
                              .Color = RGB(192, 0, 0)
                           End With
                           With .Find(findwhat:="FREEDOM | JOURNEY | PAX in PNR").Font
                              .Bold = msoFalse
                              .Name = "Heroic Condensed Medium"
                              .Size = 28
                              .Color = RGB(192, 0, 0)
                           End With
                           
                           With .Find(findwhat:=" u", after:=.Start + 30).Font
                              .Bold = msoFalse
                              .Underline = msoFalse
                              .Name = "Wingdings 3"
                              .Size = 18
                              .Color = RGB(166, 137, 93)
                           End With
                           With .Find(findwhat:=" TICKETED REVENUE (AED)").Font
                              .Bold = msoFalse
                              .Underline = msoFalse
                              .Name = "Heroic Condensed Medium"
                              .Size = 26
                              .Color = RGB(166, 137, 93)
                           End With
                           
                        End With
                    End With
                    
                        With shpCurrShape2
                        With .TextFrame.TextRange
                            '~~> Set text here
                            .Text = "Travel " & Format(Date, "MMMYY") & "-" & Format(Date + 150, "MMMYY") & _
                            " (current + 5 months) | EBI Snap " & Format(Date - 2, "DD-MMM")
                            '~~> Alignment
                            .ParagraphFormat.Alignment = 1
                           '~~> Working with font
                           With .Font
                              .Bold = msoTrue
                              .Name = "Emirates Light"
                              .Size = 18
                              .Color = RGB(0, 0, 0)
                           End With
                           With .Find(findwhat:="(current + 5 months)").Font
                              .Bold = msoFalse
                              .Name = "Emirates Light"
                              .Size = 14
                              .Color = RGB(0, 0, 0)
                           End With
                        End With
                    End With
                    
                    
                ElseIf i = 10 Then
                
                        With shpCurrShape1
                        With .TextFrame.TextRange
                            '~~> Set text here
                            .Text = "SKYWARDS OUTLOOK – CHANNEL | EOL u TICKETED REVENUE (AED)"
                            '~~> Alignment
                            .ParagraphFormat.Alignment = 1
                           '~~> Working with font
                           With .Font
                              .Bold = msoFalse
                              .Underline = msoTrue
                              .Name = "Heroic Condensed Medium"
                              .Size = 40
                              .Color = RGB(192, 0, 0)
                           End With
                           With .Find(findwhat:="CHANNEL | EOL").Font
                              .Bold = msoFalse
                              .Name = "Heroic Condensed Medium"
                              .Size = 28
                              .Color = RGB(192, 0, 0)
                           End With
                           
                           With .Find(findwhat:=" u", after:=.Start + 30).Font
                              .Bold = msoFalse
                              .Underline = msoFalse
                              .Name = "Wingdings 3"
                              .Size = 18
                              .Color = RGB(166, 137, 93)
                           End With
                           With .Find(findwhat:=" TICKETED REVENUE (AED)").Font
                              .Bold = msoFalse
                              .Underline = msoFalse
                              .Name = "Heroic Condensed Medium"
                              .Size = 26
                              .Color = RGB(166, 137, 93)
                           End With
            
                           
                        End With
                    End With
                    
                        With shpCurrShape2
                        With .TextFrame.TextRange
                            '~~> Set text here
                            .Text = "Travel " & Format(Date, "MMMYY") & "-" & Format(Date + 150, "MMMYY") & _
                            " (current + 5 months) | EBI Snap " & Format(Date - 2, "DD-MMM") & _
                            " (sorted descending Pax TY)"
                            '~~> Alignment
                            .ParagraphFormat.Alignment = 1
                           '~~> Working with font
                           With .Font
                              .Bold = msoTrue
                              .Name = "Emirates Light"
                              .Size = 18
                              .Color = RGB(0, 0, 0)
                           End With
                           With .Find(findwhat:="(current + 5 months)").Font
                              .Bold = msoFalse
                              .Name = "Emirates Light"
                              .Size = 14
                              .Color = RGB(0, 0, 0)
                           End With
                           With .Find(findwhat:="(sorted descending Pax TY)").Font
                              .Bold = msoFalse
                              .Name = "Emirates Light"
                              .Size = 14
                              .Color = RGB(0, 0, 0)
                           End With
                        End With
                    End With
                    
                  Else
                    With shpCurrShape1
                        With .TextFrame.TextRange
                            '~~> Set text here
                            .Text = "SKYWARDS OUTLOOK – REGION | POS | FLOWS u TICKETED OD PAX"
                            '~~> Alignment
                            .ParagraphFormat.Alignment = 1
                           '~~> Working with font
                           With .Font
                              .Bold = msoFalse
                              .Underline = msoTrue
                              .Name = "Heroic Condensed Medium"
                              .Size = 40
                              .Color = RGB(192, 0, 0)
                           End With
                           With .Find(findwhat:="REGION | POS | FLOWS").Font
                              .Bold = msoFalse
                              .Name = "Heroic Condensed Medium"
                              .Size = 28
                              .Color = RGB(192, 0, 0)
                           End With
                           
                           With .Find(findwhat:=" u", after:=.Start + 30).Font
                              .Bold = msoFalse
                              .Underline = msoFalse
                              .Name = "Wingdings 3"
                              .Size = 18
                              .Color = RGB(166, 137, 93)
                           End With
                           With .Find(findwhat:=" TICKETED OD PAX").Font
                              .Bold = msoFalse
                              .Underline = msoFalse
                              .Name = "Heroic Condensed Medium"
                              .Size = 26
                              .Color = RGB(166, 137, 93)
                           End With
                        End With
                    End With
                    
                        With shpCurrShape2
                        With .TextFrame.TextRange
                            '~~> Set text here
                            .Text = "Travel " & Format(Date, "MMMYY") & "-" & Format(Date + 150, "MMMYY") & _
                            " (current + 5 months) | EBI Snap " & Format(Date - 2, "DD-MMM") & _
                            " (sorted descending Pax TY)"
                            '~~> Alignment
                            .ParagraphFormat.Alignment = 1
                           '~~> Working with font
                           With .Font
                              .Bold = msoTrue
                              .Name = "Emirates Light"
                              .Size = 18
                              .Color = RGB(0, 0, 0)
                           End With
                           With .Find(findwhat:="(current + 5 months)").Font
                              .Bold = msoFalse
                              .Name = "Emirates Light"
                              .Size = 14
                              .Color = RGB(0, 0, 0)
                           End With
                           With .Find(findwhat:="(sorted descending Pax TY)").Font
                              .Bold = msoFalse
                              .Name = "Emirates Light"
                              .Size = 14
                              .Color = RGB(0, 0, 0)
                           End With
                        End With
                    End With
                    
                  End If
                       
        
        rng.Copy
        
        PPactiveSlide.Shapes.PasteSpecial ppPasteEnhancedMetafile
        
        ' change the size and shape of the picture object
        
        Set myShape = PPactiveSlide.Shapes(PPactiveSlide.Shapes.Count)
        
        'myShape.Left = 20
        'myShape.Top = 120
        
        
        With myShape

' copy this part over
        
        If i = 1 Then

                'size:
                ''1 inch = 72 points
                .Height = 72 * 5
                .Width = 72 * 12

                .ScaleHeight 0.83, msoTrue
                .ScaleWidth 0.86, msoTrue

                .LockAspectRatio = msoFalse


                'position:
                .Rotation = 0

                .Left = 30
                .Top = 110

                'Relative to original picture size = true
        
        ElseIf i = 2 Then

                'size:
                ''1 inch = 72 points
                .Height = 72 * 5
                .Width = 72 * 12

                .ScaleHeight 0.83, msoTrue
                .ScaleWidth 0.86, msoTrue

                .LockAspectRatio = msoFalse


                'position:
                .Rotation = 0

                .Left = 30
                .Top = 110

                'Relative to original picture size = true

        ElseIf i = 3 Then

                'size:
                ''1 inch = 72 points
                .Height = 72 * 5
                .Width = 72 * 12

                .ScaleHeight 0.83, msoTrue
                .ScaleWidth 0.86, msoTrue

                .LockAspectRatio = msoFalse


                'position:
                .Rotation = 0

                .Left = 30
                .Top = 110

                'Relative to original picture size = true
                
        ElseIf i = 4 Then

                'size:
                ''1 inch = 72 points
                .Height = 72 * 3.39
                .Width = 72 * 12

                .ScaleHeight 0.83, msoTrue
                .ScaleWidth 0.86, msoTrue

                .LockAspectRatio = msoFalse


                'position:
                .Rotation = 0

                .Left = 30
                .Top = 110

                'Relative to original picture size = true
                
        ElseIf i = 5 Then

                'size:
                ''1 inch = 72 points
                .Height = 72 * 3.39
                .Width = 72 * 12

                .ScaleHeight 0.83, msoTrue
                .ScaleWidth 0.86, msoTrue

                .LockAspectRatio = msoFalse


                'position:
                .Rotation = 0

                .Left = 30
                .Top = 110

                'Relative to original picture size = true
                
        ElseIf i = 6 Then

                'size:
                ''1 inch = 72 points
                .Height = 72 * 3.39
                .Width = 72 * 12

                .ScaleHeight 0.83, msoTrue
                .ScaleWidth 0.86, msoTrue

                .LockAspectRatio = msoFalse


                'position:
                .Rotation = 0

                .Left = 30
                .Top = 110

                'Relative to original picture size = true
                
        ElseIf i = 7 Then

                'size:
                ''1 inch = 72 points
                .Height = 72 * 3.39
                .Width = 72 * 12

                .ScaleHeight 0.83, msoTrue
                .ScaleWidth 0.86, msoTrue

                .LockAspectRatio = msoFalse


                'position:
                .Rotation = 0

                .Left = 30
                .Top = 110

                'Relative to original picture size = true
                
        ElseIf i = 8 Then

                'size:
                ''1 inch = 72 points
                .Height = 72 * 3.39
                .Width = 72 * 12

                .ScaleHeight 0.83, msoTrue
                .ScaleWidth 0.86, msoTrue

                .LockAspectRatio = msoFalse


                'position:
                .Rotation = 0

                .Left = 30
                .Top = 110

                'Relative to original picture size = true
                
      ElseIf i = 9 Then

                'size:
                ''1 inch = 72 points
                .Height = 72 * 5
                .Width = 72 * 12

                .ScaleHeight 0.83, msoTrue
                .ScaleWidth 0.86, msoTrue

                .LockAspectRatio = msoFalse


                'position:
                .Rotation = 0

                .Left = 30
                .Top = 110

                'Relative to original picture size = true
                
      ElseIf i = 10 Then

                'size:
                ''1 inch = 72 points
                .Height = 72 * 5
                .Width = 72 * 14

                .ScaleHeight 0.9, msoTrue
                .ScaleWidth 0.9, msoTrue

                .LockAspectRatio = msoFalse


                'position:
                .Rotation = 0

                .Left = 30
                .Top = 110

                'Relative to original picture size = true
                
        Else

                'size:
                ''1 inch = 72 points
                .Height = 72 * 3.39
                .Width = 72 * 12

                .ScaleHeight 0.83, msoTrue
                .ScaleWidth 0.86, msoTrue

                .LockAspectRatio = msoFalse


                'position:
                .Rotation = 0

                .Left = 30
                .Top = 110

                'Relative to original picture size = true
        
                
                
        End If

        End With
        
        Set MyPic = ThisWorkbook.Sheets(19).Shapes("Picture 1")
        
        MyPic.Copy
        
        PPactiveSlide.Shapes.Paste
        
        Set myShape = PPactiveSlide.Shapes(PPactiveSlide.Shapes.Count)
        
        With myShape

                'size:
                ''1 inch = 72 points
                '.Height = 72 * 3.39
                '.Width = 72 * 6.67

                '.ScaleHeight 0.9, msoTrue
                '.ScaleWidth 0.9, msoTrue

                .LockAspectRatio = msoTrue


                'position:
                .Rotation = 0

                .Left = 40
                .Top = 34

                'Relative to original picture size = true

        End With
        
   Next
   
   
   
       Set PPnewSlide = PPApp.ActivePresentation.Slides.Add(11, ppLayoutBlank)
       Set PPactiveSlide = PPApp.ActivePresentation.Slides(11)
     
        
        
        With PPactiveSlide
        
                If Not .Shapes.HasTitle Then
                    Set shpCurrShape1 = .Shapes.AddTextbox(1, 312, 250, 600, 29)
                Else
                    Set shpCurrShape1 = .Shapes.Title
                End If
                
                End With
                
               
                    With shpCurrShape1
                        With .TextFrame.TextRange
                            '~~> Set text here
                            .Text = "Appendix"
                            .ParagraphFormat.Alignment = 3
                           '~~> Working with font
                           With .Font
                              .Italic = msoTrue
                              .Bold = msoTrue
                              .Underline = msoTrue
                              .Name = "Heroic Condensed Medium"
                              .Size = 60
                              .Color = RGB(192, 0, 0)
                           End With
                        
                        End With
                    End With
   
   
   
   
   
        
Application.CutCopyMode = False
        

With PPPres
    .SaveAs fileNameString, ppSaveAsPDF
    .SaveAs fileNameString2, ppSaveAsOpenXMLPresentation
    .Close
End With


    'Quit PowerPoint
    
    PPApp.Quit

    ' Clean up

    Set PPApp = Nothing
    Set PPPres = Nothing
    Set PPactiveSlide = Nothing
    Set PPnewSlide = Nothing
    'Dim PPSlide2 As PowerPoint.Slide
    Set myShape = Nothing
    Set shpCurrShape1 = Nothing
    Set shpCurrShape2 = Nothing
    Set MyPic = Nothing

'Application.ScreenUpdating = True


'ThisWorkbook.Activate


'MsgBox ("Presentation & PDF created")


End Sub