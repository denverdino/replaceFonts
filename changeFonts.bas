Const myFontName = "Alibaba Sans"
Const myFontNameFarEast = "Alibaba PuHuiTi 2.0 55 Regular"

Sub replaceMasterFonts(master As Master)
    With master.Theme.ThemeFontScheme
        Debug.Print .MajorFont.Item(msoThemeLatin).Name
        Debug.Print .MajorFont.Item(msoThemeEastAsian).Name

        Debug.Print .MinorFont.Item(msoThemeLatin).Name
        Debug.Print .MinorFont.Item(msoThemeEastAsian).Name
        
        .MajorFont.Item(msoThemeLatin).Name = myFontName
        .MajorFont.Item(msoThemeEastAsian).Name = myFontNameFarEast
        .MinorFont.Item(msoThemeLatin).Name = myFontName
        .MinorFont.Item(msoThemeEastAsian).Name = myFontNameFarEast
    End With
    
    With master.TextStyles(ppBodyStyle)
        For i = 1 To .Levels.Count
            With .Levels(i).Font
                .Name = myFontName
                .NameFarEast = myFontNameFarEast
            End With
        Next i
    End With
    
    With master.TextStyles(ppTitleStyle)
        For i = 1 To .Levels.Count
            With .Levels(i).Font
                .Name = myFontName
                .NameFarEast = myFontNameFarEast
            End With
        Next i
    End With
    replaceShapeFonts master.Shapes
End Sub

Sub replaceShapeFonts(shapes As Shapes)
    For Each shp In shapes
        If shp.HasTextFrame Then
            With shp.TextFrame.TextRange.Font
                .Name = myFontName
                .NameFarEast = myFontNameFarEast
            End With
        End If
    Next
End Sub


Sub alibabafonts()
    
    Debug.Print "Replace fonts in slide master to my favorites ..."
    
    replaceMasterFonts ActivePresentation.SlideMaster

    Debug.Print "Replace fonts in title master to my favorites ..."
    If ActivePresentation.HasTitleMaster Then
        replaceMasterFonts ActivePresentation.TitleMaster
    End If
    
    For Each oDes In ActivePresentation.Designs
        For Each oCL In oDes.SlideMaster.CustomLayouts
            replaceShapeFonts oCL.Shapes
        Next
    Next

    Debug.Print "Replace fonts in notes master to my favorites ..."
    replaceMasterFonts ActivePresentation.NotesMaster

    Debug.Print "Replace fonts in handout master to my favorites ..."
    replaceMasterFonts ActivePresentation.HandoutMaster
    
    Debug.Print "Replace fonts in slides to my favorites ..."

    For Each sld In ActivePresentation.Slides
        replaceShapeFonts sld.Shapes
    Next

    Debug.Print "Done!"

End Sub