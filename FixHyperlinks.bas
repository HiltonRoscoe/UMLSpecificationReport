Attribute VB_Name = "HrWordSupport"
Sub FixHyperlinks()


Dim doc As Document
Dim link, i
'Loop through all open documents.
For Each doc In Application.Documents
    doc.Bookmarks.ShowHidden = True
    'Loop through all hyperlinks.
    For i = 1 To doc.Hyperlinks.Count
        'Check if a bookmark exists for the hyperlink
        Dim gotOne As Boolean
        gotOne = False
        For j = 1 To doc.Bookmarks.Count
            If doc.Bookmarks(j).Name = doc.Hyperlinks(i).SubAddress Then
                gotOne = True
                Exit For
            End If
            On Error Resume Next
        Next
        If Not gotOne Then
            ' try to find bookmark text that matches the hyperlink's text
            For j = 1 To doc.Bookmarks.Count
            ' Add "Class " to catch more
            If doc.Bookmarks(j).Range.Text = "Class " + doc.Hyperlinks(i).Range.Text Then
                Debug.Print "assigning " + doc.Bookmarks(j).Name + " for " + doc.Hyperlinks(i).SubAddress
                'doc.Hyperlinks(i).SubAddress = doc.Bookmarks(j).Name
                  With doc.Hyperlinks(i)
                  Set Rng = .Range
                  StrAddr = .Address
                  StrTxt = .TextToDisplay
                  .Delete
                End With
                doc.Hyperlinks.Add Anchor:=Rng, Address:=StrAddr, SubAddress:=doc.Bookmarks(j).Name
                Exit For
            End If
        Next
        End If
    Next
Next
End Sub
