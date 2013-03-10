Attribute VB_Name = "NewMacros"
Sub •¶ÍZ³()
' 2013/03/10
    ‘SŠp‚©‚ç”¼Šp‚Ö•ÏŠ· ("[‚O-‚X‚`-‚y‚-‚š]{1,}")
    ‹Ö~•¶šŒŸ’m ("[" & Chr(&H8740) & "-" & Chr(&H889F) & "]{1,}")
End Sub

Private Sub ‘SŠp‚©‚ç”¼Šp‚Ö•ÏŠ·(ByVal strPattern As String)
' 2013/03/10
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = strPattern
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .MatchFuzzy = False
        .MatchWildcards = True
    End With
    Do While Selection.Find.Execute
        Selection.Text = StrConv(Selection.Text, vbNarrow)
        Selection.Collapse wdCollapseEnd
    Loop
End Sub
    
Private Sub ‹Ö~•¶šŒŸ’m(ByVal strPattern As String)
' 2013/03/10
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = strPattern
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .MatchFuzzy = False
        .MatchWildcards = True
    End With
    If Selection.Find.Execute() = True Then
        MsgBox "‹Ö~•¶š‚ªŠÜ‚Ü‚ê‚Ä‚¢‚Ü‚·"
    End If
End Sub

