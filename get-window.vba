'目的のウィンドウを取得

'---------------------------------------------------------------
'タイトルが分かる場合---------------------------------------------
'---------------------------------------------------------------
If driver.Title <> "探すタイトル" Then
    On Error GoTo 見つからなかった場合の処理
    driver.SwitchToWindowByTitle ("探すタイトル")
    On Error GoTo 0
End If

'---------------------------------------------------------------
'タイトルが分からない場合-----------------------------------------
'---------------------------------------------------------------
'urlやページ内の要素で判断します。以下はURLで判断する方法
Dim i As Long
Dim isFound As Boolean

isFound = False
For i = 1 To driver.Windows.Count

    If driver.Url = "https://sekai-kabuka.com/" Then
        isFound = True
        Exit For
    End If
    
    '最後のウィンドウでSwitchToNextWindowを行うとエラーになるので確認
    If i < driver.Windows.Count Then
        driver.SwitchToNextWindow
    End If

Next

If Not isFound Then
    MsgBox "みつかりませんでした"
    Exit Sub
End If