'読み込み待ち

'---------------------------------------------------------------
'URLで判定------------------------------------------------------
'---------------------------------------------------------------

Dim waitCount
waitCount = 0

Do Until driver.Url = "targetURL"
    driver.Wait 1000
    waitCount = waitCount + 1
    
    If waitCount > 5 Then
        MsgBox "timeout"
        Exit Sub
    End If
    
Loop

'---------------------------------------------------------------
'HTMLで判定------------------------------------------------------
'---------------------------------------------------------------
Dim waitCount
waitCount = 0

Do Until InStr(driver.PageSource, "targetHTML") > 0
    driver.Wait 1000
    waitCount = waitCount + 1
    
    If waitCount > 5 Then
        MsgBox "timeout"
        Exit Sub
    End If
    
Loop

'---------------------------------------------------------------
'要素で判定------------------------------------------------------
'---------------------------------------------------------------
Dim myBy As New Selenium.By
Dim waitCount
waitCount = 0

Do Until driver.IsElementPresent(myBy.XPath("xpath")) 
    driver.Wait 1000
    waitCount = waitCount + 1
    
    If waitCount > 5 Then
        MsgBox "timeout"
        Exit Sub
    End If

Loop

'---------------------------------------------------------------
'時々エラーになる場合---------------------------------------------
'---------------------------------------------------------------
'判定の部分でエラーになる事がある
'エラーの内容はchromeのバージョンによって変わる
'手前にdriver.wait等を入れてもエラー率は下がらない
'1回目のエラーを許容して2回目の判定に進んだ場合エラーはおこらない
'根本的な解決方法は分からないけどこんな感じで回避できる

Dim waitCount As Long
waitCount = 0

Dim errorCount As Long
errorCount = 0
    
continue:
    On Error GoTo myError
    Do Until InStr(driver.PageSource, "targetHTML") > 0
        driver.Wait 1000
        waitCount = waitCount + 1
        
        If waitCount > 5 Then
            MsgBox "timeout"
            Exit Sub
        End If
        
    Loop
    On Error GoTo 0
    Exit Sub
    
myError:
    Select Case Err.Number
        Case Is = "想定しているエラー番号"
            errorCount = errorCount + 1
            If errorCount < 5 Then
                driver.Wait 1000
                Resume continue
            Else
                MsgBox "既定のエラー回数に到達しました"
                Exit Sub
            End If
        
        Case Else
            MsgBox "想定外のエラーが発生しました"
            Exit Sub
    End Select
