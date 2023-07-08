'同じ命令でもPCによって動いたり動かなかったりすることがある
'（実行系の命令では時々あるが、取得する系の命令で起こった事は今のところない）
'AさんのPCでは操作1は動くが2は動かない
'BさんのPCでは操作2は動くが1は動かない
'この時エラーにはならず、実行したかの如く次の行へ進む
'根本的な解決方法が分からない場合はこんな感じでやる

'---------------------------------------------------------------
'buttonを押したら新しいウィンドウが開くケース----------------------
'---------------------------------------------------------------
Dim windowCount As Long
windowCount = driver.Windows.Count

'ウィンドウが開く操作1
driver.FindElementById("button").Click

If driver.Windows.Count = windowCount Then
    'ウィンドウが開く操作2
    driver.ExecuteScript "document.getElementById('button').click"
    
    If driver.Windows.Count = windowCount Then
        MsgBox "ウィンドウを開けませんでした"
        Exit Sub
    End If

End If

'---------------------------------------------------------------
'何かを入力するケース--------------------------------------------
'---------------------------------------------------------------

'入力1 sendkeysは上書きではなく追加入力なので先に消す
driver.FindElementById("text").Clear
driver.FindElementById("text").SendKeys "入力する文字"
'入力2 jsは上書きなのでそのまま入力
driver.ExecuteScript "document.getElementById('text').value='入力する文字'"

If driver.FindElementById("text").Value <> "入力する文字" Then
    MsgBox "入力に失敗しました"
End If