'アップロードのダイアログを開かずに添付するファイルを設定する

'---------------------------------------------------------------
'1つ------------------------------------------------------------
'---------------------------------------------------------------
driver.FindElementById("upfile").SendKeys "ファイルのフルパス"

'---------------------------------------------------------------
'複数------------------------------------------------------------
'---------------------------------------------------------------
driver.FindElementById("upfile").SendKeys "ファイルのフルパス1"
driver.FindElementById("upfile").SendKeys "ファイルのフルパス2"