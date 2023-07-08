'既に開いているchromeの操作
'コマンドプロンプトを開き↓の様なoption付きでchromeを開くと操作可能になる
'例："C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" -remote-debugging-port=9222 --user-data-dir=C:\remote_chrome    
Dim driver As New Selenium.WebDriver
driver.SetCapability Key:="debuggerAddress", Value:="localhost:9222"
driver.Start "chrome"

