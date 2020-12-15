Attribute VB_Name = "Module1"

Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    Columns("A:A").EntireColumn.AutoFit
    Columns("B:B").EntireColumn.AutoFit
    Columns("C:C").EntireColumn.AutoFit
    Columns("D:D").EntireColumn.AutoFit
    
    'Testing Save Protocol
    
End Sub

Public Sub navyFederalGrab()

'Import the follow libraries tools > references
'1.Microsoft HTML Object Library
'2. Microsoft Internet Controls

'https://www.exceltrainingvideos.com/how-to-login-automatically-into-website-using-excel-vba/   (Logon Tutorial)
'https://www.youtube.com/watch?v=G05TrN7nt6k (General VBA Tutorial)
'https://codingislove.com/parse-html-in-excel-vba/

Dim HTMLDoc As HtmlDocument
Dim MyBrowser As InternetExplorer

Dim MyURL As String

MyURL = "https://www.navyfederal.org/"
Set MyBrowser = New InternetExplorer
MyBrowser.Silent = True
MyBrowser.navigate MyURL
MyBrowser.Visible = True

Do
Loop Until MyBrowser.readyState = READYSTATE_COMPLETE

Set HTMLDoc = MyBrowser.document
    
    HTMLDoc.getElementById("user").Value = "2751149"
    HTMLDoc.getElementById("password").Value = "++++++++++++++"
    HTMLDoc.getElementsByClassName("btn btn_sm toolbar__signin-btn")(0).Click
    
    
    Do While MyBrowser.readyState = 4: WScript.Sleep 100: Loop
    
    HTMLDoc.getElementsByClassName("account-summary-num")(0).Click
    HTMLDoc.getElementById("DownloadLink").Click
    HTMLDoc.getElementById("downloadTransactionsFileType").selectedIndex = 2
    
    HTMLDoc.getElementsByClassName("btn btn-primary btn-lg account-downloadtransactions-download-button")(0).Click
    
    'Doesnt work in step through
    'Application.SendKeys "%{S}"

    
End Sub
