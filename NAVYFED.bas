Attribute VB_Name = "Module1"

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
    HTMLDoc.getElementById("password").Value = "Football55!"
    HTMLDoc.getElementsByClassName("btn btn_sm toolbar__signin-btn")(0).Click
    
    
Application.Wait (Now + TimeValue("0:00:15"))

    HTMLDoc.getElementsByClassName("account-summary-num")(0).Click
    
Application.Wait (Now + TimeValue("0:00:05"))
    
    HTMLDoc.getElementById("DownloadLink").Click
    
Application.Wait (Now + TimeValue("0:00:03"))
    
    HTMLDoc.getElementById("downloadTransactionsFileType").selectedIndex = 2
    
Application.Wait (Now + TimeValue("0:00:02"))
    
    HTMLDoc.getElementsByClassName("btn btn-primary btn-lg account-downloadtransactions-download-button")(0).Click
    
Application.Wait (Now + TimeValue("0:00:05"))
    
    
    'Sub Routine Call
    CSVsaveAs
    
    MyBrowser.Quit
    
    Application.Wait (Now + TimeValue("0:00:03"))
    
    Dim theDate As String

    longDate = Now

    longDate = Replace(longDate, "/", ".")
    longDate = Replace(longDate, ":", ".")
    longDate = Replace(longDate, " ", "_")
    
    Name "C:\Users\nickr\Downloads\N.CSV" As "C:\Users\nickr\Documents\NAVYFEDERAL" & longDate & ".CSV"

    
End Sub

Public Sub CSVsaveAs()

    'Doesnt work in step through
    Application.SendKeys "%{N}"
    
    Application.Wait (Now + TimeValue("0:00:01"))
    Application.SendKeys "{TAB}"
    
    Application.Wait (Now + TimeValue("0:00:01"))
    Application.SendKeys "{DOWN}"
    
    Application.Wait (Now + TimeValue("0:00:01"))
    Application.SendKeys "{A}"
    
    Application.Wait (Now + TimeValue("0:00:03"))
    
    Application.SendKeys "{N}"
    
    Application.Wait (Now + TimeValue("0:00:03"))
    
    Application.SendKeys "~"
    
    Application.Wait (Now + TimeValue("0:00:10"))

End Sub

