VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Inputting_Data 
   Caption         =   "Inputting Data (v.1)"
   ClientHeight    =   8790.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13620
   OleObjectBlob   =   "Inputting_Data.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Inputting_Data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private currentPictureIndex As Integer
Private picturesFolderPath As String
Private pictureFiles As Collection

Private Sub Add_New_Trade_Click()
    ' Define the column numbers for each field
    Dim StartingBalance As Long: StartingBalance = 2
    Dim TickerColumn As Long: TickerColumn = 3 ' Column C
    Dim beforeTradeMAColumn As Long: beforeTradeMAColumn = 4 ' Column D
    Dim whyILikeTheTradeColumn As Long: whyILikeTheTradeColumn = 5 ' Column E
    Dim whyIDontLikeTheTradeColumn As Long: whyIDontLikeTheTradeColumn = 6 ' Column F
    Dim imageColumn As Long: imageColumn = 7 ' Column G
    Dim tradeRRColumn As Long: tradeRRColumn = 8 ' Column H
    Dim capitalPercentageRiskColumn As Long: capitalPercentageRiskColumn = 9 ' Column I
    Dim longOrShortColumn As Long: longOrShortColumn = 10 ' Column J
    Dim currentTimestampColumn As Long: currentTimestampColumn = 11 ' Column K
    Dim EndAccountBalance As Long: EndAccountBalance = 13 ' Column M
    Dim wsTrades As Worksheet
    Dim wsSetup As Worksheet
    Dim lastRow As Long
    
    Set wsTrades = ThisWorkbook.Worksheets("Trades")
    Set wsSetup = ThisWorkbook.Worksheets("Setup")
        
    Dim RR_Target As String
    RR_Target = wsSetup.Range("RRTarget").Value
        
    ' Check if any of the required textboxes are empty
    If Ticker_Combobox.Value = "" Then
        MsgBox "Ticker is required."
        Exit Sub
    ElseIf Before_Trade_MA.Value = "" Then
        MsgBox "Before Trade MA is required."
        Exit Sub
    ElseIf Long_or_short.Value = "" Then
        MsgBox "Long or Short is required."
        Exit Sub
    ElseIf Trade_RR_Target.Value = "" Then
        MsgBox "Trade RR Target is required."
        Exit Sub
    ElseIf Why_I_Like_The_Trade_ListBox.ListCount = 0 Then
        MsgBox "Why I Like the Trade is required."
        Exit Sub
    ElseIf Why_I_Dont_Like_The_Trade_ListBox.ListCount = 0 Then
        MsgBox "Why I Don't Like the Trade is required."
        Exit Sub
    ElseIf Trade_RR_Target.Value <= RR_Target Then
        Exit Sub
    End If
        
    Dim response As VbMsgBoxResult
    response = MsgBox("Do you want to add this trade?", vbYesNo + vbQuestion, "Confirmation")
    
    If response = vbYes Then
        
        lastRow = wsTrades.Cells(wsTrades.Rows.Count, TickerColumn).End(xlUp).Row
        
        On Error Resume Next
        wsTrades.Rows(lastRow + 1).Insert
        On Error GoTo 0 ' Disable error handling
        
        Dim lastrowBalance As Long
        lastrowBalance = wsTrades.Cells(wsTrades.Rows.Count, EndAccountBalance).End(xlUp).Row

        If lastRow >= 5 Then
            wsTrades.Cells(lastRow + 1, StartingBalance).Value = wsTrades.Cells(lastrowBalance, EndAccountBalance).Value
        Else
            wsTrades.Cells(lastRow + 1, StartingBalance).Value = wsSetup.Range("StartingAccountBalance").Value
        End If
        
        wsTrades.Cells(lastRow + 1, TickerColumn).Value = Ticker_Combobox.Value
        wsTrades.Cells(lastRow + 1, beforeTradeMAColumn).Value = Before_Trade_MA.Value
        
        Dim whyILikeItems As String
        For i = 0 To Why_I_Like_The_Trade_ListBox.ListCount - 1
            whyILikeItems = whyILikeItems & Why_I_Like_The_Trade_ListBox.List(i) & vbCrLf
        Next i
        wsTrades.Cells(lastRow + 1, whyILikeTheTradeColumn).Value = whyILikeItems
        
        Dim whyIDontLikeItems As String
        For i = 0 To Why_I_Dont_Like_The_Trade_ListBox.ListCount - 1
            whyIDontLikeItems = whyIDontLikeItems & Why_I_Dont_Like_The_Trade_ListBox.List(i) & vbCrLf
        Next i
        wsTrades.Cells(lastRow + 1, whyIDontLikeTheTradeColumn).Value = whyIDontLikeItems
               
        If currentPictureIndex > 0 And currentPictureIndex <= pictureFiles.Count Then
            On Error Resume Next
            Dim imagePath As String
            imagePath = pictureFiles.Item(currentPictureIndex)
            Me.ImagePreview.Picture = LoadPicture(imagePath)
            On Error GoTo 0
            
            If imagePath = "" Then
                MsgBox "Image not found."
                Exit Sub
            End If
            
            Dim ScaleValue As Integer
            ScaleValue = 10 ' 4x the size
            wsTrades.Cells(lastRow + 1, imageColumn).Activate
            
            ' Insert the Image and Resize
            wsTrades.Cells(lastRow + 1, imageColumn).ClearComments
            wsTrades.Cells(lastRow + 1, imageColumn).AddComment text:=""
            
            wsTrades.Cells(lastRow + 1, imageColumn).Hyperlinks.Add _
                Anchor:=wsTrades.Cells(lastRow + 1, imageColumn), _
                Address:=imagePath, _
                TextToDisplay:="Open Screenshot" ' Add a hyperlink to the folder location
            
            With wsTrades.Cells(lastRow + 1, imageColumn).Comment
                .Shape.Fill.UserPicture imagePath
                .Shape.LockAspectRatio = True
                .Shape.Width = ScaleValue * .Shape.Width
                .Visible = False
            End With
        End If
                    
        ' Run Cap Risk
        Dim RiskPercentLogicHolder As String
        If BE_Trade_Percent_Calc_Tickbox = True Then
            RiskPercentLogicHolder = "y"
        End If
        
        CalculateRiskPercentage
                  
        wsTrades.Cells(lastRow + 1, tradeRRColumn).Value = Trade_RR_Target.Value
        wsTrades.Cells(lastRow + 1, capitalPercentageRiskColumn).Value = Capital_Percentage_Risk_Calculator.Value
        wsTrades.Cells(lastRow + 1, longOrShortColumn).Value = Long_or_short.Value
        wsTrades.Cells(lastRow + 1, currentTimestampColumn).Value = Before_Trade_Current_Time.Value
        
        If RiskPercentLogicHolder <> "y" Then
            BE_Trade_Percent_Calc_Tickbox = False
            Capital_Percentage_Risk_Calculator.Value = ""
        End If
        
        ' Clear the textboxes and listboxes
        Ticker_Combobox.Value = ""
        Before_Trade_MA.Value = ""
        Long_or_short.Value = ""
        Trade_RR_Target.Value = ""
        Why_I_Like_The_Trade_ListBox.Clear
        Why_I_Dont_Like_The_Trade_ListBox.Clear
    End If

End Sub

Private Sub AF_Add_Combobox_add_Click()
    
    If AF_Add_Combobox = "" Then
        If AF_Add_Combobox_add = True Then
            AF_Thoughts_Listbox.AddItem AF_Add_Combobox.Value
            AF_Add_Combobox = ""
            AF_Add_Combobox_add = False
        End If
        MsgBox "You need to type something into the box to add it to the list", vbCritical
        AF_Add_Combobox_add = False
    End If
    
    
End Sub

Private Sub AF_Add_Combobox_add_2_Click()
    
    If AF_Add_Combobox_2 = "" Then
        If AF_Add_Combobox_add_2 = True Then
            AF_Thoughts_Listbox_2.AddItem AF_Add_Combobox_2.Value
            AF_Add_Combobox_2 = ""
            AF_Add_Combobox_add_2 = False
        End If
        MsgBox "You need to type something into the box to add it to the list", vbCritical
        AF_Add_Combobox_add_2 = False
    End If
    
End Sub


Private Sub AF_Refresh_Images_Button_Click()
    Set wsSetup = ThisWorkbook.Worksheets("Setup")
    picturesFolderPath = wsSetup.Range("J9").Value
    Set pictureFiles = New Collection
    LoadPicturesFromFolder
    SetMostRecentPicture
    UpdateImageAfterTrade
End Sub

Private Sub AF_Screenshot_Button_Click()
    Open_Screenshot_Folder
End Sub

Private Sub BE_Trade_End_Balance_Tickbox_Click()
    
    BF_Last_Ending_Balance_Textbox.Value = ""
    
    If BE_Trade_End_Balance_Tickbox = True Then
        Dim wsTrades As Worksheet
        Dim wsSetup As Worksheet
        Dim lastRow As Long
        Dim endingBalance As Variant
        Dim currencySymbol As String
        
        Set wsTrades = ThisWorkbook.Sheets("Trades")
        Set wsSetup = ThisWorkbook.Sheets("Setup")
        lastRow = wsTrades.Cells(wsTrades.Rows.Count, "M").End(xlUp).Row
        
        If lastRow > 4 Then
            endingBalance = wsTrades.Cells(lastRow, "M").Value
        Else
            endingBalance = wsTrades.Range("B5").Value
        End If

        ' Get the currency symbol from Setup sheet cell G9
        currencySymbol = wsSetup.Range("G9").Value
        
        ' Check if endingBalance is not empty and doesn't already start with the currency symbol
        If Not IsEmpty(endingBalance) And Left(endingBalance, Len(currencySymbol)) <> currencySymbol Then
            endingBalance = currencySymbol & endingBalance ' Add the currency symbol
        End If
        
        BF_Last_Ending_Balance_Textbox.Value = endingBalance
    End If
End Sub


Private Sub BE_Trade_Percent_Calc_Tickbox_Click()
    If BE_Trade_Percent_Calc_Tickbox.Value = True Then
        CalculateRiskPercentage
    Else
        Capital_Percentage_Risk_Calculator.Value = ""
    End If
End Sub


Private Sub BF_Open_Image_Location_Click()
    Open_Screenshot_Folder
End Sub

Sub Open_Screenshot_Folder()
    Dim folderPath As String
    
    ' Get the directory path from cell J9 of the "Setup" sheet
    folderPath = ThisWorkbook.Sheets("Setup").Range("J9").Value
    
    ' Check if the folder path is not empty
    If Len(folderPath) > 0 Then
        ' Check if the folder exists
        If Dir(folderPath, vbDirectory) <> "" Then
            ' Open the directory using Windows Explorer
            Call Shell("explorer.exe """ & folderPath & """", vbNormalFocus)
        Else
            MsgBox "The specified directory does not exist." & vbNewLine & "Enter your screenshot folder path in the setup sheet in cell J9", vbCritical
        End If
    Else
        MsgBox "The directory path is empty.", vbExclamation
    End If
End Sub

Private Sub BF_Refresh_Images_Button_Click()
    Set wsSetup = ThisWorkbook.Worksheets("Setup")
    picturesFolderPath = wsSetup.Range("J9").Value
    
    Set pictureFiles = New Collection
    LoadPicturesFromFolder
    SetMostRecentPicture
    UpdateImage
    
End Sub

Private Sub BF_Refresh_Trade_Time_Click()
    Before_Trade_Current_Time = Format(Now, "dd/mm/yyyy hh:mm")
End Sub

Private Sub Complete_Trade_Button_Click()
    Dim response As VbMsgBoxResult
    response = MsgBox("Do you want to add this trade?", vbYesNo + vbQuestion, "Confirmation")

    If response = vbYes Then
        Dim imageColumn As Long: imageColumn = 15
        Dim TradeEndTimeColumn As Long: TradeEndTimeColumn = 14
        Dim EndAccountBalanceColumn As Long: EndAccountBalanceColumn = 13
        Dim TakeHomeProfitColumn As Long: TakeHomeProfitColumn = 12
        Dim startingBalanceColumn As Long: startingBalanceColumn = 2
        Dim wsTrades As Worksheet
        Set wsTrades = ThisWorkbook.Worksheets("Trades")
        Dim TickerColumn As Long
        Dim rrColumn As Long
        Dim selectedValue As String
        Dim lastRow As Long
        
        ' Define the columns
        TickerColumn = 3
        TradeStartColumn = 11
        
        If Select_Trade_Listbox.ListIndex = -1 Then
            MsgBox "You need to select the trade you've completed", vbCritical
            Exit Sub
        End If
        
        selectedValue = Select_Trade_Listbox.List(Select_Trade_Listbox.ListIndex)
        
        Dim parts() As String
        parts = Split(selectedValue, " | ")
        If UBound(parts) <= 0 Then
            MsgBox "Invalid selected value."
            Exit Sub
        End If
                
        lastRow = wsTrades.Cells(wsTrades.Rows.Count, TickerColumn).End(xlUp).Row
        For i = 5 To lastRow
            If wsTrades.Cells(i, TickerColumn).Value = parts(0) And wsTrades.Cells(i, TradeStartColumn).Value = parts(1) Then
                FoundRow = i
                Exit For
            End If
        Next i
        
        If i > FoundRow Then
            MsgBox "Row not found for selected data."
            Exit Sub
        End If
        
        ' Check other conditions
        If Not IsNumeric(After_Trade_Profit_Loss.Value) Or After_Trade_Profit_Loss.Value < 0 Then
            MsgBox "Invalid profit/loss value."
            Exit Sub
        End If
        
        If Not (Profit_Confirmed_Tickbox.Value Xor Loss_Confirmed_Tickbox.Value) Then
            MsgBox "Select either Profit Confirmed or Loss Confirmed."
            Exit Sub
        End If
        
        If Not IsDate(After_Trade_Close_Time.Value) Then
            MsgBox "Invalid date/time value."
            Exit Sub
        End If
        
        ' Profit Value
        wsTrades.Cells(FoundRow, TakeHomeProfitColumn).Value = IIf(Loss_Confirmed_Tickbox.Value, -After_Trade_Profit_Loss.Value, After_Trade_Profit_Loss.Value)
        ' End Account Balance Calculation
        wsTrades.Cells(FoundRow, EndAccountBalanceColumn).Value = wsTrades.Cells(FoundRow, startingBalanceColumn).Value + wsTrades.Cells(FoundRow, TakeHomeProfitColumn).Value
        
        Dim closingTime As String
        closingTime = Format(After_Trade_Close_Time.Value, "dd/mm/yyyy hh:mm")
        wsTrades.Cells(FoundRow, TradeEndTimeColumn).Value = closingTime
           
        Dim imagePath As String
        If currentPictureIndex > 0 And currentPictureIndex <= pictureFiles.Count Then
            imagePath = pictureFiles.Item(currentPictureIndex)
            Me.ImagePreviewAfterTrade.Picture = LoadPicture(imagePath)
            If imagePath = "" Then
                MsgBox "Image not found."
                Exit Sub
            End If
            
            Dim ScaleValue As Integer
            ScaleValue = 10
            wsTrades.Cells(FoundRow, imageColumn).Activate
            
            ' Insert the Image and Resize
            wsTrades.Cells(FoundRow, imageColumn).ClearComments
            wsTrades.Cells(FoundRow, imageColumn).AddComment text:=""
            
            wsTrades.Cells(FoundRow, imageColumn).Hyperlinks.Add _
                Anchor:=wsTrades.Cells(FoundRow, imageColumn), _
                Address:=imagePath, _
                TextToDisplay:="Open Folder" ' Add a hyperlink to the folder location
            
            With wsTrades.Cells(FoundRow, imageColumn).Comment
                .Shape.Fill.UserPicture imagePath
                .Shape.LockAspectRatio = True
                .Shape.Width = ScaleValue * .Shape.Width
                .Visible = False
            End With
        End If
        
        
        ' Add all items from AF_Thoughts_Listbox into column P in the same row
        Dim thoughts As String
        For i = 0 To AF_Thoughts_Listbox.ListCount - 1
            thoughts = thoughts & AF_Thoughts_Listbox.List(i) & vbCrLf
        Next i
        wsTrades.Cells(FoundRow, 17).Value = thoughts
        
        thoughts = ""
        
        For i = 0 To AF_Thoughts_Listbox_2.ListCount - 1
            thoughts = thoughts & AF_Thoughts_Listbox_2.List(i) & vbCrLf
        Next i
        wsTrades.Cells(FoundRow, 16).Value = thoughts
        
        
    End If
End Sub


Private Sub HP_High_Impact_News_Click()
    Dim wsSetup As Worksheet
    Set wsSetup = ThisWorkbook.Worksheets("Setup")
    If wsSetup.Range("F15").Value = "y" Then
        Populatenews
    End If
End Sub

Private Sub HP_Todays_News_Click()
    Dim wsSetup As Worksheet
    Set wsSetup = ThisWorkbook.Worksheets("Setup")
    If wsSetup.Range("F15").Value = "y" Then
        Populatenews
    End If
End Sub


Private Sub Lot_Size_Calculator_Click()
    Dim accountSize As Double
    Dim entryPrice As Double
    Dim stopLoss As Double
    Dim riskPercent As Double
    Dim lotsize As Double

    ' Check if the required textboxes are empty
    If BF_Entry_Price_Calc_Textbox.Value = "" Then
        MsgBox "You need to enter an entry price", vbCritical, "Missing Information"
        Exit Sub
    End If

    If BF_Stop_Loss_Percent_Textbox.Value = "" Then
        MsgBox "You need to enter a stop loss percentage", vbCritical, "Missing Information"
        Exit Sub
    End If
    
    ' Get the account size from the last ending balance
    Dim EndBalPrevLogicHolder As String
    If BE_Trade_End_Balance_Tickbox = True Then
        EndBalPrevLogicHolder = "y"
    End If
    
    BE_Trade_End_Balance_Tickbox = True
    BE_Trade_End_Balance_Tickbox_Click
    accountSize = CDbl(BF_Last_Ending_Balance_Textbox.Value)
    
    If EndBalPrevLogicHolder <> "y" Then
        BE_Trade_End_Balance_Tickbox = False
        BF_Last_Ending_Balance_Textbox.Value = ""
    End If

    ' Get the entry price, stop-loss percent, and risk percent from the respective textboxes
    entryPrice = CDbl(BF_Entry_Price_Calc_Textbox.Value)
    stopLoss = CDbl(BF_Stop_Loss_Percent_Textbox.Value)
    
    Dim RiskPercentLogicHolder As String
    If BE_Trade_Percent_Calc_Tickbox = True Then
        RiskPercentLogicHolder = "y"
    End If
    
    CalculateRiskPercentage
          
    ' Remove the percentage symbol (%) from the string and then convert to a Double
    Dim riskPercentStr As String
    riskPercentStr = Capital_Percentage_Risk_Calculator.Value

    ' Check if the string ends with a percentage symbol
    If Right(riskPercentStr, 1) = "%" Then
        ' Remove the percentage symbol before converting to Double
        riskPercentStr = Left(riskPercentStr, Len(riskPercentStr) - 1)
    End If

    ' Convert the modified string to a Double
    riskPercent = CDbl(riskPercentStr)
    
    If riskPercent <= 0 Then
        MsgBox "Account Risk must be greater than zero.", vbCritical, "Invalid Input"
        Exit Sub
    End If
    
    If stopLoss <= 0 Then
        MsgBox "Stop Loss Value must be greater than zero.", vbCritical, "Invalid Input"
        Exit Sub
    End If
    
    ' Check if riskPercent is greater than 0 to avoid division by zero
    If riskPercent <= 0 Then
        MsgBox "Risk Percentage must be greater than zero.", vbCritical, "Invalid Input"
        Exit Sub
    End If
' ---
    If RiskPercentLogicHolder <> "y" Then
        BE_Trade_Percent_Calc_Tickbox = False
        Capital_Percentage_Risk_Calculator.Value = ""
    End If
' ---
    ' Calculate the lot size based on the formula
    Dim riskAmount As Double
    riskAmount = (accountSize * riskPercent / 100)
    lotsize = (riskAmount / (entryPrice - stopLoss)) / BF_Contracts_Per_Lot
    
    ' Check if the calculated lot size is negative, and if so, set it to 0
    If lotsize < 0 Then
        lotsize = -lotsize
    End If
    
    ' Format the lot size to display only 2 decimals and assign it to the textbox
    BF_Stop_Loss_Calculator_Textbox.Value = Format(lotsize, "0.000")

End Sub

Private Sub Open_Broker_Click()
    
    Set wsSetup = ThisWorkbook.Worksheets("Setup")
    
    Dim url As String
    url = wsSetup.Range("J4").Value
    Dim Chrome As String
    Chrome = wsSetup.Range("J7").Value
        
    On Error Resume Next
    Shell Chrome & " " & url, vbNormalFocus
    If Err.Number <> 0 Then
        MsgBox "Unable to open the web page.", vbExclamation, "Error"
    End If
    On Error GoTo 0
End Sub

Private Sub Open_News_Click()
    Set wsSetup = ThisWorkbook.Worksheets("Setup")
    
    Dim url As String
    url = wsSetup.Range("J5").Value
    Dim Chrome As String
    Chrome = wsSetup.Range("J7").Value
        
    On Error Resume Next
    Shell Chrome & " " & url, vbNormalFocus
    If Err.Number <> 0 Then
        MsgBox "Unable to open the web page.", vbExclamation, "Error"
    End If
    On Error GoTo 0
End Sub

Sub Open_Trading_View_Click()
    Set wsSetup = ThisWorkbook.Worksheets("Setup")
    
    Dim url As String
    url = wsSetup.Range("J3").Value
    Dim Chrome As String
    Chrome = wsSetup.Range("J7").Value
        
    On Error Resume Next
    Shell Chrome & " " & url, vbNormalFocus
    If Err.Number <> 0 Then
        MsgBox "Unable to open the web page.", vbExclamation, "Error"
    End If
    On Error GoTo 0
End Sub

Sub MotivationalQuotesSub()

    Dim wsSetup As Worksheet
    Set wsSetup = ThisWorkbook.Sheets("Setup")
    
    Dim lastRow As Long
    lastRow = wsSetup.Cells(wsSetup.Rows.Count, "L").End(xlUp).Row
    
    If lastRow < 3 Then
        MsgBox "No motivational quotes found in the 'Setup' sheet.", vbcritcal
        Exit Sub
    End If
    
    Dim randomIndex As Long
    randomIndex = Int((lastRow - 2) * Rnd + 1) ' Subtract 2 because quotes start from row 3
        
    Me.Quote_Of_Day_Textbox.Value = wsSetup.Cells(randomIndex + 2, "L").Value

End Sub

Sub Populatenews()
    Dim IE As Object
    Dim HTMLDoc As Object
    Dim url As String
    
    Dim currentTime As String

    If Hour(Now) < 12 Then
        currentTime = Format(Now, "hh:mm") & "am"
    Else
        currentTime = Format(Now, "hh:mm") & "pm"
    End If
    
    Home_Page_Time_Textbox.Value = Format(Now, "ddd | MMM dd | ") & currentTime
        
    Set wsSetup = ThisWorkbook.Worksheets("Setup")
    url = wsSetup.Range("J5").Value
    
    Set IE = CreateObject("InternetExplorer.Application")

    With IE
        .Visible = False ' Set to True if you want to see the IE window
        .navigate url
        
        Do While .Busy Or .readyState <> 4
            DoEvents
        Loop
        
        Set HTMLDoc = .document
    End With
    
    ' Get all rows in the table with class "calendar__row"
    Dim tableRows As Object
    Set tableRows = HTMLDoc.getElementsByClassName("calendar__row")
    
    News_Description_Combobox.Clear
    
    ' Initialize previous day, time, and impact
    Dim previousDay As String
    Dim previousTime As String
    Dim previousImpact As String
    previousDay = ""
    previousTime = ""
    previousImpact = ""
    
    ' Loop through each row and extract time, date, impact, and event
    For Each Row In tableRows
        ' Try to get the time, date, impact, and description for each row
        On Error Resume Next
        Dim timeElement As Object
        Set timeElement = Row.getElementsByClassName("calendar__cell calendar__time")(0)
        
        Dim dateElement As Object
        Set dateElement = Row.getElementsByClassName("calendar__cell calendar__date")(0)
        
        Dim impactElement As Object
        Set impactElement = Row.getElementsByClassName("calendar__cell calendar__impact")(0)
        
        Dim descriptionElement As Object
        Set descriptionElement = Row.getElementsByClassName("calendar__event-title")(0)
        On Error GoTo 0 ' Reset error handling
        
        ' Check if descriptionElement is found
        If Not descriptionElement Is Nothing Then
            ' Extract the date and description
            Dim dateText As String
            dateText = dateElement.innerText
            Dim descriptionText As String
            descriptionText = descriptionElement.innerText
            
            ' Check if timeElement is found and not empty
            If Not timeElement Is Nothing And timeElement.innerText <> "" Then
                ' Use the current time
                previousTime = timeElement.innerText
            End If
            
            ' Check if dateElement is found and not empty
            If Not dateElement Is Nothing And dateElement.innerText <> "" Then
                ' Use the current date
                previousDay = dateElement.innerText
            End If
            
            ' Check if the time contains "Day" before formatting
            If InStr(previousTime, "Day") = 0 Then
                ' Format time with leading zero if needed
                If Len(previousTime) <= 6 Then
                    previousTime = "0" & previousTime
                End If
            End If
            
            ' Check if impactElement is found and represents a high impact
            If Not impactElement Is Nothing And InStr(impactElement.innerHTML, "icon--ff-impact-red") > 0 Then
                previousImpact = "High"
            ElseIf Not impactElement Is Nothing And InStr(impactElement.innerHTML, "icon--ff-impact-ora") > 0 Then
                previousImpact = "Medium"
            ElseIf Not impactElement Is Nothing And InStr(impactElement.innerHTML, "icon--ff-impact-yel") > 0 Then
                previousImpact = "Low"
            Else
                previousImpact = "Unknown"
            End If
            
            Dim combinedInfo As String
            combinedInfo = previousDay & " | " & previousTime & " | " & previousImpact & " | " & descriptionText
            
            Dim SpecificDayAbriv As String
            SpecificDayAbriv = Left(previousDay, 3)
                        
            If HP_Todays_News.Value = True And SpecificDayAbriv <> Format(Now, "ddd") Then
                GoTo Skip
            End If

            If HP_High_Impact_News = True And Not previousImpact = "High" Then
                GoTo Skip
            End If

            ' Add to the ComboBox if it passes the filters
            News_Description_Combobox.AddItem combinedInfo

        End If
Skip:
    Next Row
    
    IE.Quit
    Set IE = Nothing
    Set HTMLDoc = Nothing
End Sub
       
Private Function QualityChecks() As Boolean

        ' Initialize error flag
    Dim HasError As Boolean
    HasError = False

    ' =============== Quality Checks
    
    ' Read website URLs and folder paths from cells
    Dim tradingViewURL As String
    Dim brokerURL As String
    Dim financialNewsURL As String
    Dim discordURL As String
    Dim chromePath As String
    Dim spotifyPath As String
    Dim screenshotsPath As String
    
    tradingViewURL = ThisWorkbook.Sheets("Setup").Range("J3").Value
    brokerURL = ThisWorkbook.Sheets("Setup").Range("J4").Value
    financialNewsURL = ThisWorkbook.Sheets("Setup").Range("J5").Value
    discordURL = ThisWorkbook.Sheets("Setup").Range("J6").Value
    chromePath = ThisWorkbook.Sheets("Setup").Range("J7").Value
    spotifyPath = ThisWorkbook.Sheets("Setup").Range("J8").Value
    screenshotsPath = ThisWorkbook.Sheets("Setup").Range("J9").Value
    
    ' Check website availability
    On Error Resume Next
    Dim xhr As Object
    Set xhr = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    ' TradingView
    xhr.Open "GET", tradingViewURL, False
    xhr.send ""
    If xhr.status <> 200 Then
        MsgBox "TradingView website is not accessible.", vbCritical
        HasError = True
    End If
    
    ' Broker
    xhr.Open "GET", brokerURL, False
    xhr.send ""
    If xhr.status <> 200 Then
        MsgBox "Broker website is not accessible.", vbCritical
        HasError = True
    End If
    
    ' Financial News
    xhr.Open "GET", financialNewsURL, False
    xhr.send ""
    If xhr.status <> 200 Then
        MsgBox "Financial News website is not accessible.", vbCritical
        HasError = True
    End If
    
    On Error GoTo 0
    
    ' Check folder existence
    If Dir(chromePath, vbDirectory) = "" Then
        MsgBox "Google Chrome folder does not exist.", vbCritical
        HasError = True
    End If
    
    If Dir(screenshotsPath, vbDirectory) = "" Then
        MsgBox "Screenshots folder does not exist.", vbCritical
        HasError = True
    End If

    ' Check for errors
    If HasError Then
        QualityChecks = True ' Set the function to True to indicate errors
        Exit Function
    End If

    ' No errors, return False
    QualityChecks = False
End Function
        
Private Sub Select_Trade_Listbox_Click()
    Dim wsTrades As Worksheet
    Dim selectedTrade As String
    Dim selectedRow As Long
    
    ' Set the relevant worksheet
    Set wsTrades = ThisWorkbook.Worksheets("Trades")
    
    ' Get the selected trade from the listbox
    selectedTrade = Select_Trade_Listbox.Value
    
    ' Find the selected trade in column F and get the corresponding value from column F
    selectedRow = 0 ' Initialize the selected row
    
    For i = 5 To wsTrades.Cells(wsTrades.Rows.Count, 3).End(xlUp).Row ' Assuming column C contains the trade identifier (Ticker)
        If wsTrades.Cells(i, 3).Value & " | " & wsTrades.Cells(i, 11).Value = selectedTrade Then
            selectedRow = i
            Exit For
        End If
    Next i
    
    ' Check if a matching trade was found
    If selectedRow > 0 Then
        ' Populate the AF_Add_Combobox with data from column F (assuming it's column 6)
        AF_Add_Combobox.Clear
        AF_Add_Combobox.AddItem wsTrades.Cells(selectedRow, 6).Value
    Else
        ' No matching trade found
        MsgBox "Selected trade not found.", vbExclamation, "Trade Not Found"
    End If
End Sub

Private Sub Settings_Demo_Trade_Tracking_Click()

    
    Dim wsSetup As Worksheet
    Set wsSetup = ThisWorkbook.Worksheets("Setup")
    
    ' Check the current value
    Dim currentValue As Variant
    currentValue = wsSetup.Range("G15").Value
    
    ' Check if the new value is different from the current value
    If Settings_Demo_Trade_Tracking.Value = True And currentValue <> "y" Then
        wsSetup.Range("G15").Value = "y"
    ElseIf Settings_Demo_Trade_Tracking.Value = False And currentValue <> "n" Then
        wsSetup.Range("G15").Value = "n"
    End If


End Sub

Private Sub Settings_Reset_Account_Button_Click()

    Dim wsTrades As Worksheet
    Set wsTrades = ThisWorkbook.Worksheets("Trades")
    Dim wsPastTrades As Worksheet
    Set wsPastTrades = ThisWorkbook.Worksheets("Past Trades")
    
    ' Check if a reason is entered
    If Settings_Account_Reset_Reason.Value = "" Then
        MsgBox "Please enter the reason why the account is getting reset", vbCritical
        Exit Sub
    End If
    
    ' Check if a new starting balance is entered
    If Settings_New_Account_Balance.Value = "" Then
        MsgBox "Please enter the new starting balance", vbCritical
        Exit Sub
    End If
    
    ' Initialize variables for counting
    Dim lastRow As Long
    Dim tradeCount As Long
    Dim missingTradeCount As Long
    Dim starttime As Date
    Dim endtime As Date
    
    lastRow = wsTrades.Cells(wsTrades.Rows.Count, 14).End(xlUp).Row ' Assuming data goes up to column Q
    starttime = wsTrades.Range("K5").Value
    endtime = wsTrades.Range("N" & lastRow).Value
    
    For i = 5 To lastRow
        If wsTrades.Cells(i, 3).Value <> "" And wsTrades.Cells(i, 14).Value = "" Then
            missingTradeCount = missingTradeCount + 1
        End If
        tradeCount = tradeCount + 1
    Next i
    
    ' Check if there are trades with missing information
    If missingTradeCount > 0 Then
        MsgBox "You have " & missingTradeCount & " trade(s) with missing information in column N." & vbNewLine & _
               "Please complete all trades before resetting the account." & vbNewLine & _
               "You can do this in the After trade tab", vbExclamation
        Exit Sub
    ElseIf tradeCount = 0 Then
        MsgBox "You've entered no trades", vbInformation
        Exit Sub
    End If
    
    Dim ConfirmationQuestion As Variant
    ConfirmationQuestion = MsgBox("You've entered " & tradeCount & " trade(s)" & vbNewLine & _
    "This is from the time period starting :" & starttime & vbNewLine & _
    "Up until :" & endtime & vbNewLine & _
    "Are you sure you want to reset the account?", vbYesNo + vbExclamation)
    
    If ConfirmationQuestion = vbNo Then Exit Sub
    
    Dim lastRowWithData As Long
    lastRowWithData = wsPastTrades.Cells(wsPastTrades.Rows.Count, "D").End(xlUp).Row + 2
    
    wsTrades.Range("B5:Q" & lastRow).Copy
    
    wsPastTrades.Cells(lastRowWithData, 4).End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
    
    lastRowWithData = wsPastTrades.Cells(wsPastTrades.Rows.Count, "D").End(xlUp).Row + 1
    

    wsPastTrades.Range("B" & lastRowWithData & ":S" & lastRowWithData).Merge
    wsPastTrades.Range("B" & lastRowWithData & ":S" & lastRowWithData).Interior.Color = RGB(255, 192, 0)
    
    wsPastTrades.Cells(lastRowWithData + 1, 2).Value = Settings_Account_Reset_Reason.Value
    wsPastTrades.Cells(lastRowWithData + 1, 19).Value = Settings_New_Account_Balance.Value ' Adjusted to column S

    'wsTrades.Range("C5:Q" & lastRow).Clear
    'wsTrades.Range("B5").Value = Settings_New_Account_Balance.Value
    
    ' Clear the reason and balance fields
    Settings_Account_Reset_Reason.Value = ""
    Settings_New_Account_Balance.Value = ""
    
    MsgBox "Account has been reset and past trades are recorded.", vbInformation
    
End Sub

Private Sub Settings_Show_News_Change()
    
    Dim wsSetup As Worksheet
    Set wsSetup = ThisWorkbook.Worksheets("Setup")
    
    ' Check the current value
    Dim currentValue As Variant
    currentValue = wsSetup.Range("F15").Value
    
    ' Check if the new value is different from the current value
    If Settings_Show_News.Value = True And currentValue <> "y" Then
        wsSetup.Range("F15").Value = "y"
        MsgBox "Please wait while the news loads", "Updating"
    ElseIf Settings_Show_News.Value = False And currentValue <> "n" Then
        wsSetup.Range("F15").Value = "n"
    End If
End Sub

Private Sub UserForm_Initialize()
    
    If QualityChecks = True Then
        'Unload Me ' Unload the UserForm if errors were found
        Exit Sub
    End If

    ' Continue initializing the UserForm
    MultiPage1.Value = 1
    Application.Wait Now + TimeValue("00:00:01")
    MultiPage1.Value = 0
End Sub

Private Sub MultiPage1_Change()
    Select Case MultiPage1.Value
        Case 0 ' Page1 is selected
            HomePageSub
        Case 1
            BeforeTradeSub
        Case 2
            AfterTradeSub
        Case 3
            SettingsSub
    End Select
End Sub

Private Sub HomePageSub()

    Me.Home_Page_Time_Textbox.Value = Now
    MotivationalQuotesSub
    
    Dim wsSetup As Worksheet
    Set wsSetup = ThisWorkbook.Worksheets("Setup")
    If wsSetup.Range("F15").Value = "y" Then
        Populatenews
    End If
        

End Sub

Private Sub BeforeTradeSub()
    
    Set wsSetup = ThisWorkbook.Worksheets("Setup")
    picturesFolderPath = wsSetup.Range("J9").Value
    
    Set pictureFiles = New Collection
    LoadPicturesFromFolder
    SetMostRecentPicture
    UpdateImage
    
    ' Calculate the current time
    Before_Trade_Current_Time = Format(Now, "dd/mm/yyyy hh:mm")

    ' Populate Long_or_short ComboBox
    Long_or_short.Clear
    Why_I_Like_The_Trade.Clear
    Why_I_dont_Like_The_Trade.Clear
    
    Long_or_short.AddItem "Long"
    Long_or_short.AddItem "Short"
    
    'Demo Account Warning
    If wsSetup.Range("G15").Value = "y" Then
        BE_Demo_Trading_Active.Value = "Warning! Demo Tracking Enabled"
    Else
        BE_Demo_Trading_Active.Value = ""
    End If
    
    Dim setupSheet As Worksheet
    Set setupSheet = Worksheets("Setup")
    
    Dim lastRowBeforeTrade As Long
    Dim lastRowWhyLike As Long
    Dim lastRowWhyDontLike As Long
    
    lastRowBeforeTrade = setupSheet.Cells(setupSheet.Rows.Count, "B").End(xlUp).Row
    lastRowWhyLike = setupSheet.Cells(setupSheet.Rows.Count, "C").End(xlUp).Row
    lastRowWhyDontLike = setupSheet.Cells(setupSheet.Rows.Count, "D").End(xlUp).Row
        
    Dim i As Long
    Before_Trade_MA.Clear
    For i = 3 To lastRowBeforeTrade
        Before_Trade_MA.AddItem setupSheet.Range("B" & i).Value
    Next i
    
    For i = 3 To lastRowWhyLike
        Why_I_Like_The_Trade.AddItem setupSheet.Range("C" & i).Value
    Next i
    
    For i = 3 To lastRowWhyDontLike
        Why_I_dont_Like_The_Trade.AddItem setupSheet.Range("D" & i).Value
    Next i
    
    CalculateTicker
    
    
End Sub

Sub SettingsSub()
    ' Check if news is activated or deactivated
    Dim wsSetup As Worksheet
    Set wsSetup = ThisWorkbook.Worksheets("Setup")
    
    ' Check if the change is due to user interaction
    If wsSetup.Range("F15").Value = "y" Then
        Settings_Show_News.Value = True
    Else
        Settings_Show_News.Value = False
    End If

    If wsSetup.Range("G15").Value = "y" Then
        Settings_Demo_Trade_Tracking.Value = True
    Else
        Settings_Demo_Trade_Tracking.Value = False
    End If


End Sub

Private Sub CalculateTicker()
    Dim wsTrades As Worksheet
    Dim lastRow As Long
    Dim tickers As Collection
    Dim ticker As Variant
    Dim tickerCounts As Object
    Dim sortedTickers() As Variant
    Dim i As Long, j As Long
    Dim temp As Variant
    Dim tickerCount As Long
    
    ' Set the relevant worksheet
    Set wsTrades = ThisWorkbook.Worksheets("Trades")
    

    Ticker_Combobox.Clear
    lastRow = wsTrades.Cells(wsTrades.Rows.Count, 3).End(xlUp).Row
    
    ' Create a collection to store unique tickers
    Set tickers = New Collection
    ' Create a dictionary to store ticker occurrence counts
    Set tickerCounts = CreateObject("Scripting.Dictionary")
    
    ' Loop through column C starting from row 5 and add unique values to the collection
    On Error Resume Next
    For i = 5 To lastRow
        tickers.Add wsTrades.Cells(i, 3).Value, CStr(wsTrades.Cells(i, 3).Value)
        
        If Not tickerCounts.Exists(wsTrades.Cells(i, 3).Value) Then
            tickerCounts(wsTrades.Cells(i, 3).Value) = 1
        Else
            tickerCounts(wsTrades.Cells(i, 3).Value) = tickerCounts(wsTrades.Cells(i, 3).Value) + 1
        End If
    Next i
    On Error GoTo 0
    
    If tickers.Count = 0 Then Exit Sub

    ReDim sortedTickers(1 To tickerCounts.Count, 1 To 2)
    i = 1
    For Each ticker In tickers
        sortedTickers(i, 1) = ticker
        sortedTickers(i, 2) = tickerCounts(ticker)
        i = i + 1
    Next ticker
    
    For i = LBound(sortedTickers) To UBound(sortedTickers) - 1
        For j = i + 1 To UBound(sortedTickers)
            If sortedTickers(j, 2) > sortedTickers(i, 2) Then
                ' Swap elements
                temp = sortedTickers(i, 1)
                sortedTickers(i, 1) = sortedTickers(j, 1)
                sortedTickers(j, 1) = temp
                
                temp = sortedTickers(i, 2)
                sortedTickers(i, 2) = sortedTickers(j, 2)
                sortedTickers(j, 2) = temp
            End If
        Next j
    Next i
    
    For i = LBound(sortedTickers) To UBound(sortedTickers)
        Ticker_Combobox.AddItem sortedTickers(i, 1)
    Next i
End Sub

Private Sub CalculateRiskPercentage()
    Dim wsTrades As Worksheet
    Dim wsSetup As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim riskRange As Range
    
    ' Set the relevant worksheets
    Set wsTrades = ThisWorkbook.Worksheets("Trades")
    Set wsSetup = ThisWorkbook.Worksheets("Setup")
    
    ' Define the column numbers
    Dim ProfitLossColumn As Long: ProfitLossColumn = 12
    Dim riskColumn As Long: riskColumn = 9 ' Column I (Capital Risk %)
    Dim startingBalanceColumn As Long: startingBalanceColumn = 2 ' Column B (Starting Account Balance)
    Dim TickerColumn As Long: TickerColumn = 3
    Dim riskPercentage As Double
    Dim matchedRow As Long: matchedRow = 0
    
    ' Define the risk range from the Setup sheet
    Set riskRange = wsSetup.Range("F3:F6")
    
    lastRow = wsTrades.Cells(wsTrades.Rows.Count, TickerColumn).End(xlUp).Row
    
    ' Get Previous Risk
    If lastRow < 5 Then
        riskPercentage = wsSetup.Range("F3").Value
    Else
        previousRiskPercentage = wsTrades.Cells(lastRow, riskColumn).Value
    End If
    
    ' Calculate Amount of losses
    For i = lastRow To 5 Step -1
        If wsTrades.Cells(i, ProfitLossColumn).Value < 0 Then
            lossCount = lossCount + 1
        Else
            Exit For
        End If
    Next i
        
    ' Find Indexable value
    Dim formattedRiskPercentage As String
    formattedRiskPercentage = CDbl(previousRiskPercentage)
    
    For i = 3 To 6
        If formattedRiskPercentage = wsSetup.Range("F" & i).Value Then
            FoundRiskCellRowRange = i
            Exit For
        End If
    Next i
            
    ' ============= INCREASE RISK
    If lossCount < 1 Then
        ' JUMP TO END OF CODE BECAUSE NO VALUES HAVE BEEN ENTERED
        If FoundRiskCellRowRange = "" Then
            previousRiskPercentage = wsSetup.Range("F3").Value
            GoTo JumpSpot
        End If
                
        'lastRow = wsTrades.Cells(wsTrades.Rows.count, riskColumn).End(xlUp).Row
                
        ' Find the next tier
        NextRiskTier = wsSetup.Range("F" & (FoundRiskCellRowRange - 1)).Value
                        
        If FoundRiskCellRowRange <= 3 Then
            NextRiskTier = wsSetup.Range("F3").Value
        Else
            NextRiskTier = wsSetup.Range("F" & (FoundRiskCellRowRange - 1)).Value
        End If
        
        ' Search for the location of formattedRiskPercentage within the risk column
        For i = lastRow To 5 Step -1 ' Start from the last row and move upwards
            If wsTrades.Cells(i, riskColumn).Value = NextRiskTier Then
                matchedRow = i
                Exit For ' Exit the loop once the value is found
            End If
        Next i
        
        ' Now matchedRow contains the row number where formattedRiskPercentage was found
        Dim NextTierTradeStartingBalance As String
        If matchedRow > 0 Then
            NextTierTradeStartingBalance = wsTrades.Cells(matchedRow, startingBalanceColumn).Value
        Else
            Exit Sub
            MsgBox "ERROR", vbCritical
        End If
        
        ' get current account balance
        Dim LastTradeAccountBalance As String
        LastTradeAccountBalance = wsTrades.Cells(lastRow, startingBalanceColumn).Value
        

        If LastTradeAccountBalance > NextTierTradeStartingBalance Then
            previousRiskPercentage = NextRiskTier
        End If
        
        ' ======== DECREASE RISK
    ElseIf lossCount >= 1 Then 'And lossCount <= riskRange.Rows.count Then
        If FoundRiskCellRowRange >= 6 Then
            previousRiskPercentage = wsSetup.Range("F6").Value
        Else
            previousRiskPercentage = wsSetup.Range("F" & (FoundRiskCellRowRange + 1)).Value
        End If
    
    End If
   
JumpSpot:
   
    formattedRiskPercentage = Format(previousRiskPercentage * 100, "0.00") & "%"
    Capital_Percentage_Risk_Calculator.Value = formattedRiskPercentage
End Sub


Private Sub AfterTradeSub()
    Dim wsTrades As Worksheet
    Dim wsSetup As Worksheet
    Dim lastRow As Long
    Dim TickerColumn As Long
    Dim statusColumn As Long
    Dim profitColumn As Long
    Dim lossColumn As Long
    
    Set wsSetup = ThisWorkbook.Worksheets("Setup")
    Set wsTrades = ThisWorkbook.Worksheets("Trades")
    
    ' Define the columns
    TickerColumn = 3 ' Column C for Ticker
    statusColumn = 4 ' Column D for Status
    profitColumn = 13 ' Column M for Profit
    lossColumn = 14 ' Column N for Loss
    
    ' Set the relevant worksheet
    Select_Trade_Listbox.Clear
    lastRow = wsTrades.Cells(wsTrades.Rows.Count, TickerColumn).End(xlUp).Row
    
    If wsSetup.Range("G15").Value = "y" Then
        AT_Demo_Trading_Active.Value = "Warning! Demo Tracking Enabled"
    Else
        AT_Demo_Trading_Active.Value = ""
    End If

    After_Trade_Close_Time = Format(Now, "dd/mm/yyyy hh:mm")

    For i = 5 To lastRow ' Assuming row 1 is header
        Dim ticker As String
        Dim status As String
        Dim profit As Variant
        Dim loss As Variant
        
        ticker = wsTrades.Cells(i, TickerColumn).Value
        status = wsTrades.Cells(i, statusColumn).Value
        profit = wsTrades.Cells(i, profitColumn).Value
        loss = wsTrades.Cells(i, lossColumn).Value
        
        ' Check if there's a value in both Status and Ticker columns
        If status <> "" And ticker <> "" Then
            ' Check if there's no value in Profit and Loss columns
            If IsEmpty(profit) And IsEmpty(loss) Then
                ' Add the item to the listbox
                Select_Trade_Listbox.AddItem ticker & " | " & wsTrades.Cells(i, 11).Value ' Assuming Column K is Risk-Reward Ratio
            End If
        End If
    Next i
    
    picturesFolderPath = wsSetup.Range("J9").Value
    Set pictureFiles = New Collection
    LoadPicturesFromFolder
    SetMostRecentPicture
    UpdateImageAfterTrade
        
End Sub

Private Sub LoadPicturesFromFolder()
    Dim fs As Object
    Dim folder As Object
    Dim file As Object
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set folder = fs.GetFolder(picturesFolderPath)
    
    For Each file In folder.Files
        If LCase(fs.GetExtensionName(file.Path)) = "jpg" Then ' Only consider JPG files
            pictureFiles.Add file.Path
        End If
    Next file
End Sub

Private Sub SetMostRecentPicture()
    Dim mostRecentDate As Date
    Dim mostRecentIndex As Integer
    Dim i As Integer
    
    mostRecentDate = DateSerial(1900, 1, 1)
    
    For i = 1 To pictureFiles.Count
        If FileDateTime(pictureFiles.Item(i)) > mostRecentDate Then
            mostRecentDate = FileDateTime(pictureFiles.Item(i))
            mostRecentIndex = i
        End If
    Next i
    
    currentPictureIndex = mostRecentIndex
End Sub

Private Sub UpdateImage()
    If currentPictureIndex > 0 And currentPictureIndex <= pictureFiles.Count Then
        On Error Resume Next
        Dim imagePath As String
        imagePath = pictureFiles.Item(currentPictureIndex)
        Me.ImagePreview.Picture = LoadPicture(imagePath)
        On Error GoTo 0
        Dim screenshotDate As Date
        Me.BE_Trade_Screenshot_Desc.Caption = " " & Format(FileDateTime(imagePath), "dd/mm/yyyy hh:mm:ss")
    End If
End Sub


Private Sub UpdateImageAfterTrade()
    If currentPictureIndex > 0 And currentPictureIndex <= pictureFiles.Count Then
        On Error Resume Next
        Dim imagePath As String
        imagePath = pictureFiles.Item(currentPictureIndex)
        Me.ImagePreviewAfterTrade.Picture = LoadPicture(imagePath)
        On Error GoTo 0
        Me.AF_Trade_Screenshot_Desc.Caption = " " & Format(FileDateTime(imagePath), "dd/mm/yyyy hh:mm:ss")
    End If
End Sub

Private Sub Change_Picture_Spinbutton_SpinUp()
    If currentPictureIndex > 1 Then
        currentPictureIndex = currentPictureIndex - 1
        UpdateImage
    End If
End Sub

Private Sub Change_Picture_Spinbutton_SpinDown()
    If currentPictureIndex < pictureFiles.Count Then
        currentPictureIndex = currentPictureIndex + 1
        UpdateImage
    End If
End Sub

Private Sub Change_Picture_Spinbutton2_SpinUp()
    If currentPictureIndex > 1 Then
        currentPictureIndex = currentPictureIndex - 1
        UpdateImageAfterTrade
    End If
End Sub

Private Sub Change_Picture_Spinbutton2_SpinDown()
    If currentPictureIndex < pictureFiles.Count Then
        currentPictureIndex = currentPictureIndex + 1
        UpdateImageAfterTrade
    End If
End Sub

Private Sub Why_I_Dont_Like_The_Trade_Add_CheckBox_Click()

    Dim reason As String
    reason = Why_I_dont_Like_The_Trade.Value
    
    If reason <> "" Then
        Why_I_Dont_Like_The_Trade_ListBox.AddItem reason
        Why_I_dont_Like_The_Trade = ""
    End If
    
    Why_I_Dont_Like_The_Trade_Add_CheckBox.Value = False

End Sub

Private Sub Why_I_Like_The_Trade_Add_CheckBox_Click()
    
    Dim reason As String
    reason = Why_I_Like_The_Trade.Value
    
    If reason <> "" Then
        Why_I_Like_The_Trade_ListBox.AddItem reason
        Why_I_Like_The_Trade.Value = ""
    End If
    
    Why_I_Like_The_Trade_Add_CheckBox.Value = False
End Sub


