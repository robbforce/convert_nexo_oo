' Convert Nexo to coinpanda (OpenOffice)
' Purpose: Convert nexo.io transaction history to the coinpanda.io template.
' Author: Nathaniel Roark
' Based on convert_nexo v1.3 by droblesa 03/18/21, https://community.accointing.com/t/nexo-integration/95/62
'----------------------------------------------------------------------------------
' Version History:
' 2023-12-11  v0.1  First build
'***********************************************************************************

' Make sure we have command-line arguments, which should contain our input file.
if WScript.Arguments.Count = 0 then
  WScript.Echo "Missing parameters"
  WScript.quit
end if

' Declare global variables
Dim sMsgTitle
sMsgTitle = "Nexo CSV File Processor for coinpanda.io"

Dim oFSO      ' Filesystem object
Dim oSrcFile  ' Transaction file
Dim oSrcFolder

Dim oSM       ' OpenOffice / LibreOffice Service Manager
Dim oDesk     ' OpenOffice / LibreOffice Desktop
Dim oDoc
Dim oSheets
Dim oSheet
Dim sFileName ' Filename without extension
sFileName = "coinpanda_nexo_transactions"

Dim iBatchCount
iBatchCount = 100
Dim iRowTotal
Dim iRowCount
Dim iSheetCount
Dim sSheetName
Dim aRangeData

Dim dteUTC
Dim sInAsset
Dim sOutAsset

' Open LibreOffice and a desktop instance.
Set oSM = WScript.CreateObject("com.sun.star.ServiceManager")
Set oDesk = oSM.createInstance("com.sun.star.frame.Desktop")

Dim aProps() 'save properties (empty)
Set oDoc = oDesk.loadComponentFromURL("private:factory/scalc", "_blank", 0, aProps)
oDoc.CurrentController.Frame.ContainerWindow.Visible = True

' Create an object for the sheet and add column headers.
Set oSheets = oDoc.getSheets()  
Set oSheet = oDoc.CurrentController.ActiveSheet
Call MakeHeaders(oSheet)

' Open the transaction file.
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oSrcFile = oFSO.OpenTextFile(WScript.Arguments(0), 1)
oSrcFolder = oFSO.GetFile(WScript.Arguments(0)).ParentFolder

'----------------------------------------------------------------------------------
' Read data from source file and process for output
' Data Structure:
' 0-Transaction, 1-Type, 2-Input Currency, 3-Input Amount, 4-Output Currency, 5-Output Amount, 6-USD Equivalent, 7-Details, 8-Date / Time
'----------------------------------------------------------------------------------
do until oSrcFile.AtEndOfStream
  NextLine = oSrcFile.ReadLine
  'aColumnText = split(NextLine, ",")
  aColumnText = TokenizeCsvFast(NextLine)

  ' Column counts start at zero. Ignore the source file header line.
  if aColumnText(0) = "Transaction" then
    iRowTotal = 0
  else
    iRowTotal = iRowTotal + 1
    iRowCount = iRowCount + 1

    ' Write transaction data to the workbook.
    dteUTC = dateadd("h", -1, cdate(aColumnText(8)))  'Makes all transactions UTC
    oSheet.getCellByPosition(0, iRowTotal).String = timeStamp(dteUTC)   'Timestamp
    
    ' Convert to a recognized asset.
    sInAsset = ConvertAsset(aColumnText(2))
    sOutAsset = ConvertAsset(aColumnText(4))

    select case aColumnText(1)
      ' Ignore these transactions.
      case "Administrator","Assimilation","Transfer In","Transfer Out","Locking Term Deposit","Unlocking Term Deposit"
        oSheet.getCellByPosition(0, iRowTotal).String = ""
        iRowTotal = iRowTotal - 1
        iRowCount = iRowCount - 1

      case "Top up Crypto","Dividend","Exchange Cashback","Fixed Term Interest","Interest"
        ' Skip empty or negative interest and credit deposits (handled with "Loan Withdrawal").
        if CSng(aColumnText(3)) <= 0 or (aColumnText(1) = "Top up Crypto" and InStr(1,aColumnText(7),"Credit",1)) then
          oSheet.getCellByPosition(0, iRowTotal).String = ""
          iRowTotal = iRowTotal - 1
          iRowCount = iRowCount - 1
        else
          oSheet.getCellByPosition(1, iRowTotal).String = "Receive"          'Type
          oSheet.getCellByPosition(4, iRowTotal).Value = aColumnText(3)      'Received Amount
          oSheet.getCellByPosition(5, iRowTotal).String = sInAsset           'Received Currency
          oSheet.getCellByPosition(8, iRowTotal).Value = aColumnText(6)      'Net Worth Amount
          oSheet.getCellByPosition(9, iRowTotal).String = "USD"              'Net Worth Currency
          select case aColumnText(1)
            case "Top up Crypto"
              if InStr(1,aColumnText(7),"Airdrop",1) then
                oSheet.getCellByPosition(10, iRowTotal).String = "Airdrop"     'Label
              end if
            case "Dividend"
              oSheet.getCellByPosition(10, iRowTotal).String = "Income"        'Label
            case "Exchange Cashback"
              oSheet.getCellByPosition(10, iRowTotal).String = "Gift"          'Label
            case "Interest","Fixed Term Interest"
              oSheet.getCellByPosition(10, iRowTotal).String = "Interest"      'Label
          end select
          oSheet.getCellByPosition(11, iRowTotal).String = aColumnText(7)    'Description
        end if
        
      case "Withdrawal","Liquidation"
        oSheet.getCellByPosition(1, iRowTotal).String = "Sent"               'Type
        oSheet.getCellByPosition(2, iRowTotal).Value = aColumnText(5)        'Sent Amount
        oSheet.getCellByPosition(3, iRowTotal).String = sOutAsset            'Sent Currency
        oSheet.getCellByPosition(8, iRowTotal).Value = aColumnText(6)        'Net Worth Amount
        oSheet.getCellByPosition(9, iRowTotal).String = "USD"                'Net Worth Currency
        select case aColumnText(1)
          case "Liquidation"
            oSheet.getCellByPosition(10, iRowTotal).String = "Expense"       'Label
        end select
        oSheet.getCellByPosition(11, iRowTotal).String = aColumnText(7)      'Description
      
      ' Reverse the asset and amount columns for "Exchange" transactions.
      case "Exchange"
        oSheet.getCellByPosition(1, iRowTotal).String = "Trade"              'Type
        ' Nexo records a negative for the outgoing exchange asset, use Abs to drop the sign.
        oSheet.getCellByPosition(2, iRowTotal).Value = Abs(aColumnText(3))   'Sent Amount
        oSheet.getCellByPosition(3, iRowTotal).String = sInAsset             'Sent Currency
        oSheet.getCellByPosition(4, iRowTotal).Value = aColumnText(5)        'Received Amount
        oSheet.getCellByPosition(5, iRowTotal).String = sOutAsset            'Received Currency
        oSheet.getCellByPosition(8, iRowTotal).Value = aColumnText(6)        'Net Worth Amount
        oSheet.getCellByPosition(9, iRowTotal).String = "USD"                'Net Worth Currency
        oSheet.getCellByPosition(11, iRowTotal).String = aColumnText(7)      'Description

      ' For loans, Nexo records the received amount and asset in the output columns.
      case "Loan Withdrawal"
        oSheet.getCellByPosition(1, iRowTotal).String = "Receive"            'Type
        oSheet.getCellByPosition(4, iRowTotal).Value = aColumnText(5)        'Received Amount
        oSheet.getCellByPosition(5, iRowTotal).String = sOutAsset            'Received Currency
        oSheet.getCellByPosition(8, iRowTotal).Value = aColumnText(6)        'Net Worth Amount
        oSheet.getCellByPosition(9, iRowTotal).String = "USD"                'Net Worth Currency
        oSheet.getCellByPosition(10, iRowTotal).String = "Receive Loan"      'Label
        oSheet.getCellByPosition(11, iRowTotal).String = aColumnText(7)      'Description

      ' These sell orders are used for repayments, but the USD value isn't recorded in the output columns.
      case "Manual Sell Order"
        oSheet.getCellByPosition(1, iRowTotal).String = "Trade"              'Type
        ' Nexo records a negative for the outgoing asset, use Abs to drop the sign.
        oSheet.getCellByPosition(2, iRowTotal).Value = Abs(aColumnText(3))   'Sent Amount
        oSheet.getCellByPosition(3, iRowTotal).String = sInAsset             'Sent Currency
        oSheet.getCellByPosition(4, iRowTotal).Value = aColumnText(6)        'Received Amount
        oSheet.getCellByPosition(5, iRowTotal).String = "USD"                'Received Currency
        oSheet.getCellByPosition(8, iRowTotal).Value = aColumnText(6)        'Net Worth Amount
        oSheet.getCellByPosition(9, iRowTotal).String = "USD"                'Net Worth Currency
        oSheet.getCellByPosition(11, iRowTotal).String = aColumnText(7)      'Description

      ' For payments, Nexo records the sent amount and asset in the input columns.
      case "Manual Repayment","Interest Additional"
        oSheet.getCellByPosition(1, iRowTotal).String = "Sent"               'Type
        oSheet.getCellByPosition(2, iRowTotal).Value = Abs(aColumnText(3))   'Sent Amount
        oSheet.getCellByPosition(3, iRowTotal).String = sInAsset             'Sent Currency
        oSheet.getCellByPosition(8, iRowTotal).Value = aColumnText(6)        'Net Worth Amount
        oSheet.getCellByPosition(9, iRowTotal).String = "USD"                'Net Worth Currency
        select case aColumnText(1)
          case "Manual Repayment"
            oSheet.getCellByPosition(10, iRowTotal).String = "Repay Loan"    'Label
          case "Interest Additional"
            oSheet.getCellByPosition(10, iRowTotal).String = "Cost"          'Label
        end select
        oSheet.getCellByPosition(11, iRowTotal).String = aColumnText(7)      'Description
    end select
  end if

  ' Save seperate files out according to the set batch count.
  if iRowCount = iBatchCount or oSrcFile.AtEndOfStream Then
    ' Excel row counts are base zero so add 1 more to the row total if we hit the end of the file.
    if oSrcFile.AtEndOfStream Then iRowTotal = iRowTotal + 1

    ' Create the new sheet with headers.
    iSheetCount = iSheetCount + 1
    sSheetName = Right("0" & iSheetCount, 2)
    Call oSheets.insertNewByName(sSheetName, oSheets.getCount())
    Set oSheetBatch = oSheets.getByName(sSheetName)
    Call MakeHeaders(oSheetBatch)

    ' To copy and paste a range, the source range, containing variant array and the target selection must all match dimensionally.
    aRangeData = Array(iRowTotal - ((iBatchCount * iSheetCount) - (iBatchCount - 1)), Array(12))
    aRangeData = oSheet.getCellRangeByPosition(0, (iBatchCount * iSheetCount) - (iBatchCount - 1), 12, iRowTotal).getDataArray()
    Call oSheetBatch.getCellRangeByPosition(0, 1, 12, iRowTotal - ((iBatchCount * iSheetCount) - iBatchCount)).setDataArray(aRangeData)

    ' Save the new sheet to a unique CSV.
    Call SaveFile(oDoc, oSheetBatch, oSrcFolder & "\" & sFileName & "_" & Right("0" & iSheetCount,2), "csv", True) 

    ' Reset some of the variables for the next run.
    aRangeData = Empty
    Set oSheetBatch = Nothing
    iRowCount = 0
  end if
loop

' Save output and close the workbook.
Call SaveFile(oDoc, oSheet, oSrcFolder & "\" & sFileName, "csv", True)
oDoc.close(True)

' Clear objects from memory.
Set oSheets = Nothing
Set oSheet = Nothing
Set oDoc = Nothing

' Close LibreOffice.
oDesk.terminate
Set oDesk = Nothing
Set oSM = Nothing

Call MsgBox("Conversion completed.", vbOKOnly, sMsgTitle)

WScript.quit
' End of script.


'***********************************************************************************
' Helper functions and routines.

Sub SaveFile(oDoc, oSheet, sFilePath, sExtension, sOverwrite)
  Dim sFilterName
  Dim aProps(1)   'Save properties (empty)
  Dim sSaveUrl

  ' FilterName determines the file format.
  If sExtension = "csv" Then
    sFilterName = "Text - txt - csv (StarCalc)"
  Else
    sFilterName = "Calc Office Open XML"
    sExtension = "xlsx"
  End if

  aProps(0) = oSM.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
  aProps(0).Name = "FilterName"
  aProps(0).Value = sFilterName
  aProps(1) = oSM.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
  aProps(1).Name = "Overwrite"
  aProps(1).Value = sOverwrite

  sSaveUrl = ConvertToURL(sFilePath & "." & sExtension)

  ' Make the sheet active before saving, specifically for csv files.
  Call oDoc.CurrentController.setActiveSheet(oSheet)
  oDoc.storeToURL sSaveUrl, aProps
End Sub

Function ConvertToURL(sFileName)
  ' Convert Windows pathnames to url
  Dim sTmpFile

  If Left(sFileName, 7) = "file://" Then
    ConvertToURL = sFileName
    Exit Function
  End If

  ConvertToURL = "file:///"
  sTmpFile = oFSO.GetAbsolutePathName(sFileName)

  ' replace any "\" by "/"
  sTmpFile = Replace(sTmpFile,"\","/") 

  ' replace any "%" by "%25"
  sTmpFile = Replace(sTmpFile,"%","%25") 

  ' replace any " " by "%20"
  sTmpFile = Replace(sTmpFile," ","%20")

  ConvertToURL = ConvertToURL & sTmpFile
End Function

Sub MakeHeaders(oSheet)
  oSheet.getCellByPosition(0, 0).String = "Timestamp (UTC)"
  oSheet.getCellByPosition(1, 0).String = "Type"
  oSheet.getCellByPosition(2, 0).String = "Sent Amount"
  oSheet.getCellByPosition(3, 0).String = "Sent Currency"
  oSheet.getCellByPosition(4, 0).String = "Received Amount"
  oSheet.getCellByPosition(5, 0).String = "Received Currency"
  oSheet.getCellByPosition(6, 0).String = "Fee Amount"
  oSheet.getCellByPosition(7, 0).String = "Fee Currency"
  oSheet.getCellByPosition(8, 0).String = "Net Worth Amount"
  oSheet.getCellByPosition(9, 0).String = "Net Worth Currency"
  oSheet.getCellByPosition(10, 0).String = "Label"
  oSheet.getCellByPosition(11, 0).String = "Description"
  oSheet.getCellByPosition(12, 0).String = "TxHash"
End Sub

Function timeStamp(myDate)
  timeStamp = Year(myDate) & "-" & Right("0" & Month(myDate),2) & "-" & Right("0" & Day(myDate),2) & " " & _  
    Right("0" & Hour(myDate),2) & ":" & Right("0" & Minute(myDate),2) & ":" &  Right("0" & Second(myDate),2) 
End Function

Function ConvertAsset(sAsset)
  select case sAsset
    case "BNBN","NEXOBNB"
      ConvertAsset = "BNB"
    case "NEXOBEP2","NEXONEXO"
      ConvertAsset = "NEXO"
    case "USDTERC"
      ConvertAsset = "USDT"
    case "LUNA2"
      ConvertAsset = "LUNA"
    case "UST"
      ConvertAsset = "USTC"
    case else
      ConvertAsset = sAsset
  end select
End Function

' A fast, hard-coded method for splitting a CSV string which contains quoted sections,
' e.g. 1,2,"Comma,Separated,Values",Comma,Separated,Values will be split to: 1, 2, "Comma,Separated,Values", Comma, Separated, Values
Function TokenizeCsvFast(sourceLine)
  Dim tokens()
  ReDim tokens(0)
  Dim newToken
  Dim newTokenNumber
  newTokenNumber = 0
  Dim inQuotes
  Dim stringPosition
  Dim newCharacter
  Dim newTokenComplete

  For stringPosition = 1 To Len(sourceLine)
    newCharacter = Mid(sourceLine, stringPosition, 1)
    newTokenComplete = False

    ' Handle quotes as an explicit case.
    If newCharacter = """" Then
      inQuotes = Not inQuotes
    ElseIf newCharacter = "," Then
      If inQuotes Then
        ' if in quotes, just build up the new token.
        newToken = newToken & newCharacter
      Else
        ' Outside of quotes, a comma separates values.
        newTokenComplete = True
      End If
    ' The terminal token may not have a terminal comma.
    ElseIf stringPosition = Len(sourceLine) Then
      newToken = newToken & newCharacter
      newTokenComplete = True
    Else
      ' Build up the new token one character at a time.
      newToken = newToken & newCharacter
    End If

    If newTokenComplete Then
      ' Add the completed token to the return array.
      ReDim Preserve tokens(newTokenNumber)
      tokens(newTokenNumber) = newToken
      newTokenNumber = newTokenNumber + 1
      ' Debug.Print newToken

      ' Start a new token.
      newToken = ""
    End If
  Next

  TokenizeCsvFast = tokens
End Function
