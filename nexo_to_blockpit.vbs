' Convert Nexo to Blockpit (OpenOffice)
' Purpose: Convert nexo.io transaction history to the accointing.com template.
' Author: Nathaniel Roark
' Based on convert_nexo v1.3 by droblesa 03/18/21, https://community.accointing.com/t/nexo-integration/95/62
'----------------------------------------------------------------------------------
' Version History:
' 2023-12-08  v0.1  First build
'***********************************************************************************

' Make sure we have command-line arguments, which should contain our input file.
if WScript.Arguments.Count = 0 then
  WScript.Echo "Missing parameters"
  WScript.quit
end if

' Declare global variables
Dim sMsgTitle
sMsgTitle = "Nexo CSV File Processor for Accointing.com"

Dim oFSO      ' Filesystem object
Dim oSrcFile  ' Transaction file
Dim oSrcFolder

Dim oSM       ' OpenOffice / LibreOffice Service Manager
Dim oDesk     ' OpenOffice / LibreOffice Desktop
Dim oWindow
Dim oSheet
Dim sFileName ' Filename without extension
Dim rowCount
Dim dteUTC
Dim sInAsset
Dim sOutAsset

' Open LibreOffice and a desktop instance.
Set oSM = WScript.CreateObject("com.sun.star.ServiceManager")
Set oDesk = oSM.createInstance("com.sun.star.frame.Desktop")

Dim aProps() 'save properties (empty)
Set oWindow = oDesk.loadComponentFromURL("private:factory/scalc", "_blank", 0, aProps)
oWindow.CurrentController.Frame.ContainerWindow.Visible = True

' Create an object for the sheet and add column headers.
Set oSheet = oWindow.CurrentController.ActiveSheet
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
  aColumnText = split(NextLine, ",")

  ' Column counts start at zero. Ignore the source file header line.
  if aColumnText(0) = "Transaction" then
    rowCount = 0
  else
    rowCount = rowCount + 1

    ' Write transaction data to the workbook.
    dteUTC = dateadd("h", -1, cdate(aColumnText(8)))  'Makes all transactions UTC
    oSheet.getCellByPosition(0, rowCount).String = timeStamp(dteUTC)   'Timestamp
    
    ' Convert to a recognized asset.
    sInAsset = ConvertAsset(aColumnText(2))
    sOutAsset = ConvertAsset(aColumnText(4))

    select case aColumnText(1)
      ' Ignore these transactions.
      case "Administrator","Assimilation","Transfer In","Transfer Out","Locking Term Deposit","Unlocking Term Deposit"
        oSheet.getCellByPosition(0, rowCount).String = ""
        rowCount = rowCount - 1

      case "Top up Crypto","Dividend","Exchange Cashback","Fixed Term Interest","Interest"
        ' Skip empty or negative interest and credit deposits (handled with "Loan Withdrawal").
        if CSng(aColumnText(3)) <= 0 or (aColumnText(1) = "Top up Crypto" and InStr(1,aColumnText(7),"Credit",1)) then
          oSheet.getCellByPosition(0, rowCount).String = ""
          rowCount = rowCount - 1
        else
          oSheet.getCellByPosition(1, rowCount).String = "Nexo"           'Integration Name
          select case aColumnText(1)
            case "Top up Crypto"
              if InStr(1,aColumnText(7),"Airdrop",1) then
                oSheet.getCellByPosition(2, rowCount).String = "Airdrop"          'Label
              else
                oSheet.getCellByPosition(2, rowCount).String = "Deposit"          'Label
              end if
            case "Dividend"
              oSheet.getCellByPosition(2, rowCount).String = "Income"             'Label
            case "Exchange Cashback"
              oSheet.getCellByPosition(2, rowCount).String = "Gift Received"      'Label
            case "Interest","Fixed Term Interest"
              oSheet.getCellByPosition(2, rowCount).String = "Interest"           'Label
          end select
          oSheet.getCellByPosition(5, rowCount).String = sInAsset         'Incoming Asset
          oSheet.getCellByPosition(6, rowCount).Value = aColumnText(3)    'Incoming Amount
          oSheet.getCellByPosition(10, rowCount).String = aColumnText(0)  'Trx. ID (optional)
        end if
        
      case "Withdrawal","Liquidation"
        oSheet.getCellByPosition(1, rowCount).String = "Nexo"           'Integration Name
        select case aColumnText(1)
          case "Withdrawal"
            oSheet.getCellByPosition(2, rowCount).String = "Withdrawal"         'Label
          case "Liquidation"
            oSheet.getCellByPosition(2, rowCount).String = "Payment"            'Label
        end select
        oSheet.getCellByPosition(3, rowCount).String = sOutAsset            'Outgoing Asset
        oSheet.getCellByPosition(4, rowCount).Value = aColumnText(5)        'Outgoing Amount
        oSheet.getCellByPosition(10, rowCount).String = aColumnText(0)      'Trx. ID (optional)
      
      ' Reverse the asset and amount columns for "Exchange" transactions.
      case "Exchange"
        oSheet.getCellByPosition(1, rowCount).String = "Nexo"               'Integration Name
        oSheet.getCellByPosition(2, rowCount).String = "Trade"              'Label
        oSheet.getCellByPosition(3, rowCount).String = sInAsset             'Outgoing Asset
        ' Nexo records a negative for the outgoing exchange asset, use Abs to drop the sign.
        oSheet.getCellByPosition(4, rowCount).Value = Abs(aColumnText(3))   'Outgoing Amount
        oSheet.getCellByPosition(5, rowCount).String = sOutAsset            'Incoming Asset
        oSheet.getCellByPosition(6, rowCount).Value = aColumnText(5)        'Incoming Amount
        oSheet.getCellByPosition(10, rowCount).String = aColumnText(0)      'Trx. ID (optional)

      case "Loan Withdrawal"
        oSheet.getCellByPosition(1, rowCount).String = "Nexo"               'Integration Name
        oSheet.getCellByPosition(2, rowCount).String = "Non-Taxable In"     'Label
        oSheet.getCellByPosition(5, rowCount).String = sOutAsset            'Incoming Asset
        oSheet.getCellByPosition(6, rowCount).Value = aColumnText(5)        'Incoming Amount
        oSheet.getCellByPosition(9, rowCount).String = "Crypto borrowing"   'Comment (optional)
        oSheet.getCellByPosition(10, rowCount).String = aColumnText(0)      'Trx. ID (optional)

      case "Manual Sell Order"
        oSheet.getCellByPosition(1, rowCount).String = "Nexo"               'Integration Name
        oSheet.getCellByPosition(2, rowCount).String = "Trade"              'Label
        oSheet.getCellByPosition(3, rowCount).String = sInAsset             'Outgoing Asset
        ' Nexo records a negative for the outgoing asset, use Abs to drop the sign.
        oSheet.getCellByPosition(4, rowCount).Value = Abs(aColumnText(3))   'Outgoing Amount
        oSheet.getCellByPosition(5, rowCount).String = "USD"                'Incoming Asset
        oSheet.getCellByPosition(6, rowCount).Value = aColumnText(6)        'Incoming Amount
        oSheet.getCellByPosition(10, rowCount).String = aColumnText(0)      'Trx. ID (optional)

      case "Manual Repayment","Interest Additional"
        oSheet.getCellByPosition(1, rowCount).String = "Nexo"               'Integration Name
        select case aColumnText(1)
          case "Manual Repayment"
            oSheet.getCellByPosition(2, rowCount).String = "Non-Taxable Out"    'Label
            oSheet.getCellByPosition(9, rowCount).String = "Crypto repayment"   'Comment (optional)
          case "Interest Additional"
            oSheet.getCellByPosition(2, rowCount).String = "Fee"                'Label
            oSheet.getCellByPosition(9, rowCount).String = "Borrowing fee"      'Comment (optional)
        end select
        oSheet.getCellByPosition(3, rowCount).String = sInAsset             'Outgoing Asset
        oSheet.getCellByPosition(4, rowCount).Value = Abs(aColumnText(3))   'Outgoing Amount
        oSheet.getCellByPosition(10, rowCount).String = aColumnText(0)      'Trx. ID (optional)
    end select
  end if
loop

' Format columns to autowidth, save output and close the workbook.
call SaveFiles(oSrcFolder & "\" & "blockpit_nexo_transactions.xlsx", oWindow, oSheet)
oWindow.close(True)

' Clear objects from memory.
Set oSheet = Nothing
Set oWindow = Nothing

' Close LibreOffice.
oDesk.terminate
Set oDesk = nothing
Set oSM = nothing

Call MsgBox("Conversion completed.", vbOKOnly, sMsgTitle)

WScript.quit
' End of script.


'***********************************************************************************
' Helper functions and routines.

Sub SaveFiles(sFilePath, oDoc, oSheet)
  Dim aProps(0) 'save properties (empty)
  Dim oProp0
  Dim sSaveUrl

  ' Set overwrite option and write into the properties array.
  Set oProp0    = oSM.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
  oProp0.Name   = "Overwrite"
  oProp0.Value  = True
  Set aProps(0) = oProp0
  
  'oSheet.columns.autofit
  'oSheet.Application.DisplayAlerts = False
  sSaveUrl = ConvertToURL(sFilePath)
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
  oSheet.getCellByPosition(0, 0).String = "Date (UTC)"
  oSheet.getCellByPosition(1, 0).String = "Integration Name"
  oSheet.getCellByPosition(2, 0).String = "Label"
  oSheet.getCellByPosition(3, 0).String = "Outgoing Asset"
  oSheet.getCellByPosition(4, 0).String = "Outgoing Amount"
  oSheet.getCellByPosition(5, 0).String = "Incoming Asset"
  oSheet.getCellByPosition(6, 0).String = "Incoming Amount"
  oSheet.getCellByPosition(7, 0).String = "Fee Asset (optional)"
  oSheet.getCellByPosition(8, 0).String = "Fee Amount (optional)"
  oSheet.getCellByPosition(9, 0).String = "Comment (optional)"
  oSheet.getCellByPosition(10, 0).String = "Trx. ID (optional)"
End Sub

Function timeStamp(myDate) 
  timeStamp = Right("0" & Day(myDate),2) & "." & Right("0" & Month(myDate),2) & "." & Year(myDate) & " " & _  
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
    case else
      ConvertAsset = sAsset
  end select
End Function
