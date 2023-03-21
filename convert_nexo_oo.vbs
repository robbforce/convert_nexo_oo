' Convert Nexo (OpenOffice)
' Purpose: Convert nexo.io transaction history to the accointing.com template.
' Author: Nathaniel Roark
'   Based on convert_nexo v1.3 by droblesa 03/18/21, https://community.accointing.com/t/nexo-integration/95/62
'----------------------------------------------------------------------------------
' Version History:
'   2022-12-27  v0.2  Adjusted columns to latest format and added some asset conversions.
'   2021-11-14  v0.1  First build
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

Dim oSM       ' OpenOffice / LibreOffice Service Manager
Dim oDesk     ' OpenOffice / LibreOffice Desktop
Dim oAccointingDoc
Dim oAccointingSheet
Dim oExchangeDoc
Dim oExchangeSheet
Dim sFileName ' Filename without extension
Dim sInAsset
Dim sAsset2

' Open LibreOffice and a desktop instance.
Set oSM = WScript.CreateObject("com.sun.star.ServiceManager")
Set oDesk = oSM.createInstance("com.sun.star.frame.Desktop")

iDoAll = msgbox("Do you want to create an Exchanges file for import into Accointing.com?", VBYesNo, sMsgTitle)

Dim aProps() 'save properties (empty)
Set oAccointingDoc = oDesk.loadComponentFromURL("private:factory/scalc", "_blank", 0, aProps)
oAccointingDoc.CurrentController.Frame.ContainerWindow.Visible = True

' Instantiate a workbook for accointing, that will not have exchange transactions.
Set oAccointingSheet = oAccointingDoc.CurrentController.ActiveSheet
MakeHeaders(oAccointingSheet)

' If the user also wants exchange transactions, instantiate a second workbook and set references.
if iDoAll = 6 then
  Set oExchangeDoc = oDesk.loadComponentFromURL("private:factory/scalc", "_blank", 0, aProps)
  oExchangeDoc.CurrentController.Frame.ContainerWindow.Visible = True
  Set oExchangeSheet = oExchangeDoc.CurrentController.ActiveSheet
  MakeHeaders(oExchangeSheet)
end if

' Open the transaction file.
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oSrcFile = oFSO.OpenTextFile(WScript.Arguments(0), 1)
oSrcFolder = oFSO.GetFile(WScript.Arguments(0)).ParentFolder

'----------------------------------------------------------------------------------
' Read data from source file and process for output
' Data Structure:
' 0-Transaction, 1-Type, 2-Currency, 3-Amount, 4-USD Equivalent, 5-Details, 6-Outstanding Loan, 7-Date/Time

' 0-Transaction, 1-Type, 2-Input Currency, 3-Input Amount, 4-Output Currency, 5-Output Amount, 6-USD Equivalent, 7-Details, 8-Outstanding Loan, 9-Date / Time
'----------------------------------------------------------------------------------
rowcountSS1 = 0
rowcountSS2 = 0

do until oSrcFile.AtEndOfStream
  NextLine = oSrcFile.ReadLine
  rowcountSS1 = rowcountSS1 + 1
  aColumnText = split(NextLine, ",")
  'sValue = ""

  ' Column counts start at zero. Ignore the source file header line.
  if aColumnText(0) = "Transaction" then
    rowcountSS1 = 0
  else
    ' Write transaction data to the workbook.
    newdate = dateadd("h", -2, cdate(aColumnText(8)))  'Makes all transactions UTC
    oAccointingSheet.getCellByPosition(1, rowcountSS1).String = newdate  'date
    
    ' Convert to a recognized asset.
    sInAsset = ConvertAsset(aColumnText(2))
    sOutAsset = ConvertAsset(aColumnText(4))

    select case aColumnText(1)
      case "Deposit","DepositToExchange"
        oAccointingSheet.getCellByPosition(0, rowcountSS1).String = "deposit"       'transactionType
        oAccointingSheet.getCellByPosition(3, rowcountSS1).String = sInAsset        'inBuyAsset
        oAccointingSheet.getCellByPosition(2, rowcountSS1).Value = aColumnText(3)   'inBuyAmount
        oAccointingSheet.getCellByPosition(9, rowcountSS1).String = aColumnText(0)  'operationId
      
      case "Interest","Dividend","FixedTermInterest"
        if aColumnText(2) <> "USD" then
          oAccointingSheet.getCellByPosition(0, rowcountSS1).String = "deposit"       'transactionType
          oAccointingSheet.getCellByPosition(8, rowcountSS1).String = "income"        'classification
          oAccointingSheet.getCellByPosition(3, rowcountSS1).String = sInAsset        'inBuyAsset
          oAccointingSheet.getCellByPosition(2, rowcountSS1).Value = aColumnText(3)   'inBuyAmount
          oAccointingSheet.getCellByPosition(9, rowcountSS1).String = aColumnText(0)  'operationId
        else
          oAccointingSheet.getCellByPosition(1, rowcountSS1).String = ""
          rowcountSS1 = rowcountSS1 - 1
        end if
        
      case "Withdrawal"
        oAccointingSheet.getCellByPosition(0, rowcountSS1).String = "withdraw"      'transactionType
        oAccointingSheet.getCellByPosition(5, rowcountSS1).String = sOutAsset       'outSellAsset
        oAccointingSheet.getCellByPosition(4, rowcountSS1).Value = aColumnText(5)   'outSellAmount
        oAccointingSheet.getCellByPosition(9, rowcountSS1).String = aColumnText(0)  'operationId
        
      case "TransferIn","TransferOut","WithdrawalCredit","Repayment","InterestAdditional","Administrator","LockingTermDeposit","UnlockingTermDeposit"
        rowcountSS1 = rowcountSS1 - 1
      
      ' Reverse the transactions for Nexo.
      case "Exchange"
        if iDoAll = 6 then
          if rowcountSS2 = 0 then
            rowcountSS2 = rowcountSS2 + 1
          end if
          oExchangeSheet.getCellByPosition(0, rowcountSS2).String = "order"         'transactionType
          oExchangeSheet.getCellByPosition(1, rowcountSS2).String = newdate         'date
          oExchangeSheet.getCellByPosition(3, rowcountSS2).String = sOutAsset       'inBuyAsset
          oExchangeSheet.getCellByPosition(2, rowcountSS2).Value = aColumnText(5)   'inBuyAmount
          oExchangeSheet.getCellByPosition(5, rowcountSS2).String = sInAsset        'outSellAsset
          oExchangeSheet.getCellByPosition(4, rowcountSS2).Value = aColumnText(3)   'outSellAmount
          oExchangeSheet.getCellByPosition(9, rowcountSS2).String = aColumnText(0)  'operationId
          rowcountSS2 = rowcountSS2 + 1
        end if
        rowcountSS1 = rowcountSS1 - 1
        
      case "ExchangeDepositedOn"
        if iDoAll = 6 then
          if rowcountSS2 = 0 then
            rowcountSS2 = rowcountSS2 + 1
          end if
          oExchangeSheet.getCellByPosition(0, rowcountSS2).String = "order"         'transactionType
          oExchangeSheet.getCellByPosition(1, rowcountSS2).String = newdate         'date
          oExchangeSheet.getCellByPosition(3, rowcountSS2).String = "USD"           'inBuyAsset
          oExchangeSheet.getCellByPosition(2, rowcountSS2).Value = aColumnText(3)   'inBuyAmount
          oExchangeSheet.getCellByPosition(5, rowcountSS2).String = "USDX"          'outSellAsset
          oExchangeSheet.getCellByPosition(4, rowcountSS2).Value = aColumnText(5)   'outSellAmount
          oExchangeSheet.getCellByPosition(9, rowcountSS2).String = aColumnText(0)  'operationId
          rowcountSS2 = rowcountSS2 + 1
        end if
        rowcountSS1 = rowcountSS1 - 1

      case "Exchange Cashback"
        if iDoAll = 6 then
          if rowcountSS2 = 0 then
            rowcountSS2 = rowcountSS2 + 1
          end if
          oExchangeSheet.getCellByPosition(0, rowcountSS2).String = "deposit"       'transactionType
          oExchangeSheet.getCellByPosition(8, rowcountSS2).String = "bounty"        'classification
          oExchangeSheet.getCellByPosition(1, rowcountSS2).String = newdate         'date
          oExchangeSheet.getCellByPosition(3, rowcountSS2).String = sInAsset        'inBuyAsset
          oExchangeSheet.getCellByPosition(2, rowcountSS2).Value = aColumnText(3)   'inBuyAmount
          oExchangeSheet.getCellByPosition(9, rowcountSS2).String = aColumnText(0)  'operationId
          rowcountSS2 = rowcountSS2 + 1
        end if
        rowcountSS1 = rowcountSS1 - 1

      case "Liquidation"
        oAccointingSheet.getCellByPosition(0, rowcountSS1).String = "withdraw"      'transactionType
        oAccointingSheet.getCellByPosition(8, rowcountSS1).String = "payment"       'classification
        oAccointingSheet.getCellByPosition(5, rowcountSS1).String = sOutAsset       'outSellAsset
        oAccointingSheet.getCellByPosition(4, rowcountSS1).Value = aColumnText(5)   'outSellAmount
        oAccointingSheet.getCellByPosition(9, rowcountSS1).String = aColumnText(0)  'operationId
    end select
  end if
loop

' Format columns to autowidth, save output and close the workbook.
call SaveFiles(oSrcFolder & "\" & "nexo_transactions_converted.xlsx", oAccointingDoc, oAccointingSheet)
oAccointingDoc.close(True)

if iDoAll = 6 then
  ' Save and close the workbook.
  call SaveFiles(oSrcFolder & "\" & "nexo_exchange_converted.xlsx", oExchangeDoc, oExchangeSheet)
  oExchangeDoc.close(True)
end if

' Clear objects from memory.
Set oAccointingSheet = Nothing
Set oAccointingDoc = Nothing
Set oExchangeSheet = Nothing
Set oExchangeDoc = Nothing

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
  oSheet.getCellByPosition(0, 0).String = "transactionType"
  oSheet.getCellByPosition(1, 0).String = "date"
  oSheet.getCellByPosition(2, 0).String = "inBuyAmount"
  oSheet.getCellByPosition(3, 0).String = "inBuyAsset"
  oSheet.getCellByPosition(4, 0).String = "outSellAmount"
  oSheet.getCellByPosition(5, 0).String = "outSellAsset"
  oSheet.getCellByPosition(6, 0).String = "feeAmount (optional)"
  oSheet.getCellByPosition(7, 0).String = "feeAsset (optional)"
  oSheet.getCellByPosition(8, 0).String = "classification (optional)"
  oSheet.getCellByPosition(9, 0).String = "operationId (optional)"
  oSheet.getCellByPosition(10, 0).String = "comments (optional)"
End Sub

Function ConvertAsset(sAsset)
  select case sAsset
    case "BNBN","NEXOBNB"
      ConvertAsset = "BNB"
    case "NEXOBEP2","NEXONEXO"
      ConvertAsset = "NEXO"
    case "USDTERC"
      ConvertAsset = "USDT"
    case "UST"
      ConvertAsset = "USTC"
    case "LUNA2"
      ConvertAsset = "LUNA"
    case else
      ConvertAsset = sAsset
  end select
End Function

Function ExchangePair(dump, pair_value)
  dumptext = split(dump, "/")
  sCoin = dumptext(pair_value)
  select case sCoin
    case "NEXONEXO"
      ExchangePair = "NEXO"
    case "USDTERC"
      ExchangePair = "USDT"
    case "NEXOBNB"
      ExchangePair = "BNB"
    case else
      ExchangePair = sCoin
  end select
End Function

Function ExchangeAmount(dump, amt_value)
  dumptext = split(dump, " ")
  tradeamt = dumptext(amt_value)
  ExchangeAmount = tradeamt
End Function
