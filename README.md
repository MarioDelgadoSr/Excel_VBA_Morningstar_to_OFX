<!-- Markdown reference: https://guides.github.com/features/mastering-markdown/ -->

# *Excel_VBA_Morningstar_to_OFX*

* This VBA program will generate an OFX file from [Morningstar](https://www.morningstar.com/)'s Portfolio export and into an [OFX formatted file](http://moneymvps.org/faq/article/8.aspx).  
* The OFX file can then be imported into [Microsoft Money Plus Sunset](https://www.microsoft.com/en-us/download/details.aspx?id=20738) to update the portfolio's stock and mutual fund prices.

With this VBA program installed in Excel, you have a reliable, free source of stock and mutual fund data to keep your Microsoft Money portfolio upto date.

## Microsoft Money Stock Price Importing Background: 

* [Obtain stock and fund quotes after July 2013](http://moneymvps.org/faq/article/651.aspx)

## Instructions

* Add [Excel_VBA_Morningstar_to_OFX.vba](https://github.com/MarioDelgadoSr/Excel_VBA_Morningstar_to_OFX/blob/master/vba/Excel_VBA_Morningstar_to_OFX.vba) to Excel.  
  It includes the main macro called *makeOFX_file*.

	* **Related**:
	
		* [How to insert and run VBA code in Excel - tutorial for beginners](https://www.ablebits.com/office-addins-blog/2013/12/06/add-run-vba-macro-excel/)
		* [Copy your macros to a Personal Macro Workbook](https://support.office.com/en-us/article/Copy-your-macros-to-a-Personal-Macro-Workbook-AA439B90-F836-4381-97F0-6E4C3F5EE566)
		
* **Edit the program** (line 53) to specify the location of the OFX file that will be generated.  By default, the program will write out to "C:\temp\quotes.ofx".

* Create a portfolio in Morningstar with 'Ticker', '$ Current Price' and 'Morningstar Rating For Funds' columns.  

	* [Video: Creating a Morningstar Portfolio](http://video.morningstar.com/us/misc/portfoliomanager/portfolio_noexisting.html)
	* ![Screen Shot of required column in custom portfolio view](https://github.com/MarioDelgadoSr/Excel_VBA_Morningstar_to_OFX/blob/master/img/portfolio.png)

* Use the Morningstar 'Export' utility (see screen shot above) to export the custom portfolio view to Excel.

* Run the *makeOFX_file* macro in Excel.  It will dynamically read the Excel-based portfolio data created by Moningstar and write out the specified file.

* Installation of Microsoft Money should assoicate .ofx file with the Microsoft Money Import Handler [mnyimprt.ext](http://moneymvps.org/faq/article/407.aspx).  
  Double clicking your file will in the file explorer will start the import handler and prompt you to start Money to continue with the import.
  
	* To automate the import, you can create a [desktop shortcut](https://answers.microsoft.com/en-us/windows/forum/windows_10-start/quick-tip-create-desktop-shortcuts-in-windows-10/d867565e-34c2-42ad-88da-ccf76a4a9820) for the quotes.ofx file.
  
* View your updated Microsoft Money Portfolio.


## Excel_VBA_Morningstar_to_OFX.vba code:

````
' Source/Documentation: https://github.com/MarioDelgadoSr/Excel_VBA_Morningstar_to_OFX

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Declare Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long

Public Function GetGUID() As String ' https://www.tek-tips.com/viewthread.cfm?qid=1509722


    Dim udtGUID As GUID

    If (CoCreateGuid(udtGUID) = 0) Then

        GetGUID = _
            String(8 - Len(Hex$(udtGUID.Data1)), "0") & Hex$(udtGUID.Data1) & _
            String(4 - Len(Hex$(udtGUID.Data2)), "0") & Hex$(udtGUID.Data2) & _
            String(4 - Len(Hex$(udtGUID.Data3)), "0") & Hex$(udtGUID.Data3) & _
            IIf((udtGUID.Data4(0) < &H10), "0", "") & Hex$(udtGUID.Data4(0)) & _
            IIf((udtGUID.Data4(1) < &H10), "0", "") & Hex$(udtGUID.Data4(1)) & _
            IIf((udtGUID.Data4(2) < &H10), "0", "") & Hex$(udtGUID.Data4(2)) & _
            IIf((udtGUID.Data4(3) < &H10), "0", "") & Hex$(udtGUID.Data4(3)) & _
            IIf((udtGUID.Data4(4) < &H10), "0", "") & Hex$(udtGUID.Data4(4)) & _
            IIf((udtGUID.Data4(5) < &H10), "0", "") & Hex$(udtGUID.Data4(5)) & _
            IIf((udtGUID.Data4(6) < &H10), "0", "") & Hex$(udtGUID.Data4(6)) & _
            IIf((udtGUID.Data4(7) < &H10), "0", "") & Hex$(udtGUID.Data4(7))
    End If

End Function


Sub makeOFX_file()

    Dim objFSO As Object
    Dim objFile As Object
    Dim rowCounter As Long

    Dim tickerColumn As Long
    Dim priceColumn As Long
    Dim mutualFundIndicatorColumn As Long
    
    Dim ofxFile As String
    Dim rangeToExport As Range
    
    Set rangeToExport = ActiveSheet.UsedRange
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    ofxFile = "C:\temp\quotes.ofx"    'Fully qualified path to the quotes.ofx file that will be written out by this macro. Example: "c:\temp\quotes.ofx"
    
    Set objFile = objFSO.CreateTextFile(ofxFile, True)          'File that will be written to via File System Object
    

    ' Determine column for 'Ticker', '$ Current Price' and 'Morningstar Rating For Funds' in the Excel file generated by Morningstar
    tickerColumn = rangeToExport.Cells.Find("Ticker").Column
    priceColumn = rangeToExport.Cells.Find("$ Current" & Chr(13) & Chr(10) & "Price").Column
    mutualFundIndicatorColumn = rangeToExport.Cells.Find("Morningstar " & Chr(13) & Chr(10) & "Rating For " & Chr(13) & Chr(10) & "Funds").Column

    'OFX File Header

    objFile.WriteLine (header("NONE"))

    objFile.WriteLine (startXML(Replace(Mid(GetGUID(), 2, 36), "-", "")))
    
    For rowCounter = 2 To rangeToExport.Rows.Count
        'To Do
        If (rangeToExport.Cells(rowCounter, tickerColumn) <> "CASH$") Then  'Filter CASH from export

            If rangeToExport.Cells(rowCounter, priceColumn) = 0 Then Exit For 'Filter any 0 priced securities	

            If rangeToExport.Cells(rowCounter, mutualFundIndicatorColumn) > 0 Then  'If mutual fund column has a rating, then it's a mutual fund
                objFile.WriteLine posmf(rangeToExport.Cells(rowCounter, tickerColumn), rangeToExport.Cells(rowCounter, priceColumn))
            Else
            
                objFile.WriteLine posstock(rangeToExport.Cells(rowCounter, tickerColumn), rangeToExport.Cells(rowCounter, priceColumn))
            End If

        End If

    Next
    
    
    objFile.WriteLine (vbCrLf & vbCrLf & "</INVPOSLIST>" & vbCrLf & vbCrLf & "</INVSTMTRS>" & vbCrLf & vbCrLf & "</INVSTMTTRNRS>" & vbCrLf & vbCrLf & "</INVSTMTMSGSRSV1>" _
                  & vbCrLf & vbCrLf & "<SECLISTMSGSRSV1>" & vbCrLf & vbCrLf & "<SECLIST>")



    For rowCounter = 2 To rangeToExport.Rows.Count
        If (rangeToExport.Cells(rowCounter, tickerColumn) <> "CASH$" And rangeToExport.Cells(rowCounter, tickerColumn) <> "VMMXX") Then
            If rangeToExport.Cells(rowCounter, priceColumn) = 0 Then Exit For  'Last Line
            'If InStr(rangeToExport.Cells(rowCounter, tickerColumn), "MorningStarImport") > 0 Then Exit For
            If rangeToExport.Cells(rowCounter, mutualFundIndicatorColumn) > 0 Then
                objFile.WriteLine mfinfo(rangeToExport.Cells(rowCounter, tickerColumn), rangeToExport.Cells(rowCounter, priceColumn))
            Else
            
                objFile.WriteLine stockinfo(rangeToExport.Cells(rowCounter, tickerColumn), rangeToExport.Cells(rowCounter, priceColumn))
            End If
        End If
    Next

    
    objFile.WriteLine (vbCrLf & vbCrLf & "</SECLIST>" & vbCrLf & vbCrLf & "</SECLISTMSGSRSV1>" & vbCrLf & vbCrLf & "</OFX>")
    
    objFile.Close
    
    Set objFSO = Nothing
    
    MsgBox ("File " & ofxFile & " generated.")
    
    
End Sub

'Create GUID header
Function header(GUID)
    header = "OFXHEADER:100" & vbCrLf & vbCrLf & "DATA:OFXSGML" & vbCrLf & vbCrLf & "VERSION:102" & vbCrLf & vbCrLf & "SECURITY:NONE" & _
              vbCrLf & vbCrLf & "ENCODING:USASCII" & vbCrLf & vbCrLf & "CHARSET:1252" & _
              vbCrLf & vbCrLf & "COMPRESSION:NONE" & vbCrLf & vbCrLf & "OLDFILEUID:NONE" & vbCrLf & vbCrLf & "NEWFILEUID:" & GUID & vbCrLf
End Function

'Start of the XML
Function startXML(GUID)
    rtnString = vbCrLf & vbCrLf & "<OFX>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<SIGNONMSGSRSV1>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<SONRS>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<STATUS>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<CODE>0</CODE>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<SEVERITY>INFO</SEVERITY>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<MESSAGE>Successful Sign On</MESSAGE>"
    rtnString = rtnString & vbCrLf & vbCrLf & "</STATUS>"
    ' Date format: YYYYMMDDHHMMSS
    rtnString = rtnString & vbCrLf & vbCrLf & "<DTSERVER>" & Year(Date) & zeroFormat(Month(Date)) & zeroFormat(Day(Date)) & _
                                                     zeroFormat(Hour(Time)) & zeroFormat(Minute(Time)) & zeroFormat(Second(Time)) & _
                                               "</DTSERVER>"
                                                     

    rtnString = rtnString & vbCrLf & vbCrLf & "<LANGUAGE>ENG</LANGUAGE>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<DTPROFUP>20010918083000</DTPROFUP>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<FI>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<ORG>broker.com</ORG>"
    rtnString = rtnString & vbCrLf & vbCrLf & "</FI>"
    rtnString = rtnString & vbCrLf & vbCrLf & "</SONRS>"
    rtnString = rtnString & vbCrLf & vbCrLf & "</SIGNONMSGSRSV1>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<INVSTMTMSGSRSV1>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<INVSTMTTRNRS>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<TRNUID>" & GUID & "</TRNUID>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<STATUS>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<CODE>0</CODE>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<SEVERITY>INFO</SEVERITY>"
    rtnString = rtnString & vbCrLf & vbCrLf & "</STATUS>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<CLTCOOKIE>4</CLTCOOKIE>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<INVSTMTRS>"
    tommorow = DateAdd("d", 1, Now())
    rtnString = rtnString & vbCrLf & vbCrLf & "<DTASOF>" & Year(tommorow) & zeroFormat(Month(tommorow)) & zeroFormat(Day(tommorow)) & "</DTASOF>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<CURDEF>USD</CURDEF>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<INVACCTFROM>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<BROKERID>dummybroker.com</BROKERID>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<ACCTID>0123456789</ACCTID>"
    rtnString = rtnString & vbCrLf & vbCrLf & "</INVACCTFROM>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<INVPOSLIST>"
    
    startXML = rtnString
    
End Function

Function posstock(strSecurity, floatPrice)
    
    rtnString = vbCrLf & vbCrLf & "<POSSTOCK>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<INVPOS>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<SECID>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<UNIQUEID>" & strSecurity & "</UNIQUEID>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<UNIQUEIDTYPE>TICKER</UNIQUEIDTYPE>"
    rtnString = rtnString & vbCrLf & vbCrLf & "</SECID>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<HELDINACCT>CASH</HELDINACCT>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<POSTYPE>LONG</POSTYPE>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<UNITS>0</UNITS>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<UNITPRICE>" & floatPrice & "</UNITPRICE>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<MKTVAL>" & floatPrice & "</MKTVAL>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<DTPRICEASOF>" & Year(Date) & zeroFormat(Month(Date)) & zeroFormat(Day(Date)) & _
                                                     zeroFormat(Hour(Time)) & zeroFormat(Minute(Time)) & "00.000[-5:EST]" & "</DTPRICEASOF>"
    rtnString = rtnString & vbCrLf & vbCrLf & "</INVPOS>"
    rtnString = rtnString & vbCrLf & vbCrLf & "</POSSTOCK>"
    posstock = rtnString

End Function


Function posmf(strSecurity, floatPrice)
    
    rtnString = vbCrLf & vbCrLf & "<POSMF>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<INVPOS>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<SECID>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<UNIQUEID>" & strSecurity & "</UNIQUEID>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<UNIQUEIDTYPE>TICKER</UNIQUEIDTYPE>"
    rtnString = rtnString & vbCrLf & vbCrLf & "</SECID>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<HELDINACCT>CASH</HELDINACCT>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<POSTYPE>LONG</POSTYPE>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<UNITS>0</UNITS>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<UNITPRICE>" & floatPrice & "</UNITPRICE>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<MKTVAL>" & floatPrice & "</MKTVAL>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<DTPRICEASOF>" & Year(Date) & zeroFormat(Month(Date)) & zeroFormat(Day(Date)) & _
                                                     zeroFormat(Hour(Time)) & zeroFormat(Minute(Time)) & "00.000[-5:EST]" & "</DTPRICEASOF>"
    rtnString = rtnString & vbCrLf & vbCrLf & "</INVPOS>"
    rtnString = rtnString & vbCrLf & vbCrLf & "</POSMF>"
    posmf = rtnString

End Function


Function stockinfo(strSecurity, floatPrice)

    rtnString = vbCrLf & vbCrLf & "<STOCKINFO>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<SECINFO>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<SECID>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<UNIQUEID>" & strSecurity & "</UNIQUEID>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<UNIQUEIDTYPE>TICKER</UNIQUEIDTYPE>"
    rtnString = rtnString & vbCrLf & vbCrLf & "</SECID>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<SECNAME>NA</SECNAME>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<TICKER>" & strSecurity & "</TICKER>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<UNITPRICE>" & floatPrice & "</UNITPRICE>"
    rtnString = rtnString & vbCrLf & vbCrLf & "</SECINFO>"
    rtnString = rtnString & vbCrLf & vbCrLf & "</STOCKINFO>"
    stockinfo = rtnString

End Function

Function mfinfo(strSecurity, floatPrice)

    rtnString = vbCrLf & vbCrLf & "<MFINFO>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<SECINFO>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<SECID>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<UNIQUEID>" & strSecurity & "</UNIQUEID>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<UNIQUEIDTYPE>TICKER</UNIQUEIDTYPE>"
    rtnString = rtnString & vbCrLf & vbCrLf & "</SECID>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<SECNAME>NA</SECNAME>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<TICKER>" & strSecurity & "</TICKER>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<UNITPRICE>" & floatPrice & "</UNITPRICE>"
    rtnString = rtnString & vbCrLf & vbCrLf & "</SECINFO>"
    rtnString = rtnString & vbCrLf & vbCrLf & "<MFTYPE>OPENEND</MFTYPE>"
    rtnString = rtnString & vbCrLf & vbCrLf & "</MFINFO>"
    mfinfo = rtnString

End Function



Function zeroFormat(intNum)

    If (intNum < 10) Then
        zeroFormat = "0" & intNum
    Else
        zeroFormat = intNum
    End If

End Function
````


## Author

* **Mario Delgado**  Github: [MarioDelgadoSr](https://github.com/MarioDelgadoSr)
* LinkedIn: [Mario Delgado](https://www.linkedin.com/in/mario-delgado-5b6195155/)
* [My Data Visualizer](http://MyDataVisualizer.com/demo/): A data visualization application using the *DataVisual* design pattern.


## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details




