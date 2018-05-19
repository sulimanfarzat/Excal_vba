Attribute VB_Name = "SQLDatabase_VBA"
Sub SQLDatabase_VBA()
    
On Error Resume Next

    'Step 1: Create the Connection String with Provider and Data Source options
    Public sSQLQry As String
    Public ReturnArray
        
    Public Conn As New ADODB.Connection
    Public mrs As New ADODB.Recordset
    Public DBPath As String, sconnect As String


    'Step 2: Create the Connection String with Provider and Data Source options
    ActiveSheet.Activate
    
    DBPath = ThisWorkbook.FullName 'Refering the sameworkbook as Data Source
    
    'You can provide the full path of your external file as shown below
    'DBPath ="C:\InputData.xlsx"
    
    sconnect = "Provider=MSDASQL.1;DSN=Connect_fromODBC;UID=your user name;PWD=your password;DBQ=database name" & DBPath & ";HDR=Yes';"
    'If any issue with MSDASQL Provider, Try the Microsoft.Jet.OLEDB:
    'sconnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBPath _
        & ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"";"
        
    'Step 3: set connection timeout Open the Connection to data source
    Conn.ConnectionTimeout = 30
    Conn.Open sconnect
    
    'Step 4: Create SQL Command String MRFIRM, MRIDEN, MRSART,MRSRN,MRSRRF,MRDTB,MRUSER, MRSRNA as Serien_NR_Zugriff
     sSQLSting = "SELECT * From your database " & _
                " WHERE ------ " & _
                " Group by ----- "
               
                

    'Step 5: Get the records by Opening this Query with in the Connected data source
     mrs.Open sSQLSting, Conn
     
     'Step 6: Copy the reords into our worksheet
     'Import Headers
        For i = 0 To mrs.Fields.Count - 1
            ActiveSheet.Range("B15").Offset(0, i) = mrs.Fields(i).Name
        Next i
        
    'Import data to destination cell
    ActiveSheet.Range("B15").Offset(1, 0).CopyFromRecordset mrs
    
     'Step 7: Close the Record Set and Connection
      'Close Recordset
      mrs.Close

      'Close Connection
      Conn.Close
      Set mrs = Nothing
      Set Conn = Nothing
     
End Sub

