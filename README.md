# Excal_vba


  <h2>1. SQLDatabase_VBA.bas</h2>
  <h3>connect SQL Database from ADODB or system with VBA in Excel</h3>
  
```
 sconnect = "Provider=MSDASQL.1;DSN=your ODBC connection name; " & _
            "UID=your user;PWD=your password;DBQ=your database" & DBPath & ";HDR=Yes';"
```


  <h3>connection String for IBM/AS400:</h3>

```
  sconnect = "PROVIDER=IBMDA400;Data Source=servername; " & _
            "DEFAULT COLLECTION=optional;USER ID=Username ;PASSWORD=KENNWORT"
```
