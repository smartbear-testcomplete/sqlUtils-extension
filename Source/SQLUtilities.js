//This code and the associated script extension are provided under a Creative Commons Attribution-ShareAlike 4.0 International License.
//You may use and modify this code according to the license, giving appropriate attribution and, if you distribute it, you must do so
//under the same license.

var databaseName;
var SQLServerName;
var connectionString;
var userName;
var passWord;


function getDatabaseName() {
    return databaseName;
}

function setDatabaseName(value) {
    databaseName = value;
}

function getSQLServerName() {
    return SQLServerName;
}

function setSQLServerName(value) {
    SQLServerName = value;
}

function getConnectionString() {
    connectionString = buildConnectionString();
    return connectionString;
}

function getUserName() {
    return userName;
}

function setUserName(value) {
    userName = value;
}

function getPassWord() {
    return passWord;
} 

function setPassWord(value) {
    passWord = value;
}

function setSQLType(value){
    Options.SQLType = aqString.ToUpper(value);    
}

function setSecurityType(value){
    if (Options.SQLType == 'MSSQL') {
        Options.SecurityType = aqString.ToUpper(value);
    }
    else {
        Options.SecurityType = 'PROMPT';
    }   
}

function buildConnectionString() {
    var localString;
    switch (Options.SQLType) {
        case 'MSSQL': 
            localString = 'Provider=SQLOLEDB.1;Persist Security Info=False;Initial Catalog=';
            localString = localString + databaseName;
            localString = localString + ';Data Source=' + SQLServerName;
            if (Options.SecurityType == 'PROMPT') {
                localString = localString + ';User ID=' + userName + ';Password=' + passWord + ';';
            }
            else {
                localString = localString + ';Integrated Security=SSPI';
            }
            break;
        case 'MYSQL_351':
            localString = 'Driver={MySQL ODBC 3.51 Driver};Server=' + SQLServerName + ';Database=' + databaseName + ';User=' + userName + ';Password=' + passWord + ';Option=3';
            break;
        case 'ORACLE':
            localString = 'Provider=MSDASQL.1;Persist Security Info=False;User ID=' + userName + ';Data Source=' + databaseName';
            break;
        default :
            Log.Warning('Unknown SQL type encountered. Defaulting to MSSQL with Integrated Security');
            Options.SQLType = 'MSSQL';
            Options.SecurityType = 'INTEGRATED';
            localString = 'Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=';
            localString = localString + databaseName;
            localString = localString + ';Data Source=' + SQLServerName;          
    }

    return localString;
}

function queryObject(query, parameters) {
    this.string = query;
    if (parameters == undefined) {
        parameters = [];
    }
    this.parameterArray = [];
    for (parameter in parameters) {
        this.parameterArray.push({name: param, value: parameters[param]});
    }
}


function getSQLFromFile(filePath) {
    return aqFile.ReadWholeTextFile(filePath, aqFile.ctANSI);
}

function open() {
//Opens the connection on the query object and sets it as the Active Connection on the command object property
    this.localConnection.Open();
    this.localCommand.CommandObject.ActiveConnection = this.localConnection;    
}

function closeAndNull() {
//Closes the connection on the object and sets the connection and command objects to null to free up memory space
    if (this.localConnection.State != 0) {
        this.localConnection.Close();
    }
    this.localConnection = null;
    this.localCommand = null;
}

function execute() {
//assigns the query to the command text, parses it for parameters, and then executes the query as a command, returning any recordset
    this.localCommand.CommandText = this.queryObject.queryString;
    this.localCommand.Parameters.ParseSQL(this.localCommand.CommandText, true);        
    for (var i = 0; i < this.queryObject.parameterArray.length; i++) {

        this.localCommand.Parameters.ParamByName(this.queryObject.parameterArray[i].name).Value = this.queryObject.parameterArray[i].value;
    }
    return this.localCommand.Execute();
}

function sqlQuery(sqlString, parameters, stringIsFile) {
    try {
        if (stringIsFile == undefined) {
            stringIsFile = false;
        }
        var localObject = {};
        var queryString;
        stringIsFile ? queryString = getSQLFromFile(sqlString) : queryString = sqlString;
        localObject.queryObject = new queryObject(sqlString, parameters);
        localObject.localConnection = ADO.CreateConnection();
        localObject.localConnection.ConnectionString = getConnectionString();
        localObject.localCommand = ADO.CreateADOCommand();
        localObject.open = open;
        localObject.execute = execute;
        localObject.closeAndNull = closeAndNull;
        return localObject;
    }
    catch (exception) {    
        Log.Error('Could not execute sqlQuery: ' + exception.message, exception.stack);
    }
}

function FormatDateForSQL(OffsetDays)
{
    try
        {
        OffsetDays = aqConvert.VarToInt(OffsetDays);
        }
    catch (e)
        {
        Log.Warning("Value of OffsetDays is non-numeric, using zero (0) instead");
        OffsetDays = 0;                                                    
        }
    
    var DateValue = aqDateTime.Today();
    DateValue = aqDateTime.AddDays(DateValue, OffsetDays);
    return aqConvert.DateTimetoFormatStr(DateValue, '%Y-%m-%d');
}
