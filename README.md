# README #

This extensions is intended to make it easier and portable to do SQL queries and operations against an SQL database without having to re-write code all over the place.  An ADODB Connection object is created in this extension and several methods are provided both to configure the environment as well as to execute SQL queries and commands on a configured SQL database.

## ABSTRACT ##

### Properties ###

* DatabaseName - Read/Write - Stores the database name for the SQL database for all queries run within the extension. Value is persisted within a TestComplete/TestExecute session but is not persisted across sessions.
* SQLServerName - Read/Write - Stores the SQL Server name for the SQL database for all queries run within the extension. Value is persisted within a TestComplete/TestExecute session but is not persisted across sessions.
* SQLUserName - Read/Write - Stores the SQL UserName for the SQL database for queries run within the extension. This field is only necessary if using either MYSQL or if using MSSQL with PROMPT database security. Value is persisted within a TestComplete/TestExecute session but is not persisted across sessions.
* SQLPassWord - Read/Write - Stores the SQL Password for the SQL database for queries run within the extension. This field is only necessary if using either MYSQL or if using MSSQL with PROMPT database security. Value is persisted within a TestComplete/TestExecute session but is not persisted across sessions.
* ConnectionString - Read Only - A connection string constructed and initialized based upon the properties entered and on the values of the SQLType and SQLSecurity

### Methods ###

The first two methods below are used to set options to persist. They can be called once and only once and never again if so desired just for initial set up.

* SetSQLType - no value returned - This sets the SQL connection type for the TestComplete environment. When this method is called and a value is passed, if the value is valid, that SQL connection type is persisted across TestComplete sessions on the same workstation. This allows the automation engineer to set up their environment for a particular SQL type and not have to reconfigure it each time.  Valid options are: MSSQL and MYSQL_351._
* SetSQLSecurityType - no value returned - This applies only to MSSQL connection type. When this method is called and a value is passed, if the value is valid, that SQL security type is persisted across TestComplete sessions on the same workstation. This allows the automation engineer to set up their environment for a particular SQL security configuration and not have to reconfigure it each time. This configures either integrated security or security that embeds a username and password. Valud values are: INTEGRATED and PROMPT

Two additional methods are provided for instantiating objects to be used for executing SQL queries.

* **NewQueryObject** - Returns an object with two properties based upon the parameters passed in. The first parameter is a string that contains the actual SQL query text itself. The second parameter is an JScript object containing any parameters needed to execute the SQL query (see example below)
* **NewSQLQuery** - Returns an object for executing an SQL query. It has one parameters, that being a query object as created by NewQueryObject.

The following method is provided to convert a date value into something meanful for SQL

* **FormateDateForSQL** - returns a date in the format of YYYY-MM-DD - parameter is an integer value - Use this method to add or subtract days from the current date (Today) and format the result into a date string that matches SQL formatting

# Example #

The following example demonstrates using this extension in a project:

```
#!javascript
//The following function can be included in any project but it only ever needs to be run the first time to set up the desired settings
function SQLOptions(){
    SQLUtilities.SetSQLType('MSSQL'); //Can be either MSSQL or MYSQL_351
	SQLUtilities.SetSQLSecurityType('INTEGRATED'); //Can be either INTEGRATED or PROMPT
}

//Now, if I want to run an SQL query against my database where I log the street address for all addresses in Pennsylvania, USA
function getPAStreetAddresses() {
    var mySQLQuery;
    var queryObject;
	var myRecords;
    SQLUtilities.DatabaseName = 'MyDatabase';
	SQLUtilities.SQLServerName = 'MyServer';
	SQLUtilities.SQLUserName = 'myusername'; //This is not necessary since I'm using integrated security. Provided in example to allow for editing and adaptation
	SQLUtilities.SQLPassWord = 'mypassword'; //This is not necessary since I'm using integrated security. Provided in example to allow for editing and adaptation
    queryObject = SQLUtilities.NewQueryObject('SELECT * FROM ADDRESSES WHERE STATE = :STATE and COUNTRY = :COUNTRY', {STATE: 'PA', COUNTRY: 'USA'})
	mySQLQuery = SQLUtilities.NewSQLQuery(queryObject);
	mySQLQuery.open();
	myRecords = mySQLQuery.execute();
	while (!myRecords.EOF) {
	    Log.Message('Street Address: ' + myRecords.Fields.Item('STREETADDRESS').Value);
		myRecords.MoveNext();
	}
	mySQLQuery.closeAndNull;
}
```

