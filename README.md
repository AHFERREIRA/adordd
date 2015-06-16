# adordd
adordd for (x)Harbour its ready!

PLEASE MAKE A DONATION!

We will all profit for sure from such development so I think its fair to ask everyone to contribute with a minimal 
importance of 20 Euros (because of the paypal costs) for a good cause.
Go to ajusera.com scroll down and click PayPal button "Doar" 
They can send you a contribution receipt if you send them a email with your details.

This is a non profit organization that performs an important social work supporting the elderly, 
young problematic, food distribution to the needy, etc.
This organization is based on the work of Father Jeronimo Usera (spanish) well known in
Spain, Portugal and South America.

You can get more information at ajusera.com (sorry its only in portuguese)

To start working:

1) Just add adordd.prg to your project and include adordd.ch

2) Set parameters as in tryadordd.prg (see in adordd.ch for syntax)

 A) SET ADO TABLES INDEX LIST TO { {"TABLE1",{"FIRST","FIRST"} }, {"TABLE2" ,{"CODID","CODID"}} }
 
 This Set is used by the SQL engine to build select with order by.
 
 Thus the fields must be separated by comma and it can include SQL functions or ASC DESC
 This SEt can not include clipper functions as they are unkown to SQL.
 
 EX:
 
SET ADO TABLES INDEX LIST TO { {"TABLE1",{"FIRST","FIRST DESC"} }, {"TABLE2" ,{"CODID","CODID"}} }

B) SET ADODBF TABLES INDEX LIST TO {  {"TABLE1",{"FIRST","FIRST"} }, {"TABLE2" ,{"CODID","STR(CODID,2,0)"}} }

This Set is used to evaluate clipper expressions such as:

&( indexkey( 0 ) )

OrdKey( )

Etc

So it must contain your actual index expressions.

C) SET ADO TEMPORAY NAMES INDEX LIST TO {"TMP","TEMP"}

Indicates the names used for temporay files at SQL level.

It must start by TMP or TEMp but can be "TMPROGER"

These temporary fies are mainy used for temporary indexes created in the SQL sever as TEMPORARY and automaticly destroied
after connection ends.

They are only visible to the user that created them.

D) SET ADO FIELDRECNO TABLES LIST TO {{"TABLE1","HBRECNO"},{"TABLE2","HBRECNO"}}

This Set lets you indicate a diferent autoinc field of the defaut per table to be used as recno()

E) SET ADO DEFAULT RECNO FIELD TO "HBRECNO"

This Set indicates the default field name to be used as recno in all tables besides the mentioned above.

If you have the same field in all tables can only use this set.

ATTENTION:

The D and or E sets are absolutly necessary and without them the navigation with adordd might be unpredictable.

F) SET ADO DEFAULT DATABASE TO "D:\WHATEVER\TEST2.mdb" SERVER TO "CSEVER" ENGINE TO ACCESS USER TO "" PASSWORD TO ""

This Set inidcates the default server and database and parameters we are using.

Connection get established here.

G)  SET ADO LOCK CONTROL SHAREPATH TO  "D:\WHATEVER" RDD TO "DBFCDX"

This set enables adordd to assure locking records and exclusive use of files as any other rdd.

You need to supply a path where adordd creates the tlocks file to control this.

This rdd file must be a rdd working with locks as ex dbfcdx.

This is not a SQL table and if you need to work in WAN and need lock control you will need to:

The connection to SQL server

and

Ex a VPN where you can access this share.

3) You can upload tables to any SQL with:

use "table" via "dbfcdx"

copy "table" to "sqtable" via "adordd

use "sqltable" //use with set connection string
or
use sqltable@connection string

//you can use a table in a new connection
use "ctable@connection string" alias "whatever"

4) Funtions you can call in adordd.prg to use in our app :

 ADOVERSION() Returns adordd version
 
 ADOSEEKSQL( nWA, lSoftSeek, cKey, lFindLast ) //returns a set of records meeting seek key
 
 ADOBEGINTRANS(nWa)
 
 ADOCOMMITTRANS(nWa)
 
 ADOROLLBACKTRANS(nWa) 
 
 ADORESETSEEK( nWa ) //resets the recodset to previous before call ADOSEEKSQL()
 
 hb_adoRddGetConnection( nWorkArea ) Returns the connection for the workarea
 
 hb_adoRddGetRecordSet( nWorkArea )  Returns the recordset for the nWorkArea 
 
 hb_adoRddGetTableName( nWorkArea )  Returns tabe name for the nWorkArea  
 
 hb_adoRddExistsTable( oCon,cTable, cIndex ) Returns .t. if table or table and index exist on the DB
 
 hb_adoRddDrop( oCon, cTable, cIndex, DBEngine ) Drops (delete) table or index in the DB
 
 hb_GetAdoConnection() Returns ado default connection

No code change in your apps with the exceptions of:

All expressions with variables for ex. index expressions the vars must be evaluated before sending it to adordd.

Deleted records are immediatly out of the table and can not be recovered again.
Thus any operations on deleted records must occour before delete the record or an error will occour.

Operations like:

delete all

while...

  if lconditon
  
     recall record
     
must be inverted

while....

  if !lcondtion
  
     delete record
     

and thats it!

