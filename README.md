# adordd
adordd for (x)Harbour

If you are working with Harbour please move all STATICS top dordd.prg before  #ifndef __XHARBOUR__
otherwise youll get compile error.

1) Just add adordd.prg to your project and include adordd.ch

2) Set parameters as in tryadordd.prg (see in adordd.ch for syntax)

3) You can upload tables to any SQL with:

use "table" via "dbfcdx"

copy "table" to "sqtable" via "adordd

"use sqltable"


No code change in your apps with the exceptions of:
All expressions with variables for ex. index expressions the vars must be evaluated before sending it to adordd.

and thats it!

12.06.15

adordd its ready.

Filters working just like any other rdd.

Performance improvments

Small bugs corrections

25.5.15

Small bugs corrections.

Dates manipulation corrected.

Scopes working

Filters working just like normal dbf syntax you dont need to change anything.

copy to and append from working

Still problems with speed of dbeval.

15.05.15

Transactions enabled
Use 
ADOBEGINTRANS()

ADOCOMMITTRANS()

ADOROLLBACKTRANS() 

Still problems with some browses where the order is by field date type

DBEVAL still didnt solve performance problem its too slow!

Im only trying it with MYSql.

ADORDD its ready! Read Important Notes below.


14.05.15

APPEND FROM and COPY TO done! Only for database type dindt test it for Delimiter,SDF, etc. We dont need it.

Adordd it has been all this week on trials in real app.

Found and corrected some minor bugs.

Still problems with some browses where the order is by field date type

DBEVAL still didnt solve performance problem its too slow!

All major functionality as any other rdd done!

Im only trying it with MYSql.

ADORDD its ready! Read Important Notes below.


09.05.15

Corrected add, update and delete records.
Reccount changed because of speed problems.
Concurrent access and locks improved Use exclusive ok

TO DO
Relations are slow we are improving it.
Trigger to check for outdated recordset.

Please try it for ex for MySql (we are only trying it now with MySQl)

Please alter in tryadordd.prg:

SET ADO DEFAULT DATABASE TO "your DB name" SERVER TO "db4free.net"  ENGINE TO "MYSQL" USER TO "you user name" PASSWORD TO "your passord"

SET ADO LOCK CONTROL SHAREPATH TO  "your shared path" RDD TO "DBFCDX"
DbCreate( "your DB name;TABLE0;MYSQL;db4free.net;your user name;your passowrd",....



06.05.15

Most part of adordd its done!

Funtions you can call in adordd.prg to use in our app :

 ADOVERSION() Returns adordd version
 
 ADOSEEKSQL( nWA, lSoftSeek, cKey, lFindLast ) //returns a set of records meeting seek key
 
 ADORESETSEEK() //resets the recodset to previous before call ADOSEEKSQL()
 
 ADOBEGINTRANS(nWa)
 
 ADOCOMMITTRANS(nWa)
 
 ADOROLLBACKTRANS(nWa) 
 
 hb_adoRddGetConnection( nWorkArea ) Returns the connection for the workarea
 
 hb_adoRddGetRecordSet( nWorkArea )  Returns the recordset for the nWorkArea 
 
 hb_adoRddGetTableName( nWorkArea )  Returns tabe name for the nWorkArea  
 
 hb_adoRddExistsTable( oCon,cTable, cIndex ) Returns .t. if table or table and index exist on the DB
 
 hb_adoRddDrop( oCon, cTable, cIndex, DBEngine ) Drops (delete) table or index in the DB
 
 hb_GetAdoConnection() Returns ado default connection
 
 
All the rest its standard rdd functions we are used to.

No code change required.

Exceptions:

All expressions with variables for ex. index expressions the vars must be evaluated before sending it to adordd.


Thats all!

IMPORTANT NOTES:

Its very important that you indicate in adordd.prg in the several places with the cursorlocation the type of cursor you need to your DB.
Not choosing the right cursor the browses become irregular.
Please remember that all browse positioning is based on absoluteposition thus a cursor that can not support it doesnt work.
Also you need a cursor that supports update() and requery().
The cursors defined in adordd support:

ACCESS
MYSQL
ORACLE

All expressions with variables for ex. index expressions the vars must be evaluated before sending it to adordd.
Filters can not be evaluated without cFilter expression.

Using Browse() you must pass delete block because it does not lock the record and adorddd raises a lock required error.
Also when you change the index key value the browse() does not re-position immediately the grid. Click right or left arrows.

adordd has been working ok with :

ACCESS
MYSQL
ORACLE

both in internal network and internet and its ok.
Internet a little slow but not crawling with tables 30.000 recs.
You must find which are the best parameters for your case and adjust it in ADO_OPEN

Look for:
Code:

  //PROPERIES AFFECTING PERFORMANCE TRY
   //oRecordSet:MaxRecords := 60
   //oRecordSet:CacheSize := 50 //records increase performance set zero returns error set great server parameters max open rows error
   //oRecordset:PageSize = 10
   //oRecordSet:MaxRecords := 15
   //oRecordset:Properties("Maximum Open Rows") := 110  //MIN TWICE THE SIZE OF CACHESIZE
 


Please post your comments and findings.

