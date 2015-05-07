# adordd
adordd for (x)Harbour

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
Filters can not be evaluated without cFilter expression.

Thats all!

IMPORTANT NOTES:

Its very imortant that you indicate in adordd.prg in the several places with the cursorlocation the type of cursor you need to your DB.
Not choosing the right cursor the browses become irregular.
Please remember that all browse positioning is based on absoluteposition thus a cursor that can not support it doesnt work.
Also you need a cursor that supports update() and requery().
The cursors defined in adordd support:

ACCESS
MYSQL
ORACLE

Changes:
Locks now work as it should.
Exclusive use in progress not validated

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

