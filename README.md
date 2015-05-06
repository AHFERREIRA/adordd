# adordd
adordd for (x)Harbour

06.05.15

Most part of adordd its done!

Please alter in adord.prg the cursorlocation as supported by the DB you are using.
Like it is it supports ACCESS and Mysql / Oracle / ADS

Funtions you can call in adordd.prg to use in our app :

 ADOVERSION() Returns adordd version
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

Please read notes in tryadordd.prg.

