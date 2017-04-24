
  //2015 AHF - Antonio H. Ferreira <disal.antonio.ferreira@gmail.com>
  //check 01_readme.pdf before using adordd
  //any application should work by setting these SETS and
  //uploading tables
 #include "fivewin.ch"
 #include "inkey.ch"

 #INCLUDE "adordd.ch"

 #include "adordd.ch"
 #ifndef __XHARBOUR__
 #include "hbcompat.ch"  //27.10.15 jose quintas advise
 #endif

 FUNCTION Main()
    LOCAL cSql :="", n, cCod, z
   
    RddRegister("ADORDD",1)
    RddSetDefault("ADORDD")

    //Index related sets
    SET ADODBF TABLES INDEX LIST TO {  { "TABLE1", { "FIRST", "FIRST" } }, { "TABLE2", { "CODID", "CODID" } } }
    
    //Works just like structural multibag indexes. It should be more than one name index for each table
    // { { "INDEX NAME SAME AS TABLE IF DIFERENT WILL NOT AUTO OPEN", { "TAG NAME" } },;
    //   { "INDEX NAME SAME AS TABLE IF DIFERENT WILL NOT AUTO OPEN" ,{ "TAG NAME"  } } }
    SET ADODBF MULTIBAG INDEX LIST TO { { "TABLE1", { "FIRST" } },;
                                        { "TABLE2" ,{ "CODID"  } } }
    
                                            
    SET ADO TEMPORAY NAMES INDEX LIST TO { "TMP", "TEMP", "A1", "ATESTE1" }  //uses "tmp1","temp1","ateste11""

    //UDFs index expressions. Please only use this if absolutly needed!
    //It takes longer to process.
    //Conversion functions as long as dont change lenght of field expressions
    //Do not need to come here!
    SET ADO INDEX UDFS TO { "IF", "&", "SUBSTR", "==", "DESCEND" }

    //If engine does not support logical fields indicate them here
    //ex tinyint if used as logical comes here!
    /*
    SET ADO TABLES LOGICAL FIELDS LIST TO  { { "TABLE1", { "SOMEFIELD" } } }
    */

    //If engine does not support precise indication of decimals like money, etc put them here
    /*
    SET ADO TABLES DECIMAL FIELDS LIST TO  { { "TABLE1", { "SOMEFIELD", 4, "ANOTHERFILED", 6 } } }
    */

    //Defining numeric field len used in index expressions WITHOUT PRECISE LEN NOTATION IN SQL TABLE
    //adordd will not work 100%
    /*
    SET ADODBF INDEX LIST FIELDTYPE NUMBER TO { { "TABLE1", { "NUMFIELD", 2 } } }
    */
    /*
    Adordd places all this information after runing hb_adoupload() in:
    BOOELANFIELDS.ADO
    DECIMALFIELDS.ADO
    Just copy it and pasted in the correspondong arrays
    */

    //Field recno and deleted related sets
    SET ADO DEFAULT RECNO FIELD TO "HBRECNO"
    //Only needed for tables with diferent from the default
    /*
    SET ADO FIELDRECNO TABLES LIST TO {{"TABLE1","????"},{"TABLE2","????"}}
    */
    
    SET ADO DEFAULT DELETED FIELD TO "HBDELETE"
    //Only needed for tables with diferent from the default
    /*
    SET ADO FIELDDELETED TABLES LIST TO {{"TABLE1","?????"},{"TABLE2","???"} }
    */

    //Lock related sets
    //Control locking in adordd for both table and record dont put final "\"
    //Uncoment a place a folder if lock set on
    //On WAN adordd needs a share over a VPN!
    /*
    SET ADO LOCK CONTROL SHAREPATH TO  "C:\TEMP" RDD TO "DBFCDX"
    */
    SET ADO FORCE LOCK OFF

    //Table names related sets
    //Table names with or without path ex. cpath_tablename or tablename
    //tables must be created or imported with the same set
    SET ADO TABLENAME WITH PATH OFF

    //If this set is on we need a path
    //SET PATH TO "C:\WHATEVER"
    /*
    SET ADO ROOT PATH TO "actual path" INSTEAD OF "uploaded path"
    */

    //Connection related sets 
    SET ADO DEFAULT DATABASE TO "TESTADORDD.MDB" SERVER TO "ACCESS" ENGINE TO "ACCESS" USER TO "" PASSWORD TO ""
    //SET ADO DEFAULT DATABASE TO "mydatabase" SERVER TO "localhost"  ENGINE TO "MYSQL" USER TO "myuser" PASSWORD TO "mypass"
    //SET ADO DEFAULT DATABASE TO "mydatabase" SERVER TO "localhost"  ENGINE TO "MARIADB" USER TO "myuser" PASSWORD TO "mypass"
    //SET ADO DEFAULT DATABASE TO "drive:\folder\mydb.DB" SERVER TO "" ENGINE TO "SQLITE" USER TO "Myusers" PASSWORD TO "Mypass"
    //SET ADO DEFAULT DATABASE TO "mydb" SERVER TO "localhost" ENGINE TO "POSTGRE" USER TO "myuser" PASSWORD TO "mypass"
    //SET ADO DEFAULT DATABASE TO "drive:\folder\mydb.FDB" SERVER TO "localhost" ENGINE TO "FIREBIRD" USER TO "SYSDBA" PASSWORD TO "masterkey"
    //SET ADO DEFAULT DATABASE TO "mydb" SERVER TO "127.0.0.1" ENGINE TO "ORACLE" USER TO "SYSTEM" PASSWORD TO "mypass"
    //other valid engines are:
    /*
    MSSQL
    INFORMIX
    ANYWHERE
    ADS  - better through native drives
    FOXPRO
    DBASE
    */
    SET EXCLUSIVE OFF  // even lock set off if we need to have locklist array we need to set thos one OFF
    SET AUTOPEN ON //Might be ON if you wish on it opens index multibag with same name as table name
    SET AUTORDER TO 1 //First index opened can be other

    /*         TRY TO INSERT 10.000.000 ROWS IN A TABLE AND THEN TRY THESE SETS WITH DIFERENT OPTIONS    */

    //For big recordsets try diferent options. Here ADORDD caches recordsets > 50 records
    //This SET is used by ADORDD and MS ADO object.
    SET ADO CACHESIZE TO 50 ASYNC ON ASYNCNOWAIT ON

    //This table will be opened with this where clause a way to reduce number of records to what we need
    //then we can change it during run time with 
    /*
    cOldSql := adowhereclause( Nwa, cSql )
    */
    SET RECORDSET OPEN WHERE CLAUSE TO { { "table1", "AGE > 39"  } }

    //Pre open and cache tables with records in the table >= nrcords or and with names of mask in table name if defined
    //This is done during app initialization to be much faster during runtime
    SET ADO PRE OPEN THRESHOLD TO 500 MASK { "TABLE" } //opens "orders";"orderinvoiced","order20" etc

    /*
    If you want to test it with your own tables comment the code below and do:

     hb_AdoUpload( "YOUR DRIVE WITH PATH FINISHING WITH \", "DBFCDX", "ACCESS OR MYSQL OR OTHER", oOverWrite .F. )

    and write your own testing routines

    THATS IT !
    */

     //ATTENTION BESIDES MSACCESS ADORDD DOESNT CREATE THE DATABASE

    //this is an idea it has not been tested but it should work
    IF !hb_adoRddExistsTable( ,"table1")
       //need to include complete path defaults to SET ADO DEFAULT DATABA
       DbCreate("table1", ;
                                 { { "CODID",   "C", 10, 0  },;
                                   { "FIRST",   "C", 30, 0  },;
                                   { "LAST",    "C", 30, 0  },;
                                   { "AGE",     "N",  8, 0  },;
                                   { "HBRECNO", "+", 11, 0  } ,;
                                   { "HBDELETE",  "L", 1,0  } }, "ADORDD" )
       
       SELE 0
       USE table1 ALIAS "TEST1"

       APPEND BLANK
       test1->First   := "HOMER si no Homer"
       test1->Last    := "Simpson"
       test1->Age     := 45
       test1->codid   := "0001"

       APPEND BLANK
       test1->First   := "Lara"
       test1->Last    := "Croft si no"
       test1->Age     := 32
       test1->codid   := "0002"
       test1->(dbcommit())

    ELSE
       SELE 0
       USE table1 ALIAS "TEST1"

    ENDIF

    IF !hb_adoRddExistsTable( ,"table2")
      //need to include complete path defaults to SET ADO DEFAULT DATABA
      DbCreate( "table2", ;
                                { { "CODID",    "C", 10, 0 },;
                                  { "ADDRESS",  "C", 30, 0 },;
                                  { "PHONE",    "C", 30, 0 },;
                                  { "EMAIL",    "C", 100,0 },;
                                  { "HBRECNO",  "+", 11,0  },;
                                  { "HBDELETE",  "L", 1,0  }}, "ADORDD" )

       SELE 0
       USE table2 ALIAS "TEST2"

       APPEND BLANK
       test2->address := "742 Evergreen Terrace"
       test2->phone   := "01 2920002"
       test2->email   := "homer@homersimpson.com"
       test2->codid   := "0001"

       APPEND BLANK
       test2->address := "Raymond Street"
       test2->phone   := "0039 29933003"
       test2->email   := "lara@laracroft.com"
       test2->codid   := "0002"
       test2->(dbcommit())

    ELSE
       SELE 0
       USE table2 ALIAS "TEST2"

    ENDIF

    IF test1->( LASTREC() ) < 3

        n := 1
        //INSERT ROWS
        DO WHILE n < 1900
           IF n > 2
              cCod := PADL(  n , 10, "0" )
              hb_GetAdoConnection():Execute("INSERT INTO TABLE1 ( CODID, [FIRST], [LAST], AGE ) VALUES ('"+;
                                            cCod+"','Firstname"+cCod+"','LastName"+cCod+"',"+;
                                            ALLTRIM( STR( HB_RANDOMINT( 30, 50 ) ) )+")")

           ENDIF
           n++

        ENDDO

    ENDIF
    
    IF test2->( LASTREC() ) < 3 
       n := 1
       z := 1
       cCod := PADL(  n , 10, "0" )
       //INSERT ROWS
       DO WHILE n < 10000
          IF n > z+10  //SOME 10 ROWS FOR EACH RECORD TABEL1
             cCod := PADL(  n , 10, "0" )
             z := n

          ENDIF

          IF n > 2
             hb_GetAdoConnection():Execute("INSERT INTO TABLE2 ( CODID, ADDRESS, PHONE, EMAIL ) VALUES ('"+cCod+;
                                           "','Street nr "+cCod+"','Phone "+cCod+"','MyMail@"+cCod+"')" )
 
          ENDIF
          n++

       ENDDO
       
    ENDIF 

    CLOSE ALL

    SELE 0
    USE table1 ALIAS "TEST1"
    SELE 0
    USE table2 ALIAS "TEST2"
    EXPALIAS()
    
    //LOCKING TRIALS
    GOTO 1  
    IF DBRLOCK(1)
       MSGINFO("TABLE2 RECORD 1 LOCKED!"+CRLF+;
               "START ANOTHER INSTANCE OF APP BEFORE CLOSING THIS MESSAGE "+;
               "AND CHECK ITS LOCKED!"+CRLF+;
               "SET ADO FORCE LOCK MUST BE ON!"+CRLF+;
               "SET EXCLUSIVE MUST BE OFF"+CRLF+;
               "IF YOU NEED LOCK ARRAY LIST"+CRLF+;
               "ALTHOUGH LOCK MIGHT BE OFF SET EXCLUSIVE MUST BE OFF") 
       EXPALIAS()
       UNLOCK

    ELSE
       MSGINFO("TABLE 2 COULD NOT LOCK RECORD 1")

    ENDIF
    
    GO TOP

    SELE TEST1
    GO TOP
    MSGINFO("BROWSE DEFAULT ORDER TABLE1 "+ordname())
    Browse()

    SELE TEST2
    GO TOP
    MSGINFO("BROWSE DEFAULT ORDER TABLE2 "+ordname())
    Browse()

    SELE TEST1
    SET RELATION TO CODID INTO TEST2
    MSGINFO("SET RELATION TO CODID FROM TABLE1 TO TABLE2")
    EXPALIAS()
    GO 1
    n := 1
    DO WHILE n < 10 .AND. !EOF()
       MSGINFO("Name "+TEST1->FIRST+" Address "+TEST2->ADDRESS)
       n++
       DBSKIP()
    ENDDO

    MSGINFO("BROWSE TABLE1")
    BROWSE()

    MSGINFO("CHANGE ORDER CREATE INDEX ON LAST TABLE1")
    INDEX ON LAST TO TMP
    SET INDEX TO TMP
    EXPALIAS()
    BROWSE()

    cSql := "CREATE VIEW CONTACTS AS SELECT TABLE1.FIRST, TABLE1.LAST,"+;
            "TABLE1.AGE, TABLE2.ADDRESS, TABLE2.EMAIL, TABLE1.HBRECNO, TABLE1.HBDELETE "+;
            "FROM TABLE1 LEFT OUTER JOIN TABLE2 ON TABLE1.CODID = TABLE2.CODID"
    MSGINFO("RUNING SQL "+cSql)

    TRY
       hb_GetAdoConnection():EXECUTE(cSql)
    CATCH
       ADOSHOWERROR( hb_GetAdoConnection())
    END

    SELE 0
    USE CONTACTS
    MSGINFO("BROWSING VIEW CONTACTS")
    BROWSE()
    INDEX ON ADDRESS TO TMP2
    SET INDEX TO TMP2
    MSGINFO("INDEXED BY ADRESS")
    EXPALIAS()
    BROWSE()
    
    //working directly with recordset in another area
    MSGINFO("GET RECORDSET FOR TABLE TEST1 "+STR(SELECT("TEST1")) )

    oRs := hb_adoRddGetRecordSet(SELECT("TEST1"))

    oRs:close()
    aa := "SELECT * FROM "+hb_adoRddGetTableName( SELECT("TEST1") )+ " WHERE FIRST = 'Lara'"

    MSGINFO("NEW SELECT FOR RECORDSET TEST1 "+AA)
    oRs:open(aa,hb_adoRddGetConnection(SELECT("TEST1")))

    SELE TEST1
    EXPALIAS()
    BROWSE()
    REINDEX
    EXPALIAS()
    BROWSE()
 
    SELE TEST2
    MSGINFO("LETS SUPPOSE THAT TABLE HAS MILLIONS OF ROWS AND RECORDSET > 100k"+CRLF+;
            "WE NEED TO REDUCE TEMPORARLY RECORDSET TO RUN SOME SUM OPERATIONS")
    cSql := adowhereclause( Select(), "phone = '01 2920002'" )
    MSGINFO("APPLY NEW WHERE CLAUSE ALL AREA INFORMATION CONTINUES THE SAME"+CRLF+;
            "phone = '01 2920002'")
    EXPALIAS()
    BROWSE()
    MSGINFO("GET BACK TO THE PREVIOUS QUERY "+CRLF+;
             csql )
    adowhereclause( Select(), csql )
    EXPALIAS()
    BROWSE()

    MSGINFO("DOES TABLE1 EXISTS ON DB ?"+CVALTOCHAR(hb_adoRddExistsTable( ,"Table1") ))
    MSGINFO("DOES TABLE3 EXISTS ON DB ?"+CVALTOCHAR(hb_adoRddExistsTable( ,"Table3") ))
    DbCloseAll()

    msginfo("End of trials!")

   RETURN nil

//display area information
FUNCTION expalias()
   LOCAL cerrorlog := "",n,j,ntarget,x

   for n = 1 to 255
      if ! Empty( Alias( n ) )
         cErrorLog += CRLF + Str( n, 3 ) + ": " + If( Select() == n,"=> ", "   " ) + ;
                      PadR( Alias( n ), 15 ) + Space( 20 ) + "RDDName: " + ;
                      ( Alias( n ) )->( RddName() ) + CRLF
         cErrorLog += "     ==============================" + CRLF
         cErrorLog += "     RecNo    RecCount    BOF   EOF" + CRLF
         cErrorLog += "    " + Transform( ( Alias( n ) )->( RecNo() ), "99999" ) + ;
                      "      " + Transform( ( Alias( n ) )->( RecCount() ), "99999" ) + ;
                      "      " + cValToChar( ( Alias( n ) )->( BoF() ) ) + ;
                      "   " + cValToChar( ( Alias( n ) )->( EoF() ) ) + CRLF + CRLF
         cErrorLog += "     Orders in use " + Space( 23 ) + "TagName" + CRLF
         for j = 1 to 15
            if ! Empty( ( Alias( n ) )->( IndexKey( j ) ) )
               cErrorLog += Space( 8 ) + ;
                            If( ( Alias( n ) )->( IndexOrd() ) == j, "=> ", "   " ) + ;
                            PadR( ( Alias( n ) )->( IndexKey( j ) ), 35 ) + ;
                            ( Alias( n ) )->( OrdName( j ) ) + ;
                            CRLF
            endif
         next
         cErrorLog += CRLF + "     Relations in use " + CRLF
         for j = 1 to 8
            if ! Empty( ( nTarget := ( Alias( n ) )->( DbRSelect( j ) ) ) )
               cErrorLog += Space( 8 ) + Str( j ) + ": " + ;
                            "TO " + ( Alias( n ) )->( DbRelation( j ) ) + ;
                            " INTO " + Alias( nTarget ) + CRLF
               // uValue = ( Alias( n ) )->( DbRelation( j ) )
               // cErrorLog += cValToChar( &( uValue ) ) + CRLF
            endif
         next
      x :=  (alias(n))->(dbrlocklist())
      cErrorLog += Space( 8 ) +"Lock recrods "+CRLF
      for j = 1 to len(x)
          cErrorLog += Space( 8 ) + "Record nr "+str(x[j])+CRLF
      next

      endif
   next
   MEMOEDIT(cerrorlog)
   
   RETURN .t.
