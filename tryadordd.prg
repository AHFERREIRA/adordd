
//2015 AHF - Antonio H. Ferreira <disal.antonio.ferreira@gmail.com>
//check 01_readme before using adordd

 FUNCTION Main()
 LOCAL cSql :=""

    RddRegister("ADORDD",1)
    RddSetDefault("ADORDD")

    SET ADODBF TABLES INDEX LIST TO {  {"TABLE1",{"FIRST","FIRST"} }, {"TABLE2" ,{"CODID","CODID"}} }
    SET ADO TEMPORAY NAMES INDEX LIST TO {"TMP","TEMP"}
    //ony needed for tables with diferent from the default
    //SET ADO FIELDRECNO TABLES LIST TO {{"TABLE1","HBRECNO"},{"TABLE2","HBRECNO"}}
    SET ADO DEFAULT RECNO FIELD TO "HBRECNO"
    SET AUTOPEN ON //might be OFF if you wish
    SET AUTORDER TO 1 // first index opened can be other

    //need to include complete path
    SET ADO DEFAULT DATABASE TO "D:\WHATEVER\TEST2.mdb" SERVER TO "CSEVER" ENGINE TO ACCESS USER TO "" PASSWORD TO ""

    //CONTROL LOCKING IN ADORDD FOR BOTH TABLE AND RECORD DONT PUT FINAL "\"
    SET ADO LOCK CONTROL SHAREPATH TO  "D:\WHATEVER" RDD TO "DBFCDX"

    SET ADO DEFAULT DELETED FIELD TO "HBDELETED" /* defining the default name for DELETED field*/

   IF !FILE(   "\test2.mdb"   )
      //need to include complete path defaults to SET ADO DEFAULT DATABA
      DbCreate("table1;\test2.mdb", ;
                                { { "CODID",   "C", 10, 0 },;
                                  { "FIRST",   "C", 30, 0 },;
                                  { "LAST",    "C", 30, 0 },;
                                  { "AGE",     "N",  8, 0 },;
                                  { "HBRECNO", "+", 10, 0  } }, "ADORDD" )
      //need to include complete path defaults to SET ADO DEFAULT DATABA
      DbCreate( "table2;\test2.mdb", ;
                                { { "CODID",    "C", 10, 0 },;
                                  { "ADDRESS",  "C", 30, 0 },;
                                  { "PHONE",    "C", 30, 0 },;
                                  { "EMAIL",    "C", 100,0 },;
                                  { "HBRECNO",  "+", 10,0  } }, "ADORDD" )

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

     CLOSE ALL
   ENDIF

   SELE 0
   USE table1 ALIAS "TEST1"
   SELE 0
   USE table2 ALIAS "TEST2"

   //LOCKING TRIAL
   GOTO 1

   IF DBRLOCK(1)
      MSGINFO("TABLE 2 RECORD 1 LOCKED! START ANOTHER "+;
              "INSTANCE OF APP BEFORE CLOSING THIS MESSAGE"+;
              " CHECK LOCK!")
      UNLOCK

   ELSE
      MSGINFO("TABLE 2 COULD NOT LOCK RECORD 1")

   ENDIF

   GO TOP

   SELE TEST1
   GO TOP
   MSGINFO("BROWSE DEFAULT ORDER TABLE1")
   Browse()

   SELE TEST 2
   GO TOP
   MSGINFO("BROWSE DEFAULT ORDER TABLE2")
   Browse()

   SELE TEST1
   SET RELATION TO CODID INTO TEST2
   MSGINFO("SET RELATION TO CODID FROM TABLE1 TO TABLE2")
   GO TOP
   DO WHILE !EOF()
      MSGINFO("Name "+TEST1->FIRST+" Address "+TEST2->ADDRESS)
      DBSKIP()
   ENDDO

   MSGINFO("BROWSE TABLE1")
   BROWSE()

   MSGINFO("CHANGE ORDER CREATE INDEX ON LAST TABLE1")
   INDEX ON LAST TO TMP
   SET INDEX TO TMP

   BROWSE()

   cSql := "CREATE VIEW CONTACTS AS SELECT TABLE1.FIRST, TABLE1.LAST,"+;
            "TABLE1.AGE, TABLE2.ADDRESS, TABLE2.EMAIL "+;
            "FROM TABLE1 LEFT OUTER JOIN TABLE2 ON TABLE1.CODID = TABLE2.CODID"
   MSGINFO("RUNING SQL "+cSql)

   TRY
      hb_GetAdoConnection()():EXECUTE(cSql)
   CATCH
      ADOSHOWERROR( hb_GetAdoConnection()())
   END

   SELE 0
   USE CONTACTS
   MSGINFO("BROWSING VIEW CONTACTS")
   BROWSE()
   INDEX ON ADDRESS TO TMP2
   SET INDEX TO TMP2
   MSGINFO("INDEXED BY ADRESS")
   BROWSE()

   //WORKING DIRECTLY WITH RECORDSET IN ANOTHER AREA
   MSGINFO("GET RECORDSET FOR TABLE TEST1 "+STR(SELECT("TEST1")) )

   oRs := hb_adoRddGetRecordSet(SELECT("TEST1"))

   oRs:close()
   aa := "SELECT * FROM "+hb_adoRddGetTableName( SELECT("TEST1") )+ " WHERE FIRST = 'Lara'"

   MSGINFO("NEW SELECT FOR RECORDSET TEST1 "+AA)
   oRs:open(aa,hb_adoRddGetConnection(SELECT("TEST1")))

   MSGINFO("CURRENT WORKAREA "+ALIAS())

   MSGINFO("BROWSE RECORDSET ALIAS TEST1")
   TEST1->(BROWSE())

   MSGINFO("DOES TABLE1 EXISTS ON DB ?"+CVALTOCHAR(hb_adoRddExistsTable( "Table1") ))
   MSGINFO("DOES TABLE3 EXISTS ON DB ?"+CVALTOCHAR(hb_adoRddExistsTable( "Table3") ))
   DbCloseAll()

RETURN nil

