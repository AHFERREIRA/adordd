*
 * Harbour Project source code:
 * ADORDD - RDD to automatically manage Microsoft ADO
 *
 * Copyright 2007 Fernando Mancera <fmancera@viaopen.com> and
 * Antonio Linares <alinares@fivetechsoft.com>
 * www - http://harbour-project.org
 *
 * Copyright 2007-2008 Miguel Angel Marchuet <miguelangel@marchuet.net>
 *  ADO_GOTOID( nWA, nRecord )
 *  ADO_GOTO( nWA, nRecord )
 *  ADO_OPEN( nWA, aOpenInfo ) some modifications
 *     Open: Excel files
 *           Paradox files
 *           Access with password
 *           FireBird
 *  ADO_CLOSE( nWA )
 *  ADO_ZAP( nWA )
 *  ADO_ORDINFO( nWA, nIndex, aOrderInfo ) some modifications
 *  ADO_RECINFO( nWA, nRecord, nInfoType, uInfo )
 *  ADO_FIELDINFO( nWA, nField, nInfoType, uInfo )
 *  ADO_FIELDNAME( nWA, nField )
 *  ADO_FORCEREL( nWA )
 *  ADO_RELEVAL( nWA, aRelInfo )
 *  ADO_EXISTS( nRdd, cTable, cIndex, ulConnect )
 *  ADO_DROP(  nRdd, cTable, cIndex, ulConnect )
 *  ADO_LOCATE( nWA, lContinue )
 *
 * www - http://www.xharbour.org
 *
 * Copyright 2015 AHF - Antonio H. Ferreira <amhf@m-homem-ferreira.pt>
 *
 * Most part has been rewriten with a diferent kind of approach
 * not deal with Catalogs - DBA responsability
 * converting indexes to selects and treat indexes a "virtual" as they really dont exist as files
 * Only allows to create temp tables
 * Seek translate to find and if it is to slow can be converted to select but after using the result seek
 * one must build a resetseek to revert to the previous select - to be done if necessary
 * Sqlparser that converts filters using standard clipper expressions to sql taking into account
 * the active indexes to build order by clause used already since 90's in our app with ADS
 * Preventing sql injection
 *
 * This program is free software; you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation; either version 2, or (at your option)
 * any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this software; see the file COPYING.txt.  If not, write to
 * the Free Software Foundation, Inc., 59 Temple Place, Suite 330,
 * Boston, MA 02111-1307 USA (or visit the web site http://www.gnu.org/).
 *
 * As a special exception, the Harbour Project gives permission for
 * additional uses of the text contained in its release of Harbour.
 *
 * The exception is that, if you link the Harbour libraries with other
 * files to produce an executable, this does not by itself cause the
 * resulting executable to be covered by the GNU General Public License.
 * Your use of that executable is in no way restricted on account of
 * linking the Harbour library code into it.
 *
 * This exception does not however invalidate any other reasons why
 * the executable file might be covered by the GNU General Public License.
 *
 * This exception applies only to the code released by the Harbour
 * Project under the name Harbour.  If you copy code from other
 * Harbour Project or Free Software Foundation releases into a copy of
 * Harbour, as the General Public License permits, the exception does
 * not apply to the code that you add in this way.  To avoid misleading
 * anyone as to the status of such modified files, you must delete
 * this exception notice from them.
 *
 * If you write modifications of your own for Harbour, it is your choice
 * whether to permit this exception to apply to your modifications.
 * If you do not wish that, delete this exception notice.
 *
 */
#ifndef __XHARBOUR__

   #include "fivewin.ch"        // as Harbour does not have TRY / CATCH
   #define UR_FI_FLAGS           6
   #define UR_FI_STEP            7
   #define UR_FI_SIZE            5 // by Lucas for Harbour


#endif

ANNOUNCE ADORDD  

#include "rddsys.ch"
#include "fileio.ch"
#include "error.ch"
#include "adordd.ch"
#include "common.ch"
#include "dbstruct.ch"
#include "dbinfo.ch"

#include "hbusrrdd.ch"  //verify that your version has the size of 7 for xarbour at least for 2008 version

#define WA_RECORDSET   1
#define WA_BOF         2
#define WA_EOF         3
#define WA_CONNECTION  4
#define WA_CATALOG     5
#define WA_TABLENAME   6
#define WA_ENGINE      7
#define WA_SERVER      8
#define WA_USERNAME    9
#define WA_PASSWORD   10
#define WA_QUERY      11
#define WA_LOCATEFOR  12
#define WA_SCOPEINFO  13
#define WA_SQLSTRUCT  14
#define WA_CONNOPEN   15
#define WA_PENDINGREL 16
#define WA_FOUND      17
#define WA_INDEXES    18 //AHF
#define WA_INDEXEXP    19 //AHF
#define WA_INDEXFOR    20 //AHF
#define WA_INDEXACTIVE 21 //AHF
#define WA_LOCKLIST    22 //AHF
#define WA_FILELOCK    23 //AHF
#define WA_INDEXUNIQUE 24//AHF
#define WA_OPENSHARED  25//AHF
#define WA_SCOPES      26//AHF
#define WA_SCOPETOP    27//AHF
#define WA_SCOPEBOT    28//AHF
#define WA_ISITSUBSET  29//AHF
#define WA_SIZE        29

#define RDD_CONNECTION 1
#define RDD_CATALOG    2

#define RDD_SIZE       2

ANNOUNCE ADORDD

//THREAD 
STATIC t_cTableName
//THREAD 
STATIC t_cEngine
//THREAD 
STATIC t_cServer
//THREAD 
STATIC t_cUserName
//THREAD 
STATIC t_cPassword
//THREAD 
STATIC t_cQuery := ""
STATIC oConnection

STATIC FUNCTION ADO_INIT( nRDD )

   LOCAL aRData := Array( RDD_SIZE )

   USRRDD_RDDDATA( nRDD, aRData )

   RETURN HB_SUCCESS

STATIC FUNCTION ADO_NEW( nWA )

   LOCAL aWAData := Array( WA_SIZE )

   aWAData[ WA_BOF ] := .F.
   aWAData[ WA_EOF ] := .F.
   aWAData[WA_INDEXES] := {}
   aWAData[WA_INDEXEXP] := {}
   aWAData[WA_INDEXFOR] := {}
   aWAData[WA_INDEXACTIVE] := 0
   aWAData[WA_LOCKLIST] := {}
   aWAData[WA_FILELOCK] := .F.
   aWAData[WA_INDEXUNIQUE] := {}
   aWAData[WA_OPENSHARED] := .T.
   aWAData[WA_SCOPES] := {}
   aWAData[WA_SCOPETOP] := {}
   aWAData[WA_SCOPEBOT] := {}
   aWAData[WA_ISITSUBSET] := .F.
   
   USRRDD_AREADATA( nWA, aWAData )

   RETURN HB_SUCCESS

STATIC FUNCTION ADO_OPEN( nWA, aOpenInfo )

   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL cName, aField, oError, nResult
   LOCAL oRecordSet, nTotalFields, n
      
   /* When there is no ALIAS we will create new one using file name */
   IF Empty( aOpenInfo[ UR_OI_ALIAS ] )
      hb_FNameSplit( aOpenInfo[ UR_OI_NAME ],, @cName )
      aOpenInfo[ UR_OI_ALIAS ] := cName
   ENDIF
   
   aOpenInfo[ UR_OI_NAME ] += ".dbf"

   hb_adoSetTable( aOpenInfo[ UR_OI_NAME ] )
   hb_adoSetEngine( "")
   hb_adoSetServer( "")
   hb_adoSetQuery( )
   hb_adoSetUser( "")
   hb_adoSetPassword( "" )
   
   IF Empty( aOpenInfo[ UR_OI_CONNECT ] )
      IF EMPTY(oConnection)
		  aWAData[ WA_CONNECTION ] :=  TOleAuto():New( "ADODB.Connection" )
		  oConnection := aWAData[ WA_CONNECTION ]
		  aWAData[ WA_TABLENAME ] := t_cTableName
		  aWAData[ WA_QUERY ] := t_cQuery
		  aWAData[ WA_USERNAME ] := t_cUserName
		  aWAData[ WA_PASSWORD ] := t_cPassword
		  aWAData[ WA_SERVER ] := t_cServer
		  aWAData[ WA_ENGINE ] := t_cEngine
		  aWAData[ WA_CONNOPEN ] := .T.
		  aWAData[ WA_OPENSHARED ] := aOpenInfo[ UR_OI_SHARED ]
		  DO CASE
		  CASE Lower( Right( aOpenInfo[ UR_OI_NAME ], 4 ) ) == ".mdb"
			 IF Empty( aWAData[ WA_PASSWORD ] )
				aWAData[ WA_CONNECTION ]:Open( "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + aOpenInfo[ UR_OI_NAME ] )
			 ELSE
				aWAData[ WA_CONNECTION ]:Open( "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + aOpenInfo[ UR_OI_NAME ] + ";Jet OLEDB:Database Password=" + AllTrim( aWAData[ WA_PASSWORD ] ) )
			 ENDIF

		  CASE Lower( Right( aOpenInfo[ UR_OI_NAME ], 4 ) ) == ".xls"
			 aWAData[ WA_CONNECTION ]:Open( "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + aOpenInfo[ UR_OI_NAME ] + ";Extended Properties='Excel 8.0;HDR=YES';Persist Security Info=False" )

		  CASE Lower( Right( aOpenInfo[ UR_OI_NAME ], 4 ) ) == ".dbf"
			 cStr :=   CFILEPATH(aOpenInfo[ UR_OI_NAME ]) 
			 cStr := SUBSTR(cStr,1,LEN(Cstr)-1)
			 cStr := "d:\followup-testes\tESTES FOLLOWUP.add"
//aWAData[ WA_CONNECTION ]:Open(adoconnect()) //trials
			 aWAData[ WA_CONNECTION ]:Open("Provider=Advantage OLE DB Provider;User ID=adssys;Data Source="+cStr+";TableType=ADS_CDX;"+;
				"Advantage Server Type=ADS_LOCAL_SERVER;")
				
//OTHE PROVIDERS FOR DBF FOR TRIALS		
// aWAData[ WA_CONNECTION ]:Open("Provider=vfpoledb.1;Data Source="+cStr+";Collating Sequence=machine;")
// aWAData[ WA_CONNECTION ]:Open( "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + cstr + ";Extended Properties=dBASE IV;User ID=Admin;Password=;" )

		  CASE Lower( Right( aOpenInfo[ UR_OI_NAME ], 3 ) ) == ".db"
			 aWAData[ WA_CONNECTION ]:Open( "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + aOpenInfo[ UR_OI_NAME ] + ";Extended Properties='Paradox 3.x';" )

		  CASE aWAData[ WA_ENGINE ] == "MYSQL"
			 aWAData[ WA_CONNECTION ]:Open( "DRIVER={MySQL ODBC 3.51 Driver};" + ;
				"server=" + aWAData[ WA_SERVER ] + ;
				";database=" + aOpenInfo[ UR_OI_NAME ] + ;
				";uid=" + aWAData[ WA_USERNAME ] + ;
				";pwd=" + aWAData[ WA_PASSWORD ] )

		  CASE aWAData[ WA_ENGINE ] == "SQL"
			 aWAData[ WA_CONNECTION ]:Open( "Provider=SQLOLEDB;" + ;
				"server=" + aWAData[ WA_SERVER ] + ;
				";database=" + aOpenInfo[ UR_OI_NAME ] + ;
				";uid=" + aWAData[ WA_USERNAME ] + ;
				";pwd=" + aWAData[ WA_PASSWORD ] )

		  CASE aWAData[ WA_ENGINE ] == "ORACLE"
			 aWAData[ WA_CONNECTION ]:Open( "Provider=MSDAORA.1;" + ;
				"Persist Security Info=False" + ;
				iif( Empty( aWAData[ WA_SERVER ] ), ;
				"", ";Data source=" + aWAData[ WA_SERVER ] ) + ;
				";User ID=" + aWAData[ WA_USERNAME ] + ;
				";Password=" + aWAData[ WA_PASSWORD ] )

		  CASE aWAData[ WA_ENGINE ] == "FIREBIRD"
			 aWAData[ WA_CONNECTION ]:Open( "Driver=Firebird/InterBase(r) driver;" + ;
				"Persist Security Info=False" + ;
				";Uid=" + aWAData[ WA_USERNAME ] + ;
				";Pwd=" + aWAData[ WA_PASSWORD ] + ;
				";DbName=" + aOpenInfo[ UR_OI_NAME ] )
		  ENDCASE
	  ELSE
	     // ITS ALREDY OPE THE ADODB CONN USE THE SAME WE WANT TRANSACTIONS WITHIN THE CONNECTION
		  aWAData[ WA_CONNECTION ] :=  oConnection
		  aWAData[ WA_TABLENAME ] := t_cTableName
		  aWAData[ WA_QUERY ] := t_cQuery
		  aWAData[ WA_USERNAME ] := t_cUserName
		  aWAData[ WA_PASSWORD ] := t_cPassword
		  aWAData[ WA_SERVER ] := t_cServer
		  aWAData[ WA_ENGINE ] := t_cEngine
		  aWAData[ WA_CONNOPEN ] := .T.
		  aWAData[ WA_OPENSHARED ] := aOpenInfo[ UR_OI_SHARED ]
      ENDIF	  
   ELSE
      // here we dont save oconnection for the next one because
	  // we assume that is not application defult conn but a temporary conn
	  //to other db system.
      aWAData[ WA_CONNECTION ] := TOleAuto():New("ADODB.Connection")
      aWAData[ WA_CONNECTION ]:Open( aOpenInfo[ UR_OI_CONNECT ] )
      aWAData[ WA_TABLENAME ] := t_cTableName
      aWAData[ WA_QUERY ] := t_cQuery
      aWAData[ WA_USERNAME ] := t_cUserName
      aWAData[ WA_PASSWORD ] := t_cPassword
      aWAData[ WA_SERVER ] := t_cServer
      aWAData[ WA_ENGINE ] := t_cEngine
      aWAData[ WA_CONNOPEN ] := .F.
   ENDIF
   /* will be initilized */
   t_cQuery := ""

   IF Empty( aWAData[ WA_QUERY ] )
      aWAData[ WA_QUERY ] := "SELECT * FROM "
   ENDIF

   oRecordSet :=  TOleAuto():New( "ADODB.Recordset" )

   IF oRecordSet == NIL
      oError := ErrorNew()
      oError:GenCode := EG_OPEN
      oError:SubCode := 1001
      oError:Description := hb_langErrMsg( EG_OPEN )
      oError:FileName := aOpenInfo[ UR_OI_NAME ]
      oError:OsCode := 0 /* TODO */
      oError:CanDefault := .T.

      UR_SUPER_ERROR( nWA, oError )
      RETURN HB_FAILURE
   ENDIF

   oRecordSet:CursorType := adOpenDynamic
   oRecordSet:CursorLocation := adUseServer  // adUseClient its slower but has avntages such always bookmaks 
   oRecordSet:LockType := adLockPessimistic
   //oRecordSet:MaxRecords := 100 ?
   //oRecordSet:CacheSize := 50 //records increase performance set zero returns error set great server parameters max open rows error

   aWAData[ WA_TABLENAME ] := SUBSTR(CFILENOPATH(aWAData[ WA_TABLENAME ] ),1,LEN(CFILENOPATH(aWAData[ WA_TABLENAME ] ))-4)

   IF aWAData[ WA_QUERY ] == "SELECT * FROM "
      oRecordSet:Open( aWAData[ WA_QUERY ] + aWAData[ WA_TABLENAME ], aWAData[ WA_CONNECTION ])
	  //to allow prop index and seek  adCmdTableDirect ?
   ELSE
      oRecordSet:Open( aWAData[ WA_QUERY ], aWAData[ WA_CONNECTION ],,,adCmdTableDirect )
   ENDIF
   
   aWAData[ WA_RECORDSET ] := oRecordSet
   aWAData[ WA_BOF ] := aWAData[ WA_EOF ] := .F.
 
   UR_SUPER_SETFIELDEXTENT( nWA, nTotalFields := oRecordSet:Fields:Count )
   
   FOR n := 1 TO nTotalFields
      aField := Array( UR_FI_SIZE )
      aField[ UR_FI_NAME ] := oRecordSet:Fields( n - 1 ):Name
	  aField[ UR_FI_TYPE ]    := ADO_FIELDSTRUCT( oRecordSet, n-1 )[7] 
      aField[ UR_FI_TYPEEXT ] := 0
  	  aField[ UR_FI_LEN ]     := ADO_FIELDSTRUCT( oRecordSet, n-1 )[3]
	  aField[ UR_FI_DEC ]     := ADO_FIELDSTRUCT( oRecordSet, n-1 )[4]
	  
#ifdef __XHARBOUR__	  
      aField[ UR_FI_FLAGS ] := 0  // xHarbour expecs this field 
      aField[ UR_FI_STEP ] := 0 // xHarbour expecs this field
#endif
	  
      UR_SUPER_ADDFIELD( nWA, aField )
   NEXT

   nResult := UR_SUPER_OPEN( nWA, aOpenInfo )

   IF nResult == HB_SUCCESS
      ADO_GOTOP( nWA )
   ENDIF
   
   //OPEN EXCLUSIVE MUST BE WITHIN A TRANSACTION!
   IF !aWAData[ WA_OPENSHARED ] 
   
       IF aWAData[ WA_CONNECTION ]:BeginTrans() > 1  //THERE IS ANOTHER TOP LEVEL TRANSACTION OPEN EXCLUSIVE
	   
  	      nResult := HB_SUCCESS 
		  aWAData[ WA_CONNECTION ]:RollbackTrans() // CLOSE TRANSACT THERE IS NOTHING TO rollback only to cncel it
	   
	   ENDIF
	   
   ENDIF	   
   
   RETURN nResult

FUNCTION ADODB_CLOSE()
 // oConnection STATIC VAR that mantains te adodb connection the same for all recordsets
 //this is to enable transactions in several recordsets because transactions is per connection
 //this it to be called within an exit proc of the application
 // or whnever we dont need it anymore.
 
   IF ! Empty( oConnection )
      IF oConnection:State != adStateClosed
         IF oConnection:State != adStateOpen
            oConnection:Cancel()
         ELSE
            oConnection:Close()
         ENDIF
      ENDIF
  ENDIF
	  
  RETURN .T.
   
STATIC FUNCTION ADO_CLOSE( nWA )

   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]
   
   //dont close connection as mugh be used by other recorsets
   // need to have all recordsets in same connection to use transactions
   IF !EMPTY( oRecordSet)
      IF oRecordSet:State = adStateOpen
         oRecordSet:Close()
      ENDIF
   ENDIF	  
   
   RETURN UR_SUPER_CLOSE( nWA )

STATIC FUNCTION ADO_GETVALUE( nWA, nField, xValue )

   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL rs := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]
   LOCAL aFieldInfo := ADO_FIELDSTRUCT( Rs, nField-1 )
   
   //MISSIGNG OLE VARLEN MODTIME ROWVER CURDUBLE FLOAT LONG AUTOINC CURRENCY BLOB IMAGE
   //DONT KNOW DEFAULT VALUES
   
   IF aWAData[ WA_EOF ] .OR. rs:EOF .OR. rs:BOF
      xValue := NIL
	   
      IF aFieldInfo[7] == HB_FT_STRING
         xValue := Space( aFieldInfo[3] )
      ENDIF
	  
	  IF aFieldInfo[7] == HB_FT_DATE
         xValue := stod("  /  /  ")
      ENDIF
	  
	  IF aFieldInfo[7] == HB_FT_INTEGER .OR.  aFieldInfo[7] ==  HB_FT_DOUBLE
		 IF aFieldInfo[4] > 0
            xValue := VAL("0."+ALLTRIM(STR(aFieldInfo[4],0)))
		 ELSE
		    xValue := 0
         ENDIF			
      ENDIF
	  
  	  IF aFieldInfo[7] == HB_FT_MEMO
         xValue := SPACE(0)
	  ENDIF
	  
	  IF aFieldInfo[7] == HB_FT_LOGICAL 
         xValue := .F.
	  ENDIF
	  
	  IF aFieldInfo[7] == HB_FT_TIMESTAMP 
	     // xValue := what is the defaut type for this?
	  ENDIF

   ELSE

      xValue := rs:Fields( nField - 1 ):Value
	  
	  IF aFieldInfo[7] == HB_FT_STRING
	     IF VALTYPE( xValue ) == "U"
            xValue := SPACE( rs:Fields( nField - 1 ):DefinedSize )
         ELSE
            xValue := PADR( xValue, rs:Fields( nField - 1 ):DefinedSize )
         ENDIF
	  ENDIF
	  
	  IF aFieldInfo[7] == HB_FT_DATE
	     IF VALTYPE( xValue ) == "U"
            xValue := SToD()
         ENDIF
	  ENDIF

	  IF aFieldInfo[7] == HB_FT_INTEGER .OR.  aFieldInfo[7] ==  HB_FT_DOUBLE
	     IF VALTYPE( xValue ) == "U"
		    IF aFieldInfo[4] > 0
               xValue := VAL("0."+ALLTRIM(STR(aFieldInfo[4],0)))
			ELSE
			   xValue := 0
            ENDIF			
         ENDIF
	  ENDIF
	  
	  IF aFieldInfo[7] == HB_FT_MEMO
	     IF VALTYPE( xValue ) == "U"
            xValue := SPACE(0)
         ENDIF
	  ENDIF
	  
	  IF aFieldInfo[7] == HB_FT_LOGICAL 
	     IF VALTYPE( xValue ) == "U"
            xValue := .F.
         ENDIF
	  ENDIF
	  
	  IF aFieldInfo[7] == HB_FT_TIMESTAMP 
	     // xValue := what is the defaut type for this?
	  ENDIF
	  
   ENDIF	  
   
   
   RETURN HB_SUCCESS
   

STATIC FUNCTION ADO_GOTO( nWA, nRecord )

   LOCAL nRecNo
   LOCAL rs := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]

   IF ADORECCOUNT(nWa,rs) <> 0 // AHF SEE FUNCTION FOR EXPLANATION rs:RecordCount <> 0
      rs:MoveFirst()
      rs:Move( nRecord - 1, 0 )
	  //rs:bookmark := nRecord
   ENDIF
   ADO_RECID( nWA, @nRecNo )

   RETURN iif( nRecord == nRecNo, HB_SUCCESS, HB_FAILURE )

STATIC FUNCTION ADO_GOTOID( nWA, nRecord )

   LOCAL nRecNo
   LOCAL rs := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]

   IF ADORECCOUNT(nWa,rs) <> 0 // AHF SEE FUNCTION FOR EXPLANATION rs:RecordCount <> 0
      rs:MoveFirst()
      rs:Move( nRecord - 1, 0 )
   ENDIF
   ADO_RECID( nWA, @nRecNo )

   RETURN iif( nRecord == nRecNo, HB_SUCCESS, HB_FAILURE )

STATIC FUNCTION ADO_GOTOP( nWA )

   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := aWAData[ WA_RECORDSET ]

   IF ADORECCOUNT(nWa,oRecordSet) != 0 // AHF SEE FUNCTION FOR EXPLANATION oRecordSet:RecordCount() != 0
      oRecordSet:MoveFirst()
   ENDIF

   aWAData[ WA_BOF ] := .F.
   aWAData[ WA_EOF ] := .F.

   RETURN HB_SUCCESS

STATIC FUNCTION ADO_GOBOTTOM( nWA )

   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := aWAData[ WA_RECORDSET ]

   oRecordSet:MoveLast()

   aWAData[ WA_BOF ] := .F.
   aWAData[ WA_EOF ] := .F.

   RETURN HB_SUCCESS

STATIC FUNCTION ADO_SKIPRAW( nWA, nToSkip )

   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := aWAData[ WA_RECORDSET ]
   LOCAL nResult := HB_SUCCESS

   IF ADORECCOUNT(nWa,oRecordSet) = 0 // AHF SEE FUNCTION FOR EXPLANATION oRecordSet:RecordCount = 0
      RETURN HB_FAILURE
   ENDIF
   
   IF nToSkip != 0
      IF aWAData[ WA_EOF ]
         IF nToSkip > 0
            RETURN HB_SUCCESS //SHOULDNET BE FAILURE?
         ENDIF
         ADO_GOBOTTOM( nWA )
         ++nToSkip
      ENDIF
      TRY
         IF aWAData[ WA_CONNECTION ]:State != adStateClosed
            IF nToSkip < 0 .AND. oRecordSet:AbsolutePosition <= - nToSkip
               oRecordSet:MoveFirst()
               aWAData[ WA_BOF ] := .T.
               aWAData[ WA_EOF ] := oRecordSet:EOF
            ELSE 
			   IF ADORECCOUNT(nWa,oRecordSet) <> 0 // AHF SEE FUNCTION FOR EXPLANATION oRecordSet:RecordCount <> 0
                  oRecordSet:Move( nToSkip )
                  aWAData[ WA_BOF ] := .F.
                  aWAData[ WA_EOF ] := oRecordSet:EOF
			   ELSE
                  aWAData[ WA_BOF ] := oRecordSet:BOF
                  aWAData[ WA_EOF ] := oRecordSet:EOF
               ENDIF			   
            ENDIF
            //ENFORCE RELATIONS SHOULD BE BELOW AFTER MOVING TO NEXT RECORD
            IF ! Empty( aWAData[ WA_PENDINGREL ] )
               IF ADO_FORCEREL( nWA ) != HB_SUCCESS
                  BREAK
               ENDIF
            ENDIF
         ELSE
            nResult := HB_FAILURE
         ENDIF
      CATCH
         nResult := HB_FAILURE
      END
   ENDIF

   RETURN nResult

STATIC FUNCTION ADO_BOF( nWA, lBof )

   LOCAL aWAData := USRRDD_AREADATA( nWA )

   lBof := aWAData[ WA_BOF ]

   RETURN HB_SUCCESS

STATIC FUNCTION ADO_EOF( nWA, lEof )

   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]
   LOCAL nResult := HB_SUCCESS

   TRY
      IF USRRDD_AREADATA( nWA )[ WA_CONNECTION ]:State != adStateClosed
         lEof := oRecordSet:Eof() //( oRecordSet:AbsolutePosition == -3 )
      ENDIF
   CATCH
      nResult := HB_FAILURE
   END

   RETURN nResult

STATIC FUNCTION ADO_DELETED( nWA, lDeleted )

   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]
   TRY
      IF oRecordSet:Status == adRecDeleted
         lDeleted := .T.
      ELSE
         lDeleted := .F.
      ENDIF
   CATCH
      lDeleted := .F.
   END

   RETURN HB_SUCCESS

STATIC FUNCTION ADO_DELETE( nWA )

   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]

   oRecordSet:Delete()

   ADO_SKIPRAW( nWA, 1 )

   RETURN HB_SUCCESS

STATIC FUNCTION ADO_RECNO( nWA, nRecNo )

   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]
   LOCAL nResult := HB_SUCCESS

   TRY
      IF USRRDD_AREADATA( nWA )[ WA_CONNECTION ]:State != adStateClosed
         nRecno := oRecordSet:BookMark
		 // its not safe see ado docs iif( oRecordSet:AbsolutePosition == -3, oRecordSet:RecordCount() + 1, oRecordSet:AbsolutePosition )
      ELSE
         nRecno := 0
         nResult := HB_FAILURE
      ENDIF
   CATCH
      nRecNo := 0
      nResult := HB_FAILURE
   END

   RETURN nResult

STATIC FUNCTION ADO_RECID( nWA, nRecNo )

   RETURN ADO_RECNO( nWA, @nRecNo )

STATIC FUNCTION ADO_RECCOUNT( nWA, nRecords )
   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]

   nRecords := ADORECCOUNT(nWA,oRecordSet) // AHF SEE FUNCTION FOR EXPLANATION oRecordSet:RecordCount()

   RETURN HB_SUCCESS

STATIC FUNCTION ADORECCOUNT(nWA,oRecordSet) //AHF
   LOCAL aAWData := USRRDD_AREADATA( nWA )
   LOCAL oCon := aAWData[WA_CONNECTION]
   LOCAL nCount := 0, cSql:="",oRs := TOleAuto():New("ADODB.Recordset") //OPEN A NEW ONE OTHERWISE PROBLEMS WITH OPEN BROWSES
   
   IF oRecordSet:CursorLocation == adUseClient
      nCount :=  oRecordSet:RecordCount
   ELSE
      IF LEN(aAWData[WA_INDEXES]) > 0 .AND. aAWData[WA_INDEXACTIVE] > 0
		 //Making it lightning faster
		 oRs:CursorLocation := adUseServer
		 oRs:CursorType := adOpenForwardOnly
		 oRs:LockType := adLockReadOnly
		 //LAST PARAMTER INSERTS cSql COUNT(*) MUST BE ALL FIELDS BECAUSE IF THERE IS A NULL FIELD COUNTS RETURNS WRONG
		 cSql := IndexBuildExp(nWA,aAWData[WA_INDEXACTIVE],aAWData,.T.) 
		 //LETS COUNT IT
		 oRs:open(cSql,oCon)
		 nCount := oRs:Fields( 0 ):Value
		 oRs:close()
	  ELSE
	     nCount :=  oRecordSet:RecordCount
	  ENDIF
   ENDIF   
   
   RETURN nCount   
   
STATIC FUNCTION ADO_PUTVALUE( nWA, nField, xValue )

   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := aWAData[ WA_RECORDSET ]
//MSGINFO("ADO PUTVALUE "+CVALTOCHAR(XVALUE)+ " "+PROCNAME(2)+ " "+CVALTOCHAR(PROCLINE(2) )  )
   IF ! aWAData[ WA_EOF ] .AND. !( oRecordSet:Fields( nField - 1 ):Value == xValue )
      oRecordSet:Fields( nField - 1 ):Value := xValue
      TRY
         oRecordSet:Update()
      CATCH
//MSGINFO("SEM UPDAYE ADO VARPUT")
      END
   ENDIF

   RETURN HB_SUCCESS

STATIC FUNCTION ADO_APPEND( nWA, lUnLockAll )

   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]
   LOCAL aWdata := USRRDD_AREADATA( nWA )
   
   HB_SYMBOL_UNUSED( lUnLockAll )

    oRecordSet:AddNew()
	
	TRY
    oRecordSet:Update()
	CATCH
	END
	
    IF !lUnlockAll
	    ADO_UNLOCK(nWA)
	ENDIF	
    AADD(aWdata[ WA_LOCKLIST ],oRecordSet:BookMark)

   RETURN HB_SUCCESS

STATIC FUNCTION ADO_ORDINFO( nWA, nIndex, aOrderInfo )

   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := aWAData[ WA_RECORDSET ]
   LOCAL nResult := HB_SUCCESS
   LOCAL cExp:="",cIndexExp := ""

   IF EMPTY(aOrderInfo[ UR_ORI_TAG ]) 
      aOrderInfo[ UR_ORI_TAG ] := 0
      IF !EMPTY(aWAData[ WA_INDEXACTIVE ])
         aOrderInfo[ UR_ORI_TAG ] := aWAData[ WA_INDEXACTIVE ]
	  ENDIF	 
   ENDIF	 
       
   DO CASE
   CASE nIndex == DBOI_EXPRESSION

   		IF ! Empty( aWAData[ WA_INDEXEXP ] ) .AND. !Empty( aOrderInfo[ UR_ORI_TAG ] ) .AND.;
           aOrderInfo[ UR_ORI_TAG ] <= len(aWAData[ WA_INDEXEXP ])
           aOrderInfo[ UR_ORI_RESULT ] := aWAData[ WA_INDEXEXP ][aOrderInfo[ UR_ORI_TAG]]
        ELSE
           aOrderInfo[ UR_ORI_RESULT ] := ""
        ENDIF
		//STRIPPING OUT INVALID EXPRESSION FOR DBFI NDEX EXPRESSION
		aOrderInfo[ UR_ORI_RESULT ] := STRTRAN(aOrderInfo[ UR_ORI_RESULT ] , ",","+")
		aOrderInfo[ UR_ORI_RESULT ] := STRTRAN(aOrderInfo[ UR_ORI_RESULT ] , "ASC","")
		aOrderInfo[ UR_ORI_RESULT ] := STRTRAN(aOrderInfo[ UR_ORI_RESULT ] , "DESC","")
		
   CASE nIndex == DBOI_CONDITION	
   
		IF ! Empty( aWAData[ WA_INDEXFOR ] ) .AND. ! Empty( aOrderInfo[ UR_ORI_TAG ] ) .AND. ;
            aOrderInfo[ UR_ORI_TAG ] <= LEN(aWAData[ WA_INDEXFOR ])
            aOrderInfo[ UR_ORI_RESULT ] := aWAData[ WA_INDEXFOR ][aOrderInfo[ UR_ORI_TAG]]
		ELSE
          aOrderInfo[ UR_ORI_RESULT ] :=""
		ENDIF
		//STRIPPING OUT INVALID EXPRESSION FOR DBF INDEX FOR EXPRESSION
		aOrderInfo[ UR_ORI_RESULT ] := STRTRAN(aOrderInfo[ UR_ORI_RESULT ] , "WHERE","FOR")
		
   CASE nIndex == DBOI_NAME
   
        IF ! Empty( aWAData[ WA_INDEXES ] ) .AND. ! Empty( aOrderInfo[ UR_ORI_TAG ] ) .AND. ;
           aOrderInfo[ UR_ORI_TAG ] <= LEN(aWAData[ WA_INDEXES ])
           aOrderInfo[ UR_ORI_RESULT ] := aWAData[ WA_INDEXES ][aOrderInfo[ UR_ORI_TAG]]
         ELSE
            aOrderInfo[ UR_ORI_RESULT ] := ""
        ENDIF
		
   CASE nIndex == DBOI_NUMBER
   
		IF VALTYPE(aOrderInfo[ UR_ORI_TAG ]) = "N"
		   aOrderInfo[ UR_ORI_RESULT ] := aOrderInfo[ UR_ORI_TAG ]
		ELSE   
            n := ASCAN(aWAData[ WA_INDEXES ],aOrderInfo[ UR_ORI_TAG])
		    IF n > 0
               aOrderInfo[ UR_ORI_RESULT ] := n
            ELSE
               aOrderInfo[ UR_ORI_RESULT ] := 0
            ENDIF
		ENDIF	
		
   CASE nIndex == DBOI_BAGNAME
   
        IF ! Empty( aWAData[ WA_INDEXES ] ) .AND. ! Empty( aOrderInfo[ UR_ORI_TAG ] ) .AND. ;
            aOrderInfo[ UR_ORI_TAG ] <= LEN(aWAData[ WA_INDEXES ])
            aOrderInfo[ UR_ORI_RESULT ] := aWAData[ WA_INDEXES ][aOrderInfo[ UR_ORI_TAG]]
        ELSE
           aOrderInfo[ UR_ORI_RESULT ] := ""
        ENDIF
		
   CASE nIndex == DBOI_BAGEXT
   
        aOrderInfo[ UR_ORI_RESULT ] := ""
		
   CASE nIndex == DBOI_ORDERCOUNT
   
        IF ! Empty( aWAData[ WA_INDEXES ] )
           aOrderInfo[ UR_ORI_RESULT ] := LEN(aWAData[ WA_INDEXES ])
        ELSE
           aOrderInfo[ UR_ORI_RESULT ] := 0
        ENDIF
		
   CASE nIndex == DBOI_FILEHANDLE
   
        aOrderInfo[ UR_ORI_RESULT ] := -1
		
   CASE nIndex == DBOI_ISCOND
   
		IF ! Empty( aWAData[ WA_INDEXFOR ] ) .AND. ! Empty( aOrderInfo[ UR_ORI_TAG ] ) .AND. ;
            aOrderInfo[ UR_ORI_TAG ] <= LEN(aWAData[ WA_INDEXFOR ])
            aOrderInfo[ UR_ORI_RESULT ] := EMPTY(aWAData[ WA_INDEXFOR ][aOrderInfo[ UR_ORI_TAG]])
		ELSE
          aOrderInfo[ UR_ORI_RESULT ] :=.F.
		ENDIF
		
   CASE nIndex == DBOI_ISDESC
   
        aOrderInfo[ UR_ORI_RESULT ] :=.F. //ITS REALLY NEVER USED
   
   CASE nIndex == DBOI_UNIQUE
   
		IF ! Empty( aWAData[ WA_INDEXUNIQUE ] ) .AND. ! Empty( aOrderInfo[ UR_ORI_TAG ] ) .AND. ;
            aOrderInfo[ UR_ORI_TAG ] <= LEN(aWAData[ WA_INDEXUNIQUE ])
            aOrderInfo[ UR_ORI_RESULT ] := EMPTY(aWAData[ WA_INDEXUNIQUE ][aOrderInfo[ UR_ORI_TAG]])
		ELSE
          aOrderInfo[ UR_ORI_RESULT ] :=.F.
		ENDIF
		
   CASE nIndex == DBOI_POSITION
   
        IF aWAData[ WA_CONNECTION ]:State != adStateClosed
           ADO_RECID( nWA, @aOrderInfo[ UR_ORI_RESULT ] )
        ELSE
           aOrderInfo[ UR_ORI_RESULT ] := 0
           nResult := HB_FAILURE
        ENDIF
		
   CASE nIndex == DBOI_RECNO
   
        IF aWAData[ WA_CONNECTION ]:State != adStateClosed
           ADO_RECID( nWA, @aOrderInfo[ UR_ORI_RESULT ] )
        ELSE
           aOrderInfo[ UR_ORI_RESULT ] := 0
           nResult := HB_FAILURE
        ENDIF
		
   CASE nIndex == DBOI_KEYCOUNT
   
        IF aWAData[ WA_CONNECTION ]:State != adStateClosed
           aOrderInfo[ UR_ORI_RESULT ] := ADORECCOUNT(nWA,oRecordSet) // AHF SEE FUNCTION FOR EXPLANATION oRecordSet:RecordCount
        ELSE
           aOrderInfo[ UR_ORI_RESULT ] := 0
           nResult := HB_FAILURE
        ENDIF
		
   CASE nIndex == DBOI_SCOPESET .OR. nIndex == DBOI_SCOPEBOTTOM .OR. nIndex == DBOI_SCOPEBOTTOMCLEAR ;
	    .OR. nIndex == DBOI_SCOPECLEAR .OR. nIndex == DBOI_SCOPETOP .OR. nIndex == DBOI_SCOPETOPCLEAR

	    aOrderInfo[ UR_ORI_RESULT ] := ADOSCOPE(nWA, AWAData,oRecordset, aOrderInfo,nIndex)
	 
   ENDCASE
   
   RETURN nResult
   
STATIC FUNCTION ADOSCOPE(nWA,aWAdata, oRecordSet, aOrderInfo,nIndex)
 LOCAL y, cScopeExp :="", cSql :=""

   //[UR_ORI_NEWVAL] comes with actual scope top or bottom and returns the former active scope if any
   IF VALTYPE(aOrderInfo[ UR_ORI_NEWVAL ]) = "B"
      aOrderInfo[ UR_ORI_NEWVAL ] := EVAL(aOrderInfo[ UR_ORI_NEWVAL ])
   ENDIF
   
   //SET SCOPE TO NO ARGS
   IF aOrderInfo[ UR_ORI_NEWVAL ] = NIL
      aOrderInfo[ UR_ORI_NEWVAL ] := ""
   ENDIF
   
   IF EMPTY(aWAdata[WA_INDEXACTIVE]) .OR. aWAdata[WA_INDEXACTIVE] = 0 //NO INDEX NO SCOPE
      aOrderInfo[ UR_ORI_RESULT ] := NIL
      RETURN HB_FAILURE
   ENDIF
   
   y:=ASCAN( aWAData[ WA_SCOPES ], aWAData[WA_INDEXACTIVE]  )

   DO CASE
   CASE nIndex == DBOI_SCOPESET //never gets called noy tested might be completly wrong!
 
       IF y > 0
		   aWAData[ WA_SCOPETOP ][y] := aOrderInfo[ UR_ORI_NEWVAL ]
		   aWAData[ WA_SCOPEBOT ][y] := aOrderInfo[ UR_ORI_NEWVAL ]
		ELSE
	       AADD( aWAData[ WA_SCOPES ],aWAData[ WA_INDEXACTIVE ])
		   AADD(aWAData[ WA_SCOPETOP ],aOrderInfo[ UR_ORI_NEWVAL ])
		   AADD(aWAData[ WA_SCOPEBOT ],aOrderInfo[ UR_ORI_NEWVAL ])
        ENDIF		
        aOrderInfo[ UR_ORI_RESULT ] := NIL

   CASE nIndex == DBOI_SCOPECLEAR //never gets called noy tested might be completly wrong!
   
        IF y > 0
		   ADEL(aWAData[ WA_SCOPES ],y,.T.)
		   ADEL(aWAData[ WA_SCOPETOP ],y,.T.)
		   ADEL(aWAData[ WA_SCOPEBOT ],y,.T.)
        ENDIF		
		aOrderInfo[ UR_ORI_RESULT ] := NIL //RETURN ACUTAL SCOPETOP NIL IF NONE

   CASE nIndex == DBOI_SCOPETOP

        IF y > 0
		   aOrderInfo[ UR_ORI_RESULT ] := aWAData[ WA_SCOPETOP ][y] //RETURN ACTUALSCOPE TOP
		   aWAData[ WA_SCOPETOP ][y] := aOrderInfo[ UR_ORI_NEWVAL ]
		   IF LEN(aWAData[ WA_SCOPEBOT ]) < y
		      AADD(aWAData[ WA_SCOPEBOT ],SPACE(LEN(CVALTOCHAR(aWAData[ WA_SCOPETOP ][y])))) //THERE INST STILL A SCOPEBOT ARRAYS MUST HAVE  SAME LEN
		   ENDIF	  
		ELSE
		   AADD(aWAData[ WA_SCOPETOP ],aOrderInfo[ UR_ORI_NEWVAL ])
		   AADD(aWAData[ WA_SCOPEBOT ],SPACE(LEN(CVALTOCHAR(aWAData[ WA_SCOPETOP ][1])))) //THERE INST STILL A SCOPEBOT ARRAYS MUST HAVE  SAME LEN
		   aOrderInfo[ UR_ORI_RESULT ] := ""
        ENDIF		
	 
   CASE nIndex == DBOI_SCOPEBOTTOM

        IF y > 0
		   aOrderInfo[ UR_ORI_RESULT ] := aWAData[ WA_SCOPEBOT ][y] //RETURN ACTUALSCOPE TOP
		   aWAData[ WA_SCOPEBOT ][y] := aOrderInfo[ UR_ORI_NEWVAL ]
   		   IF LEN(aWAData[ WA_SCOPETOP ]) < y
		      AADD(aWAData[ WA_SCOPETOP ],SPACE(LEN(CVALTOCHAR(aWAData[ WA_SCOPEBOT ][y])))) //THERE INST STILL A SCOPETOP ARRAYS MUST HAVE  SAME LEN
		   ENDIF	  
		ELSE
	       AADD( aWAData[ WA_SCOPES ],aWAData[ WA_INDEXACTIVE ])
		   AADD(aWAData[ WA_SCOPEBOT ],aOrderInfo[ UR_ORI_NEWVAL ])
		   AADD(aWAData[ WA_SCOPETOP ],SPACE(LEN(CVALTOCHAR(aWAData[ WA_SCOPEBOT ][1])))) //THERE INST STILL A SCOPETOP ARRAYS MUST HAVE  SAME LEN
		   aOrderInfo[ UR_ORI_RESULT ] := ""
        ENDIF		

   CASE nIndex == DBOI_SCOPETOPCLEAR
   
       IF y > 0
		   aOrderInfo[ UR_ORI_RESULT ] := aWAData[ WA_SCOPETOP ][y] //RETURN ACTUALSCOPE TOP
		   aWAData[ WA_SCOPETOP ][y] := SPACE(LEN(CVALTOCHAR(aWAData[ WA_SCOPEBOT ][y])))
		ELSE
		   aOrderInfo[ UR_ORI_RESULT ] := "" //RETURN ACTUALSCOPE TOP IF NONE
        ENDIF		
		
   CASE nIndex == DBOI_SCOPEBOTTOMCLEAR
   
       IF y > 0
		   aOrderInfo[ UR_ORI_RESULT ] := aWAData[ WA_SCOPEBOT ][y] //RETURN ACTUALSCOPE TOP
		   aWAData[ WA_SCOPEBOT ][y] := SPACE(LEN(CVALTOCHAR(aWAData[ WA_SCOPETOP ][y])))
		ELSE
		   aOrderInfo[ UR_ORI_RESULT ] := "" //RETURN ACTUALSCOPE TOP IF NONE
        ENDIF		

   ENDCASE

   //ONLY BUILDS QUERY AFTER ALL DONE ASSUME THAT ALWAYS CLLED IN PAIRS OTHERWISE WILL GET ERROR   
   IF nIndex = DBOI_SCOPEBOTTOM  .OR. nIndex = DBOI_SCOPEBOTTOMCLEAR 
   
      IF y = 0  //IF DIDNT FOUND ANY ITS THE FIRST ONE THAT JUST BEEN ADD
	     y := 1
	  ENDIF	 
	  
	  IF y <= LEN(aWAData[ WA_SCOPES ])  //EXIST SCOPE ARRAY ALREADY
         IF LEN(ALLTRIM(aWAData[ WA_SCOPETOP ][y]+aWAData[ WA_SCOPEBOT ][y])) > 0
            cScopeEXp := ADOPSEUDOSEEK(nWA,aWAData[ WA_SCOPETOP ][y],aWAData,,.T.,aWAData[ WA_SCOPEBOT ][y])[2]
	    ELSE
	       cScopeExp :=""
        ENDIF	  
     ELSE 
	    cScopeExp :=""
	 ENDIF
	 
      cSql := IndexBuildExp(nWA,aWAData[ WA_INDEXACTIVE ],aWAData,,cScopeExp)
	  oRecordSet:Close()
	  oRecordSet:open(cSql,aWAData[ WA_CONNECTION ])
	  
   ENDIF
   
  RETURN HB_SUCCESS
  
  
STATIC FUNCTION ADO_FIELDNAME( nWA, nField, cFieldName )

   LOCAL nResult := HB_SUCCESS
   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]

   TRY
      cFieldName := oRecordSet:Fields( nField - 1 ):Name
   CATCH
      cFieldName := ""
      nResult := HB_FAILURE
   END

   RETURN nResult

STATIC FUNCTION ADO_FIELDINFO( nWA, nField, nInfoType, uInfo )

   LOCAL nType, nLen
   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]
   LOCAL aFieldInfo := ADO_FIELDSTRUCT( oRecordSet, nField-1 )
   
   DO CASE
   CASE nInfoType == DBS_NAME
   
      uInfo := aFieldInfo[1]
	  
   CASE nInfoType == DBS_TYPE
   
      uInfo := aFieldInfo[2]
	  nType := aFieldInfo[7]
	  
      DO CASE
      CASE nType == HB_FT_STRING
         uInfo := "C"
      CASE nType == HB_FT_LOGICAL
         uInfo := "L"
      CASE nType == HB_FT_MEMO
         uInfo := "M"
      CASE nType == HB_FT_OLE
         uInfo := "G"
#ifdef HB_FT_PICTURE
      CASE nType == HB_FT_PICTURE
         uInfo := "P"
#endif
      CASE nType == HB_FT_ANY
         uInfo := "V"
      CASE nType == HB_FT_DATE
         uInfo := "D"
#ifdef HB_FT_DATETIME
      CASE nType == HB_FT_DATETIME
         uInfo := "T"
#endif
      CASE nType == HB_FT_TIMESTAMP
         uInfo := "@"
      CASE nType == HB_FT_LONG
         uInfo := "N"
      CASE nType == HB_FT_INTEGER
         uInfo := "I"
      CASE nType == HB_FT_DOUBLE
         uInfo := "B"
      OTHERWISE
         uInfo := "U"
      ENDCASE
   
   CASE nInfoType == DBS_LEN
	  
        uInfo := aFieldInfo[3]

   CASE nInfoType == DBS_DEC
       
	    uInfo := aFieldInfo[4]
	   
#ifdef DBS_FLAG
   CASE nInfoType == DBS_FLAG
      uInfo := 0
#endif
#ifdef DBS_STEP
   CASE nInfoType == DBS_STEP
      uInfo := 0
#endif
   OTHERWISE
      RETURN HB_FAILURE
   ENDCASE

   RETURN HB_SUCCESS

STATIC FUNCTION ADO_ORDLSTFOCUS( nWA, aOrderInfo )

   LOCAL nRecNo
   LOCAL aWAData    := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := aWAData[ WA_RECORDSET ]
   LOCAL cSql:="" ,n
   
   HB_SYMBOL_UNUSED( nWA )
   HB_SYMBOL_UNUSED( aOrderInfo )

   IF !EMPTY(aOrderInfo[ UR_ORI_TAG ])
   
      oRecordSet:Close()
	  /* AHF NOT NEEDED ONLY IF YOU WANT TO CHANGE IT OTHERWISE STAYS AS IT WAS WHEN OPENING IT
      oRecordSet:CursorType := adOpenDynamic
      oRecordSet:CursorLocation := adUseServer //adUseClient never use ths very slow!
      oRecordSet:LockType := adLockPessimistic
	  */
      IF VALTYPE(aOrderInfo[ UR_ORI_TAG ]) = "C"
         n := ASCAN(aWAData[ WA_INDEXES ],UPPER(aOrderInfo[ UR_ORI_TAG ]))
	  ELSE
         n := aOrderInfo[ UR_ORI_TAG ]  
	  ENDIF   
	  
      IF n = 0
         oRecordSet:Open( aWAData[ WA_TABLENAME ], aWAData[ WA_CONNECTION ])
      ELSE
	     cSql := IndexBuildExp(nWA,n,aWAData)
		 oRecordSet:Open( cSql,aWAData[ WA_CONNECTION ])
      ENDIF
      ADO_GOTOP( nWA )
      //ADO_GOTO( nWA, nRecNo )
	  aWAData[WA_ISITSUBSET] := .F.
      aOrderInfo[ UR_ORI_RESULT ] := aWAData[ WA_INDEXES ] [aWAData[ WA_INDEXACTIVE ]]
	  aWAData[ WA_INDEXACTIVE ] := n

	ELSE
	   aOrderInfo[ UR_ORI_RESULT ] := aWAData[ WA_INDEXES ] [aWAData[ WA_INDEXACTIVE ]]
	ENDIF  
	

   RETURN HB_SUCCESS

STATIC FUNCTION IndexBuildExp(nWA,nIndex,aWAData,lCountRec,myCfor)  //notgroup for adoreccount
   LOCAL cSql := "", cOrder:="", cUnique:="", cFor:=""
   
     DEFAULT lCountRec TO .F.
	 DEFAULT myCfor TO "" //when it comes ex from ado_seek to add to where clause 
	 
	 IF !lCountRec
	    cOrder := (nWA)->(ORDKEY(nIndex))
	    cOrder := " ORDER BY "+cOrder
	 ENDIF
	 
	 IF  nIndex <= LEN(aWAData[ WA_INDEXUNIQUE ])
		cUnique  := aWAData[ WA_INDEXUNIQUE ][nIndex ]+IF(lCountRec, " COUNT(*) ","*")
	 ELSE
        IF lCountRec
		   cUnique := " COUNT(*) "
        ENDIF		
	 ENDIF
	 
	 IF EMPTY(cUnique)
	    cUnique := "*"
	 ENDIF	
	 
     IF  nIndex <= LEN(aWAData[ WA_INDEXFOR ]) 
	     cFor  := " "+aWAData[ WA_INDEXFOR ][ nIndex ]
	 ENDIF
	 
     IF !EMPTY(mycFor)
	    cFor += IF(!EMPTY(cFor)," AND "," WHERE ")+mycFor
	 ENDIF
	 
	 cSql := "SELECT "+ cUnique+" FROM " + aWAData[ WA_TABLENAME ]+ IF(!EMPTY(cFor),cFor,"")+ cOrder
	
   RETURN cSql   
   
STATIC FUNCTION ADO_RAWLOCK( nWA, nAction, nRecNo )

// LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]

   /* TODO WHAT IS THIS FOR?*/

   HB_SYMBOL_UNUSED( nRecNo )
   HB_SYMBOL_UNUSED( nWA )
   HB_SYMBOL_UNUSED( nAction )

   RETURN HB_SUCCESS

STATIC FUNCTION ADO_LOCK( nWA, aLockInfo )

   LOCAL aWdata := USRRDD_AREADATA( nWA ),n
   
   HB_SYMBOL_UNUSED( nWA )
   
   IF EMPTY(aLockInfo[ UR_LI_RECORD ]) .OR. aLockInfo[ UR_LI_METHOD ] = DBLM_EXCLUSIVE
      ADO_UNLOCK( nWA, aLockInfo[ UR_LI_RECORD ] )
   ENDIF   

   /*
   UR_LI_METHOD VALUES CONSTANTS
   DBLM_EXCLUSIVE 1 RELEASE ALL AND LOCK CURRENT 
   DBLM_MULTIPLE 2 LOCK CURRENT AND ADD TO LOCKLIST
   DBLM_FILE 3 RELEASE ALL LOCKS AND FILE LOCK 
   */
   
   IF aLockInfo[UR_LI_METHOD] = DBLM_FILE
      aLockInfo[ UR_LI_RESULT ] := .T.
	  aWdata[ WA_FILELOCK ] := .T.
   ELSE
      aLockInfo[ UR_LI_RECORD ] := RecNo()
      aLockInfo[ UR_LI_RESULT ] := .T.
      AADD(aWdata[ WA_LOCKLIST ],aLockInfo[ UR_LI_RECORD ])
	ENDIF  
   
    ADOBEGINTRANS(nWa)	//START TRANSACTION
	
   RETURN HB_SUCCESS

   
STATIC FUNCTION ADO_UNLOCK( nWA, xRecID )

   LOCAL aWdata := USRRDD_AREADATA( nWA ),n

   HB_SYMBOL_UNUSED( xRecId )
   HB_SYMBOL_UNUSED( nWA )
   
   IF !EMPTY(xRecID)
      n := ASCAN(aWdata[ WA_LOCKLIST ],xRecID)
	  IF n > 0
	     ADEL(aWdata[ WA_LOCKLIST ],n,.T.)
	  ENDIF
   ELSE
      aWdata[ WA_LOCKLIST ] := {}
	  aWdata[ WA_FILELOCK ] := .F.
   ENDIF   
   
   RETURN HB_SUCCESS

   
STATIC FUNCTION ADO_FLUSH( nWA )

   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]
   LOCAL n
   
   TRY
   
      oRecordSet:Update()
	  
	  ADOCOMMITTRANS(nWa) //COMMIT TRANSACTION
	  
   CATCH
   
      ADOROLLBACKTRANS(nWa) //ROLL IT BACK
	  
   END

   RETURN HB_SUCCESS
   

FUNCTION ADOBEGINTRANS(nWa) // narea or calias

 LOCAL oCon := hb_adoRddGetConnection( nWA )
 
  IF oCon:BeginTrans() > 1 //WE ARE AREADY IN A TRANSACT
     oCon:RollbackTrans() // CLOSE TRANSACT STARTED ABOVE THERE IS NOTHING TO ROLLBACK
  ENDIF
  
RETURN .T.

FUNCTION ADOCOMMITTRANS(nWa) // narea or calias

 LOCAL oCon := hb_adoRddGetConnection( nWA ), n
 
  n := oCon:BeginTrans() 
  
  IF n > 1  //WE ARE AREADY IN ATRANSACT
     oCon:RollbackTrans() // CLOSE TRANSACT STARTED ABOVE THERE IS NOTHING TO ROLLBACK
  ENDIF
  IF n > 1 // THERE IS A ACTIVE TRANSACTION IF IT WAS 1 WE WERE NOW OPENING A TRANACT
     oCon:CommitTrans()
  ENDIF
  
RETURN .T.

FUNCTION ADOROLLBACKTRANS(nWa) // narea or calias

 LOCAL oCon := hb_adoRddGetConnection( nWA ), n
 
  n := oCon:BeginTrans() 
  
  IF n > 1  //WE ARE AREADY IN ATRANSACT
     oCon:RollbackTrans() // CLOSE TRANSACT STARTED ABOVE THERE IS NOTHING TO ROLLBACK
  ENDIF
  IF n > 1 // THERE IS A ACTIVE TRANSACTION IF IT WAS 1 WE WERE NOW OPENING A TRANACT
     oCon:RollbackTrans()
  ENDIF

RETURN .T.


STATIC FUNCTION ADO_SETLOCATE( nWA, aScopeInfo )

   LOCAL aWAData := USRRDD_AREADATA( nWA )

   aScopeInfo[ UR_SI_CFOR ] := SQLTranslate( aWAData[ WA_LOCATEFOR ] )

   aWAData[ WA_SCOPEINFO ] := aScopeInfo

   RETURN HB_SUCCESS

STATIC FUNCTION ADO_LOCATE( nWA, lContinue )

   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := aWAData[ WA_RECORDSET ]
   
   oRecordSet:Find( aWAData[ WA_SCOPEINFO ][ UR_SI_CFOR ], iif( lContinue, 1, 0 ) )

   aWAData[ WA_FOUND ] := ! oRecordSet:EOF
   aWAData[ WA_EOF ] := oRecordSet:EOF

   RETURN HB_SUCCESS

STATIC FUNCTION ADO_ORDLSTADD( nWA, aOrderInfo )
   LOCAL cTablename := USRRDD_AREADATA( nWA )[ WA_TABLENAME ]
   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL aFiles := ListIndex() //this can be build as callback function that the programer calls in a ini proc
   LOCAL aTempFiles := ListTmpNames() //this can be build as callback function that the programer calls in a ini proc
   LOCAL cExpress := "" ,cFor:="",cUnique:="",y,z
   LOCAL aTmpIndx := ListTmpIndex() 
   LOCAL aTmpExp := ListTmpExp()
   LOCAL aTmpFor := ListTmpFor()
   LOCAL aTmpUnique := ListTmpUnique()
   LOCAL cIndex , nMax ,cOrder
   
    //ATTENTION DOES NOT VERIFY IF FIELDS EXPESSION MATCH THE TABLE FIELDS
	//ADO WIL GENERATE AN ERROR OR CRASH IF SELECT FIELDS THAT NOT EXIST ON THE TABLE

	//MAYBE IT COMES WITH FILE EXTENSION AND PATH
	cOrder := CFILENOPATH(aOrderInfo[UR_ORI_BAG])
	cOrder := UPPER(CFILENOEXT(cOrder))
	
    //TMP FILES NOT PRESENT IN ListIndex ADDED TO THEIR OWN ARRAY FOR THE DURATION OF THE APP
    IF ASCAN(aTempFiles,UPPER(SUBSTR(cOrder,1,3)) ) > 0 .OR. ASCAN(aTempFiles,UPPER(SUBSTR(cOrder,1,4)) ) > 0
	   //it was added to the array by ado_ordcreate we have only to set focus
	   cIndex := cOrder //aOrderInfo[UR_ORI_BAG] CAN NOT CONTAIN PATH OR FILESXT
	   y := ASCAN(aTmpIndx,cIndex)
	   AADD( aWAData[WA_INDEXES],cIndex)
	   AADD( aWAData[WA_INDEXEXP],aTmpExp[y])
       AADD( aWAData[WA_INDEXFOR],"WHERE "+aTmpFor[y])
       AADD( aWAData[WA_INDEXUNIQUE],aTmpUnique[y])
	   aWAData[WA_INDEXACTIVE] := 1 //always qst one
	   aOrderInfo[UR_ORI_TAG] := 1 //1
  
	   ADO_ORDLSTFOCUS( nWA, aOrderInfo )
	   RETURN HB_SUCCESS
	ENDIF

	//index files present in the index not temp indexes
    y:=ASCAN( aFiles, { |z| z[1] == cTablename } )
    IF y >0
	   nMax := LEN(aFiles[y])-1
	   FOR z :=1 TO LEN( aFiles[y]) -1
		   IF aFiles[y,z+1,1] == cOrder //aOrderInfo[UR_ORI_BAG] CAN NOT CONTAIN PATH OR FILESXT
		      cIndex := aFiles[y,z+1,1]
		      cExpress:=aFiles[y,z+1,2]
			  IF LEN(aFiles[y,z+1]) >= 3 //FOR CONDITION IS PRESENT?
			     cFor := aFiles[y,z+1,3]
			  ENDIF	 
			  IF LEN(aFiles[y,z+1]) >= 4 //UNIQUE CONDITION IS PRESENT?
			     cUnique := aFiles[y,z+1,4]
			  ENDIF	 
			  EXIT
		   ENDIF	  
	   NEXT
	ELSE
       nMax := 1	
	ENDIF

	IF EMPTY(cIndex) //maybe should generate error
	   RETURN HB_FAILURE
    ENDIF
	
	AADD( aWAData[WA_INDEXES],UPPER(cIndex))
	AADD( aWAData[WA_INDEXEXP],UPPER(cExpress))
	AADD( aWAData[WA_INDEXFOR],UPPER(cFor))
	AADD( aWAData[WA_INDEXUNIQUE],UPPER(cUnique))
	aWAData[WA_INDEXACTIVE] := 1 //always qst one
	
	IF z = nMax //all indexes opened for ths table set focus and build select based on the 1st one
	   aOrderInfo[UR_ORI_TAG] := 1
	   ADO_ORDLSTFOCUS( nWA, aOrderInfo )
    ENDIF
	
   RETURN HB_SUCCESS
 
STATIC FUNCTION ADO_ORDLSTCLEAR( nWA )

   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]
   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL nRecNo
   LOCAL n
   
   aWAData[WA_INDEXES]  := {}
   aWAData[WA_INDEXEXP] := {}
   aWAData[WA_INDEXFOR] := {}
   aWAData[WA_INDEXACTIVE] := 0
   aWAData[WA_INDEXUNIQUE] := {}
   aWAData[WA_SCOPES] := {}
   aWAData[WA_SCOPETOP] := {}
   aWAData[WA_SCOPEBOT] := {}
   aWAData[WA_ISITSUBSET] := .F.
   
   ADO_RECID( nWA, @nRecNo )
   oRecordSet:Close()
   /* AHF NOT NEEDED ONLY IF YOU WANT TO CHANGE IT OTHERWISE STAYS AS IT WAS WHEN OPENING IT
   oRecordSet:CursorType := adOpenDynamic
   oRecordSet:CursorLocation := adUseServer //adUseClient
   oRecordSet:LockType := adLockPessimistic
   */
   oRecordSet:Open( aWAData[ WA_TABLENAME ], aWAData[ WA_CONNECTION ])
   ADO_GOTOP( nWA )
   ADO_GOTO( nWA, nRecNo )

   RETURN HB_SUCCESS

STATIC FUNCTION ADO_ORDCREATE( nWA, aOrderCreateInfo )

   LOCAL cTablename := USRRDD_AREADATA( nWA )[ WA_TABLENAME ]
   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL acondinfo := aOrderCreateInfo[UR_ORCR_CONDINFO]
   LOCAL aOrderInfo := ARRAY(UR_ORI_SIZE)
   LOCAL cIndex := UPPER(aOrderCreateInfo[UR_ORCR_BAGNAME])
   LOCAL aTempFiles := ListTmpNames() 
   LOCAL aTmpIndx := ListTmpIndex() 
   LOCAL aTmpExp := ListTmpExp()
   LOCAL aTmpFor := ListTmpFor()
   LOCAL aTmpUnique := ListTmpUnique() , N ,AA:={}
   
   //TMP FILES NOT PRESENT IN ListIndex
   IF ASCAN(aTempFiles,(UPPER(SUBSTR(cIndex,1,3)) )) > 0 .OR. ASCAN(aTempFiles,UPPER(SUBSTR(cIndex,1,4)) ) > 0
      y := 1
      DO WHILE ASCAN( aTmpIndx,cIndex) > 0 //no other with same name
	     y++
		 cIndex += ALLTRIM(STR(n))
	  ENDDO
   ELSE
      IF ASCAN( aWAData[WA_INDEXES],cIndex) > 0
	     // BUILD ERROR
	  ENDIF
   ENDIF
 
    AADD(aTmpIndx,UPPER(cIndex))
	AADD(aTmpExp,UPPER(STRTRAN(aOrderCreateInfo[UR_ORCR_CKEY],"+",",")) )
	AADD(aTmpFor,UPPER(STRTRAN(acondinfo[UR_ORC_CFOR],'"',"'")) )//CLEAN THE DOT .AND. .OR.
	IF acondinfo[UR_ORC_NEXT] > 0
	   AADD(aTmpUnique,UPPER(acondinfo[UR_ORC_NEXT] ))
    ELSE
	   AADD(aTmpUnique,"")
	ENDIF
	
    aOrderInfo [UR_ORI_BAG ] := cIndex
    aOrderInfo [UR_ORI_TAG ] := cIndex
	
	AADD( aWAData[WA_INDEXES],UPPER(cIndex))
	AADD( aWAData[WA_INDEXEXP],UPPER(STRTRAN(aOrderCreateInfo[UR_ORCR_CKEY],"+",",")))
	IF !EMPTY(acondinfo[UR_ORC_CFOR])
	   AADD( aWAData[WA_INDEXFOR]," WHERE "+UPPER(STRTRAN(acondinfo[UR_ORC_CFOR],'"',"'")))
    ELSE
	   AADD( aWAData[WA_INDEXFOR],"")
	ENDIF
	IF acondinfo[UR_ORC_NEXT] > 0
       AADD( aWAData[WA_INDEXUNIQUE]," TOP "+UPPER(ALLTRIM(STR(acondinfo[UR_ORC_NEXT]))))
	ELSE
       AADD( aWAData[WA_INDEXUNIQUE],"")	
	ENDIF
	
   RETURN HB_SUCCESS

STATIC FUNCTION ADO_ORDDESTROY( nWA, aOrderInfo )

   LOCAL aWAData := USRRDD_AREADATA( nWA ), n

   n:= ASCAN(aWAData[ WA_INDEXES ],aOrderInfo[ UR_ORI_TAG ])
   
   IF n > 0
      ADEL( aWAData[ WA_INDEXES ], n, .T.)
	  ADEL( aWAData[ WA_INDEXEXP ], n, .T.)
	  ADEL( aWAData[ WA_INDEXFOR ], n, .T.)
	  IF n = aWAData[ WA_INDEXACTIVE ]
	     aWAData[ WA_INDEXACTIVE ] := 0
	  ENDIF
   ENDIF

   RETURN HB_SUCCESS

STATIC FUNCTION ADO_SEEK( nWA, lSoftSeek, cKey, lFindLast )
   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]
   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL aSeek,cSql
   
   HB_SYMBOL_UNUSED( nWA )
   HB_SYMBOL_UNUSED( lSoftSeek )
   HB_SYMBOL_UNUSED( cKey )
   HB_SYMBOL_UNUSED( lFindLast )
 
   IF oRecordSet:EOF .OR. oRecordSet:BOF .OR.  aWAData[ WA_EOF ] .OR. aWAData[ WA_BOF ]
      RETURN HB_FAILURE
   ENDIF
   
   //this tell us if we are in a subset of records from previous seek
   IF aWAData[WA_ISITSUBSET]
      // ISNT WORKING STILL ADO_ORDLSTFOCUS(nWa,aWAData[INDEXACTIVE])
   ENDIF
   
   aSeek := ADOPseudoSeek(nWA,cKey,aWAData,lSoftSeek)
   
   IF aSeek[3] //no more than one field in the expression we can use find
      IF lSoftSeek 
         oRecordSet:MoveLast()
         oRecordSet:Find( aSeek[1],,adSearchBackward)
		 IF !oRecordSet:Eof()
		    oRecordSet:Move(1)
		 ENDIF	
	  ELSE
	     IF lFindLast
            oRecordSet:MoveLast()
            oRecordSet:Find( aSeek[1])
		 ELSE
            oRecordSet:MoveFirst()
            oRecordSet:Find( aSeek[1])
		 ENDIF	
	  ENDIF
   ELSE
       //attention multiple fields in cseek expression cannot emulate behaviour of lSoftSeek 
	  //more than one field in te seek expression has to be select
	  oRecordSet:Close()
      cSql := IndexBuildExp(nWA,aWAData[WA_INDEXACTIVE],aWAData,.F.,aSeek[2])
      oRecordSet:Open( cSql,aWAData[ WA_CONNECTION ] )

	  IF lFindLast
	     oRecordSet:MoveLast()
	  ENDIF
	  //TO CHECK NEXT CALLS IF WE ARE IN A SUBSSET TO REVERT TO DEFAULT SET
	  aWAData[WA_ISITSUBSET] := .T.
	  
   ENDIF	  
   
   aWAData[ WA_FOUND ] := ! oRecordSet:EOF
   aWAData[ WA_EOF ] := oRecordSet:EOF
   
   RETURN HB_SUCCESS

STATIC FUNCTION ADO_FOUND( nWA, lFound )

   LOCAL aWAData := USRRDD_AREADATA( nWA )
   lFound := aWAData[ WA_FOUND ]

   RETURN HB_SUCCESS

FUNCTION ADORDD_GETFUNCTABLE( pFuncCount, pFuncTable, pSuperTable, nRddID )

   LOCAL aADOFunc[ UR_METHODCOUNT ]

   aADOFunc[ UR_INIT ]         := (@ADO_INIT())
   aADOFunc[ UR_INFO ]         := (@ADO_INFO())
   aADOFunc[ UR_NEW ]          := (@ADO_NEW())
   aADOFunc[ UR_CREATE ]       := (@ADO_CREATE())
   aADOFunc[ UR_CREATEFIELDS ] := (@ADO_CREATEFIELDS())
   aADOFunc[ UR_OPEN ]         := (@ADO_OPEN())
   aADOFunc[ UR_CLOSE ]        := (@ADO_CLOSE())
   aADOFunc[ UR_BOF  ]         := (@ADO_BOF())
   aADOFunc[ UR_EOF  ]         := (@ADO_EOF())
   aADOFunc[ UR_DELETED ]      := (@ADO_DELETED())
   aADOFunc[ UR_SKIPRAW ]      := (@ADO_SKIPRAW())
   aADOFunc[ UR_GOTO ]         := (@ADO_GOTO())
   aADOFunc[ UR_GOTOID ]       := (@ADO_GOTOID())
   aADOFunc[ UR_GOTOP ]        := (@ADO_GOTOP())
   aADOFunc[ UR_GOBOTTOM ]     := (@ADO_GOBOTTOM())
   aADOFunc[ UR_RECNO ]        := (@ADO_RECNO())
   aADOFunc[ UR_RECID ]        := (@ADO_RECID())
   aADOFunc[ UR_RECCOUNT ]     := (@ADO_RECCOUNT())
   aADOFunc[ UR_GETVALUE ]     := (@ADO_GETVALUE())
   aADOFunc[ UR_PUTVALUE ]     := (@ADO_PUTVALUE())
   aADOFunc[ UR_DELETE ]       := (@ADO_DELETE())
   aADOFunc[ UR_APPEND ]       := (@ADO_APPEND())
   aADOFunc[ UR_FLUSH ]        := (@ADO_FLUSH())
   aADOFunc[ UR_ORDINFO ]      := (@ADO_ORDINFO())
   aADOFunc[ UR_RECINFO ]      := (@ADO_RECINFO())
   aADOFunc[ UR_FIELDINFO ]    := (@ADO_FIELDINFO())
   aADOFunc[ UR_FIELDNAME ]    := (@ADO_FIELDNAME())
   aADOFunc[ UR_ORDLSTFOCUS ]  := (@ADO_ORDLSTFOCUS())
   aADOFunc[ UR_PACK ]         := (@ADO_PACK())
   aADOFunc[ UR_RAWLOCK ]      := (@ADO_RAWLOCK())
   aADOFunc[ UR_LOCK ]         := (@ADO_LOCK())
   aADOFunc[ UR_UNLOCK ]       := (@ADO_UNLOCK())
   aADOFunc[ UR_SETFILTER ]    := (@ADO_SETFILTER())
   aADOFunc[ UR_CLEARFILTER ]  := (@ADO_CLEARFILTER())
   aADOFunc[ UR_ZAP ]          := (@ADO_ZAP())
   aADOFunc[ UR_SETLOCATE ]    := (@ADO_SETLOCATE())
   aADOFunc[ UR_LOCATE ]       := (@ADO_LOCATE())
   aADOFunc[ UR_FOUND ]        := (@ADO_FOUND())
   aADOFunc[ UR_FORCEREL ]     := (@ADO_FORCEREL())
   aADOFunc[ UR_RELEVAL ]      := (@ADO_RELEVAL())
   aADOFunc[ UR_CLEARREL ]     := (@ADO_CLEARREL())
   aADOFunc[ UR_RELAREA ]      := (@ADO_RELAREA())
   aADOFunc[ UR_RELTEXT ]      := (@ADO_RELTEXT())
   aADOFunc[ UR_SETREL ]       := (@ADO_SETREL())
   aADOFunc[ UR_ORDCREATE ]    := (@ADO_ORDCREATE())
   aADOFunc[ UR_ORDDESTROY ]   := (@ADO_ORDDESTROY())
   aADOFunc[ UR_ORDLSTADD ]    := (@ADO_ORDLSTADD())
   aADOFunc[ UR_ORDLSTCLEAR ]  := (@ADO_ORDLSTCLEAR())
   aADOFunc[ UR_EVALBLOCK ]    := (@ADO_EVALBLOCK())
   aADOFunc[ UR_SEEK ]         := (@ADO_SEEK())
   aADOFunc[ UR_EXISTS ]       := (@ADO_EXISTS())
   aADOFunc[ UR_DROP ]         := (@ADO_DROP())

   RETURN USRRDD_GETFUNCTABLE( pFuncCount, pFuncTable, pSuperTable, nRddID, ;
      /* NO SUPER RDD */, aADOFunc )

INIT PROCEDURE ADORDD_INIT()

   rddRegister( "ADORDD", RDT_FULL )

   RETURN

STATIC FUNCTION ADO_INFO(nWa, uInfoType,uReturn)
  LOCAL aWAData := USRRDD_AREADATA( nWA )
  LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]
  
  DO CASE
	CASE uInfoType == DBI_ISDBF   // 1  /* Does this RDD support DBFs? */
	
	     uReturn := .F.
		 
	CASE uInfoType == DBI_CANPUTREC  // 2  /* Can this RDD Put Records?   */
	
	     uReturn := .T.
		 
	CASE uInfoType == DBI_GETHEADERSIZE // 3  /* Data file's header size     */
	
	     uReturn := 0
		 
	CASE uInfoType == DBI_LASTUPDATE  // 4  /* The last date this file was written to  */
	
	     //SELECT QUERY DEPENDING OF THE DATABASE file INI with
		 // [SQL NAME]
		 // ConnString = "connection string"
		 // DBI_LASTUPDATE = "select query to te db" 
		 MsgInfo("Not ready yet Please consult ADORDD ADO_INFO FOR MORE INFO")
		 uReturn := CTOD("01/01/89")
		 
	CASE uInfoType == DBI_GETDELIMITER // 5  /* The delimiter (as a string)         */
	     uReturn := ""
	CASE uInfoType == DBI_SETDELIMITER // 6  /* The delimiter (as a string)         */
	     uReturn := ""
	CASE uInfoType == DBI_GETRECSIZE // 7  /* The size of 1 record in the file    */

	     ADO_RECINFO( nWA, ADO_RECID( nWA, @uReturn ), UR_DBRI_RECSIZE, @uReturn )
		 
	CASE uInfoType == DBI_GETLOCKARRAY // 8  /* An array of locked records' numbers */

         uReturn := aWAData[WA_LOCKLIST]
		 
	CASE uInfoType == DBI_TABLEEXT //  9  /* The data file's file extension      */
	
		 MSGINFO("IN SQL TABLE EXTENSION DO NOT MATTER!")
		 uReturn := NIL
		 
	CASE uInfoType == DBI_FULLPATH // 10  /* The Full path to the data file      */
		 
         uReturn := ""
		 
	CASE uInfoType == DBI_ISFLOCK // 20  /* Is there a file lock active?        */
	
	     uReturn := aWAData[WA_FILELOCK]
		 
	CASE uInfoType == DBI_CHILDCOUNT // 22  /* Number of child relations set       */
	
	     uReturn := 0
		 
	CASE uInfoType == DBI_FILEHANDLE // 23  /* The data file's OS file handle      */
	     uReturn := -1
	CASE uInfoType == DBI_BOF // 26  /* Same as bof()    */
	
	     uReturn := aWAData[WA_BOF]
		 
	CASE uInfoType == DBI_EOF // 27  /* Same as eof()    */
	
	     uReturn := aWAData[WA_EOF]
		 
	CASE uInfoType == DBI_DBFILTER // 28  /* Current Filter setting              */
	
         uReturn := oRecordSet:Filter 
		 
	CASE uInfoType == DBI_FOUND // 29  /* Same as found()  */
	
	     uReturn := aWAData[WA_FOUND]
		 
	CASE uInfoType == DBI_FCOUNT // 30  /* How many fields in a record?        */
	
	     uReturn := FCOUNT()
		 
	CASE uInfoType == DBI_LOCKCOUNT // 31  /* Number of record locks              */
	
	     uReturn := LEN(aWAData[WA_LOCKLIST])
		 
	CASE uInfoType == DBI_VALIDBUFFER  //  32  /* Is the record buffer valid?         */
	
	    IF aWAData[WA_EOF] .OR. aWAData[WA_BOF]
		   uReturn := .T.
		ELSE   
	       uReturn := oRecordSet:EditMode = adEditNone
		ENDIF   
		
	CASE uInfoType == DBI_ALIAS  // 33  /* Name (alias) for this workarea      */
	
	     uReturn := ALIAS()
		 
	CASE uInfoType == DBI_GETSCOPE // 34  /* The codeblock used in LOCATE        */
	     uReturn := ""
	CASE uInfoType == DBI_LOCKOFFSET //  35  /* The offset used for logical locking */
	     uReturn := 0
	CASE uInfoType == DBI_SHARED  //  36  /* Was the file opened shared?         */
	     uReturn := aWAData[WA_OPENSHARED]
	CASE uInfoType == DBI_MEMOEXT  //  37  /* The memo file's file extension      */
	     uReturn := ""
	CASE uInfoType == DBI_MEMOHANDLE // 38  /* File handle of the memo file        */
	     uReturn := -1
	CASE uInfoType == DBI_MEMOBLOCKSIZE  // 39  /* Memo File's block size              */
         uReturn := 0
	CASE uInfoType == DBI_DB_VERSION  //  101  /* Version of the Host driver          */
	     uReturn := "Version 2015"
	CASE uInfoType == DBI_RDD_VERSION // 102  /* current RDD's version               */
	     uReturn := "Version 2015"
       
  ENDCASE
  
 RETURN HB_SUCCESS //uReturn
 
STATIC FUNCTION ADOPSEUDOSEEK(nWA,cKey,aWAData,lSoftSeek,lBetween,cKeybottom)
 LOCAL nOrder := aWAData[WA_INDEXACTIVE]
 LOCAL cExpression := aWAData[WA_INDEXEXP][nOrder]
 LOCAL aLens := {}, n, aFields := {} , nAt := 1,cType, lNotFind := .F. ,cSqlExpression := "",nLen
  
 DEFAULT lSoftSeek TO .F.//to use like insead of =
 DEFAULT lBetween TO .F.
 
 cKey := CVALTOCHAR(cKey)
 cKeyBotom := CVALTOCHAR(cKeybottom)
 
    FOR n := 1 to (nWA)->(FCOUNT()) // we have to check all fields in table because there isnt any conspicuous mark on the expression to guide us
	    
	    nAt := AT((nWA)->(FIELDNAME(n)),cExpression)
		
	    IF nAt > 0

   		   AADD(aFields ,{(nWA)->(FIELDNAME(n)),nAt}) //nAt order of the field in the expression
		   
	   ENDIF 
	   
    NEXT
	
	//we need to have the fields with the same order as in index expression nAt
	aFields := ASORT( aFields ,,, {|x,y| x[2] < y[2] } )
	
    cExpression := ""
	cSqlExpression := ""
	
    FOR nAt := 1 TO LEN(aFields)
	   
	    nLen := FIELDSIZE(FIELDPOS(aFields[nAt,1]))
		cType := FIELDTYPE(FIELDPOS(aFields[nAt,1]))
		
		//extract from cKey the lengh og this field
		IF cType = "C" .OR. cType = "M"  
		   
		   IF !lBetween
	          cExpression += aFields[nAt,1]+IF(lSoftSeek," LIKE ", "=")+"'"+SUBSTR( cKey, 1, nLen)+"'"
		      cSqlExpression := cExpression
		   ELSE
	          cExpression += aFields[nAt,1]+" BETWEEN "+"'"+SUBSTR( cKey, 1, nLen)+"'"+;
			                 " AND "+"'"+SUBSTR( cKeyBottom, 1, nLen)+"'"
		      cSqlExpression := cExpression
           ENDIF
		   
		ELSEIF cType = "D" .OR. cType = "N"
		   
		   IF cType = "D"
		   
		      IF !lBetween
		         cExpression    += aFields[nAt,1]+ "="+"'"+ADODTOS(SUBSTR( cKey, 1, nLen))+"'" //delim might be #
		         cSqlExpression += aFields[nAt,1]+ "='"+ADODTOS(SUBSTR( cKey, 1, nLen))+"'"
		      ELSE
	             cExpression += aFields[nAt,1]+" BETWEEN "+"'"+ADODTOS(SUBSTR( cKey, 1, nLen))+"'"+;
			                 " AND "+"'"+ADODTOS(SUBSTR( cKeyBottom, 1, nLen))+"'"
		         cSqlExpression := cExpression
              ENDIF
			  
		   ELSE
		   
		      IF !lBetween
		         cExpression    += aFields[nAt,1]+ "="+"#"+SUBSTR( cKey, 1, nLen)+"#"
		         cSqlExpression += aFields[nAt,1]+ "="+SUBSTR( cKey, 1, nLen)
		      ELSE
	             cExpression += aFields[nAt,1]+" BETWEEN "+"#"+SUBSTR( cKey, 1, nLen)+"#"+;
			                 " AND "+"#"+SUBSTR( cKeyBottom, 1, nLen)+"#"
		         cSqlExpression := cExpression
              ENDIF
			  
		   ENDIF	  
		   
		ELSE

     	   lNotFind := .T.	//expression cannot be used by :Find()	
		   
		ENDIF
		
		cKey := SUBSTR(cKey,nLen+1) // take out the len of the expression already used
		IF LBetween
		   cKeybottom := SUBSTR(cKeybottom,nLen+1) // take out the len of the expression already used
		ENDIF
		
		IF nAt < LEN(aFields) //add AND last one isnt needed!
		
		   cExpression += IF(LEN(cKey) > 0 ," AND " , "")
		   cSqlExpression += IF(LEN(cKey) > 0 ," AND " , "")
		   
		ENDIF
		
		IF LEN(cKey) = 0 //there isnt more expression to look for
		   EXIT
		ENDIF   
		
    NEXT

  RETURN {cExpression,cSqlExpression,IF( lNotFind,.F.,nAt = 1)}
  

STATIC FUNCTION ADO_FIELDSTRUCT( oRs, n ) // ( oRs, nFld ) where nFld is 1 based
                                    // ( oRs, oField ) or ( oRs, cFldName )
                                    // ( oField )

   LOCAL oField, nType, uval
   LOCAL cType := 'C', nLen := 10, nDec := 0, lRW := .t.,nDBFFieldType :=  HB_FT_STRING // default
   LOCAL nFWAdoMemoSizeThreshold := 1024
   
   /*
     cType DBF TYPE "C","N","D" ETC
	 nDBFFieldType HB_FT_STRING ETC
	 based on the function FWAdoFieldStruct from Mr Rao
   */
   
   /* IF n == nil
      oField      := oRs
      oRs         := nil
   ELSEIF VALTYPE( n ) == 'O'
      oField      := n
   ELSE
      IF ValType( n ) == 'N'
         n--
      ENDIF
      TRY
         oField      := oRs:Fields( n )
      CATCH
      END
   ENDIF
   IF oField == nil
      RETURN nil
   ENDIF
   */
   oField      := oRs:Fields( n )
   nType       := oField:Type

   IF nType == adBoolean
   
      cType    := 'L'
      nLen     := 1
	  nDBFFieldType := HB_FT_LOGICAL
	  
   ELSEIF ASCAN( { adDate, adDBDate, adDBTime, adDBTimeStamp }, nType ) > 0
   
      cType    := 'D'
      nLen     := 8
      IF oRs != nil .AND. ! oRs:Eof() .AND. VALTYPE( uVal := oField:Value ) == 'T' 
	  //.AND. FW_TIMEPART( uVal ) >= 1.0 WHERE IS THIS FUNCTION?
         cType      := '@' //'T'
		 nLen := oField:DefinedSize
		 nDBFFieldType := HB_FT_TIMESTAMP // DONT KNWO IF IT IS CORRECT!
	  ELSE
	     nDBFFieldType := HB_FT_DATE
      ENDIF
	  
   ELSEIF ASCAN( { adTinyInt, adSmallInt, adInteger, adBigInt, ;
                  adUnsignedTinyInt, adUnsignedSmallInt, adUnsignedInt, ;
                  adUnsignedBigInt }, nType ) > 0
				  
      cType    := 'N'
      nLen     := oField:Precision + 1  // added 1 for - symbol
	  nDBFFieldType := HB_FT_INTEGER
	  
      IF oField:Properties( "ISAUTOINCREMENT" ):Value == .t.
         cType := '+'
         lRW   := .f.
		 nDBFFieldType := HB_FT_AUTOINC
      ENDIF
	  
   ELSEIF ASCAN( { adSingle, adDouble, adCurrency }, nType ) > 0
   
      cType    := 'N' //SHOULDNT BE "B"?
      nLen     := MIN( 19, oField:Precision-oField:NumericScale-1 ) //+ 2 )
      IF oField:NumericScale > 0 .AND. oField:NumericScale < nLen
         nDec  := oField:NumericScale
      ENDIF
      nDBFFieldType := HB_FT_INTEGER //HB_FT_DOUBLE WICH ONE IS CORRECT?
	  
   ELSEIF ASCAN( { adDecimal, adNumeric, adVarNumeric }, nType ) > 0
   
      cType    := 'N'
      nLen     := Min( 19, oField:Precision-oField:NumericScale-1 ) //+ 2 )
	  
      IF oField:NumericScale > 0 .AND. oField:NumericScale < nLen
         nDec  := oField:NumericScale
      ENDIF
	  
	  nDBFFieldType := HB_FT_INTEGER //HB_FT_LONG WICH ONE IS CORRECT?
	  
   ELSEIF ASCAN( { adBSTR, adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar }, nType ) > 0
   
      nLen     := oField:DefinedSize
	  nDBFFieldType := HB_FT_STRING
	  
      IF nType != adChar .AND. nType != adWChar .AND. nLen > nFWAdoMemoSizeThreshold
         cType := 'M'
         nLen  := 10
		 nDBFFieldType := HB_FT_MEMO
      ENDIF
	  
   ELSEIF ASCAN( { adBinary, adVarBinary, adLongVarBinary }, nType ) > 0
   
      cType := "G"
      nLen     := oField:DefinedSize
      IF nType != adBinary .AND. nLen > nFWAdoMemoSizeThreshold
         cType := 'M'
         nLen  := 10
      ENDIF
	  
      nDBFFieldType := HB_FT_OLE
	  
      IF nType != adBinary .AND. nLen > nFWAdoMemoSizeThreshold
         nDBFFieldType := HB_FT_MEMO
      ENDIF

   ELSEIF ASCAN( { adChapter, adPropVariant}, nType ) > 0
   
      cType    := 'O'
      lRW      := .f.
	  nDBFFieldType := HB_FT_MEMO
	  
   ELSEIF ASCAN( { adVariant, adIUnknown }, nType ) > 0 

      cType := "V"
      nDBFFieldType := HB_FT_ANY
	  
   ELSEIF ASCAN( { adGUID }, nType ) > 0 	  
   
      nDBFFieldType := HB_FT_STRING
	  
   ELSEIF ASCAN( { adFileTime }, nType ) > 0 	  	  
    
      cType := "T"	
      nDBFFieldType := HB_FT_DATETIME
	  
   ELSEIF ASCAN( { adEmpty, adError, adUserDefined, adIDispatch  }, nType ) > 0	  
   
      cType = 'O'
	  lRw := .t.
      nDBFFieldType := HB_FT_NONE //what is this? maybe NONE is wrong!
	  
   ELSE
   
      lRW      := .f.
	  
   ENDIF
   
   IF lAnd( oField:Attributes, 0x72100 ) .OR. ! lAnd( oField:Attributes, 8 )
      lRW      := .f.
   ENDIF
   
   RETURN { oField:Name, cType, nLen, nDec, nType, lRW, nDBFFieldType }

STATIC FUNCTION ADODTOS(cDate)
 LOCAL dDate ,cYear,cMonth,cDay

   IF AT( ".",cDate) > 0 .OR. AT("-" ,cDate) > 0 .OR. AT("/",cDate) > 0
      dDate := CTOD(cDate)  // hope to enforce set date format like this
   ELSE
      cYear  := SUBSTR(cDate,1,4)
	  cMonth := SUBSTR(cDate,5,2)
	  cDay   := SUBSTR(cDate,7,2) 
	  dDate  := CTOD(cDay+"/"+cMonth+"/"+cYear) // hope to enforce set date format like this
   ENDIF
   
   RETURN DTOC(dDate)
   
   
STATIC FUNCTION ADO_RECINFO( nWA, nRecord, nInfoType, uInfo )

   LOCAL nResult := HB_SUCCESS
   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]
   LOCAL n
   
   HB_SYMBOL_UNUSED( nWA )

#ifdef UR_DBRI_DELETED
   DO CASE
   CASE nInfoType == UR_DBRI_DELETED
   
      ADO_DELETED( nWA, @uInfo )
		
   CASE nInfoType == UR_DBRI_LOCKED
   
      FOR n:= 1 TO LEN(aWdata[ WA_LOCKLIST ])
		  IF nRecord = aWdata[ WA_LOCKLIST ][n]
             uInfo := .T.
			 EXIT
		  ENDIF
      NEXT			
		
   CASE nInfoType == UR_DBRI_RECSIZE 
   
      uInfo := 0
      FOR n := 1 TO FCOUNT()
		  uInfo += FIELDSIZE(n)
	  NEXT
		
   CASE nInfoType == UR_DBRI_RECNO
   
      nResult := ADO_RECID( nWA, @uInfo )
	  
   CASE nInfoType == UR_DBRI_UPDATED
   
      uInfo := .F.
	  
   CASE nInfoType == UR_DBRI_ENCRYPTED
   
      uInfo := .F.
	  
   CASE nInfoType == UR_DBRI_RAWRECORD
   
      uInfo := ""
	  
   CASE nInfoType == UR_DBRI_RAWMEMOS
   
      uInfo := ""
	  
   CASE nInfoType == UR_DBRI_RAWDATA
   
      nResult := ADO_GOTO( nWA, nRecord )
      uInfo := ""
	  
   ENDCASE
#else
   HB_SYMBOL_UNUSED( nRecord )
   HB_SYMBOL_UNUSED( nInfoType )
   HB_SYMBOL_UNUSED( uInfo )
#endif

   RETURN nResult
   
STATIC FUNCTION ADO_SETREL( nWA, aRelInfo )

   LOCAL aWAData := USRRDD_AREADATA( nWA ),n 
  
   IF VALTYPE(   aWAData[ WA_PENDINGREL ]) = "U"
       aWAData[ WA_PENDINGREL ] := {}
   ENDIF	   
   
   FOR n:= 1 TO LEN(aRelInfo)
      AADD(aWAData[ WA_PENDINGREL ] ,aRelInfo[n])
   NEXT
   
   RETURN HB_SUCCESS

   
STATIC FUNCTION ADO_FORCEREL( nWA )

   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL n,aPendingRel:=ARRAY(UR_RI_SIZE),nReturn := HB_SUCCESS
   
   IF !EMPTY(aWAData[ WA_PENDINGREL ])
   
      FOR n:= 1 TO LEN(aWAData[ WA_PENDINGREL ]) STEP UR_RI_SIZE //next elements next relations
       
	      ACOPY(aWAData[ WA_PENDINGREL ], aPendingRel, n, UR_RI_SIZE)
          nReturn := ADO_RELEVAL( nWA, aPendingRel )
	   
      NEXT
	  
   ENDIF
   
   RETURN nReturn
   

STATIC FUNCTION ADO_RELEVAL( nWA, aRelInfo )

   LOCAL aInfo, nReturn, nOrder, uResult

   
   nReturn := ADO_EVALBLOCK( aRelInfo[ UR_RI_PARENT ], aRelInfo[ UR_RI_BEXPR ], @uResult )

   IF nReturn == HB_SUCCESS
      /*
       *  Check the current order
       */
      aInfo := Array( UR_ORI_SIZE )
      nReturn := ADO_ORDINFO( aRelInfo[ UR_RI_CHILD ], DBOI_NUMBER, @aInfo )
	  
      IF nReturn == HB_SUCCESS
	  
         nOrder := aInfo[ UR_ORI_RESULT ]
		 
         IF nOrder != 0
            IF aRelInfo[ UR_RI_SCOPED ]
               aInfo[ UR_ORI_NEWVAL ] := uResult
               nReturn := ADO_ORDINFO( aRelInfo[ UR_RI_CHILD ], DBOI_SCOPETOP, @aInfo )
               IF nReturn == HB_SUCCESS
                  nReturn := ADO_ORDINFO( aRelInfo[ UR_RI_CHILD ], DBOI_SCOPEBOTTOM, @aInfo )
               ENDIF
            ENDIF
            IF nReturn == HB_SUCCESS
               //doesnt matter nreturn can be eof or bof sotory continunes
			   ADO_SEEK( aRelInfo[ UR_RI_CHILD ], .F., uResult, .F. )
            ENDIF
			
         ELSE
		    MSGINFO("Relations in ADO SQL with record number are not alloud! See adordd.prg")
            nReturn := ADO_GOTO( aRelInfo[ UR_RI_CHILD ], uResult )
         ENDIF
		 
      ENDIF
   ENDIF

   RETURN nReturn

STATIC FUNCTION ADO_EVALBLOCK( nArea, bBlock, uResult )

   LOCAL nCurrArea

   nCurrArea := Select()
   IF nCurrArea != nArea
      dbSelectArea( nArea )
   ELSE
      nCurrArea := 0
   ENDIF

   uResult := Eval( bBlock )

   IF nCurrArea > 0
      dbSelectArea( nCurrArea )
   ENDIF

   RETURN HB_SUCCESS

STATIC FUNCTION ADO_CLEARREL( nWA )

   LOCAL aWAData := USRRDD_AREADATA( nWA )
   
   aWAData[ WA_PENDINGREL ] := NIL
  
   RETURN HB_SUCCESS
   

STATIC FUNCTION ADO_RELAREA( nWA, nRelNo, nRelArea )

   LOCAL aWAData := USRRDD_AREADATA( nWA ),nPos
   
   nPos := nRelNo*UR_RI_SIZE-UR_RI_SIZE+UR_RI_CHILD
   
   IF !EMPTY(aWAData[ WA_PENDINGREL ])
   
       IF LEN(aWAData[ WA_PENDINGREL ]) >= nRelNo*UR_RI_SIZE

		  nRelArea := aWAData[ WA_PENDINGREL ][nPos]
		  
	   ELSE
	   
          nRelArea := 0	   
		   
       ENDIF		 
	  
   ELSE

      nRelArea := 0
     
   ENDIF
  
   
   RETURN HB_SUCCESS
   

STATIC FUNCTION ADO_RELTEXT( nWA, nRelNo, cExpr )

   LOCAL aWAData := USRRDD_AREADATA( nWA ),nPos
   
   nPos := nRelNo*UR_RI_SIZE-UR_RI_SIZE+UR_RI_CEXPR
   
   IF !EMPTY(aWAData[ WA_PENDINGREL ])
   
       IF LEN(aWAData[ WA_PENDINGREL ]) >= nRelNo*UR_RI_SIZE

		  cExpr := aWAData[ WA_PENDINGREL ][nPos]
	   ELSE
	   
          cExpr := 0	   
		   
       ENDIF		 
	  
   ELSE

      cExpr := 0
     
   ENDIF

   RETURN HB_SUCCESS
  
/*                               ADO  FUNCTIONS WORK NOT STARTED YET */

STATIC FUNCTION ADO_SETFILTER( nWA, aFilterInfo )

   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]

   oRecordSet:Filter := SQLTranslate( aFilterInfo[ UR_FRI_CEXPR ] )

   RETURN HB_SUCCESS

STATIC FUNCTION ADO_CLEARFILTER( nWA )

   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]

   TRY
      oRecordSet:Filter := ""
   CATCH
   END

   RETURN HB_SUCCESS


STATIC FUNCTION ADO_EXISTS( nRdd, cTable, cIndex, ulConnect )

   // LOCAL n
   LOCAL lRet := HB_FAILURE
   LOCAL aRData := USRRDD_RDDDATA( nRDD )

   HB_SYMBOL_UNUSED( ulConnect )

   IF ! Empty( cTable ) .AND. ! Empty( aRData[ WA_CATALOG ] )
      TRY
         // n := aRData[ WA_CATALOG ]:Tables( cTable )
         lRet := HB_SUCCESS //sql select "SHOW TABLES LIKE 'yourtable';"
      CATCH
      END
      IF ! Empty( cIndex )
         TRY
            // n := aRData[ WA_CATALOG ]:Tables( cTable ):Indexes( cIndex )
            lRet := HB_SUCCESS
         CATCH
         END
      ENDIF
   ENDIF

   RETURN lRet

STATIC FUNCTION ADO_DROP( nRdd, cTable, cIndex, ulConnect )

   // LOCAL n
   LOCAL lRet := HB_FAILURE
   LOCAL aRData := USRRDD_RDDDATA( nRDD )

   HB_SYMBOL_UNUSED( ulConnect )

   IF ! Empty( cTable ) .AND. ! Empty( aRData[ WA_CATALOG ] )
      TRY
         // n := aRData[ WA_CATALOG ]:Tables:Delete( cTable )
         lRet := HB_SUCCESS
      CATCH
      END
      IF ! Empty( cIndex )
         TRY
            // n := aRData[ WA_CATALOG ]:Tables( cTable ):Indexes:Delete( cIndex )
            lRet := HB_SUCCESS
         CATCH
         END
      ENDIF
   ENDIF

   RETURN lRet


STATIC FUNCTION ADO_ZAP( nWA )

   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := aWAData[ WA_RECORDSET ]

   IF aWAData[ WA_CONNECTION ] != NIL .AND. aWAData[ WA_TABLENAME ] != NIL
      TRY
         aWAData[ WA_CONNECTION ]:Execute( "TRUNCATE TABLE " + aWAData[ WA_TABLENAME ] )
      CATCH
         aWAData[ WA_CONNECTION ]:Execute( "DELETE * FROM " + aWAData[ WA_TABLENAME ] )
      END
      oRecordSet:Requery()
   ENDIF

   RETURN HB_SUCCESS
   
STATIC FUNCTION ADO_PACK( nWA )

// LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]

   HB_SYMBOL_UNUSED( nWA )
   
   RETURN HB_SUCCESS

STATIC FUNCTION ADO_CREATE( nWA, aOpenInfo )

   LOCAL cDataBase  := hb_tokenGet( aOpenInfo[ UR_OI_NAME ], 1, ";" )
   LOCAL cTableName := hb_tokenGet( aOpenInfo[ UR_OI_NAME ], 2, ";" )
   LOCAL cDbEngine  := hb_tokenGet( aOpenInfo[ UR_OI_NAME ], 3, ";" )
   LOCAL cServer    := hb_tokenGet( aOpenInfo[ UR_OI_NAME ], 4, ";" )
   LOCAL cUserName  := hb_tokenGet( aOpenInfo[ UR_OI_NAME ], 5, ";" )
   LOCAL cPassword  := hb_tokenGet( aOpenInfo[ UR_OI_NAME ], 6, ";" )

   LOCAL oConnection :=  TOleAuto():New( "ADODB.Connection" )
   LOCAL oCatalog :=  TOleAuto():New( "ADOX.Catalog" )
   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL oError, n

   DO CASE
   CASE Lower( Right( cDataBase, 4 ) ) == ".mdb"
      IF ! file( cDataBase )
         oCatalog:Create( "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + cDataBase )
      ENDIF
      oConnection:Open( "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + cDataBase )

   CASE Lower( Right( cDataBase, 4 ) ) == ".xls"
      IF ! file( cDataBase )
         oCatalog:Create( "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + cDataBase + ";Extended Properties='Excel 8.0;HDR=YES';Persist Security Info=False" )
      ENDIF
      oConnection:Open( "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + cDataBase + ";Extended Properties='Excel 8.0;HDR=YES';Persist Security Info=False" )

   CASE Lower( Right( cDataBase, 3 ) ) == ".db"
      IF ! file( cDataBase )
         oCatalog:Create( "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + cDataBase + ";Extended Properties='Paradox 3.x';" )
      ENDIF
      oConnection:Open( "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + cDataBase + ";Extended Properties='Paradox 3.x';" )

   CASE Lower( Right( cDataBase, 4 ) ) == ".fdb"
      IF ! file( cDataBase )
         oCatalog:Create( "Driver=Firebird/InterBase(r) driver;Uid=" + cUserName + ";Pwd=" + cPassword + ";DbName=" + cDataBase + ";" )
      ENDIF
      oConnection:Open( "Driver=Firebird/InterBase(r) driver;Uid=" + cUserName + ";Pwd=" + cPassword + ";DbName=" + cDataBase + ";" )
      oConnection:CursorLocation := adUseClient

   CASE Upper( cDbEngine ) == "MYSQL"
      oConnection:Open( ;
         "Driver={MySQL ODBC 3.51 Driver};" + ;
         "server=" + cServer + ;
         ";database=" + cDataBase + ;
         ";uid=" + cUserName + ;
         ";pwd=" + cPassword )

   ENDCASE

   TRY
      oConnection:Execute( "DROP TABLE " + cTableName )
   CATCH
   END

   TRY
      IF Lower( Right( cDataBase, 4 ) ) == ".fdb"
         oConnection:Execute( "CREATE TABLE " + cTableName + " (" + StrTran( StrTran( aWAData[ WA_SQLSTRUCT ], "[", '"' ), "]", '"' ) + ")" )
      ELSE
         oConnection:Execute( "CREATE TABLE [" + cTableName + "] (" + aWAData[ WA_SQLSTRUCT ] + ")" )
      ENDIF
   CATCH
      oError := ErrorNew()
      oError:GenCode := EG_CREATE
      oError:SubCode := 1004
      oError:Description := hb_langErrMsg( EG_CREATE ) + " (" + ;
         hb_langErrMsg( EG_UNSUPPORTED ) + ")"
      oError:FileName := aOpenInfo[ UR_OI_NAME ]
      oError:CanDefault := .T.

      FOR n := 0 TO oConnection:Errors:Count - 1
         oError:Description += oConnection:Errors( n ):Description
      NEXT

      UR_SUPER_ERROR( nWA, oError )
   END

   oConnection:Close()

   RETURN HB_SUCCESS

STATIC FUNCTION ADO_CREATEFIELDS( nWA, aStruct )

   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL n

   aWAData[ WA_SQLSTRUCT ] := ""

   FOR n := 1 TO Len( aStruct )
      IF n > 1
         aWAData[ WA_SQLSTRUCT ] += ", "
      ENDIF
      aWAData[ WA_SQLSTRUCT ] += "[" + aStruct[ n ][ DBS_NAME ] + "]"
      DO CASE
      CASE aStruct[ n ][ DBS_TYPE ] $ "C,Character"
         aWAData[ WA_SQLSTRUCT ] += " CHAR(" + str( aStruct[ n ][ DBS_LEN ] ) + ") NULL"

      CASE aStruct[ n ][ DBS_TYPE ] == "V"
         aWAData[ WA_SQLSTRUCT ] += " VARCHAR(" + str( aStruct[ n ][ DBS_LEN ] ) + ") NULL"

      CASE aStruct[ n ][ DBS_TYPE ] == "B"
         aWAData[ WA_SQLSTRUCT ] += " DOUBLE NULL"

      CASE aStruct[ n ][ DBS_TYPE ] == "Y"
         aWAData[ WA_SQLSTRUCT ] += " SMALLINT NULL"

      CASE aStruct[ n ][ DBS_TYPE ] == "I"
         aWAData[ WA_SQLSTRUCT ] += " MEDIUMINT NULL"

      CASE aStruct[ n ][ DBS_TYPE ] == "D"
         aWAData[ WA_SQLSTRUCT ] += " DATE NULL"

      CASE aStruct[ n ][ DBS_TYPE ] == "T"
         aWAData[ WA_SQLSTRUCT ] += " DATETIME NULL"

      CASE aStruct[ n ][ DBS_TYPE ] == "@"
         aWAData[ WA_SQLSTRUCT ] += " TIMESTAMP NULL DEFAULT CURRENT_TIMESTAMP"

      CASE aStruct[ n ][ DBS_TYPE ] == "M"
         aWAData[ WA_SQLSTRUCT ] += " TEXT NULL"

      CASE aStruct[ n ][ DBS_TYPE ] == "N"
         aWAData[ WA_SQLSTRUCT ] += " NUMERIC(" + str( aStruct[ n ][ DBS_LEN ] ) + ")"

      CASE aStruct[ n ][ DBS_TYPE ] == "L"
         aWAData[ WA_SQLSTRUCT ] += " LOGICAL"
      ENDCASE
   NEXT

   RETURN HB_SUCCESS

/*                   END ADO FUNCTIONS WORK  NOT STARTED YET*/


PROCEDURE hb_adoSetTable( cTableName )

   t_cTableName := cTableName

   RETURN

PROCEDURE hb_adoSetEngine( cEngine )

   t_cEngine := cEngine

   RETURN

PROCEDURE hb_adoSetServer( cServer )

   t_cServer := cServer

   RETURN

PROCEDURE hb_adoSetUser( cUser )

   t_cUserName := cUser

   RETURN

PROCEDURE hb_adoSetPassword( cPassword )

   t_cPassword := cPassword

   RETURN

PROCEDURE hb_adoSetQuery( cQuery )

   if( empty(cQuery), cQuery := "SELECT * FROM " ,cQuery)

   t_cQuery := cQuery

   RETURN

PROCEDURE hb_adoSetLocateFor( cLocateFor )

   USRRDD_AREADATA( Select() )[ WA_LOCATEFOR ] := cLocateFor

   RETURN

STATIC FUNCTION SQLTranslate( cExpr )

   IF Left( cExpr, 1 ) == '"' .AND. Right( cExpr, 1 ) == '"'
      cExpr := SubStr( cExpr, 2, Len( cExpr ) - 2 )
   ENDIF

   cExpr := StrTran( cExpr, '""' )
   cExpr := StrTran( cExpr, '"', "'" )
   cExpr := StrTran( cExpr, "''", "'" )
   cExpr := StrTran( cExpr, "==", "=" )
   cExpr := StrTran( cExpr, ".and.", "AND" )
   cExpr := StrTran( cExpr, ".or.", "OR" )
   cExpr := StrTran( cExpr, ".AND.", "AND" )
   cExpr := StrTran( cExpr, ".OR.", "OR" )
   cExpr := StrTran( cExpr, "RTRIM", "TRIM" )

   RETURN cExpr

FUNCTION hb_adoRddGetConnection( nWA )

   IF ! HB_ISNUMERIC( nWA )
      nWA := Select()
   ENDIF

   RETURN USRRDD_AREADATA( nWA )[ WA_CONNECTION ]

FUNCTION hb_adoRddGetRecordSet( nWA )

   LOCAL aWAData

   IF ! HB_ISNUMERIC( nWA )
      nWA := Select()
   ENDIF

   aWAData := USRRDD_AREADATA( nWA )

   RETURN iif( aWAData != NIL, aWAData[ WA_RECORDSET ], NIL )

function AdoConnect()

   local oDL,cConnection := "nada"
  
  
      oDL = CreateObject( "Datalinks" ):PromptNew()

      if ! Empty( oDL )
         cConnection = oDL:ConnectionString
      endif

return cConnection

FUNCTION ListIndex(nOption) 
//index files array needed for the adordd for your application
//order expressions already translated to sql DONT FORGET TO replace taitional + sign with ,
//we can and should include the SQL CONVERT to translate for ex DTOS etc
//ARRAY SPEC { {"TABLENAME",{"INDEXNAME","INDEXKEY","WHERE EXPRESSION AS USED FOR FOREXPRESSION","UNIQUE - DISTINCT ANY SQL STAT BEFORE * FROM"} }
//temporary indexes are not included gere they are create on fly and added to temindex list array
//they are only valid through the duration of the application
//the temp index name is auto given by adordd

LOCAL  Alista_fic:= { {"SECTOR",{"SECTOR","SECTORES"} },;
 {"CCLIENTE",{"COD_CLI","CODCLIENTE"},;
   {"CLIENTE","NOME"},{"CTEMP2","VENDEDOR,SECTOR,CODCLIENTE"} },;
 {"PPRODUTO",{"COD_PROD","CODIGOPROD"},;
  {"PRODUTO","PRODUTO"} },;
 {"FORNECED",{"COD_FOR","CODIGOFORN"},;
   {"FORNECED","NOME"} },;
 {"FORNPROD",{"COD_FP","CODIGOFORN,CODIGOPROD,ARMAZEM,DATAOF ASC)"},;
   {"COD_PF","CODIGOPROD,ARMAZEM,CODIGOFORN,DATAOF ASC"} ,;
   {"COD_DP","CODIGOPROD,DTOS(DATAOF)"},{"COD_DF","CODIGOFORN,DATAOF ASC"}},;
 {"DADOENC1",{"NRE_F","NRENCOMEND,CODIGOFORN,ANO ASC, SEMPREVCHE ASC"},;
  {"NRF_E","CODIGOFORN,NRENCOMEND,ANO ASC,SEMPREVCHE ASC"} },;
 {"ENCFOCLI",{"ENCPROD","DISTINCT NRENCOMEND,CODIGOPROD,ARMAZEM"},;
   {"PRODENC","CODIGOPROD,ARMAZEM,NRENCOMEND"},;
   {"FORNENC","CODIGOFORN"},;
   {"NRCAL_F","NRCALCULO,CODIGOFORN,ANO ASC,SEMPREVCHE ASC"},;
   {"NRCAL_P","NRCALCULO,CODIGOPROD,ARMAZEM,ANO ASC,SEMPREVCHE ASC"},;
   {"FCONTRT","CODIGOPROD,NRCONTRATO"} },;
 {"ENCCLIST",{"ENCCPRO","NRENCOMEND,CODIGOPROD,ARMAZEM"},;
   {"PROENCC","CODIGOPROD,ARMAZEM,NRENCOMEND"},{"MECPD","GUIA,CODCLIENTE,CODIGOPROD,ARMAZEM"},;
   {"CDENCCL","CODCLIENTE"},{"PTFACT","NRFACTUR,CODCLIENTE,CODIGOPROD,ARMAZEM"},;
   {"NRFPROD","NRFACTUR,CODIGOPROD,ARMAZEM,ANO ASC,SEMENTREGA ASC"} ,;
   {"CCONTRT","CODIGOPROD,NRCONTRATO"}},;
 {"STOCKS",{"STK_PROD","CODIGOPROD,ARMAZEM"} },;
 {"COCOFORN",{"COFORN_C","CODIGOFORN"},{"NRDOCORI","NRDOCORIGE,CODIGOFORN"} },;
 {"COCOCLI",{"COCLI_C","CODCLIENTE"},{"NRDOC","NRDOCORIGE,CODCLIENTE"} },;
 {"FACTURAS",{"FACTU","NRFACTUR"},{"ENC","NRENCOMEND"},;
   {"CODCLIE","CODCLIENTE,NRFACTUR"},{"RECFAC","RECIBO,NRFACTUR"},;
   {"CODFORE","CODIGOFORN,NRFACTUR"},{"DATAVENCE","DATAFACTUR+CONVERT(CDPAGA,SQL_NUMERIC)"} ,;
    {"DATAFCT","CODCLIENTE,CODIGOFORN,DATAFACTUR ASC,NRFACTUR"} },;
   {"MOVIME",{"MOVPROD","CODIGOPROD,ARMAZEM"} },;
 {"REGPAGA",{"REGCOD","CODCLIENTE,CODIGOFORN,NREGMOV ASC"} },;
 {"TABLANCA",{"CODLAN","DESIGNACAO"} },;
 {"MERCADO",{"COD_MERC","CODCLIENTE,CODIGOPROD"},;
   {"CD_MC_2","CODIGOPROD,CODCLIENTE"},;
   {"CD_MC_3","CODIGOFORN,CODIGOPROD,CODCLIENTE"} },;
 {"AMOSTRAS",{"AMOST_CA","CODCLIENTE,AMOSTRA"} },;
 {"VISITAS",{"COD_VIS","CODCLIENTE,DATAVISITA ASC"} },;
 {"OFERTAS",{"CLIPRO","CODCLIENTE,CODIGOPROD,ARMAZEM,DATAOFERTA ASC,NUMEROOFER,NROFER"},;
   {"PROCLI","CODIGOPROD,ARMAZEM,CODCLIENTE,DATAOFERTA ASC,NUMEROOFER,NROFER"},;
   {"CLIPRO2","CODCLIENTE,DATAOFERTA ASC,NUMEROOFER,CODIGOPROD,ARMAZEM,NROFER"},;
   {"PROCLI2","CODIGOPROD,ARMAZEM,DATAOFERTA ASC,NUMEROOFER,CODCLIENTE,NROFER"} },;
 {"CONTACTO",{"CONT_CLI","CODCLIENTE,NOME"},{"CLIANIVS","ANIVERSARI"} },;
 {"FORCONT",{"CONT_FOR","CODIGOFORN,NOME"},{"FORANIVS","ANIVERSARI"} },;
 {"VENDAS",{"VENDCLIC","CODCLIENTE,DTOS(DATAFACTUR),CODIGOPROD,ARMAZEM"},{"VENDCOMP",;
  "CODIGOPROD,ARMAZEM,DATAFACTUR ASC,CODCLIENTE"} },;
 {"LOTESTK",{"LOTESTK","CODIGOPROD,ARMAZEM,QUANTREAL ASC,DATAPROD ASC"} }, ;
 {"LOTESCC",{"LOTESCC","CODIGOPROD,ARMAZEM,LOTENR,DATAMOV ASC"},;
   {"LOTEENC","CODIGOPROD,ARMAZEM,ALIAS,DOCUMENTO"} }, ;
 {"ANOANTE",{"ANOANTE","CODIGOPROD,ARMAZEM"}},;
 {"PAISES",{"PAISES","NOMEPAIS"}},{"LOGENC",{"LOGENC","ALIAS,NRENCOMEND"}},;
 {"COMPRASH",{"COMPNRE","NRENCOMEND"} },;
 {"COMPRASE",{"COMPFOR","CODIGOFORN,DATACHEGAD ASC,CODIGOPROD,ARMAZEM"},;
 {"COMPPROD","CODIGOPROD,ARMAZEM,DATACHEGAD ASC,CODIGOFORN"} },;
 {"DESCPROD",{"DESCCODF","CODIGOPROD,CODCLIENTE,SECTOR" },;
 {"DESCDEF","DESCRICAO,CODIGOPROD,CODCLIENTE,SECTOR" } } ,;
 {"OUTROCOD",{"OUTRCOD1","CODIGOPROD,CODEQUIVAL"},{"OUTRCOD2","CODEQUIVAL,CODIGOPROD"}},;
 {"COMPOSTO",{"COD_CST","CODCOMPOST,CODIGOPROD" }},;
 {"DESPESA",{"TDESPESA","DESPESA"}},;
 {"HISTSTK",{"HISTSTK","MES"}} ,;
 {"CERTDOCS",{"CERTDOCS","NRDOC"}} ,;
 {"SAFTGUIA",{"SAFTGUIA","NRDOC"}} ,;
 {"DADOSDOC",{"DADODOC","NRDOC,CODCLIENTE,CODIGOFORN"}} ,;
 {"GUIA",{"GUIATAB","NRENCOMEND"}} ,;
 {"CCFACT",{"CCFACT","NRENCOMEND,CODIGOPROD,DATAMOV ASC"},{"CCFACT1","NRFACTUR,CODIGOPROD,DATAMOV ASC"}} ,;
 {"ACESS1",{"ACESS1","ROTINA"}},;
 {"APPT",{"APPTDATE","DATE,USERID,TIME"},{"APPTNOME","NOME,DATE,TIME"},;
 {"APPTPEND","TERMINADO,DATE,TIME,USERID"},{"APPTUSER","USERID,DATE,TIME"}} }

RETURN Alista_fic

FUNCTION ListTmpIndex()
 STATIC aTmpIndex := {}
  RETURN aTmpIndex
  
FUNCTION ListTmpExp()
 STATIC aTmpExp := {}
  RETURN aTmpExp
  
FUNCTION ListTmpFor()
 STATIC aTmpFor := {}
  RETURN aTmpFor
   
FUNCTION ListTmpUnique()
  STATIC aTmpUniques := {}
  RETURN aTmpUniques
   
 FUNCTION ListTmpNames()
  LOCAL aTmpNames := {"TMP","TEMP"}
  RETURN aTmpNames 
 
