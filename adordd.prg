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
 * Copyright 2015 AHF - Antonio H. Ferreira <disal.antonio.ferreira@gmail.com>
 * READ CAREFULLY
 *
 * This has been completly rewriten with a diferent kind of approach
 * not deal with Catalogs File indexes - DBA responsability
 * converting indexes to :SORT and treat indexes as "virtual" as they really dont exist as files
 *
 * Seek translate to :find and :filter
 * Indexes with UDFs allowed
 * Tables must have some ID field to be used as RECNO and DELETE FIELD
 * to be used as deleted flag
 * It is 100% compatible with any other dbf rdd kind of working
 * just need to compile and link it with your app.
 * No code change requeried!
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

ANNOUNCE ADORDD

#include "rddsys.ch"
#include "fileio.ch"
#include "error.ch"
#include "adordd.ch"
#include "common.ch"
#include "dbstruct.ch"
#include "dbinfo.ch"

#include "hbusrrdd.ch"  //verify that your version has the array field size of 7 for xarbour at least for 2008 version

#define WA_RECORDSET       1
#define WA_BOF             2
#define WA_EOF             3
#define WA_CONNECTION      4
#define WA_CATALOG         5
#define WA_TABLENAME       6
#define WA_ENGINE          7
#define WA_SERVER          8
#define WA_USERNAME        9
#define WA_PASSWORD       10
#define WA_QUERY          11
#define WA_LOCATEFOR      12
#define WA_SCOPEINFO      13
#define WA_SQLSTRUCT      14
#define WA_CONNOPEN       15
#define WA_PENDINGREL     16
#define WA_FOUND          17
#define WA_INDEXES        18 //AHF
#define WA_INDEXEXP       19 //AHF
#define WA_INDEXFOR       20 //AHF
#define WA_INDEXUNIQUE    24//AHF
#define WA_INDEXACTIVE    21 //AHF
#define WA_LOCKLIST       22 //AHF
#define WA_FILELOCK       23 //AHF
#define WA_OPENSHARED     25//AHF
#define WA_SCOPES         26//AHF
#define WA_SCOPETOP       27//AHF
#define WA_SCOPEBOT       28//AHF
#define WA_ISITSUBSET     29//AHF
#define WA_LASTRELKEY     30//AHF
#define WA_FILTERACTIVE   31//AHF
#define WA_FIELDRECNO     32//AHF
#define WA_TLOCKS         33//AHF
#define WA_FILEHANDLE     34//AHF
#define WA_LOCKSCHEME     35 //AHF
#define WA_CFILTERACTIVE  36//AHF
#define WA_LREQUERY       37//AHF
#define WA_RECCOUNT       38//AHF
#define WA_ABOOKMARKS     39//AHF
#define WA_INDEXDESC      40//AHF
#define WA_FIELDDELETED   41 //MAROMANO
#define WA_SIZE           41

#define RDD_CONNECTION    1
#define RDD_CATALOG       2
#define RDD_SIZE          2

#DEFINE CRLF CHR(13)+CHR(10)


STATIC t_cDataSource
STATIC t_cEngine
STATIC t_cServer
STATIC t_cUserName
STATIC t_cPassword
STATIC t_cQuery
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
   aWAData[WA_OPENSHARED] := NIL
   aWAData[WA_SCOPES] := {}
   aWAData[WA_SCOPETOP] := {}
   aWAData[WA_SCOPEBOT] := {}
   aWAData[WA_ISITSUBSET] := .F.
   aWAData[WA_FOUND] := .F.
   aWAData[WA_LASTRELKEY] := NIL
   aWAData[WA_FILTERACTIVE] := NIL
   aWAData[WA_FIELDRECNO] := NIL
   aWAData[WA_SCOPEINFO] := NIL
   aWAData[WA_TLOCKS] := {}
   aWAData[WA_FILEHANDLE ] := NIL
   aWAData[WA_LOCKSCHEME ] := ADOFORCELOCKS()  //no lock type 999
   aWAData[WA_CFILTERACTIVE ] := ""
   aWAData[WA_LREQUERY] := .F.
   aWAData[WA_RECCOUNT] := NIL //27.06.15
   aWAData[WA_ABOOKMARKS] := {}
   aWAData[WA_INDEXDESC] := {}
   aWAData[WA_FIELDDELETED] := NIL

   USRRDD_AREADATA( nWA, aWAData )


   RETURN HB_SUCCESS


STATIC FUNCTION ADO_OPEN( nWA, aOpenInfo )

   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL cName, aField, oError, nResult
   LOCAL oRecordSet, nTotalFields, n

   ADOCONNECT (nWA, aOpenInfo )

   /* When there is no ALIAS we will create new one using file name */
   IF Empty( aOpenInfo[ UR_OI_ALIAS ] )
      hb_FNameSplit( aOpenInfo[ UR_OI_NAME ],, @cName )
      aOpenInfo[ UR_OI_ALIAS ] := cName

   ENDIF

   //24.06.15 APPEND FROM
   IF PROCNAME( 1 ) == "__DBAPP"
      aOpenInfo[ UR_OI_ALIAS ] := cGetNewAlias( aOpenInfo[ UR_OI_ALIAS ] )

   ENDIF

   IF !ADOCON_CHECK()
      RETURN HB_FAILURE

   ENDIF

   //OPEN EXCLUSIVE
   IF !aOpenInfo[ UR_OI_SHARED ]
      aWAData[WA_OPENSHARED] := .F.
      IF !ADO_OPENSHARED( nWA, aWAData[WA_TABLENAME], .T.)
         oError := ErrorNew()
         oError:GenCode := EG_OPEN
         oError:SubCode := 1001
         oError:Description := hb_langErrMsg( EG_OPEN )
         oError:FileName := aOpenInfo[ UR_OI_NAME ]
         oError:OsCode := 0 // TODO
         oError:CanDefault := .T.
         NETERR(.T.)
         UR_SUPER_ERROR( nWA, oError )
         RETURN HB_FAILURE
      ENDIF

   ELSE
      aWAData[WA_OPENSHARED] := .T.
      IF ! ADO_OPENSHARED( nWA, aWAData[WA_TABLENAME], .F. )
         oError := ErrorNew()
         oError:GenCode := EG_OPEN
         oError:SubCode := 1001
         oError:Description := hb_langErrMsg( EG_OPEN )
         oError:FileName := aOpenInfo[ UR_OI_NAME ]
         oError:OsCode := 0 // TODO
         oError:CanDefault := .T.
         NETERR(.T.)
         UR_SUPER_ERROR( nWA, oError )
         RETURN HB_FAILURE
      ENDIF

   ENDIF

   oRecordSet := ADOCLASSNEW( "ADODB.Recordset" ) //TOleAuto():New( "ADODB.Recordset" )

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

   oRecordSet:CursorType     := IF(aWAData[ WA_ENGINE ] = "ACCESS", adOpenDynamic, adOpenDynamic) // adOpenKeyset adOpenDynamic
   oRecordSet:CursorLocation := IF(aWAData[ WA_ENGINE ] = "ACCESS", adUseClient, IF(aWAData[ WA_ENGINE ] = "ADS",adUseServer,adUseClient) ) //adUseServer  // adUseClient its slower but has avntages such always bookmaks
   oRecordSet:LockType       := adLockOptimistic //adLockOptimistic adLockPessimistic

   IF aOpenInfo[UR_OI_READONLY]
      oRecordSet:LockType := adLockReadOnly

   ELSE
      oRecordSet:LockType :=  IF(aWAData[ WA_ENGINE ] = "ACCESS", adLockOptimistic, adLockOptimistic) //adLockPessimistic //adLockOptimistic

   ENDIF

   //PROPERIES AFFECTING PERFORMANCE TRY
   oRecordSet:CacheSize := 30 //records increase performance set zero returns error set great server parameters max open rows error

   IF aWAData[ WA_QUERY ] == "SELECT * FROM "  //10.08.15 ORDER BY RECNO
      oRecordSet:Open( aWAData[ WA_QUERY ] + aWAData[ WA_TABLENAME ]+" ORDER BY "+ADO_GET_FIELD_RECNO(  aWAData[ WA_TABLENAME ] ), aWAData[ WA_CONNECTION ] )

   ELSE
      oRecordSet:Open( aWAData[ WA_QUERY ], aWAData[ WA_CONNECTION ] )

   ENDIF

   aWAData[ WA_RECORDSET ] := oRecordSet
   aWAData[ WA_BOF ] := aWAData[ WA_EOF ] := .F.

   UR_SUPER_SETFIELDEXTENT( nWA, nTotalFields := oRecordSet:Fields:Count )

   FOR n := 1 TO nTotalFields
      aField := Array( UR_FI_SIZE )
      aField[ UR_FI_NAME ]    := oRecordSet:Fields( n - 1 ):Name
      aField[ UR_FI_TYPE ]    := ADO_FIELDSTRUCT( oRecordSet, n-1 )[7]
      aField[ UR_FI_TYPEEXT ] := 0
      aField[ UR_FI_LEN ]     := ADO_FIELDSTRUCT( oRecordSet, n-1 )[3]
      aField[ UR_FI_DEC ]     := ADO_FIELDSTRUCT( oRecordSet, n-1 )[4]
      #ifdef __XHARBOUR__
          aField[ UR_FI_FLAGS ] := 0  // xHarbour expecs this field
          aField[ UR_FI_STEP ] := 0 // xHarbour expecs this field
      #endif

      // CHECK IF IT EXISTS RECNO FIELD
      //FROM _DBAPP DONT KNOW WHY BUT FIELD TYPE ALWAYS = N
      IF ALLTRIM( oRecordSet:Fields( n - 1 ):Name ) == ADO_GET_FIELD_RECNO(  aWAData[ WA_TABLENAME ] ) ;
        .AND. ADO_FIELDSTRUCT( oRecordSet, n-1 )[2] = "+" .OR. ;
        ( ADO_FIELDSTRUCT( oRecordSet, n-1 )[2] = "N" .AND. PROCNAME( 1 ) == "__DBAPP" )

         aWAData[WA_FIELDRECNO]:=  n - 1
         //IF IT SUPPORTS SEEK WE WILL SEEK IT ISNTEAD OF FIND IT
         IF !oRecordSet:Supports(adIndex) .OR. !oRecordSet:Supports(adSeek)
            //OTHERWISE LETS USE ADO INDEX PROP TO SPEED UP
            IF oRecordSet:CursorLocation = adUseClient
               oRecordSet:Fields( aWAData[WA_FIELDRECNO] ):Properties():Item("Optimize"):Value := 1
            ENDIF
         ENDIF

      ENDIF

      IF  ALLTRIM( oRecordSet:Fields( n - 1 ):Name )  = ADO_GET_FIELD_DELETED(aWAData [ WA_TABLENAME]);
         .AND. ADO_FIELDSTRUCT( oRecordSet, n-1 )[2] = "L"
         aWAData[WA_FIELDDELETED]:=  n - 1

      ENDIF

      UR_SUPER_ADDFIELD( nWA, aField )

   NEXT

   nResult := UR_SUPER_OPEN( nWA, aOpenInfo )

   //11.08.15 WE NEED SOME FIELD AS RECNO
   IF aWAData[WA_FIELDRECNO] = NIL
      ADO_CLOSE( nWA )
      THROW( ErrorNew( "ADORDD", 10000, 10000, "Open "+aWAData [ WA_TABLENAME] , "ADO needs field autoinc used as Recno " ) )

   ENDIF

   IF aWAData[WA_FIELDDELETED] = NIL
      ADO_CLOSE( nWA )
      THROW( ErrorNew( "ADORDD", 10000, 10000, "Open "+aWAData [ WA_TABLENAME] , "ADO needs field DELETED as logical used as DELETED()"+aWAData [ WA_TABLENAME] ) )

   ENDIF

   IF nResult == HB_SUCCESS
      ADO_GOTOP( nWA )

   ENDIF

   //auto open set and auto order
   IF SET(_SET_AUTOPEN)
      ADO_INDEXAUTOOPEN(aWAData[ WA_TABLENAME ])

   ENDIF

   //22.08.15 NEW ADDITIONS BY OTHERS CONTROL INITIALIZING
   ADO_REFRESH( nWA )

   RETURN nResult


STATIC FUNCTION ADOCONNECT(nWA,aOpenInfo)

 LOCAL aWAData := USRRDD_AREADATA( nWA )
 LOCAL aDefaults := ADODEFAULTS() //get defaults set or the sets called wth USE
 LOCAL cConnect  := hb_tokenGet( aOpenInfo[ UR_OI_NAME ], 2, "@" )

  aOpenInfo[ UR_OI_NAME ] := hb_tokenGet( aOpenInfo[ UR_OI_NAME ], 1, "@" ) //2.6.15 SEE ADORDD.CH USE DIRECTIVE

  TRY
      IF Empty( cConnect )//aOpenInfo[ UR_OI_CONNECT ] )

         IF EMPTY(oConnection)
            aWAData[ WA_CONNECTION ] :=  ADOCLASSNEW( "ADODB.Connection" ) //TOleAuto():New( "ADODB.Connection" )
            oConnection := aWAData[ WA_CONNECTION ]
            aWAData[ WA_TABLENAME ]  := UPPER(aOpenInfo[ UR_OI_NAME ])
            aWAData[ WA_QUERY ]      := t_cQuery
            aWAData[ WA_USERNAME ]   := t_cUserName
            aWAData[ WA_PASSWORD ]   := t_cPassword
            aWAData[ WA_SERVER ]     := t_cServer
            aWAData[ WA_ENGINE ]     := t_cEngine
            aWAData[ WA_CONNOPEN ]   := .T.
            aWAData[ WA_CATALOG ]    := t_cDatasource

            //23.06.15
            ADOOPENCONNECT( aWAData[ WA_CATALOG ], aWAData[ WA_SERVER ], aWAData[ WA_ENGINE ],;
                            aWAData[ WA_USERNAME ], aWAData[ WA_PASSWORD ] , aWAData[ WA_CONNECTION ] )

         ELSE
            // ITS ALREDY OPEN THE ADODB CONN USE THE SAME WE WANT TRANSACTIONS WITHIN THE CONNECTION
            aWAData[ WA_CONNECTION ] :=  oConnection
            aWAData[ WA_TABLENAME ]  := UPPER(aOpenInfo[ UR_OI_NAME ])
            aWAData[ WA_QUERY ]      := t_cQuery
            aWAData[ WA_USERNAME ]   := t_cUserName
            aWAData[ WA_PASSWORD ]   := t_cPassword
            aWAData[ WA_SERVER ]     := t_cServer
            aWAData[ WA_ENGINE ]     := t_cEngine
            aWAData[ WA_CONNOPEN ]   := .T.
            aWAData[ WA_CATALOG ]    := t_cDatasource

         ENDIF

      ELSE
         // here we dont save oconnection for the next one because
         // we assume that is not application defult conn but a temporary conn
         //to other db system.
         aWAData[ WA_CONNECTION ] := ADOCLASSNEW( "ADODB.Connection" )//TOleAuto():New("ADODB.Connection")
         aWAData[ WA_CONNECTION ]:ConnectionTimeOut := 60 //26.5.15 28800 //24.5.15 added by lucas de beltran
         aWAData[ WA_CONNECTION ]:Open( cConnect )  //2.6.15 SEE TOP OF THIS FUNCTION
         aWAData[ WA_TABLENAME ]  := UPPER(aOpenInfo[ UR_OI_NAME ])
         aWAData[ WA_QUERY ]      := t_cQuery
         aWAData[ WA_USERNAME ]   := t_cUserName
         aWAData[ WA_PASSWORD ]   := t_cPassword
         aWAData[ WA_SERVER ]     := t_cServer
         aWAData[ WA_ENGINE ]     := t_cEngine
         aWAData[ WA_CONNOPEN ]   := .F.
         aWAData[ WA_CATALOG ]    := t_cDatasource

      ENDIF

      aWAData[ WA_TABLENAME ] := CFILENOEXT(CFILENOPATH(aWAData[ WA_TABLENAME ] ))

      IF aWAData[ WA_ENGINE ] = "ADS" .OR.  aWAData[ WA_ENGINE ] = "MSSQL"
         //IF TEMP TABLE SEND IT AS TO PROVIDER
         IF UPPER( SUBSTR( aWAData[ WA_TABLENAME ],1,3) ) = "TMP" .OR. 	UPPER( SUBSTR( aWAData[ WA_TABLENAME ],1,4) ) = "TEMP"
            //temporary tables beahaves strange in loops and updates
            //aWAData[ WA_TABLENAME ] := "#"+aWAData[ WA_TABLENAME ]
            aWAData[ WA_TABLENAME ] := aWAData[ WA_TABLENAME ]
         ENDIF

      ENDIF

  CATCH
      ADOSHOWERROR(aWAData[ WA_CONNECTION ])
      RETURN HB_FAILURE

  END

  //WE DONT NEED IT ANYMORE REINITIALIZE
  t_cEngine    := NIL
  t_cServer    := NIL
  t_cUserName  := NIL
  t_cPassword  := NIL
  t_cQuery     := NIL
  t_cDataSource := NIL  //2.06.15 TO HAVE MULTIPLE DATASORCES OPEN

  RETURN HB_SUCCESS


// 23.06.15 OPEN THE CONNECTION TO ADOGETCONNECT AND ADOCONNECT
STATIC FUNCTION ADOOPENCONNECT( cDB, cServer, cEngine, cUser, cPass , oCn )

  oCn:ConnectionTimeOut := 60 //26.5.15 28800 //24.5.15 added by lucas de beltran

  DO CASE
     CASE cEngine = "DBASE"
         oCn:Open( "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + cDB +;
                   ";Extended Properties=dBASE IV;User ID="+cUser+";Password="+cPass+";" )

     CASE cEngine = "FOXPRO"
         oCn:Open( "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + cDB +;
                   ";Extended Properties=Foxpro;User ID="+cUser+";Password="+cPass+";" )

     CASE cEngine = "ACCESS"
          IF Empty( cPass )
             oCn:Open( "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + cDB  )
           ELSE
              oCn:Open( "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + cDB  + ";Jet OLEDB:Database Password=" + AllTrim( cPass ) )
           ENDIF

     CASE cEngine = "ADS"
          oCn:Open("Provider=Advantage OLE DB Provider;User ID="+cUser +;
                    ";Password="+cPass+";Data Source="+ cDB +";TableType=ADS_VFP;"+;
                    "Advantage Server Type=ADS_LOCAL_SERVER;")

     CASE cEngine == "MYSQL"
          oCn:Open( "Driver={mySQL ODBC 5.3 ANSI Driver};" + ;
                    "server=" + cServer + ;
                    ";Port=3306"+;
                    ";database=" + cDB  + ;
                    ";uid=" + cUser + ;
                    ";pwd=" + cPass+";" )

     CASE  cEngine == "MSSQL"
           oCn:Open( "Provider=SQLOLEDB;" + ;
                     "server=" + cServer + ;
                     ";database=" + cDB  + ;
                     iif(empty(cUser),";Trusted_Connection=yes",;
                         ";uid=" + cUser + ;
                         ";pwd=" + cPass ) )

     CASE cEngine == "ORACLE"
          oCn:Open( "Provider=MSDAORA.1;" + ;
                    "Persist Security Info=False" + ;
                    iif( Empty( cServer ), ;
                        "", ";Data source=" + cServer ) + ;
                        ";User ID=" + cUser + ;
                        ";Password=" + cPass )

     CASE cEngine == "FIREBIRD"  //NOT IMPLEMENTED ARRAY DBENGINES STRUCTS ETC
          oCn:Open( "Driver=Firebird/InterBase(r) driver;" + ;
                    "Persist Security Info=False" + ;
                    ";Uid=" + cUser + ;
                    ";Pwd=" + cPass + ;
                    ";DbName=" + cDB  )

      CASE cEngine == "SQLITE"
           oCn:Open( "Driver={SQLite3 ODBC Driver};" + ;
                     "Database=" + cDB   + ;
                     ";LongNames=0;Timeout=1000;NoTXN=0;SyncPragma=NORMAL;StepAPI=0;"   )

      CASE cEngine == "POSTGRE"
           oCn:Open( "Driver={PostgreSQL};Server="+cServer+";Port=5432;"+;
                     "Database="+ cDB+;
                     ";Uid="+cUser+";Pwd="+cPass+";" )

      CASE cEngine == "INFORMIX"
           oCn:Open( "Dsn='';Driver={INFORMIX 3.30 32 BIT};"+;
                     "Host="+""+";Server="+cServer+";"+;
                     "Service="+""+";Protocol=olsoctcp;"+;
                     "Database="+cDB+";Uid="+cUser+";"+;
                     "Pwd="+cPass+";" )

      CASE cEngine == "ANYWHERE"
           oCn:Open( "Driver={SQL Anywhere 12};"+;
           "Host="+cServer+";Server="+cServer+";port=2638;"+;
           "db="+cDB+;
           iif(empty(cUser),";Trusted_Connection=yes",;
                         ";uid=" + cUser + ;
                         ";pwd=" + cPass ) )

  ENDCASE

  RETURN oCn


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

   ADO_UNLOCK( nWA ) //RELEASE ALL LOCKS IN TLOCKS.DBF

   ADO_OPENSHARED( nWA, aWAData[WA_TABLENAME], .F., .T. )

   //dont close connection as mugh be used by other recorsets
   // need to have all recordsets in same connection to use transactions
   IF !EMPTY( oRecordSet)
      IF oRecordSet:State = adStateOpen
         oRecordSet:Close()

      ENDIF

   ENDIF

   oRecordSet := NIL
   aWAData[ WA_BOF ] := .F.
   aWAData[ WA_EOF ] := .F.
   aWAData[WA_INDEXES] := {}
   aWAData[WA_INDEXEXP] := {}
   aWAData[WA_INDEXFOR] := {}
   aWAData[WA_INDEXACTIVE] := 0
   aWAData[WA_LOCKLIST] := {}
   aWAData[WA_FILELOCK] := .F.
   aWAData[WA_INDEXUNIQUE] := {}
   aWAData[WA_OPENSHARED] := NIL
   aWAData[WA_SCOPES] := {}
   aWAData[WA_SCOPETOP] := {}
   aWAData[WA_SCOPEBOT] := {}
   aWAData[WA_ISITSUBSET] := .F.
   aWAData[WA_FOUND] := .F.
   aWAData[WA_LASTRELKEY] := NIL
   aWAData[WA_FILTERACTIVE] := NIL
   aWAData[WA_FIELDRECNO] := NIL
   aWAData[WA_SCOPEINFO] := NIL
   aWAData[WA_TLOCKS] := {}
   aWAData[WA_FILEHANDLE ] := NIL
   aWAData[WA_LOCKSCHEME ] := ADOFORCELOCKS()  //no lock type 999
   aWAData[WA_CFILTERACTIVE ] := ""
   aWAData[WA_LREQUERY] := .F.
   aWAData[WA_RECORDSET] := NIL //18.6.15 cleaning
   aWAData[WA_RECCOUNT] := NIL //27.06.5
   aWAData[WA_ABOOKMARKS] := {}
   aWAData[WA_INDEXDESC] := {}
   aWAData[WA_FIELDDELETED] := NIL

   RETURN UR_SUPER_CLOSE( nWA )


/*                              RECORD RELATED FUNCTION                   */

STATIC FUNCTION ADO_GET_FIELD_RECNO( cTablename )

  LOCAL cFieldName := ADODEFLDRECNO() //default recno field name
  LOCAL aFiles :=  ListFieldRecno(),n

   IF !EMPTY( aFiles ) //IS THERE A FIELD AS RECNO DIFERENT FOR THIS TABLE
      n := ASCAN( aFiles, { |z| z[1] == cTablename } )
      IF n > 0
         cFieldName := aFiles[n,2]
      ENDIF

   ENDIF

   RETURN cFieldName


STATIC FUNCTION ADO_RECINFO( nWA, nRecord, nInfoType, uInfo )

   LOCAL aWdata := USRRDD_AREADATA( nWA )
   LOCAL nResult := HB_SUCCESS
   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]
   LOCAL n

   HB_SYMBOL_UNUSED( nWA )

#ifdef DBRI_DELETED
   DO CASE
      CASE nInfoType == DBRI_DELETED
          ADO_DELETED( nWA, @uInfo )

      CASE nInfoType == DBRI_LOCKED
           FOR n:= 1 TO LEN(aWdata[ WA_LOCKLIST ])
               IF nRecord = aWdata[ WA_LOCKLIST ][n]
                  uInfo := .T.
                  EXIT
               ENDIF

           NEXT

      CASE nInfoType == DBRI_RECSIZE
           uInfo := 0
           FOR n := 1 TO FCOUNT()
               uInfo += FIELDSIZE(n)
           NEXT

      CASE nInfoType == DBRI_RECNO
           nResult := ADO_RECID( nWA, @uInfo )

      CASE nInfoType == DBRI_UPDATED
           uInfo := .F.

      CASE nInfoType == DBRI_ENCRYPTED
           uInfo := .F.

      CASE nInfoType == DBRI_RAWRECORD
           uInfo := ""

      CASE nInfoType == DBRI_RAWMEMOS
           uInfo := ""

       CASE nInfoType == DBRI_RAWDATA
            nResult := ADO_GOTO( nWA, nRecord )
            uInfo := ""

   ENDCASE
#else
   HB_SYMBOL_UNUSED( nRecord )
   HB_SYMBOL_UNUSED( nInfoType )
   HB_SYMBOL_UNUSED( uInfo )
#endif

   RETURN nResult


STATIC FUNCTION ADO_RECNO( nWA, nRecNo )

 LOCAL aWAData := USRRDD_AREADATA( nWA )
 LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]
 LOCAL nResult := HB_SUCCESS,nRecords


   IF !oRecordSet:Eof() //ADO EOF = CLIPPER LASTREC +1
      ADO_GETVALUE( nWA, aWAData[WA_FIELDRECNO]+1, @nRecNo )

   ELSE
      nRecNo :=  ADORECCOUNT( nWA, oRecordSet )+1 //14.6.15 instead of recordcount

   ENDIF


   RETURN nResult


STATIC FUNCTION ADO_RECID( nWA, nRecNo )

   RETURN ADO_RECNO( nWA, @nRecNo )


STATIC FUNCTION ADO_RECCOUNT( nWA, nRecords )
   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ],nRecNo

   IF oRecordSet:RecordCount() < 0 .OR. oRecordSet:Eof()  //14.6.15 ONLY IF IT NOT EOF OTHERWISE GHOST RECORD
      nRecords := ADORECCOUNT( nWA,oRecordSet ) // AHF SEE FUNCTION FOR EXPLANATION oRecordSet:RecordCount()

   ELSE
      ADO_RECID( nWA,@nRecNo )
      IF nRecNo > oRecordSet:RecordCount()
         nRecords := nRecNo
      ELSE
         nRecords :=  oRecordSet:RecordCount()
      ENDIF

   ENDIF

   RETURN HB_SUCCESS


STATIC FUNCTION ADORECCOUNT(nWA,oRecordSet) //AHF
   LOCAL aAWData := USRRDD_AREADATA( nWA )
   LOCAL oCon := aAWData[WA_CONNECTION]
   LOCAL nCount := 0, cSql:="",oRs := ADOCLASSNEW( "ADODB.Recordset" )//TOleAuto():New("ADODB.Recordset") //OPEN A NEW ONE OTHERWISE PROBLEMS WITH OPEN BROWSES

   IF !ADOCON_CHECK()
      RETURN 0

   ENDIF

   IF ADOEMPTYSET( oRecordSet )
      RETURN nCount

   ENDIF

   //Making it lightning faster
   oRS:CursorLocation := IF(aAWData[ WA_ENGINE ] = "ACCESS", adUseClient, adUseClient) //adUseServer  // adUseClient its slower but has avntages such always bookmaks
   oRs:CursorType     := adOpenForwardOnly
   oRs:LockType       := adLockReadOnly

   // 30.06.15
   IF aAWData[ WA_ENGINE ] = "ACCESS" //6.08.15 ONLY WITH ACCESSIT TAKES LONGER IN BIG TABLES
      cSql := "SELECT MAX("+(ADO_GET_FIELD_RECNO( aAWData[WA_TABLENAME] ))+") FROM "+aAWData[WA_TABLENAME]
   ELSE
      //30.06.15 REPLACED BY RAO NAGES IDEA next incremente key
      cSql := "SELECT `AUTO_INCREMENT` FROM INFORMATION_SCHEMA.TABLES"+;
              " WHERE TABLE_SCHEMA = '"+aAWData[ WA_CATALOG ]+"' AND TABLE_NAME = '"+aAWData[ WA_TABLENAME ]+"'"
   ENDIF

   //LETS COUNT IT
   oRs:open( cSql, oCon )
   IF ADOEMPTYSET( oRs )
      nCount := 0

   ELSE
      nCount := oRs:Fields( 0 ):Value-1

   ENDIF

   oRs:close()

   RETURN nCount


STATIC FUNCTION ADO_REFRESH( nWA, lRequery ) //22.08.15
   STATIC nSecs
   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL nReccount := ADORECCOUNT( nWA, aWAData[ WA_RECORDSET ] )
   LOCAL nRecno, xKey, aSeek, oRs, cSql
   LOCAL n, xExpression:=""

   IF nSecs <> NIL .AND. Seconds()-nSecs < 15
      RETURN HB_SUCCESS


   ENDIF

   lRequery := IF( lRequery == NIL, .F., lRequery )

   IF !aWAData[ WA_RECORDSET ]:Eof() .AND.!aWAData[ WA_RECORDSET ]:bof()
      nRecno := aWAData[ WA_RECORDSET ]:Fields(aWAData[WA_FIELDRECNO]):Value
   ELSE
      nSecs := Seconds()
      RETURN HB_SUCCESS
   ENDIF

   IF VALTYPE( aWAData[ WA_RECCOUNT ] ) == "U" //initialize
      aWAData[ WA_RECCOUNT ] := nReccount
      nSecs := Seconds()
      RETURN HB_SUCCESS

   ELSE
      // NEW RECORDS ADDED BY OTHERS?
      IF aWAData[ WA_RECCOUNT ] < nReccount .OR. lRequery //only new records
         //BUT ONLY IF NEW RECORDS ARE RELATED WITH OUR INDEXKEY
         IF aWAData[ WA_INDEXACTIVE ] > 0
            aSeek := ADOPseudoSeek(nWA,xKey,aWAData)
            cSql := IndexBuildExp(nWA,aWAData[WA_INDEXACTIVE],aWAData,.F.,IF(aSeek[3],aseek[1],aSeek[2] ) )
            oRs := TempRecordSet()
            oRs:CursorType :=   adOpenForwardOnly
            oRs:CursorLocation := adUseClient
            oRs:LockType := adLockReadOnly
            oRs:Open(cSql,aWAData[ WA_CONNECTION ] )
            oRS:Sort := ADO_GET_FIELD_RECNO(  aWAData[ WA_TABLENAME ] )

            IF oRs:RecordCount > 0
               oRs:MoveLast()
               IF oRs:Fields( aWAData[WA_FIELDRECNO] ):Value > aWAData[ WA_RECCOUNT ] //new key
                  aWAData[ WA_RECORDSET ]:Requery()
                  aWAData[ WA_LREQUERY ] := .T. //already requeried set to .f.
                  //re initialize
                  aWAData[ WA_RECCOUNT ] := nReccount
               ENDIF

            ENDIF

            oRs:Close()
            oRs := NIl

         ELSE  //NO INDEXES MUST BE A SMALL TABLE LETS REFESH IT ALL
            aWAData[ WA_RECORDSET ]:Requery()
            aWAData[ WA_RECCOUNT ] := nReccount

         ENDIF

         ADO_GOTO( nWA, nRecNo )

         IF aWAData[ WA_RECORDSET ]:Eof()
            // record does not exist anymore other app delete it!
            THROW( ErrorNew( "ADORDD", 10002, 10002, "Record move "+aWAData [ WA_TABLENAME] ,;
                   STR(  nRecNo )+ "was deleted by other app and ADORDD cant reposition " ) )

         ENDIF

      ENDIF

      nSecs := Seconds()

   ENDIF

   RETURN HB_SUCCESS


STATIC FUNCTION ADO_RESYNC( nWA, oRs ) //22.08.15

 LOCAL aWAData := USRRDD_AREADATA( nWA )
 LOCAL xKey := "", n, xExpression:="", nRecNo

   IF oRs:Bof() .OR. oRs:Eof() //nothing to do
      RETURN HB_SUCCESS
   ENDIF

   TRY
      ADO_RECID( nWA, @nRecNo )
      oRs:Resync( adAffectCurrent, adResyncAllValues )

   CATCH
      //IF WE GET HERE OUR RECORD GONE DELETED
      oRs:Requery()
      ADO_GOTO( nWA, nRecNo )

      IF aWAData[ WA_RECORDSET ]:Eof()
         // record does not exist anymore other app delete it!
         THROW( ErrorNew( "ADORDD", 10002, 10002, "Record move "+aWAData [ WA_TABLENAME] ,;
                STR(  nRecNo )+ "was deleted by other app and ADORDD cant reposition " ) )

      ENDIF

   END

   RETURN HB_SUCCESS


STATIC FUNCTION ADO_GOTO( nWA, nRecord )

   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL nRecNo
   LOCAL oRecordSet := aWAData[ WA_RECORDSET ], oClone

   IF !ADOEMPTYSET(oRecordSet)
      IF !VALTYPE(  aWAData[WA_FIELDRECNO]  ) == "U"
         IF !EMPTY(oRecordset:Filter)
            oClone := oRecordSet:Clone
            oClone:MoveFirst()
            oClone:Find(oRecordSet:Fields(aWAData[WA_FIELDRECNO]):Name+" = "+ALLTRIM(STR(nRecord,10,0)) )
            TRY
              oRecordSet:BookMark := oClone:BookMark

            CATCH
              oRecordSet:MoveLast()
              oRecordSet:MoveNext()
            END
         ELSE
            IF oRecordSet:Supports(adIndex) .AND. oRecordSet:Supports(adSeek)
               oRecordSet:Index := oRecordSet:Fields(aWAData[WA_FIELDRECNO]):Name
               oRecordSet:Seek({ ALLTRIM(STR(nRecord,10,0)) })
            ELSE
               oRecordSet:MoveFirst()
               oRecordSet:Find(oRecordSet:Fields(aWAData[WA_FIELDRECNO]):Name+" = "+ALLTRIM(STR(nRecord,10,0)) )
            ENDIF
            // IF EOF RAISE ERROR
         ENDIF

      ENDIF

      ADO_RECID( nWA, @nRecord )

      IF !EMPTY(aWAData[WA_PENDINGREL]) .AND. PROCNAME(2) <> "ADO_RELEVAL" //ENFORCE REL CHILDS BUT NOT IN A ENDLESS LOOP!
         ADO_FORCEREL( nWA )

      ENDIF

   ENDIF

   aWAData[ WA_BOF ] := oRecordSet:Bof()
   aWAData[ WA_EOF ] := oRecordSet:Eof()


   RETURN HB_SUCCESS



STATIC FUNCTION ADO_GOTOID( nWA, nRecord )
   RETURN ADO_GOTO( nWA, nRecord )


STATIC FUNCTION ADO_GOTOP( nWA )

   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := aWAData[ WA_RECORDSET ]
   LOCAL n:= LEN( aWAData[ WA_SCOPES ] ), lDeleted

   IF !ADOEMPTYSET( oRecordSet )
      ADO_REFRESH( nWA )
      IF n > 0
         n := ASCAN( aWAData[ WA_SCOPES ], aWAData[WA_INDEXACTIVE]  )
         ADO_SEEK( nWa, .T., aWAData[ WA_SCOPETOP ][n], .F. )

      ELSE
         oRecordSet:MoveFirst()

      ENDIF

      ADO_DELETED(nWA,@lDeleted)
      IF aWAData[WA_FILTERACTIVE] <> NIL .OR. ( SET( _SET_DELETED ) .AND. lDeleted )
         ADO_SKIPFILTER(nWA,1)
      ELSE
         IF !EMPTY(aWAData[WA_PENDINGREL]) .AND. PROCNAME(2) <> "ADO_RELEVAL" //ENFORCE REL CHILDS BUT NOT IN A ENDLESS LOOP!
            ADO_FORCEREL( nWA )
         ENDIF

      ENDIF

   ELSE
      aWAData[ WA_EOF ] :=  .T.
      aWAData[ WA_BOF ] :=  .T.

   ENDIF

   aWAData[ WA_EOF ] :=  oRecordSet:Eof()

   RETURN HB_SUCCESS


STATIC FUNCTION ADO_GOBOTTOM( nWA )

   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := aWAData[ WA_RECORDSET ]
   LOCAL n:= LEN( aWAData[ WA_SCOPES ] ), lDeleted

   IF !ADOEMPTYSET( oRecordSet )
      ADO_REFRESH( nWA )
      IF n > 0
         n := ASCAN( aWAData[ WA_SCOPES ], aWAData[WA_INDEXACTIVE]  )
         ADO_SEEK( nWa, .F., aWAData[ WA_SCOPEBOT ][n], .T. )
         IF !FOUND()
            ADO_SEEK( nWa, .T., aWAData[ WA_SCOPEBOT ][n], .T. )
            IF &(INDEXKEY( 0 ) ) < aWAData[ WA_SCOPETOP ][n]
               ADO_SEEK( nWa, .T., aWAData[ WA_SCOPETOP ][n], .F. )
            ENDIF
         ENDIF
      ELSE
         oRecordSet:MoveLast()

      ENDIF

      ADO_DELETED(nWA,@lDeleted)
      IF aWAData[WA_FILTERACTIVE] <> NIL .OR. ( SET( _SET_DELETED ) .AND. lDeleted )
         ADO_SKIPFILTER(nWA,-1)
      ELSE
         IF !EMPTY( aWAData[WA_PENDINGREL] ) .AND. PROCNAME( 2 ) <> "ADO_RELEVAL" //ENFORCE REL CHILDS BUT NOT IN A ENDLESS LOOP!
            ADO_FORCEREL( nWA )
         ENDIF

      ENDIF

   ELSE
      aWAData[ WA_EOF ] :=  .T.
      aWAData[ WA_BOF ] :=  .T.

   ENDIF

   aWAData[ WA_EOF ] :=  oRecordSet:Eof()

   RETURN HB_SUCCESS


STATIC FUNCTION ADO_SKIPFILTER( nWA, nRecords )

 LOCAL aWAData   := USRRDD_AREADATA( nWA )
 LOCAL oRecordSet := aWAData[ WA_RECORDSET ],lDeleted
 LOCAL  nToSkip, nRet := HB_SUCCESS

  IF ADOEMPTYSET(oRecordSet)
      nToSkip := 0
      RETURN HB_SUCCESS //SHOULDNET BE FAILURE?

   ENDIF

   nToSkip := iif( nRecords > 0, 1, iif( nRecords < 0, - 1, 0 ) )

   IF nToSkip != 0
      ADO_DELETED(nWA,@lDeleted)
      DO WHILE ( ( aWAData[ WA_FILTERACTIVE ] <> NIL .AND. !Eval( aWAData[ WA_FILTERACTIVE ] ) );
         .OR. ( Set( _SET_DELETED ) .AND. lDeleted ) );
         .AND. ( ! aWAData[ WA_EOF ] .AND.  !aWAData[ WA_BOF ] )

         IF ! ( ADO_SKIPRAW( nWA, nToSkip, aWAData ) == HB_SUCCESS )
            RETURN HB_FAILURE
         ENDIF
         ADO_DELETED(nWA,@lDeleted)

      ENDDO

      IF ! Empty( aWAData[ WA_PENDINGREL ] )
         ADO_FORCEREL( nWA )
      ENDIF

   ENDIF

   RETURN nRet


STATIC FUNCTION ADO_SKIPRAW( nWA, nToSkip )

 LOCAL aWAData := USRRDD_AREADATA( nWA )
 LOCAL oRecordSet := aWAData[ WA_RECORDSET ]
 LOCAL nResult := HB_SUCCESS,nRecno,lBof,lEof

  IF ADOEMPTYSET(oRecordSet)
      nToSkip := 0
      RETURN HB_SUCCESS //SHOULDNET BE FAILURE?

   ENDIF

   ADO_RECID(nWa,@nRecno)

   IF nToSkip != 0

      IF aWAData[ WA_EOF ]
         IF nToSkip > 0
            RETURN HB_FAILURE

         ELSE
            ADO_GOBOTTOM( nWA )
            ++nToSkip
            aWAData[ WA_EOF ] := .F.

         ENDIF

      ENDIF

      IF ( nToSkip < 0 .AND. oRecordSet:Bof() )
         ADO_GOTOP( nWa )
         aWAData[ WA_BOF ] := .T.
         aWAData[ WA_EOF ] := .F.
         RETURN HB_FAILURE

      ELSEIF nToSkip != 0
         IF ( nToSkip > 0 .AND.!oRecordSet:Eof ) .OR. ;
            ( nToSkip < 0 .AND.!oRecordSet:Bof )
            oRecordSet:Move( nToSkip )

            IF nToSkip > 0
               IF ! ADOSCOPEOK( nWA, 1 )
                  aWAData[ WA_BOF ] := .F.
                  aWAData[ WA_EOF ] := .T.

               ELSE
                  aWAData[ WA_BOF ] := .F.
                  aWAData[ WA_EOF ] := oRecordSet:EOF

               ENDIF


            ELSEIF nToSkip < 0
               IF ! ADOSCOPEOK( nWA, 0 ) .OR. oRecordSet:Bof()
                  ADO_GOTOP( nWA )
                  aWAData[ WA_BOF ] := .T.
                  aWAData[ WA_EOF ] := .F.

               ELSE
                  aWAData[ WA_BOF ] := oRecordSet:BOF//.F.
                  aWAData[ WA_EOF ] := oRecordSet:EOF

               ENDIF

            ENDIF

         ELSE
            aWAData[ WA_EOF ] := oRecordSet:EOF
            aWAData[ WA_BOF ] := oRecordSet:BOF
            RETURN HB_FAILURE

         ENDIF

      ENDIF

   ELSE
      //MAYBE REFRESH THE SET AS IN CLIPPER SKIP 0
      //ENFORCE RELATIONS SHOULD BE BELOW AFTER MOVING TO NEXT RECORD
      IF ! Empty( aWAData[ WA_PENDINGREL ] )
         ADO_FORCEREL( nWA )
      ENDIF

      RETURN HB_FAILURE
   ENDIF

   RETURN nResult


STATIC FUNCTION ADO_BOF( nWA, lBof )

   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]

   lBof := aWAData[ WA_BOF ]

   RETURN HB_SUCCESS


STATIC FUNCTION ADO_EOF( nWA, lEof )

   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]
   LOCAL nResult := HB_SUCCESS

    lEof := aWAData[ WA_EOF ]

   RETURN nResult


STATIC FUNCTION ADO_APPEND( nWA, lUnLockAll )

   LOCAL oRs := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]
   LOCAL aWdata := USRRDD_AREADATA( nWA )
   LOCAL aStruct, cBaseTable, n, oField, u
   LOCAL aCols := {}, aVals := {}
   LOCAL lAdded   := .f.,nRecNo
   LOCAL aLockInfo := ARRAY( UR_LI_SIZE ),oError, aBookMark := {}

    IF !ADOCON_CHECK()
       RETURN HB_FAILURE

    ENDIF

    aStruct := ADOSTRUCT( oRs )
    FOR n := 1 TO LEN( aStruct )
        IF aStruct[ n, 6 ]
           AADD( aCols, aStruct[ n, 1 ] )
           AADD( aVals, HB_DECODE( aStruct[ n, 2 ], 'C', Space( aStruct[ n, 3 ] ), 'D',ADONULL(), 'L', .f., ;
                 'M', SPACE(255), 'm', "", ;
                 'N', If( aStruct[ n, 3 ] == 0, 0, Val( "0." + Replicate( '0', aStruct[ n, 3 ] ) ) ), ;
                 'T', ADONULL(), '' ) )

        ENDIF

    NEXT

    IF ! EMPTY( aCols )
//       TRY
          aLockInfo[ UR_LI_RECORD ] := ADORECCOUNT(nWA,oRs)+1 //GHOST NEXT RECORD TO BE LOCKED
          aLockInfo[ UR_LI_METHOD ] := DBLM_MULTIPLE
          aLockInfo[ UR_LI_RESULT ] := .F.

          IF lUnlockAll
             ADO_UNLOCK(nWA)
          ENDIF

          ADO_LOCK( nWA, aLockInfo )

          IF !aLockInfo[ UR_LI_RESULT ]
             NETERR(.T.)
             lAdded := .F.
             oError := ErrorNew()
             oError:GenCode := EG_APPENDLOCK
             oError:SubCode := 1024
             oError:Description := hb_langErrMsg( EG_APPENDLOCK )
             oError:FileName := aWData[ WA_TABLENAME]
             oError:OsCode := 0 /* TODO */
             oError:CanDefault := .T.
             UR_SUPER_ERROR( nWA, oError )
             BREAK
          ENDIF

          oRs:AddNew( aCols, aVals )
          oRs:Update()
          IF  aWData[ WA_INDEXACTIVE ] > 0 .AND. LEN( aWData[ WA_ABOOKMARKS ] [ aWData[ WA_INDEXACTIVE ] ] ) > 0
              //WITH FOR CONDITION NO RECORDS ARE ADDED ONLY REMOVED IN ADO_PUTVALUE
              //WHEN FOR IS EVALUATED FALSE
              IF EMPTY ( aWData[WA_INDEXFOR][aWData[ WA_INDEXACTIVE ]] )
                 xBook := oRs:BookMark
                 AADD( aWData[ WA_ABOOKMARKS ][aWData[ WA_INDEXACTIVE ]] ,{oRs:BookMark ,&( INDEXKEY( 0 ) ) } )
                 IF UPPER( SUBSTR( aWData[ WA_INDEXDESC ] [ aWData[ WA_INDEXACTIVE ] ],1 ,4 ) ) = "DESC"
                    ASORT( aWData[ WA_ABOOKMARKS ][aWData[ WA_INDEXACTIVE ]], NIL, NIL, { |x,y| x[ 2 ] > y[ 2 ] } )
                 ELSE
                    ASORT( aWData[ WA_ABOOKMARKS ][aWData[ WA_INDEXACTIVE ]], NIL, NIL, { |x,y| x[ 2 ] < y[ 2 ] } )

                 ENDIF

                 aBookMarks := aWData[ WA_ABOOKMARKS ][aWData[ WA_INDEXACTIVE ]]
                 aBookMarks := ARRTRANSPOSE( aBookMarks )[ 1 ]
                 oRs:Filter := aBookMarks
                 oRs:BookMark := xBook
              ENDIF
          ENDIF

          aWData[ WA_EOF ] := oRs:Eof()
          aWData[ WA_BOF ] := oRs:Bof()

          lAdded   := .t.
          NETERR(.F.)
          ADO_RECID( nWA, @nRecNo )
          aWData[ WA_RECCOUNT ] := aLockInfo[ UR_LI_RECORD ]

/*       CATCH
          NETERR(.T.)
          ADOSHOWERROR( aWdata[WA_CONNECTION] )

       END*/

    ELSE
       NETERR(.T.)

    ENDIF

    IF !EMPTY( aWData[WA_PENDINGREL] ) .AND. PROCNAME( 2 ) <> "ADO_RELEVAL" //ENFORCE REL CHILDS BUT NOT IN A ENDLESS LOOP!
       ADO_FORCEREL( nWA )

    ENDIF

   RETURN IF( lAdded, HB_SUCCESS ,HB_FAILURE )


//APPEND FROM AND COPY TO  WHEN SRCAREA OPENED WITH ADORDD
/*
  NEXT NRECORDS LREST NRECORD NOT IMPLEMENTED!
*/
STATIC FUNCTION ADO_TRANS(  nWA, aTransInfo )

 LOCAL aScopeInfo := aTransInfo[3] //SCOPE ARRAY
 LOCAL aFields    := aTransInfo[6] //FIELDPOS EACH AREA EST AND SOURCE {FIELDPOS SOURCE, FIELDPOS DESTINTION}
 LOCAL nFields    := aTransInfo[5] //NFIELDS
 LOCAL SrcArea    := aTransInfo[1] //SOURCE AREA
 LOCAL dstArea    := aTransInfo[2] //DESTINATION AREA
 LOCAL aWAData    := USRRDD_AREADATA( nWA )
 LOCAL oRs := aWAData[ WA_RECORDSET ]
 //LOCAL oRsDst := IF( (DstArea)->(RDDNAME()) = "ADORDD",USRRDD_AREADATA( DstArea  )[ WA_RECORDSET ],NIL)
 LOCAL nRecno, oError, n

 DEFAULT  aScopeInfo[UR_SI_BWHILE] TO { ||.T. }
 DEFAULT  aScopeInfo[UR_SI_BFOR]   TO { ||.T. }

  IF !ADOCON_CHECK()
     RETURN HB_FAILURE
  ENDIF

  SELECT(SrcArea)
  nRecno := RECNO()

  DO WHILE EVAL( aScopeInfo[UR_SI_BWHILE] ) .AND.!(SrcArea)->(EOF()) //!oRs:Eof()
     IF EVAL(aScopeInfo[UR_SI_BFOR])
        (DstArea)->(DBAPPEND())
        FOR n := 1 TO (SrcArea)->(FCOUNT())
            IF (DstArea)->( FIELDTYPE(n) ) <> "+"
               (DstArea)->( FIELDPUT(n, (SrcArea)->(FIELDGET(n)) ) )
             ENDIF

        NEXT

     ENDIF

     DBSKIP() //oRs:MoveNext() TO ENFORCE RELATIONS SCOPES AND FILTERS

  ENDDO

  DBGOTO(nRecNo)

  SELECT( nWA )

  IF PROCNAME( 1 ) == "__DBCOPY"
     (DstArea)->( DBCLOSEAREA() )

  ELSEIF PROCNAME( 1 ) == "__DBAPP"
     (SrcArea)->( DBCLOSEAREA() )

  ENDIF


  RETURN HB_SUCCESS


/*                           END RECORD RELATED FUNCTION                   */

/*                                      DELETE RECALL ZAP PACK                  */
//maromano
STATIC FUNCTION ADO_GET_nFIELD_DELETED( oRs ,cTablename )

  LOCAL cFieldName := ADODEFLDDELETED() //default recno field name
  LOCAL aFiles :=  ListFieldDeleted(),n
  LOCAL aStruct := ADOSTRUCT( oRs ),nField

  IF !EMPTY( aFiles ) //IS THERE A FIELD AS Deleted DIFERENT FOR THIS TABLE
     n := ASCAN( aFiles, { |z| z[1] == cTablename } )
     IF n > 0
        cFieldName := aFiles[n,2]

     ENDIF

  ENDIF

  nField := ASCAN( aStruct, {|x| x[1] == cFieldName })
  IF nField == 0
     nField := NIL

  ENDIF

  RETURN nField


STATIC FUNCTION ADO_GET_FIELD_DELETED( cTablename )

  LOCAL cFieldName := ADODEFLDDELETED() //default recno field name
  LOCAL aFiles :=  ListFieldDeleted(),n

   IF !EMPTY( aFiles ) //IS THERE A FIELD AS Deleted DIFERENT FOR THIS TABLE
      n := ASCAN( aFiles, { |z| z[1] == cTablename } )
      IF n > 0
         cFieldName := aFiles[n,2]
      ENDIF

   ENDIF

   RETURN cFieldName


STATIC FUNCTION ADO_DELETED( nWA, lDeleted )
   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]

    IF !ADOEMPTYSET(oRecordSet)
       ADO_GETVALUE( nWA, aWAData[WA_FIELDDELETED]+1, @lDeleted )
    ENDIF

   RETURN HB_SUCCESS


STATIC FUNCTION ADO_DELETE( nWA )

   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]
   LOCAL lDeleted := .F., oError, xBook, aBookMarks := {},nRecNo

   IF !ADOCON_CHECK() .OR. ADOEMPTYSET( oRecordSet )
      RETURN HB_FAILURE

   ENDIF

   IF ADORECCOUNT( nWA,oRecordSet ) > 0 // AHF SEE FUNCTION FOR EXPLANATION oRecordSet:RecordCount()
      IF !oRecordSet:Eof .AND. !oRecordSet:Bof
         ADO_RECID( nWA, @nRecNo )
         IF ADO_ISLOCKED(aWAData[ WA_TABLENAME],nRecNo, aWAData)
            ADO_PUTVALUE( nWA, aWAData[WA_FIELDDELETED]+1, .T. )
            lDeleted = .T.
         ELSE
            //ERROR UNLOCK
            oError := ErrorNew()
            oError:GenCode := EG_UNLOCKED
            oError:SubCode := 1022
            oError:Description := hb_langErrMsg( EG_UNLOCKED )
            oError:FileName := aWAData[ WA_TABLENAME]
            oError:OsCode := 0 /* TODO */
            oError:CanDefault := .T.
            UR_SUPER_ERROR( nWA, oError )
            RETURN HB_FAILURE

         ENDIF

      ENDIF

   ENDIF

   RETURN IF( lDeleted, HB_SUCCESS, HB_FAILURE )


STATIC FUNCTION ADO_RECALL( nWA )

 LOCAL aWAData := USRRDD_AREADATA( nWA )
 LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]
 LOCAL oError, nRecNo

  IF !ADOCON_CHECK()
     RETURN HB_FAILURE

  ENDIF

  IF !ADOEMPTYSET( oRecordSet )
     ADO_RECID( nWA, @nRecNo )
     IF ADO_ISLOCKED(aWAData[ WA_TABLENAME],nRecNo, aWAData)
        ADO_PUTVALUE( nWA, aWAData[WA_FIELDDELETED]+1, .F. )

     ELSE
        //ERROR UNLOCK
        oError := ErrorNew()
        oError:GenCode := EG_UNLOCKED
        oError:SubCode := 1022
        oError:Description := hb_langErrMsg( EG_UNLOCKED )
        oError:FileName := aWAData[ WA_TABLENAME]
        oError:OsCode := 0 /* TODO */
        oError:CanDefault := .T.
        UR_SUPER_ERROR( nWA, oError )
        RETURN HB_FAILURE

     ENDIF

  ENDIF

  RETURN HB_SUCCESS


STATIC FUNCTION ADO_ZAP( nWA )

   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := aWAData[ WA_RECORDSET ]

   IF !ADOCON_CHECK()
      RETURN HB_FAILURE

   ENDIF

   IF !ADO_ISLOCKED(aWAData[ WA_TABLENAME], 0, aWAData)
      oError := ErrorNew()
      oError:GenCode     := EG_UNLOCKED
      oError:SubCode     := 1022 // EDBF_UNLOCKED
      oError:Description := hb_langErrMsg( EG_UNLOCKED )
      oError:FileName    := aWAData[ WA_TABLENAME ]
      UR_SUPER_ERROR( nWA, oError )
      RETURN HB_FAILURE
   ENDIF

   //AUTOINC FIELDS LIKE RECNO MUST BE RESET
   IF aWAData[ WA_CONNECTION ] != NIL .AND. aWAData[ WA_TABLENAME ] != NIL
      TRY
         aWAData[ WA_CONNECTION ]:Execute( "TRUNCATE TABLE " + aWAData[ WA_TABLENAME ] )

      CATCH
         aWAData[ WA_CONNECTION ]:Execute( "DELETE * FROM " + aWAData[ WA_TABLENAME ] )

      END

      ADO_REFRESH( nWA, .T. )

   ENDIF

   RETURN HB_SUCCESS


STATIC FUNCTION ADO_PACK( nWA )

 LOCAL aWAData := USRRDD_AREADATA( nWA )
 LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ],oError

  IF !ADOCON_CHECK()
     RETURN HB_FAILURE

  ENDIF

  IF !ADO_ISLOCKED(aWAData[ WA_TABLENAME], 0, aWAData)
     oError := ErrorNew()
     oError:GenCode     := EG_UNLOCKED
     oError:SubCode     := 1022 // EDBF_UNLOCKED
     oError:Description := hb_langErrMsg( EG_UNLOCKED )
     oError:FileName    := aWAData[ WA_TABLENAME ]
     UR_SUPER_ERROR( nWA, oError )
     RETURN HB_FAILURE
  ENDIF
  cSql := "DELETE FROM "+  aWAData[ WA_TABLENAME ]  + " WHERE "+ ADO_GET_FIELD_DELETED(aWAData[ WA_TABLENAME])+ " <> 0 "
  aWAData[ WA_CONNECTION ]:Execute( cSql )

  MEMOWRIT( "packsql.txt", cSql )
  ADO_REFRESH( nWA, .T. )

  RETURN HB_SUCCESS
/*                             END OF DELETE RECALL ZAP PACK              */


/*                               FIELD RELATED FUNCTIONS  */
STATIC FUNCTION ADO_GETVALUE( nWA, nField, xValue )

   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL rs := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]
   LOCAL aFieldInfo := ADO_FIELDSTRUCT( Rs, nField-1 ), nRecNo

   //MISSIGNG OLE VARLEN MODTIME ROWVER CURDUBLE FLOAT LONG CURRENCY BLOB IMAGE
   //DONT KNOW DEFAULT VALUES

   IF aWAData[ WA_EOF ] .OR. rs:EOF .OR. rs:BOF
      xValue := NIL

      IF aFieldInfo[7] == HB_FT_STRING
         xValue := Space( aFieldInfo[3] )
      ENDIF

      IF aFieldInfo[7] == HB_FT_DATE
         xValue := CTOD('')
      ENDIF

      IF aFieldInfo[7] == HB_FT_INTEGER .OR.  aFieldInfo[7] ==  HB_FT_DOUBLE
         IF aFieldInfo[4] > 0
            xValue := VAL("0."+REPLICATE("0",VAL(STR(aFieldInfo[4],10,0)) )) //VAL("0."+ALLTRIM(STR(aFieldInfo[4],0)))
         ELSE
            xValue := INT(0)
         ENDIF

      ENDIF

      IF aFieldInfo[7] == HB_FT_AUTOINC
         xValue := INT(0)
      ENDIF

      IF aFieldInfo[7] == HB_FT_MEMO
         xValue := SPACE(0)
      ENDIF

      IF aFieldInfo[7] == HB_FT_LOGICAL
         xValue := .F.
      ENDIF

      IF aFieldInfo[7] == HB_FT_TIMESTAMP
         xValue := CTOT('')
      ENDIF

   ELSE
      TRY
         xValue := rs:Fields( nField - 1 ):Value

      CATCH  //DELETED OR CHANGED RECORDS RECORDSET OUTDATED
         ADO_RECID(nWA, @nRecNo)
         ADO_REFRESH( nWA, .T. )
         ADO_GOTO(nWa, nRecNo) //IF DELETED (not found) GOES EOF THEN TRYING TO LOCK IT WILL FAIL NO UPDATE POSSIBLE

         ADO_GETVALUE( nWA, nField, xValue )

      END

      IF aFieldInfo[7] == HB_FT_STRING
         IF VALTYPE( xValue ) == "U"
            xValue := SPACE( rs:Fields( nField - 1 ):DefinedSize )
         ELSE
            xValue := PADR( xValue, rs:Fields( nField - 1 ):DefinedSize )
         ENDIF

      ENDIF

      IF aFieldInfo[7] == HB_FT_DATE
         IF VALTYPE( xValue ) == "U" .OR. FW_TTOD( xValue ) == {^ 1899/12/30 }
            xValue := CTOD('')
         ELSE
            xValue := FW_TTOD( xValue ) //24.06.15 WAS DISABLE BY MISTAKE DURING TRIALS
         ENDIF

      ENDIF

      IF aFieldInfo[7] == HB_FT_INTEGER .OR.  aFieldInfo[7] ==  HB_FT_DOUBLE
         IF VALTYPE( xValue ) == "U"
            IF aFieldInfo[4] > 0
              xValue := VAL("0."+REPLICATE("0",VAL(STR(aFieldInfo[4],10,0)) )) // ALLTRIM(STR(aFieldInfo[4],0)))
            ELSE
               xValue := INT(0)
            ENDIF

         ELSE
            IF aFieldInfo[4] = 0
               xValue := INT(xValue)
             ENDIF

         ENDIF

      ENDIF

      IF aFieldInfo[7] == HB_FT_AUTOINC
         xValue := INT(xValue)

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
         IF VALTYPE( xValue ) == "U"
            xValue := CTOT('')
         ELSE
            xValue := FW_TTOD(xValue)
         ENDIF

      ENDIF

   ENDIF


   RETURN HB_SUCCESS


STATIC FUNCTION ADO_PUTVALUE( nWA, nField, xValue )

   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := aWAData[ WA_RECORDSET ]
   LOCAL nRecNo, aOrderInfo  := ARRAY(UR_ORI_SIZE),oError,cDateFormat
   //10.08.15
   LOCAL aStruct := ADO_FIELDSTRUCT( oRecordSet, nField-1 ),xBook
   LOCAL aBookMarks := {}, bForExp, nPos

   IF !ADOCON_CHECK()
      RETURN HB_FAILURE

   ENDIF

   //DELETED FIELD ONLY UPDATABLE BY RECALL OR DBDELETE
   IF nField-1 = aWAData[WA_FIELDDELETED] .AND.( PROCNAME(1) <> "ADO_DELETE";
          .AND. PROCNAME(1) <> "ADO_RECALL" )
      RETURN HB_SUCCESS
   ENDIF

   IF ! aWAData[ WA_EOF ] .AND. !( oRecordSet:Fields( nField - 1 ):Value == xValue )
      //CHECK IF THE FIELD CA BE RW
      IF aStruct[6]
         IF aStruct[2] = "N"
            IF LEN(CVALTOCHAR(xValue)) > oRecordSet:Fields( nField - 1 ):Precision
               //round to the numericscale
               xValue := ROUND(xValue,oRecordSet:Fields( nField - 1 ):NumericScale)

               IF LEN(CVALTOCHAR(xValue)) > oRecordSet:Fields( nField - 1 ):Precision
                  oError := ErrorNew()
                  oError:GenCode := EG_DATAWIDTH
                  oError:SubCode := 1021
                  oError:subSystem := "ADORDD"
                  oError:Description := oRecordSet:Fields( nField - 1 ):Name + hb_langErrMsg( EG_DATAWIDTH )
                  oError:FileName := aWAData[ WA_TABLENAME]
                  oError:OsCode := 0 /* TODO */
                  oError:CanDefault := .T.
                  UR_SUPER_ERROR( nWA, oError )
                  RETURN HB_SUCCESS //TO CONTINUE WITH PROCESS
               ENDIF

            ENDIF

         ENDIF

         ADO_RECID(nWa,@nRecNo)

         IF ADO_ISLOCKED(aWAData[ WA_TABLENAME],nRecNo,aWAData)
            //DEFAULT DBF BEHAVIOUR TRUNCATE EXCCEDING CHARATERS
            IF aStruct[2] = "C" .OR. aStruct[2] = "M"
               xValue := SUBSTR(xValue,1,oRecordSet:Fields( nField - 1 ):DefinedSize)
            ENDIF

            IF aStruct[2] $ "DT" .AND. (EMPTY(xValue) .OR. FW_TTOD( xValue ) == {^ 1899/12/30 })
               //IF DATE IS EMPTY FIELD VALUE CAN BE "U" UPDATING IT IN THIS STATE ERRORS
               IF EMPTY(xValue) .AND. VALTYPE( oRecordSet:Fields( nField - 1 ):Value ) == "U"
                  RETURN HB_SUCCESS
               ENDIF

               xValue := ADONULL()

            ENDIF

            IF xValue == NIL
               xValue := ADONULL()
            ENDIF

            IF aStruct[2] $ "DT"
               cDateFormat := SET( _SET_DATEFORMAT )
               //IF oRecordSet:Fields( nField - 1 ):Type = adDBDate
               SET DATE FORMAT TO "YYYY-MM-DD"
               //ENDIF
            ENDIF

            //XhARBOUR HAS SOME PROBLEMS WITH DATES WITH THIS 100% OK VALTYPE( xValue )   <> "O" //ADONULL
            //VERSION XHARBOUR
            IF aStruct[2] $ "DT" //14.6.15 .AND.  VALTYPE( xValue )   <> "O"
               aWAData[ WA_CONNECTION ]:Execute( "UPDATE "+aWAData[ WA_TABLENAME ]+" SET "+;
               ADOQUOTEDCOLSQL( Trim( oRecordSet:Fields( nField - 1 ):Name ), ;
                                aWAData[ WA_ENGINE ] ) + " = " +;
               IF( VALTYPE( xValue )   <> "O", "'"+CVALTOCHAR( xValue )+"'", 'NULL' )+;
                   " WHERE " + ADOQUOTEDCOLSQL( Trim( oRecordSet:Fields(aWAData[WA_FIELDRECNO]):Name ),;
                                                aWAData[ WA_ENGINE ] )+" = "+ALLTRIM( STR( nRecNo, 10, 0 ) ) )

               //22.08.15
               //if we are in a diferent record of what we should ex was deleted by others
               // if we are with lock on it an not in :update
               // otherwise it woulnt be updated with UPDATE it do nothing
               IF PROCNAME( 1 ) <> "__DBCOPY"
                  ADO_RESYNC( nWA, oRecordSet )
               ENDIF

            ELSE
               IF 1<>1
                  //DEACTIVATED BECAUSE SOMETIMES IN A LOOP GIVES ERROR STATUS 2048
                  //CONCURRENCY WITHOUT ANYONE ELSE ITH THIS ROW!
                  //aWAData[ WA_LOCKSCHEME ] //.OR. ;// 18.06.15 to work safely without locks
                  //PROCNAME( 1 ) = "__DBCOPY" //11.08.15 copy below errors dont know why
                  oRecordSet:Fields( nField - 1 ):Value := xValue
                  TRY
                     oRecordSet:Update()
                  CATCH
                     ADOSTATUSMSG( oRecordSet:Status,oRecordSet:Fields( nField - 1 ):Name,;
                      oRecordSet:Fields( nField - 1 ):Value, xValue, oRecordSet:LockType )
                  END
               ELSE // 18.06.15 to work safely without locks
                  DO CASE
                     CASE aStruct[2] = "L"
                          IF( xValue, xValue := 1, xValue := 0 )
                          xValue := CVALTOCHAR( xValue )
                     CASE aStruct[2] = "N"
                          xValue := CVALTOCHAR( xValue )
                     OTHERWISE
                          xValue := "'"+xValue+"'"
                   ENDCASE

                   aWAData[ WA_CONNECTION ]:Execute( "UPDATE "+aWAData[ WA_TABLENAME ]+" SET "+;
                       ADOQUOTEDCOLSQL( Trim( oRecordSet:Fields( nField - 1 ):Name ),;
                       aWAData[ WA_ENGINE ] )  + " = " + xValue +" WHERE "+;
                       ADOQUOTEDCOLSQL( Trim( oRecordSet:Fields(aWAData[WA_FIELDRECNO]):Name ),;
                                        aWAData[ WA_ENGINE ] )+" = "+ALLTRIM( STR( nRecNo, 11, 0 ) ))

                   //22.08.15
                   IF PROCNAME( 1 ) <> "__DBCOPY"
                      ADO_RESYNC( nWA, oRecordSet )
                   ENDIF

               ENDIF

            ENDIF

            IF  aWAData[ WA_INDEXACTIVE ] > 0 .AND. LEN( aWAData[ WA_ABOOKMARKS ] [ aWAData[ WA_INDEXACTIVE ] ] ) > 0
                xBook := oRecordSet:BookMark
                nPos := ASCAN( aWAData[ WA_ABOOKMARKS ][aWAData[WA_INDEXACTIVE]],;
                              {|x|  x[ 1 ] = xBook  } )
                IF aWAData[WA_INDEXFOR][aWAData[ WA_INDEXACTIVE ]] <> ""
                   bForExp :=  &("{ ||"+ aWAData[WA_INDEXFOR][aWAData[ WA_INDEXACTIVE ]] +"}" )
                   IF !EVAL( bForExp )
                      ADEL( aWAData[ WA_ABOOKMARKS ] [ aWAData[ WA_INDEXACTIVE ]], nPos, .T. )
                   ENDIF
                ELSE
                   //as soon as we alter expresson of the record
                   // ONLY TO DO IF FIELD IN ORDER BY
                   IF ALLTRIM(oRecordSet:Fields( nField - 1 ):Name) $  INDEXKEY( 0 )
                      xBook := oRecordSet:BookMark
                      aWAData[ WA_ABOOKMARKS ][aWAData[ WA_INDEXACTIVE ]][ 2 ] := &( INDEXKEY( 0 ) )
                      IF UPPER( SUBSTR( aWAData[ WA_INDEXDESC ] [ aWAData[ WA_INDEXACTIVE ] ],1 ,4 ) ) = "DESC"
                         ASORT( aWAData[ WA_ABOOKMARKS ][aWAData[ WA_INDEXACTIVE ]], NIL, NIL, { |x,y| x[ 2 ] > y[ 2 ] } )

                      ELSE
                         ASORT( aWAData[ WA_ABOOKMARKS ][aWAData[ WA_INDEXACTIVE ]], NIL, NIL, { |x,y| x[ 2 ] < y[ 2 ] } )

                      ENDIF

                      aBookMarks := aWAData[ WA_ABOOKMARKS ][aWAData[ WA_INDEXACTIVE ]]
                      aBookMarks := ARRTRANSPOSE( aBookMarks )[ 1 ]
                      oRecordSet:Filter := aBookMarks
                      oRecordSet:BookMark := xBook

                   ENDIF
                ENDIF
            ENDIF

            IF aStruct[2] $ "DT"
               SET( _SET_DATEFORMAT ,cDateFormat)
            ENDIF

         ELSE
            //ERROR UNLOCK
            oError := ErrorNew()
            oError:GenCode := EG_UNLOCKED
            oError:SubCode := 1022
            oError:Description := hb_langErrMsg( EG_UNLOCKED )
            oError:FileName := aWAData[ WA_TABLENAME]
            oError:OsCode := 0 /* TODO */
            oError:CanDefault := .T.
            UR_SUPER_ERROR( nWA, oError )
            RETURN HB_FAILURE

         ENDIF

      ELSE
         RETURN HB_SUCCESS

      ENDIF

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

          CASE nType == HB_FT_TIMESTAMP
               uInfo := "T"

          CASE nType == HB_FT_INTEGER .OR. nType == HB_FT_DOUBLE //HB_FT_LONG
               uInfo := "N"

  /*      CASE nType == HB_FT_INTEGER
               uInfo := "I"

          CASE nType == HB_FT_DOUBLE
               uInfo := "B"
*/
          CASE nType == HB_FT_AUTOINC
               uInfo := "+"
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


STATIC FUNCTION ADO_FIELDSTRUCT( oRs, n ) // ( oRs, nFld ) where nFld is 1 based
                                    // ( oRs, oField ) or ( oRs, cFldName )
                                    // ( oField )

   LOCAL oField, nType, uval
   LOCAL cType := 'C', nLen := 10, nDec := 0, lRW := .t.,nDBFFieldType :=  HB_FT_STRING // default
   LOCAL nFWAdoMemoSizeThreshold := 255

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

      IF oRs != nil .AND. ! oRs:Eof() .AND. ! oRs:Bof() .AND. VALTYPE( uVal := oField:Value ) == 'T'
         //.AND. FW_TIMEPART( uVal ) >= 1.0 WHERE IS THIS FUNCTION?
         cType      := 'T'
         nDBFFieldType := HB_FT_TIMESTAMP // DONT KNWO IF IT IS CORRECT!
      ELSE
         nDBFFieldType := HB_FT_DATE
      ENDIF

   ELSEIF ASCAN( { adTinyInt, adSmallInt, adInteger, adBigInt, ;
                  adUnsignedTinyInt, adUnsignedSmallInt, adUnsignedInt, ;
                  adUnsignedBigInt }, nType ) > 0
      cType    := 'N'
      nLen     := oField:Precision //+ 1  // added 1 for - symbol
      nDBFFieldType := HB_FT_INTEGER

      TRY
         IF oField:Properties( "ISAUTOINCREMENT" ):Value == .t. //IN SOME STATES (WHERE CLAUSE) DONT KNOW WHY THIS ERRORS
            cType := '+'
            lRW   := .f.
            nDBFFieldType := HB_FT_AUTOINC
         ENDIF

      CATCH
         IF oField:name = oRs:Fields( USRRDD_AREADATA( SELECT() )[WA_FIELDRECNO] ):name //DEFINED IN ADO_OPEN FIELD RECNO
            cType := '+'
            lRW   := .f.
            nDBFFieldType := HB_FT_AUTOINC
         ENDIF

      END

   ELSEIF ASCAN( { adSingle, adDouble }, nType ) > 0
      cType    := 'N'
      nLen     := Min( 19, oField:Precision  ) //Max( 19, oField:Precision + 2 )
      nDBFFieldType := HB_FT_DOUBLE
      nDec  := 2
      /*
      IF oField:NumericScale > 0 .AND. oField:NumericScale < nLen
         nDec  := oField:NumericScale
         nDBFFieldType :=  HB_FT_DOUBLE //HB_FT_INTEGER WICH ONE IS CORRECT?
      ENDIF
      */

   ELSEIF ASCAN( { adCurrency }, nType ) > 0
      cType    := 'N'      // 'Y'
      nLen     := Min( 19, oField:Precision  )  //19
      nDec     := 2
      nDBFFieldType :=  HB_FT_DOUBLE

   ELSEIF ASCAN( { adDecimal, adNumeric, adVarNumeric }, nType ) > 0
      cType    := 'N'
      nLen     :=  Min( 19, oField:Precision  ) //Max( 19, oField:Precision + 2 )
      nDBFFieldType := HB_FT_INTEGER

      IF oField:NumericScale > 0 .AND. oField:NumericScale < nLen
         nDec  := oField:NumericScale
         nDBFFieldType :=  HB_FT_DOUBLE
      ENDIF

   ELSEIF ASCAN( { adBSTR, adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar }, nType ) > 0
      nLen     := oField:DefinedSize
      nDBFFieldType := HB_FT_STRING
      cType := "C"

      IF nType != adChar .AND. nType != adWChar .AND. nLen > nFWAdoMemoSizeThreshold
         cType := 'M'
         nLen  := 10
         nDBFFieldType := HB_FT_MEMO
      ENDIF

   ELSEIF ASCAN( { adBinary, adVarBinary, adLongVarBinary }, nType ) > 0
      nLen     := oField:DefinedSize
      IF nType != adBinary .AND. nLen > nFWAdoMemoSizeThreshold
         cType := 'm'
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
      nDBFFieldType := HB_FT_TIMESTAMP

/*   ELSEIF ASCAN( { adEmpty, adError, adUserDefined, adIDispatch  }, nType ) > 0

      cType = 'O'
      lRw := .t.
      nDBFFieldType := HB_FT_NONE //what is this? maybe NONE is wrong!
*/
   ELSE
      lRW      := .f.

   ENDIF
   /* DOESNT WORK
    IF lAnd( oField:Attributes, 0x72100 ) .OR. ! lAnd( oField:Attributes, 8 )
      lRW      := .f.
   ENDIF
   */

   RETURN { oField:Name, cType, nLen, nDec, nType, lRW, nDBFFieldType }


FUNCTION ADOSTRUCT( oRs )

   LOCAL aStruct  := {}
   LOCAL n

   FOR n := 0 TO oRs:Fields:Count()-1
      AADD( aStruct, ADO_FIELDSTRUCT( oRs, n ) )

   NEXT

RETURN aStruct
/*                          END FIELD RELATED FUNCTIONS  */


/*                                 INDEX RELATED FUNCTIONS  */
STATIC FUNCTION ADO_INDEXAUTOOPEN(cTableName)

  LOCAL aFiles := ListIndex(),y,z,nOrder := 0, nMax

  //TEMPORARY INDEXES NOT ICLUDED HERE
  //NORMALY ITS CREATED A THEN OPEN?

    y:=ASCAN( aFiles, { |z| z[1] == cTablename } )

    IF y >0
       nMax := LEN(aFiles[y])-1

       FOR z :=1 TO LEN( aFiles[y]) -1
           ORDLISTADD( aFiles[y,z+1,1] )
       NEXT

       IF SET(_SET_AUTORDER) > 0 //11.08.15 > 1
          SET ORDER TO SET(_SET_AUTORDER)

       ELSE //11.08.15 DEFAULT FIRST INDEX GETS CNTROLING ORDER BECAUSE ORDLSTADOENST DO
            //ANYTHINK CALLED FROM HERE
          SET ORDER TO 1
       ENDIF

    ENDIF

   RETURN HB_SUCCESS


STATIC FUNCTION ADO_ORDINFO( nWA, nIndex, aOrderInfo )

   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := aWAData[ WA_RECORDSET ]
   LOCAL nResult := HB_SUCCESS
   LOCAL cExp:="",cIndexExp := "",n
   LOCAl xOrderinfo := aOrderInfo[ UR_ORI_TAG ] //to leave it with same value

   //EMPTY ORDER CONSIDERED 0 CONROLING ORDER
   IF EMPTY(aOrderInfo[ UR_ORI_TAG ])
      aOrderInfo[ UR_ORI_TAG ] := 0

   ENDIF

   // IF ITS STRING CONVERT TO NUMVER
   IF VALTYPE(aOrderInfo[ UR_ORI_TAG ]) = "C"
      n := ASCAN(aWAData[ WA_INDEXES ],aOrderInfo[ UR_ORI_TAG])

      IF n > 0
         aOrderInfo[ UR_ORI_TAG ] := n
      ELSE
         aOrderInfo[ UR_ORI_TAG ] := 0  //NOT FOUND ITS CONTROLING INDEX
      ENDIF

   ENDIF

   //IF  ZERO = CONTROLING ORDER
   IF VALTYPE(aOrderInfo[ UR_ORI_TAG ]) = "N" .AND. aOrderInfo[ UR_ORI_TAG ] = 0
      aOrderInfo[ UR_ORI_TAG ] := aWAData[ WA_INDEXACTIVE ] //MIGHT CONTINUE ZERO IF NO INDEX ACTIVE

   ENDIF

   DO CASE
      CASE nIndex == DBOI_EXPRESSION

           IF ! Empty( aWAData[ WA_INDEXEXP ] ) .AND. aOrderInfo[ UR_ORI_TAG ] <= len(aWAData[ WA_INDEXEXP ])

              IF  aOrderInfo[ UR_ORI_TAG ] = 0  //CONTROLING INDEX NO ACTIVE INDEX SEE ABOVE
                  aOrderInfo[ UR_ORI_RESULT ] := ""
              ELSE
                 aOrderInfo[ UR_ORI_RESULT ] := aWAData[ WA_INDEXEXP ][aOrderInfo[ UR_ORI_TAG]]
                 //STRIPPING OUT INVALID EXPRESSION FOR DBFI NDEX EXPRESSION
                 //21.5.15 aOrderInfo[ UR_ORI_RESULT ] := STRTRAN(aOrderInfo[ UR_ORI_RESULT ] , ",","+")

                 //CONVERT TO CLIPPER EXPRESSION OTHERWISE DIFERENT FILED TYPES TYPES WILL RAISE
                 //ERROR IN THE APP CODE IN EVALUATING WITH &()
                 IF SUBSTR(PROCNAME(1),1,4) <> "ADO_" .AND. PROCNAME(1) <> "INDEXBUILDEXP" .AND. PROCNAME(1) <> "FILTER2SQL"
                    aOrderInfo[ UR_ORI_RESULT ] := KeyExprConversion( aWAData[ WA_INDEXES ][aOrderInfo[ UR_ORI_TAG]],;
                                                                  aWAData[WA_TABLENAME] )[1]
                 ENDIF
              ENDIF

           ELSE
              aOrderInfo[ UR_ORI_RESULT ] := ""

           ENDIF

      CASE nIndex == DBOI_CONDITION
           IF ! Empty( aWAData[ WA_INDEXFOR ] ) .AND. aOrderInfo[ UR_ORI_TAG ] <= LEN(aWAData[ WA_INDEXFOR ])

              IF  aOrderInfo[ UR_ORI_TAG ] = 0  //CONTROLING INDEX NO ACTIVE INDEX SEE ABOVE
                  aOrderInfo[ UR_ORI_RESULT ] := ""
              ELSE
                 aOrderInfo[ UR_ORI_RESULT ] := aWAData[ WA_INDEXFOR ][aOrderInfo[ UR_ORI_TAG]]
                 //STRIPPING OUT INVALID EXPRESSION FOR DBF INDEX FOR EXPRESSION
                 aOrderInfo[ UR_ORI_RESULT ] := STRTRAN(aOrderInfo[ UR_ORI_RESULT ] , "WHERE","FOR")

                 //CONVERT TO CLIPPER EXPRESSION OTHERWISE DIFERENT FILED TYPES TYPES WILL RAISE
                 //ERROR IN THE APP CODE IN EVALUATING WITH &()
                 IF SUBSTR(PROCNAME(1),1,4) <> "ADO_" .AND. PROCNAME(1) <> "INDEXBUILDEXP" .AND. PROCNAME(1) <> "FILTER2SQL"
                    aOrderInfo[ UR_ORI_RESULT ] := KeyExprConversion( aWAData[ WA_INDEXES ][aOrderInfo[ UR_ORI_TAG]],;
                                                                     aWAData[WA_TABLENAME] )[2]
                 ENDIF
              ENDIF

           ELSE
              aOrderInfo[ UR_ORI_RESULT ] :=""

           ENDIF

   CASE nIndex == DBOI_NAME
        IF VALTYPE(aOrderInfo[ UR_ORI_TAG ]) = "N"

           IF ! Empty( aWAData[ WA_INDEXES ] ) .AND. aOrderInfo[ UR_ORI_TAG ] <= LEN(aWAData[ WA_INDEXES ])

              IF  aOrderInfo[ UR_ORI_TAG ] = 0  //CONTROLING INDEX NO ACTIVE INDEX SEE ABOVE
                  aOrderInfo[ UR_ORI_RESULT ] := ""
              ELSE
                 aOrderInfo[ UR_ORI_RESULT ] := aWAData[ WA_INDEXES ][aOrderInfo[ UR_ORI_TAG]]
              ENDIF

           ELSE
              aOrderInfo[ UR_ORI_RESULT ] := ""
           ENDIF

        ELSE
            n := ASCAN(aWAData[ WA_INDEXES ],aOrderInfo[ UR_ORI_TAG])
            IF n > 0
               aOrderInfo[ UR_ORI_RESULT ] := aWAData[ WA_INDEXES ][n]
            ELSE
               aOrderInfo[ UR_ORI_RESULT ] := ""
            ENDIF

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
        IF VALTYPE(aOrderInfo[ UR_ORI_TAG ]) = "N"

           IF ! Empty( aWAData[ WA_INDEXES ] ) .AND. aOrderInfo[ UR_ORI_TAG ] <= LEN(aWAData[ WA_INDEXES ])
              IF  aOrderInfo[ UR_ORI_TAG ] = 0  //CONTROLING INDEX NO ACTIVE INDEX SEE ABOVE
                 aOrderInfo[ UR_ORI_RESULT ] := ""
              ELSE
                 aOrderInfo[ UR_ORI_RESULT ] := aWAData[ WA_INDEXES ][aOrderInfo[ UR_ORI_TAG]]
              ENDIF
           ELSE
              aOrderInfo[ UR_ORI_RESULT ] := ""
           ENDIF

        ELSE
           n := ASCAN(aWAData[ WA_INDEXES ],aOrderInfo[ UR_ORI_TAG])
           IF n > 0
              aOrderInfo[ UR_ORI_RESULT ] := aWAData[ WA_INDEXES ][n]
           ELSE
              aOrderInfo[ UR_ORI_RESULT ] := ""
           ENDIF

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
        IF ! Empty( aWAData[ WA_INDEXFOR ] ) .AND. aOrderInfo[ UR_ORI_TAG ] <= LEN(aWAData[ WA_INDEXFOR ])

           IF  aOrderInfo[ UR_ORI_TAG ] = 0  //CONTROLING INDEX NO ACTIVE INDEX SEE ABOVE
               aOrderInfo[ UR_ORI_RESULT ] := ""
           ELSE
              aOrderInfo[ UR_ORI_RESULT ] := !EMPTY(aWAData[ WA_INDEXFOR ][aOrderInfo[ UR_ORI_TAG]])
           ENDIF

        ELSE
           aOrderInfo[ UR_ORI_RESULT ] :=.F.

        ENDIF

   CASE nIndex == DBOI_ISDESC
        IF ! Empty( aWAData[ WA_INDEXDESC ] ) .AND. aOrderInfo[ UR_ORI_TAG ] <= LEN(aWAData[ WA_INDEXDESC ])

           IF  aOrderInfo[ UR_ORI_TAG ] = 0  //CONTROLING INDEX NO ACTIVE INDEX SEE ABOVE
               aOrderInfo[ UR_ORI_RESULT ] := .F.
           ELSE
              aOrderInfo[ UR_ORI_RESULT ] := !EMPTY(aWAData[ WA_INDEXDESC ][aOrderInfo[ UR_ORI_TAG]])
           ENDIF

        ELSE
           aOrderInfo[ UR_ORI_RESULT ] :=.F.

        ENDIF

   CASE nIndex == DBOI_UNIQUE
        IF ! Empty( aWAData[ WA_INDEXUNIQUE ] ) .AND. aOrderInfo[ UR_ORI_TAG ] <= LEN(aWAData[ WA_INDEXUNIQUE ])

           IF  aOrderInfo[ UR_ORI_TAG ] = 0  //CONTROLING INDEX NO ACTIVE INDEX SEE ABOVE
               aOrderInfo[ UR_ORI_RESULT ] := .F.
           ELSE
              aOrderInfo[ UR_ORI_RESULT ] := !EMPTY(aWAData[ WA_INDEXUNIQUE ][aOrderInfo[ UR_ORI_TAG]])
           ENDIF

        ELSE
           aOrderInfo[ UR_ORI_RESULT ] :=.F.

        ENDIF

   CASE nIndex == DBOI_POSITION
        IF aWAData[ WA_CONNECTION ]:State != adStateClosed
           aOrderInfo[ UR_ORI_RESULT ] := oRecordSet:AbsolutePosition()  //22.08.15 WAS THIS A BUG ? +1

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
        IF aWAData[ WA_CONNECTION ]:State != adStateClosed .AND. !ADOEMPTYSET(oRecordSet)
           aOrderInfo[ UR_ORI_RESULT ] := oRecordSet:RecordCount()

        ELSE
           aOrderInfo[ UR_ORI_RESULT ] := 0
           nResult := HB_FAILURE

        ENDIF

   CASE nIndex == DBOI_SCOPESET .OR. nIndex == DBOI_SCOPEBOTTOM .OR. nIndex == DBOI_SCOPEBOTTOMCLEAR ;
        .OR. nIndex == DBOI_SCOPECLEAR .OR. nIndex == DBOI_SCOPETOP .OR. nIndex == DBOI_SCOPETOPCLEAR

        aOrderInfo[ UR_ORI_RESULT ] := ADOSCOPE(nWA, AWAData,oRecordset, aOrderInfo,nIndex)

   CASE nIndex == DBOI_FULLPATH
        aOrderInfo[ UR_ORI_RESULT ] := aWAData[ WA_INDEXES ][aOrderInfo[ UR_ORI_TAG]]

   ENDCASE

   aOrderInfo[ UR_ORI_TAG ] := xOrderinfo // leave it the same

   RETURN nResult


STATIC FUNCTION ADOSCOPE( nWA, aWAdata, oRecordSet, aOrderInfo, nIndex )
 LOCAL y, cScopeExp :="", cSql :=""

   IF !ADOCON_CHECK()
      RETURN HB_FAILURE
   ENDIF

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
       CASE nIndex == DBOI_SCOPESET //never gets called not tested might be completly wrong!
            IF y > 0
               aWAData[ WA_SCOPETOP ][y] := aOrderInfo[ UR_ORI_NEWVAL ]
               aWAData[ WA_SCOPEBOT ][y] := aOrderInfo[ UR_ORI_NEWVAL ]

            ELSE
               AADD( aWAData[ WA_SCOPES ], aWAData[ WA_INDEXACTIVE ] )
               AADD( aWAData[ WA_SCOPETOP ], aOrderInfo[ UR_ORI_NEWVAL ] )
               AADD(aWAData[ WA_SCOPEBOT ],aOrderInfo[ UR_ORI_NEWVAL ])

            ENDIF
            aOrderInfo[ UR_ORI_RESULT ] := ""

       CASE nIndex == DBOI_SCOPECLEAR //never gets called noy tested might be completly wrong!
            IF y > 0
               ADEL( aWAData[ WA_SCOPES ], y, .T. )
               ADEL( aWAData[ WA_SCOPETOP ], y, .T. )
               ADEL( aWAData[ WA_SCOPEBOT ], y, .T. )
               y := 0
            ENDIF
            IF aWAData[ WA_EOF ]
               IF !ADOEMPTYSET( oRecordSet )
                  oRecordSet:MoveLast()
               ENDIF
               aWAData[ WA_EOF ] := oRecordSet:Eof()
               aWAData[ WA_BOF ] := oRecordSet:Bof()
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
               //AADD(aWAData[ WA_SCOPEBOT ],SPACE(LEN(CVALTOCHAR(aWAData[ WA_SCOPETOP ][1])))) //THERE INST STILL A SCOPEBOT ARRAYS MUST HAVE  SAME LEN
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
               //AADD(aWAData[ WA_SCOPETOP ],SPACE(LEN(CVALTOCHAR(aWAData[ WA_SCOPEBOT ][1])))) //THERE INST STILL A SCOPETOP ARRAYS MUST HAVE  SAME LEN
               aOrderInfo[ UR_ORI_RESULT ] := ""

            ENDIF

       CASE nIndex == DBOI_SCOPETOPCLEAR
            IF y > 0
               aOrderInfo[ UR_ORI_RESULT ] := aWAData[ WA_SCOPETOP ][y] //RETURN ACTUALSCOPE TOP
               ADEL( aWAData[ WA_SCOPETOP ], y, .T. )
            ELSE
               aOrderInfo[ UR_ORI_RESULT ] := "" //RETURN ACTUALSCOPE TOP IF NONE

            ENDIF

       CASE nIndex == DBOI_SCOPEBOTTOMCLEAR
            IF y > 0
               aOrderInfo[ UR_ORI_RESULT ] := aWAData[ WA_SCOPEBOT ][y] //RETURN ACTUALSCOPE TOP
               ADEL( aWAData[ WA_SCOPEBOT ], y, .T. )
               ADEL( aWAData[ WA_SCOPES ], y, .T. )
            ELSE
               aOrderInfo[ UR_ORI_RESULT ] := "" //RETURN ACTUALSCOPE TOP IF NONE
            ENDIF

   ENDCASE

  RETURN HB_SUCCESS


STATIC FUNCTION ADOSCOPEOK( nWA, nOption )

 LOCAL aWAData   := USRRDD_AREADATA( nWA )
 LOCAL lRetval := .T., y

  IF LEN( aWAData[ WA_SCOPES ] ) > 0
     y:=ASCAN( aWAData[ WA_SCOPES ], aWAData[WA_INDEXACTIVE]  )

     IF aWAData[ WA_SCOPETOP ][y]  = aWAData[ WA_SCOPEBOT ][y]
        lRetval := ! ( SUBSTR( &( INDEXKEY( 0 ) ), 1, LEN(aWAData[ WA_SCOPETOP ][y] ) ) <>  aWAData[ WA_SCOPETOP ][y] )

     ELSE
        IF nOption = 0 //top
           lRetval := ! ( SUBSTR( &( INDEXKEY( 0 ) ), 1, LEN(aWAData[ WA_SCOPETOP ][y] ) ) <  aWAData[ WA_SCOPETOP ][y] )

        ELSEIF nOption = 1 //bott
           lRetval := ! ( SUBSTR( &( INDEXKEY( 0 ) ), 1, LEN(aWAData[ WA_SCOPEBOT ][y] ) ) >  aWAData[ WA_SCOPEBOT ][y] )

        ENDIF

     ENDIF

  ENDIF

  RETURN lRetVal


STATIC FUNCTION ADO_ORDLSTFOCUS( nWA, aOrderInfo )

   LOCAL nRecNo
   LOCAL aWAData    := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := aWAData[ WA_RECORDSET ]
   LOCAL cSql:="" ,n
   LOCAL cCondition, lUnique, lDesc, cExpression, aInfo, aBookMarks := {}
   LOCAL nActual := aWAData[ WA_INDEXACTIVE ]
   LOCAL aTempFiles := ListTmpNames()

   HB_SYMBOL_UNUSED( nWA )
   HB_SYMBOL_UNUSED( aOrderInfo )

   IF !ADOCON_CHECK()
      RETURN HB_FAILURE

   ENDIF

   ADO_RECID(nWA,@nRecno)

   IF aOrderInfo[ UR_ORI_TAG ] <> NIL

      IF VALTYPE(aOrderInfo[ UR_ORI_TAG ]) = "C"
         //MAYBE IT COMES WITH FILE EXTENSION AND PATH
         aOrderInfo[ UR_ORI_TAG ] := CFILENOPATH(aOrderInfo[UR_ORI_TAG])
         aOrderInfo[ UR_ORI_TAG ] := UPPER(CFILENOEXT(aOrderInfo[ UR_ORI_TAG ]))

         n := ASCAN(aWAData[ WA_INDEXES ],UPPER(aOrderInfo[ UR_ORI_TAG ]))
      ELSE
         n := aOrderInfo[ UR_ORI_TAG ]
      ENDIF

      //IF IT IS THE SAME INDEX DO NOTHING
      //IF n = nActual
       //  RETURN HB_SUCCESS
      //ENDIF

      IF n = 0  //PHISICAL ORDER
         aWAData[ WA_INDEXACTIVE ] := 0
         aOrderInfo[ UR_ORI_RESULT ] := ""
         oRecordSet:Filter := ""
         oRecordSet:Sort := ADO_GET_FIELD_RECNO(  aWAData[ WA_TABLENAME ] )
      ELSE
         IF aWAData[ WA_INDEXACTIVE ] > 0
            aOrderInfo[ UR_ORI_RESULT ] := aWAData[ WA_INDEXES ] [aWAData[ WA_INDEXACTIVE ]]
         ELSE
            aOrderInfo[ UR_ORI_RESULT ] := ""
         ENDIF

         aWAData[ WA_INDEXACTIVE ] := n

         aInfo := Array( UR_ORI_SIZE )
         aInfo[ UR_ORI_TAG ]:= n
         ADOGETORDINFO( nWA, DBOI_EXPRESSION, @aInfo )
         cExpression := aInfo[UR_ORI_RESULT]
         ADOGETORDINFO( nWA, DBOI_CONDITION, @aInfo )
         cCondition := aInfo[UR_ORI_RESULT]
         ADOGETORDINFO( nWA, DBOI_UNIQUE, @aInfo )
         lUnique := aInfo[UR_ORI_RESULT]
         ADOGETORDINFO( nWA, DBOI_ISDESC, @aInfo )
         lDesc := aInfo[UR_ORI_RESULT]

         IF ASCAN(aTempFiles,UPPER(SUBSTR(aWAData[ WA_INDEXES ] [aWAData[ WA_INDEXACTIVE ]],1,3)) ) > 0 .OR.;
            ASCAN(aTempFiles,UPPER(SUBSTR(aWAData[ WA_INDEXES ] [aWAData[ WA_INDEXACTIVE ]],1,4)) ) > 0

            //ARRAY DATA WRKAREA ALREADY CREATEED WITH ORCREATE
            oRecordSet:Sort := ""
            aBookMarks := aWAData[ WA_ABOOKMARKS ][n]
            aBookMarks := ARRTRANSPOSE( aBookMarks )[ 1 ]
            oRecordSet:Filter := aBookMarks

         ELSE
            //INDEX FROM SET ADODBF INDEX MIGHT HAVE UDFS OR CONDITIONS
            //MUST BYUILDNOW THE ARRAY ORDERED
            IF ADOUDFINDEX( nWa, cExpression, cCondition, lUnique, lDesc, , , ,@aBookMarks )

               oRecordSet:Sort := ""
               aWAData[ WA_ABOOKMARKS ][n] := aBookMarks
               aBookMarks := ARRTRANSPOSE( aBookMarks )[ 1 ]
               oRecordSet:Filter := aBookMarks

            ELSE
               //NOT TEMP NOR UDFS OR CONDITIONS USE SORT ITS FASTER!
               oRecordSet:Filter := ""
               cSql := IndexBuildExp(nWA,n,aWAData)
               cSql := ALLTRIM( SUBSTR( cSql, AT( "ORDER BY", cSql )+9 ) )
               cSql := cSql
               oRecordSet:Sort :=  cSql

            ENDIF

         ENDIF

      ENDIF

      IF VALTYPE(nRecNo) = "N"
         IF PROCNAME(1) = "ADO_ORDLSTADD" //SET  INDEX TO THEN GO FRIST RECORD
            ADO_GOTOP( nWA )
         ELSE
            ADO_GOTO( nWA, nRecNo )
         ENDIF
      ELSE
         ADO_GOTOP( nWA )
      ENDIF

      aWAData[WA_ISITSUBSET] := .F.

   ELSE
      IF aWAData[ WA_INDEXACTIVE ] > 0
         aOrderInfo[ UR_ORI_RESULT ] := aWAData[ WA_INDEXES ] [aWAData[ WA_INDEXACTIVE ]]
      ELSE
         aOrderInfo[ UR_ORI_RESULT ] := ""
      ENDIF

   ENDIF

   RETURN HB_SUCCESS


// TO BUILD ADO INDEXES FOR FASTER FILTERS
STATIC FUNCTION ADO_OPTIMIZE( cExpression, oRecordSet )
 LOCAL n,  nTotalFields := oRecordSet:Fields:Count

   FOR n := 1 TO nTotalFields
       IF AT( oRecordSet:Fields( n - 1 ):Name, cExpression ) > 0
          IF !oRecordSet:Fields( oRecordSet:Fields( n - 1 ):Name ):Properties():Item("Optimize"):Value
             oRecordSet:Fields( oRecordSet:Fields( n - 1 ):Name ):Properties():Item("Optimize"):Value := 1
          ENDIF

       ENDIF

   NEXT

  RETURN HB_SUCCESS


STATIC FUNCTION ADO_ORDLSTADD( nWA, aOrderInfo )
   LOCAL cTablename := USRRDD_AREADATA( nWA )[ WA_TABLENAME ]
   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL aFiles := ListDbfIndex() //15.09 ListIndex()
   LOCAL aTempFiles := ListTmpNames()
   LOCAL cExpress := "" ,cFor:="",cUnique:="",cDesc := "",y,z,x
   LOCAL aTmpIndx := ListTmpIndex()
   LOCAL aTmpExp := ListTmpExp()
   LOCAL aTmpFor := ListTmpFor()
   LOCAL aTmpUnique := ListTmpUnique()
   LOCAL aTmpDbfExp := ListTmpDbfExp()
   LOCAL aTmpDbfFor := ListTmpDbfFor()
   LOCAL aTmpDbfUnique := ListTmpDbfUnique()
   LOCAL aTmpDesc := ListTmpDbfDesc()
   LOCAL aTmpBook := ListTmpBookM()
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

       y := ASCAN( aTmpIndx, cIndex )
//VERIFICAR SE JA ESTA NA LISTA E NAO ABRE
       AADD( aWAData[WA_INDEXES],cIndex )

       AADD( aWAData[WA_INDEXEXP],UPPER(aTmpExp[y] ))

       AADD( aWAData[WA_INDEXFOR],IF(!EMPTY(aTmpFor[y]),"WHERE ","")+aTmpFor[y])
       AADD( aWAData[WA_INDEXUNIQUE],aTmpUnique[y])
       AADD( aWAData[WA_INDEXDESC],aTmpDesc[y])
       AADD( aWAData[ WA_ABOOKMARKS ], aTmpBook[y])

       IF ASCAN( aWAData[WA_INDEXES],cIndex) = 1 //FIRST INDEX GETS CONTROL
          aWAData[WA_INDEXACTIVE] := 1 //always qst one
          aOrderInfo[UR_ORI_TAG] := 1 //1

          ADO_ORDLSTFOCUS( nWA, aOrderInfo )

       ENDIF

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

              IF LEN(aFiles[y,z+1]) >= 5 //DESCEND CONDITION IS PRESENT?
                 cDesc := aFiles[y,z+1,5]
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

    //15.09 NEW INDEXES
    cExpress := ADOINDEXFIELDS( nWA, cExpress, !EMPTY( cDesc ) )

    //CHECK IF INDEX ALREADY OPEN
    FOR x := 1 TO 50
       IF ORDNAME(x) = cIndex
          RETURN HB_SUCCESS
       ENDIF

    NEXT

    AADD( aWAData[WA_INDEXES],UPPER(cIndex))
    AADD( aWAData[WA_INDEXEXP],UPPER(cExpress))
    AADD( aWAData[WA_INDEXFOR],UPPER(cFor))
    AADD( aWAData[WA_INDEXUNIQUE],UPPER(cUnique))
    AADD( aWAData[WA_INDEXDESC],UPPER(cDesc))
    AADD( aWAData[ WA_ABOOKMARKS ],{})

    IF PROCNAME( 1 ) <> "ADO_INDEXAUTOOPEN"  // IT TAKES CARE OF SET ORDER
       IF LEN( aWAData[WA_INDEXES] ) = 1 //NO PREVIOUS OPENED INDEX YET FIRST OPENED GET CONTROLING ORDER
          aOrderInfo[UR_ORI_TAG] := 1
          aWAData[WA_INDEXACTIVE] := 1
          ADO_ORDLSTFOCUS( nWA, aOrderInfo )
       ENDIF

    ENDIF


   RETURN HB_SUCCESS


STATIC FUNCTION ADO_ORDLSTCLEAR( nWA )

   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]
   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL nRecNo
   LOCAL n

   IF !ADOCON_CHECK()
      RETURN HB_FAILURE

   ENDIF

   aWAData[WA_INDEXES]  := {}
   aWAData[WA_INDEXEXP] := {}
   aWAData[WA_INDEXFOR] := {}
   aWAData[WA_INDEXACTIVE] := 0
   aWAData[WA_INDEXUNIQUE] := {}
   aWAData[WA_SCOPES] := {}
   aWAData[WA_SCOPETOP] := {}
   aWAData[WA_SCOPEBOT] := {}
   aWAData[WA_ISITSUBSET] := .F.
   aWAData[WA_ABOOKMARKS] := {}
   aWAData[WA_INDEXDESC] := {}

   ADO_RECID( nWA, @nRecNo )
   oRecordSet:Sort := ADO_GET_FIELD_RECNO(  aWAData[ WA_TABLENAME ] )
   oRecordSet:Filter := ""

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
   LOCAL aTmpUnique := ListTmpUnique()
   LOCAL aTmpDbfExp := ListTmpDbfExp()
   LOCAL aTmpDbfFor := ListTmpDbfFor()
   LOCAL aTmpDbfUnique := ListTmpDbfUnique()
   LOCAL aTmpDesc := ListTmpDbfDesc()
   LOCAL aTmpBook := ListTmpBookM()
   LOCAL cForExp := ""
   LOCAL cFile := cIndex, aBookMarks := {}

    //MAYBE IT COMES WITH FILE EXTENSION AND PATH
    cIndex := CFILENOPATH(cIndex)
    cIndex := UPPER(CFILENOEXT(cIndex))

    //TMP FILES NOT PRESENT IN ListIndex
    IF ASCAN(aTempFiles,(UPPER(SUBSTR(cIndex,1,3)) )) > 0 .OR. ASCAN(aTempFiles,UPPER(SUBSTR(cIndex,1,4)) ) > 0
       //we need to write the file to allow that some function
       //returning tmp file can see that this file already exists
       MEMOWRIT(cFile,"nada")

    ELSE
       IF ASCAN( aWAData[WA_INDEXES],cIndex) > 0
         // BUILD ERROR
       ENDIF

    ENDIF

    AADD(aTmpIndx,UPPER(cIndex))
    AADD(aTmpExp,ADOINDEXFIELDS(nWA, aOrderCreateInfo[UR_ORCR_CKEY],;
         IF( !EMPTY( acondinfo ), aCondInfo[UR_ORC_DESCEND], .F.) ) )
    AADD(aTmpDbfExp, aOrderCreateInfo[UR_ORCR_CKEY] )

    IF aOrderCreateInfo[UR_ORCR_UNIQUE]
       AADD(aTmpUnique," DISTINCT " )
       AADD(aTmpDbfUnique, " UNIQUE " )

    ELSE
       AADD(aTmpUnique,"")
       AADD(aTmpDbfUnique,"")

    ENDIF

    IF !EMPTY(acondinfo)
       cForExp := IF( !EMPTY( acondinfo[UR_ORC_CFOR]), acondinfo[UR_ORC_CFOR], "" )
       AADD(aTmpFor,cForExp )//CLEAN THE DOT .AND. .OR.
       AADD(aTmpDbfFor,cForExp)
       AADD(aTmpDesc,acondinfo[ UR_ORC_DESCEND ])
       //BUILD THE INDEX
       ADOUDFINDEX( nWa, aOrderCreateInfo[UR_ORCR_BKEY],;
                    acondinfo[UR_ORC_BFOR], ;
                    aOrderCreateInfo[ UR_ORCR_UNIQUE],;
                    acondinfo[ UR_ORC_DESCEND ],acondinfo[UR_ORC_BEVAL],;
                    acondinfo[UR_ORC_STEP], acondinfo[UR_ORC_BWHILE], @aBookMarks,;
                    aOrderCreateInfo[UR_ORCR_CKEY] )

    ELSE
       AADD(aTmpFor,"")
       AADD(aTmpDbfFor,"")
       AADD(aTmpDesc, .F.)
       //BUILD THE INDEX
       ADOUDFINDEX( nWa, aOrderCreateInfo[UR_ORCR_BKEY], NIL,;
                    aOrderCreateInfo[ UR_ORCR_UNIQUE],;
                    .F., NIL, NIL, NIL, @aBookMarks,;
                    aOrderCreateInfo[UR_ORCR_CKEY] )

    ENDIF

    // only with set focus these bookmarks become active
    AADD( aTmpBook, aBookMarks )

   RETURN HB_SUCCESS


//BUILD THE ARRAY OF BOOKMARKS TO HAVE IT ORDERED WITH UDF FUNCTIONS IN INDEX EXPRESSION
STATIC FUNCTION ADOUDFINDEX( nWa, xExpression, xCondition, lUnique, lDesc, bEval,;
                             nEvery, xWhile, aBookMarks, cExpression )
 LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]
 LOCAL aWAData := USRRDD_AREADATA( nWA )
 LOCAL lRetVal := .F.
 LOCAL aUdfs := ListUdfs(), n
 LOCAL bIndexExprn, bFilterCond, nDecimals, xMark, bWhile, nStart := 1
 LOCAL xEvaluated, cSort

   bEval  := IF( bEval == NIL, { || .T. }, bEval )
   nEvery := IF( nEvery == NIL, 0, nEvery )
   cEXpression := IF( cExpression = NIL, cExpression := xExpression, cEXpression )

   IF ! EMPTY( aUdfs )
      FOR n := 1 TO LEN( aUdfs )
          IF AT( aUdfs[ n ], UPPER(cExpression) ) > 0
             lRetval := .T.
             EXIT

          ENDIF

      NEXT

   ENDIF

   IF ! EMPTY( xCondition ) .OR. ! EMPTY( xWhile )
      lRetVal = .T.
   ENDIF

   IF lRetVal //.OR. PROCNAME( 1 ) = "ADO_ORDCREATE"
      aBookMarks :={}

      IF !EMPTY( xWhile )
         nStart := 0

      ENDIF
      xWhile := IF( EMPTY( xWhile ), ".T.", xWhile )
      xCondition := IF( EMPTY( xCondition ), ".T.", xCondition )
      IF VALTYPE( xExpression ) = "B"
         bIndexExprn := xExpression

      ELSE
         bIndexExprn := &("{ || " + xExpression + "}")

      ENDIF
      IF VALTYPE( xCondition ) = "B"
         bFilterCond := xCondition

      ELSE
         bFilterCond := &("{ || " + xCondition + "}")

      ENDIF

      IF VALTYPE( xWhile ) = "B"
         bWhile := xWhile

      ELSE
         bWhile := &("{ || " + xWhile + "}")

      ENDIF

      nDecmals := SET( _SET_DECIMALS, 0 )
      n := 0

      IF nStart > 0 .AND. !ADOEMPTYSET( oRecordSet )
         oRecordSet:MoveFirst()

      ENDIF

      DO WHILE EVAL (bWhile ) .AND. !oRecordSet:Eof()

         IF nEvery > 0
            n++
            IF n = nEvery
               EVAL( bEval )
               n := 0

            ENDIF

         ENDIF
         xEvaluated := EVAL( bIndexExprn )

         IF lUnique
            IF ASCAN( aBookMarks,;
                     {|x|  SUBSTR( x[ 2 ], 1, LEN( xEvaluated ) ) = xEvaluated  } ) = 0
               IF EVAL( bFilterCond )
                  AADD( aBookMarks, { oRecordSet:BookMark, xEvaluated } )

               ENDIF
            ENDIF
         ELSE
            IF EVAL( bFilterCond )
               AADD( aBookMarks, { oRecordSet:BookMark, xEvaluated } )

             ENDIF

         ENDIF

          oRecordSet:MoveNext()

      ENDDO

      SET( _SET_DECIMALS, nDecimals )

      IF lDesc
         ASORT( aBookMarks, NIL, NIL, { |x,y| x[ 2 ] > y[ 2 ] } )

      ELSE
         ASORT( aBookMarks, NIL, NIL, { |x,y| x[ 2 ] < y[ 2 ] } )

      ENDIF

      // only with set focus these bookmarks become active
      lRetVal := .T.
   ELSE
      lRetVal := .F.
   ENDIF

   RETURN lRetVal


//USED TO HAVE THE RIGHT RETURNED VALUES FROM ADO_ORDINFO
STATIC FUNCTION ADOGETORDINFO( nWA, nOption, aInfo )
  RETURN ADO_ORDINFO( nWA, nOption, @aInfo )


STATIC FUNCTION ADO_ORDLSTREBUILD( nWA, aOrderInfo )
    //DOES NOTHING AS INDEXES ARE VIRTUAL THEY REALLY DONT EXIST AS FILES
    //ITS HERE ONLY REDEFINE SUPER FUNCTION AND TO AVOID ERROR.

   RETURN HB_SUCCESS


STATIC FUNCTION ADO_ORDDESTROY( nWA, aOrderInfo )

   LOCAL aWAData := USRRDD_AREADATA( nWA ), n
   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]

   n:= ASCAN(aWAData[ WA_INDEXES ],aOrderInfo[ UR_ORI_TAG ])

   IF n > 0
      ADEL( aWAData[ WA_INDEXES ], n, .T.)
      ADEL( aWAData[ WA_INDEXEXP ], n, .T.)
      ADEL( aWAData[ WA_INDEXFOR ], n, .T.)
      ADEL( aWAData[WA_INDEXUNIQUE], n, .T.)
      ADEL( aWAData[ WA_ABOOKMARKS ], n, .T.)
      ADEL( aWAData[WA_INDEXDESC] , n, .T.)

      IF n = aWAData[ WA_INDEXACTIVE ]
         aWAData[ WA_INDEXACTIVE ] := 0
         //11.08.15 "NATURAL ORDER"
         oRecordSet:Sort := ADO_GET_FIELD_RECNO(  aWAData[ WA_TABLENAME ] )
      ENDIF

   ENDIF

   RETURN HB_SUCCESS


STATIC FUNCTION ADOINDEXFIELDS( nWA, cExpression, lDesc )
  LOCAL n,nAt, aFields :={}, cStr := ""
  LOCAL aUdfs := ListUdfs()

    cExpression := UPPER(cExpression)
    IF ! EMPTY( aUdfs )
       FOR n := 1 TO LEN( aUdfs )
           IF AT( aUdfs[ n ], cExpression ) > 0
              RETURN cStr
              EXIT

           ENDIF

       NEXT

    ENDIF

    FOR n := 1 to (nWA)->(FCOUNT()) // we have to check all fields in table because there isnt any conspicuous mark on the expression to guide us
        nAt := AT(ALLTRIM((nWA)->(FIELDNAME(n))),cExpression)

        IF nAt > 0
           AADD( aFields , {ALLTRIM((nWA)->(FIELDNAME(n))), nAt} ) //nAt order of the field in the expression
        ENDIF

    NEXT

    //we need to have the fields with the same order as in index expression nAt
    aFields := ASORT( aFields ,,, {|x,y| x[2] < y[2] } )

    FOR n := 1 TO LEN(aFields)
        cStr += aFields[n,1]

        IF n < LEN(aFields)
           cStr += IF( lDesc, " DESC,", ",")
        ENDIF

    NEXT


   RETURN cStr


STATIC FUNCTION IndexBuildExp(nWA,nIndex,aWAData,lCountRec,myCfor)  //notgroup for adoreccount

   LOCAL cSql := "", cOrder:="", cUnique:="", cFor:=""
   LOCAL aInfo

     DEFAULT lCountRec TO .F.
     DEFAULT myCfor TO "" //when it comes ex from ado_seek to add to where clause

     IF !lCountRec
        aInfo := Array( UR_ORI_SIZE )
        aInfo[UR_ORI_TAG]:= nIndex
        ADO_ORDINFO( nWA, DBOI_EXPRESSION, @aInfo ) //(nWA)->(ORDKEY(nIndex))
        cOrder := aInfo[UR_ORI_RESULT]

        IF !EMPTY(cOrder)
           cOrder := " ORDER BY "+cOrder+","+ADO_GET_FIELD_RECNO(  aWAData[ WA_TABLENAME ] ) //21.5.15 STRTRAN(cOrder,"+",",")
        ELSE //11.08.15 DEFAULT ORDERED BY RECNO  IN DBFS RDD
           cOrder := " ORDER BY "+ADO_GET_FIELD_RECNO(  aWAData[ WA_TABLENAME ] )
        ENDIF

     ENDIF

     IF  nIndex > 0 .AND. nIndex <= LEN(aWAData[ WA_INDEXUNIQUE ])
         cUnique  := aWAData[ WA_INDEXUNIQUE ][nIndex ]+IF(lCountRec, " COUNT(*) ",aWAData[ WA_TABLENAME ]+".*")

     ELSE
        IF lCountRec
           cUnique := " COUNT(*) "
        ENDIF

     ENDIF

     IF EMPTY(cUnique)
        cUnique := aWAData[ WA_TABLENAME ]+".*"

     ENDIF

     IF  nIndex > 0 .AND. nIndex <= LEN(aWAData[ WA_INDEXFOR ])
         cFor  := " "+aWAData[ WA_INDEXFOR ][ nIndex ]

     ENDIF

     IF !EMPTY(mycFor)
        cFor += IF(!EMPTY(cFor)," AND "," WHERE ")+mycFor

     ENDIF

     cUnique := STRTRAN(  cUnique,  "#",  "") //#temp tables
     cSql := "SELECT "+ cUnique+" FROM " + aWAData[ WA_TABLENAME ]+ IF(!EMPTY(cFor),cFor,"")+ cOrder

   RETURN cSql


STATIC FUNCTION KeyExprConversion( cOrder, cTableName )

 LOCAL y, z , aFiles := ListDbfIndex(), cExpress:= "",cFor:="",cUnique :=""
 LOCAL aTempFiles := ListTmpNames()
 LOCAL aTmpIndx := ListTmpIndex()
 LOCAL aTmpDbfExp := ListTmpDbfExp()
 LOCAL aTmpDbfFor := ListTmpDbfFor()
 LOCAL aTmpDbfUnique := ListTmpDbfUnique()

      //TMP FILES NOT PRESENT IN ListIndex ADDED TO THEIR OWN ARRAY FOR THE DURATION OF THE APP
    IF ASCAN(aTempFiles,UPPER(SUBSTR(cOrder,1,3)) ) > 0 .OR. ASCAN(aTempFiles,UPPER(SUBSTR(cOrder,1,4)) ) > 0
       //it was added to the array by ado_ordcreate we have only to set focus
       y := ASCAN(aTmpIndx,cOrder)

       IF Y > 0
          cExpress := aTmpDbfExp[y]
          cFor := aTmpDbfFor[y]
          cUnique := aTmpDbfUnique[y]
       ENDIF

       RETURN {cExpress,cFor,cUnique}

    ENDIF

    y:=ASCAN( aFiles, { |z| z[1] == cTablename } )

    IF y >0
       FOR z :=1 TO LEN( aFiles[y]) -1
           IF aFiles[y,z+1,1] == cOrder
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

    ENDIF


   RETURN {cExpress,cFor,cUnique}

/*                               END INDEX RELATED FUNCTIONS  */


/*                               LOCKS RELATED FUNCTIONS */
STATIC FUNCTION ADO_RAWLOCK( nWA, nAction, nRecNo )

// LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]

   /* TODO WHAT IS THIS FOR?*/

   HB_SYMBOL_UNUSED( nRecNo )
   HB_SYMBOL_UNUSED( nWA )
   HB_SYMBOL_UNUSED( nAction )

   RETURN HB_SUCCESS


STATIC FUNCTION ADO_LOCK( nWA, aLockInfo )

   LOCAL aWdata := USRRDD_AREADATA( nWA ),lOk := .T.,nRecNo
   LOCAL oRs :=  aWdata[WA_RECORDSET],nPos

   HB_SYMBOL_UNUSED( nWA )

   IF !ADOCON_CHECK()
      RETURN HB_FAILURE

   ENDIF

   //IF WE TRY TO GETVALUE AND RESYNC NOT POSSIBLE RECORD DELETED WE ARE AT EOF
   IF oRs:Eof() .AND. PROCNAME(1) <> "ADO_APPEND"
      aLockInfo[ UR_LI_RESULT ] := .T.
      RETURN HB_FAILURE

   ENDIF

   IF (VALTYPE(aWdata[WA_OPENSHARED]) = "L"  .AND. !aWdata[WA_OPENSHARED]) .OR. aWdata[ WA_FILELOCK ]
      aLockInfo[ UR_LI_RESULT ] := .T.
      RETURN HB_SUCCESS

   ENDIF

   //FILE LOCK WE DONT NEED TO CHECK ANYTHING ELSE
   IF aLockInfo[UR_LI_METHOD] = DBLM_FILE
      IF !aWdata[ WA_FILELOCK ]
         ADO_UNLOCK( nWA )

         IF !ADO_GETLOCK(aWdata[WA_TABLENAME],"ZWTXFL",nWA)
            lOk := .F.
         ENDIF

      ENDIF

   ELSE
      IF EMPTY(aLockInfo[ UR_LI_RECORD ])
         ADO_UNLOCK( nWA )
         ADO_RECID(nWa,@nRecNo)
         aLockInfo[ UR_LI_RECORD ] := nRecNo

      ENDIF

      IF ASCAN(aWdata[ WA_LOCKLIST ],aLockInfo[ UR_LI_RECORD ]) = 0
         IF !ADO_GETLOCK(aWdata[WA_TABLENAME],aLockInfo[ UR_LI_RECORD ] ,nWA)
            lOk := .F.
         ENDIF
      ENDIF


   ENDIF

   IF lOK
      IF PROCNAME( 1 ) <> "ADO_APPEND" //NEW RECNO STILL DONT EXIST IN TABLE RECCUNT+1
         ADO_RESYNC( nWA, oRs )

      ENDIF

   ENDIF

   /*
   UR_LI_METHOD VALUES CONSTANTS
   DBLM_EXCLUSIVE 1 RELEASE ALL AND LOCK CURRENT
   DBLM_MULTIPLE 2 LOCK CURRENT AND ADD TO LOCKLIST
   DBLM_FILE 3 RELEASE ALL LOCKS AND FILE LOCK
   */
   IF lOk
      IF aLockInfo[UR_LI_METHOD] = DBLM_FILE
        aWdata[ WA_FILELOCK ] := .T.

      ELSE
         IF ASCAN(aWdata[ WA_LOCKLIST ],aLockInfo[ UR_LI_RECORD ]) = 0
            AADD(aWdata[ WA_LOCKLIST ],aLockInfo[ UR_LI_RECORD ])
         ENDIF

      ENDIF

      aLockInfo[ UR_LI_RESULT ] := .T.

   ELSE

      aLockInfo[ UR_LI_RESULT ] := .F.

   ENDIF


   RETURN HB_SUCCESS


STATIC FUNCTION ADO_UNLOCK( nWA, xRecID )

   LOCAL aWdata := USRRDD_AREADATA( nWA ),n
   LOCAL oRecordSet := aWdata[ WA_RECORDSET ], nRecNo

   HB_SYMBOL_UNUSED( xRecId )
   HB_SYMBOL_UNUSED( nWA )

   IF !ADOCON_CHECK()
      RETURN HB_FAILURE

   ENDIF

   IF !EMPTY(xRecID)
      ADO_GETUNLOCK(aWdata[ WA_TABLENAME],xRecID,nWA)  //RELEASES ONLY THIS RECORD

   ELSE
      ADO_GETUNLOCK(aWdata[ WA_TABLENAME],"ZWTXFL",nWA)

   ENDIF

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


STATIC FUNCTION ADO_GETLOCK(cTable,xRecID, nWA)
 LOCAL aWdata := USRRDD_AREADATA( nWA )
 LOCAL lRetval := .F.
 LOCAL aLockCtrl
 LOCAL currArea := SELECT(),n

  IF !aWData[WA_LOCKSCHEME ]//22.08.15 WORKING WITHOUT LOCKS BUT KEEPING LOCK ARRAY
     RETURN .T.
  ENDIF

  aLockCtrl :=  ADOLOCKCONTROL()

  xRecID := ALLTRIM(CVALTOCHAR(xRecID))

  IF SELECT("TLOCKS") = 0
     DBUSEAREA(.T.,aLockCtrl[2],aLockCtrl[1]+RDDINFO(RDDI_TABLEEXT,,aLockCtrl[2]),"TLOCKS",.T.)
     TLOCKS->(DBSETINDEX(aLockCtrl[1]+RDDINFO(RDDI_ORDBAGEXT,,aLockCtrl[2]) ))
     SELECT(currArea)

  ENDIF

  IF TLOCKS->(DBSEEK( cTable + xRecID ))
     IF TLOCKS->(DBRLOCK(  TLOCKS->(RECNO()) ))
        lRetval := .T.
     ENDIF

  ELSE
     IF !TLOCKS->(DBSEEK(SPACE(50))) //RECOVERING USED RECORDS
        TLOCKS->(DBAPPEND( .F.) ) //DOES NOT RELEASE ANY LOCKS

        IF !NETERR()
           lRetval := .T.
           REPLACE TLOCKS->CODLOCK WITH cTable+xRecID
        ENDIF

     ELSE
        IF TLOCKS->(DBRLOCK( TLOCKS->(RECNO()) ))
           lRetval := .T.
           REPLACE TLOCKS->CODLOCK WITH cTable+xRecID
        ENDIF

     ENDIF

  ENDIF

  IF lRetVal
     TLOCKS->(DBCOMMIT())
     AADD( aWData[WA_TLOCKS], { xRecID, TLOCKS->(RECNO()) } )

  ENDIF


  RETURN lRetval


STATIC FUNCTION ADO_GETUNLOCK(cTable,xRecID,nWA)
 LOCAL aWdata := USRRDD_AREADATA( nWA )
 LOCAL aLockCtrl
 LOCAL currArea := SELECT(),n,aDels := {}

  IF !aWData[WA_LOCKSCHEME ]//22.08.15 WORKING WITHOUT LOCKS BUT KEEPING LOCK ARRAY
     RETURN HB_SUCCESS
  ENDIF

  aLockCtrl :=  ADOLOCKCONTROL()

  xRecID := ALLTRIM(CVALTOCHAR(xRecID))

  IF SELECT("TLOCKS") = 0
     DBUSEAREA(.T.,aLockCtrl[2],aLockCtrl[1]+RDDINFO(RDDI_TABLEEXT,,aLockCtrl[2]),"TLOCKS",.T.)
     TLOCKS->(DBSETINDEX(aLockCtrl[1]+RDDINFO(RDDI_ORDBAGEXT,,aLockCtrl[2]) ))
     SELECT(currArea)

  ENDIF

  IF ! xRecID = "ZWTXFL"
     IF LEN( aWData[WA_TLOCKS] )  > 0
        n := ASCAN(aWData[WA_TLOCKS], {|a| a[1,1] = xRecID} )
     ELSE
         n := 0
     ENDIF

     IF n > 0
        TLOCKS->(DBGOTO(aWData[WA_TLOCKS][n,2]))

        IF TLOCKS->(ISLOCKED())
           REPLACE TLOCKS->CODLOCK WITH SPACE(50) //NOT TO GROW IT WILL BE RECYCLED NEXT TIME
        ENDIF

        TLOCKS->(DBRUNLOCK(aWData[WA_TLOCKS][n,2]))
        ADEL(aWData[WA_TLOCKS],n,.T.)
        TLOCKS->(DBCOMMIT())

     ENDIF

  ELSE

     //RELEASE ALL LOCKS
     FOR n := 1 TO LEN(aWData[WA_TLOCKS])
         TLOCKS->(DBGOTO( aWData[WA_TLOCKS][n,2] ))

         IF TLOCKS->(ISLOCKED())
            REPLACE TLOCKS->CODLOCK WITH SPACE(50) //NOT TO GROW IT WILL BE RECYCLED NEXT TIME
         ENDIF

         TLOCKS->(DBRUNLOCK( aWData[WA_TLOCKS][n,2] ))
         AADD (aDels ,n )
         TLOCKS->(DBCOMMIT())

      NEXT

  ENDIF

  IF LEN( aDels) > 0
     FOR N :=  1 TO LEN(aDels)
        ADEL( aWData[WA_TLOCKS],aDels[n], .T. )
     NEXT

  ENDIF


  RETURN HB_SUCCESS


FUNCTION ADO_ISLOCKED(cTable,xRecID,aWAData)

   IF !ADOCON_CHECK()
      RETURN HB_FAILURE

   ENDIF

   IF !aWAData[WA_LOCKSCHEME ]//22.08.15 WORKING WITHOUT LOCKS BUT KEEPING LOCK ARRAY
      RETURN .T.

   ENDIF

  RETURN IF( !aWAData[WA_OPENSHARED] .OR. aWAdata[ WA_FILELOCK ] ,;
            .T., ASCAN( aWAData[WA_LOCKLIST],  xRecID )  > 0  )


STATIC FUNCTION ADO_OPENSHARED( nWA,cTable, lExclusive, lClose )
 LOCAL aWAData := USRRDD_AREADATA( nWA )
 LOCAL lRetval := .F.
 LOCAL aLockCtrl
 LOCAL cFile

 DEFAULT lClose TO .F.

  IF !aWAData[WA_LOCKSCHEME ]
     RETURN .T.

  ENDIF

  aLockCtrl :=  ADOLOCKCONTROL()
  cFile := STRTRAN(aLockCtrl[1],"TLOCKS","")+cTable+".ctrl"

  /* //TEMPORARY TABLES ARE PRIVATE TO THE USER DONT NEED LOCK CONTROL
  IF UPPER( SUBSTR( aWAData[WA_TABLENAME] ,1,3 ) ) = "TMP" .OR. UPPER( SUBSTR( aWAData[WA_TABLENAME] ,1,4 ) ) = "TEMP" .OR.;
      UPPER( SUBSTR( aWAData[WA_TABLENAME] ,1,4 ) ) = "#TMP" .OR. UPPER( SUBSTR( aWAData[WA_TABLENAME] ,1,5 ) ) = "#TEMP"
     RETURN .T.

  ENDIF */

  IF !lClose
     IF !FILE( cFile)
        MEMOWRIT(cFile,"Nada")
     ENDIF

     aWAData[WA_FILEHANDLE] := FOPEN( (cFile), IF(lExclusive, FO_EXCLUSIVE, FO_DENYNONE  ) )
     IF FERROR() = 0
        lRetval := .T.
     ENDIF

  ELSE
     IF FCLOSE( aWAData[WA_FILEHANDLE] )
        aWAData[WA_FILEHANDLE] := NIL
     ENDIF

  ENDIF


  RETURN lRetval

/*                              END LOCKS RELATED FUNCTIONS */

/*                                             TRANSACTION RELATED FUNCTIONS */
STATIC FUNCTION ADO_FLUSH( nWA )

   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]
   LOCAL n

   IF !ADOCON_CHECK()
      RETURN HB_FAILURE
   ENDIF

   TRY
      IF VALTYPE(oRecordSet) == "O"
         IF !oRecordSet:Eof .and. !oRecordSet:Bof
            oRecordSet:Update()
         ENDIF

      ENDIF

   CATCH
      ADOSHOWERROR( USRRDD_AREADATA( nWA )[ WA_CONNECTION ] )

   END


   RETURN HB_SUCCESS


FUNCTION ADOBEGINTRANS(nWa)

 LOCAL oCon := hb_adoRddGetConnection( nWA )

 IF !ADOCON_CHECK()
     RETURN HB_FAILURE

 ENDIF

  TRY
     oCon:BeginTrans()

  CATCH
     ADOSHOWERROR(oCon)

  END

  RETURN .T.


FUNCTION ADOCOMMITTRANS(nWa)

 LOCAL oCon := hb_adoRddGetConnection( nWA )

  IF !ADOCON_CHECK()
     RETURN HB_FAILURE

  ENDIF

  TRY
     oCon:CommitTrans()

  CATCH
     ADOSHOWERROR(oCon)

  END

  RETURN .T.


FUNCTION ADOROLLBACKTRANS(nWa)

 LOCAL oCon := hb_adoRddGetConnection( nWA )
 LOCAL n,oRs, aWAData

   IF !ADOCON_CHECK()
      RETURN HB_FAILURE

   ENDIF

  TRY
     oCon:RollBackTrans()

     //UPDATE ALL RECORDSETS TO CLEAN THE CANCELED TRANSACTION FROM RECORDSETS
     FOR n := 1 TO 255
         oRs := hb_adoRddGetRecordSet(n)
         IF VALTYPE(oRs) = "O"
            ADO_REFRESH( nWA, .T. )
            aWAData := USRRDD_AREADATA( n )
         ENDIF

     NEXT

  CATCH
     ADOSHOWERROR(oCon)

  END

  RETURN .T.
/*                                      END TRANSACTION RELATED FUNCTIONS */


/*                                 SCOPE LOCATE SEEK FILTER RELATED FUNCTIONS */

STATIC FUNCTION ADO_SETFILTER( nWA, aFilterInfo )

   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]
   LOCAL aWAData := USRRDD_AREADATA( nWA )

   aWAData[WA_FILTERACTIVE]  := aFilterInfo[ UR_FRI_BEXPR ] //SAVE ACTIVE FILTER EXPRESION
   aWAData[WA_CFILTERACTIVE] := aFilterInfo[ UR_FRI_CEXPR ]

   RETURN HB_SUCCESS


STATIC FUNCTION ADO_CLEARFILTER( nWA )

 LOCAL aWAData := USRRDD_AREADATA( nWA )
 LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]


   aWAData[WA_FILTERACTIVE] := NIL //NO FILTER
   aWAData[WA_CFILTERACTIVE] := "" //NO FILTER

   RETURN HB_SUCCESS


STATIC FUNCTION ADO_FILTERTEXT(nWa,cFilterExp)
 LOCAL aWAData := USRRDD_AREADATA( nWA )

   cFilterExp := aWAData[WA_CFILTERACTIVE]

   RETURN HB_SUCCESS


STATIC FUNCTION ADO_SETLOCATE( nWA, aScopeInfo )

   LOCAL aWAData := USRRDD_AREADATA( nWA )

   aWAData[ WA_SCOPEINFO ] := aScopeInfo

   RETURN HB_SUCCESS


STATIC FUNCTION ADO_LOCATE( nWA, lContinue )

   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := aWAData[ WA_RECORDSET ]

   IF ADOEMPTYSET(oRecordSet)
      RETURN HB_FAILURE

   ENDIF

   IF !lContinue
      ADO_GOTOP(nWA) //START FROM BEGINING

   ELSE
      ADO_SKIPRAW( nWa, 1 ) //WE DONT WNAT TO FIND THIS ONE AGAIN

   ENDIF

   CURSORWAIT()

   DO WHILE !aWAData[ WA_EOF ]
      IF EVAL( aWAData[ WA_SCOPEINFO ][ UR_SI_BFOR ])
         aWAData[ WA_FOUND ] := .T.
         EXIT
      ENDIF

      ADO_SKIPRAW( nWa, 1 )
      SYSREFRESH()

   ENDDO

   CURSORARROW()

   aWAData[ WA_FOUND ] := ! oRecordSet:EOF
   aWAData[ WA_EOF ] := oRecordSet:EOF


   RETURN HB_SUCCESS


STATIC FUNCTION ADO_SEEK( nWA, lSoftSeek, cKey, lFindLast )
 LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]
 LOCAL aWAData := USRRDD_AREADATA( nWA )
 LOCAL aSeek,cSql, oRs, cTop:="", nLen, cField, oError,nPos, oRs2

   HB_SYMBOL_UNUSED( nWA )
   HB_SYMBOL_UNUSED( lSoftSeek )
   HB_SYMBOL_UNUSED( cKey )
   HB_SYMBOL_UNUSED( lFindLast )

   DEFAULT lFindLast TO .F.
   DEFAULT lSoftSeek TO .F.

   IF ADOEMPTYSET(oRecordSet)
      aWAData[ WA_FOUND ] := .F.
      aWAData[ WA_EOF ] := oRecordSet:EOF
      RETURN HB_SUCCESS

   ENDIF

   IF aWAData[WA_INDEXACTIVE] = 0
      MSGALERT("No Index active seek not allowed!") //SHOULD RAISE ERROR
      oError := ErrorNew()
      oError:GenCode := 1201
      oError:SubCode := 1201
      oError:Description :=  "Work area not indexed"
      oError:FileName := ALIAS(nWA)
      oError:OsCode := 0 // TODO
      oError:CanDefault := .F.
      UR_SUPER_ERROR( nWA, oError )
      RETURN HB_FAILURE

   ENDIF

   aSeek := ADOPseudoSeek(nWA,cKey,aWAData,lSoftSeek)

   oRs := oRecordSet:Clone
   oRs:Sort := oRecordSet:Sort

   IF aseek[3] .AND. EMPTY( oRecordSet:Filter )//ONLY ONE FIELD AND NO :FILTER
      IF lFindLast
         oRs:MoveLast()

      ELSE
         oRs:MoveFirst()

      ENDIF

      IF AT( "*", aseek[1] ) = 0  //if len ckey less than fieldlen assume partial search
         nLen := RAT( "'", aseek[1]) - AT( "'",aseek[1] )-1
         cField := ALLTRIM( SUBSTR( aseek[1], 1, AT( "=", aseek[1] )-1 ) )

         IF nLen > 0 .AND. nLen < FIELDLEN( FIELDPOS( cField ) )
            aseek[1] := STRTRAN( aseek[1], "=", " LIKE " )
            aseek[1] := STRTRAN( aseek[1], "'", "*'" , 2)
         ENDIF

      ENDIF


      IF lSoftSeek
         IF !lFindLast
            aseek[1] := STRTRAN( aseek[1], "=", ">=" )
         ELSE
            aseek[1] := STRTRAN( aseek[1], "=", "<=" )
         ENDIF

      ENDIF

      oRs:Find( aseek[1],0,IF( lFindLast, adSearchBackward, adSearchForward ) )

   ELSE //WITH :FILTER OR MOE THAN ONE FIELD
      IF !EMPTY( oRecordSet:Filter )
         IF lSoftSeek
            nPos := ASCAN( aWAData[ WA_ABOOKMARKS ][aWAData[WA_INDEXACTIVE]],;
                     {|x|  SUBSTR( x[ 2 ], 1, LEN( cKey ) ) = cKey  } )

            IF oRs:Eof() //!lFindLast
               nPos := ASCAN( aWAData[ WA_ABOOKMARKS ][aWAData[WA_INDEXACTIVE]],;
                     {|x|  SUBSTR( x[ 2 ], 1, LEN( cKey ) ) > cKey  } )

            ELSE
               IF lFindLast
                  nPos := RASCAN( aWAData[ WA_ABOOKMARKS ][aWAData[WA_INDEXACTIVE]],;
                         {|x|  SUBSTR( x[ 2 ], 1, LEN( cKey ) ) = cKey  } )

               ENDIF

            ENDIF

         ELSE
            IF !lFindLast
               nPos := ASCAN( aWAData[ WA_ABOOKMARKS ][aWAData[WA_INDEXACTIVE]],;
                      {|x| SUBSTR( x[ 2 ], 1, LEN( cKey ) ) = cKey  } )
            ELSE
               nPos := RASCAN( aWAData[ WA_ABOOKMARKS ][aWAData[WA_INDEXACTIVE]],;
                     {|x|  SUBSTR( x[ 2 ], 1, LEN( cKey ) ) = cKey  } )

            ENDIF

         ENDIF

         IF nPos > 0
            oRS:BookMark :=  aWAData[ WA_ABOOKMARKS ][aWAData[WA_INDEXACTIVE]] [ npos ][ 1 ]

         ELSE
            oRs:MoveLast()
            oRs:MoveNext()

         ENDIF

      ELSE  //WITHOUT ANY :FILTER
         IF oRecordSet:RecordCount > 2000
            ADO_OPTIMIZE( aseek[1], oRecordSet ) //SEEMS FASTER WITH THIS
         ENDIF

         IF lSoftSeek  //IS THIS RIGHT?
            oRs:Filter :=  aseek[1]
            IF oRs:Eof()
               aSeek[1] := STUFF( aseek[1], RAT( "=", aseek[1] ),1 , ">",  )
               oRs:Filter :=  aseek[1]

            ELSE
               IF lFindLast
                  oRs:MoveLast()

               ENDIF

            ENDIF

         ELSE
            oRs:Filter :=  aseek[1]

            IF ! oRs:Eof()
               IF lFindLast
                  oRs:MoveLast()
               ENDIF
            ENDIF

         ENDIF

      ENDIF

   ENDIF

   IF !oRs:eof() .AND. !oRs:bof()
      oRecordSet:Bookmark := oRs:Bookmark

   ELSE
      oRecordSet:MoveLast()
      oRecordSet:MoveNext()  //eof like the clone

   ENDIF

   aWAData[ WA_FOUND ] :=  ! oRecordSet:EOF
   aWAData[ WA_EOF ] := oRecordSet:EOF
   aWAData[ WA_BOF ] := oRecordSet:BOF

   //IS IT REALLY NECESSARY?
   IF !EMPTY(aWAData[WA_PENDINGREL]) .AND. PROCNAME(2) <> "ADO_RELEVAL" //ENFORCE REL CHILDS BUT NOT IN A ENDLESS LOOP!
      ADO_FORCEREL( nWA )
   ENDIF

   RETURN HB_SUCCESS


STATIC FUNCTION ADO_FOUND( nWA, lFound )

   LOCAL aWAData := USRRDD_AREADATA( nWA )
   lFound := aWAData[ WA_FOUND ]

   RETURN HB_SUCCESS


//build selects expression for find scopes and seeks
STATIC FUNCTION ADOPSEUDOSEEK(nWA,cKey,aWAData,lSoftSeek,lBetween,cKeybottom)

 LOCAL nOrder := aWAData[WA_INDEXACTIVE]
 LOCAL cExpression := aWAData[WA_INDEXEXP][nOrder]
 LOCAL aLens := {}, n, aFields := {} , nAt := 1,cType, lNotFind := .F. ,cSqlExpression := "",nLen
 LOCAL cFields := "",cVal := 0

 DEFAULT lSoftSeek TO .F.//to use like insead of =
 DEFAULT lBetween TO .F.

 cKey := CVALTOCHAR(cKey)
 cKeybottom := CVALTOCHAR(cKeybottom)

    FOR n := 1 to (nWA)->(FCOUNT()) // we have to check all fields in table because there isnt any conspicuous mark on the expression to guide us
       nAt := AT(ALLTRIM((nWA)->(FIELDNAME(n))),cExpression)

       IF nAt > 0
          AADD(aFields ,{ALLTRIM((nWA)->(FIELDNAME(n))),nAt}) //nAt order of the field in the expression
       ENDIF

    NEXT

    //we need to have the fields with the same order as in index expression nAt
    aFields := ASORT( aFields ,,, {|x,y| x[2] < y[2] } )

    cExpression := ""  //USE FOR :FIND NOT NEEDED FOR NOW! GAVE UP
    cSqlExpression := "" //USE FOR SELECT FROM

    FOR nAt := 1 TO LEN(aFields)
        nLen := FIELDSIZE(FIELDPOS(aFields[nAt,1]))
        cType := FIELDTYPE(FIELDPOS(aFields[nAt,1]))

        //extract from cKey the lengh og this field
        IF cType = "C" .OR. cType = "M"
           IF !lBetween
              IF !lSoftSeek
                 cExpression += aFields[nAt,1]+ "="+"'"+SUBSTR( cKey, 1, nLen)+"'"
                 cSqlExpression := cExpression
              ELSE
                 cExpression += aFields[nAt,1]+" = "+"'"+SUBSTR( cKey, 1, nLen)+"'"
                 //cSqlExpression := cExpression
                 cSqlExpression += SUBSTR( cKey, 1, nLen)

                 IF nAt > 1
                    cFields += "+"+aFields[nAt,1]
                 ELSE
                    cFields += aFields[nAt,1]
                 ENDIF
              ENDIF

           ELSE
              cExpression += aFields[nAt,1]+" BETWEEN "+"'"+SUBSTR( cKey, 1, nLen)+"'"+;
                            " AND "+"'"+SUBSTR( cKeyBottom, 1, nLen)+"'"
              cSqlExpression := cExpression
           ENDIF

        ELSEIF cType = "D" .OR. cType = "N" .OR. cType = "I" .OR. cType = "B"
           IF cType = "D"
              IF !lBetween
                 cExpression    += aFields[nAt,1]+ "="+"#"+ADODTOS(SUBSTR( cKey, 1, nLen))+"#" //delim might be #
                 cSqlExpression += aFields[nAt,1]+ "='"+ADODTOS(SUBSTR( cKey, 1, nLen))+"'"
              ELSE
                 cExpression += aFields[nAt,1]+" BETWEEN "+"'"+ADODTOS(SUBSTR( cKey, 1, nLen))+"'"+;
                           " AND "+"'"+ADODTOS(SUBSTR( cKeyBottom, 1, nLen))+"'"
                 cSqlExpression := cExpression
              ENDIF

           ELSE
              cVal := ALLTRIM(STR(VAL(SUBSTR( cKey, 1, nLen))))

              IF !lBetween
                 cExpression    += aFields[nAt,1]+ "="+cVal
                 cSqlExpression += aFields[nAt,1]+ "="+cVal
              ELSE
                 cExpression += aFields[nAt,1]+" BETWEEN "+cVal+;
                          " AND "+cVal
                 cSqlExpression := STRTRAN(cExpression,"#","")
              ENDIF

           ENDIF

        ELSEIF  cType = "L"
           nLen := 3 // although is one in the table in the stirng is 3 ex .t. or .f.

           IF SUBSTR( UPPER(cKey), 1, nLen) = ".T."
              cExpression += aFields[nAt,1]+" <> 0"
           ELSE
              cExpression += aFields[nAt,1]+" = 0"//" NOT "+aFields[nAt,1]
           ENDIF

           cExpression := STRTRAN( UPPER(cExpression), ".T.","True",1,1)
           cExpression := STRTRAN( UPPER(cExpression), ".F.","False",1,1)
           cSqlExpression := cExpression

        ELSE
          lNotFind := .T.  //expression cannot be used by :Find()

        ENDIF

        cKey := SUBSTR(cKey,nLen+1) // take out the len of the expression already used

        IF LBetween
           cKeybottom := SUBSTR(cKeybottom,nLen+1) // take out the len of the expression already used

        ENDIF

        IF nAt < LEN(aFields) //add AND last one isnt needed!
           cExpression += IF(LEN(cKey) > 0 ," AND " , "")

           IF !lSoftSeek  // EXPRESSION FILED-FIELD > "  LEN 2 FIELDS    " DONT NEED AND
              cSqlExpression += IF(LEN(cKey) > 0 ," AND " , "")
           ENDIF

        ENDIF

        IF LEN(cKey) = 0 //there isnt more expression to look for
           EXIT

        ENDIF

    NEXT

    IF lSoftSeek
       cSqlExpression := cFields+" = "+"'"+cSqlExpression+"'"

    ENDIF

  RETURN { cExpression,cSqlExpression,IF( lNotFind, .F., nAt = 1 ) }
/*                                 END LOCATE SEEK FILTER RELATED FUNCTIONS */


/*                                  RELATIONS RELATED FUNCTIONS */
STATIC FUNCTION ADO_SETREL( nWA, aRelInfo )

   LOCAL aWAData := USRRDD_AREADATA( nWA ),n

   FOR n := 1 TO LEN(aRelInfo)
       IF VALTYPE(aRelInfo[n]) = "C"
          IF AT("->",aRelInfo[n]) > 0
             aRelInfo[n] := STRTRAN(aRelInfo[n],"field->","")
          ENDIF

       ENDIF

   NEXT

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

   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL aInfo, nReturn, nOrder, uResult, oError,uResultChild

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
               //doesnt matter nreturn can be eof or bof story continunes
               //ITS MUCH MUCH FASTER!
               ADO_EVALBLOCK( aRelInfo[ UR_RI_CHILD ], aRelInfo[ UR_RI_BEXPR ], @uResultChild )
               IF uResult <> uResultChild
                  ADO_SEEK( aRelInfo[ UR_RI_CHILD ], .F., uResult, .F. )
               ENDIF
            ELSE
               oError := ErrorNew()
               oError:GenCode := 1201
               oError:SubCode := 1201
               oError:Description :=  "Work area not indexed"
               oError:FileName := ALIAS(aRelInfo[ UR_RI_CHILD ])
               oError:OsCode := 0 // TODO
               oError:CanDefault := .F.
               UR_SUPER_ERROR( nWA, oError )
               RETURN HB_FAILURE

            ENDIF

         ELSE
            oError := ErrorNew()
            oError:GenCode := 1201
            oError:SubCode := 1201
            oError:Description :=  "Work area not indexed"
            oError:FileName := ALIAS(aRelInfo[ UR_RI_CHILD ])
            oError:OsCode := 0 // TODO
            oError:CanDefault := .F.
            UR_SUPER_ERROR( nWA, oError )
            RETURN HB_FAILURE

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

   IF PROCNAME(1) <> "ADO_RELEVAL"
      // DONT KNOW WHY BUT DBEVAL ONLY WORK LIKE THIS
      //uResult := Eval( bBlock )
      UR_SUPER_EVALBLOCK( nArea, bBlock, @uResult )

   ELSE
      uResult := Eval( bBlock )

   ENDIF

   IF nCurrArea > 0
      dbSelectArea( nCurrArea )

   ENDIF


   RETURN HB_SUCCESS


STATIC FUNCTION ADO_CLEARREL( nWA )

   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL n,cAlias
   LOCAL aOrderInfo := ARRAY(UR_ORI_SIZE),nRelArea


    aWAData[ WA_PENDINGREL ] := NIL
    aWAData[ WA_LASTRELKEY ] := NIL

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
          cExpr := ""
       ENDIF

   ELSE
      cExpr := ""

   ENDIF

   RETURN HB_SUCCESS
/*                               END RELATIONS RELATED FUNCTIONS */

/*                               FILE RELATED FUNCTION */

STATIC FUNCTION ADO_CREATE( nWA, aOpenInfo  )

  LOCAL aWAData := USRRDD_AREADATA( nWA )
  LOCAL cTable  := hb_tokenGet( aOpenInfo[ UR_OI_NAME ], 1, ";" )
  LOCAL cDataBase  := hb_tokenGet( aOpenInfo[ UR_OI_NAME ], 2, ";" )
  LOCAL cDbEngine  := hb_tokenGet( aOpenInfo[ UR_OI_NAME ], 3, ";" )
  LOCAL cServer    := hb_tokenGet( aOpenInfo[ UR_OI_NAME ], 4, ";" )
  LOCAL cUserName  := hb_tokenGet( aOpenInfo[ UR_OI_NAME ], 5, ";" )
  LOCAL cPassword  := hb_tokenGet( aOpenInfo[ UR_OI_NAME ], 6, ";" )
  LOCAL cSql, cSql2, lAddAutoInc := .F.
  LOCAL oCatalog , cMarkTmp, lNoError := .T.,cTmpTable, n

   IF !ADOCON_CHECK()
      RETURN HB_FAILURE

   ENDIF

   IF EMPTY(cDbEngine) //IF NOT DEFINED USE DEFAULT
      ADODEFAULTS()

   ENDIF

   IF( ALLTRIM( cDataBase ) == "" ,cDataBase:= t_cDataSource, cDataBase )
   IF( ALLTRIM( cTable ) == "" , cTable := aOpenInfo[ UR_OI_NAME ] ,cTable)
   IF( ALLTRIM( cDbEngine ) == "" ,cDbEngine:= t_cEngine, cDbEngine )
   IF( ALLTRIM( cServer ) == "" , cServer:= t_cServer, cServer )
   IF( ALLTRIM( cUserName ) == "" , cUserName:= t_cUserName, cUserName )
   IF( ALLTRIM( cPassword ) == "" , cPassword:= t_cPassword, cPassword )

    hb_adoSetDSource(cDataBase)
    hb_adoSetEngine( cDbEngine )
    hb_adoSetServer( cServer )
    hb_adoSetUser( cUserName )
    hb_adoSetPassword( cPassword )

   IF cDbEngine = "ACCESS" //t_cEngine WITH DEFAULT VALUE BU ADODEFAULTS
      IF !FILE(cDataBase)
         oCatalog    := ADOCLASSNEW( "ADOX.Catalog" ) //TOleAuto():New( "ADOX.Catalog" )
         oCatalog:Create( "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + cDataBase )
      ENDIF

   ENDIF

   aOpenInfo[ UR_OI_NAME ] := CFILENOEXT( CFILENOPATH( cTable ) )

   ADOCONNECT(nWA,aOpenInfo)

   /*
   add HBDELETED BY MAROMANO
   */
   n := ASCAN( aWAData[ WA_SQLSTRUCT ],{ |x| x[1] = ADO_GET_FIELD_DELETED(  aWAData[ WA_TABLENAME ] ) }  )
   IF n == 0
      AADD( aWAData[ WA_SQLSTRUCT ], {  ADO_GET_FIELD_DELETED(  aWAData[ WA_TABLENAME ] ), 'L', 1, 0 } )

   ELSE  //FIX AHF CAN ALREADY EXIST AND NOT LOGICAL FIELD
      aWAData[ WA_SQLSTRUCT ][n,2] := "L"
      aWAData[ WA_SQLSTRUCT ][n,3] :=  1
      aWAData[ WA_SQLSTRUCT ][n,4] := 0

   ENDIF

   /*
   fix to add HBRECNO if it�s not present  // Lucas De Beltran 23.05.2015
   cannot be first otherwise copy to changes all fields order and values ahf 23.5.2015
   */
   n := ASCAN( aWAData[ WA_SQLSTRUCT ],{ |x| x[1] = ADO_GET_FIELD_RECNO(  aWAData[ WA_TABLENAME ] ) }  )
   IF n == 0
      AADD( aWAData[ WA_SQLSTRUCT ], {  ADO_GET_FIELD_RECNO(  aWAData[ WA_TABLENAME ] ), '+', 11, 0 } )

   ELSE  //FIX AHF CAN ALREADY EXIST AND NOT TRUE INC FIELD
      aWAData[ WA_SQLSTRUCT ][n,2] := "+"
      aWAData[ WA_SQLSTRUCT ][n,3] :=  11
      aWAData[ WA_SQLSTRUCT ][n,4] := 0

   ENDIF

   cSql := ADOSTRUCTTOSQL( aWAData,aWAData[ WA_SQLSTRUCT ],@lAddAutoInc )

   IF Lower( Right( aOpenInfo[ UR_OI_NAME ], 4 ) ) == ".dbf" .OR. aWAData[ WA_ENGINE ] = "ADS"
      aWAData[ WA_TABLENAME ] := CFILENOEXT(CFILENOPATH(aWAData[ WA_TABLENAME ] ))

   ENDIF

   cTmpTable := CFILENOEXT( CFILENOPATH( cTable ) )

   TRY
       //TEMPORARY TABLES ARE DESTROYE AUTO WHEN NO NEEDED ANYMORE
       IF UPPER( SUBSTR( cTmpTable ,1,3 ) ) = "TMP" .OR. UPPER( SUBSTR( cTmpTable ,1,4 ) ) = "TEMP"
          cMarkTmp := ADOTEMPTABLE( cDbEngine, aWAData[ WA_CONNECTION ], aWAData[ WA_TABLENAME ] )
          //we need to create the file in order to eventually test file() works in the app
          MEMOWRIT(cTable,"nada")

       ELSE
          cMarkTmp := "TABLE "

       ENDIF

       cSql := "CREATE "+cMarkTmp+  aWAData[ WA_TABLENAME ]  + " ( " + cSql + " )"

       //ORACLE NO AUTOINC ONLY SEQUENCE AS IT SHOULD BE FOR ALL
       IF aWAData[ WA_ENGINE ] = "ORACLE" .AND. lAddAutoInc
          cSql2    := "CREATE SEQUENCE " + aWAData[ WA_TABLENAME ] + "_ID_SEQ"
          cSql     += "|" + cSql2
          cSql2    := "CREATE OR REPLACE TRIGGER " + aWAData[ WA_TABLENAME ] +;
                      "_BI_TRG BEFORE INSERT ON " + aWAData[ WA_TABLENAME ]
          cSql2    += " FOR EACH ROW BEGIN "
          cSql2    += " SELECT " + aWAData[ WA_TABLENAME ] + "_ID_SEQ.nextval INTO :new.id FROM DUAL; END;"
          cSql     += "|" + cSql2

       ENDIF

       IF '|' $ cSql
          AEVAL( HB_ATokens( cSql, '|' ), { |c| aWAData[ WA_CONNECTION ]:Execute( c ) } )

       ELSE
          aWAData[ WA_CONNECTION ]:Execute( cSql )

       ENDIF

       MEMOWRIT( "creatsql.txt", cSql )

   CATCH //modified by LucasDeBeltran
         //INFORM ADOERROR
         ADOSHOWERROR( aWAData[ WA_CONNECTION ] )
         lNoError := .F.

   END

   IF lNoError .AND. ( PROCNAME(1) == "__DBCOPY" .OR. PROCNAME(2) == "__DBCREATE")// WE NEED TO LEAVE OPENED TO ADO_TRANS COMPLETE THE JOB
      IF EMPTY( aOpenInfo[ UR_OI_AREA ] )
         SELE 0

      ENDIF
      IF EMPTY( aOpenInfo[ UR_OI_ALIAS ] )
         aOpenInfo[ UR_OI_ALIAS ] := aOpenInfo[ UR_OI_NAME ]+"CP"

      ENDIF

      USE (aOpenInfo[ UR_OI_NAME ]) EXCLUSIVE ALIAS (aOpenInfo[ UR_OI_ALIAS ])

      IF NETERR()
         //ERROR TO BE RAISED!
      ENDIF

   ENDIF


   RETURN HB_SUCCESS


STATIC FUNCTION ADO_CREATEFIELDS( nWA, aStruct )

 LOCAL aWAData := USRRDD_AREADATA( nWA )//modified by LucasDeBeltran

   aWAData[ WA_SQLSTRUCT ] := aStruct //modified by LucasDeBeltran

   RETURN HB_SUCCESS


STATIC FUNCTION ADOSTRUCTTOSQL( aWAData,aStruct ,lAddAutoInc)

 LOCAL cSql := "", nCol
 LOCAL oCn := aWAData[ WA_CONNECTION ]
 LOCAL dbEngine := aWAData[ WA_ENGINE ], nver :=0
 LOCAL aEngines := { "DBASE","ACCESS","MSSQL","MYSQL","ORACLE","SQLITE",;
                     "FOXPRO","POSTGRE","INFORMIX","ANYWHERE","ADS"}
 LOCAL snDbms := ASCAN(aEngines,dbEngine),c

   //TAKEN FROM ADOFUNCS THANKS ANTONIO AND MR RAO!
   IF snDbms <  2
      MsgAlert( aEngines[snDbms] + " not supported by the function" )
      RETURN HB_FAILURE

   ENDIF

   IF dbEngine == "MSSQL"
      nVer     := VAL( oCn:Properties( "DBMS Version" ):Value )

   ENDIF

   FOR nCol := 1 TO LEN( aStruct )
       cSql  += ' ' + ADOQUOTEDCOLSQL( Trim( aStruct[nCol, 1 ] ), dbEngine)

       IF LEN( aStruct[ nCol,2 ] ) > 1
          cSql     += Trim( aStruct[ nCol,2 ] ) + ' '

       ELSE
          DO CASE

             CASE aStruct[ nCol,2 ] = '+'
                  lAddAutoInc := .T.
                  cSql  += { "", " AUTOINCREMENT", " INT IDENTITY( 1, 1 )", " INT AUTO_INCREMENT",;
                            " INT", " INTEGER"," NUMERIC"," SERIAL"," SERIAL",  " INTEGER IDENTITY", " AUTOINC" }[ snDbms ]

                  IF dbEngine <> "ADS"
                     cSql  += " PRIMARY KEY"

                  ENDIF

             CASE aStruct[ nCol,2 ] = '='
                  cSql  += { "", " DATETIME NOT NULL DEFAULT Now()", " DATETIME NOT NULL DEFAULT (GetDate())", ;
                            " TIMESTAMP DEFAULT CURRENT_TIMESTAMP", " DATE DEFAULT SysDate", ;
                            " DATETIME  DEFAULT CURRENT_TIMESTAMP","","","","","" }[ snDbms ]

             CASE aStruct[ nCol,2 ] = 'C'
                  cSql  += " VARCHAR ( " + LTrim( Str( aStruct[nCol, 3 ] ) ) + " )"

                  IF dbEngine == "ORACLE"
                     cSql  := STRTRAN( cSql, "VARCHAR", "VARCHAR2" )

                  ENDIF

             CASE aStruct[ nCol,2 ] = 'c'
                  IF dbEngine == "ORACLE"
                     cSql  += " VARCHAR2 ( " + LTrim( Str( aStruct[ nCol,3 ] ) ) + " )"

                  ELSE
                     cSql  += " VARBINARY ( " + LTrim( Str( aStruct[ nCol,3 ] ) ) + " )"

                  ENDIF

             CASE aStruct[ nCol,2 ] = 'D'
                  IF dbEngine == "MSSQL"
                     cSql  += " DateTime"  // Date dataype not compat with older servers
                                       // Even with latest providers there are some issues in usage
                  ELSE
                     cSql  += " DATE"

                  ENDIF

             CASE aStruct[ nCol,2 ] = '@'

             CASE aStruct[ nCol,2 ] = 'T'
                  cSql  += If( dbEngine == "ORACLE", " DATE", " DATETIME" )

             CASE aStruct[ nCol,2 ] = 'L'
                  IF dbEngine == "ORACLE"
                     cSql  += " NUMBER(1,0) DEFAULT 0 CHECK ( " + aStruct[ nCol,1 ] + " IN ( 0, 1 ) )"
                  ELSEIF dbEngine == "ADS"
                     cSql  += " LOGICAL"
                  ELSE
                     cSql  += " BIT DEFAULT 0"
                  ENDIF

             CASE aStruct[ nCol,2 ] = 'M'
                  cSql  += { "", " MEMO", " TEXT", " TEXT", " CLOB", " TEXT", " TEXT",;
                           " TEXT", " TEXT", " TEXT", " MEMO" }[ snDbms ]

             CASE aStruct[ nCol,2 ] = 'P'

             CASE aStruct[ nCol,2 ] = 'm'
                 IF dbEngine == "MSSQL" .AND. nVer < 9.0
                    cSql  += " IMAGE"
                 ELSE
                    cSql  += { "", " LONGBINARY", " VARBINARY(max)", " LONGBLOB", " BLOB", " BLOB",;
                               " BLOB", " BYTEA", "BLOB", " IMAGE", " BLOB" }[ snDbms ]
                 ENDIF

             CASE aStruct[ nCol,2 ] = 'N' .OR. aStruct[ nCol,2 ] = 'I'
                  c  := LTrim( Str( aStruct[ nCol,3 ] + 1 ) ) + ", " + LTrim( Str( aStruct[ nCol,4 ] ) )

                  IF dbEngine == "ORACLE"
                     cSql  += " NUMBER( " + c + " ) DEFAULT 0"

                  ELSEIF dbEngine == "ACCESS"
                     IF aStruct[ nCol,4 ] == 0 .AND. aStruct[ nCol,3 ] <= 9
                        cSql  += If( aStruct[ nCol,3 ] <= 2, " BYTE", IF( aStruct[ nCol,3 ] <= 4, " INT", " LONG" ) )
                     ELSEIF aStruct[ nCol,4 ] == 2
                        cSql  += " MONEY"
                     ELSE
                        cSql  += " DOUBLE"  // Decimal / Numeric type has issues with older versions
                     ENDIF
                  ELSEIF dbEngine == "MSSQL"
                     IF aStruct[ nCol,4 ] == 0
                        cSql  += IF( aStruct[ nCol,3 ] <= 2, " TINYINT", IF( aStruct[ nCol,3 ] <= 4, " SMALLINT", ;
                              IF( aStruct[ nCol,3 ] <= 9, " INT", " BIGINT" ) ) )
                     ELSEIF aStruct[ nCol,4 ] == 2
                        cSql  += " MONEY"
                     ELSE
                        cSql  += " DECIMAL( " + c + " )"
                     ENDIF
                  ELSEIF dbEngine == "MYSQL"
                     IF aStruct[ nCol,4 ] == 0
                        cSql  += IF( aStruct[ nCol,3 ] <= 2, " TINYINT", If( aStruct[ nCol,3 ] <= 4, " SMALLINT", ;
                                 IF( aStruct[ nCol,3 ] <= 9, " INT", " BIGINT" ) ) )
                     ELSE
                        cSql  += " DECIMAL( " + c + " )"
                     ENDIF
                  ELSEIF dbEngine == "SQLITE"
                     IF aStruct[ nCol,4 ] == 0
                        cSql  += IF( aStruct[ nCol,3 ] <= 9, " INT", " BIGINT" )
                     ELSE
                        cSql  += " DOUBLE"
                     ENDIF
                  ELSE
                     cSql  += " NUMERIC( " + c + " )"

                  ENDIF

             OTHERWISE

                  cSql  += " CHAR(1) "//+aStruct[ nCol,2 ]

          ENDCASE

       ENDIF

       IF nCol  < LEN( aStruct )
          cSql  += ","
       ENDIF

   NEXT

  RETURN cSql


STATIC FUNCTION ADOQUOTEDCOLSQL( cCol, dbEngine)
  cCol  := ADOUNQUOTE( cCol )

  DO CASE

     CASE dbEngine = "ACCESS"
          cCol     := '[' + cCol + ']'
     CASE dbEngine = "MSSQL"

     CASE dbEngine = "DBASE"
          cCol     := '[' + cCol + ']'

     CASE dbEngine = "SQLITE"
          cCol     := '"' + cCol + '"'

     CASE dbEngine = "MYSQL"

     CASE dbEngine = "FOXPRO"
          cCol     := '`' + cCol + '`'

     CASE dbEngine = "ORACLE"
          cCol     := STRTRAN( cCol, ' ', '_' )

     CASE dbEngine = "POSTGRE"
          cCol     := STRTRAN( cCol, ' ', '_' )

  ENDCASE

  RETURN cCol


STATIC FUNCTION ADOUNQUOTE( cCol )

  cCol    := ALLTRIM( cCol)

  if VALTYPE( cCol ) == 'C' .AND. LEFT( cCol, 1 ) $ '[`"'
     cCol    := ALLTRIM( SUBSTR( cCol, 2, LEN( cCol ) - 2 ) )

  endif

  RETURN cCol


STATIC FUNCTION ADOTEMPTABLE( DbEngine, oCon, cTable)
 LOCAL cMark := ""

  TRY
    oCon:Execute( "DROP TABLE " + cTable )
  CATCH
    //DO NOTHING IF DOE NOT EXIST
  END
  cMark := "TABLE "

/*
  suspended seems to work strange n updates
  sometimes not visible to the owner
  DO CASE

     CASE dbEngine = "ADS"
          // ALREADY INCLUDED IN ADOCCONET cMark := "TABLE #"
          cMark := "TABLE "

     CASE dbEngine = "ACCESS"
          cMark     := "TABLE "

     CASE dbEngine = "MSSQL"
          // ALREADY INCLUDED IN ADOCCONET cMark := "TABEL #"
          cMark := "TABLE "

     CASE dbEngine = "DBASE"
          cMark     := "TABLE "

     CASE dbEngine = "SQLITE"
          cMark    := "TEMPORARY TABLE "

     CASE dbEngine = "MYSQL"
          cMark :=   "TEMPORARY TABLE "

     CASE dbEngine = "FOXPRO"
          cMark     := "TABLE "

     CASE dbEngine = "ORACLE"
          cMark     := "GLOBAL TEMPORARY TABLE"

     CASE dbEngine = "POSTGRE"
          cMark     := "TEMPORARY TABLE "

  ENDCASE
*/
  RETURN cMark


STATIC FUNCTION ADOFILE( oCn, cTable, cIndex, cView)

   LOCAL lRet := .F.
   LOCAL oRs := ADOCLASSNEW( "ADODB.Recordset" )//TOleAuto():New( "ADODB.Recordset" )
   LOCAL aIndexes := ListIndex(), z, y

   IF !ADOCON_CHECK()
      RETURN HB_FAILURE

   ENDIF

   /*
   FIX LUCAS DE BELTRAN 23.05.2015
   */
   DEFAULT oCn TO oConnection

   IF EMPTY(oCn)
      MSGALERT("No connection estabilished! Please open some table to have a connection ")
      RETURN .F.

   ENDIF

   //FROM FW_ADOCREATETABLE
   IF ! EMPTY( cTable )
      TRY
          oRs      := oCn:OpenSchema( adSchemaTables, { nil, nil, cTable, "TABLE" } )
          lRet   := !( oRs:Bof .and. oRs:Eof )
          oRs:Close()

      CATCH
          // Older ADO version not supporting second parameter
         TRY
              oRs   := oCn:OpenSchema( adSchemaTables )

              IF ! oRs:Eof()
                 IF UPPER( SUBSTR( cTable ,1,3 ) ) = "TMP" .OR. UPPER( SUBSTR( cTable ,1,4 ) ) = "TEMP" //24.06.15
                    //oRs:Filter  := "TABLE_NAME = '" + cTable + "' AND TABLE_TYPE = 'LOCAL TEMPORARY'"
                    oRs:Filter  := "TABLE_NAME = '" + cTable + "' AND TABLE_TYPE = 'TABLE'"
                 ELSE
                    oRs:Filter  := "TABLE_NAME = '" + cTable + "' AND TABLE_TYPE = 'TABLE'"
                 ENDIF

                 lRet   := !( oRs:Bof .and. oRs:Eof )

              ENDIF

              oRs:Close()

          CATCH

              // OpenSchema(adSchemaTables) is not supported by provider
              // we do not know if the table exists
              ADOSHOWERROR( oCn )  // Comment out in final release

          END
      END

   ENDIF

   IF ! EMPTY( cIndex ) .AND. ! EMPTY(cTable)
      TRY
         //MAYBE IT COMES WITH FILE EXTENSION AND PATH
         cIndex := CFILENOPATH(cIndex)
         cIndex := UPPER(CFILENOEXT(cIndex))

         oRs      := oCn:OpenSchema( adSchemaIndexes, { nil, nil, cIndex, nil, cTable } )
         lRet   := !( oRs:Bof .and. oRs:Eof )
         oRs:Close()

      CATCH
          // OpenSchema(adSchemaTables) is not supported by provider
          // we do not know if the table exists
          ADOSHOWERROR( oCn )  // Comment out in final release

      END

   ENDIF

   //19.06.15 views
   IF ! EMPTY( cView )
      TRY
         oRs:Open(" SELECT TABLE_NAME FROM INFORMATION_SCHEMA.VIEWS ",oCn)

         IF ! oRs:Eof()
            oRs:Filter  := "TABLE_NAME = '" + cView+"'"
            lRet   := !( oRs:Bof .and. oRs:Eof )
         ENDIF

         oRs:Close()

      CATCH
         // OpenSchema(adSchemaViews) is not supported by provider
         // we do not know if the table exists
         ADOSHOWERROR( oCn )  // Comment out in final release

      END

   ENDIF


   RETURN lRet


STATIC FUNCTION ADODROP( oCon, cTable, cIndex ,cView, DBEngine)


   LOCAL lRet := .F.,cSql
   LOCAL aEngines := { "ACCESS","MSSQL","MYSQL","ORACLE","SQLITE",;
                     "FOXPRO","POSTGRE","INFORMIX","ANYWHERE","ADS"}


   IF EMPTY(oCon)
      MSGALERT("No connection estabilished! Please open some table to have a connection ")
      RETURN .F.

   ENDIF

   IF ASCAN(aEngines, DBEngine) = 0
      MSGALERT("DbEngine "+DBEngine +" not supported by adordd! "+;
               "Valid DBS are : ACCESS MSSQL MYSQL ORACLE SQLITE"+;
               "FOXPRO POSTGRE INFORMIX ANYWHERE ADS")
      RETURN .F.

   ENDIF

   IF ! EMPTY( cTable )
      TRY
         oCon:Execute( "DROP TABLE " + cTable )
         lRet := .T.

      CATCH
         ADOSHOWERROR( oCon, .f. )

      END

   ENDIF

   IF ! EMPTY( cTable ) .AND.  ! EMPTY( cIndex )
      TRY
         DO CASE

            CASE DBEngine == "ACCESS"
                 cSql  := "DROP INDEX " + cIndex + " ON " + cTable

            CASE DBEngine == "MSSQL" .OR. DBEngine == "ADS"
                 cSql  := "DROP INDEX " + cTable + '.' + cIndex

            CASE DBEngine == "MYSQL"
                 cSql  := "ALTER TABLE " + cTable + " DROP INDEX " + cIndex

         OTHERWISE
                 cSql  := "DROP INDEX " + cIndex

         ENDCASE

         oCon:Execute( cSql )
         lRet := .T.

      CATCH
          ADOSHOWERROR( oCon, .f. )

      END

   ENDIF

   IF ! EMPTY( cView )
      TRY
         oCon:Execute( "DROP VIEW " + cView )
         lRet := .T.

      CATCH
         ADOSHOWERROR( oCon, .f. )

      END

   ENDIF

   RETURN lRet
/*                             END FILE RELATED FUNCTION */


/*                                     GENERAL */
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
          uReturn := ADOLUPDATE(  aWAData  )

     CASE uInfoType == DBI_GETDELIMITER // 5  /* The delimiter (as a string)         */
          uReturn := ""

     CASE uInfoType == DBI_SETDELIMITER // 6  /* The delimiter (as a string)         */
          uReturn := ""

     CASE uInfoType == DBI_GETRECSIZE // 7  /* The size of 1 record in the file    */
          ADO_RECINFO( nWA, ADO_RECID( nWA, @uReturn ), DBRI_RECSIZE, @uReturn )

     CASE uInfoType == DBI_GETLOCKARRAY // 8  /* An array of locked records' numbers */
          uReturn := aWAData[WA_LOCKLIST]

     CASE uInfoType == DBI_TABLEEXT //  9  /* The data file's file extension      */
          uReturn := ""

     CASE uInfoType == DBI_FULLPATH // 10  /* The Full path to the data file      */
          uReturn := aWAData[WA_TABLENAME]

     CASE uInfoType == DBI_ISFLOCK // 20  /* Is there a file lock active?        */
          uReturn := aWAData[WA_FILELOCK]

     CASE uInfoType == DBI_CHILDCOUNT // 22  /* Number of child relations set       */
          uReturn := IF(LEN(aWAData[WA_PENDINGREL]) > 0, LEN(aWAData[WA_PENDINGREL]) / 7,0)

     CASE uInfoType == DBI_FILEHANDLE // 23  /* The data file's OS file handle      */
          uReturn := -1

     CASE uInfoType == DBI_BOF // 26  /* Same as bof()    */
          uReturn := aWAData[WA_BOF]

     CASE uInfoType == DBI_EOF // 27  /* Same as eof()    */
          uReturn := aWAData[WA_EOF]

     CASE uInfoType == DBI_DBFILTER // 28  /* Current Filter setting              */
          uReturn := aWAData[WA_CFILTERACTIVE]

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
          uReturn := "{||"+aWAData[ WA_LOCATEFOR ]+"}"

     CASE uInfoType == DBI_LOCKOFFSET //  35  /* The offset used for logical locking */
          uReturn := 0

     CASE uInfoType == DBI_LOCKSCHEME     //     128  /* Locking scheme used by RDD */
          uReturn := 0

     CASE uInfoType == DBI_SHARED  //  36  /* Was the file opened shared?         */
          uReturn := aWAData[WA_OPENSHARED]

     CASE uInfoType == DBI_MEMOEXT  //  37  /* The memo file's file extension      */
          uReturn := ""

     CASE uInfoType == DBI_MEMOHANDLE // 38  /* File handle of the memo file        */
          uReturn := -1

     CASE uInfoType == DBI_MEMOBLOCKSIZE  // 39  /* Memo File's block size              */
          uReturn := 0

    CASE uInfoType == DBI_ISREADONLY
         uReturn := .F.

    CASE uInfoType == DBI_DB_VERSION  //  101  /* Version of the Host driver          */
         uReturn := "Version 2015"

     CASE uInfoType == DBI_RDD_VERSION // 102  /* current RDD's version               */
          uReturn := "Version 2015"

  ENDCASE

 RETURN HB_SUCCESS //uReturn


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
   aADOFunc[ UR_SKIPFILTER ]   := (@ADO_SKIPFILTER())
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
   aADOFunc[ UR_RECALL ]       := (@ADO_RECALL())
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
   aADOFunc[ UR_FILTERTEXT ]   := (@ADO_FILTERTEXT())
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
   aADOFunc[ UR_ORDLSTREBUILD ]:= (@ADO_ORDLSTREBUILD())
   aADOFunc[ UR_EVALBLOCK ]    := (@ADO_EVALBLOCK())
   aADOFunc[ UR_SEEK ]         := (@ADO_SEEK())
   aADOFunc[ UR_TRANS ]        := (@ADO_TRANS())

   RETURN USRRDD_GETFUNCTABLE( pFuncCount, pFuncTable, pSuperTable, nRddID, ;
      /* NO SUPER RDD */, aADOFunc )


INIT PROCEDURE ADORDD_INIT()

   rddRegister( "ADORDD", RDT_FULL )

   RETURN


STATIC FUNCTION ADODTOS(cDate)
 LOCAL cYear,cMonth,cDay,dDate


   // IF YOU HAVE ADOFUNCS.PRG COMMENT THESE AND UNCOMMNED FW_DateToSQL( dDate, cType, lOracle )
   IF AT("/",cDate) = 0 .AND. AT("-",cDate) = 0 //DTOS FORMAT
      dDate := STOD(cDate)

   ELSE
      dDate := CTOD(cDate)

   ENDIF

   //cDate := Transform( DToS( FW_TToD( dDate ) ), "@R 9999-99-99" )

   cDate := DTOS(dDate)

   cYear  := SUBSTR(cDate,1,4)
   cMonth := SUBSTR(cDate,5,2)
   cDay   := SUBSTR(cDate,7,2)

   IF( EMPTY(cYear),cYear :="1901",cYear)
   IF( EMPTY(cMonth),cMonth := "01",cMonth)
   IF( EMPTY(cDay),cDay := "01",cDay)

   cDate  := cYear+"-"+cMonth+"-"+cDay // hope to enforce set date format like this

   RETURN cDate


STATIC FUNCTION ADOEMPTYSET(oRecordSet)
   RETURN  (oRecordSet:Eof() .AND.  oRecordSet:Bof() )


//from adufuncs.prg
STATIC FUNCTION SQLTranslate( cFilter )
  local cWhere
  local nAt, nLen, cToken, cDate, n
  local afunctions := {"STR(","VAL(","CVALTOCHAR(",;
                       'SOUNDEX(', "ABS(","ROUND(","LEN(","ALLTRIM(","LTRIM(","RTRIM(",;
                       "UPPER(","LOWER(","SUBSTR(",;
                       "SPACE(","DATE(","YEAR(","MONTH(",;
                       "DAY(","TIME(","IF("}
  local areplaces := { "","",""," LIKE ","","","","","","","","","","","","","","","",""}

   cWhere      := Upper( cFilter )
   //cWhere      := StrTran( StrTran( cWhere, "'", "''" ), '"', "'" )
   cWhere      := StrTran( cWhere, '"', "'" )
   cWhere      := StrTran( StrTran( cWhere, ".AND.", "AND" ), ".OR.", "OR" )
   cWhere      := StrTran( StrTran( cWhere, ".T.", "1" ), ".F.", "0" )
   cWhere      := StrTran( cWhere, "==", "=" )
   cWhere      := StrTran( cWhere, "!=", "<>" )
   cWhere      := StrTran( cWhere, "!", " NOT " )
   cWhere      := StrTran( cWhere, Alias()+"->", "" )
   cWhere      := StrTran( cWhere, "FIELD->", "" )
   if At( "!DELETED()", cWhere ) == 1; cWhere   := LTrim( SubStr( cWhere, 11 ) ); endif
   if At( "AND", cWhere ) == 1; cWhere := LTrim( SubStr( cWhere, 4 ) ); endif
   if At( "OR", cWhere ) == 1; cWhere := LTrim( SubStr( cWhere, 3 ) ); endif

   if  At("$",cWhere) > 0
      cWhere := InvertArgs(cWhere,"$")

   endif

   // Now handle dates its adpated from adofuncs because it was only considering one occurrence
   do while .t.

      for each cToken in { "STOD(", "CTOD(", "HB_STOT(", "HB_CTOT(", "STOT(", "CTOT(", "{^" }
          nAt    := At( cToken, cWhere )
          if nat > 0
             exit
          endif
      next

      if nAt = 0
         exit
      endif

      for each cToken in { "STOD(", "CTOD(", "HB_STOT(", "HB_CTOT(", "STOT(", "CTOT(", "{^" }
          nAt    := At( cToken, cWhere )
          if nAt > 0
             if Left( cToken, 1 ) == "{"
                nLen  := At( "}", SubStr( cWhere, nAt ) )
             else
                nLen  := At( ")", SubStr( cWhere, nAt ) )
             endif
             cDate := SubStr( cWhere, nAt, nLen )
#ifdef __XHARBOUR__
             if Left( cDate, 3 ) == "HB_"; cDate := SubStr( cDate, 4 ); endif
#else
             if Left( cDate, 5 ) $ "STOT(,CTOT("
                cDate    := "HB_" + cDate
             endif
             if Left( cDate, 2 ) = "{^"
                cDate    := LTrim( SubStr( cDate, 3 ) )
                cDate    := If( ':' $ cDate, "HB_STOT('", "HB_STOD('" ) + cDate + "')"
                cDate    := CharRem( "/-:} ", cDate )
             endif

#endif
             cDate  := &cDate
             cWhere := Stuff( cWhere, nAt, nLen,  DateToADO( cDate ) )
          endif

      next

   enddo

   for n:= 1 to len(afunctions)
       cWhere := StrTran( cWhere, afunctions[n], areplaces[n] )

   next

   cWhere      := StrTran( cWhere, ")", "" )


return cWhere


function InvertArgs(cString,cChar)
  local n, aTokens := HB_ATokens( cString, " ", .t. )
  local cBefore,cAfter

      for n := 1 TO Len( aTokens )
          if aTokens[n] = cChar
             aTokens[n] := " LIKE "
             cBefore := aTokens[n-1]
             cAfter := aTokens[n+1]
             aTokens[n-1] := cAfter
             aTokens[n+1] := cBefore

          endif

      next

      cString := ""

      for n:= 1 TO Len( aTokens )
          cString += aTokens[n]+" "

      next

return cString


function DateToADO( dDate, cType )

   local cRet

   if Empty( dDate )
      return nil
   endif

   DEFAULT cType TO ValType( dDate )

   if cType == 'T'
      cRet  := Transform( TToS( FW_DToT( dDate ) ), "@R 9999-99-99 99:99:99" )

   else
      cRet  := Transform( DToS( FW_TToD( dDate ) ), "@R 9999-99-99" )

   endif

return '#' + cRet + '#'


STATIC FUNCTION ADOCON_CHECK()
 LOCAL lCnOpened := .F.

   IF oConnection != NIL
      IF oConnection:State == 0
         oConnection:Close()
         TRY
            oConnection:Open()
            lCnOpened := .T.
         CATCH
            lCnOpened := .F.
         END
      ELSE
         lcnOpened := .T.
      ENDIF

   ENDIF

   RETURN lCnOpened


STATIC FUNCTION ADOLUPDATE(  aWAData  )
 LOCAL dDate
 LOCAL oRs := TempRecordSet()

  DO CASE

     CASE  aWAData[ WA_ENGINE ] = "MYSQL"
           oRs:Open(  "SELECT UPDATE_TIME FROM information_schema.tables "+;
                "WHERE  TABLE_SCHEMA = '"+ aWAData[ WA_CATALOG ] +"' AND TABLE_NAME = '"+ aWAData[ WA_TABLENAME ] +"'" )
           dDate := oRs:Fields("UPDATE_TIME"):Value
           oRs:Close()

     CASE  aWAData[ WA_ENGINE ] = "MSSQL"
           oRs:Open(  "SELECT OBJECT_NAME(OBJECT_ID) AS DatabaseName, last_user_update,*"+;
                      "FROM sys.dm_db_index_usage_stats"+;
                      "WHERE database_id = DB_ID( '"+ aWAData[ WA_CATALOG ] +"')"+;
                      "AND OBJECT_ID=OBJECT_ID('"+ aWAData[ WA_TABLENAME ] +"')"  )
           dDate := oRs:Fields("last_user_update"):Value
           oRs:Close()

      OTHERWISE
          dDate := CTOD( "31/12/1899" )
          MSGINFO("You are requesting last date table update"+CRLF+;
                  "Adordd does not support it yet to your Server!"+CRLF+;
                  "Date returned is: "+ CTOD( dDate ) )

  ENDCASE


  RETURN dDate

/*                                  END  GENERAL */


/*                    ADO SET GET FUNCTONS */

STATIC FUNCTION ADOCLASSNEW( cString )
 LOCAL cClass := ADODEFAULTS()[6]
  RETURN IF( cClass <> NIL, cClass:New( cString ), TOleAuto():New( cString ) )


STATIC FUNCTION ADOSTATUSMSG( nStatus, fName, xValue, xNewValue, LockType )
 LOCAL aMsgs := {"16 The record was not saved because its bookmark is invalid.",;
                 "64 The record was not saved because it would have affected multiple records.",;
                 "128 The record was not saved because it refers to a pending insert.",;
                 "256 The record was not saved because the operation was canceled.",;
                 "1024 The new record was not saved because of existing record locks.",;
                 "2048 The record was not saved because optimistic concurrency was in use.",;
                 "4096 The record was not saved because the user violated integrity constraints.",;
                 "8192 The record was not saved because there were too many pending changes.",;
                 "16384 The record was not saved because of a conflict with an open storage object.",;
                 "32768 The record was not saved because the computer has run out of memory.",;
                 "65536 The record was not saved because the user has insufficient permissions.",;
                 "131072 The record was not saved because it violates the struture of the underlying database.",;
                 "262144 The record has already been deleted from the data source." }
 LOCAL nPos, cErr := ""
  cErr     += CRLF + PROCNAME( 2 )  + cValToChar( PROCLINE( 2 ) )
  cErr     += CRLF + PROCNAME( 3 )  + cValToChar( PROCLINE( 3 ) )
  cErr     += CRLF + PROCNAME( 4 )  + cValToChar( PROCLINE( 4 ) )
  cErr     += CRLF + PROCNAME( 5 )  + cValToChar( PROCLINE( 5 ) )
  cErr     += CRLF + PROCNAME( 6 )  + cValToChar( PROCLINE( 6 ) )
  cErr     += CRLF + PROCNAME( 7 )  + cValToChar( PROCLINE( 7 ) )

  nPos := ASCAN( aMsgs, {| x | AT( ALLTRIM( STR( nStatus ) ), x ) > 0 } )
  MSGALERT( "Field name "+fName+CRLF+" Actual value "+CVALTOCHAR( xValue ) +CRLF+;
   " New value "+CVALTOCHAR( xNewValue )+CRLF+;
   " Record number "+STR( RECNO( ) )+" Lock Type "+STR( locktype )+ CRLF+ aMsgs[ nPos ]+CRLF+cErr )

  RETURN NIL


FUNCTION ADOSHOWERROR( oCn, lSilent )

   LOCAL nErr, oErr, cErr

   DEFAULT oCn TO oConnection
   DEFAULT lSilent TO .F.

   IF ( nErr := oCn:Errors:Count ) > 0
      oErr  := oCn:Errors( nErr - 1 )
      IF ! lSilent
         WITH OBJECT oErr
            cErr     := oErr:Description
            cErr     += CRLF + 'Source       : ' + oErr:Source
            cErr     += CRLF + 'NativeError  : ' + cValToChar( oErr:NativeError )
            cErr     += CRLF + 'Error Source : ' + oErr:Source
            cErr     += CRLF + 'Sql State    : ' + oErr:SQLState
            cErr     += CRLF + REPLICATE( '-', 50 )
            cErr     += CRLF + PROCNAME( 1 ) + "( " + cValToChar( PROCLINE( 1 ) ) + " )"
            cErr     += CRLF + PROCNAME( 2 )  + cValToChar( PROCLINE( 2 ) )
            cErr     += CRLF + PROCNAME( 3 )  + cValToChar( PROCLINE( 3 ) )
            cErr     += CRLF + PROCNAME( 4 )  + cValToChar( PROCLINE( 4 ) )
            cErr     += CRLF + PROCNAME( 5 )  + cValToChar( PROCLINE( 5 ) )
            cErr     += CRLF + PROCNAME( 6 )  + cValToChar( PROCLINE( 6 ) )
            cErr     += CRLF + PROCNAME( 7 )  + cValToChar( PROCLINE( 7 ) )

            MSGALERT( cErr, IF( oCn:Provider = NIL, "ADO ERROR",oCn:Provider ) )
         END
      ENDIF

   ELSE
      MSGALERT( "ADO ERROR UNKNOWN"+;
                CRLF + PROCNAME( 1 )  + cValToChar( PROCLINE( 1 ) ) +;
                CRLF + PROCNAME( 2 )  + cValToChar( PROCLINE( 2 ) )+;
                CRLF + PROCNAME( 3 )  + cValToChar( PROCLINE( 3 ) )+;
                CRLF + PROCNAME( 4 )  + cValToChar( PROCLINE( 4 ) )+;
                CRLF + PROCNAME( 5 )  + cValToChar( PROCLINE( 5 ) )+;
                CRLF + PROCNAME( 6 )  + cValToChar( PROCLINE( 6 ) )+;
                CRLF + PROCNAME( 7 )  + cValToChar( PROCLINE( 7 ) )  )

   ENDIF

   RETURN oErr


PROCEDURE hb_adoSetDSource( cDB )

   t_cDataSource := cDB

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


FUNCTION hb_adoRddGetTableName( nWA )

   LOCAL aWAData

   IF ! HB_ISNUMERIC( nWA )
      nWA := Select()

   ENDIF

   aWAData := USRRDD_AREADATA( nWA )

   RETURN iif( aWAData != NIL, aWAData[ WA_TABLENAME ], NIL )


FUNCTION hb_adoRddGetTables( oCn )

   LOCAL lRet := .T., aTables := {}
   LOCAL oRs := ADOCLASSNEW( "ADODB.Recordset" )//TOleAuto():New( "ADODB.Recordset" )

   IF !ADOCON_CHECK()
      RETURN HB_FAILURE

   ENDIF

   DEFAULT oCn TO oConnection

   IF EMPTY(oCn)
      MSGALERT("No connection estabilished! Please open some table to have a connection ")
      RETURN .F.

   ENDIF

   TRY
       oRs      := oCn:OpenSchema( adSchemaTables )
       lRet   := !( oRs:Bof .and. oRs:Eof )

   CATCH
          // Older ADO version not supporting second parameter
       TRY
          oRs   := oCn:OpenSchema( adSchemaTables )

          IF ! oRs:Eof()
             oRs:Filter  := "TABLE_NAME = '" + cTable + "' AND TABLE_TYPE = 'TABLE'"
             lRet   := !( oRs:Bof .and. oRs:Eof )

          ENDIF

       CATCH
          lRet := .F.
          // OpenSchema(adSchemaTables) is not supported by provider
          // we do not know if the table exists
          ADOSHOWERROR( oCn )  // Comment out in final release

       END
   END

   IF lRet
      oRs:MoveFirst()
      DO WHILE ! oRs:Eof()
         AADD( aTables, oRs:Fields( 2 ):Value )
         oRs:moveNext()
      ENDDO

   ENDIF

   oRs:Close()

  RETURN aTables


FUNCTION hb_adoRddExistsTable( oCn, cTable, cIndex, cView )
   RETURN ADOFILE( oCn, cTable, cIndex, cView )

FUNCTION hb_adoRddDrop( oCn, cTable, cIndex, cView, DBEngine )
   RETURN ADODROP( oCn, cTable, cIndex, cView,  DBEngine )


// field name bit to use as deleted per each table {{"CTABLE","CFIELDNAME"} }
FUNCTION ListFieldDeleted( aList )

 STATIC aListFieldDelete

  IF !EMPTY(aList)
     aListFieldDelete := aList
  ENDIF

  RETURN aListFieldDelete


/* default field to be used as deleted */
FUNCTION ADODEFLDDELETED( cFieldName )
 STATIC cDelName := "HBDELETED"

  IF !EMPTY(cFieldName)
     cDelName := cFieldName
  ENDIF

  RETURN cDelName


FUNCTION ListIndex(aList) //ATTENTION ALL MUST BE UPPERCASE
//index files array needed for the adordd for your application
//order expressions already translated to sql DONT FORGET TO replace taitional + sign with ,
//we can and should include the SQL CONVERT to translate for ex DTOS etc
//ARRAY SPEC { {"TABLENAME",{"INDEXNAME","INDEXKEY","WHERE EXPRESSION AS USED FOR FOREXPRESSION","UNIQUE - DISTINCT ANY SQL STAT BEFORE * FROM"} }
//temporary indexes are not included gere they are create on fly and added to temindex list array
//they are only valid through the duration of the application
//the temp index name is auto given by adordd

 STATIC Alista_fic

   IF !EMPTY(aList)
      Alista_fic := aList

   ENDIF

  RETURN Alista_fic


// array with same tables and indexes as lustindex but with original clipper index expressions
//aray has to be the same structure as for ListIndex (see above)
//indexes not present inthis list will return indexexpressions as per ListIndex
FUNCTION ListDbfIndex( aList )
 STATIC AClipper_fic

   IF !EMPTY(aList)
      AClipper_fic := aList

   ENDIF

  RETURN AClipper_fic


// field name autoinc to use as recno per each table {{"CTABLE","CFIELDNAME"} }
FUNCTION ListFieldRecno( aList )

 STATIC aListFieldRecno

    IF !EMPTY(aList)
      aListFieldRecno := aList
   ENDIF

   RETURN aListFieldRecno


//index temporary names {"TMP","TEMP","ETC"}
FUNCTION ListTmpNames(aList)
 STATIC aTmpNames

   IF !EMPTY(aList)
      aTmpNames := aList

   ENDIF

   RETURN aTmpNames


//index udfs

FUNCTION ListUdfs(aList)
 STATIC aUdfs

   IF !EMPTY(aList)
      aUdfs := aList

   ENDIF

   RETURN aUdfs

 /* default values for adordd to use if not present in open or create */
FUNCTION ADODEFAULTS( cDB, cServer, cEngine, cUser, cPass, cClass, lGetThem )
 STATIC aDefaults := {}

   DEFAULT lGetThem TO .T. //DEFAULT CALL WITHOUT PARAMS GET DEFAULT ARRAY

   IF !lGetThem	//RESET THEM
      aDefaults := {}
      AADD(aDefaults, cDB )
      AADD(aDefaults, cServer )
      AADD(aDefaults, cEngine )
      AADD(aDefaults, cUser )
      AADD(aDefaults, cPass )
      AADD(aDefaults, cClass )
      oConnection := ADOGETCONNECT(  cDB, cServer, cEngine, cUser, cPass  )

   ELSE
      DEFAULT t_cQuery TO "SELECT * FROM "
      DEFAULT t_cUserName TO aDefaults[4]
      DEFAULT t_cPassword TO aDefaults[5]
      DEFAULT t_cServer TO aDefaults[2]
      DEFAULT t_cEngine TO aDefaults[3]
      DEFAULT t_cDataSource TO aDefaults[1]

   ENDIF

   RETURN aDefaults


STATIC FUNCTION ADOGETCONNECT( cDB, cServer, cEngine, cUser, cPass  )
 LOCAL oCn := ADOCLASSNEW( "ADODB.Connection" )//TOleAuto():New( "ADODB.Connection" )

  TRY

     oCn := ADOOPENCONNECT( cDB, cServer, cEngine, cUser, cPass , oCn )

  CATCH
     ADOSHOWERROR( oCn )
     oCn := nil
     QUIT  //lucas deBeltran

  END

  RETURN oCn


/* default field to be used as recno */
FUNCTION ADODEFLDRECNO( cFieldName )
 STATIC cName := "HBRECNO"

  IF !EMPTY(cFieldName)
      cName := cFieldName
  ENDIF

   RETURN cName


/* THESE ARE FILLED WITH INFORMATION FROM ADO_CREATE (INDEX) THEY ONLY LIVE THROUGH APP*/
STATIC FUNCTION ListTmpIndex(aList)
 STATIC aTmpIndex := {}
  RETURN aTmpIndex

//sql index exp
STATIC FUNCTION ListTmpExp(aList)
 STATIC aTmpExp := {}
  RETURN aTmpExp

//dbf index exp
STATIC FUNCTION ListTmpDbfExp(aList)
 STATIC aTmpDbfExp := {}
  RETURN aTmpDbfExp


//SQL FOR EXP
STATIC FUNCTION ListTmpFor(aList)
 STATIC aTmpFor := {}
  RETURN aTmpFor


//DBF FOR EXP
STATIC FUNCTION ListTmpDbfFor()
 STATIC aTmpDbfFor := {}
  RETURN aTmpDbfFor


//dbf UNIQUE EXP
STATIC FUNCTION ListTmpDbfUnique()
 STATIC aTmpDbfUnique := {}
 RETURN aTmpDbfUnique


//DBF UNIQUE EXP
STATIC FUNCTION ListTmpUnique()
  STATIC aTmpUniques := {}
  RETURN aTmpUniques

//index temporary descend
FUNCTION ListTmpDbfDesc()
 STATIC aTmpDesc := {}
 RETURN aTmpDesc

//bookmarks tmp indexes
FUNCTION ListTmpBookM()
 STATIC aTmpBookM := {}
 RETURN aTmpBookM


STATIC FUNCTION TempRecordSet() //USED IN ADO_SEEK AVOID OVERTIME NEW OBJ RECORDSET
  STATIC oRs

  IF EMPTY(oRs)
     oRs := ADOCLASSNEW( "ADODB.Recordset" )//TOleAuto():New( "ADODB.Recordset" )
  ENDIF

  RETURN oRs

FUNCTION hb_GetAdoConnection()//supply app the con object
  RETURN oConnection


//ceate table for record lock control
FUNCTION ADOLOCKCONTROL(cPath,cRdd)
 STATIC cFile,rRdd

 LOCAL cTable
 LOCAL cIndex

 FIELD CODLOCK

  DEFAULT cRdd TO "DBFCDX"
  DEFAULT cPath TO SUBSTR(ALLTRIM(cFilePath( GetModuleFileName( GetInstance() ) ) ),1,;
                       LEN(ALLTRIM(cFilePath( GetModuleFileName( GetInstance() ) ) ))-1)
  IF EMPTY(rRdd)
     rRdd := cRdd

  ENDIF

  IF EMPTY(cFile)
     cFile := cPath+"\TLOCKS"

  ENDIF

  cTable := cPath+"\TLOCKS"+RDDINFO(RDDI_TABLEEXT,,rRdd)
  cIndex := cPath+"\TLOCKS"+RDDINFO(RDDI_ORDBAGEXT,,rRdd)

  IF !FILE(cTable)
     DBCREATE(cTable,;
              { {"CODLOCK","C",50,0 }},;
              rRdd,.T.,"TLOCKS")
     INDEX ON CODLOCK TO (cIndex)

  ENDIF


  RETURN {cFile,rRdd}


FUNCTION ADOFORCELOCKS(lOn) //force lock control buy ado
 STATIC lLockScheme := .T.

  IF VALTYPE( lOn ) = "L"
     lLockScheme := lOn

  ENDIF

  RETURN lLockScheme


FUNCTION ADOVIRTUALDELETE(lOn) //Use Field to delete Records
 STATIC lDeleteScheme := .F.

  IF VALTYPE( lOn ) = "L"
     lDeleteScheme := lOn

  ENDIF

  RETURN lDeleteScheme


FUNCTION ADOVERSION()
//version string = nr of version . post date() / sequencial nr in the same post date
RETURN "AdoRdd Version 1.081015/1"

/*                   END ADO SET GET FUNCTONS */

function fwAdoConnect()

   local oDL,   cConnection := ""

   oDL = CreateObject( "Datalinks" ):PromptNew()

   if ! Empty( oDL )
      cConnection = oDL:ConnectionString
   endif


return nil

//TO DEBGUG
FUNCTION ARRAYTOCHAR(AARRAY)
LOCAL N

FOR N:= 1 TO LEN(AARRAY)
   AARRAY[N] := CVALTOCHAR(AARRAY[N] )

NEXT

RETURN AARRAY
//----------------------------------------------------------------------------//

#ifdef __XHARBOUR__

function AdoNull()   ; return VTWrapper( 1, nil )
function AdoDefault(); return OleDefaultArg()

#else

#define WIN_VT_NULL                  1
#define WIN_VT_ERROR                10
#define WIN_DISP_E_PARAMNOTFOUND ( 0x80020004 )

function AdoNull()   ; return __OleVariantNew( WIN_VT_NULL)  // WIN_VT_NULL

function AdoDefault()

   local pFunc := HB_FuncPtr( "OLEDEFAULTARG" )

   if Empty( pFunc )
      return   __oleVariantNew( WIN_VT_ERROR, WIN_DISP_E_PARAMNOTFOUND )
   endif

return HB_Exec( pFunc )

#endif

//----------------------------------------------------------------------------//

#pragma BEGINDUMP

#include <windows.h>
#include <hbapi.h>

HB_FUNC( FW_TTOD )
{
   hb_retdl( hb_pardl( 1 ) );
}

HB_FUNC( FW_DTOT )
{

#ifdef __XHARBOUR__
   hb_retdtl( hb_pardl( 1 ), hb_part( 1 ) );
#else
   long lJulian;
   long lMilliSecs;

   hb_partdt( &lJulian, &lMilliSecs, 1 );
   hb_rettdt( lJulian, lMilliSecs );
#endif
}

#pragma ENDDUMP

function ArrTranspose( aArray, lSquare )

   local nRows, nCols, nRow, nCol, nWidth
   local aNew

   DEFAULT lSquare TO .f.

   nRows          := Len( aArray )
   if lSquare
      nCols       := Len( aArray[ 1 ] )
   else
      nCols       := 1
      for nRow := 1 to nRows
         if ValType( aArray[ nRow ] ) == 'A'
            nCols    := Max( nCols, Len( aArray[ nRow ] ) )
         endif
      next
   endif

   aNew           := Array( nCols, nRows )
   for nRow := 1 to nRows
      if ValType( aArray[ nRow ] ) == 'A'
         nWidth  := Len( aArray[ nRow ] )
         for nCol := 1 to nWidth
            aNew[ nCol, nRow ]   := aArray[ nRow, nCol ]
         next
      else
         aNew[ 1, nRow ]      := aArray[ nRow ]
      endif
   next

return aNew


#ifndef __XHARBOUR__

   #xcommand TRY  => BEGIN SEQUENCE WITH {| oErr | Break( oErr ) }
   #xcommand CATCH [<!oErr!>] => RECOVER [USING <oErr>] <-oErr->
   #xcommand FINALLY => ALWAYS

   #include "fivewin.ch"        // as Harbour does not have TRY / CATCH IF YOU DONT HAVE COMENT THIS LINE
   #define UR_FI_FLAGS           6
   #define UR_FI_STEP            7
   #define UR_FI_SIZE            5 // by Lucas for Harbour

//13.04.15 functions given by thefull to compile with Harbour WITHOUT FIVEWIN
function cValToChar( u ); return CStr( u )
function MsgInfo( u ) ; return Alert( u )
function MsgAlert( u ); return Alert( u )

function cFilePath( cPathMask )   // returns path of a filename

   local n := RAt( "\", cPathMask ), cDisk

return If( n > 0, Upper( Left( cPathMask, n ) ),;
           ( cDisk := cFileDisc( cPathMask ) ) + If( ! Empty( cDisk ), "\", "" ) )

function cFileNoPath( cPathMask )

    local n := RAt( "\", cPathMask )

return If( n > 0 .and. n < Len( cPathMask ),;
           Right( cPathMask, Len( cPathMask ) - n ),;
           If( ( n := At( ":", cPathMask ) ) > 0,;
           Right( cPathMask, Len( cPathMask ) - n ),;
           cPathMask ) )

function cFileNoExt( cPathMask ) // returns the filename without ext

   local cName := AllTrim( cFileNoPath( cPathMask ) )
   local n     := RAt( ".", cName )

return AllTrim( If( n > 0, Left( cName, n - 1 ), cName ) )

function cFileDisc( cPathMask )  // returns drive of the path

return If( At( ":", cPathMask ) == 2, ;
           Upper( Left( cPathMask, 2 ) ), "" )

#pragma BEGINDUMP
#include <hbapi.h>

HB_FUNC( LAND )
{
   hb_retl( ( hb_parnl( 1 ) & hb_parnl( 2 ) ) != 0 );
}

#pragma ENDDUMP

#endif

