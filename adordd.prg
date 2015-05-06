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
 *
 * Most part has been rewriten with a diferent kind of approach
 * not deal with Catalogs File indexes - DBA responsability
 * converting indexes to selects and treat indexes as "virtual" as they really dont exist as files
 * Only allows to create temp tables
 * Seek translate to find and if it is to slow can be converted to select but after using the result seek
 * one must call a resetseek to revert to the previous select SEE ADOSEEL.. AND ADORSETSEEK
 * Tables must have some ID field to be used as RECNO
 * Sqlparser that converts filters using standard clipper expressions to sql taking into account
 * the active indexes to build order by clause 
 * Preventing sql injection TO DO
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
#define WA_LASTRELKEY  30//AHF
#define WA_FILTERACTIVE  31//AHF
#define WA_FIELDRECNO 32//AHF
#define WA_TLOCKS 33//AHF
#define WA_SIZE        33

#define RDD_CONNECTION 1
#define RDD_CATALOG    2

#define RDD_SIZE       2

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
 
   ADOCONNECT(nWA,aOpenInfo)
   
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

   oRecordSet:CursorType :=  IF(aWAData[ WA_ENGINE ] = "ACCESS", adOpenKeyset, adOpenDynamic) // adOpenKeyset adOpenDynamic
   oRecordSet:CursorLocation := IF(aWAData[ WA_ENGINE ] = "ACCESS", adUseServer, adUseClient) //adUseServer  // adUseClient its slower but has avntages such always bookmaks 
   oRecordSet:LockType :=    IF(aWAData[ WA_ENGINE ] = "ACCESS", adLockOptimistic, adLockOptimistic) //adLockOptimistic adLockPessimistic

   IF aOpenInfo[UR_OI_READONLY]
      oRecordSet:LockType := adLockReadOnly
   ELSE
      oRecordSet:LockType :=  IF(aWAData[ WA_ENGINE ] = "ACCESS", adLockOptimistic, adLockPessimistic) //adLockPessimistic //adLockOptimistic
   ENDIF	  

   //PROPERIES AFFECTING PERFORMANCE TRY
   //oRecordSet:MaxRecords := 60
   //oRecordSet:CacheSize := 50 //records increase performance set zero returns error set great server parameters max open rows error
   //oRecordset:PageSize = 10
   //oRecordSet:MaxRecords := 15
   //oRecordset:Properties("Maximum Open Rows") := 110  //MIN TWICE THE SIZE OF CACHESIZE

   IF aWAData[ WA_QUERY ] == "SELECT * FROM "
      oRecordSet:Open( aWAData[ WA_QUERY ] + aWAData[ WA_TABLENAME ], aWAData[ WA_CONNECTION ])
   ELSE
      oRecordSet:Open( aWAData[ WA_QUERY ], aWAData[ WA_CONNECTION ],,,adCmdTableDirect )
   ENDIF

   aWAData[ WA_RECORDSET ] := oRecordSet
   aWAData[ WA_BOF ] := aWAData[ WA_EOF ] := .F.
 
   UR_SUPER_SETFIELDEXTENT( nWA, nTotalFields := oRecordSet:Fields:Count )
   
   //FIELD TO BE USED AS RECNO FO THIS TABLE 
   cName := ADO_GET_FIELD_RECNO(  aWAData[ WA_TABLENAME ] )
   
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
      IF UPPER(cName) = UPPER(oRecordSet:Fields( n - 1 ):Name)
	  
         aWAData[WA_FIELDRECNO]:=  n - 1 
		 //IF IT SUPPORTS SEEK WE WILL SEEK IT ISNTEAD OF FIND IT
		 IF !oRecordSet:Supports(adIndex) .OR. !oRecordSet:Supports(adSeek)
		    //OTHERWISE LETS USE ADO INDEX PROP TO SPEED UP
            IF  oRecordSet:CursorLocation = adUseClient	 
			    oRecordSet:Fields( aWAData[WA_FIELDRECNO] ):Properties("Optimize") := 1 
		    ENDIF
		 ENDIF	
		 
	  ENDIF
	  
      UR_SUPER_ADDFIELD( nWA, aField )
   NEXT

   nResult := UR_SUPER_OPEN( nWA, aOpenInfo )
  
   IF nResult == HB_SUCCESS
      ADO_GOTOP( nWA )
   ENDIF
   
   //auto open set and auto order
   IF SET(_SET_AUTOPEN)
   
      ADO_INDEXAUTOOPEN(aWAData[ WA_TABLENAME ])
	  
   ENDIF
   
   
   RETURN nResult
   
   
STATIC FUNCTION ADOCONNECT(nWA,aOpenInfo)

 LOCAL aWAData := USRRDD_AREADATA( nWA )
 LOCAL aDefaults := ADODEFAULTS() //get defaults set or the sets called wth USE

TRY
  IF Empty( aOpenInfo[ UR_OI_CONNECT ] )
   
     IF EMPTY(oConnection)

	    aWAData[ WA_CONNECTION ] :=  TOleAuto():New( "ADODB.Connection" )
	    oConnection := aWAData[ WA_CONNECTION ]
		 
		aWAData[ WA_TABLENAME ]  := UPPER(aOpenInfo[ UR_OI_NAME ])
		aWAData[ WA_QUERY ]      := t_cQuery
		aWAData[ WA_USERNAME ]   := t_cUserName
		aWAData[ WA_PASSWORD ]   := t_cPassword
		aWAData[ WA_SERVER ]     := t_cServer
		aWAData[ WA_ENGINE ]     := t_cEngine
		aWAData[ WA_CONNOPEN ]   := .T.
    
		DO CASE
		  
		  CASE Lower( Right( aOpenInfo[ UR_OI_NAME ], 4 ) ) == ".mdb" .OR. aWAData[ WA_ENGINE ] = "ACCESS"
		  
			 IF Empty( aWAData[ WA_PASSWORD ] )
				aWAData[ WA_CONNECTION ]:Open( "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + t_cDataSource  )
			 ELSE
				aWAData[ WA_CONNECTION ]:Open( "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + t_cDataSource  + ";Jet OLEDB:Database Password=" + AllTrim( aWAData[ WA_PASSWORD ] ) )
			 ENDIF

		  CASE Lower( Right( aOpenInfo[ UR_OI_NAME ], 4 ) ) == ".xls" 
		  
			 aWAData[ WA_CONNECTION ]:Open( "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + t_cDataSource  + ";Extended Properties='Excel 8.0;HDR=YES';Persist Security Info=False" )

		  CASE Lower( Right( aOpenInfo[ UR_OI_NAME ], 4 ) ) == ".dbf" .OR. aWAData[ WA_ENGINE ] = "ADS"
		  
			 aWAData[ WA_CONNECTION ]:Open("Provider=Advantage OLE DB Provider;User ID="+aWAData[ WA_USERNAME ] +;
					 ";Data Source="+ t_cDataSource +";TableType=ADS_VFP;"+;
					 "Advantage Server Type=ADS_LOCAL_SERVER;")
			 
			 aWAData[ WA_TABLENAME ] := CFILENOEXT(CFILENOPATH(aWAData[ WA_TABLENAME ] ))
             IF	UPPER( SUBSTR( aWAData[ WA_TABLENAME ],1,3) ) = "TMP" .OR. 	UPPER( SUBSTR( aWAData[ WA_TABLENAME ],1,3) ) = "TEMP"
			    aWAData[ WA_TABLENAME ] := "#"+aWAData[ WA_TABLENAME ]
             ENDIF
			 
			 //OTHE PROVIDERS FOR DBF FOR TRIALS		
			 // aWAData[ WA_CONNECTION ]:Open("Provider=vfpoledb.1;Data Source="+t_cDataSource+";Collating Sequence=machine;")
			 // aWAData[ WA_CONNECTION ]:Open( "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + t_cDataSource + ";Extended Properties=dBASE IV;User ID=Admin;Password=;" )

		  CASE Lower( Right( aOpenInfo[ UR_OI_NAME ], 3 ) ) == ".db"
		  
			 aWAData[ WA_CONNECTION ]:Open( "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + t_cDataSource + ";Extended Properties='Paradox 3.x';" )
			 
		  CASE aWAData[ WA_ENGINE ] == "MYSQL"

			 aWAData[ WA_CONNECTION ]:Open( "Driver={mySQL ODBC 5.3 ANSI Driver};" + ;
				"server=" + aWAData[ WA_SERVER ] + ;
				";Port=3306;Option=4"+;
				";database=" + t_cDataSource  + ;
				";uid=" + aWAData[ WA_USERNAME ] + ;
				";pwd=" + aWAData[ WA_PASSWORD ]+";" )
				
		  CASE aWAData[ WA_ENGINE ] == "SQL"
		  
			 aWAData[ WA_CONNECTION ]:Open( "Provider=SQLOLEDB;" + ;
				"server=" + aWAData[ WA_SERVER ] + ;
				";database=" + t_cDataSource  + ;
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
				";DbName=" + t_cDataSource  )
			 
		  ENDCASE		 
		 
	 ELSE

	    // ITS ALREDY OPEN THE ADODB CONN USE THE SAME WE WANT TRANSACTIONS WITHIN THE CONNECTION
		aWAData[ WA_CONNECTION ] :=  oConnection
	    aWAData[ WA_TABLENAME ]  := aOpenInfo[ UR_OI_NAME ]
		aWAData[ WA_QUERY ]      := t_cQuery
		aWAData[ WA_USERNAME ]   := t_cUserName
		aWAData[ WA_PASSWORD ]   := t_cPassword
		aWAData[ WA_SERVER ]     := t_cServer
		aWAData[ WA_ENGINE ]     := t_cEngine
		aWAData[ WA_CONNOPEN ]   := .T.
		 
     ENDIF	  
	  
  ELSE

     // here we dont save oconnection for the next one because
	 // we assume that is not application defult conn but a temporary conn
	 //to other db system.
     aWAData[ WA_CONNECTION ] := TOleAuto():New("ADODB.Connection")
	 aWAData[ WA_CONNECTION ]:Open( aOpenInfo[ UR_OI_CONNECT ] )
	 aWAData[ WA_TABLENAME ]  := aOpenInfo[ UR_OI_NAME ]
	 aWAData[ WA_QUERY ]      := t_cQuery
	 aWAData[ WA_USERNAME ]   := t_cUserName
	 aWAData[ WA_PASSWORD ]   := t_cPassword
	 aWAData[ WA_SERVER ]     := t_cServer
	 aWAData[ WA_ENGINE ]     := t_cEngine
	 aWAData[ WA_CONNOPEN ]   := .F.
	  
  ENDIF

  aWAData[ WA_TABLENAME ] := CFILENOEXT(CFILENOPATH(aWAData[ WA_TABLENAME ] ))
  IF aWAData[ WA_ENGINE ] = "ADS"
     //IF TEMP TABLE SEND IT AS TO PROVIDER
     IF UPPER( SUBSTR( aWAData[ WA_TABLENAME ],1,3) ) = "TMP" .OR. 	UPPER( SUBSTR( aWAData[ WA_TABLENAME ],1,4) ) = "TEMP"
	    aWAData[ WA_TABLENAME ] := "#"+aWAData[ WA_TABLENAME ]
	 ENDIF	
  ENDIF
CATCH
	ADOSHOWERROR(aWAData[ WA_CONNECTION ])
END
   //WE DONT NEED IT ANYMORE REINITIALIZE
  t_cTableName := NIL
  t_cEngine    := NIL
  t_cServer    := NIL
  t_cUserName  := NIL
  t_cPassword  := NIL
  t_cQuery     := NIL
 
  RETURN HB_SUCCESS

   
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
   ADO_UNLOCK( nWA ) //RELEASE ALL LOCKS IN TLOCKS.DBF
   
   RETURN UR_SUPER_CLOSE( nWA )
   
  
/*                              RECORD RELATED FUNCTION                   */

STATIC FUNCTION ADO_GET_FIELD_RECNO( cTablename )

  LOCAL cFieldName := ADODEFLDRECNO() //default recno field name
  LOCAL aFiles :=  ListFieldRecno(),n

   IF !EMPTY(aFiles) //IS THERE A FIELD AS RECNO DIFERENT FOR THIS TABLE
   
      n := ASCAN( aFiles, { |z| z[1] == cTablename } )
	  
	  IF n > 0
	     cFieldName := aFiles[n,2]
	  ENDIF
	  
   ENDIF
   
   RETURN cFieldName
   

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

   
STATIC FUNCTION ADO_RECNO( nWA, nRecNo )
   
   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]
   LOCAL nResult := HB_SUCCESS,nRecords 
   
   
    IF !EMPTY(aWAData[WA_FIELDRECNO])  // 100% SUPPORTED AND SAFE

	   IF !oRecordSet:Eof()
  
	      ADO_GETVALUE( nWA, aWAData[WA_FIELDRECNO]+1, @nRecNo ) 
 		  
	   ELSE

           nRecNo := ADORECCOUNT(nWA,oRecordSet)+1 //  ADO_RECCOUNT( nWA, @nRecNo )
		  
	   ENDIF
	   
    ELSE
	
	   IF oRecordSet:Supports(adBookmark) 
	
	      /* Although the Supports method may return True for a given functionality, it does not guarantee that 
	      the provider can make the feature available under all circumstances. 
	      The Supports method simply returns whether the provider can support the specified functionality,
	      assuming certain conditions are met. For example, the Supports method may indicate that a 
	      Recordset object supports updates even though the cursor is based on a multiple table join, 
	      some columns of which are not updatable*/
	      IF oRecordSet:Eof() .or. oRecordSet:Bof()
	   
	         nRecno := 0
		  
	      ELSE	  
		  
              nRecno := oRecordSet:BookMark
			  
		  ENDIF
		  
       ELSE		
	   
          //ATTENTION NOT WORKING CORRECTLY WITH DELETED ROWS!2	
          nRecno := IF( oRecordSet:AbsolutePosition == adPosEOF, oRecordSet:RecordCount() + 1, oRecordSet:AbsolutePosition )
	      //MUST TAKE OUT THE DELETED ROWS! OTHERWISE WRONG NRECNO 
		  //TODO nRecno := nRecno-nDeletedRows 
		  
       ENDIF	
	   
	ENDIF
	
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
    
   IF ADOEMPTYSET(oRecordSet)	
      RETURN nCount
   ENDIF	  
   //Making it lightning faster
   
   oRS:CursorLocation := IF(aAWData[ WA_ENGINE ] = "ACCESS", adUseClient, adUseServer) //adUseServer  // adUseClient its slower but has avntages such always bookmaks 
   oRs:CursorType := adOpenForwardOnly
   oRs:LockType := adLockReadOnly

   IF !EMPTY(aAWData[WA_FIELDRECNO])  //RECCOUNT/LASTREC = MAX NUMBER OF FIELD RECNO
   
      cSql := "SELECT MAX("+(ADO_GET_FIELD_RECNO( aAWData[WA_TABLENAME] ))+") FROM "+aAWData[WA_TABLENAME]
	  
   ELSE	//NO FIELD RECNO RECCOUNT/LASTREC = NR OF ROWS
   
      //LAST PARAMTER INSERTS cSql COUNT(*) MUST BE ALL FIELDS BECAUSE IF THERE IS A NULL FIELD COUNTS RETURNS WRONG
      cSql := "SELECT COUNT(*) FROM "+aAWData[WA_TABLENAME]
	  
   ENDIF	  
   
   //LETS COUNT IT
   oRs:open(cSql,oCon)
   nCount := oRs:Fields( 0 ):Value
   oRs:close()
    
   RETURN nCount   

   
STATIC FUNCTION ADO_GOTO( nWA, nRecord )
   
   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL nRecNo
   LOCAL oRecordSet := aWAData[ WA_RECORDSET ]
   LOCAL nRecs := ADORECCOUNT(nWa,oRecordSet) // AHF SEE FUNCTION FOR EXPLANATION rs:RecordCount <> 0
   
   IF !ADOEMPTYSET(oRecordSet) //.AND. nRecord < nRecs  .AND. nRecord > 0

	  IF !EMPTY(aWAData[WA_FIELDRECNO]) 
	  
         IF oRecordSet:Supports(adIndex) .AND. oRecordSet:Supports(adSeek)
		    oRecordSet:Index := oRecordSet:Fields(aWAData[WA_FIELDRECNO]):Name
			oRecordSet:Seek({ ALLTRIM(STR(nRecord,0)) })
		 ELSE
            oRecordSet:MoveFirst()
		    oRecordSet:Find(oRecordSet:Fields(aWAData[WA_FIELDRECNO]):Name+" = "+ALLTRIM(STR(nRecord,0)) )
		 ENDIF	
         // IF EOF RAISE ERROR
	  ELSE
	  
	     IF oRecordSet:Supports(adBookmark)

     		//WORKAROUND IT GETS HERE AS INTEGER WITHOUT DECIMALS
			//ATTENTION ITS A VARIANT TYPE CA BE ANY VALUE
	        nRecord := VAL(CVALTOCHAR(nRecord)+".00")
			
            oRecordSet:BookMark := nRecord //READ NOTES IN ADO_RECNO
		 
	     ELSE	 

		     oRecordSet:AbsolutePosition := Max( 1, Min( nRecord, oRecordSet:RecordCount() ) )
	  
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
   
   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL nRecNo
   LOCAL oRecordSet := aWAData[ WA_RECORDSET ]
   LOCAL nRecs := ADORECCOUNT(nWa,oRecordSet) // AHF SEE FUNCTION FOR EXPLANATION rs:RecordCount <> 0
   
   IF !ADOEMPTYSET(oRecordSet) //.AND. nRecord < nRecs .AND. nRecord > 0

	  IF !EMPTY(aWAData[WA_FIELDRECNO]) 
	  
         IF oRecordSet:Supports(adIndex) .AND. oRecordSet:Supports(adSeek)
		    oRecordSet:Index := oRecordSet:Fields(aWAData[WA_FIELDRECNO]):Name
			oRecordSet:Seek({ ALLTRIM(STR(nRecord,0)) })
		 ELSE
            oRecordSet:MoveFirst()
		    oRecordSet:Find(oRecordSet:Fields(aWAData[WA_FIELDRECNO]):Name+" = "+ALLTRIM(STR(nRecord,0)) )
	     ENDIF
		 
	  ELSE
	     IF oRecordSet:Supports(adBookmark)
		 
		    //WORKAROUND IT GETS HERE AS INTEGER WITHOUT DECIMALS
			//ATTENTION ITS A VARIANT TYPE CA BE ANY VALUE
	        nRecord := VAL(CVALTOCHAR(nRecord)+".00")

            oRecordSet:BookMark := nRecord //READ NOTES IN ADO_RECNO

 		 ELSE	
		 
		    oRecordSet:AbsolutePosition := Max( 1, Min( nRecord, oRecordSet:RecordCount() ) )
	   
		 ENDIF	
     
	  ENDIF
	  
	  ADO_RECID( nWA, @nRecord )  
	  
      IF !EMPTY(aWAData[WA_PENDINGREL]) .AND. PROCNAME(2) <> "ADO_RELEVAL" //ENFORCE REL CHILDS BUT NOT IN A ENDLESS LOOP!
         ADO_FORCEREL( nWA ) 
      ENDIF

   ENDIF
   
   aWAData[ WA_EOF ] := oRecordSet:Eof()
   aWAData[ WA_BOF ] := oRecordSet:Bof()
   

   RETURN HB_SUCCESS 

   
STATIC FUNCTION ADO_GOTOP( nWA )

   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := aWAData[ WA_RECORDSET ]

   IF !ADOEMPTYSET(oRecordSet) 
      oRecordSet:MoveFirst()
   ENDIF

   aWAData[ WA_EOF ] := oRecordSet:Eof()
   //CANT DO THIS SKIPRAW WROSK WRONG aWAData[ WA_BOF ] := oRecordSet:Bof()

   RETURN HB_SUCCESS

   
STATIC FUNCTION ADO_GOBOTTOM( nWA )

   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := aWAData[ WA_RECORDSET ]

   IF !ADOEMPTYSET(oRecordSet) 
      oRecordSet:MoveLast()
   ENDIF	  

   aWAData[ WA_EOF ] := oRecordSet:Eof()
   //CANT DO THIS SKIPRAW WROSK WRONG aWAData[ WA_BOF ] := oRecordSet:Bof()

   RETURN HB_SUCCESS
   

STATIC FUNCTION ADO_SKIPRAW( nWA, nToSkip )

   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := aWAData[ WA_RECORDSET ]
   LOCAL nResult := HB_SUCCESS
   LOCAL nRecords := ADORECCOUNT(nWa,oRecordSet)// AHF SEE FUNCTION FOR EXPLANATION oRecordSet:RecordCount = 0
   
   IF nRecords = 0 
      nToSkip := 0
      RETURN HB_FAILURE
   ENDIF
   
   IF ADOEMPTYSET(oRecordSet)  
      nTotSkip := 0
	  RETURN HB_SUCCESS //SHOULDNET BE FAILURE?
   ENDIF
   
   IF nToSkip != 0
   
      IF aWAData[ WA_EOF ]
         IF nToSkip > 0
            RETURN HB_SUCCESS //SHOULDNET BE FAILURE?
         ENDIF
         ADO_GOBOTTOM( nWA )
         ++nToSkip
      ENDIF
	  
      IF nToSkip < 0 .AND. oRecordSet:AbsolutePosition <= - nToSkip
         oRecordSet:MoveFirst()
         aWAData[ WA_BOF ] := .T.
         aWAData[ WA_EOF ] := oRecordSet:EOF
      ELSE
         oRecordSet:Move( nToSkip )
         aWAData[ WA_BOF ] := .F.
         aWAData[ WA_EOF ] := oRecordSet:EOF
      ENDIF
	  
      //ENFORCE RELATIONS SHOULD BE BELOW AFTER MOVING TO NEXT RECORD
      IF ! Empty( aWAData[ WA_PENDINGREL ] )
         ADO_FORCEREL( nWA ) 
      ENDIF
	  
   ELSE
      //SKIP 0 SHOULD FLUSH ALL DATA ACC TO XHARBOUR USER MANUAL
   ENDIF

   RETURN nResult
   

STATIC FUNCTION ADO_BOF( nWA, lBof )

   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]
   
   lBof := aWAData[ WA_BOF ]

   RETURN HB_SUCCESS
   

STATIC FUNCTION ADO_EOF( nWA, lEof )

   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]
   LOCAL nResult := HB_SUCCESS

    lEof := oRecordSet:Eof() 
   

   RETURN nResult


STATIC FUNCTION ADO_APPEND( nWA, lUnLockAll )

   LOCAL oRs := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]
   LOCAL aWdata := USRRDD_AREADATA( nWA )
   LOCAL aStruct, cBaseTable, n, oField, u
   LOCAL aCols := {}, aVals := {}
   LOCAL lAdded   := .f.,nRecNo
   LOCAL aLockInfo := ARRAY(UR_LI_SIZE),oError
   
	aStruct := ADOSTRUCT( oRs )
	FOR n := 1 TO LEN( aStruct )
        IF aStruct[ n, 6 ]
            AADD( aCols, aStruct[ n, 1 ] )
            AADD( aVals, HB_DECODE( aStruct[ n, 2 ], 'C', Space( aStruct[ n, 3 ] ), 'D', CToD( '' ), 'L', .f., ;
                 'M', "", 'm', "", '+', 0, ;
                 'N', If( aStruct[ n, 3 ] == 0, 0, Val( "0." + Replicate( '0', aStruct[ n, 3 ] ) ) ), ;
                 'T', CTOT( '' ), '' ) )
				 
			//hb_DeCode( aStruct[ n, 2 ], 'C', SPACE( aStruct[ n, 3 ] ), ;
            //   'D', CTOD("  /  /  "), 'L', .f., 'M', SPACE(255), 'N', 0, 'T', "", '' ) )
         ENDIF
    NEXT
	
    IF ! EMPTY( aCols )

       TRY
	  
         aLockInfo[ UR_LI_RECORD ] := ADORECCOUNT(nWA,oRs)+1 //GHOST NEXT RECORD
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
            oError:FileName := aWAData[ WA_TABLENAME]
            oError:OsCode := 0 /* TODO */
            oError:CanDefault := .T.
			UR_SUPER_ERROR( nWA, oError )
            BREAK
		 ENDIF
		 
         oRs:AddNew( aCols, aVals )
		 oRs:Update()
		 
		 aWData[ WA_EOF ] := oRs:Eof()
		 aWData[ WA_BOF ] := oRs:Bof()
		 
         lAdded   := .t.
		 NETERR(.F.)	
	 
		 ADO_RECID( nWA, @nRecNo )
		 
    CATCH
	
         NETERR(.T.)
		 ADOSHOWERROR(aWdata[WA_CONNECTION])
    END
	  
   ELSE
   
      NETERR(.T.)
	  
   ENDIF
   
   IF !EMPTY(aWData[WA_PENDINGREL]) .AND. PROCNAME(2) <> "ADO_RELEVAL" //ENFORCE REL CHILDS BUT NOT IN A ENDLESS LOOP!
      ADO_FORCEREL( nWA ) 
   ENDIF

   RETURN IF(lAdded,HB_SUCCESS,HB_FAILURE)

//APPEND FROM AND COPY TO   
STATIC FUNCTION ADO_TRANS(  nDstArea, aFieldsStru, bFor, bWhile, nNext, nRecord, lRest )

msginfo("dest reea " +cvaltochar(ndstarea)+;
"dest bfor " +cvaltochar(bfor) +;
"dest bwhile " +cvaltochar(bwhile) +;
"dest nnext " +cvaltochar(nnext) +;
"dest nrecord " +cvaltochar(nRecord) +;
"dest lrest " +cvaltochar(lRest) )

MSGSELECT(ARRAYTOCHAR(aFieldsStru))

  RETURN HB_SUCCESS


STATIC FUNCTION ADO_TRANSREC( aTransItem )
MSGSELECT(ARRAYTOCHAR(aTransItem))
  RETURN HB_SUCCESS
  

/*                           END RECORD RELATED FUNCTION                   */   

/*                                      DELETE RECALL ZAP PACK                  */
STATIC FUNCTION ADO_DELETED( nWA, lDeleted )

   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]

    IF !ADOEMPTYSET(oRecordSet)
	   IF !oRecordSet:Eof .AND. !oRecordSet:Bof
          IF oRecordSet:Status = adRecDeleted 
             lDeleted := .T.
          ELSE
             lDeleted := .F.
          ENDIF
	   ELSE
 	      lDeleted := .F.
	   ENDIF	  
	ELSE
       lDeleted := .F.	
    ENDIF
	
   RETURN HB_SUCCESS
   

STATIC FUNCTION ADO_DELETE( nWA )

   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]
   LOCAL tmp, lDeleted := .F.,nRecNo

   IF ADORECCOUNT(nWA,oRecordSet) > 0 // AHF SEE FUNCTION FOR EXPLANATION oRecordSet:RecordCount()
   
      IF !oRecordSet:Eof .AND. !oRecordSet:Bof
	  
 		 ADO_RECID(nWa,@nRecNo)
		 IF ADO_ISLOCKED(aWAData[ WA_TABLENAME],nRecNo, aWAData) 
		    tmp = oRecordSet:AbsolutePosition //SAME USED IN ADOFUNCS
            oRecordSet:Delete()
		    oRecordSet:Update()
            oRecordSet:AbsolutePosition := Max( 1, Min( tmp, oRecordSet:RecordCount() ) ) //SAME USED IN ADOFUNCS
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
	  
   RETURN IF(lDeleted,HB_SUCCESS,HB_FAILURE)
   

STATIC FUNCTION ADO_RECALL(nRecno)
   MSGALERT("RECALL NOT POSSIBLE IN SQL!")
   RETURN HB_SUCCESS
   
   
STATIC FUNCTION ADO_ZAP( nWA )

   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := aWAData[ WA_RECORDSET ]
//AUTOINC FIELDS LIKE RECNO MUST BE RESET
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

   //DOES NOTHING BECAUSE RECORDS ARE AUTOMATICLY REMOVED WHEN DELETED
   HB_SYMBOL_UNUSED( nWA )
   
   RETURN HB_SUCCESS
/*                             END OF DELETE RECALL ZAP PACK              */


/*                               FIELD RELATED FUNCTIONS  */
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
         xValue := CTOD('') 
      ENDIF
	  
	  IF aFieldInfo[7] == HB_FT_INTEGER .OR.  aFieldInfo[7] ==  HB_FT_DOUBLE
	  
		 IF aFieldInfo[4] > 0
            xValue := VAL("0."+REPLICATE("0",VAL(STR(aFieldInfo[4],0)) )) //VAL("0."+ALLTRIM(STR(aFieldInfo[4],0)))
		 ELSE
		    xValue := INT(0)
         ENDIF	
		 
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
            xValue := FW_TTOD('') 
	     ELSE
 		    xValue := FW_TTOD( xValue ) 
         ENDIF
		 
	  ENDIF

	  IF aFieldInfo[7] == HB_FT_INTEGER .OR.  aFieldInfo[7] ==  HB_FT_DOUBLE 
	  
	     IF VALTYPE( xValue ) == "U"
		 
		    IF aFieldInfo[4] > 0
               xValue := VAL("0."+REPLICATE("0",VAL(STR(aFieldInfo[4],0)) )) // ALLTRIM(STR(aFieldInfo[4],0)))
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
            xValue := FW_DTOT(xValue)		 
         ENDIF
		 
	  ENDIF
	  
   ENDIF	  
   
   
   RETURN HB_SUCCESS

   
STATIC FUNCTION ADO_PUTVALUE( nWA, nField, xValue )

   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := aWAData[ WA_RECORDSET ]
   LOCAL nRecNo, aOrderInfo  := ARRAY(UR_ORI_SIZE),oError
   

   IF ! aWAData[ WA_EOF ] .AND. !( oRecordSet:Fields( nField - 1 ):Value == xValue )
 
      //CHECK IF THE FIELD CA BE RW
	  IF ADO_FIELDSTRUCT( oRecordSet, nField-1 )[6]
	  
	     IF ADO_FIELDSTRUCT( oRecordSet, nField-1 )[2] = "N"
		 
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

		 ADO_RECID(nWa,@nRecNo)
		 IF ADO_ISLOCKED(aWAData[ WA_TABLENAME],nRecNo,aWAData) 

            oRecordSet:Fields( nField - 1 ):Value := xValue
	        oRecordSet:Update()

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
		 
		 // ONLY TO DO IF FIELD IN ORDER BY
		 ADO_ORDINFO( nWA, DBOI_EXPRESSION, aOrderInfo )

		 IF ALLTRIM(oRecordSet:Fields( nField - 1 ):Name) $  aOrderInfo[ UR_ORI_RESULT]

		    ADO_RECID(nWa,@nRecNo)
		    oRecordSet:Requery()
		    ADO_GOTO(nWA,nRecNo)
			
         ENDIF	 
		 
      ELSE
	  
	     RETURN HB_FAILURE
		 
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
#ifdef HB_FT_DATETIME
      CASE nType == HB_FT_DATETIME
         uInfo := "T"
#endif
      CASE nType == HB_FT_TIMESTAMP
         uInfo := "@"
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
      IF oRs != nil .AND. ! oRs:Eof() .AND. VALTYPE( uVal := oField:Value ) == 'T' 
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

      IF oField:Properties( "ISAUTOINCREMENT" ):Value == .t.
         cType := '+'
         lRW   := .f.
		 nDBFFieldType := HB_FT_AUTOINC
      ENDIF
	  
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
      nDBFFieldType := HB_FT_DATETIME
	  
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

   FOR n := 1 TO oRs:Fields:Count()
      AADD( aStruct, ADO_FIELDSTRUCT( oRs, n-1 ) )
   NEXT

RETURN aStruct   
/*                          END FIELD RELATED FUNCTIONS  */


/*                                 INDEX RELATED FUNCTIONS  */
STATIC FUNCTION ADO_INDEXAUTOOPEN(cTableName)

  LOCAL aFiles := ListIndex(),y,z,nOrder := 0
  
  //TEMPORARY INDEXES NOT ICLUDED HERE
  //NORMALY ITS CREATED A THEN OPEN?
  
    y:=ASCAN( aFiles, { |z| z[1] == cTablename } )
	
    IF y >0
	
	   nMax := LEN(aFiles[y])-1
	   
	   FOR z :=1 TO LEN( aFiles[y]) -1
	   
  	       ORDLISTADD( aFiles[y,z+1,1] )
		   
	   NEXT
	   
	ENDIF
	
	IF SET(_SET_AUTORDER) > 1
	   // IN ADO_ORDLSTADD DEFAULTS TO SET FOCUS ORDER 1 IF DIFERENT ASSUMES THIS
	   SET ORDER TO SET(_SET_AUTORDER)
	ENDIF
	
   RETURN HB_SUCCESS
   

STATIC FUNCTION ADO_ORDINFO( nWA, nIndex, aOrderInfo )

   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := aWAData[ WA_RECORDSET ]
   LOCAL nResult := HB_SUCCESS
   LOCAL cExp:="",cIndexExp := "",ERR
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
		      aOrderInfo[ UR_ORI_RESULT ] := STRTRAN(aOrderInfo[ UR_ORI_RESULT ] , ",","+")
		   
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
   
        aOrderInfo[ UR_ORI_RESULT ] :=.F. //ITS REALLY NEVER USED
   
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
		   aOrderInfo[ UR_ORI_RESULT ] := oRecordSet:AbsolutePosition()+1
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
           aOrderInfo[ UR_ORI_RESULT ] := oRecordSet:RecordCount() //ADORECCOUNT(nWA,oRecordSet) // AHF SEE FUNCTION FOR EXPLANATION oRecordSet:RecordCount
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
  
  
STATIC FUNCTION ADO_ORDLSTFOCUS( nWA, aOrderInfo )

   LOCAL nRecNo 
   LOCAL aWAData    := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := aWAData[ WA_RECORDSET ]
   LOCAL cSql:="" ,n
   
   HB_SYMBOL_UNUSED( nWA )
   HB_SYMBOL_UNUSED( aOrderInfo )
   
   IF !EMPTY(aWAData[WA_FIELDRECNO]) 
      ADO_RECID(nWA,@nRecno)
   ENDIF
   
   IF aOrderInfo[ UR_ORI_TAG ] <> NIL
      //TRY
      oRecordSet:Close()
	  //CATCH
	   // ADOSHOWERROR(aWAData[ WA_CONNECTION ])
	  // END	
	  /* AHF NOT NEEDED ONLY IF YOU WANT TO CHANGE IT OTHERWISE STAYS AS IT WAS WHEN OPENING IT
      oRecordSet:CursorType := adOpenDynamic
      oRecordSet:CursorLocation := adUseServer //adUseClient never use ths very slow!
      oRecordSet:LockType := adLockPessimistic
	  */
      IF VALTYPE(aOrderInfo[ UR_ORI_TAG ]) = "C"

         //MAYBE IT COMES WITH FILE EXTENSION AND PATH
	     aOrderInfo[ UR_ORI_TAG ] := CFILENOPATH(aOrderInfo[UR_ORI_TAG])
	     aOrderInfo[ UR_ORI_TAG ] := UPPER(CFILENOEXT(aOrderInfo[ UR_ORI_TAG ]))
	  
         n := ASCAN(aWAData[ WA_INDEXES ],UPPER(aOrderInfo[ UR_ORI_TAG ]))
	  ELSE
         n := aOrderInfo[ UR_ORI_TAG ]  
	  ENDIF   

      IF n = 0  //PHISICAL ORDER

		 aWAData[ WA_INDEXACTIVE ] := 0
		 aOrderInfo[ UR_ORI_RESULT ] := ""
	     cSql := IndexBuildExp(nWA,n,aWAData)

         oRecordSet:Open( cSql, aWAData[ WA_CONNECTION ])
		 
      ELSE

	     IF aWAData[ WA_INDEXACTIVE ] > 0
            aOrderInfo[ UR_ORI_RESULT ] := aWAData[ WA_INDEXES ] [aWAData[ WA_INDEXACTIVE ]]
		 ELSE
		    aOrderInfo[ UR_ORI_RESULT ] := ""
         ENDIF		 
		 
		 aWAData[ WA_INDEXACTIVE ] := n
	     cSql := IndexBuildExp(nWA,n,aWAData)
		 
		 oRecordSet:Open( cSql,aWAData[ WA_CONNECTION ])

      ENDIF
	  
	  IF !EMPTY(aWAData[WA_FIELDRECNO]) .AND. VALTYPE(nRecNo) = "N"
	     IF PROCNAME(1) = "ADO_ORDLSTADD" //SET  INDEX TO THEN GO FRIST RECORD
		    ADO_GOTOP( nWA )
		 ELSE
	        ADO_GOTO( nWA, nRecNo )
	     ENDIF		
	  ELSE	 
         ADO_GOTOP( nWA )
      ENDIF

	  aWAData[WA_ISITSUBSET] := .F.
	  
 	  ADO_SETFILTER( nWA, aWAData[ WA_FILTERACTIVE ] ) //ENFORCE ANY ACIVE FILTER
	  
	ELSE

	   IF aWAData[ WA_INDEXACTIVE ] > 0
          aOrderInfo[ UR_ORI_RESULT ] := aWAData[ WA_INDEXES ] [aWAData[ WA_INDEXACTIVE ]]
	   ELSE
		  aOrderInfo[ UR_ORI_RESULT ] := ""
       ENDIF		 
	   
	ENDIF  


   RETURN HB_SUCCESS
   

STATIC FUNCTION ADO_ORDLSTADD( nWA, aOrderInfo )
   LOCAL cTablename := USRRDD_AREADATA( nWA )[ WA_TABLENAME ]
   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL aFiles := ListIndex() 
   LOCAL aTempFiles := ListTmpNames() 
   LOCAL cExpress := "" ,cFor:="",cUnique:="",y,z,x
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
	   
       AADD( aWAData[WA_INDEXFOR],IF(!EMPTY(aTmpFor[y]),"WHERE ","")+aTmpFor[y])
       AADD( aWAData[WA_INDEXUNIQUE],aTmpUnique[y])

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
			  EXIT
		   ENDIF	  
	   NEXT
	ELSE
       nMax := 1	
	ENDIF

	IF EMPTY(cIndex) //maybe should generate error
	   RETURN HB_FAILURE
    ENDIF

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
   LOCAL aTmpUnique := ListTmpUnique() 
   LOCAL aTmpDbfExp := ListTmpDbfExp()
   LOCAL aTmpDbfFor := ListTmpDbfFor()
   LOCAL aTmpDbfUnique := ListTmpDbfUnique()
   LOCAL cForExp := ""
   LOCAL cFile := cIndex
   
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
	AADD(aTmpExp,ADOINDEXFIELDS(nWA,UPPER(STRTRAN(aOrderCreateInfo[UR_ORCR_CKEY],"+",","))) )
	AADD(aTmpDbfExp,aOrderCreateInfo[UR_ORCR_CKEY])
    
	IF !EMPTY(acondinfo) 
		cForExp := IF(!EMPTY(acondinfo[UR_ORC_CWHILE]),acondinfo[UR_ORC_CWHILE],"") +IF(!EMPTY(acondinfo[UR_ORC_CFOR]), " AND "+acondinfo[UR_ORC_CFOR],"")

		AADD(aTmpFor,UPPER(STRTRAN(cForExp,'"',"'")) )//CLEAN THE DOT .AND. .OR.

		AADD(aTmpDbfFor,cForExp)
		IF acondinfo[UR_ORC_NEXT] > 0
		   AADD(aTmpUnique,UPPER(acondinfo[UR_ORC_NEXT] ))
		   AADD(aTmpDbfUnique,UPPER(acondinfo[UR_ORC_NEXT] ))
		ELSE
		   AADD(aTmpUnique,"")
		   AADD(aTmpDbfUnique,"")
		ENDIF
    ELSE
	   AADD(aTmpFor,"")
	   AADD(aTmpDbfFor,"")
	   AADD(aTmpUnique,"")
       AADD(aTmpDbfUnique,"")
	ENDIF
	
    aOrderInfo [UR_ORI_BAG ] := cIndex
    aOrderInfo [UR_ORI_TAG ] := cIndex
	

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

   
STATIC FUNCTION ADOINDEXFIELDS(nWA,cExpression)
  LOCAL n,nAt, aFields :={}, cStr := ""
  
    cExpression := UPPER(cExpression)
	
    FOR n := 1 to (nWA)->(FCOUNT()) // we have to check all fields in table because there isnt any conspicuous mark on the expression to guide us
	    
	    nAt := AT(ALLTRIM((nWA)->(FIELDNAME(n))),cExpression)
		
	    IF nAt > 0

   		   AADD(aFields ,{ALLTRIM((nWA)->(FIELDNAME(n))),nAt}) //nAt order of the field in the expression
		   
	   ENDIF 
	   
    NEXT
	
	//we need to have the fields with the same order as in index expression nAt
	aFields := ASORT( aFields ,,, {|x,y| x[2] < y[2] } )

    FOR n := 1 TO LEN(aFields)
	    cStr += aFields[n,1]
		IF n < LEN(aFields)
		   cStr += ","
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
		
	       cOrder := " ORDER BY "+STRTRAN(cOrder,"+",",")
		   
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
   
   HB_SYMBOL_UNUSED( nWA )

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
	  
	  IF !ADO_GETLOCK(aWdata[WA_TABLENAME],aLockInfo[ UR_LI_RECORD ] ,nWA)
         lOk := .F.
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
	  
         AADD(aWdata[ WA_LOCKLIST ],aLockInfo[ UR_LI_RECORD ])
		 
	  ENDIF  
	  
      aLockInfo[ UR_LI_RESULT ] := .T.
	  
   ELSE
   
      aLockInfo[ UR_LI_RESULT ] := .F.  
	 
   ENDIF

   
   RETURN HB_SUCCESS

   
STATIC FUNCTION ADO_UNLOCK( nWA, xRecID )

   LOCAL aWdata := USRRDD_AREADATA( nWA ),n
   
   HB_SYMBOL_UNUSED( xRecId )
   HB_SYMBOL_UNUSED( nWA )
 
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
 LOCAL aLockCtrl :=  ADOLOCKCONTROL()
 LOCAL currArea := SELECT(),n
 
  xRecID := ALLTRIM(CVALTOCHAR(xRecID))
  
  IF SELECT("TLOCKS") = 0
     DBUSEAREA(.T.,aLockCtrl[2],aLockCtrl[1]+".DBF","TLOCKS",.T.)
     TLOCKS->(DBSETINDEX(aLockCtrl[1]+".CDX"))	 
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
 LOCAL aLockCtrl :=  ADOLOCKCONTROL()
 LOCAL currArea := SELECT(),n,aDels := {}
 
  xRecID := ALLTRIM(CVALTOCHAR(xRecID)) 
  
  IF SELECT("TLOCKS") = 0
     DBUSEAREA(.T.,aLockCtrl[2],aLockCtrl[1]+".DBF","TLOCKS",.T.)
     TLOCKS->(DBSETINDEX(aLockCtrl[1]+".CDX"))	 
	 SELECT(currArea)	 
  ENDIF
  
  
  IF ! xRecID = "ZWTXFL"
  
     n := ASCAN(aWData[WA_TLOCKS], {|a,z| a[z,1] = xRecID} ) 

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
	 
         IF aWData[WA_TLOCKS][n,1] <> "EXCLUSIVE" .OR. PROCNAME(2) = "ADO_CLOSE"//EXCLUSIVE ONLY WHEN WE CLOSE TABLE
		
	        TLOCKS->(DBGOTO(aWData[WA_TLOCKS][n,2]))
			
			IF TLOCKS->(ISLOCKED())
	           REPLACE TLOCKS->CODLOCK WITH SPACE(50) //NOT TO GROW IT WILL BE RECYCLED NEXT TIME
			ENDIF   
			
	        TLOCKS->(DBRUNLOCK(aWData[WA_TLOCKS][n,2]))
		    AADD(aDels,n)
 		    TLOCKS->(DBCOMMIT())
			
	     ENDIF	
		 
	 NEXT
      	
  ENDIF
  
  IF LEN( aDels) > 0
  
     FOR N :=  1 TO LEN(aDels)
	     ADEL(aWData[WA_TLOCKS],aDels[n],.T.) 
	 NEXT
 
  ENDIF
  
  
  RETURN HB_SUCCESS

  
FUNCTION ADO_ISLOCKED(cTable,xRecID,aWAData)  
  xRecID := ALLTRIM(CVALTOCHAR(xRecID)) 
  RETURN IF( aWAData[WA_OPENSHARED] ,ASCAN(aWAData[WA_TLOCKS], {|x| x[1] = ALLTRIM(CVALTOCHAR(xRecID))} )  > 0  ,.T.)
  
  
STATIC FUNCTION ADO_OPENSHARED( nWA,cTable, lExclusive)
 LOCAL aWAData := USRRDD_AREADATA( nWA )
 LOCAL lRetval := .F.
 LOCAL aLockCtrl :=  ADOLOCKCONTROL()
 LOCAL currArea := SELECT()
 
 /* IF SELECT("TLOCKS") = 0
     DBUSEAREA(.T.,aLockCtrl[2],aLockCtrl[1]+".DBF","TLOCKS",.T.)
     TLOCKS->(DBSETINDEX(aLockCtrl[1]+".CDX"))	 
	 SELECT(currArea)
  ENDIF

  IF !lExclusive
  
     IF TLOCKS->(DBSEEK( cTable + "EXCLUSIVE" ))
  
        IF TLOCKS->(DBRLOCK(  TLOCKS->(RECNO()) ))  //CAN BE OPENED SHARED?

		   TLOCKS->(DBRUNLOCK(  TLOCKS->(RECNO()) ))
	       lRetval := .T.
		   
	    ENDIF
		
	 ENDIF

     IF lRetval
	 
		IF TLOCKS->(DBSEEK( cTable + "SHARED" ))
		
		   DO WHILE TLOCKS->CODLOCK == cTable + "SHARED" .AND. !TLOCKS->(EOF())
		   
		      IF TLOCKS->(DBRLOCK(  TLOCKS->(RECNO()) )) //MARKED AS OPEN SHARED IF CANNOT GET LOCK SOMEONE HAS IT ALREADY
			     TLOCKS->(DBRUNLOCK(  TLOCKS->(RECNO()) ))
				 EXIT
			  ENDIF
			  
			  TLOCKS->(DBSKIP())
			  
           ENDDO	
		   
		   lRetval := .T.
		   
        ELSE
		
		   IF !TLOCKS->(DBSEEK(SPACE(50))) //RECOVERING USED RECORDS
	 
              TLOCKS->(DBAPPEND( .F.) ) //DOES NOT RELEASEANY LOCKS
	          IF !NETERR()
		         lRetval := .T.
		         REPLACE TLOCKS->CODLOCK WITH cTable+"SHARED"
		      ENDIF   
		
           ELSE
	 
              IF TLOCKS->(DBRLOCK( TLOCKS->(RECNO()) ))
	             lRetval := .T.
	             REPLACE TLOCKS->CODLOCK WITH cTable+"SHARED"
	          ENDIF	
		
	       ENDIF	
		   
		ENDIF
		
	 ENDIF
	 
  ELSE	 //EXCLUSIVE
  
     IF TLOCKS->(DBSEEK( cTable + "SHARED" ))
  
        IF TLOCKS->(DBRLOCK(  TLOCKS->(RECNO()) ))  //CAN BE OPENED SEXCLUSIVE?

		   TLOCKS->(DBRUNLOCK(  TLOCKS->(RECNO()) ))
	       lRetval := .T.
		   
	    ENDIF
		
	 ENDIF

     IF lRetval
	 
		IF TLOCKS->(DBSEEK( cTable + "EXCLUSIVE" ))
		
		   IF TLOCKS->(DBRLOCK(  TLOCKS->(RECNO()) )) 
		      lRetval := .T.
		   ENDIF
		   
        ELSE
		
		   IF !TLOCKS->(DBSEEK(SPACE(50))) //RECOVERING USED RECORDS
	 
              TLOCKS->(DBAPPEND( .F.) ) //DOES NOT RELEASEANY LOCKS
	          IF !NETERR()
		         lRetval := .T.
		         REPLACE TLOCKS->CODLOCK WITH cTable+"EXCLUSIVE"
		      ENDIF   
		
           ELSE
	 
              IF TLOCKS->(DBRLOCK( TLOCKS->(RECNO()) ))
	             lRetval := .T.
	             REPLACE TLOCKS->CODLOCK WITH cTable+"EXCLUSIVE"
	          ENDIF	
		
	       ENDIF	
		   
		ENDIF
		
	 ENDIF
  
  ENDIF
  
  IF lRetVal 
  
     TLOCKS->(DBCOMMIT())
	 IF lExclusive
        AADD(aWAData[WA_TLOCKS], { cTable + "EXCLUSIVE", TLOCKS->(RECNO()) } )
	 ENDIF	
	 
  ENDIF

*/
lretval := .t.  
  RETURN lRetval  
  
/*                              END LOCKS RELATED FUNCTIONS */   

/*                                             TRANSACTION RELATED FUNCTIONS */
STATIC FUNCTION ADO_FLUSH( nWA )

   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]
   LOCAL n
   
   TRY
   
      oRecordSet:Update()
	  
   CATCH
	  
   END
   

   RETURN HB_SUCCESS

   
FUNCTION ADOBEGINTRANS(nWa) 

 LOCAL oCon := hb_adoRddGetConnection( nWA )

 RETURN .T. //TRIALS
 
  IF oCon:BeginTrans() > 1 //WE ARE AREADY IN A TRANSACT
     oCon:RollbackTrans() // CLOSE TRANSACT STARTED ABOVE THERE IS NOTHING TO ROLLBACK
  ENDIF
  
RETURN .T.

FUNCTION ADOCOMMITTRANS(nWa)

 LOCAL oCon := hb_adoRddGetConnection( nWA ), n

 RETURN .T. //TRIALS
 
  n := oCon:BeginTrans() 
  
  IF n > 1  //WE ARE AREADY IN ATRANSACT
     oCon:RollbackTrans() // CLOSE TRANSACT STARTED ABOVE THERE IS NOTHING TO ROLLBACK
  ENDIF
  IF n > 1 // THERE IS A ACTIVE TRANSACTION IF IT WAS 1 WE WERE NOW OPENING A TRANACT
     oCon:CommitTrans()
  ENDIF
  
RETURN .T.

FUNCTION ADOROLLBACKTRANS(nWa) 

 LOCAL oCon := hb_adoRddGetConnection( nWA ), n
 
RETURN .T. //TRIALS

  n := oCon:BeginTrans() 
  
  IF n > 1  //WE ARE AREADY IN ATRANSACT
     oCon:RollbackTrans() // CLOSE TRANSACT STARTED ABOVE THERE IS NOTHING TO ROLLBACK
  ENDIF
  IF n > 1 // THERE IS A ACTIVE TRANSACTION IF IT WAS 1 WE WERE NOW OPENING A TRANACT
     oCon:RollbackTrans()
  ENDIF

RETURN .T.
/*                                      END TRANSACTION RELATED FUNCTIONS */


/*                                     LOCATE SEEK FILTER RELATED FUNCTIONS */
STATIC FUNCTION ADO_SETFILTER( nWA, aFilterInfo )

   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]
   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL cSql :="",oError, nRecNo 
 
      IF VALTYPE(aFilterInfo) = "A" .AND. EMPTY(aFilterInfo[ UR_FRI_CEXPR ])
	  
         MSGINFO("ADORDD doesnt supports dbfilters without cFilterExpression!"+;
		         PROCNAME(1)+" "+STR(PROCLINE(1)))

         oError := ErrorNew()
         oError:GenCode := EG_ARG
         oError:SubCode := 1112
         oError:Description := "dbfilter without cExpr not valid in ADORDD"
         oError:FileName := "<nenhuma>"
         oError:OsCode := 0 
         oError:CanDefault := .F.

         RETURN HB_FAILURE
		 
	  ENDIF
	  
	  IF VALTYPE(aFilterInfo) = "A"
	  
         aWAData[WA_FILTERACTIVE] := aFilterInfo[ UR_FRI_CEXPR ] //SAVE ACTIVE FILTER EXPRESION
		 
	  ELSE	 //CHECKING ACTVE FILTER IF ONE
	  
	     IF EMPTY(aWAData[WA_FILTERACTIVE])
		    RETURN HB_SUCCESS  //NONE CONTINUE WITH CURRENT TASK
		 ENDIF
		 
	  ENDIF
	  
	  ADO_RECID( nWA, @nRecNo )
	  
	  cSql :=  Filter2Sql(nWA, aWAData[WA_TABLENAME], aWAData[WA_FILTERACTIVE])

      oRecordSet:Close()   
	  oRecordSet:Open(cSql,aWAData[WA_CONNECTION])

      ADO_GOTOID( nWA, nRecNo )
	  IF oRecordSet:Eof() //does not have this rec in filter lets gotop
	     ADO_GOTOP(nWa)
	  ENDIF
	  
   RETURN HB_SUCCESS

   
STATIC FUNCTION ADO_CLEARFILTER( nWA )

   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]
   LOCAL aOrderInfo := ARRAY(UR_ORI_SIZE)
   
   aWAData[WA_FILTERACTIVE] := NIL //NO FILTER
 
   IF !"CLOSE" $ PROCNAME(1) //IF WE CLOSING IT DO NOTHING UR_SUPER_CLOSE CALL THIS!
      // THIS CREATE A NEW SELECT WITH ACTIVE ORDER ALL RECORDS
      ADO_ORDLSTFOCUS( nWA, aOrderInfo )  // UR_ORI_TAG = nil UR_ORI_RESULT = current order
      aOrderInfo[UR_ORI_TAG] := aOrderInfo[UR_ORI_RESULT] //taG = current order
      ADO_ORDLSTFOCUS( nWA, aOrderInfo ) // new set with current order NO FILTER
   ENDIF
   
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
      ADO_SKIPRAW(nWA,1) //WE DONT WNAT TO FIND THIS ONE AGAIN
   ENDIF
   CURSORWAIT()
   DO WHILE !aWAData[ WA_EOF ]
      IF EVAL( aWAData[ WA_SCOPEINFO ][ UR_SI_BFOR ])
	     aWAData[ WA_FOUND ] := .T.
	     EXIT
	  ENDIF	 
      ADO_SKIPRAW(nWA,1)
	  SYSREFRESH()
   ENDDO
   CURSORARROW()
    
   aWAData[ WA_FOUND ] := ! oRecordSet:EOF
   aWAData[ WA_EOF ] := oRecordSet:EOF
  
   
   RETURN HB_SUCCESS
   
   
STATIC FUNCTION ADO_SEEK( nWA, lSoftSeek, cKey, lFindLast )

   RETURN ADOSEEKSQLFIND( nWA, lSoftSeek, cKey, lFindLast )
   
   
/* NEW VERSION */
STATIC FUNCTION ADOSEEKSQLFIND( nWA, lSoftSeek, cKey, lFindLast )
   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]
   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL aSeek,cSql,oRs := TempRecordSet()
   
   HB_SYMBOL_UNUSED( nWA )
   HB_SYMBOL_UNUSED( lSoftSeek )
   HB_SYMBOL_UNUSED( cKey )
   HB_SYMBOL_UNUSED( lFindLast )
   
   DEFAULT lFindLast TO .F.
   DEFAULT lSoftSeek TO .F.

 /*  IF ADOEMPTYSET(oRecordSet)
      aWAData[ WA_FOUND ] := ! oRecordSet:EOF
      aWAData[ WA_EOF ] := oRecordSet:EOF
      RETURN HB_SUCCESS
   ENDIF
   */
   IF aWAData[WA_INDEXACTIVE] = 0
      MSGALERT("No Index active seek not allowed!") //SHOULD RAISE ERROR
	  RETURN HB_FAILURE
   ENDIF
   
   aSeek := ADOPseudoSeek(nWA,cKey,aWAData,lSoftSeek)
   IF lSoftSeek
      
   ENDIF

   cSql := IndexBuildExp(nWA,aWAData[WA_INDEXACTIVE],aWAData,.F.,IF(aSeek[3],aseek[1],aSeek[2] ))

   oRs:CursorType :=   adOpenDynamic 
   oRs:CursorLocation := IF(aWAData[ WA_ENGINE ] = "ACCESS", adUseClient, adUseServer) //adUseServer  // adUseClient its slower but has avntages such always bookmaks 
   oRs:LockType := IF(aWAData[ WA_ENGINE ] = "ACCESS", adLockOptimistic, adLockPessimistic)   
  
   oRs:Open(cSql,aWAData[ WA_CONNECTION ] )
  
   IF lSoftSeek .AND. ADOEMPTYSET(oRs)  // DIDNT FIND SFOTSEEK ON LOOK FOR THE NEXT KEY where field > key
      cSql := STRTRAN(cSql," = "," > ")
      cSql := STRTRAN(cSql,"SELECT ","SELECT TOP 1 ") //first following record SOFTSEEK
 	  oRs:Close()
	  oRs:Open(cSql,aWAData[ WA_CONNECTION ] )
   ENDIF

   ADOFINDREC(nWA,aWAdata,oRs,oRecordSet,lFindLast)
 
   RETURN HB_SUCCESS
   
   
STATIC FUNCTION ADOFINDREC(nWA,aWAdata,oRs,oRecordSet,lFindLast)


   IF !ADOEMPTYSET(oRs) //FOUNDED!
  
      IF lFindLast
	     oRs:MoveLast()
	  ELSE
	     oRs:MoveFirst()
	  ENDIF	  
	 
	  IF !EMPTY(aWAData[WA_FIELDRECNO])  // 100% SUPPORTED AND SAFE
	  
         IF oRecordSet:Supports(adIndex) .AND. oRecordSet:Supports(adSeek)
		    oRecordSet:Index := oRecordSet:Fields(aWAData[WA_FIELDRECNO]):Name
			oRecordSet:Seek({ ALLTRIM(STR(oRS:Fields(aWAData[WA_FIELDRECNO]):Value,0)) })
		 ELSE
		    oRecordSet:MoveFirst()
		    oRecordSet:Find(oRecordSet:Fields(aWAData[WA_FIELDRECNO]):Name+" = "+ALLTRIM(STR(oRS:Fields(aWAData[WA_FIELDRECNO]):Value,0)) )
         ENDIF
		 
 	  ELSE

		 IF oRS:Supports(adBookmark) 
		
		    //Although the Supports method may return True for a given functionality, it does not guarantee that 
		    //the provider can make the feature available under all circumstances. 
		    //The Supports method simply returns whether the provider can support the specified functionality,
		    //assuming certain conditions are met. For example, the Supports method may indicate that a 
		    //Recordset object supports updates even though the cursor is based on a multiple table join, 
		    //some columns of which are not updatable
		    IF oRS:Eof() .or. oRS:Bof()
		    ELSE	  
			   oRecordSet:BookMark := oRS:BookMark //not guarantee to work
		    ENDIF
		   
		 ELSE		
		
		    //ATTENTION NOT WORKING CORRECTLY WITH DELETED ROWS!2	
		    oRecordSet:AbsolutePosition := IF( oRS:AbsolutePosition == adPosEOF, oRS:RecordCount() + 1, oRS:AbsolutePosition )
		    //MUST TAKE OUT THE DELETED ROWS! OTHERWISE WRONG NRECNO 
		    //TODO nRecno := nRecno-nDeletedRows 
		   
		 ENDIF	
		
	  ENDIF
	 
   ELSE  //NOT FOUND!

       IF !ADOEMPTYSET(oRecordSet) 
	      oRecordSet:MoveLast()
	      oRecordSet:MoveNext()  // eof()
	   ENDIF	  

   ENDIF

   oRs:close()
	  
   aWAData[ WA_FOUND ] := ! oRecordSet:EOF
   aWAData[ WA_EOF ] := oRecordSet:EOF
   
   //TO CHECK NEXT CALLS IF WE ARE IN A SUBSSET TO REVERT TO DEFAULT SET
   aWAData[WA_ISITSUBSET] := .F.

   IF !EMPTY(aWAData[WA_PENDINGREL]) .AND. PROCNAME(2) <> "ADO_RELEVAL" //ENFORCE REL CHILDS BUT NOT IN A ENDLESS LOOP!
      ADO_FORCEREL( nWA ) 
   ENDIF

  RETURN HB_SUCCESS
  
  
//VERSION WITHOUT :FIND o becalled directly from the app
//you must then call ADORESETSEEK when you dont need anymore this subset
FUNCTION ADOSEEKSQL( nWA, lSoftSeek, cKey, lFindLast )
   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]
   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL aSeek,cSql
   
   HB_SYMBOL_UNUSED( nWA )
   HB_SYMBOL_UNUSED( lSoftSeek )
   HB_SYMBOL_UNUSED( cKey )
   HB_SYMBOL_UNUSED( lFindLast )
   
   DEFAULT lFindLast TO .F.
   DEFAULT lSoftSeek TO .F.

   IF aWAData[WA_INDEXACTIVE] = 0
      MSGALERT("No Index active seek not allowed!") //SHOULD RAISE ERROR
	  RETURN HB_FAILURE
   ENDIF
   
   aSeek := ADOPseudoSeek(nWA,cKey,aWAData,lSoftSeek)
   
   cSql := IndexBuildExp(nWA,aWAData[WA_INDEXACTIVE],aWAData,.F.,IF(aSeek[3],aseek[1],aSeek[2] ))
   oRecordSet:Close() 
   oRecordSet:Open(cSql,aWAData[ WA_CONNECTION ] )
   
   //TO CHECK NEXT CALLS IF WE ARE IN A SUBSSET TO REVERT TO DEFAULT SET
   aWAData[WA_ISITSUBSET] := .T.

   IF !ADOEMPTYSET(oRecordSet) //FOUND!
	  
	  IF lFindLast
		 oRecordSet:MoveLast()
	  ELSE
		 oRecordSet:MoveFirst()
	  ENDIF	  
		 
   ENDIF

   aWAData[ WA_FOUND ] := ! oRecordSet:EOF
   aWAData[ WA_EOF ] := oRecordSet:EOF

   IF !EMPTY(aWAData[WA_PENDINGREL]) .AND. PROCNAME(1) <> "ADO_RELEVAL" //ENFORCE REL CHILDS BUT NOT IN A ENDLESS LOOP!
      ADO_FORCEREL( nWA ) 
   ENDIF
 
   RETURN HB_SUCCESS
   
   
FUNCTION ADORESETSEEK()  //RESET THE RECORDSET WITHOUT PREVIOUS SEEK (WHERE) EXPRESSION

   LOCAL nWA := SELECT()
   LOCAL oRecordSet := USRRDD_AREADATA( nWA )[ WA_RECORDSET ]
   LOCAL aWAData := USRRDD_AREADATA( nWA )
   LOCAL cSql

   cSql := IndexBuildExp(nWA,aWAData[WA_INDEXACTIVE],aWAData,.F.)
   oRecordSet:Close() 
   oRecordSet:Open(cSql,aWAData[ WA_CONNECTION ] )

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
 cKeyBotom := CVALTOCHAR(cKeybottom)
 
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
		         cExpression    += aFields[nAt,1]+ "="+"'"+ADODTOS(SUBSTR( cKey, 1, nLen))+"'" //delim might be #
		         cSqlExpression += aFields[nAt,1]+ "='"+ADODTOS(SUBSTR( cKey, 1, nLen))+"'"
		      ELSE
	             cExpression += aFields[nAt,1]+" BETWEEN "+"'"+ADODTOS(SUBSTR( cKey, 1, nLen))+"'"+;
			                 " AND "+"'"+ADODTOS(SUBSTR( cKeyBottom, 1, nLen))+"'"
		         cSqlExpression := cExpression
              ENDIF
			  
		   ELSE
		      
			  cVal := ALLTRIM(STR(VAL(SUBSTR( cKey, 1, nLen))))
			  
		      IF !lBetween
		         cExpression    += aFields[nAt,1]+ "="+"#"+cVal+"#"
		         cSqlExpression += aFields[nAt,1]+ "="+cVal
		      ELSE
	             cExpression += aFields[nAt,1]+" BETWEEN "+"#"+cVal+"#"+;
			                 " AND "+"#"+cVal+"#"
		         cSqlExpression := STRTRAN(cExpression,"#","")
              ENDIF
			  
		   ENDIF	  
		   
		ELSEIF  cType = "L" 
		
		   nLen := 3 // although is one in the table in the stirng is 3 ex .t. or .f.
		   IF SUBSTR( UPPER(cKey), 1, nLen) = ".T."
	          cExpression += aFields[nAt,1]
		   ELSE
		      cExpression += " NOT "+aFields[nAt,1]
           ENDIF		   
		   cExpression := STRTRAN( UPPER(cExpression), ".T.","True",1,1)
		   cExpression := STRTRAN( UPPER(cExpression), ".F.","True",1,1)
		   cSqlExpression := cExpression

		ELSE

     	   lNotFind := .T.	//expression cannot be used by :Find()	
		   
		ENDIF
		
		cKey := SUBSTR(cKey,nLen+1) // take out the len of the expression already used
		IF LBetween
		   cKeybottom := SUBSTR(cKeybottom,nLen+1) // take out the len of the expression already used
		ENDIF
		
		IF nAt < LEN(aFields) //add AND last one isnt needed!
		   
		   IF !lSoftSeek  // EXPRESSION FILED-FIELD > "  LEN 2 FIELDS    " DONT NEED AND
		      cExpression += IF(LEN(cKey) > 0 ," AND " , "")
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
	
  RETURN {cExpression,cSqlExpression,IF( lNotFind,.F.,nAt = 1)}
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
   LOCAL aInfo, nReturn, nOrder, uResult
 
   nReturn := ADO_EVALBLOCK( aRelInfo[ UR_RI_PARENT ], aRelInfo[ UR_RI_BEXPR ], @uResult )
  
   IF nReturn == HB_SUCCESS

/*      IF VALTYPE(aWAData[WA_LASTRELKEY]) <> "U" 
         IF aWAData[WA_LASTRELKEY] == uResult //KEY DIDNT CHANGED DONT HAVE TO SEEK AGAIN
		    RETURN nReturn
	     ELSE
		    aWAData[WA_LASTRELKEY] := uResult
		 ENDIF	
      ENDIF
*/
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
            ELSE			   
               /* should raise error child not indexed*/
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
   
   IF PROCNAME(1) <> "ADO_RELEVAL"
      // DONT KNOW WHY BUT DBEVAL ONLY WORK LIKE THIS
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
STATIC FUNCTION ADO_CREATE( nWA, aOpenInfo )

  LOCAL aWAData := USRRDD_AREADATA( nWA )
  LOCAL cDataBase  := hb_tokenGet( aOpenInfo[ UR_OI_NAME ], 1, ";" )
  LOCAL cTable     := hb_tokenGet( aOpenInfo[ UR_OI_NAME ], 2, ";" )
  LOCAL cDbEngine  := hb_tokenGet( aOpenInfo[ UR_OI_NAME ], 3, ";" )
  LOCAL cServer    := hb_tokenGet( aOpenInfo[ UR_OI_NAME ], 4, ";" )
  LOCAL cUserName  := hb_tokenGet( aOpenInfo[ UR_OI_NAME ], 5, ";" )
  LOCAL cPassword  := hb_tokenGet( aOpenInfo[ UR_OI_NAME ], 6, ";" )
  LOCAL cSql, cSql2, lAddAutoInc := .F.
  LOCAL oCatalog , cMarkTmp
 
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
	     oCatalog    := TOleAuto():New( "ADOX.Catalog" )
	     oCatalog:Create( "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + cDataBase )
	  ENDIF 
   ENDIF

 
   aOpenInfo[ UR_OI_NAME ] := CFILENOEXT( CFILENOPATH( cTable ) )

   ADOCONNECT(nWA,aOpenInfo)
   
   cSql := ADOSTRUCTTOSQL( aWAData,aWAData[ WA_SQLSTRUCT ],@lAddAutoInc )

   IF Lower( Right( aOpenInfo[ UR_OI_NAME ], 4 ) ) == ".dbf" .OR. aWAData[ WA_ENGINE ] = "ADS"
      aWAData[ WA_TABLENAME ] := CFILENOEXT(CFILENOPATH(aWAData[ WA_TABLENAME ] ))
   ENDIF
   
   
   TRY
       //TEMPORARY TABLES ARE DESTROYE AUTO WHEN NO NEEDED ANYMORE
       IF UPPER( SUBSTR( aWAData[ WA_TABLENAME ] ,1,3 ) ) = "TMP" .OR. UPPER( SUBSTR( aWAData[ WA_TABLENAME ] ,1,4 ) ) = "TEMP"
          cMarkTmp := ADOTEMPTABLE( cDbEngine )
		  //we need to create the file in order to eventually test file() works in the app
		  MEMOWRIT(cTable,"nada")
	   ELSE
          cMarkTmp	:= "TABLE "
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
	  
      END
 
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
          CASE aStruct[ nCol,2 ] = 'N'
             c     := LTrim( Str( aStruct[ nCol,3 ] + 1 ) ) + ", " + LTrim( Str( aStruct[ nCol,4 ] ) )
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
		 
            cSql  += " CHAR(1) "+aStruct[ nCol,2 ]
			
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

FUNCTION ADOTEMPTABLE(DbEngine)
 LOCAL cMark := ""
 
  DO CASE 

  CASE dbEngine = "ADS"
     cMark := "TABLE #"
  CASE dbEngine = "ACCESS"
     cMark     := "TABLE "
  CASE dbEngine = "MSSQL"
     cMark := "TABEL #"
  CASE dbEngine = "DBASE"
	 cMark     := "TABLE "
  CASE dbEngine = "SQLITE"
	 cMark    := "TEMPORARY TABLE "
  CASE dbEngine = "MYSQL"
     cMark :=  "TEMPORARY TABLE "
  CASE dbEngine = "FOXPRO"
	 cMark     := "TABLE "
  CASE dbEngine = "ORACLE"
     cMark     := "GLOBAL TEMPORARY TABLE"
  CASE dbEngine = "POSTGRE"
	 cMark     := "TEMPORARY TABLE "
  ENDCASE
 
  RETURN cMark

  
STATIC FUNCTION ADOFILE( oCn, cTable, cIndex)

   LOCAL lRet := .F.
   LOCAL oRs := TOleAuto():New( "ADODB.Recordset" )
   LOCAL aIndexes := ListIndex(), z, y 
   
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
                 oRs:Filter  := "TABLE_NAME = '" + cTable + "' AND TABLE_TYPE = 'TABLE'"
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
 
   RETURN lRet
  

STATIC FUNCTION ADODROP( oCon, cTable, cIndex ,DBEngine)

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
	     ADOSHOWERROR( aWAData[ WA_CONNECTION ], .f. )
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
		 
         oCn:Execute( cSql )
         lRet := .T.

	  CATCH
	  
	     ADOSHOWERROR( aWAData[ WA_CONNECTION ], .f. )
		 
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
	
	     uReturn := IF(LEN(aWAData[WA_PENDINGREL]) > 0, LEN(aWAData[WA_PENDINGREL]) / 7,0)
		 
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
	
	     uReturn := "{||"+aWAData[ WA_LOCATEFOR ]+"}"
		 
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
   aADOFunc[ UR_TRANS ]        := (@ADO_TRANS())
   aADOFunc[ UR_TRANSREC ]     := (@ADO_TRANSREC())

  
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
   
   cDate := DTOS(dDate)
   
   cYear  := SUBSTR(cDate,1,4)
   cMonth := SUBSTR(cDate,5,2)
   cDay   := SUBSTR(cDate,7,2) 
   
   IF( EMPTY(cYear),cYear :="1901",cYear)
   IF( EMPTY(cMonth),cMonth := "01",cMonth)
   IF( EMPTY(cDay),cDay := "01",cDay)
   
   cDate  := cYear+"-"+cMonth+"-"+cDay // hope to enforce set date format like this
   
   //FW_DateToSQL( CTOD(cDate), "D",  )
   
   RETURN cDate

   
STATIC FUNCTION ADOEMPTYSET(oRecordSet)
   RETURN (oRecordSet:Eof() .AND.  oRecordSet:Bof() )


STATIC FUNCTION SQLTranslate( cExpr )
 LOCAL n, nSpacePos, aStr:= {}
 
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
   cEXpr := STRTRAN( cExpr,"FIELD->","")
   cEXpr := STRTRAN( cExpr,".",",")

   nSpacePos := 1
   FOR n := 1 TO LEN(cExpr)
       IF SUBSTR(cExpr,n,1) = " "
	      nSpacePos := n
	   ENDIF
       IF SUBSTR(cExpr,n,2) = "->"
	      AADD( aStr,SUBSTR(cExpr, nSpacePos,n+1) )
		  nSpacePos := 0
	   ENDIF
   NEXT
   
   FOR N:= 1 TO LEN(aStr)
       cExpr := STRTRAN(cExpr,aStr[n],"")
   NEXT
   
   RETURN cExpr
   
   
FUNCTION Filter2Sql(nArea,cTableName,cFiltro) //ahf taken from old app sorry its portuguese

LOCAL n, cSelect :="",cAliasCorr := alias(nArea)
LOCAL nat := 0
LOCAL calias,nandou := 0,nsizealias := 0
LOCAL tmpexp := "%$",cvalor 
LOCAL corderexp :=""
LOCAL operators := {' DIFERENTE ', ' DIFERENTE ', ' MENORIGUAL ', ' MAIORIGUAL ','.',' = ', ' < ', ' > ', ;
        ' LIKE ', ' SOUNDEX( ', ' AND ', ' OR ',' NOT '  ,;
		" ABS( "," ROUND( "," CHAR_LENGTH( "," TRIM( "," LTRIM( ",;
        " RTRIM( "," UPPER( "," LOWER( ",;
        " SUBSTRING( ",;
        " SPACE( ",;
        " CURRDATE( "," YEAR( "," MONTH( ",;
        " DAYOFMONTH( "," CURRTIME( "," IIF( "}
//THESE HAVE TO BE ACC TO HE DB USED THESE FOR ADS		
LOCAL aitems := {'<>', '< >','<=', '>=','->','=', '<', '>', ;
        '$', 'Soundex(', '.AND.', '.OR.','.NOT.' ,;
		"Abs(","Round(","Len(","AllTrim(","LTrim(","RTrim(",;
        "Upper(","Lower(","Substr(",;
        "Space(","Date(","Year(","Month(",;
        "Day(","Time(","If("}
LOCAL lProximo := .f.,cKFilter,acampos := {}, cfilter:="" ,cordfor :=""
LOCAL cselalias := cTableName,cjoin :=""
LOCAL arelfields :={},areldbfs := {},aaliases := {},nparams := 0,cnewordexp:="",cnewordfor:=""
LOCAL aOrders := {}, nconvert :=0
//18.2.01
//THESE HAVE TO BE ACC TO HE DB USED THESE FOR ADS
LOCAL anaopermitidos := {"DTOS(","CTOD(","DTOC(","STR(","VAL(","EMPTY("}
LOCAL asqlini := {"CONVERT(","CONVERT(","CONVERT(","CONVERT(","CONVERT(","TRIM("}
LOCAL asqlfim := {",SQL_CHAR)",",SQL_DATE)",",SQL_CHAR)",",SQL_CHAR)",",SQL_NUMERIC)",")=''"},nNot
LOCAL aInfo,nRelArea,cExpRel

  //18.2.01 COMO STRING JA NAO E UPPER TEMOS DE SUBSTITUIR TODOS OS ITEMS POR UPPR ITEM
  for n := 1 to len(aitems)
      cfiltro := strtran(cFiltro,aitems[n],upper(aitems[n])) 
	  cFiltro := strtran(cFiltro,upper(substr(aitems[n],1,1))+substr(aitems[n],2),upper(aitems[n]))  //capitalizadas
	  cFiltro := strtran(cFiltro,lower(aitems[n]),upper(aitems[n]))
  next
  for n := 1 to len(anaopermitidos)
      cfiltro := strtran(cFiltro,anaopermitidos[n],upper(anaopermitidos[n])) 
	  cfiltro := strtran(cFiltro,lower(anaopermitidos[n]),upper(anaopermitidos[n])) 
	  cFiltro := strtran(cFiltro,upper(substr(anaopermitidos[n],1,1))+lower(substr(anaopermitidos[n],2)),upper(anaopermitidos[n]))  //capitalizadas
  next
  
  //ISTO PARA DADA SCALAR FUNCTION SUBSITIUI FUNCOES NAO SUPORTADAS PELO SQL MAS SO POR FILTROS
  for nNot := 1 to len(anaopermitidos)
      n:=1
      do while n <= len(cfiltro)
         nat := 0
         nat:= at(anaopermitidos[nNot], cFiltro)
	     if nat >0
	        cFiltro := stuff(cFiltro,nat,len(anaopermitidos[nNot]),asqlini[nNot] )
		    nat := at(")",cfiltro,nat)
		    cFiltro := stuff(cFiltro,nat,1,asqlfim[nNot] )
	     endif
         n++
      enddo
  next	  
  //18.2.15

  //expressao for indice
  aInfo := Array( UR_ORI_SIZE )
  ADO_ORDINFO( nArea, DBOI_CONDITION, @aInfo ) //(nWA)->(ORDfor(nIndex))
  cordfor := aInfo[UR_ORI_RESULT]
  if !empty(aInfo[UR_ORI_RESULT])
     cordfor := " AND "+upper(aInfo[UR_ORI_RESULT])
  endif
  
  //expressao ordem indice
  aInfo := Array( UR_ORI_SIZE )
  ADO_ORDINFO( nArea, DBOI_EXPRESSION, @aInfo ) //(nWA)->(ORDKEY(nIndex))
  cOrderExp := aInfo[UR_ORI_RESULT]

  //expressao orderfor tudo cm aliases para o alias corrente nunca tem aliases colocamos
  for n:= 1 to (cAliasCorr)->(fcount())
      cordfor := strtran(cordfor,upper((cAliasCorr)->(fieldname(n))),;
	             upper(cTableName+"."+(cAliasCorr)->(fieldname(n))))
	  //PARA EVITAR QUE PRIO EXEMPLO DATAGUIA SEJA SUBSTITUIDA POR GUIA FICANDO DATAALIAS.GUIA
	  //26.4.15 CHANGED IN BY $
	  if substr(cfiltro,at(alltrim(upper((cAliasCorr)->(fieldname(n)))),cfiltro)-1,1) $ " (),.##%$=<>"			 
         cfiltro := strtran(cfiltro,alltrim(upper((cAliasCorr)->(fieldname(n)))),;
	             upper( cTableName+"."+alltrim((cAliasCorr)->(fieldname(n)))))
      endif				 
	  x := at(ALLTRIM(UPPER((cAliasCorr)->(fieldname(n))) ),corderexp ) 
	  if x > 0 //index para campos orderarem pela mesma ordem que indice
	     aadd(aorders,alltrim(str(x,0))+"->"+upper(cTableName+"."+(cAliasCorr)->(fieldname(n)))+" ASC#% ")
	 endif
  next

  asort(aorders)
  for n:= 1 to len(aorders)
      cnewordexp += alltrim(substr(aorders[n],at("->",aorders[n])+2))
  next

  corderexp:= cnewordexp	

 //corderexp := strtran(corderexp,"+","")
  //ULTIMA VIRGULA SEM NADA A SEGUIR
  corderexp := substr(alltrim(corderexp),1,len(alltrim(corderexp))-2)
  
  while .t. //expressao aliases filtro

      nat := at("->",cfiltro)

      if nat > 0
         //temos de retirar este "->" para
         //nao passarmos por ele outra vez
         //ate acabramos
         cfiltro := stuff(cfiltro,nat,2,tmpexp)

         nandou := 0

         //temos de andar para tras ate ao primeiro espaco
         //para apanhar nome do alias
         for n:= nat-1 to 0 step -1

             //se chegamos a um espaco ou ao principio do expfiltro
             //temos o alias
             cvalor := substr(cfiltro,abs(n),1)

             if cvalor == " " .or. cvalor $ ".;,+/-*=()[]&%$#'" .or. n = 0

                //extraimos o alias
                calias := alltrim(substr(cfiltro,nat-nandou,nandou))
                nsizealias := len(calias)
                if upper(calias) <> upper(cselalias)
                  //transformamos o alias virtual
                  //que pode ou nao ser o desta janela no alias
                  //verdadeiro e depois no virtual desta janela
				  aadd(Aaliases,calias)
                  calias := AliasToTableName(calias)
                
                   //agora juntamos a expfiltro
                   cfiltro := stuff(cfiltro,nat-nandou,nsizealias,calias)
				 endif  
                exit
             endif
             nandou ++
         next
      else
         exit
      endif
  enddo
 //parte final expressao juntar tudo
 cfiltro :="SELECT * "+;
            "FROM "+cTableName+" "+cjoin+;
            "WHERE "+cfilter+cfiltro+cordfor+IF(!EMPTY(corderexp)," ORDER BY "+corderexp,"")
 
 //voltar a colocar as "->"
 cfiltro := strtran(cfiltro,tmpexp,"->")
 cfiltro := strtran(cfiltro,"->",".")
 cfiltro := strtran(cfiltro,'"',"'")
 
 for n := 1 to len(aitems)
     //if operators[n] <> ""
	    cfiltro := strtran(cFiltro,upper(aitems[n]),operators[n])
	 //endif
 next
 cfiltro := strtran(cfiltro,"DIFERENTE","<>") //PARA NAO INTERFERIR COM < E >
 cfiltro := strtran(cfiltro,"MAIORIGUAL",">=") //PARA NAO INTERFERIR COM < E >
 cfiltro := strtran(cfiltro,"MENORIGUAL","<=") //PARA NAO INTERFERIR COM < E >
 cfiltro := strtran(cfiltro,"#%",",")
 //limpar parantesis ) SEM PAR
 n :=1
 lproximo := .f.
 do while n <= len(cFiltro)
    if substr(cFiltro,n,1) = "("
	   lproximo := .t.
	endif   
    if substr(cFiltro,n,1) = ")"
	   if !lproximo
          cfiltro := stuff(cfiltro,n,1,"")
          lproximo := .f.
        endif		  
    endif
    n++
enddo
 cSelect := cfiltro


RETURN cSelect //end


FUNCTION AliasToTableName(cAlias)

 LOCAL afiles := ListIndex()
 LOCAL x,cNome, lenfiles := len(afiles)

 cnome := calias
 
  for x := 1 to lenfiles
      if at(afiles[x,1],cAlias) > 0
         cnome := afiles[x,1]
         exit
      endif
  next

RETURN cNome
   
/*                                  END  GENERAL */

   
/*                    ADO SET GET FUNCTONS */
   
   
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

            MSGALERT( cErr, IF( oCn:Provider = NIL, "ADO ERROR",oCn:Provider ) )
         END
      ENDIF
   ELSE
      MSGALERT( "ADO ERROR UNKNOWN" )
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
   
  
FUNCTION hb_adoRddExistsTable( oCn, cTable, cIndex )
   RETURN ADOFILE( oCn, cTable, cIndex )

FUNCTION hb_adoRddDrop( oCn, cTable, cIndex, DBEngine )
   RETURN ADODROP( oCn, cTable, cIndex, DBEngine )
   
   
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

   
 /* default values for adordd to use if not present in open or create */
FUNCTION ADODEFAULTS( cDB, cServer, cEngine, cUser, cPass,lGetThem )   
 STATIC aDefaults := {}

   DEFAULT lGetThem TO .T. //DEFAULT CALL WITHOUT PARAMS GET DEFAULT ARRAY
   
   IF !lGetThem	//RESET THEM
   
      aDefaults := {}
      AADD(aDefaults, cDB )
      AADD(aDefaults, cServer )
      AADD(aDefaults, cEngine )
      AADD(aDefaults, cUser )
      AADD(aDefaults, cPass )	
	  
   ELSE
  
	  DEFAULT t_cQuery TO "SELECT * FROM "
	  DEFAULT t_cUserName TO aDefaults[4]
	  DEFAULT t_cPassword TO aDefaults[5]
	  DEFAULT t_cServer TO aDefaults[2]
	  DEFAULT t_cEngine TO aDefaults[3]
	  DEFAULT t_cDataSource TO aDefaults[1]
   
   ENDIF
 
   RETURN aDefaults

   
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
   
   
//SQL UNIQUE EXP   
STATIC FUNCTION ListTmpDbfUnique()
 STATIC aTmpDbfUnique := {} 
 RETURN aTmpDbfUnique 

 
//DBF UNIQUE EXP 
STATIC FUNCTION ListTmpUnique()
  STATIC aTmpUniques := {}
  RETURN aTmpUniques
  

STATIC FUNCTION TempRecordSet() //USED IN ADO_SEEK AVOID OVERTIME NEW OBJ RECORDSET 
  STATIC oRs
  
  IF EMPTY(oRs)
     oRs := TOleAuto():New( "ADODB.Recordset" )
  ENDIF
  
  RETURN oRs
  
FUNCTION hb_GetAdoConnection()//supply app the con object
  RETURN oConnection

//ceate table for record lock control
FUNCTION ADOLOCKCONTROL(cPath,cRdd)
 STATIC cFile,rRdd
 
 LOCAL lRetVal := .F.
 LOCAL cTable 
 LOCAL cIndex 
  
  DEFAULT cRdd TO "DBFCDX"
  
  IF !EMPTY(cRdd)
     rRdd := cRdd
  ENDIF	 
  IF !EMPTY(cPath)
     cFile := cPath+"\TLOCKS"
  ELSE
     RETURN {cFile,rRdd}
  ENDIF	 
  
  cTable := cPath+"\TLOCKS.DBF"
  cIndex := cPath+"\TLOCKS.CDX"
  
  TRY
  
      IF !FILE(cTable)
	  
         DBCREATE(cTable,;
		 { {"CODLOCK","C",50,0 }},;
		 "DBFCDX",.T.,"TLOCKS")
         INDEX ON CODLOCK TO (cIndex)
		 
      ENDIF
	  
	  lRetVal := .T.
	  
  CATCH
  
      lRetVal := .F.
  END

 
  RETURN lRetVal
  
  
FUNCTION ADOVERSION()  
//version string = nr of version . post date() / sequencial nr in the same post date
RETURN "AdoRdd Version 1.060515/1"

/*                   END ADO SET GET FUNCTONS */
   
//TO DEBGUG
FUNCTION ARRAYTOCHAR(AARRAY)
LOCAL N
FOR N:= 1 TO LEN(AARRAY)
   AARRAY[N] := CVALTOCHAR(AARRAY[N] )
NEXT
RETURN AARRAY

function fwAdoConnect()

   local oDL,   cConnection := ""

   oDL = CreateObject( "Datalinks" ):PromptNew()

   if ! Empty( oDL )
      cConnection = oDL:ConnectionString
   endif


return nil

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
 
 
