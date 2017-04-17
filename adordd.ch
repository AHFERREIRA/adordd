/*
 * Harbour Project source code:
 * ADORDD - RDD to automatically manage Microsoft ADO
 *
 * Copyright 2007 Fernando Mancera <fmancera@viaopen.com> and
 * Antonio Linares <alinares@fivetechsoft.com>
 * www - http://harbour-project.org
 *
 * Copyright 2015 AHF - Antonio H. Ferreira <disal.antonio.ferreira@gmail.com>
 * Constant Group: Supports
 * COMANDS AFTER LOCATE
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

#ifndef _ADORDD_CH
#define _ADORDD_CH

/* Cursor Type */
#define adOpenForwardOnly               0
#define adOpenKeyset                    1
#define adOpenDynamic                   2
#define adOpenStatic                    3

/* Lock Types */
#define adLockReadOnly                  1
#define adLockPessimistic               2
#define adLockOptimistic                3
#define adLockBatchOptimistic           4

/* Field Types */
#define adEmpty                         0
#define adTinyInt                       16
#define adSmallInt                      2
#define adInteger                       3
#define adBigInt                        20
#define adUnsignedTinyInt               17
#define adUnsignedSmallInt              18
#define adUnsignedInt                   19
#define adUnsignedBigInt                21
#define adSingle                        4
#define adDouble                        5
#define adCurrency                      6
#define adDecimal                       14
#define adNumeric                       131
#define adBoolean                       11
#define adError                         10
#define adUserDefined                   132
#define adVariant                       12
#define adIDispatch                     9
#define adIUnknown                      13
#define adGUID                          72
#define adDate                          7
#define adDBDate                        133
#define adDBTime                        134
#define adDBTimeStamp                   135
#define adBSTR                          8
#define adChar                          129
#define adVarChar                       200
#define adLongVarChar                   201
#define adWChar                         130
#define adVarWChar                      202
#define adLongVarWChar                  203
#define adBinary                        128
#define adVarBinary                     204
#define adLongVarBinary                 205
#define adChapter                       136
#define adFileTime                      64
#define adPropVariant                   138
#define adVarNumeric                    139
#define adArray                         /* &H2000 */

#define adRecDeleted                    4 /*Indicates that the record was deleted.*/

#define adUseNone                       1
#define adUseServer                     2
#define adUseClient                     3
#define adUseClientBatch                3


#define adKeyForeign                    2

/* Constant Group: ObjectStateEnum */
#define adStateClosed                   0
#define adStateOpen                     1
#define adStateConnecting               2
#define adStateExecuting                4
#define adStateFetching                 8

/* Constant Group: SchemaEnum */
#define adSchemaProviderSpecific        -1
#define adSchemaAsserts                 0
#define adSchemaCatalogs                1
#define adSchemaCharacterSets           2
#define adSchemaCollations              3
#define adSchemaColumns                 4
#define adSchemaCheckConstraints        5
#define adSchemaConstraintColumnUsage   6
#define adSchemaConstraintTableUsage    7
#define adSchemaKeyColumnUsage          8
#define adSchemaReferentialContraints   9
#define adSchemaReferentialConstraints  9
#define adSchemaTableConstraints        10
#define adSchemaColumnsDomainUsage      11
#define adSchemaIndexes                 12
#define adSchemaColumnPrivileges        13
#define adSchemaTablePrivileges         14
#define adSchemaUsagePrivileges         15
#define adSchemaProcedures              16
#define adSchemaSchemata                17
#define adSchemaSQLLanguages            18
#define adSchemaStatistics              19
#define adSchemaTables                  20
#define adSchemaTranslations            21
#define adSchemaProviderTypes           22
#define adSchemaViews                   23
#define adSchemaViewColumnUsage         24
#define adSchemaViewTableUsage          25
#define adSchemaProcedureParameters     26
#define adSchemaForeignKeys             27
#define adSchemaPrimaryKeys             28
#define adSchemaProcedureColumns        29
#define adSchemaDBInfoKeywords          30
#define adSchemaDBInfoLiterals          31
#define adSchemaCubes                   32
#define adSchemaDimensions              33
#define adSchemaHierarchies             34
#define adSchemaLevels                  35
#define adSchemaMeasures                36
#define adSchemaProperties              37
#define adSchemaMembers                 38
#define adSchemaTrustees                39
#define adSchemaFunctions               40
#define adSchemaActions                 41
#define adSchemaCommands                42
#define adSchemaSets                    43

/* Constant Group: Supports */
#define adAddNew                        0x1000400 /* Supports the AddNew method to add new records. */
#define adApproxPosition                0x0004000 /* Supports the AbsolutePosition and AbsolutePage properties. */
#define adBookmark                      0x0002000 /* Supports the Bookmark property to gain access to specific records. */
#define adDelete                        0x1000800 /* Supports the Delete method to delete records. */
#define adFind                          0x0080000 /* Supports the Find method to locate a row in a Recordset. */
#define adHoldRecords                   0x0000100 /* Retrieves more records or changes the next position without committing all pending changes. */
#define adIndex                         0x0100000 /* Supports the Index property to name an index. */
#define adMovePrevious                  0x0000200 /* Supports the MoveFirst and MovePrevious methods, and Move or GetRows methods to move the current record position backward without requiring bookmarks. */
#define adNotify                        0x0040000 /* Indicates that the underlying data provider supports notifications (which determines whether Recordset events are supported). */
#define adResync                        0x0020000 /* Supports the Resync method to update the cursor with the data that is visible in the underlying database. */
#define adSeek                          0x0200000 /* Supports the Seek method to locate a row in a Recordset. */
#define adUpdate                        0x1008000 /* Supports the Update method to modify existing data. */
#define adUpdateBatch                   0x0010000

/* Command type */
#define adCmdUnspecified                -1
#define adCmdUnknown                    8
#define adCmdText                       1
#define adCmdTable                      2
#define adCmdStoredProc                 4
#define adCmdFile                       256
#define adCmdTableDirect                512

/* Execute type */
#define adAsyncExecute                  16  /* Indicates that the command should execute asynchronously.This value cannot be combined with the CommandTypeEnum value adCmdTableDirect.*/
#define adAsyncFetch                    32  /* Indicates that the remaining rows after the initial quantity specified in the CacheSize property should be retrieved asynchronously.*/
#define adAsyncFetchNonBlocking         64  /* Indicates that the main thread never blocks while retrieving. If the requested row has not been retrieved, the current row automatically moves to the end of the file.*/
                                            /* If you open a Recordset from a Stream containing a persistently stored Recordset, adAsyncFetchNonBlocking will not have an effect; the operation will be synchronous and blocking.*/
                                            /* adAsynchFetchNonBlocking has no effect when the adCmdTableDirect option is used to open the Recordset.*/
#define adExecuteNoRecords              128 /* Indicates that the command text is a command or stored procedure that does not return rows (for example, a command that only inserts data). If any rows are retrieved, they are discarded and not returned. */
                                            /*adExecuteNoRecords can only be passed as an optional parameter to the Command or Connection Execute method.*/
#define adExecuteStream                 256 /* Indicates that the results of a command execution should be returned as a stream.*/
                                            /*adExecuteStream can only be passed as an optional parameter to the Command Execute method.*/
#define adExecuteRecord                 512 /*Indicates that the CommandText is a command or stored procedure that returns a single row which should be returned as a Record object.*/
#define adOptionUnspecified             -1  /* Indicates that the command is unspecified.*/

/* Editmodes type */
#define adEditNone                      0  /* Indicates that no editing operation is in progress.*/
#define adEditInProgress                1 /*Indicates that data in the current record has been modified but not saved.*/
#define adEditAdd                       2 /*Indicates that the AddNew method has been called, and the current record in the copy buffer is a new record that has not been saved in the database.*/
#define adEditDelete                    4 /* Indicates that the current record has been deleted.*/

*/ Find modes */
#define adSearchBackward               -1
#define adSearchForward                 1

/* Transactions */
#define adXactAbortRetaining    262144 /*Performs retaining aborts by calling RollbackTrans to automatically start a new transaction. Not all providers support this behavior.*/
#define adXactCommitRetaining   131072 /*Performs retaining commits by calling CommitTrans to automatically start a new transaction. Not all providers support this behavior.*/

/*Position Enum*/
#define adPosBOF                -2 /*Indicates that the current record pointer is at BOF (that is, the BOF property is True).*/
#define adPosEOF                -3 /*Indicates that the current record pointer is at EOF (that is, the EOF property is True).*/
#define adPosUnknown            -1 /*Indicates that the Recordset is empty, the current position is unknown, or the provider does not support the AbsolutePage or AbsolutePosition property.*/

/*Resync AffectEnum */
#define adAffectAll               3  /* If there is not a Filter applied to the Recordset, affects all records.
                                        If the Filter property is set to a string criteria (such as "Author='Smith'"), then the operation affects visible records in the current chapter.
                                        If the Filter property is set to a member of the FilterGroupEnum or an array of bookmarks, then the operation will affect all rows of the Recordset.
                                        Note Note	adAffectAll is hidden in the Visual Basic Object Browser. */
#define adAffectAllChapters       4  /* Affects all records in all sibling chapters of the Recordset, including those not visible via any Filter that is currently applied.*/
#define adAffectCurrent           1  /* Affects only the current record.*/
#define adAffectGroup             2  /* Affects only records that satisfy the current Filter property setting. You must set the Filter property to a FilterGroupEnum value or an array of Bookmarks to use this option.*/

/* Resync Enum*/
#define adResyncAllValues         2  /* Default. Overwrites data, and pending updates are canceled.*/
#define adResyncUnderlyingValues  1  /* Does not overwrite data, and pending updates are not canceled.*/

/*Field Staus */
#define adFieldPendingChange      4 /* Indicates either that the field has been deleted and then re-added, perhaps with a different data type, or that the value of the field which previously had a status of adFieldOK has changed. The final form of the field will modify the Fields collection after the Update method is called./*

/*Save PersistFormatEnum */
#define adPersistADTG             0  /*This value indicates Microsoft Advanced Data TableGram (ADTG) format.*/
#define adPersistXML              1  /* This value indicates Extensible Markup Language (XML) format. */



//YOU CAN ALSO USE CTABLE@CON STRING
#command USE <(db)> [VIA <rdd>] [ALIAS <a>] [<nw: NEW>] ;
            [<ex: EXCLUSIVE>] [<sh: SHARED>] [<ro: READONLY>] ;
            [CODEPAGE <cp>] [INDEX <(index1)> [, <(indexN)>]] ;
            [ WHERE <cQuery> ]  =>;
         [ hb_adoSetQuery( <cQuery> ) ; ] ;
         dbUseArea( <.nw.>, <rdd>, <(db)>, <(a)>, ;
                    iif( <.sh.> .OR. <.ex.>, ! <.ex.>, NIL ), <.ro.> , [<cp>] ) ;
         [; dbSetIndex( <(index1)> )] ;
         [; dbSetIndex( <(indexN)> )]


/* sets for adordd */
//NOT NEEDED ANYMORE #command SET ADO TABLES INDEX LIST TO <array>  => ListIndex( <array>) /* defining index array list */

#command SET ADODBF TABLES INDEX LIST TO <array>  => ListDbfIndex( <array>) /* defining index array list with clipper expressions */

#command SET ADODBF MULTIBAG INDEX LIST TO <array>  => ListMultibagfIndex( <array>) /* defining index array list for all compound indexes */

#command SET ADODBF INDEX LIST FIELDTYPE NUMBER TO <array>  => ListFNumberIndex( <array>) /* defining numeric field len used in index expressions */

#command SET ADO TEMPORAY NAMES INDEX LIST TO <array>  => ListTmpNames( <array>) /* defining temporary index array list of names*/

#command SET ADO FIELDRECNO TABLES LIST TO <array>  => ListFieldRecno( <array>) /* defining temporary index array list of names*/

#command SET ADO DEFAULT RECNO FIELD TO <cname>  => ADODEFLDRECNO( <cname> ) /* defining the default name for id recno autoinc*/

#command SET ADO DEFAULT DATABASE TO <cDB> SERVER TO <cServer> [PORT TO <cPort>] ENGINE TO <cEngine> [USER TO <cUser>];
  [PASSWORD TO <cPass>] [CLASSNAME <cClass>]=> ADODEFAULTS( <cDB>, <cServer>, <cPort>, <cEngine>, <cUser>, <cPass>, <cClass>,.F.) /* defining the defaults for ado open and create*/

#command SET ADO LOCK CONTROL SHAREPATH TO <cPath> RDD TO <cRdd> => ADOLOCKCONTROL( <cPath>, <cRdd> ) /* defines path for table for record lock control D:\PATH */

#command SET ADO FORCE LOCK <x:ON,OFF>  => ADOFORCELOCKS( Upper( <(x)> ) == "ON" ) /* ADO locks files and records ?*/

#command SET ADO INDEX UDFS TO <array> => ListUdfs( <array> )

#command SET ADO DEFAULT DELETED FIELD TO <cname>  => ADODEFLDDELETED( <cname> ) /* defining the default name for DELETED field*/

#command SET ADO FIELDDELETED TABLES LIST TO <array>  => ListFieldDeleted( <array>) /* defining temporary Delete array list of names*/

#command SET ADO TABLES LOGICAL FIELDS LIST TO <array>  => ListFieldLogical( <array>) /* defining logical field array list of names*/

#command SET ADO TABLES DECIMAL FIELDS LIST TO <array>  => ListFieldDecimal( <array>) /* defining decimals by field array list of names*/

#command SET ADO TABLENAME WITH PATH <x:ON,OFF> => ADOTABLEWITHPATH( Upper( <(x)> ) == "ON" ) /* table name = path_tablename instead of only tablename */

//USE AS DEFAULT TO CACHE TABLES BASED ON THE NUMBER OF THE RECORDS OF THE RECORDSET
//SETTING IT TO 0 DOESNT CACHE ANY RECORDSETS
#command SET ADO CACHESIZE TO <nCache> ASYNC <x:ON,OFF> ASYNCNOWAIT <y:ON,OFF> => ADOPARAMETERS( <nCache>, Upper( <(x)> ) == "ON", Upper( <(y)> ) == "ON" )

//USED TO PREOPEN TABLES BASED ON THE NUMBER OF COUNTED RECORDS IN THE TABLE 
//SETTING IT TO 0 DOES NOT PREOPEN ANY TABLE BASED ON NR OF RECORDS
#command SET ADO PRE OPEN THRESHOLD TO <nRecords> [ MASK <aMask> ] => ADOPREOPENTHRESHOLD( <nRecords>, <aMask> )

#command SET ADO ROOT PATH TO <cNewPath> INSTEAD OF <cOldPath> => ADOROOTPATH( <cNewPath>, <cOldPath> )

#command SET RECORDSET OPEN WHERE CLAUSE TO <array> => ListTableQuery( <array> )  /*query to use when open each mentioned table*/

#endif
