 FUNCTION Main()
 
  RddRegister("ADORDD",1) 
  RddSetDefault("ADORDD") 

  //sql aware tables index expressions 
  //ATTENTION ALL MUST BE UPPERCASE
  //index files array needed for the adordd for your application
  //order expressions already translated to sql DONT FORGET TO replace taitional + sign with ,
  //ARRAY SPEC { {"TABLENAME",{"INDEXNAME","INDEXKEY","WHERE EXPRESSION AS USED FOR FOREXPRESSION","UNIQUE - DISTINCT ANY SQL STAT BEFORE * FROM"} }
  //temporary indexes are not included here they are create on fly and added to temindex list array
  //they are only valid through the duration of the application
  //the temp index name is auto given by adordd
  SET ADO TABLES INDEX LIST TO { {"ENCCLIST",{"ENCCPRO","NRENCOMEND,CODIGOPROD,ARMAZEM"},;
	   {"PROENCC","CODIGOPROD,ARMAZEM,NRENCOMEND"},{"MECPD","GUIA,CODCLIENTE,CODIGOPROD,ARMAZEM"},;
	   {"CDENCCL","CODCLIENTE"},{"PTFACT","NRFACTUR,CODCLIENTE,CODIGOPROD,ARMAZEM"},;
	   {"NRFPROD","NRFACTUR,CODIGOPROD,ARMAZEM,ANO ASC,SEMENTREGA ASC"} ,;
	   {"CCONTRT","CODIGOPROD,NRCONTRATO"}} }
	 
	//dbfs type index expression 
    //ATTENTION ALL MUST BE UPPERCASE
    //index files array needed for the adordd for your application 
    //order expressions as dbfs type
    //ARRAY SPEC { {"TABLENAME",{"INDEXNAME","INDEXKEY","FOR EXPRESSION","UNIQUE" } }
    //temporary indexes are not included here they are create on fly and added to temindex list array
    //they are only valid through the duration of the application
    //the temp index name is auto given by adordd	
    SET ADODBF TABLES INDEX LIST TO {  {"ENCCLIST",{"ENCCPRO","NRENCOMEND+CODIGOPROD+ARMAZEM"},;
     {"PROENCC","CODIGOPROD+ARMAZEM+NRENCOMEND"},{"MECPD","GUIA+CODCLIENTE+CODIGOPROD+ARMAZEM"},;
     {"CDENCCL","CODCLIENTE"},{"PTFACT","NRFACTUR+CODCLIENTE+CODIGOPROD+ARMAZEM"},;
     {"NRFPROD","NRFACTUR+CODIGOPROD+ARMAZEM+STR(VAL(ANO),4,0)+STR(VAL(SEMENTREGA),2,0)"} ,;
     {"CCONTRT","CODIGOPROD+NRCONTRATO"}} }

    //temporary index names
    SET ADO TEMPORAY NAMES INDEX LIST TO {"TMP","TEMP"}
	
	//each table autoinc field used as recno 
    SET ADO FIELDRECNO TABLES LIST TO {{"ENCCLIST","HBRECNO"},{"FACTURAS","HBRECNO"}}
	
	//default table autoinc field used as recno
    SET ADO DEFAULT RECNO FIELD TO "HBRECNO"

	SET AUTOPEN ON //might be OFF if you wish
	SET AUTORDER TO 1 // first index opened can be other
	
	//set default parameters to adordd if you do not USE COMMAND or dont pretend to include this info
	//set it here
	SET ADO DEFAULT DATABASE TO "d:\followup-testes\TESTES FOLLOWUP.add" SERVER TO "ADS" ENGINE TO "ADS" ;
	USER TO "adssys" PASSWORD TO ""

/*               THE ONLY CHANGES IN YOUR APP CODE END HERE! (SHOULD)              */


/*                                T R I A L S
   PEASE READ THIS CAREFULLY!
   
   PLEASE REMEMBER THAT ALTHOUGH ADORDD STILL AND MIGHT WORK WITHOUT ANY AUTOINC FIELD AS RECNO
   RESULT WILL BE UNPREDICTABLE IN SOME CIRCUNSTANCES OR ERROR MIGHT OCCOUR.
   THE FINAL RELEASE WIL NOT WORK WITHOUT SUCH A FIELD  
   
   INDEXES WITH DATES AND IN SOME BROWSE WITHIN A DATE SCOPE MOVEMENT HAS STILL SOME PROBLEMS   
   
   LOCATES HAVE TO BE CHANGED TO:
   
   EX:
   IF RDDSNAME() = "ADORDD"
      hb_adoSetLocateFor( "pessoa = "+"'"+(MeuNome())+"'")
	  locate for "pessoa = "+"'"+(MeuNome())+"'"
   ELSE	  
      locate for rtrim(MeuNome()) $ (caliasi)->pessoa
   ENDIF
   
   WHEN YOU DELETE A RECORD YOU CANT ACCESS IT ANYMORE. THUS CODE LIKE THIS IS ILLEGAL:

   DELETE RECORD
   BLANKREC

   THIS MUST BE CHANGED TO
   
   IF RDDSNAME() = "ADORDD"
      BLANKREC
      DELETE RECORD
   ELSE
      DELETE RECORD
      BLANKREC
   ENDIF
   
   FILTERS ARE REALLY SELECTS USING FILTER2SQL() IN ADORDD THAT ALLOW THE USE OF NORMAL DBF FILTERS
   THIS IS COFIGURED FOR ADS OR USUSAL SQL SINTAX BUT YOU CAN CONFIGURE IT FOR ANY DB 
   
   BESIDES THESE CHANGES APP SHOULD RUN WITHOUT ANY CODE LOGIC CHANGE

   PLEASE REPORT ANY BUGS! THANKS!   */	
   
	DBCREATE("database;table;dbengine;server;user;password",;
            {{"field1","+",10,0},;
             {"field2","N",10,3},;
             {"field3","D",8,0},;
             {"field4","L",1,0},;
             {"field5","N",10,0},;
             {"fieldmemo","M",10,0}},"ADORDD",.T.,"Trial") 
	
	APPEND BLANK
	REPLACE FIELD2 WITH ADORDD
	REPLACE FIELD3 WITH DATE()
	BROWSE()
	CLOSE TRIAL
	
	SELE 0
	/*
	WE DONT USE THIS COMMAND SO IT IS NOT TESTED IF YOU FIND SOME PROBLEM CHECK ADORDD.CH
	
    USE cTable VIA adordd ALIAS calias NEW SHARED CODEPAGE whatever INDEX cindex1;
      FROM DATABASE cDatabase FROM SERVER cServer QUERY "SELECT  * FROM "	USER admin PASSWORD admin
	  
	The tested way its to define ADO DEFAULTS see above and just call normal USE   
	*/
	
	SELE 1
	USE ....
	SELE 2
	USE ....
	SET INDEX TO
	SELE 1
	SET RELATION ....
	
	NREG  := RECNO()
	SEEK ???
	DO WHILE ???
	   MSGINFO(???)
	   DBSKIP()
	ENDDO
	
	//try changing index focus
	ORDSETFOCUS( ???? )
	BROWSE()
	
	//try set up filter
	SET FILTER TO ????
	BROWSE()
	
	//etc... or just linked adordd to you app and check it.
	
RETURN nil	
	
	
	