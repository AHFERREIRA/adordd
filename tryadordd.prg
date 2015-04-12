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
	
	USE ....
	
	//try changing index focus
	BROWSE()
	
	//try set up filter
	
	BROWSE()
	
	//etc... or just linked to you app and check it.
	
RETURN nil	
	
	
	