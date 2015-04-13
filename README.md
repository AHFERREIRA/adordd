# adordd
adordd for (x)Harbour

12.04.15

Most of many bugs found and corrected

Filters done!
Function FILTER2SQL parse 90% of usual filter expressions to SQL SELECTs with index order and for condition awareness.
Filter in adordd must be with cExpressions otherwise will raise error.
Vars in filters must be evaluated before, as adordd (SQL DB) does not know those vars.

Ex dbSetFilter( {|| &(cSearch) $ chave}, "'"+(cSearch)+"'"+' $ chave' )

index on the expressions must be:

Ex Index expression in (x)harbour array
  index on dtos(data)+str(val(cstr))
Ex Index expression in adordd array
  index on data+cstr

Missing features:

LOCATES - Milestone 15.4.15

Milestone end of April

CREATE
APPEND FROM
COPY TO
COPY STRUCT EXT
SORT
UPDATE
COUNT
TOTAL
SUM
AVERAGE
RESET SELECTS TO DEFAULT AFTER A SEEK WITH MORE THAN ONE FIELD IN SEEK EXPRESSION BY ISSUING
A GO TO NRECORD TO A RECNO OUT OF THE ACTUAL SELECT
EX:
nrecord := recno()
Seek "whatever"+"wahtever too" //2 fields
do while seek expression true
enddo
go to nrecord // comes back to select before the seek

CONCURRENT TABLE/RECORD ACESS (ready but not really tested)

We are getting there.
 
