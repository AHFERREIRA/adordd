# adordd
adordd for (x)Harbour

25.04.15

Most part of adordd its done!
Its working on trial phase in real app.

On trial new rdcored locking system to emulate completly as other rdds.
Absolutly no code change required.

Still some problems with NULL values and dates being corrected.

The speed can be increased with the use of indexes on the server side for the fields used as RECNO  by adorddd.
Configuration of cacheSize and MaxRecords can also has a huge impact of speed.

There are still some issues please read notes in tryadordd.prg.

