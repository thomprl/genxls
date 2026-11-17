GENXLS

This is a clone of the late Scott Klement's HSSF POI Service Program with a few extra features that have been added.  

This will create an XLSX IFS File based on a SQL Select Statement with some generic formatting and filtering auto applied.  

INSTALL

- Download the JAR Files and put them in a IFS Folder on your IBMi from the following URL. https://github.com/thomprl/genxls/releases/tag/1.0.0
- Download the source
- Modify the IFS location to find the JAR files in GENXLS.SQLRPGLE.  
- Create the IBMi Objects

```
CHGCURLIB CURLIB(yourlib)
CRTRPGMOD MODULE(HSSFR414) SRCFILE(QRPGLESRC) DBGVIEW(*LIST)
CRTSRVPGM SRVPGM(HSSFR414) MODULE(HSSFR414) ACTGRP(HSSFR414) 
               EXPORT(*SRCFILE) SRCFILE(QSRVSRC) 
               TEXT('Excel spreadsheet utilities - POI 3.14')
CRTBNDDIR BNDDIR(HSSF14)
ADDBNDDIRE BNDDIR(HSSF14) OBJ((HSSFR414)) 
CRTSQLRPGI OBJ(GENXLS) SRCMBR(*OBJ) COMMIT(*NONE) OBJTYPE(*PGM) 
               OPTION(*EVENTF) RPGPPOPT(*LVL2) DBGVIEW(*SOURCE)
               TEXT('Generate XLSX from SQL - POI 3.14')
CRTCMD CMD(GENXLS) PGM(GENXLS) TEXT('Generate XLSX from SQL - POI 3.14')
```

Usage:
```
 GENXLS TITLE(QCUSTCDT) SQLSTMT('select * from qiws.QCUSTCDT') IFSPTHFILE('/tmp/QCUSTCDT.xlsx')
```
