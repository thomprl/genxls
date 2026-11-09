GENXLS

This is a clone of the late Scott Klement's HSSF POI Service Program with a few extra features that have been added.  

This will create an XLSX IFS File based on a SQL Select Statement with some generic formatting and filtering auto applied.  


INSTALL

- Download the JAR Files and put them in a IFS Folder on your IBMi from the following URL. https://github.com/thomprl/genxls/releases/tag/1.0.0
- Modify the IFS location to find the JAR files in GENXLS.SQLRPGLE.  
- Create the binding Directory
```
CRTBNDDIR HSSF14
QSYS/ADDBNDDIRE BNDDIR(HSSF14) OBJ((HSSFR414)) 
```
