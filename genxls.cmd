             CMD        PROMPT('Generate XLSX From SQL')          
             PARM       KWD(FILETYPE) TYPE(*CHAR) LEN(4) RSTD(*YES) +  
                          DFT('XLSX') SPCVAL((XLSX) (CSV) (HTML) +     
                          (XLS)) PROMPT('File Type')                   
             PARM       KWD(TITLE) TYPE(*CHAR) LEN(100) MIN(1) +       
                          CASE(*MIXED) PROMPT('Title')                 
             PARM       KWD(SQLSTMT) TYPE(*CHAR) LEN(4096) MIN(1) +    
                          PROMPT('SQL Statement')                      
             PARM       KWD(IFSPTHFILE) TYPE(*CHAR) LEN(1024) MIN(1) + 
                          PROMPT('IFS Path and Filename')              
             PARM       KWD(RTNERR) TYPE(*CHAR) LEN(1) +               
                          PROMPT('Return Error')                       
             PARM       KWD(AUTOFILTER) TYPE(*LGL) RSTD(*YES) +        
                          DFT(*YES) SPCVAL((*YES '1') (*NO '0')) +     
                          PROMPT('Add auto filter (XLS/XLSX)')         
