**free
Ctl-Opt DftActGrp(*No) ActGrp(*Caller) OPTION(*SRCSTMT: *NODEBUGIO: *NOUNREF) Main(Main);
Ctl-Opt Thread(*SERIALIZE) BndDir('HSSF14');

// -----------------------------------------------------------------------
//
// Purpose: Run an SQL SELECT and write the result to an XLS, XLSX, CSV
//          or HTML file in the IFS. POI 3.14 version of GENXLS.
//
// Information was gathered from the following sites
// https://www.ibm.com/docs/en/i/7.4.0?topic=statements-get-descriptor
// https://www.itjungle.com/2016/07/19/fhg071916-story01/
// https://www.ibm.com/docs/en/i/7.4.0?topic=statements-get-descriptor#rbafzgetdescr__datatypecode
// http://www.easy400.net/hssfcgi/html/page5.htm
// https://www.itjungle.com/2013/08/21/fhg082113-story01/

Dcl-Pr Main ExtPgm('GENXLS');
  pIfsFileType CHAR(4) const;
  pTitle CHAR(100) const;
  pSqlStatement CHAR(4096) const;
  pIfsPathAndFilename CHAR(1024) const;
  pErrorCode CHAR(1) OPTIONS(*OMIT);
  pAutoFilter IND const OPTIONS(*NOPASS);
End-Pr;

Dcl-S vSqlStmt varchar(4096);
Dcl-S vForDescriptor varchar(20) INZ('vals');

Dcl-Ds t_dsColumnAttribute Template Qualified;
  columnName varchar(128);
  columnLabel varchar(128);
  columnDataType int(10);
  columnLength int(10);
  columnPrecision int(10);
  columnScale int(10);
  columnType char(10);
  columnDateTimeIntervalCode int(10);
  IncludeColumn IND;
End-Ds;

Dcl-Ds t_dsColumnValues Template Qualified;
  columnValue varchar(4000);
End-Ds;

Dcl-Ds t_dsColumnList Template Qualified;
  vNumberOfColumns int(10);
  dsColumnAttribute LIKEDS(t_dsColumnAttribute) DIM(300);
End-Ds;

Dcl-S gBook   like(SSWorkbook);
Dcl-S gSheet like(SSSheet);
Dcl-S gLgHeading   like(SSCellStyle);
Dcl-S gSmHeading   like(SSCellStyle);
Dcl-S gColHeading   like(SSCellStyle);
Dcl-S gNumeric   like(SSCellStyle);
Dcl-S gText   like(SSCellStyle);
Dcl-S gDates   like(SSCellStyle);
Dcl-S gNumericND like(SSCellStyle);
Dcl-S gXLSX IND;
Dcl-S gRowCount int(10) INZ(0);
Dcl-S gRow LIKE(SSRow);
Dcl-S gTitle CHAR(100);
Dcl-S gIfsPathAndFilename CHAR(1024);
Dcl-S gIfsFileType CHAR(4);
// Spreadsheet row of the column headings and the last column used
Dcl-S gHeaderRow INT(10);
Dcl-S gLastColumn INT(10) INZ(-1);
// Add filter drop-downs to the column headings (AUTOFILTER parameter)
Dcl-S gAutoFilter IND INZ(*ON);

// Return codes passed back in pErrorCode (RTNERR on the command)
Dcl-C RTN_OK '0';
Dcl-C RTN_ERROR '1';
Dcl-C RTN_NO_ROWS '2';
Dcl-S gReturnCode CHAR(1) INZ(RTN_OK);
// First error message, sent to the caller when the program ends
Dcl-S gErrorMessage VARCHAR(256);
// Data rows written to the output file
Dcl-S gRowsWritten INT(10);

// Column widths are in 1/256ths of a character: 200 characters
Dcl-C MAX_COLUMN_WIDTH 51200;
// Width of the filter arrow button: 3 characters
Dcl-C FILTER_ARROW_WIDTH 768;

// Constants for DB2 data types
Dcl-C DB2_SMALLINT 5;
Dcl-C DB2_INTEGER 4;
Dcl-C DB2_BIGINT 25;
Dcl-C DB2_DECIMAL 3;
Dcl-C DB2_NUMERIC 2;
Dcl-C DB2_REAL 7;
Dcl-C DB2_FLOAT 6;
Dcl-C DB2_DOUBLE 8;
Dcl-C DB2_FLOAT7 -360;
Dcl-C DB2_CHAR 1;
Dcl-C DB2_VARCHAR 12;
Dcl-C DB2_CLOB 40;
Dcl-C DB2_GRAPHIC -95;
Dcl-C DB2_VARGRAPHIC -96;
Dcl-C DB2_DBCLOB -350;
Dcl-C DB2_BINARY -2;
Dcl-C DB2_VARBINARY -3;
Dcl-C DB2_BLOB 30;
Dcl-C DB2_DATE 9;

/include *libl/qrpglesrc,hssf14_h

Dcl-Proc Main;
  Dcl-Pi *N;
    pIfsFileType CHAR(4) const;
    pTitle CHAR(100) const;
    pSqlStatement CHAR(4096) const;
    pIfsPathAndFilename CHAR(1024) const;
    pErrorCode CHAR(1) OPTIONS(*OMIT);
    pAutoFilter IND const OPTIONS(*NOPASS);
  End-Pi;

  Dcl-Ds dsColumnList Qualified;
    vNumberOfColumns int(10);
    dsColumnAttribute LIKEDS(t_dsColumnAttribute) DIM(300);
  End-Ds;

  Dcl-S vAbnormalEnd IND;

  Exec Sql
    Set Option Commit = *None, CloSQLCSR = *EndMod, DatFmt = *ISO, TIMFMT = *ISO;

  // Globals keep their values from the last call (linear-main program in
  // the caller's activation group), so start every call clean
  Reset gReturnCode;
  Reset gErrorMessage;
  Reset gRowsWritten;
  Reset gRowCount;
  Reset gHeaderRow;
  Reset gLastColumn;
  Reset gAutoFilter;

  // Auto filter is on unless the caller passes *NO
  If %parms >= %parmnum(pAutoFilter);
    gAutoFilter = pAutoFilter;
  EndIf;

  If pSqlStatement = '';
    SetError('SQL statement is required');
    Return;
  Else;
    vSqlStmt = pSqlStatement;
  EndIf;

  // Make sure it's a SELECT statement and not a Update/Insert/Delete
  // or something else that could be harmful
  If CheckSqlStatement(vSqlStmt);
    SetError('Invalid SQL statement; only a SELECT is allowed');
    *InLR = *On;
    Return;
  EndIf;

  SetJavaPath();

  gIfsFileType = %upper(pIfsFileType);
  gTitle = pTitle;
  gIfsPathAndFilename = pIfsPathAndFilename;

  // Make sure it's a valid type
  If not (gIfsFileType in %list('XLS' : 'XLSX' : 'CSV' : 'HTML'));
    SetError('Invalid file type ' + %trim(pIfsFileType)
        + '; use XLS, XLSX, CSV or HTML');
    Return;
  EndIf;

  // Get column attributes for the SQL statement
  dsColumnList = GetColumnAttributes(vSqlStmt);
  If gReturnCode = RTN_ERROR;
    Return;
  EndIf;

  // Get Row Data
  GetRowData(dsColumnList);

  // Runs on every return, including a crash
  On-Exit vAbnormalEnd;
    If vAbnormalEnd;
      SetError('GENXLS ended abnormally');
    EndIf;

    // Not checked: it fails harmlessly if the descriptor was never allocated
    Exec Sql
      Deallocate descriptor local :vForDescriptor;

    // Tell the caller how it went; from the command line this shows at
    // the bottom of the screen. On-Exit runs as its own procedure, so
    // *CALLER would be MAIN; '*PGMBDY' : 1 is whoever called the program.
    // Errors use CPF9898 (severity 40) so they reach the job log even when
    // the job only logs severity 30 and up
    Select;
      When gReturnCode = RTN_OK;
        Snd-Msg *COMP ('GENXLS wrote ' + %char(gRowsWritten) + ' rows to '
            + %trimr(gIfsPathAndFilename)) %target('*PGMBDY' : 1);
      When gReturnCode = RTN_NO_ROWS;
        Snd-Msg *INFO ('GENXLS: the SQL statement returned no rows; '
            + 'no file was created') %target('*PGMBDY' : 1);
      Other;
        Snd-Msg *DIAG %msg('CPF9898' : 'QCPFMSG' : 'GENXLS failed: '
            + gErrorMessage + '. See the job log for details')
            %target('*PGMBDY' : 1);
    EndSl;

    // RTNERR is a return value, so it's *NULL when the caller leaves it out
    If %addr(pErrorCode) <> *NULL;
      pErrorCode = gReturnCode;
    EndIf;

End-Proc;

Dcl-Proc GetColumnAttributes;
  Dcl-Pi *N LIKEDS(dsColumnList);
    pSqlStatement varchar(4096);
  End-Pi;

  Dcl-S vCount int(10);
  Dcl-S vIndex int(10);
  Dcl-S vColumnName varchar(128);
  Dcl-S vColumnLabel varchar(256);
  Dcl-S vColumnLength int(10);
  Dcl-S vColumnType int(10);
  Dcl-S vColumnPrecision int(10);
  Dcl-S vColumnScale int(10);
  Dcl-S vDateTimeIntervalCode int(10);

  Dcl-Ds dsColumnList Qualified;
    vNumberOfColumns int(10);
    dsColumnAttribute LIKEDS(t_dsColumnAttribute) DIM(300);
  End-Ds;

  Exec Sql PREPARE s1 FROM :pSqlStatement;
  If SqlFailed('PREPARE');
    Return dsColumnList;
  EndIf;

  Exec Sql
    ALLOCATE DESCRIPTOR local :vForDescriptor WITH MAX 300;   // 300 = max columns you expect
  If SqlFailed('ALLOCATE DESCRIPTOR');
    Return dsColumnList;
  EndIf;

  Exec Sql
    DESCRIBE s1 USING SQL DESCRIPTOR local :vForDescriptor;
  If SqlFailed('DESCRIBE');
    Return dsColumnList;
  EndIf;

  Exec Sql
    GET DESCRIPTOR :vForDescriptor :vCount = COUNT;
  If SqlFailed('GET DESCRIPTOR COUNT');
    Return dsColumnList;
  EndIf;

  dsColumnList.vNumberOfColumns = vCount;

  // Now loop through columns
  For vIndex = 1 to vCount;
    Exec Sql GET DESCRIPTOR :vForDescriptor VALUE :vIndex
        :vColumnName   = DB2_SYSTEM_COLUMN_NAME,
        :vColumnLabel = DB2_LABEL,
        :vColumnType = TYPE,
        :vColumnLength = LENGTH,
        :vColumnPrecision = PRECISION,
        :vColumnScale = SCALE,
        :vDateTimeIntervalCode = DATETIME_INTERVAL_CODE;
    If SqlFailed('GET DESCRIPTOR column ' + %char(vIndex));
      Return dsColumnList;
    EndIf;

    If vColumnName = ' ' and vColumnLabel <> ' ';
      vColumnName = vColumnLabel;
    ElseIf vColumnName <> ' ' and vColumnLabel = ' ';
      vColumnLabel = vColumnName;
    ElseIf vColumnName = ' ' and vColumnLabel = ' ';
      vColumnLabel = 'FIELD_' + %editc(vIndex:'Z');
      vColumnName = 'FIELD_' + %editc(vIndex:'Z');
    EndIf;

    dsColumnList.dsColumnAttribute(vIndex).columnName = vColumnName;
    dsColumnList.dsColumnAttribute(vIndex).columnLabel = SquishSpaces(vColumnLabel);
    dsColumnList.dsColumnAttribute(vIndex).columnDataType = vColumnType;
    dsColumnList.dsColumnAttribute(vIndex).columnLength = vColumnLength;
    dsColumnList.dsColumnAttribute(vIndex).columnPrecision = vColumnPrecision;
    dsColumnList.dsColumnAttribute(vIndex).columnScale = vColumnScale;
    dsColumnList.dsColumnAttribute(vIndex).columnDateTimeIntervalCode = vDateTimeIntervalCode;

    If vColumnType = DB2_CHAR or vColumnType = DB2_VARCHAR;
      dsColumnList.dsColumnAttribute(vIndex).columnType = 'CHAR';
      dsColumnList.dsColumnAttribute(vIndex).IncludeColumn = *On;
      Exec Sql
        Set Descriptor Local :vForDescriptor
        Value :vIndex
                TYPE = 12,
                DB2_CCSID = 37,
                LENGTH = 400;
    ElseIf vColumnType = DB2_DECIMAL or vColumnType = DB2_NUMERIC;
      dsColumnList.dsColumnAttribute(vIndex).columnType = 'DECIMAL';
      dsColumnList.dsColumnAttribute(vIndex).IncludeColumn = *On;
      Exec Sql
        Set Descriptor Local :vForDescriptor
        Value :vIndex
                TYPE = 12,
                DB2_CCSID = 37,
                LENGTH = 400;
    ElseIf vColumnType = DB2_SMALLINT or vColumnType = DB2_INTEGER or vColumnType = DB2_BIGINT;
      dsColumnList.dsColumnAttribute(vIndex).columnType = 'INTEGER';
      dsColumnList.dsColumnAttribute(vIndex).IncludeColumn = *On;
      Exec Sql
        Set Descriptor Local :vForDescriptor
        Value :vIndex
                TYPE = 12,
                DB2_CCSID = 37,
                LENGTH = 400;
    ElseIf vColumnType = DB2_DATE;
      // Date, time and timestamp all have TYPE 9. DESCRIBE already set them
      // up, and SET DESCRIPTOR with TYPE 9 and no LENGTH fails (SQL0804)
      dsColumnList.dsColumnAttribute(vIndex).columnType = 'DATE';
      dsColumnList.dsColumnAttribute(vIndex).IncludeColumn = *On;
    ElseIf vColumnType = DB2_REAL or vColumnType = DB2_FLOAT or vColumnType = DB2_DOUBLE
          or vColumnType = DB2_FLOAT7;
      dsColumnList.dsColumnAttribute(vIndex).columnType = 'FLOAT';
      dsColumnList.dsColumnAttribute(vIndex).IncludeColumn = *On;
      Exec Sql
        Set Descriptor Local :vForDescriptor
        Value :vIndex
                TYPE = 12,
                DB2_CCSID = 37,
                LENGTH = 400;
    ElseIf vColumnType = DB2_CLOB or vColumnType = DB2_DBCLOB;
      dsColumnList.dsColumnAttribute(vIndex).columnType = 'CLOB';
      dsColumnList.dsColumnAttribute(vIndex).IncludeColumn = *Off;
      Exec Sql
        Set Descriptor Local :vForDescriptor
        Value :vIndex
                TYPE = 12,
                DB2_CCSID = 37,
                LENGTH = 400;
    ElseIf vColumnType = DB2_GRAPHIC or vColumnType = DB2_VARGRAPHIC;
      dsColumnList.dsColumnAttribute(vIndex).columnType = 'GRAPHIC';
      dsColumnList.dsColumnAttribute(vIndex).IncludeColumn = *Off;
      Exec Sql
        Set Descriptor Local :vForDescriptor
        Value :vIndex
                TYPE = 12,
                DB2_CCSID = 37,
                LENGTH = 400;
    ElseIf vColumnType = DB2_BINARY or vColumnType = DB2_VARBINARY or vColumnType = DB2_BLOB;
      dsColumnList.dsColumnAttribute(vIndex).columnType = 'BINARY';
      dsColumnList.dsColumnAttribute(vIndex).IncludeColumn = *Off;
      Exec Sql
        Set Descriptor Local :vForDescriptor
        Value :vIndex
                TYPE = 12,
                DB2_CCSID = 37,
                LENGTH = 400;
    Else;
      dsColumnList.dsColumnAttribute(vIndex).columnType = 'OTHER';
      dsColumnList.dsColumnAttribute(vIndex).IncludeColumn = *Off;
      Exec Sql
        Set Descriptor Local :vForDescriptor
        Value :vIndex
                TYPE = 12,
                DB2_CCSID = 37,
                LENGTH = 400;
    EndIf;
    // Each branch above except DATE runs one SET DESCRIPTOR; for DATE this
    // checks the GET DESCRIPTOR again, which already passed
    If SqlFailed('SET DESCRIPTOR column ' + %char(vIndex));
      Return dsColumnList;
    EndIf;

  EndFor;
  Return dsColumnList;
End-Proc;

///
// SqlFailed
// Check SQLCOD after an SQL statement. On an error, write a diagnostic
// message to the job log and set the return code to RTN_ERROR.
// @param  pAction Short description of the statement that just ran.
// @return *ON if the statement failed, *OFF if it worked.
///
Dcl-Proc SqlFailed;
  Dcl-Pi *N IND;
    pAction VARCHAR(50) CONST;
  End-Pi;

  If SQLCOD >= 0;
    Return *OFF;
  EndIf;

  // CPF9898 (severity 40) so it reaches the job log
  Snd-Msg *DIAG %msg('CPF9898' : 'QCPFMSG' : 'GENXLS SQL error '
      + %char(SQLCOD) + ' (SQLSTATE ' + SQLSTT + ') on ' + pAction);
  SetError('SQL error ' + %char(SQLCOD) + ' on ' + pAction);
  Return *ON;
End-Proc SqlFailed;

///
// SetError
// Set the return code to RTN_ERROR and keep the message for the caller.
// Only the first error is kept, because later ones usually follow from it.
// @param  pMessage Short description of what went wrong.
// @return None
///
Dcl-Proc SetError;
  Dcl-Pi *N;
    pMessage VARCHAR(256) CONST;
  End-Pi;

  If gReturnCode <> RTN_ERROR;
    gReturnCode = RTN_ERROR;
    gErrorMessage = pMessage;
  EndIf;
End-Proc SetError;

///
// SquishSpaces
// Trim the text and replace each run of blanks between words with one blank.
// Column headings (DB2_LABEL) come back as 20-byte segments, so a heading
// like 'Comm Pay' + 'Grid' arrives as 'Comm Pay            Grid'.
// @param  pText Text to clean up.
// @return The trimmed text with single blanks between words.
///
Dcl-Proc SquishSpaces;
  Dcl-Pi *N VARCHAR(256);
    pText VARCHAR(256) CONST;
  End-Pi;

  Dcl-S vText VARCHAR(256);

  vText = %trim(pText);

  Dow %scan('  ' : vText) > 0;
    vText = %scanrpl('  ' : ' ' : vText);
  EndDo;

  Return vText;
End-Proc SquishSpaces;

///
// EscapeText
// Escape text for the output file type in gIfsFileType.
// CSV doubles each quote. HTML replaces & " < > with entities.
// Other file types get the text back unchanged.
// @param  pText Text to escape.
// @return The escaped text.
///
Dcl-Proc EscapeText;
  Dcl-Pi *N VARCHAR(16000);
    pText VARCHAR(4000) CONST;
  End-Pi;

  Dcl-S vText VARCHAR(16000);

  vText = pText;

  If gIfsFileType = 'CSV';
    vText = %scanrpl('"' : '""' : vText);
  ElseIf gIfsFileType = 'HTML';
    // & first, so the & in the other entities isn't escaped again
    vText = %scanrpl('&' : '&amp;' : vText);
    vText = %scanrpl('"' : '&quot;' : vText);
    vText = %scanrpl('<' : '&lt;' : vText);
    vText = %scanrpl('>' : '&gt;' : vText);
  EndIf;

  Return vText;
End-Proc EscapeText;

// Make sure it's a SELECT statement and not a Update/Insert/Delete
// or something else that could be harmful
Dcl-Proc CheckSqlStatement;
  Dcl-Pi *N IND;
    pSqlStatement varchar(4096) const;
  End-Pi;
  Dcl-S vError IND INZ(*OFF);

  Exec Sql
    SELECT 1
      INTO :vError
      FROM TABLE(QSYS2.PARSE_STATEMENT(:pSqlStatement)) AS A
    WHERE SQL_STATEMENT_TYPE <> 'QUERY' FETCH FIRST ROW ONLY;

  If vError or SqlCod < 0;
    Return *On;
  Else;
    Return *Off;
  EndIf;

End-Proc;

// Get Row Data and create Excel file
Dcl-Proc GetRowData;
  Dcl-Pi *N;
    pColumnList LikeDS(t_dsColumnList) const;
  End-Pi;

  Dcl-Ds dsColumnValues LIKEDS(t_dsColumnValues) DIM(300);
  Dcl-S vIndex INT(10);
  Dcl-S vColumnValue LIKE(dsColumnValues.columnValue);
  Dcl-S vFirstTime IND INZ(*On);
  Dcl-S vColumnCount INT(10);
  Dcl-S vWrkDecimalField PACKED(26:8);
  Dcl-S vWrkNoDecimalField PACKED(26:0);
  Dcl-S vWrkDate DATE;
  Dcl-S vWrkTime TIME;
  Dcl-S vWrkTimestamp TIMESTAMP;
  Dcl-S vDetailString VARCHAR(32000);

  Exec Sql
    Declare rowDataCsr cursor for s1;

  Exec Sql
    Open rowDataCsr;

  If sqlcod = -502;
    Exec Sql
      Close rowDataCsr;
    Exec Sql
      Open rowDataCsr;
  EndIf;
  If SqlFailed('OPEN cursor');
    Return;
  EndIf;

  Dow 1=1;
    Exec Sql
    Fetch From rowDataCsr
      Into SQL descriptor :vForDescriptor;

    If sqlcod < 0 or sqlcod = 100;
      If sqlcod < 0;
        SqlFailed('FETCH');
      EndIf;
      Leave;
    EndIf;

    If vFirstTime;
      vFirstTime = *Off;

      // Records found let's generate the spreadsheet
      If gIfsFileType In %list('XLS' : 'XLSX');
        // If it's XLS or XLSX Create Spreadsheet and write column headers
        CreateSpreadsheet();
        CreateSpreadsheetHeaderRow(pColumnList);
      ElseIf gIfsFileType = 'CSV';
        // If it's CSV Create a CSV Header Row
        CreateCSVHeaderRow(pColumnList);
      ElseIf gIfsFileType = 'HTML';
        // Implement HTML generation logic her
        CreateHTMLHeaderRow(pColumnList);
      EndIf;

      // The CSV or HTML header write failed
      If gReturnCode = RTN_ERROR;
        Leave;
      EndIf;
    EndIf;

    // Process Row Data here - for example, write to Excel file
    Reset dsColumnValues;

    gRowCount += 1;
    If gIfsFileType In %list('XLS' : 'XLSX');
      gRow = SSSheet_createRow(gSheet:gRowCount);
    EndIf;

    vColumnCount = 0;
    vDetailString = '';

    For vIndex = 1 to pColumnList.vNumberOfColumns;
      If not pColumnList.dsColumnAttribute(vIndex).IncludeColumn;
        Iter;
      EndIf;

      Reset vColumnValue;
      Reset vWrkDate;
      Reset vWrkTime;
      Reset vWrkTimestamp;

      If pColumnList.dsColumnAttribute(vIndex).columnType = 'DATE';
        If pColumnList.dsColumnAttribute(vIndex).columnDateTimeIntervalCode = 1;
          Exec Sql
            GET DESCRIPTOR :vForDescriptor
              value :vIndex :vWrkDate = DATA;
          vColumnValue = %char(vWrkDate);
        ElseIf pColumnList.dsColumnAttribute(vIndex).columnDateTimeIntervalCode = 2;
          Exec Sql
            GET DESCRIPTOR :vForDescriptor
              value :vIndex :vWrkTime = DATA;
          vColumnValue = %char(vWrkTime);
        ElseIf pColumnList.dsColumnAttribute(vIndex).columnDateTimeIntervalCode = 3;
          Exec Sql
            GET DESCRIPTOR :vForDescriptor
              value :vIndex :vWrkTimestamp = DATA;
          vColumnValue = %char(vWrkTimestamp);
        EndIf;
      Else;
        Exec Sql
          GET DESCRIPTOR :vForDescriptor
            value :vIndex :vColumnValue = DATA;
      EndIf;
      // Each branch above runs one GET DESCRIPTOR
      If SqlFailed('GET DESCRIPTOR DATA column ' + %char(vIndex));
        Leave;
      EndIf;

      dsColumnValues(vIndex).columnValue = vColumnValue;

      If vColumnCount > 0 and gIfsFileType = 'CSV';
        vDetailString += ',';
      EndIf;
      If vColumnCount = 0 and gIfsFileType = 'HTML';
        vDetailString += '<tr>';
      EndIf;

      If pColumnList.dsColumnAttribute(vIndex).columnType = 'CHAR';
        If gIfsFileType In %list('XLS' : 'XLSX');
          ss_text( gRow: vColumnCount : %trim(vColumnValue)   : gText);
        ElseIf gIfsFileType = 'CSV';
          vDetailString += '"' + EscapeText(%trim(vColumnValue)) + '"';
        ElseIf gIfsFileType = 'HTML';
          vDetailString += '<td>' + EscapeText(%trim(vColumnValue)) + '</td>';
        EndIf;

      ElseIf pColumnList.dsColumnAttribute(vIndex).columnType = 'DATE';
        If gIfsFileType In %list('XLS' : 'XLSX');
          ss_text( gRow: vColumnCount : %trim(vColumnValue)   : gText);
        ElseIf gIfsFileType = 'CSV';
          vDetailString += '"' + EscapeText(%trim(vColumnValue)) + '"';
        ElseIf gIfsFileType = 'HTML';
          vDetailString += '<td>' + EscapeText(%trim(vColumnValue)) + '</td>';
        EndIf;

      ElseIf pColumnList.dsColumnAttribute(vIndex).columnType = 'INTEGER';
        Monitor;
          vWrkNoDecimalField = %dec(vColumnValue: 26: 0);
        On-Error;
          vWrkNoDecimalField = 0;
        EndMon;
        If gIfsFileType In %list('XLS' : 'XLSX');
          ss_num ( gRow: vColumnCount : vWrkNoDecimalField : gNumericND);
        ElseIf gIfsFileType = 'CSV';
          vDetailString +=  %char(vWrkNoDecimalField);
        ElseIf gIfsFileType = 'HTML';
          vDetailString += '<td>' + %char(vWrkNoDecimalField) + '</td>';
        EndIf;

      ElseIf (pColumnList.dsColumnAttribute(vIndex).columnType = 'DECIMAL' or
            pColumnList.dsColumnAttribute(vIndex).columnType = 'FLOAT') and
            pColumnList.dsColumnAttribute(vIndex).columnScale <> 0;
        Monitor;
          // FLOAT values come back in E notation (1.2345E2), which %dec can't read
          If pColumnList.dsColumnAttribute(vIndex).columnType = 'FLOAT';
            vWrkDecimalField = %dech(%float(vColumnValue) : 26 : 8);
          Else;
            vWrkDecimalField = %dec(vColumnValue: 26: 8);
          EndIf;
        On-Error;
          vWrkDecimalField = 0;
        EndMon;
        If gIfsFileType In %list('XLS' : 'XLSX');
          ss_num ( gRow: vColumnCount : vWrkDecimalField : gNumeric);
        ElseIf gIfsFileType = 'CSV';
          vDetailString +=  %trim(%editc(vWrkDecimalField:'P'));
        ElseIf gIfsFileType = 'HTML';
          vDetailString += '<td>' + %trim(%editc(vWrkDecimalField:'P')) + '</td>';
        EndIf;

      ElseIf (pColumnList.dsColumnAttribute(vIndex).columnType = 'DECIMAL' or
            pColumnList.dsColumnAttribute(vIndex).columnType = 'FLOAT') and
            pColumnList.dsColumnAttribute(vIndex).columnScale = 0;
        Monitor;
          // FLOAT values come back in E notation (1.2345E2), which %dec can't read
          If pColumnList.dsColumnAttribute(vIndex).columnType = 'FLOAT';
            vWrkDecimalField = %dech(%float(vColumnValue) : 26 : 8);
          Else;
            vWrkDecimalField = %dec(vColumnValue: 26: 8);
          EndIf;
        On-Error;
          vWrkDecimalField = 0;
        EndMon;
        If gIfsFileType In %list('XLS' : 'XLSX');
          ss_num ( gRow: vColumnCount : vWrkDecimalField : gNumericND);
        ElseIf gIfsFileType = 'CSV';
          vDetailString +=  %trim(%editc(vWrkDecimalField:'P'));
        ElseIf gIfsFileType = 'HTML';
          vDetailString += '<td>' + %trim(%editc(vWrkDecimalField:'P')) + '</td>';
        EndIf;
      EndIf;

      vColumnCount += 1;
    EndFor;

    // A GET DESCRIPTOR in the column loop failed
    If gReturnCode = RTN_ERROR;
      Leave;
    EndIf;

    If gIfsFileType = 'HTML';
      vDetailString += '</tr>';
    EndIf;

    If gIfsFileType = 'CSV' or gIfsFileType = 'HTML';
      // Write detail row to CSV or HTML file
      Exec Sql
        CALL QSYS2.IFS_WRITE_UTF8(rtrim(:gIfsPathAndFilename),
                            rtrim(:vDetailString),
                            OVERWRITE => 'APPEND',
                            END_OF_LINE => 'CRLF');
      If SqlFailed('IFS_WRITE_UTF8 row');
        Leave;
      EndIf;
    EndIf;

    gRowsWritten += 1;
  EndDo;

  Exec Sql
    Close rowDataCsr;
  SqlFailed('CLOSE cursor');

  If gReturnCode = RTN_ERROR;
    // Don't save a partial spreadsheet
    If gIfsFileType In %list('XLS' : 'XLSX') and Not vFirstTime;
      ss_end_object_group();
    EndIf;
    Return;
  EndIf;

  // No rows: no file is written
  If vFirstTime;
    gReturnCode = RTN_NO_ROWS;
    Return;
  EndIf;

  If gIfsFileType = 'HTML';
    // vDetailString still holds the last row, so start it over
    vDetailString = '</table>';
    // Write end of HTML Table
    Exec Sql
      CALL QSYS2.IFS_WRITE_UTF8(rtrim(:gIfsPathAndFilename),
                          rtrim(:vDetailString),
                          OVERWRITE => 'APPEND',
                          END_OF_LINE => 'CRLF');
    SqlFailed('IFS_WRITE_UTF8 end of table');
  EndIf;

  If gIfsFileType In %list('XLS' : 'XLSX');
    FinishSpreadsheet();
    // A bad path or missing authority shows up here as a Java exception
    Monitor;
      ss_save(gBook: %trimr(gIfsPathAndFilename));
    On-Error;
      SetError('Could not save ' + %trimr(gIfsPathAndFilename));
    EndMon;
    ss_end_object_group();
  EndIf;

End-Proc;

// Implement logic to create Excel file and write column headers based on dsColumnList
Dcl-Proc CreateSpreadsheet;

  Dcl-S vColumnCount INT(10);

  ss_begin_object_group(100);

  If %upper(gIfsFileType) = 'XLSX';
    gBook = new_XSSFWorkbook();
  ElseIf %upper(gIfsFileType) = 'XLS';
    gBook = new_HSSFWorkbook();
  EndIf;

  gSheet = ss_NewSheet(gBook: 'Sheet1');
  // Create Cell Styles
  CreateSpreadsheetCellStyles();
End-Proc;

Dcl-Proc CreateSpreadsheetHeaderRow;
  Dcl-Pi *N;
    pColumnList LikeDS(t_dsColumnList) const;
  End-Pi;

  Dcl-S vIndex INT(10);
  Dcl-S vPrintSetup LIKE(SSPrintSetup);
  Dcl-S vColumnCount INT(10);
  Dcl-S vColumnWidth INT(10);
  Dcl-S vTitleRange LIKE(CellRangeAddress);

  Dcl-C TRUE '1';
  Dcl-C FALSE '0';

  gRowCount = 0;

  If pColumnList.vNumberOfColumns = 0;
    Return;
  Else;
    gRow = SSSheet_createRow(gSheet:gRowCount);
  EndIf;

  // Create header row with column labels
  ss_text( gRow : 0
        : %trimr(gTitle):  gLgHeading);

  gRowCount += 2;
  gRow = SSSheet_createRow(gSheet:gRowCount);
  gHeaderRow = gRowCount;

  For vIndex = 1 to pColumnList.vNumberOfColumns;

    If not pColumnList.dsColumnAttribute(vIndex).IncludeColumn;
      Iter;
    EndIf;

    // If Column Label is longer than Column Length, use Column Length to set column width
    If %len(%trim(pColumnList.dsColumnAttribute(vIndex).columnLabel)) >
            pColumnList.dsColumnAttribute(vIndex).columnLength;
      // Use Column Label length to set column width
      vColumnWidth =
          %len(%trimr(pColumnList.dsColumnAttribute(vIndex).columnLabel)) * 356;
      // Else use Column Length to set column width
    Else;
      // Use Column Length to set column width
      vColumnWidth = pColumnList.dsColumnAttribute(vIndex).columnLength * 356;
    EndIf;

    // Excel fails above 255 characters, so never go above 200
    If vColumnWidth > MAX_COLUMN_WIDTH;
      vColumnWidth = MAX_COLUMN_WIDTH;
    EndIf;
    SSSheet_setColumnWidth(gSheet : vColumnCount : vColumnWidth);

    // Create header row with column labels
    ss_text( gRow : vColumnCount
        : %trimr(pColumnList.dsColumnAttribute(vIndex).columnLabel):  gColHeading);

    vColumnCount += 1;
  EndFor;

  gLastColumn = vColumnCount - 1;

  // Merge the title across the columns, so auto-sizing column A
  // doesn't make it as wide as the title
  If gLastColumn > 0;
    vTitleRange = new_CellRangeAddress(0 : 0 : 0 : gLastColumn);
    SSSheet_addMergedRegion(gSheet : vTitleRange);
    ss_freeLocalRef(vTitleRange);
  EndIf;

  // Setup PageSetup Properties
  vPrintSetup = SSSheet_getPrintSetup(gSheet);
  SSPrintSetup_setLandscape(vPrintSetup: TRUE);
  SSPrintSetup_setFitWidth(vPrintSetup: 1);
  SSPrintSetup_setFitHeight(vPrintSetup: 0);
  // Set the top line to repeat on each printed page
  SS_setRepeating(gBook: gSheet: -1: -1: 0: 0 );
  // Set the page to Auto Page Break
  SSSheet_setAutobreaks(gSheet: TRUE);

  // Freeze the top header row when viewing the spreadsheet
  SSSheet_createFreezePane(gSheet: 0: 3: 0: 3);
  // Set the footers to print the date, page numbers and filename
  SS_footer_setRight (gSheet: SS_footer_file());
  SS_footer_setCenter (gSheet: 'Page ' + SS_footer_page()
                                    +  ' of ' + SS_footer_numPages());
  SS_footer_setLeft  (gSheet: 'Run Date: ' + %editc(udate:'Y'));

End-Proc;

///
// FinishSpreadsheet
// Add filter drop-downs to the column heading row (when gAutoFilter is
// on) and size each column to fit the heading and the first
// AUTOSIZE_ROWS data rows. Call after the last data row, before ss_save.
// If POI fails, the columns keep the widths set with the heading row.
// @return None
///
Dcl-Proc FinishSpreadsheet;
  Dcl-S vColumn INT(10);
  Dcl-S vColumnWidth INT(10);
  Dcl-S vLastRow INT(10);
  Dcl-S vHeadingChars FLOAT(8);
  Dcl-S vDataChars FLOAT(8);
  Dcl-S vRange LIKE(CellRangeAddress);
  Dcl-S vAutoFilter LIKE(SSAutoFilter);

  // Data rows to read when sizing columns; reading every row is slow
  Dcl-C AUTOSIZE_ROWS 20;

  If gLastColumn < 0;
    Return;
  EndIf;

  // Filter drop-downs on the heading row, covering every data row
  If gAutoFilter;
    Monitor;
      vRange = new_CellRangeAddress(gHeaderRow : gRowCount : 0 : gLastColumn);
      vAutoFilter = SSSheet_setAutoFilter(gSheet : vRange);
      ss_freeLocalRef(vAutoFilter);
      ss_freeLocalRef(vRange);
    On-Error;
    EndMon;
  EndIf;

  // Size each column to fit the wider of its heading and the first
  // AUTOSIZE_ROWS data rows
  vLastRow = gHeaderRow + AUTOSIZE_ROWS;
  If vLastRow > gRowCount;
    vLastRow = gRowCount;
  EndIf;

  For vColumn = 0 to gLastColumn;
    Monitor;
      // -1 means no cells in those rows
      vHeadingChars = SheetUtil_getColumnWidth(gSheet : vColumn : *OFF
          : gHeaderRow : gHeaderRow);
      vDataChars = -1;
      If vLastRow > gHeaderRow;
        vDataChars = SheetUtil_getColumnWidth(gSheet : vColumn : *OFF
            : gHeaderRow + 1 : vLastRow);
      EndIf;

      // The heading is centered, so leave room for the filter arrow on
      // both sides or the right end of the heading goes under the arrow
      vColumnWidth = 0;
      If vHeadingChars > 0;
        vColumnWidth = %inth(vHeadingChars * 256);
        If gAutoFilter;
          vColumnWidth += FILTER_ARROW_WIDTH * 2;
        EndIf;
      EndIf;
      If vDataChars > 0 and %inth(vDataChars * 256) > vColumnWidth;
        vColumnWidth = %inth(vDataChars * 256);
      EndIf;

      // 0 means no cells to measure: keep the current width
      If vColumnWidth > 0;
        If vColumnWidth > MAX_COLUMN_WIDTH;
          vColumnWidth = MAX_COLUMN_WIDTH;
        EndIf;
        SSSheet_setColumnWidth(gSheet : vColumn : vColumnWidth);
      EndIf;
    On-Error;
    EndMon;
  EndFor;
End-Proc FinishSpreadsheet;

Dcl-Proc CreateCSVHeaderRow;
  Dcl-Pi *N;
    pColumnList LikeDS(t_dsColumnList) const;
  End-Pi;

  Dcl-S vIndex INT(10);
  Dcl-S vFirstTime IND INZ(*On);
  Dcl-S vHeaderRow VARCHAR(32000);

  For vIndex = 1 to pColumnList.vNumberOfColumns;

    If not pColumnList.dsColumnAttribute(vIndex).IncludeColumn;
      Iter;
    EndIf;

    If vFirstTime;
      vHeaderRow = '"'
          + EscapeText(%trimr(pColumnList.dsColumnAttribute(vIndex).columnLabel))
          + '"';
      vFirstTime = *Off;
    Else;
      vHeaderRow += ',' + '"'
          + EscapeText(%trimr(pColumnList.dsColumnAttribute(vIndex).columnLabel))
          + '"';
    EndIf;

  EndFor;

  // Write header row to CSV file
  Exec Sql
    CALL QSYS2.IFS_WRITE_UTF8(rtrim(:gIfsPathAndFilename),
                        rtrim(:vHeaderRow),
                        OVERWRITE => 'REPLACE',
                        END_OF_LINE => 'CRLF');
  SqlFailed('IFS_WRITE_UTF8 CSV header');

End-Proc;

Dcl-Proc CreateHTMLHeaderRow;
  Dcl-Pi *N;
    pColumnList LikeDS(t_dsColumnList) const;
  End-Pi;

  Dcl-S vIndex INT(10);
  Dcl-S vFirstTime IND INZ(*On);
  Dcl-S vHeaderRow VARCHAR(32000);

  For vIndex = 1 to pColumnList.vNumberOfColumns;

    If not pColumnList.dsColumnAttribute(vIndex).IncludeColumn;
      Iter;
    EndIf;

    If vFirstTime;
      vHeaderRow = '<table border="1"><tr>';
      vFirstTime = *Off;
    EndIf;

    vHeaderRow += '<th><b>'
        + EscapeText(%trimr(pColumnList.dsColumnAttribute(vIndex).columnLabel))
        + '</b></th>';

  EndFor;

  vHeaderRow += '</tr>';

  // Write header row to HTML file
  Exec Sql
    CALL QSYS2.IFS_WRITE_UTF8(rtrim(:gIfsPathAndFilename),
                        rtrim(:vHeaderRow),
                        OVERWRITE => 'REPLACE',
                        END_OF_LINE => 'CRLF');
  SqlFailed('IFS_WRITE_UTF8 HTML header');

End-Proc;



// CreateCellStyles(): Create the different display styles
//     used for cells in this Excel workbook.
//
// NOTE: Uses the following global variables:
//        Book, LgHeading, SmHeading, ColHeading, Numeric
//        Text, Dates.
Dcl-Proc CreateSpreadsheetCellStyles;

  Dcl-S vLgFont like(SSFont);
  Dcl-S vSmFont like(SSFont);
  Dcl-S vChFont like(SSFont);
  Dcl-S vDataFmt like(SSDataFormat);
  Dcl-S vTempStr like(jstring);
  Dcl-S vNumFmt int(5);
  Dcl-S vDateFmt int(5);

  //
  //  Create a cell style for the large, centered
  //  title at the top of the report.
  //
  gLgHeading = SSWorkbook_createCellStyle(gBook);

  vLgFont = SSWorkbook_createFont(gBook);
  SSFont_setBoldweight(vLgFont: BOLDWEIGHT_BOLD);
  SSFont_setFontHeightInPoints(vLgFont: 16);
  SSCellStyle_setFont(gLgHeading: vLgFont);

  SSCellStyle_setAlignment(gLgHeading: ALIGN_LEFT);

  //
  // Create a cell style for the smaller text that
  //  will be printed below the main heading.
  //

  gSmHeading = SSWorkbook_createCellStyle(gBook);

  vSmFont = SSWorkbook_createFont(gBook);
  SSFont_setFontHeightInPoints(vSmFont: 8);
  SSCellStyle_setFont(gSmHeading: vSmFont);

  SSCellStyle_setAlignment(gSmHeading: ALIGN_LEFT);

  //
  // Create a cell style for the column headings.
  // These are bold and have a border line at the bottom
  //

  gColHeading = SSWorkbook_createCellStyle(gBook);

  vChFont = SSWorkbook_createFont(gBook);
  SSFont_setBoldweight(vChFont: BOLDWEIGHT_BOLD);
  SSCellStyle_setFont(gColHeading: vChFont);

  SSCellStyle_setAlignment(gColHeading: ALIGN_CENTER);
  SSCellStyle_setBorderBottom(gColHeading: BORDER_THIN);
  SSCellStyle_setBorderTop(gColHeading: BORDER_THIN);
  SSCellStyle_setBorderLeft(gColHeading: BORDER_THIN);
  SSCellStyle_setBorderRight(gColHeading: BORDER_THIN);
  SSCellStyle_setFillForegroundColor(gColHeading: COLOR_YELLOW);
  ssCellStyle_setFillPattern(gColHeading : SS_PATTERN_SOLID_FOREGROUND);

  //
  // Create a cell style for numbers so that they are
  //  right-aligned and the number is formatted nicely.
  //

  gNumeric = SSWorkbook_createCellStyle(gBook);

  vDataFmt = SSWorkbook_createDataFormat(gBook);
  vTempStr = new_String('#,##0.00');
  vNumFmt  = SSDataFormat_getFormat(vDataFmt: vTempStr);
  SSCellStyle_setDataFormat(gNumeric: vNumFmt);
  SSCellStyle_setBorderBottom(gNumeric : BORDER_THIN);
  SSCellStyle_setBorderLeft(gNumeric : BORDER_THIN);
  SSCellStyle_setBorderRight(gNumeric : BORDER_THIN);
  SSCellStyle_setBorderTop(gNumeric : BORDER_THIN);
  SSCellStyle_setAlignment(gNumeric: ALIGN_RIGHT);

  gNumericND = SSWorkbook_createCellStyle(gBook);
  vDataFmt = SSWorkbook_createDataFormat(gBook);
  // vTempStr = new_String('#,##0');
  // vNumFmt = SSDataFormat_getFormat(vDataFmt: vTempStr);
  // SSCellStyle_setDataFormat(gNumericND: vNumFmt);
  SSCellStyle_setAlignment(gNumericND: ALIGN_RIGHT);
  SSCellStyle_setBorderBottom(gNumericND: BORDER_THIN);
  SSCellStyle_setBorderLeft(gNumericND: BORDER_THIN);
  SSCellStyle_setBorderRight(gNumericND: BORDER_THIN);
  SSCellStyle_setBorderTop(gNumericND: BORDER_THIN);

  //
  // Create a cell style for text so that it's
  //  left-aligned
  //

  gText = SSWorkbook_createCellStyle(gBook);
  SSCellStyle_setAlignment(gText: ALIGN_LEFT);
  SSCellStyle_setBorderBottom(gText: BORDER_THIN);
  SSCellStyle_setBorderLeft(gText: BORDER_THIN);
  SSCellStyle_setBorderRight(gText: BORDER_THIN);
  SSCellStyle_setBorderTop(gText: BORDER_THIN);

  //
  // Create a cell style for dates.  Dates in Excel
  //  are numbers that are formatted in a particular
  //  way.
  //

  gDates = SSWorkbook_createCellStyle(gBook);
  vDataFmt = SSWorkbook_createDataFormat(gBook);
  vTempStr = new_String('m/d/yy');
  vDateFmt = SSDataFormat_getFormat(vDataFmt: vTempStr);
  SSCellStyle_setDataFormat(gDates: vDateFmt);
  SSCellStyle_setDataFormat(gDates: vDateFmt);
  SSCellStyle_setBorderBottom(gDates: BORDER_THIN);
  SSCellStyle_setBorderLeft(gDates: BORDER_THIN);
  SSCellStyle_setBorderRight(gDates: BORDER_THIN);
  SSCellStyle_setBorderTop(gDates: BORDER_THIN);

End-Proc;

Dcl-Proc SetJavaPath;


  Dcl-Pr qcmdexc  ExtPgm('QCMDEXC');
    *N char(3000) const options(*varsize);
    *N packed(15:5) const;
  End-Pr qcmdexc;

  Dcl-S vCmdString char(3000);

  // Set Java CLASSPATH
  vCmdString = 'RMVENVVAR ENVVAR(IBM_JAVA_OPTIONS)';
  Monitor;
    qcmdexc(vCmdString : %len(%trim(vCmdString)));
  On-Error;
  EndMon;

  // POI 3.14 needs Java 6 or later. These settings only take effect
  // if the JVM has not already started in this job.
  vCmdString = 'ADDENVVAR ENVVAR(JAVA_HOME) VALUE(' +
             '''' + '/QOpenSys/QIBM/ProdData/JavaVM/jdk80/64bit' + '''' + ')' +
                     ' REPLACE(*YES)';
  Monitor;
    qcmdexc(vCmdString : %len(%trim(vCmdString)));
  On-Error;
  EndMon;

  vCmdString = 'ADDENVVAR ENVVAR(CLASSPATH) VALUE(' +
             '''' + '/home/rthompson/poi3.14/poi-3.14-20160307.jar' +
             ':/home/rthompson/poi3.14/poi-ooxml-3.14-20160307.jar' +
             ':/home/rthompson/poi3.14/poi-ooxml-schemas-3.14-20160307.jar' +
             ':/home/rthompson/poi3.14/poi-scratchpad-3.14-20160307.jar' +
             ':/home/rthompson/poi3.14/xmlbeans-2.6.0.jar' +
             ':/home/rthompson/poi3.14/commons-codec-1.10.jar' +
             ':/home/rthompson/poi3.14/commons-logging-1.2.jar' +
             ':/home/rthompson/poi3.14/curvesapi-1.03.jar' + '''' + ')' +
                     ' REPLACE(*YES)';
  Monitor;
    qcmdexc(vCmdString : %len(%trim(vCmdString)));
  On-Error;
  EndMon;

  vCmdString = 'ADDENVVAR ENVVAR(QIBM_RPG_JAVA_PROPERTIES) ' +
                     'VALUE(' + '''' + '-Djava.awt.headless=true;' +
                     '-Dos400.awt.native=true;' + '''' + ')' +
                     ' REPLACE(*YES)';
  Monitor;
    qcmdexc(vCmdString : %len(%trim(vCmdString)));
  On-Error;
  EndMon;

End-Proc;
