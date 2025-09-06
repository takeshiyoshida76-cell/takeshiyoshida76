       IDENTIFICATION          DIVISION.
       PROGRAM-ID           .  WRTETIME.
       AUTHOR.                 TY.
       
       ENVIRONMENT             DIVISION.
       INPUT-OUTPUT            SECTION.
       FILE-CONTROL.
       SELECT  OUT-FILE        ASSIGN  TO  "MYFILE"
       ORGANIZATION            IS  LINE    SEQUENTIAL.
       
       DATA                    DIVISION.
       FILE                    SECTION.
       FD  OUT-FILE
           LABEL   RECORDS ARE STANDARD.
       01  OUT-REC             PIC X(80).
       
       WORKING-STORAGE         SECTION.
       01  WS-DATE-TIME.
       *  YYYYMMDD
         05  WS-DATE           PIC 9(8).
       *  HHMMEE
         05  WS-TIME           PIC 9(6).
       01  WS-DATE-R           REDEFINES   WS-DATE.
         05  WS-YEAR           PIC 9(4).
         05  WS-MONTH          PIC 9(2).
         05  WS-DAY            PIC 9(2).
       01  WS-TIME-R           REDEFINES   WS-TIME.
         05  WS-HOUR           PIC 9(2).
         05  WS-MINUTE         PIC 9(2).
         05  WS-SECOND         PIC 9(2).
       01  WS-OUT-REC.
         05  FILLER            PIC X( )    VALUE   "CURRENT TIME =".
         05  WS-NOWTIME.
           10  WS-NOWYEAR      PIC 9(4).
           10  FILLER          PIC X(1)    VALUE   "/".
           10  WS-NOWMONTH     PIC 9(2).
           10  FILLER          PIC X(1)    VALUE   "/".
           10  WS-NOWDAY       PIC 9(2)
           10  FILLER          PIC X(1)    VALUE   SPACE.
           10  WS-NOWHOUR      PIC 9(2).
           10  FILLER          PIC X(1)    VALUE   ":".
           10  WS-NOWMINUTE    PIC 9(2).
           10  FILLER          PIC X(1)    VALUE   ":".
           10  WS-NOWSECOND    PIC 9(2).
           10  FILLER          PIC X(1)    VALUE   SPACE.
       PROCEDURE               DIVISION.
       MAIN-LOGIC              SECTION.
          OPEN    OUTPUT       OUT-FILE
       
       *  WRITE MESSAGE
          MOVE    "THIS PROGRAM IS WRITTEN IN COBOL."
                               TO  OUT-REC.
       *  GET SYSTEM DATETIME
          ACCEPT  WS-DATE      FROM    DATE.
          ACCEPT  WS-TIME      FROM    TIME.
       
       *  EDIT MESSAGE+
          MOVE    WS-YEAR      TO  WS-NOWYEAR.
          MOVE    WS-MONTH     TO  WS-NOWMONTH.
          MOVE    WS-DAY       TO  WS-NOWDAY.
          MOVE    WS-HOUR      TO  WS-NOWHOUR.
          MOVE    WS-MINUTE    TO  WS-NOWMINUTE.
          MOVE    WS-SECOND    TO  WS-NOWSECOND.
       
       *  WRITE OUTFILE
          WRITE OUT-REC        FROM    WS-OUT-REC.
       
          CLOSE   OUT-FILE
          STOP    RUN.
