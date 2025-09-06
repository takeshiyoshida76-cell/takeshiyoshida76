       IDENTIFICATION                      DIVISION.
       PROGRAM-ID. WRTETIME.
       AUTHOR. TY.

       ENVIRONMENT                         DIVISION.
       INPUT-OUTPUT                        SECTION.
       FILE-CONTROL.
           SELECT OUT-FILE
           ASSIGN TO "MYFILE.TXT"
           ORGANIZATION IS LINE SEQUENTIAL
           ACCESS MODE IS SEQUENTIAL
           FILE STATUS IS FILE-STATUS-CODE.

       DATA                                DIVISION.
       FILE                                SECTION.
       FD OUT-FILE.
       01 OUT-REC                          PIC X(80).

       WORKING-STORAGE                     SECTION.
       01 FILE-STATUS-CODE                 PIC X(2).
          88 FILE-OK                       VALUE "00".
       
       01 WS-DATE-TIME-DATA.
          05 WS-CURRENT-DATE               PIC 9(8).
          05 WS-CURRENT-TIME               PIC 9(6).
          05 WS-CURRENT-MILLISEC           PIC 9(2).
          05 WS-CURRENT-GMT                PIC S9(4).

       01 WS-FORMATTED-DATE-TIME.
          05 WS-YEAR                       PIC 9(4).
          05 FILLER                        PIC X(1) VALUE "/".
          05 WS-MONTH                      PIC 9(2).
          05 FILLER                        PIC X(1) VALUE "/".
          05 WS-DAY                        PIC 9(2).
          05 FILLER                        PIC X(1) VALUE SPACE.
          05 WS-HOUR                       PIC 9(2).
          05 FILLER                        PIC X(1) VALUE ":".
          05 WS-MINUTE                     PIC 9(2).
          05 FILLER                        PIC X(1) VALUE ":".
          05 WS-SECOND                     PIC 9(2).
       
       01 WS-OUT-LINE.
          05 FILLER                        PIC X(22)
             VALUE "THIS PROGRAM IS WRITTEN IN COBOL.".
       
       01 WS-TIME-LINE.
          05 FILLER                        PIC X(15)
             VALUE "CURRENT TIME = ".
          05 WS-TIME-PART                  PIC X(16).

       PROCEDURE                           DIVISION.
       MAIN-LOGIC                          SECTION.
      * OPEN FILE IN APPEND MODE.
           OPEN EXTEND OUT-FILE.
           IF NOT FILE-OK
               DISPLAY "ERROR OPENING FILE: " FILE-STATUS-CODE
               STOP RUN
           END-IF.

      * WRITE FIRST LINE TO FILE.
           WRITE OUT-REC FROM WS-OUT-LINE.
           IF NOT FILE-OK
               DISPLAY "ERROR WRITING TO FILE: " FILE-STATUS-CODE
               CLOSE OUT-FILE
               STOP RUN
           END-IF.
           
      * GET CURRENT DATE AND TIME FROM SYSTEM.
           ACCEPT WS-DATE-TIME-DATA FROM DATE YYYYMMDD.
           ACCEPT WS-DATE-TIME-DATA FROM TIME HHMMSS.

      * MOVE COMPONENTS TO THE FORMATTED OUTPUT VARIABLE.
           MOVE WS-CURRENT-DATE TO WS-FORMATTED-DATE-TIME.
           MOVE WS-CURRENT-TIME TO WS-FORMATTED-DATE-TIME.

      * MOVE FORMATTED TIME STRING TO OUTPUT LINE.
           MOVE WS-FORMATTED-DATE-TIME TO WS-TIME-PART.
           
      * WRITE THE SECOND LINE TO FILE.
           WRITE OUT-REC FROM WS-TIME-LINE.
           IF NOT FILE-OK
               DISPLAY "ERROR WRITING TO FILE: " FILE-STATUS-CODE
               CLOSE OUT-FILE
               STOP RUN
           END-IF.

      * CLOSE THE FILE.
           CLOSE OUT-FILE.
           IF NOT FILE-OK
               DISPLAY "ERROR WRITING TO FILE: " FILE-STATUS-CODE
               STOP RUN
           END-IF.

           DISPLAY "LOG WRITTEN TO MYFILE.TXT SUCCESSFULLY.".
           STOP RUN.
