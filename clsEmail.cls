VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=============================================================================================
'                                 Simple Mail Transfer Protocol Email Class
'
'=============================================================================================

'Description:   This is a class that was programmmed by me to reduce the hassle of sending mails
'               over SMTP. It has the encoding/decoding routines of commonly used UUEncode and as
'               well as Base64, corformed to MIME standard. It is pretty much unfinished yet, as
'               it doesn't has any POP ability as yet (not even POP Before SMTP auth), but able
'               to send mail in a jiffy. Error trapping also still minimal, n little bit dodgy.
'               ABout the encoding routines, it might not be as fast as other implementations,
'               so if you find it too slow, you can find other better ones at PSC
'
'               Any comments please direct them to cryospherez@yahoo.com
'
'               Enjoy!


'---------------------------------------------------------------------------------------------
'                                       Member Variables Declaration
'---------------------------------------------------------------------------------------------

Option Explicit

'Buffer for Data
Private strCurrentCommand As String
Private strCurrentData As String

Private blnTimeout As Boolean
Private intRetry As Integer



'Attachments variables
Private m_HaveAttachment As Boolean
Private m_FilesAttached() As String
Private m_AttachedFilesChanged As Boolean

Private m_ReadyForSend As Boolean

'Sockets
Private WithEvents SMTPSock As Winsock
Attribute SMTPSock.VB_VarHelpID = -1
Private WithEvents POPSock As Winsock
Attribute POPSock.VB_VarHelpID = -1


' member variable for SMTPConnected property
Private m_SMTPConnected As Boolean
' member variable for POPConnected property
Private m_POPConnected As Boolean

' member variable for SMTPHostname property
Private m_SMTPHostname As String
' member variable for SMTPPort property
Private m_SMTPPort As Integer
' member variable for POPHostname property
Private m_POPHostname As String
' member variable for POPPort property
Private m_POPPort As Integer
' member variable for Username property
Private m_Username As String
' member variable for EmailAddress property
Private m_EmailAddress As String
' member variable for SenderAddress property
Private m_SenderAddress As String
' member variable for RecepientAddress property
Private m_RecepientAddress As String
' member variable for Subject property
Private m_Subject As String
' member variable for MessageBody property
Private m_MessageBody As String
' member variable for CarbonCopy property
Private m_CarbonCopy As String
' member variable for POPAuthorisation property
Private m_POPAuthorisation As Boolean


'---------------------------------------------------------------------------------------------
'                                           Events Declaration
'---------------------------------------------------------------------------------------------




'Error Events
Public Event SMTPError(intErrorNumber As Integer, strDescription As String)
Public Event POPError(intErrorNumber As Integer, strDescription As String)
Public Event UUCodeError(intErrorNumber As Integer, strDescription As String)

'Return attachments queued files path
Public Event FilesInQueue(strFilePath As String, intIndex As Integer)

Public Event SentSuccessful()
Public Event SMTPConnect()
Public Event SMTPDisconnect()

'---------------------------------------------------------------------------------------------
'                                           Constants Declaration
'---------------------------------------------------------------------------------------------

'Standard SMTP Error Constants
Const ERROR_SERVICE_NOT_AVAILABLE = 421         '<domain> Service not available, closing transmission channel
                                                '(This may be a reply to any command if the service knows it
                                                'must shut down)
Const ERROR_MAILBOX_BUSY = 450                  'Requested mail action not taken: mailbox unavailable
                                                '(e.g., mailbox busy)
Const ERROR_PROCESSING_ERROR = 451              'Requested action aborted: local error in processing
Const ERROR_INSUFFICIENT_SYSTEM_STORAGE = 452   'Requested action not taken: insufficient system storage
Const ERROR_COMMAND_UNRECOGNIZED = 500          'Syntax error, command unrecognized
                                                '(This may include errors such as command line too long)
Const ERROR_SYNTAX_ERROR = 501                  'Syntax error in parameters or arguments
Const ERROR_COMMAND_NOT_IMPLEMENTED = 502       'Command not implemented
Const ERROR_BAD_SEQUENCE = 503                  'Bad sequence of commands
Const ERROR_COMMAND_NOT_PARAMETER = 504         'Command Not parameter
Const ERROR_MAILBOX_UNAVAILABLE = 550           'Requested Not Action: mailbox unavailable
                                                '(e.g., mailbox not found, no access, or command rejected
                                                'for policy reasons)
Const ERROR_USER_NOT_LOCAL = 551                'User not local; please try <forward-path>
Const ERROR_EXCEEDED_STORAGE_ALLOCATION = 552   'Requested mail action aborted: exceeded storage allocation
Const ERROR_MAILBOX_NOT_NAME = 553              'Requested Not Action: mailbox Not Name
                                                '(e.g., mailbox syntax incorrect)
Const ERROR_NO_SERVICE = 554            'Transaction failed  (Or, in the case of a connection-opening
                                                'response, "No SMTP service here")
                                                
'Self-Defined Mail Error
Const ERROR_NO_SUBJECT = 601
Const ERROR_NO_MESSAGEBODY = 602
Const ERROR_NO_SENDER_ADDRESS = 603
Const ERROR_NO_RECEIVER_ADDRESS = 604

'Self-Defined Non-Standard SMTP Related Errors
Const ERROR_SMTP_NOT_CONNECTED = 701
Const ERROR_SMTP_HOST_INVALID = 702
Const ERROR_SMTP_NO_SOCKET = 703

'Self-Defined Non-Standard POP Related Errors
Const ERROR_POP_NOT_CONNECTED = 801
Const ERROR_POP_NOT_COMPLETE = 802
Const ERROR_POP_PASSWORD_INVALID = 803
Const ERROR_POP_NO_SOCKET = 804

'Self-Defined Non-Standard Connection Related Errors
Const ERROR_TIMEOUT = 900
Const ERROR_TRANSACTION_IN_PROCESS = 901
Const ERROR_TRANSACTION_FAILED = 902

'Self Defined Error Constants for UUCode/Base64 Encoding/Decoding
Const ERROR_UUENCODING = 1001
Const ERROR_UUDECODING = 1002
Const ERROR_B64ENCODING = 1003
Const ERROR_B64DECODING = 1004

'Self-Defined Constants
Const sngTimeout As Single = 5
Const intNumOfRetry As Integer = 5

'Boundary for MIME parts
Const MIMEBoundary = "123451234512345"

'=============================================================================================
'=============================================================================================
'                                       ENCODING/DECODING SECTION
'=============================================================================================
'=============================================================================================


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'                                          Files UUencoding
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'Description:- Can only encode 1 file per call, but appending job is possible

'---------------------------------------------------------------------------------------------


Public Function UUEncode(strFilename As String, blnAppend As Boolean, Optional strOutputFile As String) As Integer

Dim strFileTitle    As String
Dim lngFileSize     As Long
Dim intFileNum      As Integer
Dim intFileNum2     As Integer
Dim lngNumOfLines   As Long
Dim intCharCode     As Integer
Dim intCounter      As Integer
Dim strBuffer       As String
Dim strEncodedLine  As String
Dim lngLineNumber   As Long
Dim intByteNumber   As Integer

'Error catcher
On Error GoTo Erh:

'Extract the title of the file.
For intCounter = Len(strFilename) To 1 Step -1
    If Mid$(strFilename, intCounter, 1) = "\" Then
        strFileTitle = Mid$(strFilename, intCounter + 1)
        
        Exit For
    End If
Next intCounter

'Get the file size
lngFileSize = FileLen(strFilename)
'Get the number of lines of encoded chars.
lngNumOfLines = (lngFileSize \ 45) + 1

'Open the input file.
intFileNum = FreeFile
Open strFilename For Binary Access Read As intFileNum



'Open the output file
intFileNum2 = FreeFile
If strOutputFile = "" And blnAppend Then
    Open strFilename & ".uue" For Append As intFileNum2
    Print #intFileNum2, ""
    Print #intFileNum2, "begin 664 " & strFileTitle
ElseIf blnAppend Then
    Open strOutputFile For Append As intFileNum2
    Print #intFileNum2, ""
    Print #intFileNum2, "begin 664 " & strFileTitle
ElseIf strOutputFile = "" Then
    Open strFilename & ".uue" For Output As intFileNum2
    Print #intFileNum2, "begin 664 " & strFileTitle
Else
    Open strOutputFile For Output As intFileNum2
    Print #intFileNum2, "begin 664 " & strFileTitle
End If



'Extract chunks of 45 bytes from source file, and write it in output file.
For lngLineNumber = 1 To lngNumOfLines

    'Prepare the buffer for extracted bytes.
    If lngLineNumber = lngNumOfLines Then
        strBuffer = Space$(lngFileSize Mod 45)
        
        'If the last line have nothing (The file perfectly divisible by 45)
        If strBuffer = "" Then Exit For
        
    Else
        strBuffer = Space$(45)
    End If
    
    'Read the bytes and store it into the buffer.
    Get intFileNum, , strBuffer
    
    'Prepare the marker at the beginning of line, with details of how many bytes in it.
    strEncodedLine = Chr$(Len(strBuffer) + 32)
    
    'Add padding of ASCII Character 0 for last line if it is not perfectly divisible by 3
    If (Len(strBuffer) Mod 3) Then
        strBuffer = strBuffer & String$(3 - (Len(strBuffer) Mod 3), Chr$(0))
    End If
    
    'Start UUEncoding with steps of 3 bytes
    For intByteNumber = 1 To Len(strBuffer) Step 3
    
        'For the first byte.
        intCharCode = Asc(Mid$(strBuffer, intByteNumber, 1)) \ 4 + 32
        
        strEncodedLine = strEncodedLine & IIf(intCharCode = 32, Chr$(96), Chr$(intCharCode))
        
        
        'First Byte + Second Byte
        intCharCode = (Asc(Mid$(strBuffer, intByteNumber, 1)) Mod 4) * 16 _
                    + (Asc(Mid$(strBuffer, intByteNumber + 1, 1)) \ 16) + 32
        
        strEncodedLine = strEncodedLine & IIf(intCharCode = 32, Chr$(96), Chr$(intCharCode))
        
        'First Byte + Second Byte + Third Byte
        intCharCode = (Asc(Mid$(strBuffer, intByteNumber + 1, 1)) Mod 16) * 4 _
                    + (Asc(Mid$(strBuffer, intByteNumber + 2, 1)) \ 64) + 32
        
        strEncodedLine = strEncodedLine & IIf(intCharCode = 32, Chr$(96), Chr$(intCharCode))
        
        'First Byte + Second Byte + Third Byte + Fourth Byte
        intCharCode = Asc(Mid$(strBuffer, intByteNumber + 2, 1)) Mod 64 + 32
        
        strEncodedLine = strEncodedLine & IIf(intCharCode = 32, Chr$(96), Chr$(intCharCode))
    
    Next intByteNumber
    
    'Write into the file.
    Print #intFileNum2, strEncodedLine

Next lngLineNumber

'Mark the end of file.
Print #intFileNum2, Chr$(96) & vbCrLf & "end"

'Close the files
Close intFileNum
Close intFileNum2

'Return error free

UUEncode = 0
Exit Function


'Error handler
Erh:

'Return the error number
RaiseEvent UUCodeError(Err.Number, Err.Description)
MsgBox Err.Description

End Function


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'                                          Files UUDecoding
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'Description:- Unlike UUEncoding routine, UUDecoding routine here can decode multiple files

'---------------------------------------------------------------------------------------------


Public Sub UUDecode(strFilename As String, Optional strPath As String)


Dim intFileNum      As Integer
Dim intFileNum2     As Integer
Dim intCounter      As Integer
Dim strOutputFile   As String
Dim lngPos          As Long
Dim strDataLine     As String
Dim strDecodedLine  As String
Dim intLenChar      As Integer

On Error GoTo Erh:

'Open UUEncoded file
intFileNum = FreeFile
Open strFilename For Input As intFileNum

'Find the beginning of a file.
Do
    Line Input #intFileNum, strDataLine
    
    If Trim$(Left$(strDataLine, 9)) Like "begin ###" Then
    
        'Extract the file name from the line
        strOutputFile = Mid$(Trim$(strDataLine), 11)
        
        'Open the output file.
        intFileNum2 = FreeFile
        
        'Check if optional output path is specified
        If strPath = "" Then
        
            'Extract the path from the original UUEncoded file
            For intCounter = Len(strFilename) To 1 Step -1
                If Mid$(strFilename, intCounter, 1) = "\" Then
                    strPath = Left$(strFilename, intCounter)
                    
                    Exit For
                End If
            Next intCounter
            'Open the file in the original file path
            Open strPath & strOutputFile For Binary As intFileNum2
            
        Else
            'Use the path provided
            Open strPath & strOutputFile For Binary As intFileNum2
        End If
        
        Do
            'Read a line
            Line Input #intFileNum, strDataLine
            
            'If this is the last line, or a dodgy new beginning,
            'stop reading n close output file
            If Trim$(strDataLine) = Chr$(96) Then
                
                Line Input #intFileNum, strDataLine
                
                If Trim$(strDataLine) = "end" Then
                    Exit Do
                End If
                
            ElseIf Trim$(strDataLine) = "end" Then
            
                Exit Do
                
            ElseIf Trim$(strDataLine) <> "" Then
                
                'Extract the first char out of the line
                intLenChar = Asc(Left$(strDataLine, 1)) - 32
                strDataLine = Mid$(strDataLine, 2)
                
                
                
                'Check if any ' is in the line, and replace it with a space
                Do
                    lngPos = InStr(strDataLine, Chr$(96))
                    
                    If lngPos <> 0 Then
                        strDataLine = Left$(strDataLine, lngPos - 1) & " " & Mid$(strDataLine, lngPos + 1)
                    Else
                        Exit Do
                    End If
                
                Loop
                
                
                'UUDecode the line
                For intCounter = 1 To Len(strDataLine) Step 4
                    strDecodedLine = strDecodedLine & Chr((Asc(Mid$(strDataLine, intCounter, 1)) - 32) * 4 _
                                                        + (Asc(Mid$(strDataLine, intCounter + 1, 1)) - 32) \ 16)
                                                        
                    strDecodedLine = strDecodedLine & Chr((Asc(Mid$(strDataLine, intCounter + 1, 1)) Mod 16) * 16 _
                                                        + (Asc(Mid$(strDataLine, intCounter + 2, 1)) - 32) \ 4)
                                                        
                    strDecodedLine = strDecodedLine & Chr((Asc(Mid$(strDataLine, intCounter + 2, 1)) Mod 4) * 64 _
                                                         + Asc(Mid$(strDataLine, intCounter + 3, 1)) - 32)
                Next intCounter
                
                ''Check the line for any missing char
                If Len(strDecodedLine) < intLenChar Then
                    RaiseEvent UUCodeError(ERROR_UUDECODING, "There are missing bytes in file " & strOutputFile & ". The file might be corrupted")
                End If
                
                
                'Write the decoded line to output file
                Put #intFileNum2, , strDecodedLine
                
                'Reset the decoded line
                strDecodedLine = ""
                
            End If
            
            
            
        
        Loop
        
        'Close output file
        Close #intFileNum2
        
    End If
    
Loop Until EOF(intFileNum)

'Close input file
Close #intFileNum

Exit Sub

Erh:

RaiseEvent UUCodeError(ERROR_UUDECODING, Err.Number & ":" & Err.Description)
MsgBox "Error error"

End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'                                          File Base64 Encoding
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'Description:-  Encode files using Base64 encoding. Attach MIME headers too.
'               Appending job possible

'---------------------------------------------------------------------------------------------

Public Sub Base64Encode(strFilename As String, blnAppend As Boolean, Optional strOutputFile As String, Optional blnFinalFile As Boolean = False)

'A dodgy error trapping..like usual.ahahaha
On Error GoTo Erh:

'Base 64 table
Const BASE64_TABLE As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"


Dim strEncodedLine  As String
Dim strBuffer       As String
Dim strFileTitle    As String
Dim strMIMEHeader   As String
Dim intCounter      As Integer
Dim lngLineNumber   As Long
Dim intFileNum      As Integer
Dim intFileNum2     As Integer
Dim lngNumOfLines   As Long
Dim lngFileSize     As Long

'Get the file size
lngFileSize = FileLen(strFilename)

'Find the number of data lines to be encoded in 76 chunks
If (lngFileSize Mod 76) Then
    lngNumOfLines = lngFileSize \ 57 + 1
Else
    lngNumOfLines = lngFileSize \ 57
End If


'Extract the title of the file.
For intCounter = Len(strFilename) To 1 Step -1
    If Mid$(strFilename, intCounter, 1) = "\" Then
        strFileTitle = Mid$(strFilename, intCounter + 1)
        
        Exit For
        
    End If
Next intCounter

'Open the input file
intFileNum = FreeFile
Open strFilename For Binary As intFileNum

'Construct MIME header for attachment
strMIMEHeader = "Content-Type: application/octet-stream;" & vbCrLf & vbTab & "name=" & """" & strFileTitle & """" & vbCrLf & _
                "Content-Transfer-Encoding: base64" & vbCrLf & _
                "Content-Disposition: attachment;" & vbCrLf & vbTab & "filename=" & """" & strFileTitle & """" & vbCrLf

'Open the output file
intFileNum2 = FreeFile
If strOutputFile = "" And blnAppend Then
    Open strFilename & ".b64" For Append As intFileNum2
    'Print #intFileNum2, ""
    Print #intFileNum2, strMIMEHeader
ElseIf strOutputFile <> "" And blnAppend Then
    Open strOutputFile For Append As intFileNum2
    'Print #intFileNum2, ""
    Print #intFileNum2, strMIMEHeader
ElseIf strOutputFile = "" Then
    Open strFilename & ".b64" For Output As intFileNum2
    Print #intFileNum2, strMIMEHeader
Else
    Open strOutputFile For Output As intFileNum2
    Print #intFileNum2, strMIMEHeader
End If


For lngLineNumber = 1 To lngNumOfLines
    
    
    
    'Prepare the buffer for extracted bytes.
    If lngLineNumber = lngNumOfLines Then
        strBuffer = Space$(lngFileSize Mod 57)
    Else
        strBuffer = Space$(57)
    End If
    
    'Read the bytes and store it into the buffer.
    Get #intFileNum, , strBuffer
    
    
    
    'Start Base64 Encoding
    For intCounter = 1 To (Len(strBuffer) - Len(strBuffer) Mod 3) Step 3
    
        strEncodedLine = strEncodedLine + Mid(BASE64_TABLE, _
                            (Asc(Mid(strBuffer, intCounter, 1)) \ 4) + 1, 1)
       
        strEncodedLine = strEncodedLine + Mid(BASE64_TABLE, _
                            ((Asc(Mid(strBuffer, intCounter, 1)) Mod 4) * 16 _
                            + Asc(Mid(strBuffer, intCounter + 1, 1)) \ 16) + 1, 1)
       
        strEncodedLine = strEncodedLine + Mid(BASE64_TABLE, _
                            ((Asc(Mid(strBuffer, intCounter + 1, 1)) Mod 16) * 4 _
                            + Asc(Mid(strBuffer, intCounter + 2, 1)) \ 64) + 1, 1)
       
        strEncodedLine = strEncodedLine + Mid(BASE64_TABLE, _
                            (Asc(Mid(strBuffer, intCounter + 2, 1)) Mod 64) + 1, 1)
     Next intCounter
     
    'Adding pads '=' at the last line if the line is consists of 16 or 8 bytes
    If (Len(strBuffer) Mod 3) = 2 Then
         
        strEncodedLine = strEncodedLine + Mid(BASE64_TABLE, _
                        (Asc(Mid(strBuffer, intCounter, 1)) \ 4) + 1, 1)
         
        strEncodedLine = strEncodedLine + Mid(BASE64_TABLE, _
                        (Asc(Mid(strBuffer, intCounter, 1)) Mod 4) * 16 _
                       + Asc(Mid(strBuffer, intCounter + 1, 1)) \ 16 + 1, 1)
         
        strEncodedLine = strEncodedLine + Mid(BASE64_TABLE, _
                        (Asc(Mid(strBuffer, intCounter + 1, 1)) Mod 16) * 4 + 1, 1)
         
        strEncodedLine = strEncodedLine & "="
         
    ElseIf (Len(strBuffer) Mod 3) = 1 Then
         
        strEncodedLine = strEncodedLine + Mid(BASE64_TABLE, _
                         Asc(Mid(strBuffer, intCounter, 1)) \ 4 + 1, 1)
         
        strEncodedLine = strEncodedLine + Mid(BASE64_TABLE, _
                        (Asc(Mid(strBuffer, intCounter, 1)) Mod 4) * 16 + 1, 1)
         
        strEncodedLine = strEncodedLine & "=="
        
    End If
    
    'Write into the output file
    Print #intFileNum2, strEncodedLine
    
    'Reset the encoded line to nothing
    strEncodedLine = ""

Next lngLineNumber
            
'Mark the end of file.
Print #intFileNum2, ""

If blnFinalFile Then
    Print #intFileNum2, "--" & MIMEBoundary & "--"
Else
    Print #intFileNum2, "--" & MIMEBoundary
End If

'Close the file
Close #intFileNum
Close #intFileNum2

Exit Sub

Erh:

RaiseEvent UUCodeError(ERROR_B64ENCODING, Err.Number & ":" & Err.Description)

End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'                                          File Base64 Decoding
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'Description:- Decode files using Base64 encoding. Use MIME headers. Support multiple files
'              (I think).

'---------------------------------------------------------------------------------------------


Public Sub Base64Decode(strFilename As String, Optional strPath As String)

'Base 64 table
Const BASE64_TABLE As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"


Dim intFileNum      As Integer
Dim intFileNum2     As Integer
Dim intCounter      As Integer
Dim strOutputFile   As String
Dim lngPos          As Long
Dim strDataLine     As String
Dim strDecodedLine  As String
Dim intLenChar      As Integer

On Error GoTo Erh:

'Open UUEncoded file
intFileNum = FreeFile
Open strFilename For Input As intFileNum

'Find the beginning of a file.
Do
    Line Input #intFileNum, strDataLine
    
    If UCase$(Trim$(strDataLine)) Like "CONTENT?TYPE:*APPLICATION/OCTET?STREAM*" Then
    
        'Extract the file name from the header
        strOutputFile = Mid$(strDataLine, InStr(1, strDataLine, "=") + 1)
        
        'Remove quotes if it exists.
        If Left$(strOutputFile, 1) = """" Then
            strOutputFile = Mid$(strOutputFile, 2, Len(strOutputFile) - 2)
        End If
        
        'Check the header whether this is really a base64 encoded stuff
        Do Until UCase$(Trim$(strDataLine)) Like "CONTENT?TRANSFER?ENCODING:*"
            Line Input #intFileNum, strDataLine
        Loop
        
        'If it is not encoded in Base64
        If Trim$(UCase$(Mid$(strDataLine, InStr(1, strDataLine, ":") + 1))) <> "BASE64" Then
            Exit Sub
        End If
        
        'Skipped the rest of the header until the blank line between the header n the b64 stuff
        While Trim$(strDataLine) <> ""
            Line Input #intFileNum, strDataLine
        Wend
        
        'Open the output file.
        intFileNum2 = FreeFile
        
        'Check if optional output path is specified
        If strPath = "" Then
        
            'Extract the path from the original UUEncoded file
            For intCounter = Len(strFilename) To 1 Step -1
                If Mid$(strFilename, intCounter, 1) = "\" Then
                    strPath = Left$(strFilename, intCounter)
                    
                    Exit For
                    
                End If
            Next intCounter
            'Open the file in the original file path
            Open strPath & strOutputFile For Binary As intFileNum2
            
        Else
            'Use the path provided
            Open strPath & strOutputFile For Binary As intFileNum2
        End If
        
        Do
            'Read a line
            
            Line Input #intFileNum, strDataLine
            
            'If this is the last line, or a dodgy new beginning,
            'stop reading n close output file
            If Trim$(strDataLine) = "====" Then
                Exit Do
            ElseIf Trim$(strDataLine) = "" Then
            
                Exit Do
                
            ElseIf Trim$(strDataLine) <> "" Then
    
                
                
                
                'UUDecode the line
                For intCounter = 1 To Len(strDataLine) Step 4
                    
                    If Mid$(strDataLine, intCounter + 2) = "=" And Mid$(strDataLine, intCounter + 3) = "=" Then
                    
                        strDecodedLine = strDecodedLine & Chr$((B64CharCode(Mid$(strDataLine, intCounter, 1)) - 1) * 4 _
                                                            + (B64CharCode(Mid$(strDataLine, intCounter + 1, 1)) - 1) \ 16)
                    
                    ElseIf Mid$(strDataLine, intCounter + 3) = "=" Then
                        
                        strDecodedLine = strDecodedLine & Chr$((B64CharCode(Mid$(strDataLine, intCounter, 1)) - 1) * 4 _
                                                            + (B64CharCode(Mid$(strDataLine, intCounter + 1, 1)) - 1) \ 16)
                                                        
                        strDecodedLine = strDecodedLine & Chr$(((B64CharCode(Mid$(strDataLine, intCounter + 1, 1)) - 1) Mod 16) * 16 _
                                                            + (B64CharCode(Mid$(strDataLine, intCounter + 2, 1)) - 1) \ 4)
                    
                    Else
                    
                        strDecodedLine = strDecodedLine & Chr$((B64CharCode(Mid$(strDataLine, intCounter, 1)) - 1) * 4 _
                                                            + (B64CharCode(Mid$(strDataLine, intCounter + 1, 1)) - 1) \ 16)
                                                        
                        strDecodedLine = strDecodedLine & Chr$(((B64CharCode(Mid$(strDataLine, intCounter + 1, 1)) - 1) Mod 16) * 16 _
                                                            + (B64CharCode(Mid$(strDataLine, intCounter + 2, 1)) - 1) \ 4)
                                                        
                        strDecodedLine = strDecodedLine & Chr$(((B64CharCode(Mid$(strDataLine, intCounter + 2, 1)) - 1) Mod 4) * 64 _
                                                            + B64CharCode(Mid$(strDataLine, intCounter + 3, 1)) - 1)
                    End If
                    
                Next intCounter
                
                ''Check the line for any missing char
                If Len(strDecodedLine) < intLenChar Then
                    RaiseEvent UUCodeError(ERROR_UUDECODING, "There are missing bytes in file " & strOutputFile & ". The file might be corrupted")
                End If
                
                
                'Write the decoded line to output file
                Put #intFileNum2, , strDecodedLine
                
                'Reset the decoded line
                strDecodedLine = ""
                
            End If
            
            
            
        
        Loop Until EOF(intFileNum)
        
        'Close output file
        Close #intFileNum2
        
    End If
    
Loop Until EOF(intFileNum)

'Close input file
Close #intFileNum

Exit Sub

Erh:

RaiseEvent UUCodeError(ERROR_UUDECODING, Err.Number & ":" & Err.Description)


End Sub

Private Function B64CharCode(strB64Char As String) As Integer

'Base 64 table
Const BASE64_TABLE As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

B64CharCode = InStr(BASE64_TABLE, strB64Char)

End Function



'=============================================================================================
'=============================================================================================
'                               SIMPLE MAIL TRANSFER PROTOCOL (SMTP) SECTION
'=============================================================================================
'=============================================================================================



Public Sub SMTPProcessData(strData As String)
    
Dim strRCPTAddress
Static intCounter As Integer
Dim blnSendingAddress As Boolean

    
'Stop timeout timer
blnTimeout = False
    
    
Select Case Val(Left$(strData, 3))

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'                                      SMTP Transactions Process
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


'SMTP Server is ready
Case 220

    SMTPSendData "HELO [" & SMTPSock.LocalIP & "]"
    strCurrentCommand = "HELO"
    
    TimeOut

'---------------------------------------------------------------------------------------------

'Previous command is succesful.
Case 250
    
    
    
    Select Case strCurrentCommand
    
    '-----------------------------------------------------------------------------------------
    Case "HELO"
        
        
        
        'Send Sender Address
        SMTPSendData "MAIL FROM:<" & m_SenderAddress & ">"
        strCurrentCommand = "MAIL"
        
        TimeOut
        
    '-----------------------------------------------------------------------------------------
    Case "MAIL"
        
        
        'Send Recepient Address(s)
        
        If blnSendingAddress = False Then
            If Trim$(m_CarbonCopy) <> "" Then
                strRCPTAddress = Split(m_RecepientAddress & "," & m_CarbonCopy, ",")
            Else
                strRCPTAddress = Split(m_RecepientAddress, ",")
            End If
            
            blnSendingAddress = True
        End If
        
        
        If intCounter < UBound(strRCPTAddress) Then
            SMTPSendData "RCPT TO:<" & strRCPTAddress(intCounter) & ">"
            intCounter = intCounter + 1
            
            TimeOut
            
        ElseIf intCounter = UBound(strRCPTAddress) Then
            
            SMTPSendData "RCPT TO:<" & strRCPTAddress(intCounter) & ">"
            intCounter = 0
            strCurrentCommand = "RCPT"
            blnSendingAddress = False
            
            TimeOut
            
        End If
        
    '-----------------------------------------------------------------------------------------
    Case "RCPT"
    
        
        'Notify the server that the body is coming up.
        SMTPSendData "DATA"
        
        TimeOut
    
    '-----------------------------------------------------------------------------------------
    Case "DATA"
        
        'Sent successful.Mark ready for next delivery
        
        RaiseEvent SentSuccessful
        m_ReadyForSend = True
    
    Case "RSET"
        
        'Somehow an error occured. Restarted the process.
        'Potential looping
        
        If intRetry = intNumOfRetry Then
            
            intRetry = 0
            'Stop retrying
            RaiseEvent SMTPError(ERROR_BAD_SEQUENCE, "Unable to performs proper transaction with the server")
            
            'Flip the switch to ready for sending next mail, though same error likely to happen?
            m_ReadyForSend = True
            
        Else
            
            'Reinitiliaze mail sending process
            
            SMTPSendData "MAIL FROM:<" & m_SenderAddress & ">"
            strCurrentCommand = "MAIL"
        
            TimeOut
            
        End If
        
    End Select
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'                                     SMTP Message Composing Process
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Case 354

    
    'Server is ready for Message
    SMTPSendData ProcessHeader & m_MessageBody
    
    If m_HaveAttachment Then
    
        'Send a MIME multipart boundary.
        SMTPSendData "--" & MIMEBoundary
        'Encode n send the files in queue
        
        EncodeAndSend
        
        'Mark end of message
        SMTPSendData "."
    Else
        'Mark end of message
        SMTPSendData "."
    End If
    
    strCurrentCommand = "DATA"
    
    TimeOut
    
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'                                    SMTP Error Handling Process
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'                                        TEMPORARY ERRORS

Case ERROR_SERVICE_NOT_AVAILABLE '421
'---------------------------------------------------------------------------------------------
Case ERROR_MAILBOX_BUSY '450
'---------------------------------------------------------------------------------------------
Case ERROR_PROCESSING_ERROR '451
'---------------------------------------------------------------------------------------------
Case ERROR_INSUFFICIENT_SYSTEM_STORAGE '452



'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'                                    CRITICAL/PERMANENT ERRORS

Case ERROR_COMMAND_UNRECOGNIZED '500

    Select Case strCurrentCommand
    
    Case "EHLO"
        'The server does not support Extended Helo, use trad. Helo instead
        
        SMTPSendData "HELO [" & SMTPSock.LocalIP & "]"
        strCurrentCommand = "HELO"
        
        TimeOut
    
    
    End Select
    
    'Raise Error
    RaiseEvent SMTPError(Val(Left$(strData, 3)), "SMTPError: " & Trim$(Mid$(strData, 4)))
'---------------------------------------------------------------------------------------------
Case ERROR_SYNTAX_ERROR '501

    Select Case strCurrentCommand
    
    Case "MAIL"
    
        'The server does not support extended Mail command. Exclude the message size.
    
        SMTPSendData "MAIL FROM:<" & m_SenderAddress & ">"
        strCurrentCommand = "MAIL"
            
        TimeOut
        
    End Select
    
    'Raise Error
    RaiseEvent SMTPError(Val(Left$(strData, 3)), "SMTPError: " & Trim$(Mid$(strData, 4)))
'---------------------------------------------------------------------------------------------
Case ERROR_COMMAND_NOT_IMPLEMENTED '502

'Extended Commands not available at this server.

    Select Case strCurrentCommand
    
    Case "EHLO"
        'The server does not support Extended Helo, use trad. Helo instead
        
        SMTPSendData "HELO [" & SMTPSock.LocalIP & "]"
        strCurrentCommand = "HELO"
        
        TimeOut
        
    Case "MAIL"
    
        'The server does not support extended Mail command. Exclude the message size.
    
        SMTPSendData "MAIL FROM:<" & m_SenderAddress & ">"
        strCurrentCommand = "MAIL"
            
        TimeOut
    
    End Select

    'Raise Error
    RaiseEvent SMTPError(Val(Left$(strData, 3)), "SMTPError: " & Trim$(Mid$(strData, 4)))
'---------------------------------------------------------------------------------------------
Case ERROR_BAD_SEQUENCE '503

    'Somehow some commands are sent not in sequence ?
    'MOstly because of some error in the client itself.Probably should just reset n resend the msg
    
    Select Case strCurrentCommand
    
    Case "MAIL"
        'Most probably some problem with EHLO/HELO
        
        SMTPSendData "RSET"
        
        TimeOut
    
    Case "RCPT"
        'Most probably some problem with MAIL
        
        SMTPSendData "RSET"
        
        TimeOut
        
    Case "DATA"
        'Most probably some problem with RCPT
        
        SMTPSendData "RSET"
        
        TimeOut
    
    Case Else
    
        SMTPSendData "RSET"
        
        TimeOut
    
    End Select
    
    'Raise Error
    RaiseEvent SMTPError(Val(Left$(strData, 3)), "SMTPError: " & Trim$(Mid$(strData, 4)))
'---------------------------------------------------------------------------------------------
Case ERROR_COMMAND_NOT_PARAMETER '504

    'Raise Error
    RaiseEvent SMTPError(Val(Left$(strData, 3)), "SMTPError: " & Trim$(Mid$(strData, 4)))
'---------------------------------------------------------------------------------------------
Case ERROR_MAILBOX_UNAVAILABLE '550

    'Raise Error
    RaiseEvent SMTPError(Val(Left$(strData, 3)), "SMTPError: " & Trim$(Mid$(strData, 4)))
'---------------------------------------------------------------------------------------------
Case ERROR_USER_NOT_LOCAL '551

'Parse out the given address

    'Raise Error
    RaiseEvent SMTPError(Val(Left$(strData, 3)), "SMTPError: " & Trim$(Mid$(strData, 4)))
'---------------------------------------------------------------------------------------------
Case ERROR_EXCEEDED_STORAGE_ALLOCATION '552
    'Raise Error
    RaiseEvent SMTPError(Val(Left$(strData, 3)), "SMTPError: " & Trim$(Mid$(strData, 4)))
'---------------------------------------------------------------------------------------------
Case ERROR_MAILBOX_NOT_NAME '553
    'Raise Error
    RaiseEvent SMTPError(Val(Left$(strData, 3)), "SMTPError: " & Trim$(Mid$(strData, 4)))
'---------------------------------------------------------------------------------------------
Case ERROR_NO_SERVICE  '554

    'Server does not serve SMTP, or reject connection for some unknown reason
    
    
    'Close socket
    SMTPSock.Close
    
    'Raise error event
    RaiseEvent SMTPError(ERROR_SMTP_HOST_INVALID, "The server does not server SMTP, or you are rejected from the service.")

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

End Select

    
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'                                          SMTP Transactions Process
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



Private Function ProcessHeader() As String

Dim strMIMEMessage          As String
Dim strMIMEContent          As String
Dim strMIMEMessageContent   As String
Dim strMIMEClient           As String
Dim strMIMEMailInfo         As String
Dim strMIMEVersion          As String







strMIMEMailInfo = "DATE: " & Format(Now, "dd mmm yy ttttt") & vbCrLf & _
                  "FROM: " & Trim$(m_SenderAddress) & vbCrLf & _
                  "TO: " & Trim$(m_RecepientAddress) & vbCrLf & _
                  "SUBJECT: " & m_Subject
            
strMIMEClient = "X-Mailer: " & App.ProductName & vbCrLf & _
                "X-Version: " & App.Major & "." & App.Minor & vbCrLf & _
                "X-CompanyName: " & App.CompanyName
                
strMIMEVersion = "MIME-Version: 1.0"

strMIMEContent = "Content-Type: multipart/mixed;" & vbCrLf & vbTab & _
                 "boundary=" & """" & MIMEBoundary & """"

strMIMEMessageContent = "This is a multi-part message in MIME format." & vbCrLf & _
                        "--" & MIMEBoundary & vbCrLf & _
                        "Content-Type: text/plain;" & vbCrLf & _
                         vbTab & "charset=" & """" & "iso-8859-1" & """" & vbCrLf & _
                        "Content-Transfer-Encoding: 7bit"

ProcessHeader = strMIMEMailInfo & vbCrLf & _
                strMIMEVersion & vbCrLf & _
                strMIMEContent & vbCrLf & _
                strMIMEClient & vbCrLf & _
                vbCrLf & _
                strMIMEMessageContent & vbCrLf & _
                vbCrLf

                
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'                                          SMTP Transactions Process
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Private Function PrepareMessage(strMessage As String) As String
' This function check for a period at the beginning of line, and add one extra if found
' This is to make sure the server doesn't terminate transition of the body message due to client's period

Dim vCRLF
Dim vStart

vStart = 1

Do
    vCRLF = InStr(vStart, strMessage, vbCrLf)
    If vCRLF <> 0 Then
        If Mid$(strMessage, vCRLF + 2, 1) = "." Then
            strMessage = Left$(strMessage, vCRLF + 1) & "." & Mid$(strMessage, vCRLF + 2)
        End If
        
        vStart = vCRLF + 3
    End If
Loop Until vCRLF = 0

End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'                                        SMTP Transactions Process
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Public Sub SendMail()

If m_ReadyForSend = False Then
    RaiseEvent SMTPError(ERROR_TRANSACTION_IN_PROCESS, "Email transaction in process. Terminate current job before sending another mail")
    Exit Sub
Else
    m_ReadyForSend = False
End If

'Check if the sockets for operations are ready.
If SMTPSock Is Nothing Then
    RaiseEvent SMTPError(ERROR_SMTP_NO_SOCKET, "No socket for SMTP connection specified")
    Exit Sub
End If

If POPAuthorisation And POPSock Is Nothing Then
    RaiseEvent POPError(ERROR_POP_NO_SOCKET, "No socket for POP connection specified")
    Exit Sub
End If

'Check all required fields are filled
If m_SMTPHostname = "" Or _
   m_SMTPPort = 0 Then
    RaiseEvent SMTPError(ERROR_SMTP_HOST_INVALID, "Required SMTP connection parameters missing")
    Exit Sub
End If

If m_SenderAddress = "" Then
    RaiseEvent SMTPError(ERROR_NO_SENDER_ADDRESS, "No sender address specified")
    Exit Sub
End If

If m_RecepientAddress = "" Then
    RaiseEvent SMTPError(ERROR_NO_RECEIVER_ADDRESS, "No receiver address specified")
    Exit Sub
End If

If m_Subject = "" Then
    RaiseEvent SMTPError(ERROR_NO_SUBJECT, "No subject specified")
End If

If m_MessageBody = "" Then
    RaiseEvent SMTPError(ERROR_NO_MESSAGEBODY, "No message body specified")
End If

'If POP Authorisation is required for sending emails
If m_POPAuthorisation = True Then
    If (m_POPHostname = "" Or _
        m_POPPort = 0) Then
            RaiseEvent POPError(ERROR_POP_NOT_COMPLETE, "Required POP connection parameters missing")
            Exit Sub
    End If
    
    'Connect to POP Server for initiating authoriation process.
    
    'TODO: POP Before SMTP Authorisation
    
End If


'All ready. Connect to SMTP server now
If SMTPConnected = False Then
    SMTPSock.Connect m_SMTPHostname, m_SMTPPort
Else
    'Reinitiate process (Equivalent to RSET command)
    SMTPSendData "HELO [" & SMTPSock.LocalIP & "]"
    strCurrentCommand = "HELO"
    
    TimeOut
    
End If


End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'                                        Attachments Handling
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'User-accessible functions

Public Sub AddFiles(strFilenames As String)

Dim strTempFiles()  As String
Dim intCounter      As Integer
Dim intCounter2     As Integer

Dim intPos As Integer

'Exit routine if only empty string is passed.
If Trim$(strFilenames) = "" Then
    Exit Sub
End If

'Can pass multiple files at once, which must each enclosed in quotes and separated by
'white space. . e.g "a.txt" "b.txt" . It's okay without quotes for single file.

'Turn the files paths string to array of path
strTempFiles = Split(strFilenames, """" & " " & """")


For intCounter = 0 To UBound(strTempFiles)

    'Remove starting quotes
    If Left$(strTempFiles(intCounter), 1) = """" Then
        strTempFiles(intCounter) = Mid$(strTempFiles(intCounter), 2)
    End If
    
    'Remove trailing quotes
    If Right$(strTempFiles(intCounter), 1) = """" Then
        strTempFiles(intCounter) = Mid$(strTempFiles(intCounter), Len(strTempFiles(intCounter)) - 1)
    End If
    
    'Add the files to the message's queue of files
    If Dir(strTempFiles(intCounter)) <> "" Then
    
        'Flip the condition if not yet
        If m_HaveAttachment = False Then
    
            m_HaveAttachment = True
        
            'Prepare file queue array
            ReDim m_FilesAttached(0)
            
            'Assign the valid path to the last element of the queue
            m_FilesAttached(0) = strTempFiles(intCounter)
            
        Else
    
            'Redimension the file queue array
            ReDim Preserve m_FilesAttached(UBound(m_FilesAttached) + 1)
            'Assign the valid path to the last element of the queue
            m_FilesAttached(UBound(m_FilesAttached)) = strTempFiles(intCounter)
        
        End If
        
        'Set of attachments changed, so need reencoding work
        m_AttachedFilesChanged = True

    End If
    
Next intCounter

End Sub

Public Sub ListFiles()

Dim intCounter As Integer

If m_HaveAttachment Then

    For intCounter = 0 To UBound(m_FilesAttached)
    
        'Raise event that report the file
        RaiseEvent FilesInQueue(m_FilesAttached(intCounter), intCounter)
    
    Next intCounter

End If

End Sub

Public Sub RemoveFile(intFileIndex As Integer)

Dim strTempFiles() As String

Dim intCounter As Integer


'Check the index for validity
If intFileIndex < LBound(m_FilesAttached) Or intFileIndex > UBound(m_FilesAttached) Then
    Exit Sub
End If



If UBound(m_FilesAttached) = 0 Then
    'If it is the last element, just erase the array
    Erase m_FilesAttached
    'No files in queue, hence no attachment
    
    m_HaveAttachment = False
    
Else
    'Redim the temporary array less one element compared to the file's queue array
    ReDim strTempFiles(UBound(m_FilesAttached))
    
    'Copy all values minus the one to be removed.
    For intCounter = 0 To UBound(m_FilesAttached)
        If intCounter < intFileIndex Then
            strTempFiles(intCounter) = m_FilesAttached(intCounter)
        ElseIf intCounter > intFileIndex Then
            strTempFiles(intCounter - 1) = m_FilesAttached(intCounter)
        End If
    Next intCounter
    
    'Redimension the queue array
    ReDim m_FilesAttached(UBound(strTempFiles) + 1)
    
    'Assign back all values
    m_FilesAttached = strTempFiles

End If

'Set of attachments changed, so need reencoding work
m_AttachedFilesChanged = True

End Sub

Public Sub ClearFiles()

'Eaasy...just erase the array, n flip the switch

Erase m_FilesAttached

m_HaveAttachment = False

End Sub

'Class functions

Private Sub EncodeAndSend(Optional blnUseBase64 As Boolean = True)


'TODO: Sort of a flag/switch or whatsoever, so that if same set attachments
'      is sent again, the temp file is not deleted, and no unnecessary encoding work.


Dim strTempPath     As String
Dim intFileNum      As Integer
Dim intCounter      As Integer
Dim strBuffer       As String


'On error goto Erh:

'Get the app's path
strTempPath = App.Path
If Right$(strTempPath, 1) <> "\" Then strTempPath = strTempPath & "\"

'Kill any stupid files out there...
If Dir(strTempPath & "Temp.dat") <> "" Then Kill strTempPath & "Temp.dat"

If m_AttachedFilesChanged Then

    'Encode all files in the queue
    If blnUseBase64 Then
        If UBound(m_FilesAttached) = 0 Then
            'The file is the only file anyway
            Base64Encode m_FilesAttached(0), True, strTempPath & "Temp.dat", True
        Else
            For intCounter = 0 To UBound(m_FilesAttached) - 1
                Base64Encode m_FilesAttached(intCounter), True, strTempPath & "Temp.dat"
            Next
            'Mark the file as the last file
            Base64Encode m_FilesAttached(UBound(m_FilesAttached)), True, strTempPath & "Temp.dat", True
        End If
    Else
        For intCounter = 0 To UBound(m_FilesAttached)
            UUEncode m_FilesAttached(intCounter), True, strTempPath & "Temp.dat"
        Next
    End If

End If

'Open the file and send bit by bit
intFileNum = FreeFile
Open strTempPath & "Temp.dat" For Binary As intFileNum

 Do Until EOF(intFileNum)
 
  strBuffer = Space$(8192)
  Get #intFileNum, , strBuffer
  
  'Send the chunk
  SMTPSock.SendData strBuffer
  
  'TODO: Some sort of timeout mechanism
  
 Loop
 
 'Send the end of file
 SMTPSock.SendData vbCrLf
  
 'Close the file
 Close intFileNum
 
 'Files sent successfully

Exit Sub

Erh:

End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'                                        SMTP Data Send
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Private Sub SMTPSendData(strData As String)

'Send data to the server.
SMTPSock.SendData strData & vbCrLf



End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'                                        SMTP Connection Timeout
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'Some might want to set different timeout duration for different phase/commands, which can be
'done here.

Private Sub TimeOut()

Dim sngCurrentTime As Single
sngCurrentTime = Timer

blnTimeout = True

Do
    DoEvents
Loop Until blnTimeout = False Or (Timer > (sngCurrentTime + sngTimeout))

If blnTimeout = True Then

    RaiseEvent SMTPError(ERROR_TIMEOUT, "Connection Timeout")
    
    'Close socket
    SMTPSock.Close
    
End If


End Sub

Public Sub Terminate()

If SMTPSock Is Nothing Then

Else
    If SMTPSock.State <> sckClosed Then
        SMTPSock.Close
    End If
End If

End Sub


'=============================================================================================
'=============================================================================================
'                                       Class Properties
'=============================================================================================
'=============================================================================================


Public Property Get POPAuthorisation() As Boolean
    POPAuthorisation = m_POPAuthorisation
End Property

Public Property Let POPAuthorisation(ByVal newValue As Boolean)
    m_POPAuthorisation = newValue
End Property



Public Property Let SMTPSocket(ByVal sckMail As Winsock)
    Set SMTPSock = sckMail
End Property

Public Property Let POPSocket(ByVal sckMail As Winsock)
    Set POPSock = sckMail
End Property

Public Property Get POPConnected() As Boolean
    POPConnected = m_POPConnected
End Property

Public Property Let POPConnected(ByVal newValue As Boolean)
    m_POPConnected = newValue
End Property

Public Property Get SMTPConnected() As Boolean
    SMTPConnected = m_SMTPConnected
End Property

Public Property Let SMTPConnected(ByVal newValue As Boolean)
    m_SMTPConnected = newValue
End Property

Public Property Get CCAddress() As String
    CCAddress = m_CarbonCopy
End Property

Public Property Let CCAddress(ByVal newValue As String)
    m_CarbonCopy = newValue
End Property

Public Property Get MessageBody() As String
    MessageBody = m_MessageBody
End Property

Public Property Let MessageBody(ByVal newValue As String)
    m_MessageBody = newValue
End Property

Public Property Get Subject() As String
    Subject = m_Subject
End Property

Public Property Let Subject(ByVal newValue As String)
    m_Subject = newValue
End Property

Public Property Get RecepientAddress() As String
    RecepientAddress = m_RecepientAddress
End Property

Public Property Let RecepientAddress(ByVal newValue As String)
    m_RecepientAddress = newValue
End Property

Public Property Get SenderAddress() As String
    SenderAddress = m_SenderAddress
End Property

Public Property Let SenderAddress(ByVal newValue As String)
    m_SenderAddress = newValue
End Property

Public Property Get EmailAddress() As String
    EmailAddress = m_EmailAddress
End Property

Public Property Let EmailAddress(ByVal newValue As String)
    m_EmailAddress = newValue
End Property

Public Property Get Username() As String
    Username = m_Username
End Property

Public Property Let Username(ByVal newValue As String)
    m_Username = newValue
End Property

Property Get POPPort() As Integer
    POPPort = m_POPPort
End Property

Property Let POPPort(ByVal newValue As Integer)
    m_POPPort = newValue
End Property

Public Property Get POPHostname() As String
    POPHostname = m_POPHostname
End Property

Public Property Let POPHostname(ByVal newValue As String)
    m_POPHostname = newValue
End Property

Public Property Get SMTPPort() As Integer
    SMTPPort = m_SMTPPort
End Property

Public Property Let SMTPPort(ByVal newValue As Integer)
    m_SMTPPort = newValue
End Property

Public Property Get SMTPHostname() As String
    SMTPHostname = m_SMTPHostname
End Property

Public Property Let SMTPHostname(ByVal newValue As String)
    m_SMTPHostname = newValue
End Property


Private Sub Class_Initialize()

m_ReadyForSend = True

End Sub


'=============================================================================================
'=============================================================================================
'                                       SMTP/POP Sockets Routines
'=============================================================================================
'=============================================================================================

Private Sub SMTPSock_Close()

'Force close
If SMTPSock.State <> sckClosed Then
    SMTPSock.Close
    
    m_ReadyForSend = False
    SMTPConnected = False
End If

End Sub

Private Sub SMTPSock_Connect()

'Mark connected
SMTPConnected = True

MsgBox "Connected"

End Sub

Private Sub SMTPSock_DataArrival(ByVal bytesTotal As Long)

Dim strData As String

SMTPSock.GetData strData, vbString

MsgBox strData

SMTPProcessData strData


End Sub

Private Sub SMTPSock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

SMTPConnected = False

If SMTPSock.State <> sckClosed Then
    SMTPSock.Close
    
    m_ReadyForSend = False
    SMTPConnected = False
End If

End Sub

'                                               END OF CODES
'----------------------------------------------------------------------------------------
