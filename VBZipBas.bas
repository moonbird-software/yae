Attribute VB_Name = "VBZipBas"
Option Explicit
'---------------------------------------------------
' Sample VB 5 code to drive zip32.dll
' Contributed to the Info-Zip project by Mike Le Voi
'
' Contact me at: mlevoi@modemss.brisnet.org.au
'
' Visit my home page at: http://modemss.brisnet.org.au/~mlevoi
'
' Use this code at your own risk. Nothing implied or warranted
' to work on your machine :-)
'---------------------------------------------------

'argv
Public Type ZIPnames
    s(0 To 99) As String
End Type

'ZPOPT is used to set options in the zip32.dll
Private Type ZPOPT
    fSuffix As Long
    fEncrypt As Long
    fSystem As Long
    fVolume As Long
    fExtra As Long
    fNoDirEntries As Long
    fExcludeDate As Long
    fIncludeDate As Long
    fVerbose As Long
    fQuiet As Long
    fCRLF_LF As Long
    fLF_CRLF As Long
    fJunkDir As Long
    fRecurse As Long
    fGrow As Long
    fForce As Long
    fMove As Long
    fDeleteEntries As Long
    fUpdate As Long
    fFreshen As Long
    fJunkSFX As Long
    fLatestTime As Long
    fComment As Long
    fOffsets As Long
    fPrivilege As Long
    fEncryption As Long
    fRepair As Long
    flevel As Byte
    date As String ' 8 bytes long
    szRootDir As String ' up to 256 bytes long
End Type

Private Type ZIPUSERFUNCTIONS
    DLLPrnt As Long
    DLLPASSWORD As Long
    DLLCOMMENT As Long
    DLLSERVICE As Long
End Type

'Structure ZCL - not used by VB
'Private Type ZCL
'    argc As Long            'number of files
'    filename As String      'Name of the Zip file
'    fileArray As ZIPnames   'The array of filenames
'End Type

' Call back "string" (sic)
Private Type CBChar
    ch(4096) As Byte
End Type

'Local declares
Dim MYOPT As ZPOPT
' Dim MYZCL As ZCL
Dim MYUSER As ZIPUSERFUNCTIONS

'This assumes zip32.dll is in your \windows\system directory!
Private Declare Function ZpInit Lib "zip32.dll" _
(ByRef Zipfun As ZIPUSERFUNCTIONS) As Long ' Set Zip Callbacks

Private Declare Function ZpSetOptions Lib "zip32.dll" _
(ByRef Opts As ZPOPT) As Long ' Set Zip options

Private Declare Function ZpGetOptions Lib "zip32.dll" _
() As ZPOPT ' used to check encryption flag only

Private Declare Function ZpArchive Lib "zip32.dll" _
(ByVal argc As Long, ByVal funame As String, ByRef argv As ZIPnames) As Long ' Real zipping action

Global vbzipinf As String, crlf$

' Puts a function pointer in a structure
Function FnPtr(ByVal lp As Long) As Long
    FnPtr = lp
End Function

' Callback for zip32.dll
Function DLLPrnt(ByRef fname As CBChar, ByVal x As Long) As Long
    Dim s0$, xx As Long

    ' always put this in callback routines!
    On Error Resume Next
    s0 = ""
    For xx = 0 To x
        If fname.ch(xx) = 0 Then xx = 99999 Else s0 = s0 + Chr(fname.ch(xx))
    Next xx
    ' vbzipinf = vbzipinf + s0
    DoEvents
    DLLPrnt = 0
End Function

' Callback for zip32.dll
Function DllPass(ByRef s1 As Byte, x As Long, _
    ByRef s2 As Byte, _
    ByRef s3 As Byte) As Long

    ' always put this in callback routines!
    On Error Resume Next
    ' not supported - always return 1
    DllPass = 1
End Function

' Callback for zip32.dll
Function DllComm(ByRef s1 As CBChar) As CBChar
    
    ' always put this in callback routines!
    On Error Resume Next
    ' not supported always return \0
    s1.ch(0) = vbNullString
    DllComm = s1
End Function

'Main Subroutine
Function VBZip(argc As Integer, zipname As String, _
        mynames As ZIPnames, junk As Integer, _
        recurse As Integer, updat As Integer, _
        freshen As Integer, basename As String) As Long
    
    Dim hmem As Long, xx As Integer
    Dim retcode As Long
    
    On Error Resume Next ' nothing will go wrong :-)
    
    ' Set address of callback functions
    MYUSER.DLLPrnt = FnPtr(AddressOf DLLPrnt)
    MYUSER.DLLPASSWORD = FnPtr(AddressOf DllPass)
    MYUSER.DLLCOMMENT = FnPtr(AddressOf DllComm)
    MYUSER.DLLSERVICE = 0& ' not coded yet :-)
    retcode = ZpInit(MYUSER)
    
    ' Set zip options
    MYOPT.fSuffix = 0        ' include suffixes (not yet implemented)
    MYOPT.fEncrypt = 0       ' 1 if encryption wanted
    MYOPT.fSystem = 0        ' 1 to include system/hidden files
    MYOPT.fVolume = 0        ' 1 if storing volume label
    MYOPT.fExtra = 0         ' 1 if including extra attributes
    MYOPT.fNoDirEntries = 0  ' 1 if ignoring directory entries
    MYOPT.fExcludeDate = 0   ' 1 if excluding files earlier than a specified date
    MYOPT.fIncludeDate = 0   ' 1 if including files earlier than a specified date
    MYOPT.fVerbose = 0       ' 1 if full messages wanted
    MYOPT.fQuiet = 0         ' 1 if minimum messages wanted
    MYOPT.fCRLF_LF = 0       ' 1 if translate CR/LF to LF
    MYOPT.fLF_CRLF = 0       ' 1 if translate LF to CR/LF
    MYOPT.fJunkDir = junk    ' 1 if junking directory names
    MYOPT.fRecurse = recurse ' 1 if recursing into subdirectories
    MYOPT.fGrow = 0          ' 1 if allow appending to zip file
    MYOPT.fForce = 0         ' 1 if making entries using DOS names
    MYOPT.fMove = 0          ' 1 if deleting files added or updated
    MYOPT.fDeleteEntries = 0 ' 1 if files passed have to be deleted
    MYOPT.fUpdate = updat    ' 1 if updating zip file--overwrite only if newer
    MYOPT.fFreshen = freshen ' 1 if freshening zip file--overwrite only
    MYOPT.fJunkSFX = 0       ' 1 if junking sfx prefix
    MYOPT.fLatestTime = 0    ' 1 if setting zip file time to time of latest file in archive
    MYOPT.fComment = 0       ' 1 if putting comment in zip file
    MYOPT.fOffsets = 0       ' 1 if updating archive offsets for sfx Files
    MYOPT.fPrivilege = 0     ' 1 if not saving privelages
    MYOPT.fEncryption = 0    'Read only property!
    MYOPT.fRepair = 0        ' 1=> fix archive, 2=> try harder to fix
    MYOPT.flevel = 0         ' compression level - should be 0!!!
    MYOPT.date = vbNullString ' "12/31/79"? US Date?
    MYOPT.szRootDir = basename
    
    ' Set options
    retcode = ZpSetOptions(MYOPT)
    
    ' ZCL not needed in VB
    ' MYZCL.argc = 2
    ' MYZCL.filename = "c:\wiz\new.zip"
    ' MYZCL.fileArray = MYNAMES
    
    ' Go for it!
    retcode = ZpArchive(argc, zipname, mynames)
    
    VBZip = retcode
End Function

