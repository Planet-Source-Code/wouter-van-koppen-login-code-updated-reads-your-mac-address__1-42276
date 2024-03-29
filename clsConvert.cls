VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const NCBASTAT = &H33
Private Const NCBNAMSZ = 16
Private Const HEAP_ZERO_MEMORY = &H8
Private Const HEAP_GENERATE_EXCEPTIONS = &H4
Private Const NCBRESET = &H32
Private Const EM_SETPASSWORDCHAR = &HCC

Private Type NCB
    ncb_command          As Byte
    ncb_retcode          As Byte
    ncb_lsn              As Byte
    ncb_num              As Byte
    ncb_buffer           As Long
    ncb_length           As Integer
    ncb_callname         As String * NCBNAMSZ
    ncb_name             As String * NCBNAMSZ
    ncb_rto              As Byte
    ncb_sto              As Byte
    ncb_post             As Long
    ncb_lana_num         As Byte
    ncb_cmd_cplt         As Byte
    ncb_reserve(9)       As Byte
    ncb_event            As Long
End Type

Private Type ADAPTER_STATUS
    adapter_address(5)   As Byte
    rev_major            As Byte
    reserved0            As Byte
    adapter_type         As Byte
    rev_minor            As Byte
    duration             As Integer
    frmr_recv            As Integer
    frmr_xmit            As Integer
    iframe_recv_err      As Integer
    xmit_aborts          As Integer
    xmit_success         As Long
    recv_success         As Long
    iframe_xmit_err      As Integer
    recv_buff_unavail    As Integer
    t1_timeouts          As Integer
    ti_timeouts          As Integer
    Reserved1            As Long
    free_ncbs            As Integer
    max_cfg_ncbs         As Integer
    max_ncbs             As Integer
    xmit_buf_unavail     As Integer
    max_dgram_size       As Integer
    pending_sess         As Integer
    max_cfg_sess         As Integer
    max_sess             As Integer
    max_sess_pkt_size    As Integer
    name_count           As Integer
End Type

Private Type NAME_BUFFER
    name                 As String * NCBNAMSZ
    name_num             As Integer
    name_flags           As Integer
End Type

Private Type ASTAT
    adapt                As ADAPTER_STATUS
    NameBuff(30)         As NAME_BUFFER
End Type


Private Declare Function Netbios Lib "netapi32.dll" (pncb As NCB) As Byte

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)

Private Declare Function GetProcessHeap Lib "kernel32" () As Long

Private Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, _
        ByVal dwFlags As Long, ByVal dwBytes As Long) As Long

Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, _
        ByVal dwFlags As Long, lpMem As Any) As Long


'I'm sorry that the code is less explained but I
'never do this
'And I'm dutch my english isn't very good
'FEEL FREE TO STEAL AND DISTRIBUTE - but please
'vote for me!!!

Public Function UsercodeToKeycode(UserCode As String) As String
    Dim Length As Integer
    Dim teller As Integer
    Dim KeyCode As String
    Dim AscCode As Integer
    Length = Len(UserCode)
    Do While teller <= Length - 1
        teller = teller + 1
        AscCode = 500 - Val("&H" & Asc(Mid(UserCode, teller, 1)))
        KeyCode = KeyCode & AscCode & "-"
    Loop
    KeyCode = Left(KeyCode, Len(KeyCode) - 1)
    UsercodeToKeycode = KeyCode
End Function

Public Function KeycodetoUsercode(KeyCode As String) As String
    Dim TextByNow As String
    Dim sign1 As Integer
    Dim sign2 As Integer
    Dim noMore As Boolean
    sign1 = 1
    sign2 = 1
    KeyCode = KeyCode & "-"
    Do While noMore = False
        sign2 = InStr(sign2, KeyCode, "-")
        If sign2 = Len(KeyCode) Then
            noMore = True
        End If
        TextByNow = TextByNow & Chr$(500 - Hex(Mid(KeyCode, sign1, sign2 - sign1)))
        sign1 = sign2 + 1
        sign2 = sign2 + 1
    Loop
    KeycodetoUsercode = TextByNow
End Function

Public Sub CreateKey(Folder As String, Value As String)
    Dim b As Object
    On Error Resume Next
    Set b = CreateObject("wscript.shell")
    b.RegWrite Folder, Value
End Sub

Public Sub CreateIntegerKey(Folder As String, Value As Integer)
    Dim b As Object
    On Error Resume Next
    Set b = CreateObject("wscript.shell")
    b.RegWrite Folder, Value, "REG_DWORD"
End Sub

Public Function ReadKey(Value As String) As String
    Dim b As Object
    On Error Resume Next
    Set b = CreateObject("wscript.shell")
    ReadKey = b.RegRead(Value)
End Function

Public Sub DeleteKey(Value As String)
    Dim b As Object
    On Error Resume Next
    Set b = CreateObject("Wscript.Shell")
    b.RegDelete Value
End Sub

Public Function GetMACAddress(LanaNumber As Long) As String
    Dim udtNCB       As NCB
    Dim bytResponse  As Byte
    Dim udtASTAT     As ASTAT
    Dim udtTempASTAT As ASTAT
    Dim lngASTAT     As Long
    Dim strOut       As String
    Dim X            As Integer
    
    udtNCB.ncb_command = NCBRESET
    bytResponse = Netbios(udtNCB)
    udtNCB.ncb_command = NCBASTAT
    udtNCB.ncb_lana_num = LanaNumber
    udtNCB.ncb_callname = "* "
    udtNCB.ncb_length = Len(udtASTAT)
    lngASTAT = HeapAlloc(GetProcessHeap(), HEAP_GENERATE_EXCEPTIONS Or HEAP_ZERO_MEMORY, udtNCB.ncb_length)
    strOut = ""
    If lngASTAT Then
        udtNCB.ncb_buffer = lngASTAT
        bytResponse = Netbios(udtNCB)
        CopyMemory udtASTAT, udtNCB.ncb_buffer, Len(udtASTAT)
        With udtASTAT.adapt
            For X = 0 To 5
                strOut = strOut & Right$("00" & Hex$(.adapter_address(X)), 2)
            Next X
        End With
        HeapFree GetProcessHeap(), 0, lngASTAT
    End If
    GetMACAddress = strOut
End Function


