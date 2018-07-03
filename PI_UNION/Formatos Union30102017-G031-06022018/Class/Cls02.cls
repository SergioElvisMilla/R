VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cls02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Enum Enum_ReportOriention
    OrientUndefined = 0
    OrientPortrait = vbPRORPortrait
    OrientLandscape = vbPRORLandscape
End Enum
Private Type DEVMODE
    dmDeviceName    As String * 32
    dmSpecVersion   As Integer
    dmDriverVersion As Integer
    dmSize          As Integer
    dmDriverExtra   As Integer
    dmFields        As Long
    dmOrientation   As Integer
    dmPaperSize     As Integer
    dmPaperLength   As Integer
    dmPaperWidth    As Integer
    dmScale         As Integer
    dmCopies        As Integer
    dmDefaultSource As Integer
    dmPrintQuality  As Integer
    dmColor         As Integer
    dmDuplex        As Integer
    dmYResolution   As Integer
    dmTTOption      As Integer
    dmCollate       As Integer
    dmFormName      As String * 32
    dmUnusedPadding As Integer
    dmBitsPerPel    As Integer
    dmPelsWidth     As Long
    dmPelsHeight    As Long
    dmDisplayFlags  As Long
    dmDisplayFrequency As Long
End Type
Private Type PRINTER_DEFAULTS
    pDatatype       As String
    pDevMode        As Long
    DesiredAccess   As Long
End Type
Private Type PRINTER_INFO_2
    pServerName     As Long
    pPrinterName    As Long
    pShareName      As Long
    pPortName       As Long
    pDriverName     As Long
    pComment        As Long
    pLocation       As Long
    pDevMode        As Long
    pSepFile        As Long
    pPrintProcessor As Long
    pDatatype       As Long
    pParameters     As Long
    pSecurityDescriptor As Long
    Attributes      As Long
    Priority        As Long
    DefaultPriority As Long
    StartTime       As Long
    UntilTime       As Long
    Status          As Long
    cJobs           As Long
    AveragePPM      As Long
End Type
Private Const DM_IN_BUFFER      As Long = 8
Private Const DM_OUT_BUFFER     As Long = 2
Private Const DM_ORIENTATION    As Long = &H1
Private Const NULLPTR           As Long = 0&
Private Const PRINTER_ACCESS_ADMINISTER  As Long = &H4
Private Const PRINTER_ACCESS_USE         As Long = &H8
Private Const STANDARD_RIGHTS_REQUIRED   As Long = &HF0000
Private Const PRINTER_ALL_ACCESS         As Long = (STANDARD_RIGHTS_REQUIRED Or PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As Any) As Long
Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal hWnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, pDevModeOutput As Any, pDevModeInput As Any, ByVal fMode As Long) As Long
Private Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal cbBuf As Long, pcbNeeded As Long) As Long
Private Declare Function SetPrinter Lib "winspool.drv" Alias "SetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal Command As Long) As Long
Public Function SetPrinterOrientation(ByVal eOrientation As Enum_ReportOriention) As Boolean
    On Error GoTo Solucion
    Dim bDevMode()      As Byte
    Dim bPrinterInfo2() As Byte
    Dim hPrinter        As Long
    Dim lResult         As Long
    Dim nSize           As Long
    Dim sPrnName        As String
    Dim dm              As DEVMODE
    Dim pd              As PRINTER_DEFAULTS
    Dim pi2             As PRINTER_INFO_2
    sPrnName = Printer.DeviceName
    pd.DesiredAccess = PRINTER_ALL_ACCESS
    If OpenPrinter(sPrnName, hPrinter, pd) Then
        Call GetPrinter(hPrinter, 2&, 0&, 0&, nSize)
        ReDim bPrinterInfo2(1 To nSize) As Byte
        lResult = GetPrinter(hPrinter, 2, bPrinterInfo2(1), nSize, nSize)
        Call CopyMemory(pi2, bPrinterInfo2(1), Len(pi2))
        nSize = DocumentProperties(0&, hPrinter, sPrnName, 0&, 0&, 0)
        ReDim bDevMode(1 To nSize)
        If pi2.pDevMode Then
           Call CopyMemory(bDevMode(1), ByVal pi2.pDevMode, Len(dm))
        Else
           Call DocumentProperties(0&, hPrinter, sPrnName, bDevMode(1), 0&, DM_OUT_BUFFER)
        End If
        Call CopyMemory(dm, bDevMode(1), Len(dm))
        With dm
            .dmOrientation = eOrientation
            .dmFields = DM_ORIENTATION
        End With
        Call CopyMemory(bDevMode(1), dm, Len(dm))
        Call DocumentProperties(0&, hPrinter, sPrnName, bDevMode(1), bDevMode(1), DM_IN_BUFFER Or DM_OUT_BUFFER)
        pi2.pDevMode = VarPtr(bDevMode(1))
        lResult = SetPrinter(hPrinter, 2, pi2, 0&)
        Call ClosePrinter(hPrinter)
        SetPrinterOrientation = True
    Else
        SetPrinterOrientation = False
    End If
    Exit Function
Solucion:
    MsgBox "No se puede acceder a las propiedades de la impresora", vbCritical, "Error"
End Function
Public Function GetPrinterOrientation(DeviceName As String, hDC As Long) As Enum_ReportOriention
    On Error GoTo Solucion
    Dim hPrinter    As Long
    Dim nSize       As Long
    Dim pDevMode    As DEVMODE
    Dim aDevMode()  As Byte
    If OpenPrinter(DeviceName, hPrinter, NULLPTR) Then
       nSize = DocumentProperties(NULLPTR, hPrinter, DeviceName, NULLPTR, NULLPTR, 0)
       ReDim aDevMode(1 To nSize)
       nSize = DocumentProperties(NULLPTR, hPrinter, DeviceName, aDevMode(1), NULLPTR, DM_OUT_BUFFER)
       Call CopyMemory(pDevMode, aDevMode(1), Len(pDevMode))
       GetPrinterOrientation = pDevMode.dmOrientation
       Call ClosePrinter(hPrinter)
    Else
       GetPrinterOrientation = OrientUndefined
    End If
    Exit Function
Solucion:
    MsgBox "No se puede acceder a las propiedades de la impresora", vbCritical, "Error"
End Function
