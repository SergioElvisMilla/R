VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cls05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

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

Private Type PRINTER_INFO_5
        pPrinterName As String
        pPortName As String
        Attributes As Long
        DeviceNotSelectedTimeout As Long
        TransmissionRetryTimeout As Long
End Type
Const PRINTER_ATTRIBUTE_QUEUED = &H1
Const PRINTER_ATTRIBUTE_DIRECT = &H2
Const PRINTER_ATTRIBUTE_DEFAULT = &H4
Const PRINTER_ATTRIBUTE_SHARED = &H8
Const PRINTER_ATTRIBUTE_NETWORK = &H10
Const PRINTER_ATTRIBUTE_HIDDEN = &H20
Const PRINTER_ATTRIBUTE_LOCAL = &H40
Private Type PRINTER_DEFAULTS
        pDatatype As String
        pDevMode As DEVMODE
        DesiredAccess As Long
End Type


Private Const DM_IN_BUFFER      As Long = 8
Private Const DM_OUT_BUFFER     As Long = 2
Private Const DM_ORIENTATION    As Long = &H1
Private Const DM_PAPERSIZE = &H2&
Private Const DM_PAPERLENGTH = &H4&
Private Const DM_PAPERWIDTH = &H8&
Private Const DM_SCALE = &H10&
Private Const DM_XX1 = &H20&
Private Const DM_XX2 = &H40&
Private Const DM_XX3 = &H80&

'32 64 128 TrackDefault

Private Const DM_COPIES = &H100&
Private Const DM_DEFAULTSOURCE = &H200&
Private Const DM_PRINTQUALITY = &H400&
Private Const DM_COLOR = &H800&
Private Const DM_DUPLEX = &H1000&
Private Const DM_YRESOLUTION = &H2000&
Private Const DM_TTOPTION = &H4000&
Private Const DM_COLLATE As Long = &H8000
Private Const DM_FORMNAME As Long = &H10000



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


Private bCnfAnt As Boolean, sDeviceAnt As String, sCnfDmAnt As DEVMODE, oPrinter As Printer, sDevicePre As String
Public Function SetPrinterOrientation(ByVal sDevice As String, ByVal eOrientation As Enum_ReportOriention3, Optional iPaperSize As Integer, Optional iScale As Integer, Optional iPaperWidth As Integer, Optional iPaperLength As Integer) As Boolean
    On Error GoTo Solucion
    Dim bDevMode()      As Byte
    Dim bPrinterInfo2() As Byte
    Dim hPrinter        As Long
    Dim lResult         As Long
    Dim nSize           As Long
    Dim dm              As DEVMODE
    Dim pd              As PRINTER_DEFAULTS
    Dim pi2             As PRINTER_INFO_2
    Dim iCnt            As Integer
    pd.DesiredAccess = PRINTER_ALL_ACCESS
    If OpenPrinter(sDevice, hPrinter, pd) Then
        Call GetPrinter(hPrinter, 2&, 0&, 0&, nSize)
        ReDim bPrinterInfo2(1 To nSize) As Byte
        lResult = GetPrinter(hPrinter, 2, bPrinterInfo2(1), nSize, nSize)
        Call CopyMemory(pi2, bPrinterInfo2(1), Len(pi2))
        nSize = DocumentProperties(0&, hPrinter, sDevice, 0&, 0&, 0)
        ReDim bDevMode(1 To nSize)
        If pi2.pDevMode Then
           Call CopyMemory(bDevMode(1), ByVal pi2.pDevMode, Len(dm))
        Else
           Call DocumentProperties(0&, hPrinter, sDevice, bDevMode(1), 0&, DM_OUT_BUFFER)
        End If
        Call CopyMemory(dm, bDevMode(1), Len(dm))
        If Not bCnfAnt Then
            sDeviceAnt = Printer.DeviceName
            With sCnfDmAnt
                .dmDeviceName = dm.dmDeviceName: .dmSpecVersion = dm.dmSpecVersion: .dmDriverVersion = dm.dmDriverVersion: .dmSize = dm.dmSize: .dmDriverExtra = dm.dmDriverExtra: .dmFields = dm.dmFields: .dmOrientation = dm.dmOrientation: .dmPaperSize = dm.dmPaperSize: .dmPaperLength = dm.dmPaperLength: .dmPaperWidth = dm.dmPaperWidth: .dmScale = dm.dmScale: .dmCopies = dm.dmCopies: .dmDefaultSource = dm.dmDefaultSource: .dmPrintQuality = dm.dmPrintQuality: .dmColor = dm.dmColor: .dmDuplex = dm.dmDuplex: .dmYResolution = dm.dmYResolution: .dmTTOption = dm.dmTTOption: .dmCollate = dm.dmCollate: .dmFormName = dm.dmFormName: .dmUnusedPadding = dm.dmUnusedPadding: .dmBitsPerPel = dm.dmBitsPerPel: .dmPelsWidth = dm.dmPelsWidth: .dmPelsHeight = dm.dmPelsHeight: .dmDisplayFlags = dm.dmDisplayFlags: .dmDisplayFrequency = dm.dmDisplayFrequency
            End With
        End If
        With dm
            If iPaperWidth > 0 Then
                .dmOrientation = eOrientation
                .dmPaperSize = iPaperSize
                .dmScale = iScale
                .dmPaperWidth = iPaperWidth
                .dmPaperLength = iPaperLength
                .dmFields = DM_ORIENTATION Or DM_PAPERSIZE Or DM_PAPERLENGTH Or DM_PAPERWIDTH Or DM_SCALE
            ElseIf iPaperSize > 0 Then
                .dmOrientation = eOrientation
                .dmPaperSize = iPaperSize
                .dmFields = DM_ORIENTATION Or DM_PAPERSIZE
            Else
                .dmOrientation = eOrientation
                .dmFields = DM_ORIENTATION
            End If
        End With
        If Not bCnfAnt Then sCnfDmAnt.dmFields = dm.dmFields: bCnfAnt = True
        Call CopyMemory(bDevMode(1), dm, Len(dm))
        Call DocumentProperties(0&, hPrinter, sDevice, bDevMode(1), bDevMode(1), DM_IN_BUFFER Or DM_OUT_BUFFER)
        pi2.pDevMode = VarPtr(bDevMode(1))
        If (Trim$(sDevice) <> Trim$(sDeviceAnt)) Then pi2.Attributes = PRINTER_ATTRIBUTE_DEFAULT Or pi2.Attributes
        lResult = SetPrinter(hPrinter, 2, pi2, 0&)
        Call ClosePrinter(hPrinter)
        SetPrinterOrientation = True
    Else
        SetPrinterOrientation = False
    End If
    Call SetPrinterName(sDevice)
    sDevicePre = sDevice
    Exit Function
Solucion:
    MsgBox "No se puede acceder a las propiedades de la impresora", vbCritical, "Error"
End Function
Private Sub SetPrinterName(sDevice As String)
    For Each oPrinter In Printers
        If Trim$(oPrinter.DeviceName) = Trim$(sDevice) Then Set Printer = oPrinter: Exit For
    Next
End Sub
Public Function ReturnCfg() As Boolean
    If Trim$(sDeviceAnt) <> "" Then
        For Each oPrinter In Printers
            If Trim$(oPrinter.DeviceName) = Trim$(sDeviceAnt) Then Set Printer = oPrinter: Exit For
        Next
        On Error GoTo Solucion
        Dim hPrinter As Long, lResult As Long, pd As PRINTER_DEFAULTS, nSize As Long, bPrinterInfo2() As Byte, pi2 As PRINTER_INFO_2, dm As DEVMODE
        pd.DesiredAccess = PRINTER_ALL_ACCESS
        If OpenPrinter(sDevicePre, hPrinter, pd) Then
            Call GetPrinter(hPrinter, 2&, 0&, 0&, nSize)
            ReDim bPrinterInfo2(1 To nSize) As Byte
            lResult = GetPrinter(hPrinter, 2, bPrinterInfo2(1), nSize, nSize)
            Call CopyMemory(pi2, bPrinterInfo2(1), Len(pi2))
            nSize = DocumentProperties(0&, hPrinter, sDevicePre, 0&, 0&, 0)
            ReDim bDevMode(1 To nSize)
            If pi2.pDevMode Then
               Call CopyMemory(bDevMode(1), ByVal pi2.pDevMode, Len(dm))
            Else
               Call DocumentProperties(0&, hPrinter, sDevicePre, bDevMode(1), 0&, DM_OUT_BUFFER)
            End If
            Call CopyMemory(dm, bDevMode(1), Len(dm))
            With dm
                .dmDeviceName = sCnfDmAnt.dmDeviceName: .dmSpecVersion = sCnfDmAnt.dmSpecVersion: .dmDriverVersion = sCnfDmAnt.dmDriverVersion: .dmSize = sCnfDmAnt.dmSize: .dmDriverExtra = sCnfDmAnt.dmDriverExtra: .dmFields = sCnfDmAnt.dmFields: .dmOrientation = sCnfDmAnt.dmOrientation: .dmPaperSize = sCnfDmAnt.dmPaperSize: .dmPaperLength = sCnfDmAnt.dmPaperLength: .dmPaperWidth = sCnfDmAnt.dmPaperWidth: .dmScale = sCnfDmAnt.dmScale: .dmCopies = sCnfDmAnt.dmCopies: .dmDefaultSource = sCnfDmAnt.dmDefaultSource
                .dmPrintQuality = sCnfDmAnt.dmPrintQuality: .dmColor = sCnfDmAnt.dmColor: .dmDuplex = sCnfDmAnt.dmDuplex: .dmYResolution = sCnfDmAnt.dmYResolution: .dmTTOption = sCnfDmAnt.dmTTOption: .dmCollate = sCnfDmAnt.dmCollate: .dmFormName = sCnfDmAnt.dmFormName: .dmUnusedPadding = sCnfDmAnt.dmUnusedPadding: .dmBitsPerPel = sCnfDmAnt.dmBitsPerPel: .dmPelsWidth = sCnfDmAnt.dmPelsWidth: .dmPelsHeight = sCnfDmAnt.dmPelsHeight: .dmDisplayFlags = sCnfDmAnt.dmDisplayFlags: .dmDisplayFrequency = sCnfDmAnt.dmDisplayFrequency
            End With
            Call CopyMemory(bDevMode(1), dm, Len(dm))
            Call DocumentProperties(0&, hPrinter, sDevicePre, bDevMode(1), bDevMode(1), DM_IN_BUFFER Or DM_OUT_BUFFER)
            pi2.pDevMode = VarPtr(bDevMode(1))
            lResult = SetPrinter(hPrinter, 2, pi2, 0&)
            Call ClosePrinter(hPrinter)
            If (Trim$(sDevicePre) <> Trim$(sDeviceAnt)) Then
                ReturnCfg = True And ReturnDfl
            Else
                ReturnCfg = True
            End If
        Else
            ReturnCfg = False
        End If
        sDeviceAnt = ""
        Exit Function
Solucion:
        MsgBox "No se puede acceder a las propiedades de la impresora", vbCritical, "Error"
    End If
End Function
Public Function ReturnDfl() As Boolean
    On Error GoTo Solucion
    Dim bPrinterInfo2() As Byte
    Dim hPrinter        As Long
    Dim lResult         As Long
    Dim nSize           As Long
    Dim pd              As PRINTER_DEFAULTS
    Dim pi2             As PRINTER_INFO_2
    pd.DesiredAccess = PRINTER_ALL_ACCESS
    If OpenPrinter(sDeviceAnt, hPrinter, pd) Then
        Call GetPrinter(hPrinter, 2&, 0&, 0&, nSize)
        ReDim bPrinterInfo2(1 To nSize) As Byte
        lResult = GetPrinter(hPrinter, 2, bPrinterInfo2(1), nSize, nSize)
        Call CopyMemory(pi2, bPrinterInfo2(1), Len(pi2))
        pi2.Attributes = PRINTER_ATTRIBUTE_DEFAULT Or pi2.Attributes
        lResult = SetPrinter(hPrinter, 2, pi2, 0&)
        Call ClosePrinter(hPrinter)
        ReturnDfl = True
    Else
        ReturnDfl = False
    End If
    Exit Function
Solucion:
    MsgBox "No se puede acceder a las propiedades de la impresora", vbCritical, "Error"
End Function
Private Sub Class_Terminate()
    Call ReturnCfg
End Sub
