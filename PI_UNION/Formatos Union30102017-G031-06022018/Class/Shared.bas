Attribute VB_Name = "Shared"
Public Function xVal(Expression As String) As Double
    If Trim$(Expression) = "" Then Expression = 0
    xVal = Val(Replace(Expression, ",", ""))
End Function

Public Function xTrim(ByVal sCad As String, Optional ByVal IncCom As Boolean = False) As String
    If IncCom Then
        xTrim = "'" & Trim$("" & sCad) & "'"
    Else
        xTrim = Trim$("" & sCad)
    End If
End Function

Public Function GetCampoCab(ByVal ICnx As ADODB.Connection, ByVal CodDoc As String, ByVal NumDoc As String, ByVal CodEmp As String, ByVal Campo As String, ByVal sCondicion As String) As String
    Dim rsTmp As ADODB.Recordset, sArcCab As String
    Set rsTmp = ICnx.Execute("Select CodDoc,ArcCab,ArcDet From CfgDoc00 Where CodDoc = " & xTrim(CodDoc) & " And CodEmp = " & xTrim(CodEmp), , adCmdText)
    If Not (rsTmp.EOF And rsTmp.BOF) Then
        sArcCab = xTrim(rsTmp.Fields("ArcCab").Value)
        Call Close_RecordSet(rsTmp)
    Else
        Exit Function
    End If
    If xTrim(sArcCab) = "" Then Exit Function
    Set rsTmp = ICnx.Execute("Select " & xTrim(Campo) & vbCrLf & _
                             "From " & xTrim(sArcCab) & vbCrLf & _
                             "Where CodDoc = " & xTrim(CodDoc, True) & " And NumDoc = " & xTrim(NumDoc, True) & IIf(xTrim(sCondicion) = "", "", xTrim(sCondicion)), , adCmdText)
    If Not (rsTmp.EOF And rsTmp.BOF) Then
        GetCampoCab = Trim$("" & rsTmp.Fields(0).Value)
    End If
    Call Close_RecordSet(rsTmp)
End Function

Public Sub Close_RecordSet(ByRef Rs As ADODB.Recordset)
    If Not Rs Is Nothing Then
        If Rs.State <> adStateClosed Then Rs.Close
        Set Rs = Nothing
    End If
End Sub


