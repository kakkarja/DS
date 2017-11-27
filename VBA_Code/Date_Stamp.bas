Attribute VB_Name = "Module1"
Option Explicit

'''START DATE - STAMP'''

Sub DStamp()
Set TXW = ActiveSheet
    On Error GoTo Oboy
    ActiveSheet.Activate
    
    If TXW.ProtectContents = False Then
        KalenderS.Show vbModeless
    Else
        MsgBox "This Add-In is functioning." _
            & " Please try again in new" _
            & " Worksheet.", , "Calendar Stamp"
    End If
Oboy:
Set TXW = Nothing
End Sub


'''USERFORM DATE - STAMP (CALENDAR 2017-2021)'''

Dim i As Long
Dim j As Long
Dim k As Long
Dim D As Long
Dim Bul() As Variant
Dim Hit() As Variant
Dim BVal() As Variant
Dim H() As Variant
Dim M As String
Dim Lab As Variant
Dim rg As Range
Dim Tx As String
Dim CD As Date
Dim DT As String
Dim Def As Date, CDe As Long
Private Sub D1_Click()
    If RangeC = "" Then
        D1 = False
    Else
        RangeC = Format(CDe, "DD/MM/YYYY")
        DT = Format(CDe, "DD/MM/YY")
        RangeC = DT
    End If
    Stmp
End Sub

Private Sub D2_Click()
    If RangeC = "" Then
        D2 = False
    Else
        RangeC = Format(CDe, "DD/MM/YYYY")
        DT = Format(CDe, "DD, MMMM yyyy")
        RangeC = DT
    End If
    Stmp
End Sub

Private Sub D3_Click()
    If RangeC = "" Then
        D3 = False
    Else
        RangeC = Format(CDe, "DD/MM/YYYY")
        DT = Format(CDe, "DD-MM-YYYY")
        RangeC = DT
    End If
    Stmp
End Sub

Private Sub D4_Click()
    If RangeC = "" Then
        D4 = False
    Else
        RangeC = Format(CDe, "DD/MM/YYYY")
        DT = Format(CDe, "D/M/YY")
        RangeC = DT
    End If
    Stmp
End Sub

Private Sub D5_Click()
    If RangeC = "" Then
        D5 = False
    Else
        RangeC = Format(CDe, "DD/MM/YYYY")
        DT = Format(CDe, "DDDD, DD/MM/YYYY")
        RangeC = DT
    End If
    Stmp
End Sub

Private Sub D6_Click()
    If RangeC = "" Then
        D6 = False
    Else
        RangeC = Format(CDe, "DD/MM/YYYY")
        If CDate(RangeC) = Date Then
            DT = Format(CDe, "DD/MM/YYYY") & " " & Format(Time, "hh:mm")
            RangeC = DT
        Else
            RangeC = "Works only for today date"
            GoTo bye
        End If
    End If
    Stmp
bye:
End Sub

Private Sub D7_Click()
    If RangeC = "" Then
        D7 = False
    Else
        RangeC = Format(CDe, "DD/MM/YYYY")
        DT = Format(CDe, "DDDD, DD MMM 'YY")
        RangeC = DT
    End If
    Stmp
End Sub

Private Sub Label1_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label1.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label1.Caption))
        CDe = CLng(CD)
        RangeC = Format(CDe, "DD/MM/YYYY")
    End If
    CRB
    Stmp
End Sub

Private Sub Label10_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label10.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label10.Caption))
        CDe = CLng(CD)
        RangeC = Format(CDe, "DD/MM/YYYY")
    End If
    CRB
    Stmp
End Sub

Private Sub Label11_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label11.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label11.Caption))
        CDe = CLng(CD)
        RangeC = Format(CDe, "DD/MM/YYYY")
    End If
    CRB
    Stmp
End Sub

Private Sub Label12_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label12.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label12.Caption))
        CDe = CLng(CD)
        RangeC = Format(CDe, "DD/MM/YYYY")
    End If
    CRB
    Stmp
End Sub

Private Sub Label13_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label13.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label13.Caption))
        CDe = CLng(CD)
        RangeC = Format(CDe, "DD/MM/YYYY")
    End If
    CRB
    Stmp
End Sub

Private Sub Label14_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label14.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label14.Caption))
        CDe = CLng(CD)
        RangeC = Format(CDe, "DD/MM/YYYY")
    End If
    CRB
    Stmp
End Sub

Private Sub Label15_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label15.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label15.Caption))
        CDe = CLng(CD)
        RangeC = Format(CDe, "DD/MM/YYYY")
    End If
    CRB
    Stmp
End Sub

Private Sub Label16_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label16.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label16.Caption))
        CDe = CLng(CD)
        RangeC = Format(CDe, "DD/MM/YYYY")
    End If
    CRB
    Stmp
End Sub

Private Sub Label17_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label17.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label17.Caption))
        CDe = CLng(CD)
        RangeC = Format(CDe, "DD/MM/YYYY")
    End If
    CRB
    Stmp
End Sub

Private Sub Label18_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label18.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label18.Caption))
        CDe = CLng(CD)
        RangeC = Format(CDe, "DD/MM/YYYY")
    End If
    CRB
    Stmp
End Sub

Private Sub Label19_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label19.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label19.Caption))
        CDe = CLng(CD)
        RangeC = Format(CDe, "DD/MM/YYYY")
    End If
    CRB
    Stmp
End Sub

Private Sub Label2_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label2.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label2.Caption))
        CDe = CLng(CD)
        RangeC = Format(CDe, "DD/MM/YYYY")
    End If
    CRB
    Stmp
End Sub

Private Sub Label20_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label20.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label20.Caption))
        CDe = CLng(CD)
        RangeC = Format(CDe, "DD/MM/YYYY")
    End If
    CRB
    Stmp
End Sub

Private Sub Label21_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label21.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label21.Caption))
        CDe = CLng(CD)
        RangeC = Format(CDe, "DD/MM/YYYY")
    End If
    CRB
    Stmp
End Sub

Private Sub Label22_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label22.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label22.Caption))
        CDe = CLng(CD)
        RangeC = Format(CDe, "DD/MM/YYYY")
    End If
    CRB
    Stmp
End Sub

Private Sub Label23_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label23.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label23.Caption))
        CDe = CLng(CD)
        RangeC = Format(CDe, "DD/MM/YYYY")
    End If
    CRB
    Stmp
End Sub

Private Sub Label24_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label24.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label24.Caption))
        CDe = CLng(CD)
        RangeC = Format(CDe, "DD/MM/YYYY")
    End If
    CRB
    Stmp
End Sub

Private Sub Label25_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label25.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label25.Caption))
        CDe = CLng(CD)
        RangeC = Format(CDe, "DD/MM/YYYY")
    End If
    CRB
    Stmp
End Sub

Private Sub Label26_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label26.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label26.Caption))
        CDe = CLng(CD)
        RangeC = Format(CDe, "DD/MM/YYYY")
    End If
    CRB
    Stmp
End Sub

Private Sub Label27_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label27.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label27.Caption))
        CDe = CLng(CD)
        RangeC = Format(CDe, "DD/MM/YYYY")
    End If
    CRB
    Stmp
End Sub

Private Sub Label28_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label28.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label28.Caption))
        CDe = CLng(CD)
        RangeC = Format(CDe, "DD/MM/YYYY")
    End If
    CRB
    Stmp
End Sub

Private Sub Label29_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label29.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label29.Caption))
        CDe = CLng(CD)
        RangeC = Format(CDe, "DD/MM/YYYY")
    End If
    CRB
    Stmp
End Sub

Private Sub Label3_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label3.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label3.Caption))
        CDe = CLng(CD)
        RangeC = Format(CDe, "DD/MM/YYYY")
    End If
    CRB
    Stmp
End Sub

Private Sub Label30_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label30.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label30.Caption))
        CDe = CLng(CD)
        RangeC = Format(CDe, "DD/MM/YYYY")
    End If
    CRB
    Stmp
End Sub

Private Sub Label31_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label31.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label31.Caption))
        CDe = CLng(CD)
        RangeC = Format(CDe, "DD/MM/YYYY")
    End If
    CRB
    Stmp
End Sub

Private Sub Label32_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label32.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label32.Caption))
        CDe = CLng(CD)
        RangeC = Format(CDe, "DD/MM/YYYY")
    End If
    CRB
    Stmp
End Sub

Private Sub Label33_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label33.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label33.Caption))
        CDe = CLng(CD)
        RangeC = Format(CDe, "DD/MM/YYYY")
    End If
    CRB
    Stmp
End Sub

Private Sub Label34_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label34.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label34.Caption))
        CDe = CLng(CD)
        RangeC = Format(CDe, "DD/MM/YYYY")
    End If
    CRB
    Stmp
End Sub

Private Sub Label35_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label35.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label35.Caption))
        CDe = CLng(CD)
        RangeC = Format(CDe, "DD/MM/YYYY")
    End If
    CRB
    Stmp
End Sub

Private Sub Label36_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label36.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label36.Caption))
        CDe = CLng(CD)
        RangeC = Format(CDe, "DD/MM/YYYY")
    End If
    CRB
    Stmp
End Sub

Private Sub Label37_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label37.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label37.Caption))
        CDe = CLng(CD)
        RangeC = Format(CDe, "DD/MM/YYYY")
    End If
    CRB
    Stmp
End Sub

Private Sub Label38_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label38.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label38.Caption))
        CDe = CLng(CD)
        RangeC = Format(CDe, "DD/MM/YYYY")
    End If
    CRB
    Stmp
End Sub

Private Sub Label39_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label39.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label39.Caption))
        CDe = CLng(CD)
        RangeC = Format(CDe, "DD/MM/YYYY")
    End If
    CRB
    Stmp
End Sub

Private Sub Label4_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label4.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label4.Caption))
        CDe = CLng(CD)
        RangeC = Format(CDe, "DD/MM/YYYY")
    End If
    CRB
    Stmp
End Sub

Private Sub Label40_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label40.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label40.Caption))
        CDe = CLng(CD)
        RangeC = Format(CDe, "DD/MM/YYYY")
    End If
    CRB
    Stmp
End Sub

Private Sub Label41_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label41.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label41.Caption))
        CDe = CLng(CD)
        RangeC = Format(CDe, "DD/MM/YYYY")
    End If
    CRB
    Stmp
End Sub

Private Sub Label42_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label42.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label42.Caption))
        CDe = CLng(CD)
        RangeC = Format(CDe, "DD/MM/YYYY")
    End If
    CRB
    Stmp
End Sub

Private Sub Label5_Click()
            Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label5.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label5.Caption))
        CDe = CLng(CD)
        RangeC = Format(CDe, "DD/MM/YYYY")
    End If
    CRB
    Stmp
End Sub

Private Sub Label6_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label6.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label6.Caption))
        CDe = CLng(CD)
        RangeC = Format(CDe, "DD/MM/YYYY")
    End If
    CRB
    Stmp
End Sub

Private Sub Label7_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label7.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label7.Caption))
        CDe = CLng(CD)
        RangeC = Format(CDe, "DD/MM/YYYY")
    End If
    CRB
    Stmp
End Sub

Private Sub Label8_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label8.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label8.Caption))
        CDe = CLng(CD)
        RangeC = Format(CDe, "DD/MM/YYYY")
    End If
    CRB
    Stmp
End Sub

Private Sub Label9_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label9.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label9.Caption))
        CDe = CLng(CD)
        RangeC = Format(CDe, "DD/MM/YYYY")
    End If
    CRB
    Stmp
End Sub

Private Sub RangeC_Change()
    'If IsDate(RangeC) = True Then
    '    Def = Format(RangeC, "dd/mm/yyyy;@")
    'End If
End Sub

Private Sub RangeC_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    RangeC = Replace(ActiveCell.Address, "$", "")
End Sub

Private Sub SOnOf_Change()
    If SOnOf = True Then
        SOnOf.Caption = "Sound ON"
    Else
        SOnOf.Caption = "Sound OFF"
    End If
End Sub

Private Sub SpinButton1_Change()
    Select Case Years
        Case Is = "2017"
            Y2017
        Case Is = "2018"
            Y2018
        Case Is = "2019"
            Y2019
        Case Is = "2020"
            Y2020
        Case Is = "2021"
            Y2021
        Case Else
            Years = Year(Date)
    End Select
End Sub

Private Sub DD()
    For j = 1 To 42
        Controls("label" & j).Caption = ""
    Next j
    j = 0
End Sub

Private Sub Stmp()
    If Not RangeC = "" Then
        'On Error Resume Next
        ActiveCell.NumberFormat = "@"
        If D1 = True Then
            'D1 = False
            ActiveCell.Value = DT
            j = ActiveCell.Column
            Cells.Columns(j).AutoFit
            j = 0
            DT = ""
        ElseIf D2 = True Then
            'D2 = False
            ActiveCell.Value = DT
            j = ActiveCell.Column
            Cells.Columns(j).AutoFit
            j = 0
            DT = ""
        ElseIf D3 = True Then
            'D3 = False
            ActiveCell.Value = DT
            j = ActiveCell.Column
            Cells.Columns(j).AutoFit
            j = 0
            DT = ""
        ElseIf D4 = True Then
            'D4 = False
            ActiveCell.Value = DT
            j = ActiveCell.Column
            Cells.Columns(j).AutoFit
            j = 0
            DT = ""
        ElseIf D5 = True Then
            'D5 = False
            ActiveCell.Value = DT
            j = ActiveCell.Column
            Cells.Columns(j).AutoFit
            j = 0
            DT = ""
        ElseIf D6 = True Then
            'D6 = False
            If DT <> Format(Now, "DD/MM/YYYY HH:MM") Then
                DT = ""
                GoTo Good
            End If
            ActiveCell.Value = DT
            j = ActiveCell.Column
            Cells.Columns(j).AutoFit
            j = 0
            DT = ""
        ElseIf D7 = True Then
            'D7 = False
            ActiveCell.Value = DT
            j = ActiveCell.Column
            Cells.Columns(j).AutoFit
            j = 0
            DT = ""
        Else
            RangeC = Format(CDe, "DD/MM/YYYY")
            ActiveCell.Value = Format(CDe, "DD/MM/YYYY")
            j = ActiveCell.Column
            Cells.Columns(j).AutoFit
            j = 0
        End If
        On Error GoTo Good
        'CDe = CLng(Def)
            If CDate(CDe) = Date Then
                If SOnOf = False Then
                    GoTo Good
                Else
                    Application.Speech. _
                    Speak Day(CDe) _
                    & MonthName(Month(CDe), False) & _
                    Left(Format(CDe, "YYYY"), 1) & "thousand" & _
                    Format(CDe, "YY") & ", " & Hour(Time) & _
                    " O'clock" & Minute(Time) & "Minutes"
                End If
            Else
                If SOnOf = False Then
                    GoTo Good
                Else
                    Application.Speech. _
                    Speak Day(CDe) _
                    & MonthName(Month(CDe), False) & _
                    Left(Format(CDe, "YYYY"), 1) & "thousand" & _
                    Format(CDe, "YY")
                End If
            End If
        CD = 0
    End If
Good:
CD = 0
End Sub
Private Sub CRB()
Dim RB As Integer
Dim CR As String
    If D1 = True Then
        D1 = False
    ElseIf D2 = True Then
        D2 = False
    ElseIf D3 = True Then
        D3 = False
    ElseIf D4 = True Then
        D4 = False
    ElseIf D5 = True Then
        D5 = False
    ElseIf D6 = True Then
        D6 = False
    ElseIf D7 = True Then
        D7 = False
    End If
End Sub

Private Sub UserForm_Initialize()
    With Years
        .AddItem "2017"
        .AddItem "2018"
        .AddItem "2019"
        .AddItem "2020"
        .AddItem "2021"
    End With
    With Controls
        With .Item("Minggu")
            .Caption = Format(Weekday(1, vbSunday), "ddd")
        End With
        With .Item("Senin")
            .Caption = Format(Weekday(2, vbSunday), "ddd")
        End With
        With .Item("Selasa")
            .Caption = Format(Weekday(3, vbSunday), "ddd")
        End With
        With .Item("Rabu")
            .Caption = Format(Weekday(4, vbSunday), "ddd")
        End With
        With .Item("Kamis")
            .Caption = Format(Weekday(5, vbSunday), "ddd")
        End With
        With .Item("Jumat")
            .Caption = Format(Weekday(6, vbSunday), "ddd")
        End With
        With .Item("Sabtu")
            .Caption = Format(Weekday(7, vbSunday), "ddd")
        End With
    End With
    'Bul = _
    'Array(0, "Januari", "Februari", "Maret", "April" _
    ', "Mei", "Juni", "Juli", "Agustus", "September" _
    ', "Oktober", "November", "Desember")
    RangeC.Locked = True
    M = Format(Date, "m")
    For i = 1 To 12
        If M = i Then
            SpinButton1.Value = i
            SpinButton1_Change
            Exit Sub
        End If
    Next i
End Sub

Private Sub Years_Change()
    With Years
        If .Locked = True Then
            .Locked = False
            SpinButton1_Change
        Else
            SpinButton1_Change
            .Locked = True
        End If
    End With
End Sub

Private Sub Years_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    With Years
        If .Locked = True Then
            .Locked = False
        Else
            .Locked = True
        End If
    End With
End Sub



Private Sub Y2017()
Dim Minggu As Integer, Senin As Integer, Selasa As Integer _
, Rabu As Integer, Kamis As Integer, Jumat As Integer _
, Sabtu As Integer
    BVal = Array _
    (0, 42736, 42767, 42795, 42826, 42856, 42887, 42917, _
    42948, 42979, 43009, 43040, 43070)
    
    Bul = _
    Array(0, Format(42736, "mmmm"), _
    Format(42767, "mmmm"), Format(42795, "mmmm"), _
    Format(42826, "mmmm"), Format(42856, "mmmm"), _
    Format(42887, "mmmm"), Format(42917, "mmmm"), _
    Format(42948, "mmmm"), Format(42979, "mmmm"), _
    Format(43009, "mmmm"), Format(43040, "mmmm"), _
    Format(43070, "mmmm"))
    
    Hit = Array(0, 31, 28, 31, 30, 31, 30, 31, 31, 30, _
    31, 30, 31)
    
    H = Array(0, Format(Weekday(1, vbSunday), "dddd"), _
    Format(Weekday(2, vbSunday), "dddd"), _
    Format(Weekday(3, vbSunday), "dddd"), _
    Format(Weekday(4, vbSunday), "dddd"), _
    Format(Weekday(5, vbSunday), "dddd"), _
    Format(Weekday(6, vbSunday), "dddd"), _
    Format(Weekday(7, vbSunday), "dddd"))
    
    For i = 1 To UBound(Bul)
        If SpinButton1.Value = i Then
            Bulan.Caption = Bul(i)
                For k = 1 To UBound(H)
                Minggu = 1
                Senin = 2
                Selasa = 3
                Rabu = 4
                Kamis = 5
                Jumat = 6
                Sabtu = 7
                        If Format(BVal(i), "dddd") _
                        = H(k) Then
                            Select Case H(k)
                                Case Format(Weekday(1, vbSunday), "dddd")
                                    DD
                                    For D = Minggu To Hit(i)
                                        With Controls("Label" & D)
                                            .Caption = D
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date)).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0

                                    Exit Sub
                                Case Format(Weekday(2, vbSunday), "dddd")
                                    DD
                                    For D = Senin To Hit(i) + 1
                                        With Controls("Label" & D)
                                            .Caption = D - 1
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 1).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0

                                    Exit Sub
                                Case Format(Weekday(3, vbSunday), "dddd")
                                    DD
                                    For D = Selasa To Hit(i) + 2
                                        With Controls("Label" & D)
                                            .Caption = D - 2
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 2).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0
                                    Exit Sub
                                Case Format(Weekday(4, vbSunday), "dddd")
                                    DD
                                    For D = Rabu To Hit(i) + 3
                                        With Controls("Label" & D)
                                            .Caption = D - 3
                                            .Font.Bold = False
                                        End With
                                    Next
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 3).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0
                                    Exit Sub
                                Case Format(Weekday(5, vbSunday), "dddd")
                                    DD
                                    For D = Kamis To Hit(i) + 4
                                        With Controls("Label" & D)
                                            .Caption = D - 4
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 4).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0
                                    Exit Sub
                                Case Format(Weekday(6, vbSunday), "dddd")
                                    DD
                                    For D = Jumat To Hit(i) + 5
                                        With Controls("Label" & D)
                                            .Caption = D - 5
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 5).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0
                                    Exit Sub
                                Case Format(Weekday(7, vbSunday), "dddd")
                                    DD
                                    For D = Sabtu To Hit(i) + 6
                                        With Controls("Label" & D)
                                            .Caption = D - 6
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 6).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0
 
                                    Exit Sub
                            End Select
                            
                        End If
                Next k
        End If
    Next i
End Sub

Private Sub Y2018()
Dim Minggu As Integer, Senin As Integer, Selasa As Integer _
, Rabu As Integer, Kamis As Integer, Jumat As Integer _
, Sabtu As Integer
    BVal = Array _
    (0, 43101, 43132, 43160, 43191, 43221, 43252, 43282, _
    43313, 43344, 43374, 43405, 43435)
    
    Bul = _
    Array(0, Format(43101, "mmmm"), _
    Format(43132, "mmmm"), Format(43160, "mmmm"), _
    Format(43191, "mmmm"), Format(43221, "mmmm"), _
    Format(43252, "mmmm"), Format(43282, "mmmm"), _
    Format(43313, "mmmm"), Format(43344, "mmmm"), _
    Format(43374, "mmmm"), Format(43405, "mmmm"), _
    Format(43435, "mmmm"))
    
    Hit = Array(0, 31, 28, 31, 30, 31, 30, 31, 31, 30, _
    31, 30, 31)
    
    H = Array(0, Format(Weekday(1, vbSunday), "dddd"), _
    Format(Weekday(2, vbSunday), "dddd"), _
    Format(Weekday(3, vbSunday), "dddd"), _
    Format(Weekday(4, vbSunday), "dddd"), _
    Format(Weekday(5, vbSunday), "dddd"), _
    Format(Weekday(6, vbSunday), "dddd"), _
    Format(Weekday(7, vbSunday), "dddd"))
    
    For i = 1 To UBound(Bul)
        If SpinButton1.Value = i Then
            Bulan.Caption = Bul(i)
                For k = 1 To UBound(H)
                Minggu = 1
                Senin = 2
                Selasa = 3
                Rabu = 4
                Kamis = 5
                Jumat = 6
                Sabtu = 7
                        If Format(BVal(i), "dddd") _
                        = H(k) Then
                            Select Case H(k)
                                Case Format(Weekday(1, vbSunday), "dddd")
                                    DD
                                    For D = Minggu To Hit(i)
                                        With Controls("Label" & D)
                                            .Caption = D
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date)).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0

                                    Exit Sub
                                Case Format(Weekday(2, vbSunday), "dddd")
                                    DD
                                    For D = Senin To Hit(i) + 1
                                        With Controls("Label" & D)
                                            .Caption = D - 1
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 1).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0

                                    Exit Sub
                                Case Format(Weekday(3, vbSunday), "dddd")
                                    DD
                                    For D = Selasa To Hit(i) + 2
                                        With Controls("Label" & D)
                                            .Caption = D - 2
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 2).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0
                                    Exit Sub
                                Case Format(Weekday(4, vbSunday), "dddd")
                                    DD
                                    For D = Rabu To Hit(i) + 3
                                        With Controls("Label" & D)
                                            .Caption = D - 3
                                            .Font.Bold = False
                                        End With
                                    Next
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 3).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0
                                    Exit Sub
                                Case Format(Weekday(5, vbSunday), "dddd")
                                    DD
                                    For D = Kamis To Hit(i) + 4
                                        With Controls("Label" & D)
                                            .Caption = D - 4
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 4).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0
                                    Exit Sub
                                Case Format(Weekday(6, vbSunday), "dddd")
                                    DD
                                    For D = Jumat To Hit(i) + 5
                                        With Controls("Label" & D)
                                            .Caption = D - 5
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 5).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0
                                    Exit Sub
                                Case Format(Weekday(7, vbSunday), "dddd")
                                    DD
                                    For D = Sabtu To Hit(i) + 6
                                        With Controls("Label" & D)
                                            .Caption = D - 6
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 6).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0
 
                                    Exit Sub
                            End Select
                            
                        End If
                Next k
        End If
    Next i
End Sub
Private Sub Y2019()
Dim Minggu As Integer, Senin As Integer, Selasa As Integer _
, Rabu As Integer, Kamis As Integer, Jumat As Integer _
, Sabtu As Integer
    BVal = Array _
    (0, 43466, 43497, 43525, 43556, 43586, 43617, 43647, _
    43678, 43709, 43739, 43770, 43800)
    
    Bul = _
    Array(0, Format(43466, "mmmm"), _
    Format(43497, "mmmm"), Format(43525, "mmmm"), _
    Format(43556, "mmmm"), Format(43586, "mmmm"), _
    Format(43617, "mmmm"), Format(43647, "mmmm"), _
    Format(43678, "mmmm"), Format(43709, "mmmm"), _
    Format(43739, "mmmm"), Format(43770, "mmmm"), _
    Format(43800, "mmmm"))
    
    Hit = Array(0, 31, 28, 31, 30, 31, 30, 31, 31, 30, _
    31, 30, 31)
    
    H = Array(0, Format(Weekday(1, vbSunday), "dddd"), _
    Format(Weekday(2, vbSunday), "dddd"), _
    Format(Weekday(3, vbSunday), "dddd"), _
    Format(Weekday(4, vbSunday), "dddd"), _
    Format(Weekday(5, vbSunday), "dddd"), _
    Format(Weekday(6, vbSunday), "dddd"), _
    Format(Weekday(7, vbSunday), "dddd"))
    
    For i = 1 To UBound(Bul)
        If SpinButton1.Value = i Then
            Bulan.Caption = Bul(i)
                For k = 1 To UBound(H)
                Minggu = 1
                Senin = 2
                Selasa = 3
                Rabu = 4
                Kamis = 5
                Jumat = 6
                Sabtu = 7
                        If Format(BVal(i), "dddd") _
                        = H(k) Then
                            Select Case H(k)
                                Case Format(Weekday(1, vbSunday), "dddd")
                                    DD
                                    For D = Minggu To Hit(i)
                                        With Controls("Label" & D)
                                            .Caption = D
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date)).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0

                                    Exit Sub
                                Case Format(Weekday(2, vbSunday), "dddd")
                                    DD
                                    For D = Senin To Hit(i) + 1
                                        With Controls("Label" & D)
                                            .Caption = D - 1
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 1).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0

                                    Exit Sub
                                Case Format(Weekday(3, vbSunday), "dddd")
                                    DD
                                    For D = Selasa To Hit(i) + 2
                                        With Controls("Label" & D)
                                            .Caption = D - 2
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 2).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0
                                    Exit Sub
                                Case Format(Weekday(4, vbSunday), "dddd")
                                    DD
                                    For D = Rabu To Hit(i) + 3
                                        With Controls("Label" & D)
                                            .Caption = D - 3
                                            .Font.Bold = False
                                        End With
                                    Next
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 3).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0
                                    Exit Sub
                                Case Format(Weekday(5, vbSunday), "dddd")
                                    DD
                                    For D = Kamis To Hit(i) + 4
                                        With Controls("Label" & D)
                                            .Caption = D - 4
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 4).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0
                                    Exit Sub
                                Case Format(Weekday(6, vbSunday), "dddd")
                                    DD
                                    For D = Jumat To Hit(i) + 5
                                        With Controls("Label" & D)
                                            .Caption = D - 5
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 5).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0
                                    Exit Sub
                                Case Format(Weekday(7, vbSunday), "dddd")
                                    DD
                                    For D = Sabtu To Hit(i) + 6
                                        With Controls("Label" & D)
                                            .Caption = D - 6
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 6).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0
 
                                    Exit Sub
                            End Select
                            
                        End If
                Next k
        End If
    Next i
End Sub

Private Sub Y2020()
Dim Minggu As Integer, Senin As Integer, Selasa As Integer _
, Rabu As Integer, Kamis As Integer, Jumat As Integer _
, Sabtu As Integer
    BVal = Array _
    (0, 43831, 43862, 43891, 43922, _
    43952, 43983, 44013, 44044, 44075 _
    , 44105, 44136, 44166)
    
    Bul = _
    Array(0, Format(43831, "mmmm"), _
    Format(43862, "mmmm"), Format(43891, "mmmm"), _
    Format(43922, "mmmm"), Format(43952, "mmmm"), _
    Format(43983, "mmmm"), Format(44013, "mmmm"), _
    Format(44044, "mmmm"), Format(44075, "mmmm"), _
    Format(44105, "mmmm"), Format(44136, "mmmm"), _
    Format(44166, "mmmm"))
    
    Hit = Array(0, 31, 28, 31, 30, 31, 30, 31, 31, 30, _
    31, 30, 31)
    
    H = Array(0, Format(Weekday(1, vbSunday), "dddd"), _
    Format(Weekday(2, vbSunday), "dddd"), _
    Format(Weekday(3, vbSunday), "dddd"), _
    Format(Weekday(4, vbSunday), "dddd"), _
    Format(Weekday(5, vbSunday), "dddd"), _
    Format(Weekday(6, vbSunday), "dddd"), _
    Format(Weekday(7, vbSunday), "dddd"))
    
    For i = 1 To UBound(Bul)
        If SpinButton1.Value = i Then
            Bulan.Caption = Bul(i)
                For k = 1 To UBound(H)
                Minggu = 1
                Senin = 2
                Selasa = 3
                Rabu = 4
                Kamis = 5
                Jumat = 6
                Sabtu = 7
                        If Format(BVal(i), "dddd") _
                        = H(k) Then
                            Select Case H(k)
                                Case Format(Weekday(1, vbSunday), "dddd")
                                    DD
                                    For D = Minggu To Hit(i)
                                        With Controls("Label" & D)
                                            .Caption = D
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date)).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0

                                    Exit Sub
                                Case Format(Weekday(2, vbSunday), "dddd")
                                    DD
                                    For D = Senin To Hit(i) + 1
                                        With Controls("Label" & D)
                                            .Caption = D - 1
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 1).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0

                                    Exit Sub
                                Case Format(Weekday(3, vbSunday), "dddd")
                                    DD
                                    For D = Selasa To Hit(i) + 2
                                        With Controls("Label" & D)
                                            .Caption = D - 2
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 2).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0
                                    Exit Sub
                                Case Format(Weekday(4, vbSunday), "dddd")
                                    DD
                                    For D = Rabu To Hit(i) + 3
                                        With Controls("Label" & D)
                                            .Caption = D - 3
                                            .Font.Bold = False
                                        End With
                                    Next
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 3).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0
                                    Exit Sub
                                Case Format(Weekday(5, vbSunday), "dddd")
                                    DD
                                    For D = Kamis To Hit(i) + 4
                                        With Controls("Label" & D)
                                            .Caption = D - 4
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 4).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0
                                    Exit Sub
                                Case Format(Weekday(6, vbSunday), "dddd")
                                    DD
                                    For D = Jumat To Hit(i) + 5
                                        With Controls("Label" & D)
                                            .Caption = D - 5
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 5).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0
                                    Exit Sub
                                Case Format(Weekday(7, vbSunday), "dddd")
                                    DD
                                    For D = Sabtu To Hit(i) + 6
                                        With Controls("Label" & D)
                                            .Caption = D - 6
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 6).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0
 
                                    Exit Sub
                            End Select
                            
                        End If
                Next k
        End If
    Next i
End Sub

Private Sub Y2021()
Dim Minggu As Integer, Senin As Integer, Selasa As Integer _
, Rabu As Integer, Kamis As Integer, Jumat As Integer _
, Sabtu As Integer
    BVal = Array _
    (0, 44197, 44228, 44256, 44287, 44317, 44348, _
    44378, 44409, 44440, 44470, 44501, 44531)
    
    Bul = _
    Array(0, Format(44197, "mmmm"), _
    Format(44228, "mmmm"), Format(44256, "mmmm"), _
    Format(44287, "mmmm"), Format(44317, "mmmm"), _
    Format(44348, "mmmm"), Format(44378, "mmmm"), _
    Format(44409, "mmmm"), Format(44440, "mmmm"), _
    Format(44470, "mmmm"), Format(44501, "mmmm"), _
    Format(44531, "mmmm"))
    
    Hit = Array(0, 31, 28, 31, 30, 31, 30, 31, 31, 30, _
    31, 30, 31)
    
    H = Array(0, Format(Weekday(1, vbSunday), "dddd"), _
    Format(Weekday(2, vbSunday), "dddd"), _
    Format(Weekday(3, vbSunday), "dddd"), _
    Format(Weekday(4, vbSunday), "dddd"), _
    Format(Weekday(5, vbSunday), "dddd"), _
    Format(Weekday(6, vbSunday), "dddd"), _
    Format(Weekday(7, vbSunday), "dddd"))
    
    For i = 1 To UBound(Bul)
        If SpinButton1.Value = i Then
            Bulan.Caption = Bul(i)
                For k = 1 To UBound(H)
                Minggu = 1
                Senin = 2
                Selasa = 3
                Rabu = 4
                Kamis = 5
                Jumat = 6
                Sabtu = 7
                        If Format(BVal(i), "dddd") _
                        = H(k) Then
                            Select Case H(k)
                                Case Format(Weekday(1, vbSunday), "dddd")
                                    DD
                                    For D = Minggu To Hit(i)
                                        With Controls("Label" & D)
                                            .Caption = D
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date)).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0

                                    Exit Sub
                                Case Format(Weekday(2, vbSunday), "dddd")
                                    DD
                                    For D = Senin To Hit(i) + 1
                                        With Controls("Label" & D)
                                            .Caption = D - 1
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 1).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0

                                    Exit Sub
                                Case Format(Weekday(3, vbSunday), "dddd")
                                    DD
                                    For D = Selasa To Hit(i) + 2
                                        With Controls("Label" & D)
                                            .Caption = D - 2
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 2).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0
                                    Exit Sub
                                Case Format(Weekday(4, vbSunday), "dddd")
                                    DD
                                    For D = Rabu To Hit(i) + 3
                                        With Controls("Label" & D)
                                            .Caption = D - 3
                                            .Font.Bold = False
                                        End With
                                    Next
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 3).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0
                                    Exit Sub
                                Case Format(Weekday(5, vbSunday), "dddd")
                                    DD
                                    For D = Kamis To Hit(i) + 4
                                        With Controls("Label" & D)
                                            .Caption = D - 4
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 4).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0
                                    Exit Sub
                                Case Format(Weekday(6, vbSunday), "dddd")
                                    DD
                                    For D = Jumat To Hit(i) + 5
                                        With Controls("Label" & D)
                                            .Caption = D - 5
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 5).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0
                                    Exit Sub
                                Case Format(Weekday(7, vbSunday), "dddd")
                                    DD
                                    For D = Sabtu To Hit(i) + 6
                                        With Controls("Label" & D)
                                            .Caption = D - 6
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 6).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0
 
                                    Exit Sub
                            End Select
                            
                        End If
                Next k
        End If
    Next i
End Sub

