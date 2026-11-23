Option Explicit
' ============================================================
' Key auto : คีย์ผลนับจากเว็บ Count-Weight ลงฟอร์ม STINV010 (Inventory Countings)
' วิธีใช้: กดปุ่ม "Key auto" ในเว็บ -> ได้ไฟล์ keys-xxx.xlsx -> รันมาโคร Key_Auto -> เลือกไฟล์นั้น
' ============================================================
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
    Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#End If

' ===== ปรับค่าตรงนี้ให้ตรงกับการคีย์จริงของฟอร์ม =====
' ลำดับการคีย์ส่วนหัว (ใช้ {ST}=Stocktake {EMP}=Employee {AISLE} {BAY}) และปุ่ม {TAB} / {ENTER}
Const HEADER_KEYS As String = "{ST}{TAB}{EMP}{TAB}{AISLE}{TAB}{BAY}{TAB}"
' ลำดับการคีย์แต่ละบรรทัด (ใช้ {ART}=Article {QTY}=Quantity)
Const ROW_KEYS As String = "{ART}{TAB}{QTY}{TAB}"
Const START_DELAY As Long = 5000      ' เวลาให้คลิกช่องแรกของฟอร์ม (มิลลิวินาที)
Const KEY_DELAY As Long = 150         ' หน่วงหลังคีย์แต่ละบรรทัด
Const PAUSE_NEW_BAY As Boolean = True ' True = หยุดถามก่อนขึ้น Bay ใหม่ (ให้ตั้งฟอร์มเอง)

Sub Key_Auto()
    Dim fn As Variant, wb As Workbook, v As Variant
    Dim i As Long, n As Long, bays As Long
    Dim st As String, emp As String
    Dim k As String, curKey As String, prevKey As String
    Dim hdr As String, rw As String

    fn = Application.GetOpenFilename("Excel (*.xlsx),*.xlsx", , "เลือกไฟล์คีย์ที่ดาวน์โหลดจากเว็บ Count-Weight")
    If VarType(fn) = vbBoolean Then Exit Sub

    Set wb = Workbooks.Open(CStr(fn), ReadOnly:=True)
    v = wb.Sheets(1).UsedRange.Value
    wb.Close SaveChanges:=False

    n = UBound(v, 1) - 1
    If n < 1 Then MsgBox "ไม่พบข้อมูลในไฟล์": Exit Sub

    For i = 2 To UBound(v, 1)
        k = BayKey(v, i)
        If k <> prevKey Then bays = bays + 1: prevKey = k
    Next i

    st = InputBox("เลข Stocktake:", "Key auto")
    If st = "" Then Exit Sub
    emp = InputBox("รหัสพนักงาน (Employee):", "Key auto")
    If emp = "" Then Exit Sub

    If MsgBox("พร้อมคีย์ " & n & " บรรทัด (" & bays & " Bay)" & vbCrLf & _
              "กด OK แล้วคลิกช่องแรกของฟอร์ม STINV010 ภายใน " & START_DELAY / 1000 & " วินาที" & vbCrLf & _
              "(กดค้างปุ่ม Esc เพื่อหยุด)", vbOKCancel, "Key auto") = vbCancel Then Exit Sub
    Sleep START_DELAY

    For i = 2 To UBound(v, 1)
        k = BayKey(v, i)
        If k <> curKey Then
            If curKey <> "" And PAUSE_NEW_BAY Then
                If MsgBox("ขึ้น Bay ใหม่ " & k & vbCrLf & "ตั้งฟอร์มให้พร้อมแล้วกด OK จากนั้นคลิกช่องแรกภายใน " & _
                          START_DELAY / 1000 & " วินาที", vbOKCancel, "Key auto") = vbCancel Then Exit For
                Sleep START_DELAY
            End If
            curKey = k
            hdr = HEADER_KEYS
            hdr = Replace(hdr, "{ST}", SK(st))
            hdr = Replace(hdr, "{EMP}", SK(emp))
            hdr = Replace(hdr, "{AISLE}", SK(Trim(CStr(v(i, 1)))))
            hdr = Replace(hdr, "{BAY}", SK(Trim(CStr(v(i, 2)))))
            Application.SendKeys hdr, True
            DoEvents
            Sleep KEY_DELAY
        End If

        If (GetAsyncKeyState(&H1B) And &H8000) <> 0 Then
            MsgBox "หยุดการคีย์แล้ว (คีย์ไปแล้ว " & i - 2 & " บรรทัด)"
            Exit For
        End If

        rw = ROW_KEYS
        rw = Replace(rw, "{ART}", SK(Trim(CStr(v(i, 3)))))
        rw = Replace(rw, "{QTY}", SK(Replace(CStr(Round(CDbl(v(i, 4)), 3)), ",", ".")))
        Application.SendKeys rw, True
        DoEvents
        Sleep KEY_DELAY
    Next i

    Application.SendKeys "{NUMLOCK}", True   ' เหมือนโค้ดเดิม (คืนค่า NumLock) ถ้า NumLock เพี้ยนให้ลบบรรทัดนี้
    DoEvents
    MsgBox "คีย์เสร็จแล้ว", vbInformation, "Key auto"
End Sub

Private Function BayKey(v As Variant, i As Long) As String
    BayKey = Trim(CStr(v(i, 1))) & "-" & Trim(CStr(v(i, 2))) & "-" & Trim(CStr(v(i, 6)))
End Function

' ครอบอักขระพิเศษของ SendKeys ( + ^ % ~ ( ) { } [ ] )
Private Function SK(ByVal s As String) As String
    Dim i As Long, c As String, r As String
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If InStr("+^%~(){}[]", c) > 0 Then r = r & "{" & c & "}" Else r = r & c
    Next i
    SK = r
End Function
