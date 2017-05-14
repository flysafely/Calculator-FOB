Attribute VB_Name = "Module1"
Public Function myRound(ByVal sglN As String, ByRef lngW As Long) As String
On Error GoTo err1
        '四舍五入函数
        'sglN   '要计算的数值
        'lngN   '将要保留的小数位数
        Dim lngN     As Long       '字符总长
        Dim lngD     As Long       '记录小数点位置
        Dim lngC     As Long       '小数位数
        Dim sglX     As String       '小数点后lngW-1位以前的数字
        Dim lngX2     As Long         '保存lngW位的数字(要保留的小数最未位)
        Dim lngX3     As Long         '保存lngW+1位的数字(要舍去的小数第一位)
        
        '计算小数点位置
        lngD = InStr(sglN, ".")
        lngN = Len(sglN)
        
        If lngD = 0 Then
                myRound = sglN
        Else
                sglX = Left(sglN, lngD + (lngW - 1))
                lngC = lngN - lngD
                If lngC > lngW Then
                        lngX2 = Mid(sglN, lngD + lngW, 1)
                        lngX3 = Mid(sglN, lngD + lngW + 1, 1)
                        If lngX3 > 4 Then lngX2 = lngX2 + 1
                        
                        If lngW = 1 Then
                                myRound = sglX & lngX2
                        Else
                                myRound = sglX & lngX2
                        End If
                Else
                        myRound = sglN
                End If
        End If
        
Exit Function
err1:
myRound = sglN
End Function
