Attribute VB_Name = "Module1"
Public Function myRound(ByVal sglN As String, ByRef lngW As Long) As String
On Error GoTo err1
        '�������뺯��
        'sglN   'Ҫ�������ֵ
        'lngN   '��Ҫ������С��λ��
        Dim lngN     As Long       '�ַ��ܳ�
        Dim lngD     As Long       '��¼С����λ��
        Dim lngC     As Long       'С��λ��
        Dim sglX     As String       'С�����lngW-1λ��ǰ������
        Dim lngX2     As Long         '����lngWλ������(Ҫ������С����δλ)
        Dim lngX3     As Long         '����lngW+1λ������(Ҫ��ȥ��С����һλ)
        
        '����С����λ��
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
