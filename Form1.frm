VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ʱ��ۼ�����  Version 1.1.1"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7995
   DrawMode        =   8  'Xor Pen
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   ScaleHeight     =   6120
   ScaleWidth      =   7995
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox Text17 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   3480
      TabIndex        =   74
      Text            =   "��������Ʒ����"
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "���ݱ���"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   73
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000010&
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   72
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox Text16 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   600
      MousePointer    =   12  'No Drop
      MultiLine       =   -1  'True
      OLEDragMode     =   1  'Automatic
      TabIndex        =   50
      Top             =   4920
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000010&
      Caption         =   "�� ��"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   48
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox Text14 
      Height          =   375
      Left            =   3720
      TabIndex        =   45
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox Text13 
      Height          =   375
      Left            =   1320
      TabIndex        =   44
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox Text12 
      Height          =   375
      Left            =   3720
      TabIndex        =   43
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   1320
      TabIndex        =   42
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   3720
      TabIndex        =   41
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   840
      TabIndex        =   40
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   3720
      TabIndex        =   39
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   2040
      TabIndex        =   38
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   3720
      TabIndex        =   37
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   2040
      TabIndex        =   36
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000010&
      Caption         =   "��     ��"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   34
      Top             =   4080
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   3720
      TabIndex        =   11
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "���װ���"
      Height          =   735
      Left            =   120
      TabIndex        =   66
      Top             =   1080
      Width           =   1335
      Begin VB.OptionButton Option4 
         Caption         =   "�⾶"
         Height          =   375
         Left            =   720
         TabIndex        =   71
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton Option3 
         Caption         =   "�ھ�"
         Height          =   375
         Left            =   120
         TabIndex        =   70
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "����ѡ��"
      Height          =   615
      Left            =   600
      TabIndex        =   67
      Top             =   3480
      Width           =   1815
      Begin VB.OptionButton Option2 
         Caption         =   "RMB"
         Height          =   255
         Left            =   1080
         TabIndex        =   69
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "USD"
         Height          =   255
         Left            =   240
         TabIndex        =   68
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox Text15 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   600
      TabIndex        =   46
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label Label36 
      Caption         =   "�������ݱ������ⲿTXT��"
      Height          =   255
      Index           =   16
      Left            =   5640
      TabIndex        =   79
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label Label36 
      Caption         =   "���ݱ�����ָ����������"
      Height          =   255
      Index           =   15
      Left            =   5640
      TabIndex        =   78
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Label36 
      Caption         =   "��Ԫ,Ҳ����ÿ������Ԫ"
      Height          =   255
      Index           =   14
      Left            =   5640
      TabIndex        =   77
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label Label36 
      Caption         =   "�˹�������Ŀ����ÿ�ֶ�"
      Height          =   255
      Index           =   13
      Left            =   5640
      TabIndex        =   76
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label10 
      Caption         =   "��Ʒ����"
      Height          =   255
      Left            =   3480
      TabIndex        =   75
      Top             =   4680
      Width           =   855
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      DrawMode        =   1  'Blackness
      Height          =   375
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   5400
      X2              =   5400
      Y1              =   0
      Y2              =   6120
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      Height          =   855
      Left            =   5520
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label Label37 
      Caption         =   "�Ĳ�Ʒ���ۣ���"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5640
      TabIndex        =   65
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Label Label36 
      Caption         =   "һ��,һ�޵ȵ�λ"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   12
      Left            =   5640
      TabIndex        =   64
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label Label36 
      Caption         =   "���Ϊһ��,һ��"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   11
      Left            =   5640
      TabIndex        =   63
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label Label36 
      Caption         =   "��װ����һ�з���"
      Height          =   255
      Index           =   10
      Left            =   5640
      TabIndex        =   62
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label Label36 
      Caption         =   "������,��ǩ,��ֽ,͸��"
      Height          =   255
      Index           =   9
      Left            =   5640
      TabIndex        =   61
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label36 
      Caption         =   "(��ֽ���װ����)�����"
      Height          =   255
      Index           =   6
      Left            =   5640
      TabIndex        =   60
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label36 
      Caption         =   "�ڰ�������ָ��ȥ���װ"
      Height          =   255
      Index           =   3
      Left            =   5640
      TabIndex        =   59
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label36 
      Caption         =   "��Ŀ����д��0��"
      Height          =   255
      Index           =   8
      Left            =   5640
      TabIndex        =   58
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label36 
      Caption         =   "����Һ��,����Һ�����"
      Height          =   255
      Index           =   7
      Left            =   5640
      TabIndex        =   57
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label36 
      Caption         =   "һ�����ж��ٵ�λ��Ʒ"
      Height          =   255
      Index           =   5
      Left            =   5640
      TabIndex        =   56
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label36 
      Caption         =   "�������е�������ָ"
      Height          =   255
      Index           =   4
      Left            =   5640
      TabIndex        =   55
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label36 
      Caption         =   "ÿ�����ж�����/��/Ƭ"
      Height          =   255
      Index           =   2
      Left            =   5640
      TabIndex        =   54
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label36 
      Caption         =   "ÿ��/ÿ��/ÿ��/ÿ��/"
      Height          =   255
      Index           =   1
      Left            =   5640
      TabIndex        =   53
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label36 
      Caption         =   "�ڰ�����е�������ָ"
      Height          =   255
      Index           =   0
      Left            =   5640
      TabIndex        =   52
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label35 
      Caption         =   "��ʷ��¼"
      Height          =   375
      Left            =   720
      TabIndex        =   51
      Top             =   4680
      Width           =   735
   End
   Begin VB.Shape Shape1 
      Height          =   975
      Left            =   480
      Top             =   4800
      Width           =   2775
   End
   Begin VB.Label Label34 
      Caption         =   "Label34"
      Height          =   1095
      Left            =   5520
      TabIndex        =   49
      Top             =   4680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   7200
      Picture         =   "Form1.frx":0ECA
      Top             =   5040
      Width           =   720
   End
   Begin VB.Label Label27 
      BackColor       =   &H00FFFFC0&
      Caption         =   " Copyright 2012-2013 Hangzhou Shenhuang Nonwovens Co.,Ltd                                        by:Authur Xu"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   47
      Top             =   5880
      Width           =   8055
   End
   Begin VB.Label Label33 
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   35
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label Label32 
      Caption         =   "Ԫ/m2"
      Height          =   255
      Left            =   4680
      TabIndex        =   33
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label31 
      Caption         =   "Ԫ/m3"
      Height          =   255
      Left            =   2280
      TabIndex        =   32
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label Label30 
      Caption         =   "Ԫ/��"
      Height          =   255
      Left            =   4680
      TabIndex        =   31
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label29 
      Caption         =   "Ԫ/��"
      Height          =   255
      Left            =   4680
      TabIndex        =   30
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label28 
      Caption         =   "g/m2"
      Height          =   255
      Left            =   1920
      TabIndex        =   29
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label26 
      Caption         =   "Ԫ/��"
      Height          =   255
      Left            =   2280
      TabIndex        =   28
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label25 
      Caption         =   "Ԫ/��or��λ"
      Height          =   255
      Left            =   1800
      TabIndex        =   27
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label24 
      Caption         =   "ֽ�䵥��:"
      Height          =   255
      Left            =   2880
      TabIndex        =   26
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label23 
      Caption         =   "�˷ѵ���:"
      Height          =   255
      Left            =   480
      TabIndex        =   25
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label22 
      Caption         =   "���ϵ���:"
      Height          =   255
      Left            =   2880
      TabIndex        =   24
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label21 
      Caption         =   "�ڰ�����:"
      Height          =   255
      Left            =   480
      TabIndex        =   23
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label20 
      Caption         =   "Һ�����:"
      Height          =   255
      Left            =   2880
      TabIndex        =   22
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label19 
      Caption         =   "�˹�����"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label18 
      Caption         =   "��λ"
      Height          =   255
      Left            =   4680
      TabIndex        =   20
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label17 
      Caption         =   "M"
      Height          =   255
      Left            =   3000
      TabIndex        =   19
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label Label16 
      Caption         =   "M"
      Height          =   255
      Left            =   4680
      TabIndex        =   18
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label15 
      Caption         =   "����:"
      Height          =   255
      Left            =   3240
      TabIndex        =   17
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label14 
      Caption         =   "���:"
      Height          =   255
      Left            =   3240
      TabIndex        =   16
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label13 
      Caption         =   "���:"
      Height          =   255
      Left            =   1560
      TabIndex        =   15
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label12 
      Caption         =   "M"
      Height          =   255
      Left            =   3000
      TabIndex        =   14
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Label11 
      Caption         =   "�䳤:"
      Height          =   255
      Left            =   1560
      TabIndex        =   13
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label9 
      Caption         =   "g/m2"
      Height          =   255
      Left            =   4680
      TabIndex        =   12
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label8 
      Caption         =   "����:"
      Height          =   255
      Left            =   3240
      TabIndex        =   10
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "Ƭ"
      Height          =   255
      Left            =   2520
      TabIndex        =   9
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   "����:"
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "M"
      Height          =   255
      Left            =   4680
      TabIndex        =   6
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "ֽ��:"
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "M"
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "ֽ��:"
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "�ڰ����:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim q As Integer
Dim p As Integer
Dim a As Integer
Dim b As Integer
Dim c As Integer
Dim d As Integer
Dim e As Integer
Dim f As Integer
Dim g As Integer
Dim h As Integer
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim l As Integer
Dim m As Integer
Dim o As Integer
Dim n As Integer

Dim unitworkcost As Double
Dim Deviation1 As Double
Dim Deviation2 As Double
Dim num1 As Double
Dim num2 As Double
Dim num3 As Double
Dim num4 As Double
Dim num5 As Double
Dim num6 As Double
Dim num7 As Double
Dim num8 As Double

Dim num10 As Double

Dim huilv As Double
Dim pcsprice As Double

Dim date1 As String
Dim date2 As String
Dim date3 As String
Dim date4 As String
Dim date5 As String
Dim date6 As String
Dim date7 As String
Dim date8 As String
Dim date9 As String
Dim date10 As String
Dim date11 As String
Dim date12 As String
Dim date13 As String
Dim date14 As String
Dim date15 As String

Dim History As String
Dim clothprice As String
Dim sheetlenght As String
Dim sheetwidth As String
Dim amount As String
Dim amount1 As String
Dim gsm As String
Dim workcost As String
Dim liquidprice As String
Dim innerpackage As String
Dim outerlenght As String
Dim outerwidth As String
Dim outerhigh As String
Dim outerpaperprice As String
Dim yunfei As String


Private Sub Command1_Click()
n = n + 1
 
aa:
sheetlenght = Text1
date1 = Text1
If (IsNumeric(Text1.Text) = False) Then
 a = MsgBox("ֽ������Ϊ����", 1, "���󣡣�")
 If a = vbOK Then
 Text1.Text = InputBox("��������ȷ������.", "ע�⣡")
 GoTo aa
 ElseIf a = vbCancel Then
 Stop
 End If
 End If
 
bb:
sheetwidth = Text2
date2 = Text2
If (IsNumeric(Text2.Text) = False) Then
 b = MsgBox("ֽ�����Ϊ����", 1, "���󣡣�")
 If b = vbOK Then
 Text2.Text = InputBox("��������ȷ������.", "ע�⣡")
 GoTo bb
 ElseIf b = vbCancel Then
 Stop
 End If
 End If
 
cc:
amount = Text3
date3 = Text3
If (IsNumeric(Text3.Text) = False) Then
 c = MsgBox("��������Ϊ����", 1, "���󣡣�")
 If c = vbOK Then
 Text3.Text = InputBox("��������ȷ������.", "ע�⣡")
 GoTo cc
 ElseIf c = vbCancel Then
 Stop
 End If
 End If
 
dd:
gsm = Text4
date4 = Text4
If (IsNumeric(Text4.Text) = False) Then
 d = MsgBox("���ر���Ϊ����", 1, "���󣡣�")
 If d = vbOK Then
 Text4.Text = InputBox("��������ȷ������.", "ע�⣡")
 GoTo dd
 ElseIf d = vbCancel Then
 Stop
 End If
 End If
 
ee:
outerlenght = Text5
date5 = Text5
If (IsNumeric(Text5.Text) = False) Then
 e = MsgBox("�䳤����Ϊ����", 1, "���󣡣�")
 If e = vbOK Then
 Text5.Text = InputBox("��������ȷ������.", "ע�⣡")
 GoTo ee
 ElseIf e = vbCancel Then
 Stop
 End If
 End If
 
ff:
outerwidth = Text6
date6 = Text6
If (IsNumeric(Text6.Text) = False) Then
 f = MsgBox("������Ϊ����", 1, "���󣡣�")
 If f = vbOK Then
 Text6.Text = InputBox("��������ȷ������.", "ע�⣡")
 GoTo ff
 ElseIf f = vbCancel Then
 Stop
 End If
 End If
 
gg:
outerhigh = Text7
date7 = Text7
If (IsNumeric(Text7.Text) = False) Then
 g = MsgBox("��߱���Ϊ����", 1, "���󣡣�")
 If g = vbOK Then
 Text7.Text = InputBox("��������ȷ������.", "ע�⣡")
 GoTo gg
 ElseIf g = vbCancel Then
 Stop
 End If
 End If
 
hh:
amount1 = Text8
date8 = Text8
If (IsNumeric(Text8.Text) = False) Then
 h = MsgBox("��������Ϊ����", 1, "���󣡣�")
 If h = vbOK Then
 Text8.Text = InputBox("��������ȷ������.", "ע�⣡")
 GoTo hh
 ElseIf h = vbCancel Then
 Stop
 End If
 End If
 
ii:
workcost = Text9
date9 = Text9
If (IsNumeric(Text9.Text) = False) Then
 i = MsgBox("�˹����ñ���Ϊ����", 1, "���󣡣�")
 If i = vbOK Then
 Text9.Text = InputBox("��������ȷ������.", "ע�⣡")
 GoTo ii
 ElseIf i = vbCancel Then
 Stop
 End If
 End If
 
jj:
liquidprice = Text10
date10 = Text10
If (IsNumeric(Text10.Text) = False) Then
 j = MsgBox("Һ����ñ���Ϊ����,���Ϊ�ɽ�����0", 1, "���󣡣�")
 If j = vbOK Then
 Text10.Text = InputBox("��������ȷ������.", "ע�⣡")
 GoTo jj
 ElseIf j = vbCancel Then
 Stop
 End If
 End If
 
kk:
innerpackage = Text11
date11 = Text11
If (IsNumeric(Text11.Text) = False) Then
 k = MsgBox("�ڰ����ñ���Ϊ����", 1, "���󣡣�")
 If k = vbOK Then
 Text11.Text = InputBox("��������ȷ������.", "ע�⣡")
 GoTo kk
 ElseIf k = vbCancel Then
 Stop
 End If
 End If
 
ll:
clothprice = Text12
date12 = Text12
If (IsNumeric(Text12.Text) = False) Then
 l = MsgBox("���ϵ��۱���Ϊ����", 1, "���󣡣�")
 If l = vbOK Then
 Text12.Text = InputBox("��������ȷ������.", "ע�⣡")
 GoTo ll
 ElseIf l = vbCancel Then
 Stop
 End If
 End If
 
mm:
yunfei = Text13
date13 = Text13
If (IsNumeric(Text13.Text) = False) Then
 m = MsgBox("�˷ѵ��۱���Ϊ����", 1, "���󣡣�")
 If m = vbOK Then
 Text13.Text = InputBox("��������ȷ������.", "ע�⣡")
 GoTo mm
 ElseIf m = vbCancel Then
 Stop
 End If
 End If
 
oo:
outerpaperprice = Text14
date14 = Text14
If (IsNumeric(Text14.Text) = False) Then
 o = MsgBox("ֽ�䵥�۱���Ϊ����", 1, "���󣡣�")
 If o = vbOK Then
 Text14.Text = InputBox("��������ȷ������.", "ע�⣡")
 GoTo oo
 ElseIf o = vbCancel Then
 Stop
 End If
 End If

If Val(Text10.Text) = 0 Then
num10 = 6.8
ElseIf Val(Text10.Text) <> 0 Then
num10 = 6.6
End If

If Option3.Value = False And Option4.Value = False Then
q = MsgBox("Ĭ������Ϊ���ھ���", vbYesNo, "�������")
If q = vbYes Then
Option3.Value = True
ElseIf q = vbNo Then
Option4.Value = True
End If
End If

If Option3.Value = True Then
Deviation1 = 0.015
Deviation2 = 0.02
ElseIf Option4.Value = True Then
Deviation1 = 0
Deviation2 = 0
End If

If Option1.Value = False And Option2.Value = False Then
p = MsgBox("�Ƿ�Ҫѡ����Ԫ���ۣ�", vbYesNo, "����δѡ��")
If p = vbYes Then
Option1.Value = True
ElseIf p = vbNo Then Option2.Value = True
End If
End If

If Option1.Value = True Then
huilv = num10
ElseIf Option2.Value = True Then huilv = 1
End If


 
pcsprice = Val(clothprice) / 1000000


num1 = Val(pcsprice) * Val(sheetlenght) * Val(sheetwidth) * Val(gsm) * Val(amount) * 1.02

num2 = ((Val(outerlenght) + Deviation1) + (Val(outerwidth) + Deviation1) + 0.08) * ((Val(outerwidth) + Deviation1) + (Val(outerhigh) + Deviation2) + 0.04) * Val(outerpaperprice)
num3 = num2 / Val(amount1)

If Val(workcost) > 100 Then
unitworkcost = Val(workcost) / 1000000
num4 = Val(sheetlenght) * Val(sheetwidth) * Val(gsm) * 1.02 * Val(amount) * unitworkcost

ElseIf workcost < 100 Then
num4 = Val(workcost)
End If

num5 = (Val(outerlenght) + Deviation1) * (Val(outerwidth) + Deviation1) * (Val(outerhigh) + Deviation2) * Val(yunfei)
num6 = num5 / Val(amount1)

num7 = (Val(Format(num1, "0.###")) + Val(Format(num3, "0.###")) + Val(Format(num4, "0.###")) + Val(Format(num6, "0.###")) + liquidprice + innerpackage) * 1.06 * 1.1
num8 = num7 / huilv

Text15 = Format(num8, "0.###")
date15 = Text15
Label34.Caption = "����" & n & "->" & num8

History = History & vbCrLf & Label34.Caption
Text16 = History

End Sub

Private Sub Picture1_Click()

End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Text15.Text = ""

End Sub

Private Sub Label34_()
Label34.Caption = "����1" & ":" & num7 & Chr(13) & "����2"

End Sub


Private Sub Command3_Click()
Text1.Text = date1
Text2.Text = date2
Text3.Text = date3
Text4.Text = date4
Text5.Text = date5
Text6.Text = date6
Text7.Text = date7
Text8.Text = date8
Text9.Text = date9
Text10.Text = date10
Text11.Text = date11
Text12.Text = date12
Text13.Text = date13
Text14.Text = date14
Text15.Text = date15
End Sub

Private Sub Command4_Click()
Dim n1 As String
Dim n2 As String
If Option1.Value = True Then
n1 = "USD"
ElseIf Option2.Value = True Then
n1 = "RMB"
End If
If Option3.Value = True Then
n2 = "�ھ�"
ElseIf Option4.Value = True Then
n2 = "�⾶"
End If
Name App.Path & "\�������ݼ�¼.csv" As App.Path & "\�������ݼ�¼.txt"

Open App.Path & "\�������ݼ�¼.txt" For Append As #1
Print #1, " "; ","; Text17; " "; ","; " "; " "; " "; Text1; "X"; Text2; " "; " "; " "; ","; " "; " "; " "; Text3; " "; " "; " "; ","; " "; " "; " "; Text4; " "; " "; " "; ","; " "; n2; Text5; "X"; Text6; "X"; Text7; " "; " "; " "; ","; " "; " "; " "; Text8; " "; " "; " "; ","; " "; " "; " "; Text9; " "; " "; " "; ","; " "; " "; " "; Text10; " "; " "; " "; ","; " "; " "; " "; Text11; " "; " "; " "; ","; " "; " "; " "; Text12; " "; " "; " "; ","; " "; " "; n1; Text15; " "; ","; " "; vbCrLf;
Close #1

Name App.Path & "\�������ݼ�¼.txt" As App.Path & "\�������ݼ�¼.csv"
End Sub
