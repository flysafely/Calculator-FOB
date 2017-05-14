VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "申皇报价计算器  Version 1.1.1"
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
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text17 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   3480
      TabIndex        =   74
      Text            =   "请输入商品名称"
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "数据保存"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
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
      Caption         =   "返回数据"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
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
      Caption         =   "重 置"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
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
      Caption         =   "计     算"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
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
      Caption         =   "外包装规格"
      Height          =   735
      Left            =   120
      TabIndex        =   66
      Top             =   1080
      Width           =   1335
      Begin VB.OptionButton Option4 
         Caption         =   "外径"
         Height          =   375
         Left            =   720
         TabIndex        =   71
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton Option3 
         Caption         =   "内径"
         Height          =   375
         Left            =   120
         TabIndex        =   70
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "货币选择"
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
         Name            =   "宋体"
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
      Caption         =   "各项数据保存至外部TXT中"
      Height          =   255
      Index           =   16
      Left            =   5640
      TabIndex        =   79
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label Label36 
      Caption         =   "数据保存是指将计算器中"
      Height          =   255
      Index           =   15
      Left            =   5640
      TabIndex        =   78
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Label36 
      Caption         =   "少元,也可填每袋多少元"
      Height          =   255
      Index           =   14
      Left            =   5640
      TabIndex        =   77
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label Label36 
      Caption         =   "人工费用项目填入每吨多"
      Height          =   255
      Index           =   13
      Left            =   5640
      TabIndex        =   76
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label10 
      Caption         =   "商品名称"
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
      Caption         =   "的产品单价！！"
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
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
      Caption         =   "一盒,一罐等单位"
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
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
      Caption         =   "结果为一包,一卷"
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
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
      Caption         =   "包装袋等一切费用"
      Height          =   255
      Index           =   10
      Left            =   5640
      TabIndex        =   62
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label Label36 
      Caption         =   "其他盖,标签,贴纸,透明"
      Height          =   255
      Index           =   9
      Left            =   5640
      TabIndex        =   61
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label36 
      Caption         =   "(外纸箱包装费用)以外的"
      Height          =   255
      Index           =   6
      Left            =   5640
      TabIndex        =   60
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label36 
      Caption         =   "内包费用是指除去外包装"
      Height          =   255
      Index           =   3
      Left            =   5640
      TabIndex        =   59
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label36 
      Caption         =   "项目处填写“0”"
      Height          =   255
      Index           =   8
      Left            =   5640
      TabIndex        =   58
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label36 
      Caption         =   "若无液体,请在液体费用"
      Height          =   255
      Index           =   7
      Left            =   5640
      TabIndex        =   57
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label36 
      Caption         =   "一箱中有多少单位产品"
      Height          =   255
      Index           =   5
      Left            =   5640
      TabIndex        =   56
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label36 
      Caption         =   "外包规格中的数量是指"
      Height          =   255
      Index           =   4
      Left            =   5640
      TabIndex        =   55
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label36 
      Caption         =   "每罐中有多少张/抽/片"
      Height          =   255
      Index           =   2
      Left            =   5640
      TabIndex        =   54
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label36 
      Caption         =   "每包/每盒/每卷/每袋/"
      Height          =   255
      Index           =   1
      Left            =   5640
      TabIndex        =   53
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label36 
      Caption         =   "内包规格中的数量是指"
      Height          =   255
      Index           =   0
      Left            =   5640
      TabIndex        =   52
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label35 
      Caption         =   "历史记录"
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
      Caption         =   "结果"
      BeginProperty Font 
         Name            =   "仿宋_GB2312"
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
      Caption         =   "元/m2"
      Height          =   255
      Left            =   4680
      TabIndex        =   33
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label31 
      Caption         =   "元/m3"
      Height          =   255
      Left            =   2280
      TabIndex        =   32
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label Label30 
      Caption         =   "元/吨"
      Height          =   255
      Left            =   4680
      TabIndex        =   31
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label29 
      Caption         =   "元/件"
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
      Caption         =   "元/件"
      Height          =   255
      Left            =   2280
      TabIndex        =   28
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label25 
      Caption         =   "元/吨or单位"
      Height          =   255
      Left            =   1800
      TabIndex        =   27
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label24 
      Caption         =   "纸箱单价:"
      Height          =   255
      Left            =   2880
      TabIndex        =   26
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label23 
      Caption         =   "运费单价:"
      Height          =   255
      Left            =   480
      TabIndex        =   25
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label22 
      Caption         =   "布料单价:"
      Height          =   255
      Left            =   2880
      TabIndex        =   24
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label21 
      Caption         =   "内包费用:"
      Height          =   255
      Left            =   480
      TabIndex        =   23
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label20 
      Caption         =   "液体费用:"
      Height          =   255
      Left            =   2880
      TabIndex        =   22
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label19 
      Caption         =   "人工费用"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label18 
      Caption         =   "单位"
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
      Caption         =   "数量:"
      Height          =   255
      Left            =   3240
      TabIndex        =   17
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label14 
      Caption         =   "箱宽:"
      Height          =   255
      Left            =   3240
      TabIndex        =   16
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label13 
      Caption         =   "箱高:"
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
      Caption         =   "箱长:"
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
      Caption         =   "克重:"
      Height          =   255
      Left            =   3240
      TabIndex        =   10
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "片"
      Height          =   255
      Left            =   2520
      TabIndex        =   9
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   "数量:"
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
      Caption         =   "纸宽:"
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
      Caption         =   "纸长:"
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "内包规格:"
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
 a = MsgBox("纸长必须为数字", 1, "错误！！")
 If a = vbOK Then
 Text1.Text = InputBox("请输入正确的数字.", "注意！")
 GoTo aa
 ElseIf a = vbCancel Then
 Stop
 End If
 End If
 
bb:
sheetwidth = Text2
date2 = Text2
If (IsNumeric(Text2.Text) = False) Then
 b = MsgBox("纸宽必须为数字", 1, "错误！！")
 If b = vbOK Then
 Text2.Text = InputBox("请输入正确的数字.", "注意！")
 GoTo bb
 ElseIf b = vbCancel Then
 Stop
 End If
 End If
 
cc:
amount = Text3
date3 = Text3
If (IsNumeric(Text3.Text) = False) Then
 c = MsgBox("数量必须为数字", 1, "错误！！")
 If c = vbOK Then
 Text3.Text = InputBox("请输入正确的数字.", "注意！")
 GoTo cc
 ElseIf c = vbCancel Then
 Stop
 End If
 End If
 
dd:
gsm = Text4
date4 = Text4
If (IsNumeric(Text4.Text) = False) Then
 d = MsgBox("克重必须为数字", 1, "错误！！")
 If d = vbOK Then
 Text4.Text = InputBox("请输入正确的数字.", "注意！")
 GoTo dd
 ElseIf d = vbCancel Then
 Stop
 End If
 End If
 
ee:
outerlenght = Text5
date5 = Text5
If (IsNumeric(Text5.Text) = False) Then
 e = MsgBox("箱长必须为数字", 1, "错误！！")
 If e = vbOK Then
 Text5.Text = InputBox("请输入正确的数字.", "注意！")
 GoTo ee
 ElseIf e = vbCancel Then
 Stop
 End If
 End If
 
ff:
outerwidth = Text6
date6 = Text6
If (IsNumeric(Text6.Text) = False) Then
 f = MsgBox("箱宽必须为数字", 1, "错误！！")
 If f = vbOK Then
 Text6.Text = InputBox("请输入正确的数字.", "注意！")
 GoTo ff
 ElseIf f = vbCancel Then
 Stop
 End If
 End If
 
gg:
outerhigh = Text7
date7 = Text7
If (IsNumeric(Text7.Text) = False) Then
 g = MsgBox("箱高必须为数字", 1, "错误！！")
 If g = vbOK Then
 Text7.Text = InputBox("请输入正确的数字.", "注意！")
 GoTo gg
 ElseIf g = vbCancel Then
 Stop
 End If
 End If
 
hh:
amount1 = Text8
date8 = Text8
If (IsNumeric(Text8.Text) = False) Then
 h = MsgBox("数量必须为数字", 1, "错误！！")
 If h = vbOK Then
 Text8.Text = InputBox("请输入正确的数字.", "注意！")
 GoTo hh
 ElseIf h = vbCancel Then
 Stop
 End If
 End If
 
ii:
workcost = Text9
date9 = Text9
If (IsNumeric(Text9.Text) = False) Then
 i = MsgBox("人工费用必须为数字", 1, "错误！！")
 If i = vbOK Then
 Text9.Text = InputBox("请输入正确的数字.", "注意！")
 GoTo ii
 ElseIf i = vbCancel Then
 Stop
 End If
 End If
 
jj:
liquidprice = Text10
date10 = Text10
If (IsNumeric(Text10.Text) = False) Then
 j = MsgBox("液体费用必须为数字,如果为干巾请填0", 1, "错误！！")
 If j = vbOK Then
 Text10.Text = InputBox("请输入正确的数字.", "注意！")
 GoTo jj
 ElseIf j = vbCancel Then
 Stop
 End If
 End If
 
kk:
innerpackage = Text11
date11 = Text11
If (IsNumeric(Text11.Text) = False) Then
 k = MsgBox("内包费用必须为数字", 1, "错误！！")
 If k = vbOK Then
 Text11.Text = InputBox("请输入正确的数字.", "注意！")
 GoTo kk
 ElseIf k = vbCancel Then
 Stop
 End If
 End If
 
ll:
clothprice = Text12
date12 = Text12
If (IsNumeric(Text12.Text) = False) Then
 l = MsgBox("布料单价必须为数字", 1, "错误！！")
 If l = vbOK Then
 Text12.Text = InputBox("请输入正确的数字.", "注意！")
 GoTo ll
 ElseIf l = vbCancel Then
 Stop
 End If
 End If
 
mm:
yunfei = Text13
date13 = Text13
If (IsNumeric(Text13.Text) = False) Then
 m = MsgBox("运费单价必须为数字", 1, "错误！！")
 If m = vbOK Then
 Text13.Text = InputBox("请输入正确的数字.", "注意！")
 GoTo mm
 ElseIf m = vbCancel Then
 Stop
 End If
 End If
 
oo:
outerpaperprice = Text14
date14 = Text14
If (IsNumeric(Text14.Text) = False) Then
 o = MsgBox("纸箱单价必须为数字", 1, "错误！！")
 If o = vbOK Then
 Text14.Text = InputBox("请输入正确的数字.", "注意！")
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
q = MsgBox("默认数据为“内径”", vbYesNo, "数据类别！")
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
p = MsgBox("是否要选择美元报价？", vbYesNo, "货币未选择！")
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
Label34.Caption = "数据" & n & "->" & num8

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
Label34.Caption = "数据1" & ":" & num7 & Chr(13) & "数据2"

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
n2 = "内径"
ElseIf Option4.Value = True Then
n2 = "外径"
End If
Name App.Path & "\计算数据记录.csv" As App.Path & "\计算数据记录.txt"

Open App.Path & "\计算数据记录.txt" For Append As #1
Print #1, " "; ","; Text17; " "; ","; " "; " "; " "; Text1; "X"; Text2; " "; " "; " "; ","; " "; " "; " "; Text3; " "; " "; " "; ","; " "; " "; " "; Text4; " "; " "; " "; ","; " "; n2; Text5; "X"; Text6; "X"; Text7; " "; " "; " "; ","; " "; " "; " "; Text8; " "; " "; " "; ","; " "; " "; " "; Text9; " "; " "; " "; ","; " "; " "; " "; Text10; " "; " "; " "; ","; " "; " "; " "; Text11; " "; " "; " "; ","; " "; " "; " "; Text12; " "; " "; " "; ","; " "; " "; n1; Text15; " "; ","; " "; vbCrLf;
Close #1

Name App.Path & "\计算数据记录.txt" As App.Path & "\计算数据记录.csv"
End Sub
