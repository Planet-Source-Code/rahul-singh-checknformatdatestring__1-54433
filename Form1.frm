VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Test Function"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   5505
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   4560
      TabIndex        =   11
      Text            =   "4"
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   4560
      TabIndex        =   10
      Text            =   "2"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox Text3 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   4560
      TabIndex        =   9
      Text            =   "2"
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox Text2 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   4560
      TabIndex        =   3
      Text            =   "/"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Text            =   "12.3.79"
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "2.3.79   -  02/03/2079 12-3-9  -  12/03/2009 2|3|9     -  02/03/1979  12\12\121 - 12/12/2121"
      Height          =   975
      Left            =   1200
      TabIndex        =   15
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "Some Outputs for above settings:"
      Height          =   255
      Left            =   480
      TabIndex        =   14
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label Label7 
      Caption         =   "Input Date separators Allowed                            /  -  \  .  |  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   480
      TabIndex        =   13
      Top             =   240
      Width           =   2700
   End
   Begin VB.Label Label6 
      Caption         =   "Output Date String"
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Year Digits"
      Height          =   255
      Left            =   3720
      TabIndex        =   8
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Date Digits"
      Height          =   255
      Left            =   3720
      TabIndex        =   7
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Month Digits"
      Height          =   255
      Left            =   3600
      TabIndex        =   6
      Top             =   900
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Input Date String"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Date Separator"
      Height          =   255
      Left            =   3360
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblOP 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1440
      TabIndex        =   2
      Top             =   1320
      Width           =   1005
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    lblOP.Caption = checkNformatDate(Text1.Text, Text3.Text, Text4.Text, Text5.Text, Text2.Text)
End Sub

'---------------------------------------------------------------------------------------
' Procedure    : checkNformatDate
' DateTime     : 6/17/2004 15:54
' Author       : RAHUL SINGH (2102037)
' Modified on  : 6/17/2004 15:54
' Purpose      :
'---------------------------------------------------------------------------------------

Public Function checkNformatDate(str_date As String, date_digit As Integer, month_digit As Integer, year_digit As Integer, Separator As String) As String
    If (date_digit <> 1 And date_digit <> 2) Or (month_digit <> 1 And month_digit <> 2) Or (year_digit <> 2 And year_digit <> 4) Then
        checkNformatDate = ""
        Exit Function
    End If
    
Dim x, tdate, tmonth, tyear
Dim start_pos, end_pos As Integer
    str_date = Trim(str_date)
    SearchString = str_date
    If InStr(str_date, "/") Then
        Searchchar = "/"
    ElseIf InStr(str_date, "-") Then
        Searchchar = "-"
    ElseIf InStr(str_date, "\") Then
        Searchchar = "\"
    ElseIf InStr(str_date, "|") Then
        Searchchar = "|"
    ElseIf InStr(str_date, ".") Then
        Searchchar = "."
    End If
    l = Len(str_date)
    x = InStr(1, SearchString, Searchchar, 1)
    If x > 1 Then
        end_pos = x - 1
        tdate = Trim(Mid(str_date, 1, end_pos))
        If Val(tdate) > 31 Then
            Exit Function
        End If
        start_pos = x + 1
        x = InStr(x + 1, SearchString, Searchchar, 1)
        If x > 0 Then
            end_pos = x - start_pos
            If end_pos = 0 Then
                Exit Function
            End If
            tmonth = Trim(Mid(str_date, start_pos, end_pos))
            If Val(tmonth) > 12 Then
                Exit Function
            End If
            start_pos = x + 1
            If Len(Trim(Mid(str_date, start_pos, l))) > 4 Then
                Exit Function
            End If
            tyear = Trim(Mid(str_date, start_pos, l))
            If InStr(tyear, Searchchar) Then
                tyear = Trim(Mid(str_date, l, l))
                If tyear = Searchchar Then
                    Exit Function
                End If
            End If
        Else
            Exit Function
        End If
    Else
        Exit Function
    End If
    If date_digit = 1 Then
        tdate = Val(tdate)
    ElseIf date_digit = 2 Then
        If Len(tdate) = 1 Then
            tdate = "0" & tdate
        End If
    End If
    If month_digit = 1 Then
        tmonth = Val(tmonth)
    ElseIf month_digit = 2 Then
        If Len(tmonth) = 1 Then
            tmonth = "0" & tmonth
        End If
    End If
    If year_digit = 2 Then
        If Len(tyear) = 1 Then
            tyear = "0" & tyear
        End If
    ElseIf year_digit = 4 Then
        If Len(tyear) = 1 Then
            tyear = "200" & tyear
        ElseIf Len(tyear) = 2 Then
            tyear = "20" & tyear
        ElseIf Len(tyear) = 3 Then
            tyear = "2" & tyear
        End If
    End If
    checkNformatDate = tdate & Separator & tmonth & Separator & tyear
End Function
