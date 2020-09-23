VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   7110
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1740
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   900
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   975
      Left            =   1740
      TabIndex        =   0
      Top             =   1740
      Width           =   2745
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Cls
    Print ValidateEmail(Text1)
End Sub

Public Function ValidateEmail(ByRef sEmail As String) As Boolean
    Dim bValid As Boolean ' Is email valid?
    Dim bFlag As Boolean ' Multipurpose boolean value
    Dim vntEmail As Variant ' Splitted Email
    Dim vntDomain As Variant ' Splitted Domain
    Dim vntValidDomainExt As Variant ' Valid domain extensions
    Dim lCount As Long ' Loop variable
    Dim lChars As Long ' Second loop variable
    ' Validates a email of form "a.bc_123@server.ext"
    
    ' Insert valid domain extensions in variable
    vntValidDomainExt = Array("no", "com", "edu", "gov", "int", "mil", "net", "org", "info", "biz", "pro", "name", "coop", "museum", "aero", _
                                                                "af", "al", "dz", "as", "ad", "ao", "ai", "aq", "ag", "ar", "am", "aw", "ac", "au", "at", "az", "bs", _
                                                                "bh", "bd", "bb", "By", "be", "bz", "bj", "bm", "bt", "bo", "ba", "bw", "bv", "br", "io", "bn", "bg", _
                                                                "bf", "bi", "kh", "cm", "ca", "cv", "ky", "cf", "td", "cs", "cl", "cn", "cx", "cc", "co", "km", "cg", "ck", _
                                                                "cr", "ci", "hr", "cu", "cy", "cz", "dk", "dj", "dm", "do", "tp", "ec", "eg", "sv", "gq", "er", "ee", "et", _
                                                                "fk", "fo", "fj", "fi", "fr", "gf", "pf", "tf", "ga", "gm", "ge", "de", "gh", "gi", "gr", "gl", "gd", "gp", _
                                                                "gu", "gt", "gg", "gn", "gw", "gy", "ht", "hm", "va", "hn", "hk", "hu", "is", "In", "id", "ir", "iq", "ie", _
                                                                "im", "il", "it", "jm", "jp", "je", "jo", "kz", "ke", "ki", "kp", "kr", "kw", "kg", "la", "lv", "lb", "ls", "lr", _
                                                                "ly", "li", "lt", "lu", "mo", "mk", "mg", "mw", "my", "mv", "ml", "mt", "mh", "mq", "mr", "mu", "yt", _
                                                                "mx", "fm", "md", "mc", "mn", "ms", "ma", "mz", "mm", "na", "nr", "np", "nl", "an", "nc", "nz", "ni", _
                                                                "ne", "ng", "nu", "nf", "mp", "no", "om", "pk", "pw", "ps", "pa", "pg", "py", "pe", "ph", "pn", "pl", _
                                                                "pt", "pr", "qa", "re", "ro", "ru", "rw", "kn", "lc", "vc", "ws", "sm", "st", "sa", "sn", "sc", "sl", "sg", _
                                                                "sk", "si", "sb", "so", "za", "gs", "es", "lk", "sh", "pm", "sd", "sr", "sj", "sz", "se", "ch", "sy", "tw", _
                                                                "tj", "tz", "th", "tg", "tk", "to", "tt", "tn", "tr", "tm", "tc", "tv", "ug", "ua", "ae", "gb", "uk", "us", _
                                                                "um", "uy", "su", "uz", "vu", "ve", "vn", "vg", "vi", "wf", "eh", "ye", "yu", "cd", "zm", "zr", "zw")
    
    ' Emails are normally composed of only lower-case characters
    ' so lower-case the email and since the parameter is ByRef, the
    ' email will be lower-cased back there.
    sEmail = LCase(sEmail)
    ' If sEmail contains "@"
    If InStr(sEmail, "@") Then
        ' Split email on "@"
        vntEmail = Split(sEmail, "@")
        ' Asuure only 1 "@"
        If UBound(vntEmail) = 1 Then
            ' Starts with alphanumeric character
            If Left(vntEmail(0), 1) Like "[a-z]" Or Left(vntEmail(0), 1) Like "[0-9]" Then
                bFlag = True
                ' Assure all other characters are alphanumeric or "." or "_"
                For lCount = 2 To Len(vntEmail(0))
                    If Not (Mid(vntEmail(0), lCount, 1) Like "[a-z]" Or Mid(vntEmail(0), lCount, 1) Like "[0-9]" Or Mid(vntEmail(0), lCount, 1) = "." Or Mid(vntEmail(0), lCount, 1) Like "_") Then bFlag = False
                Next lCount
                If bFlag Then
                    ' If domain part contains "."
                    If InStr(vntEmail(1), ".") Then
                        ' Split domains on "."
                        vntDomain = Split(vntEmail(1), ".")
                        bFlag = True
                        ' Assure all domains characters are alphanumeric and domain length is over 1 (>= 2)
                        For lCount = LBound(vntDomain) To UBound(vntDomain)
                            If Len(vntDomain(lCount)) < 2 Then bFlag = False
                            For lChars = 1 To Len(vntDomain(lCount))
                                If Not (Mid(vntDomain(lCount), lChars, 1) Like "[a-z]" Or Mid(vntDomain(lCount), lChars, 1) Like "[0-9]") Then bFlag = False
                                If Not bFlag Then Exit For
                            Next lChars
                            If Not bFlag Then Exit For
                        Next lCount
                        If bFlag Then
                            bFlag = False
                            ' Check if email domain extension is valid
                            For lCount = LBound(vntValidDomainExt) To UBound(vntValidDomainExt)
                                If vntDomain(UBound(vntDomain)) = vntValidDomainExt(lCount) Then
                                    bFlag = True
                                    Exit For
                                End If
                            Next lCount
                            If bFlag Then
                                ' If email has passed through all this, it's valid
                                bValid = True
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    ValidateEmail = bValid
End Function
