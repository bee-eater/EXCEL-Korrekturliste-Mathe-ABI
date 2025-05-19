Attribute VB_Name = "M1_Global"
Option Explicit

'-----------------------------------------------------
' KONSTANTEN
'-----------------------------------------------------
Public Const DevMode = 0
Public Const WbPw = ""
Public Const Version = "v2.0.2"

Public Const WbNameConfig = "Config"
Public Const WbNameGradeKey = "Notenspiegel"
Public Const WbNameGradeSheet = "Noten"
Public Const WbNamePrintSheet = "Print"
Public Const WbNameTestDaten = "TestData"

Public Const CfgNameChart = "GradeChart"

Public Const CfgVLookUpPoints = "!$B$3:$C$302,2,0" ' SVERWEIS auf Punkte
Public Const CfgVLookUpGrades = "!$B$3:$D$302,3,0" ' SVERWEIS auf Note
Public Const CfgVLookUpUpDown = "!$B$3:$E$302,4,0" ' SVERWEIS auf Grenzfälle

Public Const CfgMaxSheets = 5        ' Anzahl der Teilbereiche - 1
Public Const CfgFirstSect = "$F$4"   ' Zelle mit dem ersten Teilbereich (2 Spalten jeweils)
Public Const CfgExerCount = "$F$21"  ' Zelle in der die Anzahl der angelegten Teilaufgaben steht
Public Const CfgFirstPupi = "$B$5"   ' Zelle an der die Schüler beginnen
Public Const CfgNumOfPupi = "$C$45"  ' Zelle in der die Anzahl der Schüler steht
Public Const CfgAbiDate = "$G$27"    ' Zelle in der das Datum steht
Public Const CfgAbiClass = "$G$29"   ' Zelle in der der Kurs steht
Public Const CfgAbiTeacher = "$G$28" ' Zelle in der der Kursleiter steht
Public Const CfgAbiTitle = "$F$2"    ' Zelle mit dem Titel der Arbeit
Public Const CfgUpdateInfo = "$J$29" ' Zelle mit der Update-Info

Public Const CfgPrintNameCol = 5     ' Spalte für Namen auf Druckseite

Public Const CfgColStart = 2
Public Const CfgRowStart = 2
Public Const CfgColOffsetFirstEx = 2
Public Const CfgRowOffsetFirstEx = 3
Public Const CfgRowOffsetFirstPupil = 5

'-----------------------------------------------------
' GLOBALE VARIABLEN
'-----------------------------------------------------
'Abbruch
Public cmdAbortAll As Boolean
'Farben
Public gClrHeader As Long
Public gClrTheme1 As Long
Public gClrTheme2 As Long

Public gClrTabGrades As Long
Public gClrTabSections As Long
Public gClrTabPrint As Long

'Grade limits
Public gClrMinus2 As Long
Public gClrMinus1 As Long
Public gClrPlus1 As Long
Public gClrPlus2 As Long

'Vars
Public gNumOfPupils As Integer
Public gSheetCnt As Integer

'-----------------------------------------------------
' GLOBALE VARIABLEN INITIALISIEREN
'-----------------------------------------------------
Public Function Init()

    gClrHeader = RGB(196, 215, 155)
    gClrTheme1 = RGB(217, 217, 217)
    gClrTheme2 = RGB(217, 217, 217)
    
    gClrTabGrades = RGB(0, 176, 240)
    gClrTabSections = RGB(146, 208, 80)
    gClrTabPrint = RGB(255, 255, 0)
    
    gClrMinus2 = RGB(146, 208, 80)
    gClrMinus1 = RGB(205, 255, 189)
    gClrPlus1 = RGB(255, 151, 151)
    gClrPlus2 = RGB(255, 0, 0)
    
    gNumOfPupils = Worksheets(WbNameConfig).Range(CfgNumOfPupi).Value

    Dim i As Integer
    gSheetCnt = 0
    For i = 0 To CfgMaxSheets
        If WSExists(Worksheets(WbNameConfig).Range(CfgFirstSect).Offset(0, i * 2).Value) Then
            gSheetCnt = gSheetCnt + 1
        End If
    Next i
    
End Function
