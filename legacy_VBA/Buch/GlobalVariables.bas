Attribute VB_Name = "GlobalVariables"
Option Explicit

' Email path and information variables
Public pathBriefFuerEMAIL As String
Public elternEMAIL As String

' Row and column indicators
Public letzteSpalte As Integer
Public elternLetzte As Integer
Public letzteRow As Integer
Public RowFurTabelle As Integer

' Parent/account data
Public nameEltern As String
Public Konto As String
Public Adresse As String

' Financial summary variables
Public sumEndJahr As Double
Public sumTarif As Double
Public sunZahlung As Double
Public sumTarifJahr As Double

' Time tracking variables
Public aktuelMonat As Integer
Public letzteKartei As Integer

' Position tracking variables
Public absoluteRow As Integer
Public absoluteColumn As Integer

' Variables for delayed search
Public tempListBoxRef As String
Public tempTextBoxValue As String
Public tempLastParentRow As Integer

' Performance metrics
Public searchTimeStart As Double
Public searchTimeEnd As Double
