Attribute VB_Name = "Module1"
Declare Function BeginPath Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function EndPath Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function StrokePath Lib "gdi32" (ByVal hdc As Long) As Long


