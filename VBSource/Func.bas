Attribute VB_Name = "Func"




Public Function TxtFilterToStr(Text As String) As String
   TxtFilterToStr = Text
   If IsNumeric(Text) = True Then TxtFilterToStr = ""
   If Mid(Text, 1, 1) = " " Then TxtFilterToStr = ""
   
End Function

Public Function TxtFilterToNumber(Text As String) As String
    TxtFilterToNumber = Text
    If IsNumeric(Text) = False Then TxtFilterToNumber = ""
    If Mid(Text, 1, 1) = " " Then TxtFilterToNumber = ""
End Function





