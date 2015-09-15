Sub 本学期课程处理()

    Dim rows As Integer
    Dim rg As Range
    
    
    rows = ActiveSheet.[A65536].End(xlUp).Row '表行数
    
    a = Asc("A") '得到Ascii码
    o = Asc("O")
    
    For c = a To o
        For r = 1 To rows
            c_letter = Chr(c)
            cell = c_letter & r '&强制连接
            Set rg = Range(cell) 'range注意用set
            if c_letter = "G" Then
            	v = Replace(rg.value, "*", "")
            	rg.value = v
            End If
            if c_letter = "H" Then
            	v = Replace(rg.value, "－", "-")
            	rg.value = v
            End If
            If IsEmpty(rg) = True Then
                v = Switch(c_letter = "A", "东财", c_letter = "B", "0", c_letter = "C", "东财", c_letter = "D", "0", c_letter = "E", "0", c_letter = "F", "考试", c_letter = "G", "东财教师", c_letter = "H", "100-101周", c_letter = "I", "0", c_letter = "J", "0~0", c_letter = "K", "东财", c_letter = "L", "东财", c_letter = "M", "0", c_letter = "N", "0", c_letter = "O", "0")
            	rg.value = v
            End If


                
        Next r
    Next c
    
End Sub


'正则处理: \((\d*#|研\S*)\)