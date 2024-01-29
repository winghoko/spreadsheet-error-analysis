Option Explicit

' MIT License
'
' Copyright (c) 2024 winghoko
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.

Public Function WLINEST(y_range As Range, x_range As Range, u_range As Range, _
    Optional constant As Boolean = True, _
    Optional absolute As Boolean = False, _
    Optional stats As Boolean = False _
) As Variant

    Dim length, counter As Integer
    Dim sum_w, sum_wx, sum_wxx, sum_wy, sum_wxy As Double
    Dim w, x, y As Double
    
    length = y_range.Cells.Count
    
    ' Various weighed sums, computed in a single loop
    sum_w = 0
    sum_wx = 0
    sum_wxx = 0
    sum_wy = 0
    sum_wxy = 0
    For counter = 1 To length
        w = u_range.Cells(counter, 1).Value
        w = 1 / w / w
        x = x_range.Cells(counter, 1).Value
        y = y_range.Cells(counter, 1).Value
        sum_w = sum_w + w
        sum_wx = sum_wx + w * x
        sum_wxx = sum_wxx + w * x * x
        sum_wy = sum_wy + w * y
        sum_wxy = sum_wxy + w * x * y
    Next counter
    
    ' Fitted slope and intercept
    Dim inv_det, slope, intercept As Double
    If constant Then
        inv_det = 1 / (sum_w * sum_wxx - sum_wx * sum_wx)
        slope = inv_det * (sum_w * sum_wxy - sum_wx * sum_wy)
        intercept = inv_det * (sum_wxx * sum_wy - sum_wx * sum_wxy)
    Else
        inv_det = 1 / sum_wxx
        slope = inv_det * sum_wxy
    End If
    
    
    Dim out() As Double
    
    If stats Then
        
        ' When both fitted values and statistics are asked for
        
        Dim chisq As Double
        Dim df As Integer
        
        ' Compute weighed chi-square assuming y uncertainties absolute
        chisq = 0
        For counter = 1 To length
            w = u_range.Cells(counter, 1).Value
            w = 1 / w / w
            x = x_range.Cells(counter, 1).Value
            y = y_range.Cells(counter, 1).Value
            chisq = chisq + w * (y - intercept - slope * x) ^ 2
        Next counter
        
        ' Compute degree of freedom
        If constant Then
            df = length - 2
        Else
            df = length - 1
        End If
        
        ' Compute uncertainties in fitted parameters
        Dim u_slope, u_intercept As Variant
        If constant Then
            u_intercept = (inv_det * sum_wxx) ^ 0.5
            u_slope = (inv_det * sum_w) ^ 0.5
        Else
            u_intercept = inv_det ^ 0.5
            u_slope = Application.NA()
        End If
         
        ' adjust uncertainties in fitted parameter
        ' when uncertainties in y are relative
        If Not (absolute) Then
            u_slope = u_slope * (chisq / df) ^ 0.5
            u_intercept = u_intercept * (chisq / df) ^ 0.5
        End If
        
        ' Construct the output array
        ReDim out(3, 2) As Double
        out(0, 0) = slope
        out(0, 1) = intercept
        out(1, 0) = u_slope
        out(1, 1) = u_intercept
        out(2, 0) = chisq
        out(2, 1) = df
                
    Else
        
        ' When only fitted parameters are asked for
        
        ' Construct the output array
        ReDim out(1, 2) As Double
        out(0, 0) = slope
        out(0, 1) = intercept

    End If
    
    ' Set the value of array to be returned
    WLINEST = out

End Function


Public Function WLINEST_HELP()
    
    Dim msg As String
    Dim result As Variant
    
    msg = "Weighted least-square fit for best linear trend." & vbNewLine & vbNewLine
    msg = msg & "If stats is FALSE, returns the row [fitted slope, fitted intercept]."
    msg = msg & vbNewLine & vbNewLine
    msg = msg & "If stats is true, returns the 3-by-2 table "
    msg = msg & "[[fitted slope, fitted intercept], [slope error, intercept error], "
    msg = msg & "[unscaled chi-square, degree of freedom]]."
    msg = msg & vbNewLine & vbNewLine & "INPUT PARAMETERS" & vbNewLine & vbNewLine
    msg = msg & "y_range - values of dependent (y) variables as a single-column range."
    msg = msg & vbNewLine & vbNewLine
    msg = msg & "x_range - values of independent (x) variable as a single-column range."
    msg = msg & vbNewLine & vbNewLine
    msg = msg & "u_range - values of uncertainty (standard deviation) in dependent (y) variable "
    msg = msg & "as a single-column range." & vbNewLine & vbNewLine
    msg = msg & "constant - [Optional, default=TRUE] if TRUE, the y-intercept (b) is calculated; "
    msg = msg & "otherwise it is set at 0." & vbNewLine & vbNewLine
    msg = msg & "absolute - [Optional, default=FALSE] if TRUE, uncertainty are treated as absolute; "
    msg = msg & "otherwise, they are treated as relative ratios. "
    msg = msg & "(NOTE: absolute affects only the error in the fitted parameters)." & vbNewLine & vbNewLine
    msg = msg & "stats - [Optional, default=FALSE] if TRUE, return additional regression statistics; "
    msg = msg & "otherwise return only fitted values."

    result = MsgBox(msg, Buttons:=vbOKOnly, Title:="Help on the WLINEST() Function")
    WLINEST_HELP = "=WLINEST_HELP()"

End Function
