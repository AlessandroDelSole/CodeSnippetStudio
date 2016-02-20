Imports System.Globalization
Imports Microsoft.CodeAnalysis

Public Class LocationConverter
    Implements IValueConverter

    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert
        Dim loc = CType(value, Location)
        Return CStr(loc.GetLineSpan.StartLinePosition.Line + 1)
    End Function

    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
        Throw New NotImplementedException()
    End Function
End Class
