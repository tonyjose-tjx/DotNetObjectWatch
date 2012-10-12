Partial Public Class IEnumerableImplementaion

    Friend Shared Function TakeImplementation(ByVal q As Object, ByVal rest As String) As IEnumerable

        Dim parameters() As String = rest.Substring(rest.IndexOf("(") + 1, rest.LastIndexOf(")") - rest.IndexOf("(") - 1).Split(",")

        If parameters.Count <> 1 Then
            Throw New ArgumentException("Invalid number of parameters passed to Take() method")
        End If

        Dim count As Integer
        If Not Integer.TryParse(parameters(0), count) Then
            Throw New Exception("The parameter passed to take method is not an Integer.")
        End If

        If Not TypeOf q Is IEnumerable Then
            Throw New Exception("Take can be called only to IEnumerable types")
        End If

        Dim iterItem As IEnumerable = q
        Dim i As Integer = 0
        Dim items As New List(Of Object)

        For Each item As Object In iterItem

            If i < count Then
                items.Add(item)
            Else
                Exit For
            End If

            i += 1

        Next

        Return items

    End Function



End Class
