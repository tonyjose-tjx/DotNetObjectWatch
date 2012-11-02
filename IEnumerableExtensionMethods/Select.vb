Partial Public Class IEnumerableImplementaion


    Friend Shared Function SelectImplementation(q As Object, rest As String) As IEnumerable


        Dim parameters() As String = rest.Substring(rest.IndexOf("(") + 1, rest.LastIndexOf(")") - rest.IndexOf("(") - 1).Split(",")

        If TypeOf q Is IEnumerable Then
            Dim ienuQ As IEnumerable = q

            Dim enumIter = ienuQ.GetEnumerator()
            enumIter.MoveNext()

            Dim itemType = enumIter.Current.GetType
            Dim parameterRepresentation As String = parameters.First()
            Dim parameterName As String = parameterRepresentation.Substring((parameterRepresentation.IndexOf("(") + 1), parameterRepresentation.IndexOf(")") - (parameterRepresentation.IndexOf("(") + 1))
            Dim expression As String = parameterRepresentation.Substring(parameterRepresentation.IndexOf(")") + 1)
            Dim linqParameter As System.Linq.Expressions.ParameterExpression = Expressions.Expression.Parameter(itemType, parameterName)
            Dim exp As System.Linq.Expressions.LambdaExpression = ExpressionBuilder.Linq.Dynamic.ParseLambda({linqParameter}, Nothing, expression)
            Dim collection As New List(Of Object)
            Dim lambdaEx As [Delegate] = exp.Compile

            For Each element As Object In ienuQ
                collection.Add(lambdaEx.DynamicInvoke(element))
            Next


            Dim result As String = Newtonsoft.Json.JsonConvert.SerializeObject(collection)

            Return result


        Else

            Throw New Exception("Select is not supported for the type.")

        End If


    End Function

End Class
