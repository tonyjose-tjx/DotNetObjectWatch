Public Class IEnumerableImplementaion

    Friend Shared Function WhereEvaluator(ByVal q As Object, ByVal rest As String) As IEnumerable


        Dim parameters() As String = rest.Substring(rest.IndexOf("(") + 1, rest.LastIndexOf(")") - rest.IndexOf("(") - 1).Split(",")

        If TypeOf q Is IEnumerable Then
            Dim ienuQ As IEnumerable = q

            Dim enumIter = ienuQ.GetEnumerator()
            enumIter.MoveNext()

            Dim itemType = enumIter.Current.GetType

            Dim parameterRepresentation As String = parameters.First()
            Dim parameterName As String = parameterRepresentation.Substring((parameterRepresentation.IndexOf("(") + 1), parameterRepresentation.IndexOf(")") - (parameterRepresentation.IndexOf("(") + 1))
            Dim expression As String = parameterRepresentation.Substring(parameterRepresentation.IndexOf(")") + 1)

            Dim linqParameter As System.Linq.Expressions.ParameterExpression = System.Linq.Expressions.Expression.Parameter(itemType, parameterName)
            Dim exp As System.Linq.Expressions.LambdaExpression = ExpressionBuilder.Linq.Dynamic.DynamicExpression.ParseLambda({linqParameter}, Nothing, expression)

            Dim collection As New List(Of Object)
            Dim lambdaEx As [Delegate] = exp.Compile
            For Each element As Object In ienuQ
                If lambdaEx.DynamicInvoke(element) Then
                    collection.Add(element)
                End If
            Next


            Return collection


        Else
            Throw New Exception("Where is not supported for the type.")
        End If

    End Function


End Class
