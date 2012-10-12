Imports System.Linq
Imports System.Collections.Generic
Imports System.Text
Imports System.Reflection



Public Enum MethodExecutionPermission
    AllMethods
    IEnumerableMethodsOnly
    NoMethods
End Enum




Public Class ObjectWatch

    Private _methodPermission As MethodExecutionPermission



    Public Sub New(ByVal methodExecutionPermission As MethodExecutionPermission)
        _methodPermission = methodExecutionPermission
    End Sub


    Public Function GetResults(ByVal query As String, ByVal item As Object) As String

        Try

            Dim inputString As String = query
            Dim instructions As List(Of String) = GetInstructionSet(inputString)



            For Each element As String In instructions

                If item IsNot Nothing Then
                    item = TypeReflectionTesting(item, element)
                End If

            Next

            If item IsNot Nothing Then
                Try
                    Dim result As String = Newtonsoft.Json.JsonConvert.SerializeObject(item)
                    Return result
                Catch ex As Exception
                    Return "Unable to serialize the results. Please check the item is serializable."
                End Try

            Else
                Return "Nothing"
            End If

        Catch ex As Exception

            Return "Unable to perform the query!"

        End Try

    End Function



    Private Function GetInstructionSet(ByVal inputText As String) As List(Of String)


        Dim calls As New List(Of String)

        Dim entireSetOfChars = inputText.ToCharArray()
        Dim ignoreTillEnd As Boolean = False
        Dim bracketCount As Integer = 0
        Dim indexPosition As Integer = 0
        Dim lastCount As Integer = 0
        For Each ch As Char In entireSetOfChars


            If ch = "(" AndAlso Not ignoreTillEnd Then

                If bracketCount = 0 Then
                    calls.Add(inputText.Substring(lastCount, indexPosition - lastCount))
                    lastCount = indexPosition
                End If

                bracketCount += 1


            ElseIf ch = ")" AndAlso Not ignoreTillEnd Then

                bracketCount -= 1


            ElseIf ch = """"c Then

                ignoreTillEnd = Not ignoreTillEnd

            ElseIf ch = "." Then


                If bracketCount = 0 Then

                    calls.Add(inputText.Substring(lastCount, indexPosition - lastCount))

                    lastCount = indexPosition

                End If



            End If




            indexPosition += 1

        Next

        calls.Add(inputText.Substring(lastCount, indexPosition - lastCount))


        Dim continousExecutionCalls As New List(Of String)


        For i As Integer = 0 To calls.Count - 1

            Dim item = calls(i)


            If item.StartsWith(".") Then

                continousExecutionCalls.Add(calls(i))

            ElseIf item.StartsWith("(") AndAlso item.EndsWith(")") Then

                If continousExecutionCalls.Last.EndsWith(")") Then

                    continousExecutionCalls.Add(calls(i))

                Else
                    Dim lastElement = continousExecutionCalls.Last
                    continousExecutionCalls(continousExecutionCalls.Count - 1) = lastElement & calls(i)


                End If

            End If


        Next


        Return continousExecutionCalls


    End Function



    Private Function TypeReflectionTesting(ByVal item As Object, ByVal command As String) As Object

        Dim q As Object = item


        If command.StartsWith(".") Then
            'This could be a property call or a method call. Or even a field access. 
            'As field access is nasty to be done in this way. I am not going to allow that.
            'Property access is fine as longer as as it is not an indexer property. 
            'If that is property it should start with ()
            Dim rest As String = command.Substring(1)

            If Not command.Contains("(") OrElse Not command.Contains(")") Then

                Dim prop As Object = q.GetType.GetProperty(rest).GetValue(q, Nothing)
                Return prop

            Else

                If Not _methodPermission = MethodExecutionPermission.NoMethods Then


                    'This could be a method.
                    Dim methodName As String = rest.Substring(0, rest.IndexOf("("))
                    Dim methodInfo As Reflection.MethodInfo = q.GetType.GetMethod(methodName)

                    If methodInfo IsNot Nothing Then

                        If _methodPermission = MethodExecutionPermission.AllMethods Then

                            Return MethodInvocation(q, rest, methodInfo)

                        Else

                            Return "Method Invocations are not permitted! The current setting is " & _methodPermission.ToString

                        End If


                    Else



                        'Extension methods
                        If methodName = "Where" Then

                            Return IEnumerableImplementaion.WhereEvaluator(q, rest)

                        ElseIf methodName = "Select" Then

                            Return IEnumerableImplementaion.SelectImplementation(q, rest)

                        ElseIf methodName = "Take" Then

                            Return IEnumerableImplementaion.TakeImplementation(q, rest)

                        Else

                            Return "Method not found!"

                        End If

                    End If

                Else

                    Return "Method Execution Is Not Permitted!"

                End If



            End If

        ElseIf command.StartsWith("(") AndAlso command.EndsWith(")") Then
            'Probably and indexer

            Return IndexerPropertyEvaluator(command, q)

        Else

            Return Nothing

        End If

    End Function


    Private Function StringToObject(ByVal value As String, ByVal type As Type) As Object

        If type = GetType(System.String) Then

            Return value

        ElseIf type = GetType(System.Int16) Then

            Return Int16.Parse(value)

        ElseIf type = GetType(System.Int32) Then

            Return Int32.Parse(value)

        ElseIf type = GetType(System.Int64) Then

            Return Int64.Parse(value)

        ElseIf type = GetType(System.Double) Then

            Return Double.Parse(value)

        ElseIf type = GetType(System.Decimal) Then

            Return Decimal.Parse(value)

        ElseIf type = GetType(System.Guid) Then

            Return Guid.Parse(value)

        ElseIf type = GetType(System.Single) Then

            Return Single.Parse(value)

        ElseIf type = GetType(System.DateTime) Then

            Return DateTime.Parse(value)

        ElseIf type = GetType(System.Byte) Then

            Return Byte.Parse(value)

        ElseIf type = GetType(System.Boolean) Then

            Return Boolean.Parse(value)

        Else

            Return False

        End If



    End Function

   





    Private Function MethodInvocation(ByVal q As Object, ByVal rest As String, ByVal methodInfo As Reflection.MethodInfo) As String

        Dim parameters() As String = rest.Substring(rest.IndexOf("(") + 1, rest.LastIndexOf(")") - rest.IndexOf("(") - 1).Split(",")
        Dim objectCollection As New List(Of Object)
        Dim params() As Reflection.ParameterInfo = methodInfo.GetParameters()
        Dim count As Integer = 0

        For Each param As Reflection.ParameterInfo In params

            Dim value As String = parameters(count).ToString
            Dim type As Type = param.ParameterType


            If type = GetType(System.String) Then

                value = value.Substring(value.IndexOf(""""), value.LastIndexOf(""""))

            End If

            Dim convertedObj As Object = StringToObject(value, type)
            objectCollection.Add(convertedObj)

            count += 1

        Next

        Dim returnvalue = methodInfo.Invoke(q, objectCollection.ToArray)
        Dim result As String = Newtonsoft.Json.JsonConvert.SerializeObject(returnvalue)

        Return result

    End Function


 
    Private Function IndexerPropertyEvaluator(ByVal command As String, ByVal obj As Object) As String

        Dim newString As String = command.Replace("(", "").Replace(")", "")
        Dim properties() As MemberInfo = obj.GetType.GetDefaultMembers
        Dim propInfo As PropertyInfo
        Dim indexParameters() As ParameterInfo
        Dim paramters() As String

        For Each defaultType As Reflection.MemberInfo In properties

            If defaultType.MemberType = MemberTypes.Property Then


                propInfo = obj.GetType.GetProperty(defaultType.Name)
                indexParameters = propInfo.GetIndexParameters
                paramters = newString.Split(","c)
                Dim paramCollection As New List(Of Object)
                Dim count As Integer = 0

                For Each param As ParameterInfo In indexParameters

                    Dim paramValue As String = paramters(count).ToString
                    Dim type As Type = param.ParameterType


                    If type = GetType(System.String) Then

                        paramValue = paramValue.Substring(paramValue.IndexOf(""""), paramValue.LastIndexOf(""""))

                    End If

                    Dim convertedObj As Object = StringToObject(paramValue, type)
                    paramCollection.Add(convertedObj)

                    count += 1

                Next

                Dim value As Object = propInfo.GetValue(obj, paramCollection.ToArray)
                Return value

            End If

        Next
        Return String.Empty
    End Function


End Class