Imports System
Imports System.IO
Public Class BankAccount
    Private Structure Transaction
        Public amount As Double
        Public date1 As Date
        Public reference As String
        Public type As Boolean
    End Structure
    Private Structure LoadFile
        Public transaction() As Transaction
        Public balance As Double
    End Structure
    Private balance As Double
    Private transactions(0) As Transaction
    Function MakeTransaction(ByVal type As Boolean, ByVal length As Integer, ByVal balance As Double)
        Dim check As Boolean
        Console.WriteLine(" Please type the amount of money transferred")
        Console.WriteLine()
        If transactions.Length = 1 And transactions(0).reference = "" Then
            If type = False Then
                transactions(0).type = False
            Else
                transactions(0).type = True
            End If
            Do
                Try
                    Console.Write(" $")
                    transactions(0).amount = Console.ReadLine()
                    Console.WriteLine()
                Catch ex As Exception
                    Console.WriteLine(" Please try again and enter a valid number")
                    Console.WriteLine()
                End Try
            Loop Until transactions(0).amount <> 0
            Console.WriteLine(" Now enter the date of this transaction in the form MM/DD/(YY)YY")
            Console.WriteLine()
            Do
                Console.Write(" ")
                Try
                    transactions(0).date1 = Console.ReadLine()
                    Console.WriteLine()
                    check = True
                Catch ex As Exception
                    Console.WriteLine()
                    Console.WriteLine(" Please try again and enter a valid date")
                    Console.WriteLine()
                    check = False
                End Try
            Loop Until check = True
            Console.WriteLine(" Now please enter a refernce to identify this transaction (this could be a description about what is was for)")
            Console.WriteLine()
            Console.Write(" ")
            transactions(0).reference = Console.ReadLine()
            If transactions(0).reference = "" Or transactions(0).reference = " " Then
                transactions(0).reference = "-"
            End If
        Else
            ReDim Preserve transactions(length)
            If type = False Then
                transactions(length).type = False
            Else
                transactions(length).type = True
            End If
            Do
                Console.Write(" $")
                transactions(length).amount = Console.ReadLine()
                Console.WriteLine()
                If transactions(length).amount = 0 Then
                    Console.WriteLine(" Please try again")
                    Console.WriteLine()
                End If
            Loop Until transactions(length).amount <> 0
            Console.WriteLine(" Now enter the date of this transaction in the form MM/DD/(YY)YY")
            Console.WriteLine()
            Do
                Console.Write(" ")
                Try
                    transactions(length).date1 = Console.ReadLine()
                    Console.WriteLine()
                    check = True
                Catch ex As Exception
                    Console.WriteLine()
                    Console.WriteLine(" Please try again and enter a valid date")
                    Console.WriteLine()
                    check = False
                End Try
            Loop Until check = True
            Console.WriteLine(" Now please enter a refernce to identify this transaction (this could be a description about what is was for)")
            Console.WriteLine()
            Console.Write(" ")
            transactions(length).reference = Console.ReadLine()
            If transactions(length).reference = "" Or transactions(length).reference = " " Then
                transactions(length).reference = "-"
            End If
        End If
        Return transactions
    End Function
    Sub OutputStatement(ByRef transactions() As Transaction, ByVal balance As Double)
        Dim x As Integer
        Dim y As Integer
        Dim z As Integer
        If transactions(0).reference <> "" Then
            Console.Write(" Date       |")
            Console.Write(" Withdrawal or deposit? |")
            Console.Write(" Amount    |")
            Console.Write(" Reference")
            Console.WriteLine()

            For x = 0 To transactions.Length - 1
                Console.Write(" ")
                Console.Write(FormatDateTime(transactions(x).date1, DateFormat.ShortDate))
                For z = 1 To 11 - Len(CStr(transactions(x).date1))
                    Console.Write(" ")
                Next
                Console.Write("| ")
                If transactions(x).type = True Then
                    Console.Write("Withdrawal")
                    Console.Write("             | $")
                Else
                    Console.Write("Deposit")
                    Console.Write("                | $")
                End If
                Console.Write(transactions(x).amount)
                For y = 0 To 8 - Len(transactions(x).amount.ToString)
                    Console.Write(" ")
                Next
                Console.Write("| ")
                Console.Write(transactions(x).reference)
                Console.WriteLine()
            Next
            Console.WriteLine()
        Else
            Console.WriteLine(" There are no transactions to display - To add some please use the 'Deposit' or 'Withdraw' options in the menu")
        End If
        Console.WriteLine()
        Console.Write(" Your balance is: $")
        Console.WriteLine(balance)
    End Sub
    Sub Save(ByRef filePath As String, ByRef balance As Double, ByRef transactions() As Transaction)
        Dim file As System.IO.StreamWriter
        Dim x As Integer
        filePath = filePath & "1"
        file = My.Computer.FileSystem.OpenTextFileWriter(filePath, False)
        file.WriteLine(balance)
        For x = 0 To transactions.Length - 1
            file.Write(transactions(x).amount)
            file.Write(" ")
            file.Write(FormatDateTime(transactions(x).date1, DateFormat.ShortDate))
            file.Write(" ")
            If transactions(x).type = True Then
                file.Write("1")
            ElseIf transactions(x).type = False Then
                file.Write("0")
            End If
            file.Write(" ")
            file.Write(transactions(x).reference)
            file.WriteLine()
        Next
        file.Close()
    End Sub
    Function Load(ByRef filePath As String, ByRef balance As Double)
        Dim load1 As LoadFile
        Dim temp As String
        Dim a As Integer
        Dim b As Integer = 1
        Dim x As Integer
        Dim line As String = ""
        Dim loader As StreamReader
        loader = New StreamReader(filePath)
        load1.balance = loader.ReadLine()
        Do
            line = loader.ReadLine()
            a += 1
        Loop Until line Is Nothing
        loader.Close()
        loader = New StreamReader(filePath)
        ReDim load1.transaction(a - 2)
        Dim temptransactions(a - 2) As String
        loader.ReadLine()
        For x = 0 To a - 2
            temptransactions(x) = loader.ReadLine()
        Next
        loader.Close()
        For x = 0 To a - 2
            temp = ""
            Do Until Mid(temptransactions(x), b, 1) = " "
                temp = temp & (Mid(temptransactions(x), b, 1))
                b += 1
            Loop
            load1.transaction(x).amount = CDbl(temp)
            b += 1
            temp = ""
            Do Until Mid(temptransactions(x), b, 1) = " "
                temp = temp & (Mid(temptransactions(x), b, 1))
                b += 1
            Loop
            load1.transaction(x).date1 = CDate(temp)
            b += 1
            If Mid(temptransactions(x), b, 1) = "1" Then
                load1.transaction(x).type = True
            ElseIf Mid(temptransactions(x), b, 1) = "0" Then
                load1.transaction(x).type = False
            End If
            b += 1
            load1.transaction(x).reference = Mid(temptransactions(x), b + 1, Len(temptransactions(x)) - b)
            b = 1
        Next
        Return load1
    End Function
End Class
