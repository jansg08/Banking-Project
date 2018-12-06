Imports System
Imports System.IO
Module Module1
    Public Structure Transaction
        Public amount As Double
        Public date1 As Date
        Public reference As String
        Public type As Boolean
    End Structure
    Public Structure LoadFile
        Public transaction() As Transaction
        Public balance As Double
    End Structure
    Sub Main()
        Dim filePath As String = "c:\Users\Jack\Documents\stmt.txt"
        Dim transactions(0) As Transaction
        Dim balance As Double
        Dim choice As Char
        Dim startingBalance As Boolean = True
        Dim type As Boolean
        Dim data As LoadFile
        Dim length As Integer = transactions.Length
        Dim checker As Boolean
        Dim loader As StreamReader
        Console.SetWindowSize(146, 30)
        transactions(0).reference = ""
        If System.IO.File.Exists(filePath) Then
            loader = New StreamReader(filePath)
            If loader.ReadLine() IsNot Nothing Then
                loader.Close()
                startingBalance = False
                ReDim data.transaction(CountLines(loader, filePath) - 2)
                data = Load(filePath, balance, transactions, loader)
                balance = data.balance
                transactions = data.transaction
            End If
            loader.Close()
        Else
            Warning()
        End If
        loader.Close()
        Do
            menu()
            Console.Write(" ")
            choice = UCase(Console.ReadLine())
            Select Case choice
                Case "A"
                    Console.WriteLine()
                    If startingBalance Then
                        Console.WriteLine(" Please enter your starting balance")
                        Do
                            Try
                                Console.WriteLine()
                                Console.Write(" $")
                                balance = Console.ReadLine()
                            Catch ex As Exception
                                Console.WriteLine()
                                Console.WriteLine(" Please try again and enter a valid number")
                            End Try
                        Loop Until balance <> 0
                        startingBalance = False
                    Else
                        Console.WriteLine(" A starting balance has already been set so you will have to reset the account (using the 'R' option) if you want to change this")
                    End If
                Case "B"
                    Console.WriteLine()
                    If startingBalance Then
                        Console.WriteLine(" No data has been input - please select option 'A' to set a starting balance")
                    Else
                        OutputStatement(transactions, balance)
                    End If
                Case "C"
                    'Makes a deposit
                    type = False
                    length = transactions.Length
                    Console.WriteLine()
                    If startingBalance Then
                        Console.WriteLine(" Please set a starting balance before entering transactions")
                    Else
                        If transactions(0).reference <> "" Then
                            ReDim Preserve transactions(length)
                        End If
                        transactions = MakeTransaction(transactions, type, length, balance)
                        balance += transactions(transactions.Length - 1).amount
                    End If
                Case "D"
                    'Makes a withdrawal
                    type = True
                    length = transactions.Length
                    Console.WriteLine()
                    If startingBalance Then
                        Console.WriteLine(" Please set a starting balance before entering transactions")
                    Else
                        If transactions(0).reference <> "" Then
                            ReDim Preserve transactions(length)
                        End If
                        transactions = MakeTransaction(transactions, type, length, balance)
                        balance -= transactions(transactions.Length - 1).amount
                    End If
                Case "E"
                    'Runs statistics
                    Console.WriteLine()
                    Console.WriteLine(" This option is currently under construction. Sorry for any inconvenience caused.")
                Case "F"
                    'Creates a printable statement
                    Console.WriteLine()
                    Console.WriteLine(" This option is currently under construction. Sorry for any inconvenience caused.")
                Case "R"
                    'Resets all values back to 0
                    Console.WriteLine()
                    Console.WriteLine(" The following procedure will erase all data stored in this bank account and close the program. Are you sure you want to proceed? [Y/N]")
                    Console.WriteLine()
                    Console.Write(" ")
                    If UCase(Console.ReadLine()) = "Y" Then
                        choice = "x"
                        checker = True
                    End If
                    Console.WriteLine(" This option is currently under construction. Sorry for any inconvenience caused.")
                Case "Z"
                    Console.WriteLine()
                    Instructions()
            End Select
        Loop Until LCase(choice) = "x"
        If transactions(0).reference <> "" And checker = False Then
            Save(filePath, balance, transactions)
        ElseIf checker Then
            System.IO.File.WriteAllText(filePath, Nothing)
        End If
    End Sub
    Function MakeTransaction(ByVal transactions() As Transaction, ByVal type As Boolean, ByVal length As Integer, ByVal balance As Double) As Transaction()
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
    Function Load(ByRef filePath As String, ByRef balance As Double, ByRef transactions() As Transaction, ByVal loader As StreamReader) As LoadFile
        Dim load1 As LoadFile
        Dim temp As String
        Dim a As Integer
        Dim b As Integer = 1
        Dim x As Integer
        Dim line As String = ""
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
    Function CountLines(ByVal loader As StreamReader, ByVal filePath As String) As Integer
        Dim a As Integer = 1
        loader = New StreamReader(filePath)
        Do
            loader.ReadLine()
            a += 1
        Loop Until loader.readline() Is Nothing
        Return a
        loader.Close()
    End Function
    Sub menu()
        Dim filename As String = "c:\Users\Jack\Documents\file.txt"
        Console.WriteLine()
        Console.WriteLine(" Please choose from the following options:")
        Console.WriteLine()
        Console.WriteLine(" A. Set starting balance")
        Console.WriteLine(" B. Show balance and statement")
        Console.WriteLine(" C. Deposit")
        Console.WriteLine(" D. Withdraw")
        Console.WriteLine(" E. View statistics")
        Console.WriteLine(" F. Create a printable statement")
        Console.WriteLine(" R. Reset balance and all transactions")
        Console.WriteLine(" X. Save and Exit")
        Console.WriteLine(" Z. Intructions for creating text file")
        Console.WriteLine()
    End Sub
    Sub Warning()
        Console.WriteLine(" ------------------------------------------------------------------------------------------------------------------------------------------------")
        Console.WriteLine(" **It appears you haven't already created a text file to store your balance and other details, so please do this now (press z for instructions)**")
        Console.WriteLine(" ------------------------------------------------------------------------------------------------------------------------------------------------")
        Console.WriteLine()
    End Sub
    Sub Instructions()
        Console.WriteLine("  ______________________________________________________________")
        Console.WriteLine(" |                                                              |")
        Console.WriteLine(" | 1. Go to File Explorer                                       |")
        Console.WriteLine(" | 2. Go to your 'Documents' folder                             |")
        Console.WriteLine(" | 3. Right-click in that folder and select 'New'               |")
        Console.WriteLine(" | 4. Select 'Text Document' and name it 'stmt'                 |")
        Console.WriteLine(" | 5. Once completed, close this window and restart the program |")
        Console.WriteLine(" |______________________________________________________________|")
    End Sub
End Module