Imports System
Imports System.IO

Module Program
    Sub Main()
        'Declare variables
        Dim doc As String
        Dim result As New List(Of Dictionary(Of String, String))
        Dim output As String = "Name,email" & vbCrLf
        'Manage inputs
        Console.WriteLine("Input your source .vcf file path: ")
        Dim inputPath As String = Console.ReadLine()
        Console.Clear()
        Console.WriteLine("Input your output .csv file  path: ")
        Dim outputPath As String = Console.ReadLine()
        Console.Clear()
        Console.WriteLine("Your input path is: " & inputPath)
        Console.WriteLine("Your output path is: " & outputPath)
        Console.WriteLine("Continue? Y/N")
        'Console.Clear()

        Dim key As Char = Console.ReadKey().KeyChar
        If key.ToString.ToLower <> "y" Then
            Return
        End If
        'Open Source file
        Using reader As StreamReader = New StreamReader(inputPath)
            doc = reader.ReadToEnd
        End Using
        'Split by vCard
        Dim splitDoc As String() = doc.Split("END:VCARD")

        'create dictionary of results
        For Each s As String In splitDoc
            If Not s = String.Empty Then
                s = s.Replace("BEGIN:VCARD", "").Replace(vbCrLf, " ")
                'declarations
                Dim exValues As New Dictionary(Of String, String)
                'get name value
                Dim nameStart As Integer = s.IndexOf("FN:") + 3
                Dim nameEnd As Integer = s.IndexOf(" N:")
                Dim nameVal As String = s.Substring(nameStart, nameEnd - nameStart).Trim.Replace(",", "")
                'Get email value
                Dim mailStart = s.IndexOf("EMAIL:") + 6
                Dim mailVal As String = s.Substring(mailStart).Trim
                'add to dict
                exValues.Add("Name", nameVal)
                exValues.Add("Email", mailVal)
                result.Add(exValues)
            End If
        Next s

        'Create output and save
        For Each item As Dictionary(Of String, String) In result
            If item.ContainsKey("Name") And item.ContainsKey("Email") Then
                output &= item.Item("Name") & "," & item.Item("Email") & vbCrLf
            End If

        Next item
        'Report outputs
        Using writer As StreamWriter = New StreamWriter(outputPath)
            writer.Write(output)
        End Using

        Console.WriteLine(vbCr & result.Count())
    End Sub

End Module
