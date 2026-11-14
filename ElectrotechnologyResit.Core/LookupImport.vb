Imports System.Globalization
Imports System.IO
Imports System.Text
Imports OfficeOpenXml

Namespace ElectrotechnologyResit.Core

    Public Class LookupImportService
        Private ReadOnly _students As IStudentDirectory
        Private ReadOnly _teachers As ITeacherDirectory
        Private ReadOnly _units As IUnitDirectory

        Public Sub New(students As IStudentDirectory, teachers As ITeacherDirectory, units As IUnitDirectory)
            _students = students
            _teachers = teachers
            _units = units
        End Sub

        Public Async Function ImportStudentsAsync(rows As IEnumerable(Of StudentRecord)) As Task(Of ImportTally)
            Dim tally As New ImportTally()
            If rows Is Nothing Then Return tally
            For Each row In rows
                If row Is Nothing OrElse String.IsNullOrWhiteSpace(row.StudentId) Then
                    tally.Skipped += 1
                    Continue For
                End If
                Dim kind = Await _students.UpsertAsync(row).ConfigureAwait(False)
                Count(tally, kind)
            Next
            Return tally
        End Function

        Public Async Function ImportTeachersAsync(rows As IEnumerable(Of TeacherRecord)) As Task(Of ImportTally)
            Dim tally As New ImportTally()
            If rows Is Nothing Then Return tally
            For Each row In rows
                If row Is Nothing OrElse String.IsNullOrWhiteSpace(row.FullName) Then
                    tally.Skipped += 1
                    Continue For
                End If
                Dim kind = Await _teachers.UpsertAsync(row).ConfigureAwait(False)
                Count(tally, kind)
            Next
            Return tally
        End Function

        Public Async Function ImportUnitsAsync(rows As IEnumerable(Of UnitRecord)) As Task(Of ImportTally)
            Dim tally As New ImportTally()
            If rows Is Nothing Then Return tally
            For Each row In rows
                If row Is Nothing OrElse String.IsNullOrWhiteSpace(row.UnitCode) Then
                    tally.Skipped += 1
                    Continue For
                End If
                Dim kind = Await _units.UpsertAsync(row).ConfigureAwait(False)
                Count(tally, kind)
            Next
            Return tally
        End Function

        Private Shared Sub Count(tally As ImportTally, kind As UpsertKind)
            If kind = UpsertKind.Inserted Then
                tally.Inserted += 1
            Else
                tally.Updated += 1
            End If
        End Sub
    End Class

    Public NotInheritable Class LookupFileReader
        Private Sub New()
        End Sub

        Public Shared Function ReadStudents(path As String) As IReadOnlyList(Of StudentRecord)
            Dim table = ReadTable(path)
            Dim id = RequireColumn(table.Headers, "Student ID", "StudentID", "Student Id")
            Dim given = FindHeader(table.Headers, "Given Name", "Student Given Name", "First Name", "Student Firstname", "Firstname")
            Dim family = FindHeader(table.Headers, "Family Name", "Student Family Name", "Surname", "Last Name", "Student Surname")
            Dim email = FindHeader(table.Headers, "Personal Email", "Student Personal Email", "Email", "Student Email")
            Dim block = FindHeader(table.Headers, "Block Group Code", "Block Group", "Blockgroup")
            Dim rows As New List(Of StudentRecord)
            For Each line In table.Rows
                rows.Add(New StudentRecord With {
                    .StudentId = Cell(line, id),
                    .GivenName = Cell(line, given),
                    .FamilyName = Cell(line, family),
                    .PersonalEmail = Cell(line, email),
                    .BlockGroupCode = Cell(line, block)
                })
            Next
            Return rows
        End Function

        Public Shared Function ReadTeachers(path As String) As IReadOnlyList(Of TeacherRecord)
            Dim table = ReadTable(path)
            Dim name = RequireColumn(table.Headers, "Full Name", "Teacher Full Name", "Teacher_Full_Name", "Name", "Teacher")
            Dim email = RequireColumn(table.Headers, "Email", "Teacher Email")
            Dim rows As New List(Of TeacherRecord)
            For Each line In table.Rows
                rows.Add(New TeacherRecord With {.FullName = Cell(line, name), .Email = Cell(line, email)})
            Next
            Return rows
        End Function

        Public Shared Function ReadUnits(path As String) As IReadOnlyList(Of UnitRecord)
            Dim table = ReadTable(path)
            Dim code = RequireColumn(table.Headers, "Unit Code", "Unit_Code", "Code", "Unit")
            Dim title = RequireColumn(table.Headers, "Unit Title", "Unit_Title", "Title", "Unit Name")
            Dim rows As New List(Of UnitRecord)
            For Each line In table.Rows
                rows.Add(New UnitRecord With {.UnitCode = Cell(line, code), .UnitTitle = Cell(line, title)})
            Next
            Return rows
        End Function

        Private Class TableData
            Public Property Headers As String()
            Public Property Rows As List(Of String())
        End Class

        Private Shared Function ReadTable(path As String) As TableData
            If String.IsNullOrWhiteSpace(path) OrElse Not File.Exists(path) Then
                Throw New FileNotFoundException("The import file was not found.")
            End If
            Dim extension = IO.Path.GetExtension(path)
            If extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) Then Return ReadExcel(path)
            If extension.Equals(".csv", StringComparison.OrdinalIgnoreCase) OrElse extension.Equals(".txt", StringComparison.OrdinalIgnoreCase) Then
                Return ReadCsv(File.ReadAllText(path, Encoding.UTF8))
            End If
            Throw New InvalidOperationException("Choose a CSV or Excel (.xlsx) file.")
        End Function

        Private Shared Function ReadCsv(text As String) As TableData
            If text IsNot Nothing AndAlso text.Length > 0 AndAlso text(0) = ChrW(&HFEFF) Then text = text.Substring(1)
            Dim parsed = ParseCsv(If(text, ""))
            If parsed.Count = 0 Then Throw New InvalidOperationException("The file has no header row.")
            Return New TableData With {.Headers = parsed(0), .Rows = parsed.Skip(1).ToList()}
        End Function

        Public Shared Function ParseCsv(text As String) As List(Of String())
            Dim rows As New List(Of String())
            Dim row As New List(Of String)
            Dim cell As New StringBuilder()
            Dim inQuotes = False
            Dim i = 0
            While i < text.Length
                Dim ch = text(i)
                If inQuotes Then
                    If ch = """"c Then
                        If i + 1 < text.Length AndAlso text(i + 1) = """"c Then
                            cell.Append(""""c)
                            i += 1
                        Else
                            inQuotes = False
                        End If
                    Else
                        cell.Append(ch)
                    End If
                ElseIf ch = """"c Then
                    inQuotes = True
                ElseIf ch = ","c Then
                    row.Add(cell.ToString())
                    cell.Clear()
                ElseIf ch = vbLf Then
                    row.Add(cell.ToString())
                    cell.Clear()
                    rows.Add(row.ToArray())
                    row = New List(Of String)
                ElseIf ch <> vbCr Then
                    cell.Append(ch)
                End If
                i += 1
            End While
            If cell.Length > 0 OrElse row.Count > 0 Then
                row.Add(cell.ToString())
                rows.Add(row.ToArray())
            End If
            Return rows
        End Function

        Private Shared Function ReadExcel(path As String) As TableData
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial
            Using package As New ExcelPackage(New FileInfo(path))
                Dim sheet = package.Workbook.Worksheets.FirstOrDefault()
                If sheet Is Nothing OrElse sheet.Dimension Is Nothing Then
                    Throw New InvalidOperationException("The workbook has no rows.")
                End If
                Dim lastColumn = sheet.Dimension.End.Column
                Dim lastRow = sheet.Dimension.End.Row
                Dim headers(lastColumn - 1) As String
                For columnIndex = 1 To lastColumn
                    headers(columnIndex - 1) = sheet.Cells(1, columnIndex).Text
                Next
                Dim rows As New List(Of String())
                For rowNumber = 2 To lastRow
                    Dim line(lastColumn - 1) As String
                    Dim any = False
                    For columnIndex = 1 To lastColumn
                        line(columnIndex - 1) = sheet.Cells(rowNumber, columnIndex).Text
                        If Not String.IsNullOrWhiteSpace(line(columnIndex - 1)) Then any = True
                    Next
                    If any Then rows.Add(line)
                Next
                Return New TableData With {.Headers = headers, .Rows = rows}
            End Using
        End Function

        Private Shared Function RequireColumn(headers As String(), ParamArray aliases As String()) As Integer
            Dim index = FindHeader(headers, aliases)
            If index < 0 Then Throw New InvalidOperationException("The file is missing a required column (" & String.Join(" or ", aliases) & ").")
            Return index
        End Function

        Private Shared Function FindHeader(headers As String(), ParamArray aliases As String()) As Integer
            If headers Is Nothing Then Return -1
            For index = 0 To headers.Length - 1
                Dim header = If(headers(index), "").Trim()
                For Each candidate In aliases
                    If header.Equals(candidate, StringComparison.OrdinalIgnoreCase) Then Return index
                Next
            Next
            Return -1
        End Function

        Private Shared Function Cell(row As String(), index As Integer) As String
            If index < 0 OrElse row Is Nothing OrElse index >= row.Length Then Return ""
            Return If(row(index), "").Trim()
        End Function
    End Class

End Namespace
