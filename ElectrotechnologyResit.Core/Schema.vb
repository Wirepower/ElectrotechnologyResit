Namespace ElectrotechnologyResit.Core

    Public NotInheritable Class SharePointColumns
        Public Const StudentId As String = "StudentID"
        Public Const StudentFirstname As String = "StudentFirstname"
        Public Const Surname As String = "Surname"
        Public Const Unit As String = "Unit"
        Public Const UnitName As String = "UnitName"
        Public Const Blockgroup As String = "Blockgroup"
        Public Const AllocatedTeacher As String = "AllocatedTeacher"
        Public Const AllocatedTeacherEmail As String = "AllocatedTeacherEmail"
        Public Const AttemptNo As String = "AttemptNo"
        Public Const ResitDate As String = "ResitDate"
        Public Const EnergyspaceClassCreated As String = "EnergyspaceClassCreated"
        Public Const EnergyspaceAssessmentBooked As String = "EnergyspaceAssessmentBooked"
        Public Const Status As String = "Status"
        Public Const StudentGivenName As String = "StudentGivenName"
        Public Const StudentFamilyName As String = "StudentFamilyName"
        Public Const StudentPersonalEmail As String = "StudentPersonalEmail"
        Public Const BlockGroupCode As String = "BlockGroupCode"
        Public Const TeacherFullName As String = "TeacherFullName"
        Public Const Email As String = "Email"
        Public Const UnitCode As String = "UnitCode"
        Public Const UnitTitle As String = "UnitTitle"
        Public Const QualificationCode As String = "QualificationCode"
        Public Const QualificationTitle As String = "QualificationTitle"
        Public Const Version As String = "Version"
        Public Const ClassGroupSuffix As String = "ClassGroupSuffix"
        Public Const AbbreviationRules As String = "AbbreviationRules"

        Private Sub New()
        End Sub
    End Class

    Public NotInheritable Class SharePointListSchema
        Private Sub New()
        End Sub

        Public Shared Function Build(names As SharePointListNames) As IReadOnlyList(Of ListSpec)
            Dim configured = If(names, SharePointListNames.Default())
            Return New List(Of ListSpec) From {
                Bookings(configured.Bookings),
                Students(configured.Students),
                Teachers(configured.Teachers),
                Units(configured.Units),
                Qualifications(configured.Qualifications),
                Signature(configured.SignatureLibrary)
            }
        End Function

        Private Shared Function TextColumn(name As String, displayName As String, Optional indexed As Boolean = False, Optional unique As Boolean = False) As ColumnSpec
            Return New ColumnSpec With {
                .Name = name,
                .DisplayName = displayName,
                .Kind = ColumnKind.Text,
                .Indexed = indexed OrElse unique,
                .EnforceUniqueValues = unique
            }
        End Function

        Private Shared Function Bookings(displayName As String) As ListSpec
            Return New ListSpec With {
                .DisplayName = displayName,
                .Template = ListTemplateKind.GenericList,
                .Columns = New List(Of ColumnSpec) From {
                    TextColumn(SharePointColumns.StudentId, "Student ID", unique:=True),
                    TextColumn(SharePointColumns.Unit, "Unit"),
                    TextColumn(SharePointColumns.QualificationCode, "QualificationCode", indexed:=True),
                    New ColumnSpec With {.Name = SharePointColumns.ResitDate, .DisplayName = "Resit date", .Kind = ColumnKind.DateOnly, .Indexed = True},
                    TextColumn(SharePointColumns.Blockgroup, "Blockgroup"),
                    TextColumn(SharePointColumns.AllocatedTeacher, "AllocatedTeacher"),
                    TextColumn(SharePointColumns.Status, "status"),
                    TextColumn(SharePointColumns.AllocatedTeacherEmail, "AllocatedTeacher email"),
                    New ColumnSpec With {.Name = SharePointColumns.EnergyspaceClassCreated, .DisplayName = "EnergyspaceClassCreated", .Kind = ColumnKind.YesNo},
                    New ColumnSpec With {.Name = SharePointColumns.EnergyspaceAssessmentBooked, .DisplayName = "EnergyspaceAssessmentBooked", .Kind = ColumnKind.YesNo},
                    TextColumn(SharePointColumns.AttemptNo, "AttemptNo"),
                    TextColumn(SharePointColumns.StudentFirstname, "Student Firstname"),
                    TextColumn(SharePointColumns.Surname, "Surname"),
                    TextColumn(SharePointColumns.Email, "Email"),
                    TextColumn(SharePointColumns.UnitName, "Unit Name")
                }
            }
        End Function

        Private Shared Function Students(displayName As String) As ListSpec
            Return New ListSpec With {
                .DisplayName = displayName,
                .Template = ListTemplateKind.GenericList,
                .Columns = New List(Of ColumnSpec) From {
                    TextColumn(SharePointColumns.StudentId, "Student ID", unique:=True),
                    TextColumn(SharePointColumns.StudentGivenName, "Student Given Name"),
                    TextColumn(SharePointColumns.StudentFamilyName, "Student Family Name"),
                    TextColumn(SharePointColumns.StudentPersonalEmail, "Student Personal Email"),
                    TextColumn(SharePointColumns.BlockGroupCode, "Block Group Code")
                }
            }
        End Function

        Private Shared Function Teachers(displayName As String) As ListSpec
            Return New ListSpec With {
                .DisplayName = displayName,
                .Template = ListTemplateKind.GenericList,
                .Columns = New List(Of ColumnSpec) From {
                    TextColumn(SharePointColumns.TeacherFullName, "Teacher Full Name", indexed:=True),
                    TextColumn(SharePointColumns.Email, "Email")
                }
            }
        End Function

        Private Shared Function Units(displayName As String) As ListSpec
            Return New ListSpec With {
                .DisplayName = displayName,
                .Template = ListTemplateKind.GenericList,
                .Columns = New List(Of ColumnSpec) From {
                    TextColumn(SharePointColumns.UnitCode, "Unit Code", indexed:=True),
                    TextColumn(SharePointColumns.UnitTitle, "Unit Title"),
                    TextColumn(SharePointColumns.QualificationCode, "QualificationCode", indexed:=True)
                }
            }
        End Function

        Private Shared Function Qualifications(displayName As String) As ListSpec
            Return New ListSpec With {
                .DisplayName = displayName,
                .Template = ListTemplateKind.GenericList,
                .Columns = New List(Of ColumnSpec) From {
                    TextColumn(SharePointColumns.QualificationCode, "QualificationCode", unique:=True),
                    TextColumn(SharePointColumns.QualificationTitle, "Qualification title"),
                    TextColumn(SharePointColumns.Version, "Version"),
                    TextColumn(SharePointColumns.Status, "Status", indexed:=True),
                    TextColumn(SharePointColumns.ClassGroupSuffix, "ClassGroupSuffix"),
                    TextColumn(SharePointColumns.AbbreviationRules, "AbbreviationRules")
                }
            }
        End Function

        Private Shared Function Signature(displayName As String) As ListSpec
            Return New ListSpec With {
                .DisplayName = displayName,
                .Template = ListTemplateKind.DocumentLibrary,
                .Columns = New List(Of ColumnSpec)
            }
        End Function
    End Class

    Public NotInheritable Class SchemaComparer
        Private Sub New()
        End Sub

        Public Shared Function Compare(expected As IEnumerable(Of ListSpec), actual As IEnumerable(Of ExistingList)) As SchemaReport
            Dim report As New SchemaReport()
            Dim lists = If(actual, Enumerable.Empty(Of ExistingList)()).ToList()
            For Each spec In If(expected, Enumerable.Empty(Of ListSpec)())
                Dim found = lists.FirstOrDefault(Function(list) String.Equals(list.DisplayName, spec.DisplayName, StringComparison.OrdinalIgnoreCase))
                If found Is Nothing Then
                    report.MissingLists.Add(spec.DisplayName)
                    Continue For
                End If
                If found.Template <> spec.Template Then
                    Dim expectedKind = If(spec.Template = ListTemplateKind.DocumentLibrary, "a document library", "a list")
                    report.TemplateMismatches.Add(spec.DisplayName & " is not " & expectedKind)
                    Continue For
                End If
                Dim columns = If(found.Columns, New List(Of ExistingColumn)())
                For Each column In If(spec.Columns, New List(Of ColumnSpec)())
                    Dim existing = columns.FirstOrDefault(Function(item) String.Equals(item.Name, column.Name, StringComparison.OrdinalIgnoreCase))
                    If existing Is Nothing Then
                        report.MissingColumns.Add(New SchemaGap With {.ListName = spec.DisplayName, .ColumnName = column.Name})
                        Continue For
                    End If
                    If column.Indexed AndAlso Not existing.Indexed Then
                        report.IndexGaps.Add(New SchemaGap With {.ListName = spec.DisplayName, .ColumnName = column.Name})
                    End If
                    If column.EnforceUniqueValues AndAlso Not existing.EnforceUniqueValues Then
                        report.UniqueGaps.Add(New SchemaGap With {.ListName = spec.DisplayName, .ColumnName = column.Name})
                    End If
                Next
            Next
            Return report
        End Function
    End Class

    Public Interface ISharePointSchemaStore
        Function ReadListsAsync() As Task(Of IReadOnlyList(Of ExistingList))
        Function CreateListAsync(spec As ListSpec) As Task
        Function AddColumnAsync(listDisplayName As String, column As ColumnSpec) As Task
        Function UpdateColumnAsync(listDisplayName As String, columnName As String, indexed As Boolean, enforceUniqueValues As Boolean) As Task
    End Interface

    Public Interface ISharePointProvisioner
        Function VerifyAsync() As Task(Of SchemaReport)
        Function EnsureAsync() As Task(Of SchemaReport)
    End Interface

    ''' <summary>
    ''' Creates missing lists and columns and turns on missing indexes.
    ''' Safe to run again. It never deletes a list, a column, or a row.
    ''' </summary>
    Public Class ListProvisioner
        Implements ISharePointProvisioner

        Private ReadOnly _store As ISharePointSchemaStore
        Private ReadOnly _names As SharePointListNames

        Public Sub New(store As ISharePointSchemaStore, names As SharePointListNames)
            _store = store
            _names = If(names, SharePointListNames.Default())
        End Sub

        Public Async Function VerifyAsync() As Task(Of SchemaReport) Implements ISharePointProvisioner.VerifyAsync
            Dim actual = Await _store.ReadListsAsync().ConfigureAwait(False)
            Return SchemaComparer.Compare(SharePointListSchema.Build(_names), actual)
        End Function

        Public Async Function EnsureAsync() As Task(Of SchemaReport) Implements ISharePointProvisioner.EnsureAsync
            Dim expected = SharePointListSchema.Build(_names)
            For pass As Integer = 1 To 4
                Dim report = SchemaComparer.Compare(expected, Await _store.ReadListsAsync().ConfigureAwait(False))
                If report.IsSatisfied Then Return report

                Dim acted = False
                For Each listName In report.MissingLists
                    Dim spec = expected.First(Function(item) String.Equals(item.DisplayName, listName, StringComparison.OrdinalIgnoreCase))
                    Await _store.CreateListAsync(spec).ConfigureAwait(False)
                    acted = True
                Next
                If acted Then Continue For

                For Each gap In report.MissingColumns
                    Dim column = FindColumn(expected, gap)
                    Await _store.AddColumnAsync(gap.ListName, column).ConfigureAwait(False)
                    acted = True
                Next
                If acted Then Continue For

                Dim patched As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                For Each gap In report.IndexGaps
                    patched.Add(gap.ToString())
                Next
                For Each gap In report.UniqueGaps
                    patched.Add(gap.ToString())
                Next
                For Each key In patched
                    Dim gap = report.IndexGaps.Concat(report.UniqueGaps).First(Function(item) item.ToString().Equals(key, StringComparison.OrdinalIgnoreCase))
                    Dim column = FindColumn(expected, gap)
                    Await _store.UpdateColumnAsync(gap.ListName, column.Name, column.Indexed OrElse column.EnforceUniqueValues, column.EnforceUniqueValues).ConfigureAwait(False)
                    acted = True
                Next
                If Not acted Then Return report
            Next
            Return SchemaComparer.Compare(expected, Await _store.ReadListsAsync().ConfigureAwait(False))
        End Function

        Private Shared Function FindColumn(expected As IEnumerable(Of ListSpec), gap As SchemaGap) As ColumnSpec
            Dim list = expected.First(Function(item) String.Equals(item.DisplayName, gap.ListName, StringComparison.OrdinalIgnoreCase))
            Return list.Columns.First(Function(item) String.Equals(item.Name, gap.ColumnName, StringComparison.OrdinalIgnoreCase))
        End Function
    End Class

    ''' <summary>In-memory lists used to test setup. There is no delete operation.</summary>
    Public Class MemorySchemaStore
        Implements ISharePointSchemaStore

        Public ReadOnly Operations As New List(Of String)
        Public ReadOnly Lists As New List(Of ExistingList)
        Public Marker As String = "kept"

        Public Function ReadListsAsync() As Task(Of IReadOnlyList(Of ExistingList)) Implements ISharePointSchemaStore.ReadListsAsync
            Dim copy As New List(Of ExistingList)
            For Each list In Lists
                Dim columns As New List(Of ExistingColumn)
                If list.Columns IsNot Nothing Then
                    For Each column In list.Columns
                        columns.Add(New ExistingColumn With {
                            .Name = column.Name,
                            .Id = column.Id,
                            .Indexed = column.Indexed,
                            .EnforceUniqueValues = column.EnforceUniqueValues
                        })
                    Next
                End If
                copy.Add(New ExistingList With {
                    .DisplayName = list.DisplayName,
                    .Id = list.Id,
                    .Template = list.Template,
                    .Columns = columns
                })
            Next
            Return Task.FromResult(Of IReadOnlyList(Of ExistingList))(copy)
        End Function

        Public Function CreateListAsync(spec As ListSpec) As Task Implements ISharePointSchemaStore.CreateListAsync
            Operations.Add("create:" & spec.DisplayName)
            Dim existing = Lists.FirstOrDefault(Function(item) String.Equals(item.DisplayName, spec.DisplayName, StringComparison.OrdinalIgnoreCase))
            If existing Is Nothing Then
                Lists.Add(New ExistingList With {
                    .DisplayName = spec.DisplayName,
                    .Id = Guid.NewGuid().ToString("N"),
                    .Template = spec.Template,
                    .Columns = New List(Of ExistingColumn)
                })
            End If
            Return Task.CompletedTask
        End Function

        Public Function AddColumnAsync(listDisplayName As String, column As ColumnSpec) As Task Implements ISharePointSchemaStore.AddColumnAsync
            Operations.Add("add:" & listDisplayName & "." & column.Name)
            Dim list = RequireList(listDisplayName)
            Dim existing = list.Columns.FirstOrDefault(Function(item) String.Equals(item.Name, column.Name, StringComparison.OrdinalIgnoreCase))
            If existing Is Nothing Then
                list.Columns.Add(New ExistingColumn With {
                    .Name = column.Name,
                    .Id = Guid.NewGuid().ToString("N"),
                    .Indexed = column.Indexed,
                    .EnforceUniqueValues = column.EnforceUniqueValues
                })
            End If
            Return Task.CompletedTask
        End Function

        Public Function UpdateColumnAsync(listDisplayName As String, columnName As String, indexed As Boolean, enforceUniqueValues As Boolean) As Task Implements ISharePointSchemaStore.UpdateColumnAsync
            Operations.Add("patch:" & listDisplayName & "." & columnName)
            Dim list = RequireList(listDisplayName)
            Dim column = list.Columns.First(Function(item) String.Equals(item.Name, columnName, StringComparison.OrdinalIgnoreCase))
            column.Indexed = indexed
            column.EnforceUniqueValues = enforceUniqueValues
            Return Task.CompletedTask
        End Function

        Private Function RequireList(displayName As String) As ExistingList
            Dim list = Lists.FirstOrDefault(Function(item) String.Equals(item.DisplayName, displayName, StringComparison.OrdinalIgnoreCase))
            If list Is Nothing Then Throw New InvalidOperationException("List '" & displayName & "' was not found.")
            Return list
        End Function
    End Class

End Namespace
