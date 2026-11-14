Imports System.IO
Imports ElectrotechnologyResit.Core
Imports Xunit

Public Class SampleDataTests
    <Fact>
    Public Sub Seed_is_fictional_and_covers_several_nights()
        Dim today = New Date(2026, 11, 14)
        Dim seed = SampleDataCatalog.Build(today)
        Assert.InRange(seed.Bookings.Count, 12, 15)
        Assert.Equal(3, seed.Bookings.Select(Function(item) item.ResitDate).Distinct().Count())
        Assert.Contains(seed.Bookings, Function(item) item.ResitDate = today)
        Assert.Contains(seed.Bookings, Function(item) item.ResitDate = today.AddDays(-16))
        Assert.Contains(seed.Bookings, Function(item) item.EnergyspaceCreated)
        Assert.Contains(seed.Bookings, Function(item) Not item.EnergyspaceCreated)
        For Each booking In seed.Bookings
            Assert.Matches("^90000\d{2}$", booking.StudentId)
            Assert.EndsWith("@example.com", booking.Email)
            Assert.EndsWith("@example.com", booking.TeacherEmail)
            Assert.DoesNotContain("vu.edu.au", booking.Email, StringComparison.OrdinalIgnoreCase)
        Next
        For Each teacher In seed.Teachers
            Assert.EndsWith("@example.com", teacher.Email)
        Next
        Dim expectedUnits = New String() {"UEECD0019", "UEECD0020", "UEECD0046", "UEEEL0005", "UEEEL0012", "UEEEL0021", "UEEEL0025"}
        Assert.Equal(expectedUnits.Length, seed.Units.Count)
        For Each code In expectedUnits
            Assert.Contains(seed.Units, Function(item) item.UnitCode = code)
            Assert.Contains(seed.Bookings, Function(item) item.UnitCode = code)
        Next
        Assert.Equal(14, seed.Bookings.Count)
    End Sub

    <Fact>
    Public Async Function Sample_store_persists_and_reset_restores_the_seed() As Task
        Dim folder = Path.Combine(Path.GetTempPath(), "resit-sample-" & Guid.NewGuid().ToString("N"))
        Dim today = New Date(2026, 11, 14)
        Try
            Dim store As New SampleDataStore(folder, today)
            Dim original = (Await store.GetBookingsAsync()).Count
            Assert.InRange(original, 12, 15)
            Await store.UpsertAsync(New ResitBooking With {
                .StudentId = "9000099",
                .FirstName = "Quinn",
                .Surname = "Adler",
                .Email = "quinn.adler@example.com",
                .UnitCode = "UEECD0019",
                .UnitName = "Fabricate, assemble and dismantle utilities industry components",
                .TeacherName = "Jordan Hale",
                .TeacherEmail = "not-used@example.com",
                .AttemptNo = "1",
                .ResitDate = today
            }, False)
            Dim reopened As New SampleDataStore(folder, today)
            Dim booking = (Await reopened.GetBookingsAsync()).Single(Function(item) item.StudentId = "9000099")
            Assert.Equal("Quinn", booking.FirstName)
            Assert.Equal("jordan.hale@example.com", booking.TeacherEmail)
            Await reopened.DeleteOnOrBeforeAsync(today.AddDays(-14))
            Assert.DoesNotContain(Await reopened.GetBookingsAsync(), Function(item) item.ResitDate <= today.AddDays(-14))
            reopened.Reset(today)
            Assert.Equal(original, (Await reopened.GetBookingsAsync()).Count)
            Assert.Null(Await reopened.FindByStudentIdAsync("9000099"))
            Assert.NotNull(Await reopened.GetCurrentImageAsync())
        Finally
            If Directory.Exists(folder) Then Directory.Delete(folder, True)
        End Try
    End Function
End Class
