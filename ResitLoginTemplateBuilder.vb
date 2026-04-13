Imports System.IO
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing
Imports WdColor = DocumentFormat.OpenXml.Wordprocessing.Color

''' <summary>Creates the default Resit login .docx (Open XML). Layout/fonts match the ESSW handout sample (Arial PC block, TNR class list, Source Sans daily password).</summary>
Friend Module ResitLoginTemplateBuilder
    Private Const ColorText As String = "000000"
    Private Const ColorClassGroup As String = "444444"
    Private Const ColorDailyRed As String = "FF0000"
    Private Const FillPanel As String = "EEF0F8"
    Private Const FillWhite As String = "FFFFFF"

    Private Const FontPc As String = "Arial"
    Private Const FontClassGroup As String = "Times New Roman"
    Private Const FontDaily As String = "Source Sans Pro"

    ' Word font sizes are in half-points.
    Private Const SzDate As String = "24"      ' 12 pt — date line (top right)
    Private Const SzPc As String = "28"        ' 14 pt — PC login block
    Private Const SzSection As String = "32"   ' 16 pt — rule, section headings
    Private Const SzDailyLabel As String = "48"   ' 24 pt — “Daily Energyspace Password:”
    Private Const SzDailyPassword As String = "144" ' 72 pt — daily password (sample)

    ''' <summary>Horizontal rule under the PC block — kept short so it stays on one line at 16 pt on A4.</summary>
    Private ReadOnly SeparatorLine As String = New String("-"c, 52)

    ''' <summary>Creates the file next to the app if it does not exist.</summary>
    Friend Sub EnsureDefaultTemplateExists()
        Dim path = ResitLoginTemplatePaths.GetBundledTemplateFullPath()
        If File.Exists(path) Then Return
        Try
            CreateDefaultTemplate(path)
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("ResitLoginTemplateBuilder: " & ex.ToString())
        End Try
    End Sub

    Friend Sub CreateDefaultTemplate(outputPath As String)
        Dim dir = Path.GetDirectoryName(outputPath)
        If Not String.IsNullOrEmpty(dir) Then Directory.CreateDirectory(dir)

        Using wordDoc As WordprocessingDocument = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document)
            Dim mainPart As MainDocumentPart = wordDoc.AddMainDocumentPart()
            Dim body As New Body()

            body.AppendChild(ParaDateRight("{{RESIT_NIGHT_DATE}}"))
            body.AppendChild(SpacerPara())
            body.AppendChild(SpacerPara())

            body.AppendChild(ParaPcLoginBlock())
            body.AppendChild(ParaSeparatorRule())
            body.AppendChild(ParaClassGroupsHeading())
            body.AppendChild(ParaClassGroupsPlaceholder())
            body.AppendChild(SpacerPara())
            body.AppendChild(SpacerPara())
            body.AppendChild(SpacerPara())

            body.AppendChild(ParaDailyPasswordLabel())
            body.AppendChild(SpacerPara())
            body.AppendChild(ParaDailyPasswordValue())

            body.AppendChild(DefaultSectionProperties())

            mainPart.Document = New Document(body)
            mainPart.Document.Save()
        End Using
    End Sub

    Private Function DefaultSectionProperties() As SectionProperties
        Return New SectionProperties(
            New PageSize() With {.Width = 11906, .Height = 16838},
            New PageMargin() With {.Top = 1440, .Right = 1440, .Bottom = 1440, .Left = 1440, .Header = 708, .Footer = 708, .Gutter = 0},
            New Columns() With {.Space = "708"},
            New DocGrid() With {.LinePitch = 360})
    End Function

    Private Function SpacerPara() As Paragraph
        Dim p As New Paragraph()
        p.AppendChild(New ParagraphProperties(New SpacingBetweenLines() With {.After = "0", .Line = "240", .LineRule = LineSpacingRuleValues.Auto}))
        Dim r As New Run()
        r.AppendChild(New RunProperties(New FontSize() With {.Val = SzSection}, New RunFonts() With {.Ascii = FontClassGroup, .HighAnsi = FontClassGroup}))
        r.AppendChild(MakeText(""))
        p.AppendChild(r)
        Return p
    End Function

    Private Function ParaDateRight(placeholder As String) As Paragraph
        Dim p As New Paragraph()
        p.AppendChild(New ParagraphProperties(
            New Justification() With {.Val = JustificationValues.Right},
            New SpacingBetweenLines() With {.After = "0", .Line = "240", .LineRule = LineSpacingRuleValues.Auto}))
        Dim r As New Run()
        r.AppendChild(New RunProperties(
            New FontSize() With {.Val = SzDate},
            New FontSizeComplexScript() With {.Val = SzDate},
            New WdColor() With {.Val = ColorText},
            New RunFonts() With {.Ascii = FontPc, .HighAnsi = FontPc, .ComplexScript = FontPc}))
        r.AppendChild(MakeText(placeholder))
        p.AppendChild(r)
        Return p
    End Function

    ''' <summary>Single paragraph with line breaks — matches sample PC block.</summary>
    Private Function ParaPcLoginBlock() As Paragraph
        Dim p As New Paragraph()
        p.AppendChild(New ParagraphProperties(New SpacingBetweenLines() With {.After = "120", .Line = "240", .LineRule = LineSpacingRuleValues.Auto}))

        p.AppendChild(RunPcText("PC LOGIN DETAILS (for those that don't have login details on a PC)"))
        p.AppendChild(New Run(New Break()))
        p.AppendChild(RunPcText("{{USERNAME_LINE}}"))
        p.AppendChild(New Run(New Break()))
        p.AppendChild(RunPcText("{{PASSWORD_LINE}}"))
        Return p
    End Function

    Private Function RunPcText(text As String) As Run
        Dim r As New Run()
        r.AppendChild(New RunProperties(
            New FontSize() With {.Val = SzPc},
            New FontSizeComplexScript() With {.Val = SzPc},
            New WdColor() With {.Val = ColorText},
            New RunFonts() With {.Ascii = FontPc, .HighAnsi = FontPc, .ComplexScript = FontPc}))
        r.AppendChild(MakeText(text))
        Return r
    End Function

    Private Function ParaSeparatorRule() As Paragraph
        Dim p As New Paragraph()
        p.AppendChild(New ParagraphProperties(New SpacingBetweenLines() With {.After = "120", .Line = "240", .LineRule = LineSpacingRuleValues.Auto}))
        Dim r As New Run()
        r.AppendChild(New RunProperties(
            New FontSize() With {.Val = SzSection},
            New FontSizeComplexScript() With {.Val = SzSection},
            New WdColor() With {.Val = ColorText},
            New RunFonts() With {.Ascii = FontPc, .HighAnsi = FontPc, .ComplexScript = FontPc}))
        r.AppendChild(MakeText(SeparatorLine))
        p.AppendChild(r)
        Return p
    End Function

    Private Function ParaClassGroupsHeading() As Paragraph
        Dim p As New Paragraph()
        p.AppendChild(New ParagraphProperties(New SpacingBetweenLines() With {.After = "120", .Line = "240", .LineRule = LineSpacingRuleValues.Auto}))
        Dim r As New Run()
        r.AppendChild(New RunProperties(
            New FontSize() With {.Val = SzSection},
            New FontSizeComplexScript() With {.Val = SzSection},
            New WdColor() With {.Val = ColorText},
            New RunFonts() With {.Ascii = FontPc, .HighAnsi = FontPc, .ComplexScript = FontPc}))
        r.AppendChild(MakeText("Energyspace resit class groups:"))
        p.AppendChild(r)
        Return p
    End Function

    Private Function ParaClassGroupsPlaceholder() As Paragraph
        Dim p As New Paragraph()
        Dim pPr As New ParagraphProperties(
            New Shading() With {.Val = ShadingPatternValues.Clear, .Color = "auto", .Fill = FillWhite},
            New SpacingBetweenLines() With {.Line = "270", .LineRule = LineSpacingRuleValues.AtLeast, .After = "0"})
        p.AppendChild(pPr)
        Dim r As New Run()
        r.AppendChild(New RunProperties(
            New FontSize() With {.Val = SzSection},
            New FontSizeComplexScript() With {.Val = SzSection},
            New WdColor() With {.Val = ColorClassGroup},
            New RunFonts() With {.Ascii = FontClassGroup, .HighAnsi = FontClassGroup, .EastAsia = FontClassGroup}))
        r.AppendChild(MakeText("{{CLASS_GROUPS}}"))
        p.AppendChild(r)
        Return p
    End Function

    Private Function ParaDailyPasswordLabel() As Paragraph
        Dim p As New Paragraph()
        p.AppendChild(New ParagraphProperties(
            New Shading() With {.Val = ShadingPatternValues.Clear, .Color = "auto", .Fill = FillWhite},
            New SpacingBetweenLines() With {.Line = "270", .LineRule = LineSpacingRuleValues.AtLeast, .After = "0"}))
        Dim r As New Run()
        r.AppendChild(New RunProperties(
            New FontSize() With {.Val = SzDailyLabel},
            New FontSizeComplexScript() With {.Val = SzDailyLabel},
            New WdColor() With {.Val = ColorText},
            New RunFonts() With {.Ascii = FontDaily, .HighAnsi = FontDaily, .ComplexScript = FontDaily},
            New Shading() With {.Val = ShadingPatternValues.Clear, .Color = "auto", .Fill = FillPanel}))
        r.AppendChild(MakeText("Daily Energyspace Password: "))
        p.AppendChild(r)
        Return p
    End Function

    Private Function ParaDailyPasswordValue() As Paragraph
        Dim p As New Paragraph()
        p.AppendChild(New ParagraphProperties(
            New Justification() With {.Val = JustificationValues.Center},
            New Shading() With {.Val = ShadingPatternValues.Clear, .Color = "auto", .Fill = FillWhite},
            New SpacingBetweenLines() With {.Line = "270", .LineRule = LineSpacingRuleValues.AtLeast, .After = "0"}))
        Dim r As New Run()
        r.AppendChild(New RunProperties(
            New FontSize() With {.Val = SzDailyPassword},
            New FontSizeComplexScript() With {.Val = SzDailyPassword},
            New WdColor() With {.Val = ColorDailyRed},
            New RunFonts() With {.Ascii = FontDaily, .HighAnsi = FontDaily, .ComplexScript = FontDaily},
            New Shading() With {.Val = ShadingPatternValues.Clear, .Color = "auto", .Fill = FillPanel}))
        r.AppendChild(MakeText("{{DAILY_ENERGYSPACE_PASSWORD}}"))
        p.AppendChild(r)
        Return p
    End Function

    Private Function MakeText(text As String) As Text
        Dim t As New Text(If(text, ""))
        t.Space = SpaceProcessingModeValues.Preserve
        Return t
    End Function
End Module
