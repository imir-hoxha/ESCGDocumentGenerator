using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using Thm15 = DocumentFormat.OpenXml.Office2013.Theme;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using M = DocumentFormat.OpenXml.Math;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using V = DocumentFormat.OpenXml.Vml;
using W14 = DocumentFormat.OpenXml.Office2010.Word;

namespace ConsoleApp1
{
    public class GeneratedClass
    {
        // Creates a WordprocessingDocument.
        public void CreatePackage(string filePath)
        {
            using (WordprocessingDocument package = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                CreateParts(package);
            }
        }

        // Adds child parts and generates content of the specified part.
        private void CreateParts(WordprocessingDocument document)
        {
            ExtendedFilePropertiesPart extendedFilePropertiesPart1 = document.AddNewPart<ExtendedFilePropertiesPart>("rId3");
            GenerateExtendedFilePropertiesPart1Content(extendedFilePropertiesPart1);

            MainDocumentPart mainDocumentPart1 = document.AddMainDocumentPart();
            GenerateMainDocumentPart1Content(mainDocumentPart1);

            ThemePart themePart1 = mainDocumentPart1.AddNewPart<ThemePart>("rId8");
            GenerateThemePart1Content(themePart1);

            WebSettingsPart webSettingsPart1 = mainDocumentPart1.AddNewPart<WebSettingsPart>("rId3");
            GenerateWebSettingsPart1Content(webSettingsPart1);

            WordprocessingPeoplePart wordprocessingPeoplePart1 = mainDocumentPart1.AddNewPart<WordprocessingPeoplePart>("rId7");
            GenerateWordprocessingPeoplePart1Content(wordprocessingPeoplePart1);

            DocumentSettingsPart documentSettingsPart1 = mainDocumentPart1.AddNewPart<DocumentSettingsPart>("rId2");
            GenerateDocumentSettingsPart1Content(documentSettingsPart1);

            StyleDefinitionsPart styleDefinitionsPart1 = mainDocumentPart1.AddNewPart<StyleDefinitionsPart>("rId1");
            GenerateStyleDefinitionsPart1Content(styleDefinitionsPart1);

            FontTablePart fontTablePart1 = mainDocumentPart1.AddNewPart<FontTablePart>("rId6");
            GenerateFontTablePart1Content(fontTablePart1);

            WordprocessingCommentsExPart wordprocessingCommentsExPart1 = mainDocumentPart1.AddNewPart<WordprocessingCommentsExPart>("rId5");
            GenerateWordprocessingCommentsExPart1Content(wordprocessingCommentsExPart1);

            WordprocessingCommentsPart wordprocessingCommentsPart1 = mainDocumentPart1.AddNewPart<WordprocessingCommentsPart>("rId4");
            GenerateWordprocessingCommentsPart1Content(wordprocessingCommentsPart1);

            SetPackageProperties(document);
        }

        // Generates content of extendedFilePropertiesPart1.
        private void GenerateExtendedFilePropertiesPart1Content(ExtendedFilePropertiesPart extendedFilePropertiesPart1)
        {
            Ap.Properties properties1 = new Ap.Properties();
            properties1.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            Ap.Template template1 = new Ap.Template();
            template1.Text = "Normal.dotm";
            Ap.TotalTime totalTime1 = new Ap.TotalTime();
            totalTime1.Text = "2";
            Ap.Pages pages1 = new Ap.Pages();
            pages1.Text = "2";
            Ap.Words words1 = new Ap.Words();
            words1.Text = "379";
            Ap.Characters characters1 = new Ap.Characters();
            characters1.Text = "2285";
            Ap.Application application1 = new Ap.Application();
            application1.Text = "Microsoft Office Word";
            Ap.DocumentSecurity documentSecurity1 = new Ap.DocumentSecurity();
            documentSecurity1.Text = "0";
            Ap.Lines lines1 = new Ap.Lines();
            lines1.Text = "30";
            Ap.Paragraphs paragraphs1 = new Ap.Paragraphs();
            paragraphs1.Text = "11";
            Ap.ScaleCrop scaleCrop1 = new Ap.ScaleCrop();
            scaleCrop1.Text = "false";
            Ap.Company company1 = new Ap.Company();
            company1.Text = "European Commission";
            Ap.LinksUpToDate linksUpToDate1 = new Ap.LinksUpToDate();
            linksUpToDate1.Text = "false";
            Ap.CharactersWithSpaces charactersWithSpaces1 = new Ap.CharactersWithSpaces();
            charactersWithSpaces1.Text = "2653";
            Ap.SharedDocument sharedDocument1 = new Ap.SharedDocument();
            sharedDocument1.Text = "false";
            Ap.HyperlinksChanged hyperlinksChanged1 = new Ap.HyperlinksChanged();
            hyperlinksChanged1.Text = "false";
            Ap.ApplicationVersion applicationVersion1 = new Ap.ApplicationVersion();
            applicationVersion1.Text = "16.0000";

            properties1.Append(template1);
            properties1.Append(totalTime1);
            properties1.Append(pages1);
            properties1.Append(words1);
            properties1.Append(characters1);
            properties1.Append(application1);
            properties1.Append(documentSecurity1);
            properties1.Append(lines1);
            properties1.Append(paragraphs1);
            properties1.Append(scaleCrop1);
            properties1.Append(company1);
            properties1.Append(linksUpToDate1);
            properties1.Append(charactersWithSpaces1);
            properties1.Append(sharedDocument1);
            properties1.Append(hyperlinksChanged1);
            properties1.Append(applicationVersion1);

            extendedFilePropertiesPart1.Properties = properties1;
        }

        // Generates content of mainDocumentPart1.
        private void GenerateMainDocumentPart1Content(MainDocumentPart mainDocumentPart1)
        {
            Document document1 = new Document() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se wp14" } };
            document1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            document1.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
            document1.AddNamespaceDeclaration("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
            document1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            document1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            document1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            document1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            document1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            document1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            document1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            document1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            document1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            document1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            document1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            document1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            document1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            document1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            document1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            document1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            Body body1 = new Body();

            Paragraph paragraph1 = new Paragraph() { RsidParagraphMarkRevision = "00E075BD", RsidParagraphAddition = "005F6638", RsidParagraphProperties = "005F6638", RsidRunAdditionDefault = "005F6638", ParagraphId = "7D64A9AE", TextId = "77777777" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Heading1" };
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { After = "360" };

            ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
            RunFonts runFonts1 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Color color1 = new Color() { Val = "auto" };
            FontSize fontSize1 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties1.Append(runFonts1);
            paragraphMarkRunProperties1.Append(color1);
            paragraphMarkRunProperties1.Append(fontSize1);
            paragraphMarkRunProperties1.Append(fontSizeComplexScript1);

            paragraphProperties1.Append(paragraphStyleId1);
            paragraphProperties1.Append(spacingBetweenLines1);
            paragraphProperties1.Append(paragraphMarkRunProperties1);

            Run run1 = new Run();

            RunProperties runProperties1 = new RunProperties();
            RunFonts runFonts2 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Italic italic1 = new Italic();
            Color color2 = new Color() { Val = "auto" };
            FontSize fontSize2 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "24" };

            runProperties1.Append(runFonts2);
            runProperties1.Append(italic1);
            runProperties1.Append(color2);
            runProperties1.Append(fontSize2);
            runProperties1.Append(fontSizeComplexScript2);
            Text text1 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text1.Text = "(C1) ";

            run1.Append(runProperties1);
            run1.Append(text1);

            Run run2 = new Run() { RsidRunProperties = "00E075BD" };

            RunProperties runProperties2 = new RunProperties();
            RunFonts runFonts3 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Italic italic2 = new Italic();
            Color color3 = new Color() { Val = "auto" };
            FontSize fontSize3 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "24" };

            runProperties2.Append(runFonts3);
            runProperties2.Append(italic2);
            runProperties2.Append(color3);
            runProperties2.Append(fontSize3);
            runProperties2.Append(fontSizeComplexScript3);
            Text text2 = new Text();
            text2.Text = "FINLAND";

            run2.Append(runProperties2);
            run2.Append(text2);

            Run run3 = new Run();

            RunProperties runProperties3 = new RunProperties();
            RunFonts runFonts4 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Italic italic3 = new Italic();
            Color color4 = new Color() { Val = "auto" };
            FontSize fontSize4 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "24" };

            runProperties3.Append(runFonts4);
            runProperties3.Append(italic3);
            runProperties3.Append(color4);
            runProperties3.Append(fontSize4);
            runProperties3.Append(fontSizeComplexScript4);
            Text text3 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text3.Text = " – (C24) Lead DG";

            run3.Append(runProperties3);
            run3.Append(text3);

            Run run4 = new Run();

            RunProperties runProperties4 = new RunProperties();
            RunFonts runFonts5 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Color color5 = new Color() { Val = "auto" };
            FontSize fontSize5 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "24" };

            runProperties4.Append(runFonts5);
            runProperties4.Append(color5);
            runProperties4.Append(fontSize5);
            runProperties4.Append(fontSizeComplexScript5);
            Text text4 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text4.Text = ": (C2) ";

            run4.Append(runProperties4);
            run4.Append(text4);

            Run run5 = new Run() { RsidRunProperties = "00E075BD" };

            RunProperties runProperties5 = new RunProperties();
            RunFonts runFonts6 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Color color6 = new Color() { Val = "auto" };
            FontSize fontSize6 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "24" };

            runProperties5.Append(runFonts6);
            runProperties5.Append(color6);
            runProperties5.Append(fontSize6);
            runProperties5.Append(fontSizeComplexScript6);
            Text text5 = new Text();
            text5.Text = "F";

            run5.Append(runProperties5);
            run5.Append(text5);

            Run run6 = new Run();

            RunProperties runProperties6 = new RunProperties();
            RunFonts runFonts7 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Color color7 = new Color() { Val = "auto" };
            FontSize fontSize7 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript7 = new FontSizeComplexScript() { Val = "24" };

            runProperties6.Append(runFonts7);
            runProperties6.Append(color7);
            runProperties6.Append(fontSize7);
            runProperties6.Append(fontSizeComplexScript7);
            Text text6 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text6.Text = "ailure to adopt and/or to communicate all measures for transposing Council Directive ";

            run6.Append(runProperties6);
            run6.Append(text6);

            Run run7 = new Run() { RsidRunProperties = "00E075BD" };

            RunProperties runProperties7 = new RunProperties();
            RunFonts runFonts8 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Color color8 = new Color() { Val = "auto" };
            FontSize fontSize8 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript8 = new FontSizeComplexScript() { Val = "24" };

            runProperties7.Append(runFonts8);
            runProperties7.Append(color8);
            runProperties7.Append(fontSize8);
            runProperties7.Append(fontSizeComplexScript8);
            Text text7 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text7.Text = "2009/119/EC ";

            run7.Append(runProperties7);
            run7.Append(text7);

            Run run8 = new Run();

            RunProperties runProperties8 = new RunProperties();
            RunFonts runFonts9 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Color color9 = new Color() { Val = "auto" };
            FontSize fontSize9 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript9 = new FontSizeComplexScript() { Val = "24" };

            runProperties8.Append(runFonts9);
            runProperties8.Append(color9);
            runProperties8.Append(fontSize9);
            runProperties8.Append(fontSizeComplexScript9);
            Text text8 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text8.Text = "of 14 September ";

            run8.Append(runProperties8);
            run8.Append(text8);

            Run run9 = new Run() { RsidRunProperties = "00E075BD" };

            RunProperties runProperties9 = new RunProperties();
            RunFonts runFonts10 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Color color10 = new Color() { Val = "auto" };
            FontSize fontSize10 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript10 = new FontSizeComplexScript() { Val = "24" };

            runProperties9.Append(runFonts10);
            runProperties9.Append(color10);
            runProperties9.Append(fontSize10);
            runProperties9.Append(fontSizeComplexScript10);
            Text text9 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text9.Text = "2009 ";

            run9.Append(runProperties9);
            run9.Append(text9);

            Run run10 = new Run();

            RunProperties runProperties10 = new RunProperties();
            RunFonts runFonts11 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Color color11 = new Color() { Val = "auto" };
            FontSize fontSize11 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript11 = new FontSizeComplexScript() { Val = "24" };

            runProperties10.Append(runFonts11);
            runProperties10.Append(color11);
            runProperties10.Append(fontSize11);
            runProperties10.Append(fontSizeComplexScript11);
            Text text10 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text10.Text = "imposing an ";

            run10.Append(runProperties10);
            run10.Append(text10);
            ProofError proofError1 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run11 = new Run();

            RunProperties runProperties11 = new RunProperties();
            RunFonts runFonts12 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Color color12 = new Color() { Val = "auto" };
            FontSize fontSize12 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript12 = new FontSizeComplexScript() { Val = "24" };

            runProperties11.Append(runFonts12);
            runProperties11.Append(color12);
            runProperties11.Append(fontSize12);
            runProperties11.Append(fontSizeComplexScript12);
            Text text11 = new Text();
            text11.Text = "obligatrion";

            run11.Append(runProperties11);
            run11.Append(text11);
            ProofError proofError2 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            Run run12 = new Run();

            RunProperties runProperties12 = new RunProperties();
            RunFonts runFonts13 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Color color13 = new Color() { Val = "auto" };
            FontSize fontSize13 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript13 = new FontSizeComplexScript() { Val = "24" };

            runProperties12.Append(runFonts13);
            runProperties12.Append(color13);
            runProperties12.Append(fontSize13);
            runProperties12.Append(fontSizeComplexScript13);
            Text text12 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text12.Text = " on member states to maintain minimum stocks of crude oil and/or petroleum products ";

            run12.Append(runProperties12);
            run12.Append(text12);

            Run run13 = new Run() { RsidRunProperties = "00E075BD" };

            RunProperties runProperties13 = new RunProperties();
            RunFonts runFonts14 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Color color14 = new Color() { Val = "auto" };
            FontSize fontSize14 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript14 = new FontSizeComplexScript() { Val = "24" };

            runProperties13.Append(runFonts14);
            runProperties13.Append(color14);
            runProperties13.Append(fontSize14);
            runProperties13.Append(fontSizeComplexScript14);
            Text text13 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text13.Text = "– ";

            run13.Append(runProperties13);
            run13.Append(text13);
            CommentRangeStart commentRangeStart1 = new CommentRangeStart() { Id = "0" };

            Run run14 = new Run() { RsidRunProperties = "00E075BD" };

            RunProperties runProperties14 = new RunProperties();
            RunFonts runFonts15 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Color color15 = new Color() { Val = "auto" };
            FontSize fontSize15 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript15 = new FontSizeComplexScript() { Val = "24" };

            runProperties14.Append(runFonts15);
            runProperties14.Append(color15);
            runProperties14.Append(fontSize15);
            runProperties14.Append(fontSizeComplexScript15);
            Text text14 = new Text();
            text14.Text = "2";

            run14.Append(runProperties14);
            run14.Append(text14);

            Run run15 = new Run();

            RunProperties runProperties15 = new RunProperties();
            RunFonts runFonts16 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Color color16 = new Color() { Val = "auto" };
            FontSize fontSize16 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript16 = new FontSizeComplexScript() { Val = "24" };

            runProperties15.Append(runFonts16);
            runProperties15.Append(color16);
            runProperties15.Append(fontSize16);
            runProperties15.Append(fontSizeComplexScript16);
            Text text15 = new Text();
            text15.Text = "01";

            run15.Append(runProperties15);
            run15.Append(text15);

            Run run16 = new Run() { RsidRunProperties = "00E075BD" };

            RunProperties runProperties16 = new RunProperties();
            RunFonts runFonts17 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Color color17 = new Color() { Val = "auto" };
            FontSize fontSize17 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript17 = new FontSizeComplexScript() { Val = "24" };

            runProperties16.Append(runFonts17);
            runProperties16.Append(color17);
            runProperties16.Append(fontSize17);
            runProperties16.Append(fontSizeComplexScript17);
            Text text16 = new Text();
            text16.Text = "5/";

            run16.Append(runProperties16);
            run16.Append(text16);

            Run run17 = new Run();

            RunProperties runProperties17 = new RunProperties();
            RunFonts runFonts18 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Color color18 = new Color() { Val = "auto" };
            FontSize fontSize18 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript18 = new FontSizeComplexScript() { Val = "24" };

            runProperties17.Append(runFonts18);
            runProperties17.Append(color18);
            runProperties17.Append(fontSize18);
            runProperties17.Append(fontSizeComplexScript18);
            Text text17 = new Text();
            text17.Text = "2072";

            run17.Append(runProperties17);
            run17.Append(text17);
            CommentRangeEnd commentRangeEnd1 = new CommentRangeEnd() { Id = "0" };

            Run run18 = new Run();

            RunProperties runProperties18 = new RunProperties();
            RunStyle runStyle1 = new RunStyle() { Val = "CommentReference" };
            RunFonts runFonts19 = new RunFonts() { AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi, EastAsiaTheme = ThemeFontValues.MinorHighAnsi, ComplexScriptTheme = ThemeFontValues.MinorBidi };
            Bold bold1 = new Bold() { Val = false };
            BoldComplexScript boldComplexScript1 = new BoldComplexScript() { Val = false };
            Color color19 = new Color() { Val = "auto" };

            runProperties18.Append(runStyle1);
            runProperties18.Append(runFonts19);
            runProperties18.Append(bold1);
            runProperties18.Append(boldComplexScript1);
            runProperties18.Append(color19);
            CommentReference commentReference1 = new CommentReference() { Id = "0" };

            run18.Append(runProperties18);
            run18.Append(commentReference1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);
            paragraph1.Append(run2);
            paragraph1.Append(run3);
            paragraph1.Append(run4);
            paragraph1.Append(run5);
            paragraph1.Append(run6);
            paragraph1.Append(run7);
            paragraph1.Append(run8);
            paragraph1.Append(run9);
            paragraph1.Append(run10);
            paragraph1.Append(proofError1);
            paragraph1.Append(run11);
            paragraph1.Append(proofError2);
            paragraph1.Append(run12);
            paragraph1.Append(run13);
            paragraph1.Append(commentRangeStart1);
            paragraph1.Append(run14);
            paragraph1.Append(run15);
            paragraph1.Append(run16);
            paragraph1.Append(run17);
            paragraph1.Append(commentRangeEnd1);
            paragraph1.Append(run18);

            Paragraph paragraph2 = new Paragraph() { RsidParagraphMarkRevision = "00E075BD", RsidParagraphAddition = "005F6638", RsidParagraphProperties = "005F6638", RsidRunAdditionDefault = "005F6638", ParagraphId = "19CE7305", TextId = "77777777" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines() { Before = "240", After = "120", Line = "240", LineRule = LineSpacingRuleValues.Auto };

            ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
            RunFonts runFonts20 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Arial,Times New Roman", ComplexScript = "Arial" };
            Bold bold2 = new Bold();
            BoldComplexScript boldComplexScript2 = new BoldComplexScript();
            Languages languages1 = new Languages() { EastAsia = "en-GB" };

            paragraphMarkRunProperties2.Append(runFonts20);
            paragraphMarkRunProperties2.Append(bold2);
            paragraphMarkRunProperties2.Append(boldComplexScript2);
            paragraphMarkRunProperties2.Append(languages1);

            paragraphProperties2.Append(spacingBetweenLines2);
            paragraphProperties2.Append(paragraphMarkRunProperties2);

            Run run19 = new Run() { RsidRunProperties = "00E075BD" };

            RunProperties runProperties19 = new RunProperties();
            RunFonts runFonts21 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Arial,Times New Roman", ComplexScript = "Arial" };
            Bold bold3 = new Bold();
            BoldComplexScript boldComplexScript3 = new BoldComplexScript();
            Languages languages2 = new Languages() { EastAsia = "en-GB" };

            runProperties19.Append(runFonts21);
            runProperties19.Append(bold3);
            runProperties19.Append(boldComplexScript3);
            runProperties19.Append(languages2);
            Text text18 = new Text();
            text18.Text = "Incriminated Fact";

            run19.Append(runProperties19);
            run19.Append(text18);
            CommentRangeStart commentRangeStart2 = new CommentRangeStart() { Id = "1" };

            Run run20 = new Run() { RsidRunProperties = "00E075BD" };

            RunProperties runProperties20 = new RunProperties();
            RunFonts runFonts22 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Arial,Times New Roman", ComplexScript = "Arial" };
            Bold bold4 = new Bold();
            BoldComplexScript boldComplexScript4 = new BoldComplexScript();
            Languages languages3 = new Languages() { EastAsia = "en-GB" };

            runProperties20.Append(runFonts22);
            runProperties20.Append(bold4);
            runProperties20.Append(boldComplexScript4);
            runProperties20.Append(languages3);
            Text text19 = new Text();
            text19.Text = ":";

            run20.Append(runProperties20);
            run20.Append(text19);
            CommentRangeEnd commentRangeEnd2 = new CommentRangeEnd() { Id = "1" };

            Run run21 = new Run() { RsidRunProperties = "00E075BD" };

            RunProperties runProperties21 = new RunProperties();
            RunFonts runFonts23 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Arial,Times New Roman", ComplexScript = "Arial" };
            Bold bold5 = new Bold();
            BoldComplexScript boldComplexScript5 = new BoldComplexScript();
            Languages languages4 = new Languages() { EastAsia = "en-GB" };

            runProperties21.Append(runFonts23);
            runProperties21.Append(bold5);
            runProperties21.Append(boldComplexScript5);
            runProperties21.Append(languages4);
            CommentReference commentReference2 = new CommentReference() { Id = "1" };

            run21.Append(runProperties21);
            run21.Append(commentReference2);

            paragraph2.Append(paragraphProperties2);
            paragraph2.Append(run19);
            paragraph2.Append(commentRangeStart2);
            paragraph2.Append(run20);
            paragraph2.Append(commentRangeEnd2);
            paragraph2.Append(run21);

            Paragraph paragraph3 = new Paragraph() { RsidParagraphMarkRevision = "00E075BD", RsidParagraphAddition = "005F6638", RsidParagraphProperties = "005F6638", RsidRunAdditionDefault = "005F6638", ParagraphId = "201F5591", TextId = "77777777" };

            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines3 = new SpacingBetweenLines() { Before = "240", After = "120", Line = "240", LineRule = LineSpacingRuleValues.Auto };

            ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
            RunFonts runFonts24 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Arial,Times New Roman", ComplexScript = "Arial" };
            Bold bold6 = new Bold();
            BoldComplexScript boldComplexScript6 = new BoldComplexScript();
            Languages languages5 = new Languages() { EastAsia = "en-GB" };

            paragraphMarkRunProperties3.Append(runFonts24);
            paragraphMarkRunProperties3.Append(bold6);
            paragraphMarkRunProperties3.Append(boldComplexScript6);
            paragraphMarkRunProperties3.Append(languages5);

            paragraphProperties3.Append(spacingBetweenLines3);
            paragraphProperties3.Append(paragraphMarkRunProperties3);

            Run run22 = new Run();

            RunProperties runProperties22 = new RunProperties();
            Italic italic4 = new Italic();

            runProperties22.Append(italic4);
            Text text20 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text20.Text = "(C18) ";

            run22.Append(runProperties22);
            run22.Append(text20);

            Run run23 = new Run() { RsidRunProperties = "00E075BD" };

            RunProperties runProperties23 = new RunProperties();
            Italic italic5 = new Italic();

            runProperties23.Append(italic5);
            Text text21 = new Text();
            text21.Text = "Failure to adopt and/or communicate all measures for transposing Council Directive 2009/119/EC of 14 September";

            run23.Append(runProperties23);
            run23.Append(text21);

            Run run24 = new Run();

            RunProperties runProperties24 = new RunProperties();
            Italic italic6 = new Italic();

            runProperties24.Append(italic6);
            Text text22 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text22.Text = " 2009 imposing ";

            run24.Append(runProperties24);
            run24.Append(text22);

            Run run25 = new Run() { RsidRunProperties = "00E075BD" };

            RunProperties runProperties25 = new RunProperties();
            Italic italic7 = new Italic();

            runProperties25.Append(italic7);
            Text text23 = new Text();
            text23.Text = "an obligation on Member States to maintain minimum stocks of crude oil and/or petroleum products";

            run25.Append(runProperties25);
            run25.Append(text23);

            Run run26 = new Run() { RsidRunProperties = "00E075BD" };

            RunProperties runProperties26 = new RunProperties();
            Italic italic8 = new Italic();

            runProperties26.Append(italic8);
            CarriageReturn carriageReturn1 = new CarriageReturn();
            Text text24 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text24.Text = " ";

            run26.Append(runProperties26);
            run26.Append(carriageReturn1);
            run26.Append(text24);

            paragraph3.Append(paragraphProperties3);
            paragraph3.Append(run22);
            paragraph3.Append(run23);
            paragraph3.Append(run24);
            paragraph3.Append(run25);
            paragraph3.Append(run26);

            Paragraph paragraph4 = new Paragraph() { RsidParagraphAddition = "005F6638", RsidParagraphProperties = "005F6638", RsidRunAdditionDefault = "005F6638", ParagraphId = "6D2271D7", TextId = "77777777" };

            ParagraphProperties paragraphProperties4 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId2 = new ParagraphStyleId() { Val = "Briefingtext" };

            ParagraphMarkRunProperties paragraphMarkRunProperties4 = new ParagraphMarkRunProperties();
            Bold bold7 = new Bold();

            paragraphMarkRunProperties4.Append(bold7);

            paragraphProperties4.Append(paragraphStyleId2);
            paragraphProperties4.Append(paragraphMarkRunProperties4);

            paragraph4.Append(paragraphProperties4);

            Paragraph paragraph5 = new Paragraph() { RsidParagraphMarkRevision = "006515F2", RsidParagraphAddition = "005F6638", RsidParagraphProperties = "005F6638", RsidRunAdditionDefault = "005F6638", ParagraphId = "5D7640F6", TextId = "77777777" };

            ParagraphProperties paragraphProperties5 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId3 = new ParagraphStyleId() { Val = "Briefingtext" };

            ParagraphMarkRunProperties paragraphMarkRunProperties5 = new ParagraphMarkRunProperties();
            Bold bold8 = new Bold();

            paragraphMarkRunProperties5.Append(bold8);

            paragraphProperties5.Append(paragraphStyleId3);
            paragraphProperties5.Append(paragraphMarkRunProperties5);
            CommentRangeStart commentRangeStart3 = new CommentRangeStart() { Id = "2" };

            Run run27 = new Run();

            RunProperties runProperties27 = new RunProperties();
            Bold bold9 = new Bold();

            runProperties27.Append(bold9);
            Text text25 = new Text();
            text25.Text = "Next decision step in the cycle";

            run27.Append(runProperties27);
            run27.Append(text25);
            CommentRangeEnd commentRangeEnd3 = new CommentRangeEnd() { Id = "2" };

            Run run28 = new Run();

            RunProperties runProperties28 = new RunProperties();
            RunStyle runStyle2 = new RunStyle() { Val = "CommentReference" };
            RunFonts runFonts25 = new RunFonts() { AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi, ComplexScriptTheme = ThemeFontValues.MinorBidi };

            runProperties28.Append(runStyle2);
            runProperties28.Append(runFonts25);
            CommentReference commentReference3 = new CommentReference() { Id = "2" };

            run28.Append(runProperties28);
            run28.Append(commentReference3);

            Run run29 = new Run();

            RunProperties runProperties29 = new RunProperties();
            Bold bold10 = new Bold();

            runProperties29.Append(bold10);
            Text text26 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text26.Text = ": ";

            run29.Append(runProperties29);
            run29.Append(text26);

            Run run30 = new Run();

            RunProperties runProperties30 = new RunProperties();
            RunStyle runStyle3 = new RunStyle() { Val = "CommentReference" };
            RunFonts runFonts26 = new RunFonts() { AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi, ComplexScriptTheme = ThemeFontValues.MinorBidi };

            runProperties30.Append(runStyle3);
            runProperties30.Append(runFonts26);
            CommentReference commentReference4 = new CommentReference() { Id = "3" };

            run30.Append(runProperties30);
            run30.Append(commentReference4);

            paragraph5.Append(paragraphProperties5);
            paragraph5.Append(commentRangeStart3);
            paragraph5.Append(run27);
            paragraph5.Append(commentRangeEnd3);
            paragraph5.Append(run28);
            paragraph5.Append(run29);
            paragraph5.Append(run30);

            Paragraph paragraph6 = new Paragraph() { RsidParagraphMarkRevision = "00E075BD", RsidParagraphAddition = "005F6638", RsidParagraphProperties = "005F6638", RsidRunAdditionDefault = "005F6638", ParagraphId = "1CA700F3", TextId = "77777777" };

            ParagraphProperties paragraphProperties6 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines4 = new SpacingBetweenLines() { Before = "240", After = "120", Line = "240", LineRule = LineSpacingRuleValues.Auto };

            ParagraphMarkRunProperties paragraphMarkRunProperties6 = new ParagraphMarkRunProperties();
            RunFonts runFonts27 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSizeComplexScript fontSizeComplexScript19 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties6.Append(runFonts27);
            paragraphMarkRunProperties6.Append(fontSizeComplexScript19);

            paragraphProperties6.Append(spacingBetweenLines4);
            paragraphProperties6.Append(paragraphMarkRunProperties6);

            Run run31 = new Run() { RsidRunProperties = "00E075BD" };

            RunProperties runProperties31 = new RunProperties();
            RunFonts runFonts28 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSizeComplexScript fontSizeComplexScript20 = new FontSizeComplexScript() { Val = "24" };

            runProperties31.Append(runFonts28);
            runProperties31.Append(fontSizeComplexScript20);
            Text text27 = new Text();
            text27.Text = "Reasoned opinion 258";

            run31.Append(runProperties31);
            run31.Append(text27);

            Run run32 = new Run();

            RunProperties runProperties32 = new RunProperties();
            RunFonts runFonts29 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSizeComplexScript fontSizeComplexScript21 = new FontSizeComplexScript() { Val = "24" };

            runProperties32.Append(runFonts29);
            runProperties32.Append(fontSizeComplexScript21);
            Text text28 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text28.Text = " ";

            run32.Append(runProperties32);
            run32.Append(text28);

            Run run33 = new Run() { RsidRunProperties = "00E075BD" };

            RunProperties runProperties33 = new RunProperties();
            RunFonts runFonts30 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSizeComplexScript fontSizeComplexScript22 = new FontSizeComplexScript() { Val = "24" };

            runProperties33.Append(runFonts30);
            runProperties33.Append(fontSizeComplexScript22);
            Text text29 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text29.Text = "(ex226) - Notification ";

            run33.Append(runProperties33);
            run33.Append(text29);

            Run run34 = new Run();

            RunProperties runProperties34 = new RunProperties();
            RunFonts runFonts31 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSizeComplexScript fontSizeComplexScript23 = new FontSizeComplexScript() { Val = "24" };

            runProperties34.Append(runFonts31);
            runProperties34.Append(fontSizeComplexScript23);
            Text text30 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text30.Text = "on ";

            run34.Append(runProperties34);
            run34.Append(text30);

            Run run35 = new Run() { RsidRunProperties = "00E075BD" };

            RunProperties runProperties35 = new RunProperties();
            RunFonts runFonts32 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSizeComplexScript fontSizeComplexScript24 = new FontSizeComplexScript() { Val = "24" };

            runProperties35.Append(runFonts32);
            runProperties35.Append(fontSizeComplexScript24);
            Text text31 = new Text();
            text31.Text = "23/10/2015";

            run35.Append(runProperties35);
            run35.Append(text31);

            paragraph6.Append(paragraphProperties6);
            paragraph6.Append(run31);
            paragraph6.Append(run32);
            paragraph6.Append(run33);
            paragraph6.Append(run34);
            paragraph6.Append(run35);

            Paragraph paragraph7 = new Paragraph() { RsidParagraphAddition = "005F6638", RsidParagraphProperties = "005F6638", RsidRunAdditionDefault = "005F6638", ParagraphId = "4DDC2D1D", TextId = "77777777" };

            ParagraphProperties paragraphProperties7 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines5 = new SpacingBetweenLines() { Before = "240", After = "120", Line = "240", LineRule = LineSpacingRuleValues.Auto };

            ParagraphMarkRunProperties paragraphMarkRunProperties7 = new ParagraphMarkRunProperties();
            RunFonts runFonts33 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Arial,Times New Roman", ComplexScript = "Arial" };
            Bold bold11 = new Bold();
            BoldComplexScript boldComplexScript7 = new BoldComplexScript();
            Languages languages6 = new Languages() { EastAsia = "en-GB" };

            paragraphMarkRunProperties7.Append(runFonts33);
            paragraphMarkRunProperties7.Append(bold11);
            paragraphMarkRunProperties7.Append(boldComplexScript7);
            paragraphMarkRunProperties7.Append(languages6);

            paragraphProperties7.Append(spacingBetweenLines5);
            paragraphProperties7.Append(paragraphMarkRunProperties7);

            paragraph7.Append(paragraphProperties7);

            Paragraph paragraph8 = new Paragraph() { RsidParagraphMarkRevision = "003602E0", RsidParagraphAddition = "005F6638", RsidParagraphProperties = "005F6638", RsidRunAdditionDefault = "005F6638", ParagraphId = "12BB8C47", TextId = "77777777" };

            ParagraphProperties paragraphProperties8 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines6 = new SpacingBetweenLines() { Before = "240", After = "120", Line = "240", LineRule = LineSpacingRuleValues.Auto };

            ParagraphMarkRunProperties paragraphMarkRunProperties8 = new ParagraphMarkRunProperties();
            RunFonts runFonts34 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Times New Roman", ComplexScript = "Arial" };
            Bold bold12 = new Bold();
            FontSizeComplexScript fontSizeComplexScript25 = new FontSizeComplexScript() { Val = "24" };
            Languages languages7 = new Languages() { EastAsia = "en-GB" };

            paragraphMarkRunProperties8.Append(runFonts34);
            paragraphMarkRunProperties8.Append(bold12);
            paragraphMarkRunProperties8.Append(fontSizeComplexScript25);
            paragraphMarkRunProperties8.Append(languages7);

            paragraphProperties8.Append(spacingBetweenLines6);
            paragraphProperties8.Append(paragraphMarkRunProperties8);
            CommentRangeStart commentRangeStart4 = new CommentRangeStart() { Id = "4" };

            Run run36 = new Run() { RsidRunProperties = "003602E0" };

            RunProperties runProperties36 = new RunProperties();
            RunFonts runFonts35 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Arial,Times New Roman", ComplexScript = "Arial" };
            Bold bold13 = new Bold();
            BoldComplexScript boldComplexScript8 = new BoldComplexScript();
            Languages languages8 = new Languages() { EastAsia = "en-GB" };

            runProperties36.Append(runFonts35);
            runProperties36.Append(bold13);
            runProperties36.Append(boldComplexScript8);
            runProperties36.Append(languages8);
            Text text32 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text32.Text = "Last formal step ";

            run36.Append(runProperties36);
            run36.Append(text32);
            CommentRangeEnd commentRangeEnd4 = new CommentRangeEnd() { Id = "4" };

            Run run37 = new Run();

            RunProperties runProperties37 = new RunProperties();
            RunStyle runStyle4 = new RunStyle() { Val = "CommentReference" };

            runProperties37.Append(runStyle4);
            CommentReference commentReference5 = new CommentReference() { Id = "4" };

            run37.Append(runProperties37);
            run37.Append(commentReference5);

            paragraph8.Append(paragraphProperties8);
            paragraph8.Append(commentRangeStart4);
            paragraph8.Append(run36);
            paragraph8.Append(commentRangeEnd4);
            paragraph8.Append(run37);

            Paragraph paragraph9 = new Paragraph() { RsidParagraphAddition = "005F6638", RsidParagraphProperties = "005F6638", RsidRunAdditionDefault = "005F6638", ParagraphId = "400131FF", TextId = "77777777" };

            ParagraphProperties paragraphProperties9 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId4 = new ParagraphStyleId() { Val = "Briefingtext" };

            paragraphProperties9.Append(paragraphStyleId4);

            Run run38 = new Run() { RsidRunProperties = "00E075BD" };
            Text text33 = new Text();
            text33.Text = "Formal notice Art. 258 TFEU";

            run38.Append(text33);

            Run run39 = new Run();
            Text text34 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text34.Text = " ";

            run39.Append(text34);

            Run run40 = new Run() { RsidRunProperties = "00E075BD" };
            Text text35 = new Text();
            text35.Text = "on 18/06/2015";

            run40.Append(text35);

            paragraph9.Append(paragraphProperties9);
            paragraph9.Append(run38);
            paragraph9.Append(run39);
            paragraph9.Append(run40);

            Paragraph paragraph10 = new Paragraph() { RsidParagraphAddition = "005F6638", RsidParagraphProperties = "005F6638", RsidRunAdditionDefault = "005F6638", ParagraphId = "08E7D1F1", TextId = "77777777" };

            ParagraphProperties paragraphProperties10 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId5 = new ParagraphStyleId() { Val = "Briefingtext" };

            ParagraphMarkRunProperties paragraphMarkRunProperties9 = new ParagraphMarkRunProperties();
            RunFonts runFonts36 = new RunFonts() { EastAsia = "Arial,Times New Roman" };
            Bold bold14 = new Bold();
            BoldComplexScript boldComplexScript9 = new BoldComplexScript();
            Languages languages9 = new Languages() { EastAsia = "en-GB" };

            paragraphMarkRunProperties9.Append(runFonts36);
            paragraphMarkRunProperties9.Append(bold14);
            paragraphMarkRunProperties9.Append(boldComplexScript9);
            paragraphMarkRunProperties9.Append(languages9);

            paragraphProperties10.Append(paragraphStyleId5);
            paragraphProperties10.Append(paragraphMarkRunProperties9);

            paragraph10.Append(paragraphProperties10);

            Paragraph paragraph11 = new Paragraph() { RsidParagraphAddition = "005F6638", RsidParagraphProperties = "005F6638", RsidRunAdditionDefault = "005F6638", ParagraphId = "406DB6E9", TextId = "77777777" };

            ParagraphProperties paragraphProperties11 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId6 = new ParagraphStyleId() { Val = "Briefingtext" };

            ParagraphMarkRunProperties paragraphMarkRunProperties10 = new ParagraphMarkRunProperties();
            Bold bold15 = new Bold();

            paragraphMarkRunProperties10.Append(bold15);

            paragraphProperties11.Append(paragraphStyleId6);
            paragraphProperties11.Append(paragraphMarkRunProperties10);
            CommentRangeStart commentRangeStart5 = new CommentRangeStart() { Id = "5" };

            Run run41 = new Run() { RsidRunProperties = "006515F2" };

            RunProperties runProperties38 = new RunProperties();
            Bold bold16 = new Bold();

            runProperties38.Append(bold16);
            Text text36 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text36.Text = "History of ";

            run41.Append(runProperties38);
            run41.Append(text36);

            Run run42 = new Run();

            RunProperties runProperties39 = new RunProperties();
            Bold bold17 = new Bold();

            runProperties39.Append(bold17);
            Text text37 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text37.Text = "previous ";

            run42.Append(runProperties39);
            run42.Append(text37);

            Run run43 = new Run() { RsidRunProperties = "006515F2" };

            RunProperties runProperties40 = new RunProperties();
            Bold bold18 = new Bold();

            runProperties40.Append(bold18);
            Text text38 = new Text();
            text38.Text = "formal decisions";

            run43.Append(runProperties40);
            run43.Append(text38);
            CommentRangeEnd commentRangeEnd5 = new CommentRangeEnd() { Id = "5" };

            Run run44 = new Run();

            RunProperties runProperties41 = new RunProperties();
            RunStyle runStyle5 = new RunStyle() { Val = "CommentReference" };
            RunFonts runFonts37 = new RunFonts() { AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi, ComplexScriptTheme = ThemeFontValues.MinorBidi };

            runProperties41.Append(runStyle5);
            runProperties41.Append(runFonts37);
            CommentReference commentReference6 = new CommentReference() { Id = "5" };

            run44.Append(runProperties41);
            run44.Append(commentReference6);

            paragraph11.Append(paragraphProperties11);
            paragraph11.Append(commentRangeStart5);
            paragraph11.Append(run41);
            paragraph11.Append(run42);
            paragraph11.Append(run43);
            paragraph11.Append(commentRangeEnd5);
            paragraph11.Append(run44);

            Paragraph paragraph12 = new Paragraph() { RsidParagraphMarkRevision = "00E075BD", RsidParagraphAddition = "005F6638", RsidParagraphProperties = "005F6638", RsidRunAdditionDefault = "005F6638", ParagraphId = "4C0ED080", TextId = "77777777" };

            ParagraphProperties paragraphProperties12 = new ParagraphProperties();
            Justification justification1 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties11 = new ParagraphMarkRunProperties();
            RunFonts runFonts38 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSizeComplexScript fontSizeComplexScript26 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties11.Append(runFonts38);
            paragraphMarkRunProperties11.Append(fontSizeComplexScript26);

            paragraphProperties12.Append(justification1);
            paragraphProperties12.Append(paragraphMarkRunProperties11);

            Run run45 = new Run() { RsidRunProperties = "00E075BD" };

            RunProperties runProperties42 = new RunProperties();
            RunFonts runFonts39 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSizeComplexScript fontSizeComplexScript27 = new FontSizeComplexScript() { Val = "24" };

            runProperties42.Append(runFonts39);
            runProperties42.Append(fontSizeComplexScript27);
            Text text39 = new Text();
            text39.Text = "Formal notice Art. 258 TFEU";

            run45.Append(runProperties42);
            run45.Append(text39);

            Run run46 = new Run();

            RunProperties runProperties43 = new RunProperties();
            RunFonts runFonts40 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSizeComplexScript fontSizeComplexScript28 = new FontSizeComplexScript() { Val = "24" };

            runProperties43.Append(runFonts40);
            runProperties43.Append(fontSizeComplexScript28);
            Text text40 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text40.Text = " ";

            run46.Append(runProperties43);
            run46.Append(text40);

            Run run47 = new Run() { RsidRunProperties = "00E075BD" };

            RunProperties runProperties44 = new RunProperties();
            RunFonts runFonts41 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSizeComplexScript fontSizeComplexScript29 = new FontSizeComplexScript() { Val = "24" };

            runProperties44.Append(runFonts41);
            runProperties44.Append(fontSizeComplexScript29);
            Text text41 = new Text();
            text41.Text = "on 18/06/2015";

            run47.Append(runProperties44);
            run47.Append(text41);

            paragraph12.Append(paragraphProperties12);
            paragraph12.Append(run45);
            paragraph12.Append(run46);
            paragraph12.Append(run47);

            Paragraph paragraph13 = new Paragraph() { RsidParagraphAddition = "005F6638", RsidParagraphProperties = "005F6638", RsidRunAdditionDefault = "005F6638", ParagraphId = "6F619FC1", TextId = "77777777" };

            ParagraphProperties paragraphProperties13 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines7 = new SpacingBetweenLines() { Before = "240", After = "120", Line = "240", LineRule = LineSpacingRuleValues.Auto };

            ParagraphMarkRunProperties paragraphMarkRunProperties12 = new ParagraphMarkRunProperties();
            RunFonts runFonts42 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Arial,Times New Roman", ComplexScript = "Arial" };
            Bold bold19 = new Bold();
            BoldComplexScript boldComplexScript10 = new BoldComplexScript();
            Languages languages10 = new Languages() { EastAsia = "en-GB" };

            paragraphMarkRunProperties12.Append(runFonts42);
            paragraphMarkRunProperties12.Append(bold19);
            paragraphMarkRunProperties12.Append(boldComplexScript10);
            paragraphMarkRunProperties12.Append(languages10);

            paragraphProperties13.Append(spacingBetweenLines7);
            paragraphProperties13.Append(paragraphMarkRunProperties12);

            paragraph13.Append(paragraphProperties13);

            Paragraph paragraph14 = new Paragraph() { RsidParagraphAddition = "005F6638", RsidParagraphProperties = "005F6638", RsidRunAdditionDefault = "005F6638", ParagraphId = "0F3AA67A", TextId = "77777777" };

            ParagraphProperties paragraphProperties14 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines8 = new SpacingBetweenLines() { Before = "240", After = "120", Line = "240", LineRule = LineSpacingRuleValues.Auto };

            ParagraphMarkRunProperties paragraphMarkRunProperties13 = new ParagraphMarkRunProperties();
            RunFonts runFonts43 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Arial,Times New Roman", ComplexScript = "Arial" };
            Bold bold20 = new Bold();
            BoldComplexScript boldComplexScript11 = new BoldComplexScript();
            Languages languages11 = new Languages() { EastAsia = "en-GB" };

            paragraphMarkRunProperties13.Append(runFonts43);
            paragraphMarkRunProperties13.Append(bold20);
            paragraphMarkRunProperties13.Append(boldComplexScript11);
            paragraphMarkRunProperties13.Append(languages11);

            paragraphProperties14.Append(spacingBetweenLines8);
            paragraphProperties14.Append(paragraphMarkRunProperties13);
            CommentRangeStart commentRangeStart6 = new CommentRangeStart() { Id = "6" };

            Run run48 = new Run() { RsidRunProperties = "003602E0" };

            RunProperties runProperties45 = new RunProperties();
            RunFonts runFonts44 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Arial,Times New Roman", ComplexScript = "Arial" };
            Bold bold21 = new Bold();
            BoldComplexScript boldComplexScript12 = new BoldComplexScript();
            Languages languages12 = new Languages() { EastAsia = "en-GB" };

            runProperties45.Append(runFonts44);
            runProperties45.Append(bold21);
            runProperties45.Append(boldComplexScript12);
            runProperties45.Append(languages12);
            Text text42 = new Text();
            text42.Text = "Context";

            run48.Append(runProperties45);
            run48.Append(text42);
            CommentRangeEnd commentRangeEnd6 = new CommentRangeEnd() { Id = "6" };

            Run run49 = new Run();

            RunProperties runProperties46 = new RunProperties();
            RunStyle runStyle6 = new RunStyle() { Val = "CommentReference" };

            runProperties46.Append(runStyle6);
            CommentReference commentReference7 = new CommentReference() { Id = "6" };

            run49.Append(runProperties46);
            run49.Append(commentReference7);

            Run run50 = new Run();

            RunProperties runProperties47 = new RunProperties();
            RunFonts runFonts45 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Arial,Times New Roman", ComplexScript = "Arial" };
            Bold bold22 = new Bold();
            BoldComplexScript boldComplexScript13 = new BoldComplexScript();
            Languages languages13 = new Languages() { EastAsia = "en-GB" };

            runProperties47.Append(runFonts45);
            runProperties47.Append(bold22);
            runProperties47.Append(boldComplexScript13);
            runProperties47.Append(languages13);
            Text text43 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text43.Text = " ";

            run50.Append(runProperties47);
            run50.Append(text43);

            paragraph14.Append(paragraphProperties14);
            paragraph14.Append(commentRangeStart6);
            paragraph14.Append(run48);
            paragraph14.Append(commentRangeEnd6);
            paragraph14.Append(run49);
            paragraph14.Append(run50);

            Paragraph paragraph15 = new Paragraph() { RsidParagraphMarkRevision = "00C73951", RsidParagraphAddition = "005F6638", RsidParagraphProperties = "005F6638", RsidRunAdditionDefault = "005F6638", ParagraphId = "5811576F", TextId = "77777777" };

            ParagraphProperties paragraphProperties15 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines9 = new SpacingBetweenLines() { Before = "240", After = "120", Line = "240", LineRule = LineSpacingRuleValues.Auto };

            ParagraphMarkRunProperties paragraphMarkRunProperties14 = new ParagraphMarkRunProperties();
            RunFonts runFonts46 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSizeComplexScript fontSizeComplexScript30 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties14.Append(runFonts46);
            paragraphMarkRunProperties14.Append(fontSizeComplexScript30);

            paragraphProperties15.Append(spacingBetweenLines9);
            paragraphProperties15.Append(paragraphMarkRunProperties14);

            Run run51 = new Run() { RsidRunProperties = "00C73951" };

            RunProperties runProperties48 = new RunProperties();
            RunFonts runFonts47 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSizeComplexScript fontSizeComplexScript31 = new FontSizeComplexScript() { Val = "24" };

            runProperties48.Append(runFonts47);
            runProperties48.Append(fontSizeComplexScript31);
            Text text44 = new Text();
            text44.Text = "19/04/2006";

            run51.Append(runProperties48);
            run51.Append(text44);

            paragraph15.Append(paragraphProperties15);
            paragraph15.Append(run51);

            Paragraph paragraph16 = new Paragraph() { RsidParagraphAddition = "005F6638", RsidParagraphProperties = "005F6638", RsidRunAdditionDefault = "005F6638", ParagraphId = "2C54113F", TextId = "77777777" };

            ParagraphProperties paragraphProperties16 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId7 = new ParagraphStyleId() { Val = "Briefingtext" };

            paragraphProperties16.Append(paragraphStyleId7);
            ProofError proofError3 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

            Run run52 = new Run();
            Text text45 = new Text();
            text45.Text = "1.FACTS";

            run52.Append(text45);
            ProofError proofError4 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

            paragraph16.Append(paragraphProperties16);
            paragraph16.Append(proofError3);
            paragraph16.Append(run52);
            paragraph16.Append(proofError4);

            Paragraph paragraph17 = new Paragraph() { RsidParagraphMarkRevision = "00E075BD", RsidParagraphAddition = "005F6638", RsidParagraphProperties = "005F6638", RsidRunAdditionDefault = "005F6638", ParagraphId = "30266226", TextId = "77777777" };

            ParagraphProperties paragraphProperties17 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId8 = new ParagraphStyleId() { Val = "Briefingtext" };
            Indentation indentation1 = new Indentation() { Start = "720" };

            ParagraphMarkRunProperties paragraphMarkRunProperties15 = new ParagraphMarkRunProperties();
            FontSize fontSize19 = new FontSize() { Val = "20" };

            paragraphMarkRunProperties15.Append(fontSize19);

            paragraphProperties17.Append(paragraphStyleId8);
            paragraphProperties17.Append(indentation1);
            paragraphProperties17.Append(paragraphMarkRunProperties15);

            Run run53 = new Run() { RsidRunProperties = "00E075BD" };

            RunProperties runProperties49 = new RunProperties();
            FontSize fontSize20 = new FontSize() { Val = "20" };

            runProperties49.Append(fontSize20);
            Text text46 = new Text();
            text46.Text = "Failure to adopt and/or to notify the national measures (MNE) transposing Council Directive 2009/119/EC of 14 September 2009 imposing an obligation on Member States to maintain minimum stocks of crude oil and/or petroleum products (OJ L 265, 9.10.2009, p. 9).";

            run53.Append(runProperties49);
            run53.Append(text46);

            paragraph17.Append(paragraphProperties17);
            paragraph17.Append(run53);

            Paragraph paragraph18 = new Paragraph() { RsidParagraphMarkRevision = "00E075BD", RsidParagraphAddition = "005F6638", RsidParagraphProperties = "005F6638", RsidRunAdditionDefault = "005F6638", ParagraphId = "4B90117D", TextId = "77777777" };

            ParagraphProperties paragraphProperties18 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId9 = new ParagraphStyleId() { Val = "Briefingtext" };
            Indentation indentation2 = new Indentation() { Start = "720" };

            ParagraphMarkRunProperties paragraphMarkRunProperties16 = new ParagraphMarkRunProperties();
            FontSize fontSize21 = new FontSize() { Val = "20" };

            paragraphMarkRunProperties16.Append(fontSize21);

            paragraphProperties18.Append(paragraphStyleId9);
            paragraphProperties18.Append(indentation2);
            paragraphProperties18.Append(paragraphMarkRunProperties16);

            Run run54 = new Run() { RsidRunProperties = "00E075BD" };

            RunProperties runProperties50 = new RunProperties();
            FontSize fontSize22 = new FontSize() { Val = "20" };

            runProperties50.Append(fontSize22);
            Text text47 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text47.Text = "Article 25(1) of the Directive provides that ";

            run54.Append(runProperties50);
            run54.Append(text47);
            ProofError proofError5 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

            Run run55 = new Run() { RsidRunProperties = "00E075BD" };

            RunProperties runProperties51 = new RunProperties();
            FontSize fontSize23 = new FontSize() { Val = "20" };

            runProperties51.Append(fontSize23);
            Text text48 = new Text();
            text48.Text = "\" Member";

            run55.Append(runProperties51);
            run55.Append(text48);
            ProofError proofError6 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

            Run run56 = new Run() { RsidRunProperties = "00E075BD" };

            RunProperties runProperties52 = new RunProperties();
            FontSize fontSize24 = new FontSize() { Val = "20" };

            runProperties52.Append(fontSize24);
            Text text49 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text49.Text = " States shall bring into force the laws, regulations and administrative provisions necessary to comply with this Directive by 31 December 2012.\"";

            run56.Append(runProperties52);
            run56.Append(text49);

            paragraph18.Append(paragraphProperties18);
            paragraph18.Append(run54);
            paragraph18.Append(proofError5);
            paragraph18.Append(run55);
            paragraph18.Append(proofError6);
            paragraph18.Append(run56);

            Paragraph paragraph19 = new Paragraph() { RsidParagraphAddition = "005F6638", RsidParagraphProperties = "005F6638", RsidRunAdditionDefault = "005F6638", ParagraphId = "70FBBF66", TextId = "77777777" };

            ParagraphProperties paragraphProperties19 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId10 = new ParagraphStyleId() { Val = "Briefingtext" };
            Indentation indentation3 = new Indentation() { Start = "720" };

            paragraphProperties19.Append(paragraphStyleId10);
            paragraphProperties19.Append(indentation3);

            Run run57 = new Run() { RsidRunProperties = "00E075BD" };

            RunProperties runProperties53 = new RunProperties();
            FontSize fontSize25 = new FontSize() { Val = "20" };

            runProperties53.Append(fontSize25);
            Text text50 = new Text();
            text50.Text = "By 9/09/2015, Finland had notified several measures transposing the Directive. Finland declared full transposition by 31/12/2012.";

            run57.Append(runProperties53);
            run57.Append(text50);

            paragraph19.Append(paragraphProperties19);
            paragraph19.Append(run57);

            Paragraph paragraph20 = new Paragraph() { RsidParagraphAddition = "005F6638", RsidParagraphProperties = "005F6638", RsidRunAdditionDefault = "005F6638", ParagraphId = "04EAA7D7", TextId = "77777777" };

            ParagraphProperties paragraphProperties20 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId11 = new ParagraphStyleId() { Val = "Briefingtext" };

            paragraphProperties20.Append(paragraphStyleId11);

            Run run58 = new Run();
            Text text51 = new Text();
            text51.Text = "2. PROCEDURE";

            run58.Append(text51);

            paragraph20.Append(paragraphProperties20);
            paragraph20.Append(run58);

            Paragraph paragraph21 = new Paragraph() { RsidParagraphMarkRevision = "00E075BD", RsidParagraphAddition = "005F6638", RsidParagraphProperties = "005F6638", RsidRunAdditionDefault = "005F6638", ParagraphId = "39359948", TextId = "77777777" };

            ParagraphProperties paragraphProperties21 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId12 = new ParagraphStyleId() { Val = "Briefingtext" };
            Indentation indentation4 = new Indentation() { Start = "720" };

            ParagraphMarkRunProperties paragraphMarkRunProperties17 = new ParagraphMarkRunProperties();
            FontSize fontSize26 = new FontSize() { Val = "20" };

            paragraphMarkRunProperties17.Append(fontSize26);

            paragraphProperties21.Append(paragraphStyleId12);
            paragraphProperties21.Append(indentation4);
            paragraphProperties21.Append(paragraphMarkRunProperties17);

            Run run59 = new Run() { RsidRunProperties = "00E075BD" };

            RunProperties runProperties54 = new RunProperties();
            FontSize fontSize27 = new FontSize() { Val = "20" };

            runProperties54.Append(fontSize27);
            LastRenderedPageBreak lastRenderedPageBreak1 = new LastRenderedPageBreak();
            Text text52 = new Text();
            text52.Text = "18/1/2013 – notification MNEs (full transposition)";

            run59.Append(runProperties54);
            run59.Append(lastRenderedPageBreak1);
            run59.Append(text52);

            paragraph21.Append(paragraphProperties21);
            paragraph21.Append(run59);

            Paragraph paragraph22 = new Paragraph() { RsidParagraphMarkRevision = "00E075BD", RsidParagraphAddition = "005F6638", RsidParagraphProperties = "005F6638", RsidRunAdditionDefault = "005F6638", ParagraphId = "58461F25", TextId = "77777777" };

            ParagraphProperties paragraphProperties22 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId13 = new ParagraphStyleId() { Val = "Briefingtext" };
            Indentation indentation5 = new Indentation() { Start = "720" };

            ParagraphMarkRunProperties paragraphMarkRunProperties18 = new ParagraphMarkRunProperties();
            FontSize fontSize28 = new FontSize() { Val = "20" };

            paragraphMarkRunProperties18.Append(fontSize28);

            paragraphProperties22.Append(paragraphStyleId13);
            paragraphProperties22.Append(indentation5);
            paragraphProperties22.Append(paragraphMarkRunProperties18);

            Run run60 = new Run() { RsidRunProperties = "00E075BD" };

            RunProperties runProperties55 = new RunProperties();
            FontSize fontSize29 = new FontSize() { Val = "20" };

            runProperties55.Append(fontSize29);
            Text text53 = new Text();
            text53.Text = "19/06/2015 – LFN sent";

            run60.Append(runProperties55);
            run60.Append(text53);

            paragraph22.Append(paragraphProperties22);
            paragraph22.Append(run60);

            Paragraph paragraph23 = new Paragraph() { RsidParagraphMarkRevision = "00E075BD", RsidParagraphAddition = "005F6638", RsidParagraphProperties = "005F6638", RsidRunAdditionDefault = "005F6638", ParagraphId = "1525ACEE", TextId = "77777777" };

            ParagraphProperties paragraphProperties23 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId14 = new ParagraphStyleId() { Val = "Briefingtext" };
            Indentation indentation6 = new Indentation() { Start = "720" };

            ParagraphMarkRunProperties paragraphMarkRunProperties19 = new ParagraphMarkRunProperties();
            FontSize fontSize30 = new FontSize() { Val = "20" };

            paragraphMarkRunProperties19.Append(fontSize30);

            paragraphProperties23.Append(paragraphStyleId14);
            paragraphProperties23.Append(indentation6);
            paragraphProperties23.Append(paragraphMarkRunProperties19);

            Run run61 = new Run() { RsidRunProperties = "00E075BD" };

            RunProperties runProperties56 = new RunProperties();
            FontSize fontSize31 = new FontSize() { Val = "20" };

            runProperties56.Append(fontSize31);
            Text text54 = new Text();
            text54.Text = "17/08/2015 – reply Finland (full transposition)";

            run61.Append(runProperties56);
            run61.Append(text54);

            paragraph23.Append(paragraphProperties23);
            paragraph23.Append(run61);

            Paragraph paragraph24 = new Paragraph() { RsidParagraphAddition = "005F6638", RsidParagraphProperties = "005F6638", RsidRunAdditionDefault = "005F6638", ParagraphId = "6167DBCF", TextId = "77777777" };

            ParagraphProperties paragraphProperties24 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId15 = new ParagraphStyleId() { Val = "Briefingtext" };

            paragraphProperties24.Append(paragraphStyleId15);

            Run run62 = new Run();
            Text text55 = new Text();
            text55.Text = "3. ANALYSIS";

            run62.Append(text55);

            paragraph24.Append(paragraphProperties24);
            paragraph24.Append(run62);

            Paragraph paragraph25 = new Paragraph() { RsidParagraphMarkRevision = "00E075BD", RsidParagraphAddition = "005F6638", RsidParagraphProperties = "005F6638", RsidRunAdditionDefault = "005F6638", ParagraphId = "68687AA0", TextId = "77777777" };

            ParagraphProperties paragraphProperties25 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId16 = new ParagraphStyleId() { Val = "Briefingtext" };
            Indentation indentation7 = new Indentation() { Start = "720" };

            ParagraphMarkRunProperties paragraphMarkRunProperties20 = new ParagraphMarkRunProperties();
            FontSize fontSize32 = new FontSize() { Val = "20" };

            paragraphMarkRunProperties20.Append(fontSize32);

            paragraphProperties25.Append(paragraphStyleId16);
            paragraphProperties25.Append(indentation7);
            paragraphProperties25.Append(paragraphMarkRunProperties20);

            Run run63 = new Run() { RsidRunProperties = "00E075BD" };

            RunProperties runProperties57 = new RunProperties();
            FontSize fontSize33 = new FontSize() { Val = "20" };

            runProperties57.Append(fontSize33);
            Text text56 = new Text();
            text56.Text = "After assessment of the measures notified by Finland, the Commission services considered that the following provisions of Directive 2009/119/EC had not been transposed into the national legal order: Article 2(b) and (m), Article 5(1), first subparagraph,";

            run63.Append(runProperties57);
            run63.Append(text56);

            paragraph25.Append(paragraphProperties25);
            paragraph25.Append(run63);

            Paragraph paragraph26 = new Paragraph() { RsidParagraphMarkRevision = "00E075BD", RsidParagraphAddition = "005F6638", RsidParagraphProperties = "005F6638", RsidRunAdditionDefault = "005F6638", ParagraphId = "0B353195", TextId = "77777777" };

            ParagraphProperties paragraphProperties26 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId17 = new ParagraphStyleId() { Val = "Briefingtext" };
            Indentation indentation8 = new Indentation() { Start = "720" };

            ParagraphMarkRunProperties paragraphMarkRunProperties21 = new ParagraphMarkRunProperties();
            FontSize fontSize34 = new FontSize() { Val = "20" };

            paragraphMarkRunProperties21.Append(fontSize34);

            paragraphProperties26.Append(paragraphStyleId17);
            paragraphProperties26.Append(indentation8);
            paragraphProperties26.Append(paragraphMarkRunProperties21);

            Run run64 = new Run() { RsidRunProperties = "00E075BD" };

            RunProperties runProperties58 = new RunProperties();
            FontSize fontSize35 = new FontSize() { Val = "20" };

            runProperties58.Append(fontSize35);
            Text text57 = new Text();
            text57.Text = "Article 5(1), second subparagraph, first sentence, Article 5(2), Article 8(1), second subparagraph, Article 16(1) and (2), Annexes I and III. In their reply, the Finnish authorities partially addressed the issues raised in the letter of formal notice. However, the Commission services consider that the following provisions have not been transposed: Article 2(m), Article 5(1), first subparagraph, Article 5(1), second subparagraph, first sentence, Annexes I and III.";

            run64.Append(runProperties58);
            run64.Append(text57);

            paragraph26.Append(paragraphProperties26);
            paragraph26.Append(run64);

            Paragraph paragraph27 = new Paragraph() { RsidParagraphAddition = "005F6638", RsidParagraphProperties = "005F6638", RsidRunAdditionDefault = "005F6638", ParagraphId = "12BBD25C", TextId = "77777777" };

            ParagraphProperties paragraphProperties27 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId18 = new ParagraphStyleId() { Val = "Briefingtext" };

            paragraphProperties27.Append(paragraphStyleId18);

            Run run65 = new Run();
            Text text58 = new Text();
            text58.Text = "4. CONCLUSION/PROPOSAL OF THE SERVICE";

            run65.Append(text58);

            paragraph27.Append(paragraphProperties27);
            paragraph27.Append(run65);

            Paragraph paragraph28 = new Paragraph() { RsidParagraphMarkRevision = "00E075BD", RsidParagraphAddition = "005F6638", RsidParagraphProperties = "005F6638", RsidRunAdditionDefault = "005F6638", ParagraphId = "4A81D3DC", TextId = "77777777" };

            ParagraphProperties paragraphProperties28 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId19 = new ParagraphStyleId() { Val = "Briefingtext" };
            Indentation indentation9 = new Indentation() { Start = "720" };

            paragraphProperties28.Append(paragraphStyleId19);
            paragraphProperties28.Append(indentation9);

            Run run66 = new Run() { RsidRunProperties = "00E075BD" };

            RunProperties runProperties59 = new RunProperties();
            FontSize fontSize36 = new FontSize() { Val = "20" };

            runProperties59.Append(fontSize36);
            Text text59 = new Text();
            text59.Text = "In view of the above analysis, it is considered that Finland has not yet adopted and/or notified all the measures necessary to fully transpose Directive 2009/119/EC and therefore failed to fulfil its obligations under Art. 25(1) of this Directive. Consequently, it is proposed to issue a reasoned opinion in M-10/15.";

            run66.Append(runProperties59);
            run66.Append(text59);

            Run run67 = new Run() { RsidRunProperties = "00E075BD" };

            RunProperties runProperties60 = new RunProperties();
            FontSize fontSize37 = new FontSize() { Val = "20" };

            runProperties60.Append(fontSize37);
            CarriageReturn carriageReturn2 = new CarriageReturn();

            run67.Append(runProperties60);
            run67.Append(carriageReturn2);

            Run run68 = new Run() { RsidRunProperties = "003602E0" };
            Text text60 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text60.Text = ". ";

            run68.Append(text60);

            paragraph28.Append(paragraphProperties28);
            paragraph28.Append(run66);
            paragraph28.Append(run67);
            paragraph28.Append(run68);

            Paragraph paragraph29 = new Paragraph() { RsidParagraphAddition = "005F6638", RsidParagraphProperties = "005F6638", RsidRunAdditionDefault = "005F6638", ParagraphId = "71A980DA", TextId = "77777777" };

            ParagraphProperties paragraphProperties29 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines10 = new SpacingBetweenLines() { Before = "240", After = "120", Line = "240", LineRule = LineSpacingRuleValues.Auto };

            ParagraphMarkRunProperties paragraphMarkRunProperties22 = new ParagraphMarkRunProperties();
            RunFonts runFonts48 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Arial,Times New Roman", ComplexScript = "Arial" };
            Bold bold23 = new Bold();
            BoldComplexScript boldComplexScript14 = new BoldComplexScript();
            Languages languages14 = new Languages() { EastAsia = "en-GB" };

            paragraphMarkRunProperties22.Append(runFonts48);
            paragraphMarkRunProperties22.Append(bold23);
            paragraphMarkRunProperties22.Append(boldComplexScript14);
            paragraphMarkRunProperties22.Append(languages14);

            paragraphProperties29.Append(spacingBetweenLines10);
            paragraphProperties29.Append(paragraphMarkRunProperties22);

            paragraph29.Append(paragraphProperties29);

            Paragraph paragraph30 = new Paragraph() { RsidParagraphMarkRevision = "003602E0", RsidParagraphAddition = "005F6638", RsidParagraphProperties = "005F6638", RsidRunAdditionDefault = "005F6638", ParagraphId = "0EC1148C", TextId = "77777777" };

            ParagraphProperties paragraphProperties30 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId20 = new ParagraphStyleId() { Val = "Briefingcontact" };

            paragraphProperties30.Append(paragraphStyleId20);
            CommentRangeStart commentRangeStart7 = new CommentRangeStart() { Id = "7" };

            Run run69 = new Run() { RsidRunProperties = "003602E0" };
            Text text61 = new Text();
            text61.Text = "Contact";

            run69.Append(text61);
            CommentRangeEnd commentRangeEnd7 = new CommentRangeEnd() { Id = "7" };

            Run run70 = new Run();

            RunProperties runProperties61 = new RunProperties();
            RunStyle runStyle7 = new RunStyle() { Val = "CommentReference" };
            RunFonts runFonts49 = new RunFonts() { AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi, ComplexScriptTheme = ThemeFontValues.MinorBidi };
            Bold bold24 = new Bold() { Val = false };

            runProperties61.Append(runStyle7);
            runProperties61.Append(runFonts49);
            runProperties61.Append(bold24);
            CommentReference commentReference8 = new CommentReference() { Id = "7" };

            run70.Append(runProperties61);
            run70.Append(commentReference8);

            Run run71 = new Run() { RsidRunProperties = "003602E0" };
            Text text62 = new Text();
            text62.Text = ":";

            run71.Append(text62);

            Run run72 = new Run() { RsidRunProperties = "003602E0" };
            TabChar tabChar1 = new TabChar();

            run72.Append(tabChar1);
            CommentRangeStart commentRangeStart8 = new CommentRangeStart() { Id = "8" };

            Run run73 = new Run();
            Text text63 = new Text();
            text63.Text = "N/A";

            run73.Append(text63);
            CommentRangeEnd commentRangeEnd8 = new CommentRangeEnd() { Id = "8" };

            Run run74 = new Run();

            RunProperties runProperties62 = new RunProperties();
            RunStyle runStyle8 = new RunStyle() { Val = "CommentReference" };
            RunFonts runFonts50 = new RunFonts() { AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi, ComplexScriptTheme = ThemeFontValues.MinorBidi };
            Bold bold25 = new Bold() { Val = false };

            runProperties62.Append(runStyle8);
            runProperties62.Append(runFonts50);
            runProperties62.Append(bold25);
            CommentReference commentReference9 = new CommentReference() { Id = "8" };

            run74.Append(runProperties62);
            run74.Append(commentReference9);
            ProofError proofError7 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

            Run run75 = new Run();
            Text text64 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text64.Text = ", ";

            run75.Append(text64);

            Run run76 = new Run() { RsidRunProperties = "003602E0" };
            Text text65 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text65.Text = " (";

            run76.Append(text65);
            ProofError proofError8 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

            Run run77 = new Run() { RsidRunProperties = "003602E0" };
            Text text66 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text66.Text = "DG ";

            run77.Append(text66);

            Run run78 = new Run();
            Text text67 = new Text();
            text67.Text = "ENER)";

            run78.Append(text67);

            Run run79 = new Run() { RsidRunProperties = "003602E0" };
            Text text68 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text68.Text = ", ";

            run79.Append(text68);

            Run run80 = new Run();
            Text text69 = new Text();
            text69.Text = "N/A";

            run80.Append(text69);

            Run run81 = new Run() { RsidRunProperties = "003602E0" };
            Break break1 = new Break();
            Text text70 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text70.Text = "Updated on: ";

            run81.Append(break1);
            run81.Append(text70);

            Run run82 = new Run();
            Text text71 = new Text();
            text71.Text = "05/10/";

            run82.Append(text71);

            Run run83 = new Run() { RsidRunProperties = "003602E0" };
            Text text72 = new Text();
            text72.Text = "201";

            run83.Append(text72);

            Run run84 = new Run();
            Text text73 = new Text();
            text73.Text = "5";

            run84.Append(text73);

            paragraph30.Append(paragraphProperties30);
            paragraph30.Append(commentRangeStart7);
            paragraph30.Append(run69);
            paragraph30.Append(commentRangeEnd7);
            paragraph30.Append(run70);
            paragraph30.Append(run71);
            paragraph30.Append(run72);
            paragraph30.Append(commentRangeStart8);
            paragraph30.Append(run73);
            paragraph30.Append(commentRangeEnd8);
            paragraph30.Append(run74);
            paragraph30.Append(proofError7);
            paragraph30.Append(run75);
            paragraph30.Append(run76);
            paragraph30.Append(proofError8);
            paragraph30.Append(run77);
            paragraph30.Append(run78);
            paragraph30.Append(run79);
            paragraph30.Append(run80);
            paragraph30.Append(run81);
            paragraph30.Append(run82);
            paragraph30.Append(run83);
            paragraph30.Append(run84);

            Paragraph paragraph31 = new Paragraph() { RsidParagraphMarkRevision = "005F6638", RsidParagraphAddition = "00C33253", RsidRunAdditionDefault = "00C33253", ParagraphId = "203E0153", TextId = "77777777" };

            ParagraphProperties paragraphProperties31 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties23 = new ParagraphMarkRunProperties();
            Languages languages15 = new Languages() { Val = "en-US" };

            paragraphMarkRunProperties23.Append(languages15);

            paragraphProperties31.Append(paragraphMarkRunProperties23);
            BookmarkStart bookmarkStart1 = new BookmarkStart() { Name = "_GoBack", Id = "9" };
            BookmarkEnd bookmarkEnd1 = new BookmarkEnd() { Id = "9" };

            paragraph31.Append(paragraphProperties31);
            paragraph31.Append(bookmarkStart1);
            paragraph31.Append(bookmarkEnd1);

            SectionProperties sectionProperties1 = new SectionProperties() { RsidRPr = "005F6638", RsidR = "00C33253" };
            PageSize pageSize1 = new PageSize() { Width = (UInt32Value)12240U, Height = (UInt32Value)15840U };
            PageMargin pageMargin1 = new PageMargin() { Top = 1440, Right = (UInt32Value)1440U, Bottom = 1440, Left = (UInt32Value)1440U, Header = (UInt32Value)708U, Footer = (UInt32Value)708U, Gutter = (UInt32Value)0U };
            Columns columns1 = new Columns() { Space = "708" };
            DocGrid docGrid1 = new DocGrid() { LinePitch = 360 };

            sectionProperties1.Append(pageSize1);
            sectionProperties1.Append(pageMargin1);
            sectionProperties1.Append(columns1);
            sectionProperties1.Append(docGrid1);

            body1.Append(paragraph1);
            body1.Append(paragraph2);
            body1.Append(paragraph3);
            body1.Append(paragraph4);
            body1.Append(paragraph5);
            body1.Append(paragraph6);
            body1.Append(paragraph7);
            body1.Append(paragraph8);
            body1.Append(paragraph9);
            body1.Append(paragraph10);
            body1.Append(paragraph11);
            body1.Append(paragraph12);
            body1.Append(paragraph13);
            body1.Append(paragraph14);
            body1.Append(paragraph15);
            body1.Append(paragraph16);
            body1.Append(paragraph17);
            body1.Append(paragraph18);
            body1.Append(paragraph19);
            body1.Append(paragraph20);
            body1.Append(paragraph21);
            body1.Append(paragraph22);
            body1.Append(paragraph23);
            body1.Append(paragraph24);
            body1.Append(paragraph25);
            body1.Append(paragraph26);
            body1.Append(paragraph27);
            body1.Append(paragraph28);
            body1.Append(paragraph29);
            body1.Append(paragraph30);
            body1.Append(paragraph31);
            body1.Append(sectionProperties1);

            document1.Append(body1);

            mainDocumentPart1.Document = document1;
        }

        // Generates content of themePart1.
        private void GenerateThemePart1Content(ThemePart themePart1)
        {
            A.Theme theme1 = new A.Theme() { Name = "Office Theme" };
            theme1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.ThemeElements themeElements1 = new A.ThemeElements();

            A.ColorScheme colorScheme1 = new A.ColorScheme() { Name = "Office" };

            A.Dark1Color dark1Color1 = new A.Dark1Color();
            A.SystemColor systemColor1 = new A.SystemColor() { Val = A.SystemColorValues.WindowText, LastColor = "000000" };

            dark1Color1.Append(systemColor1);

            A.Light1Color light1Color1 = new A.Light1Color();
            A.SystemColor systemColor2 = new A.SystemColor() { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" };

            light1Color1.Append(systemColor2);

            A.Dark2Color dark2Color1 = new A.Dark2Color();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "44546A" };

            dark2Color1.Append(rgbColorModelHex1);

            A.Light2Color light2Color1 = new A.Light2Color();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "E7E6E6" };

            light2Color1.Append(rgbColorModelHex2);

            A.Accent1Color accent1Color1 = new A.Accent1Color();
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "5B9BD5" };

            accent1Color1.Append(rgbColorModelHex3);

            A.Accent2Color accent2Color1 = new A.Accent2Color();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "ED7D31" };

            accent2Color1.Append(rgbColorModelHex4);

            A.Accent3Color accent3Color1 = new A.Accent3Color();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "A5A5A5" };

            accent3Color1.Append(rgbColorModelHex5);

            A.Accent4Color accent4Color1 = new A.Accent4Color();
            A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex() { Val = "FFC000" };

            accent4Color1.Append(rgbColorModelHex6);

            A.Accent5Color accent5Color1 = new A.Accent5Color();
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex() { Val = "4472C4" };

            accent5Color1.Append(rgbColorModelHex7);

            A.Accent6Color accent6Color1 = new A.Accent6Color();
            A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex() { Val = "70AD47" };

            accent6Color1.Append(rgbColorModelHex8);

            A.Hyperlink hyperlink1 = new A.Hyperlink();
            A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex() { Val = "0563C1" };

            hyperlink1.Append(rgbColorModelHex9);

            A.FollowedHyperlinkColor followedHyperlinkColor1 = new A.FollowedHyperlinkColor();
            A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex() { Val = "954F72" };

            followedHyperlinkColor1.Append(rgbColorModelHex10);

            colorScheme1.Append(dark1Color1);
            colorScheme1.Append(light1Color1);
            colorScheme1.Append(dark2Color1);
            colorScheme1.Append(light2Color1);
            colorScheme1.Append(accent1Color1);
            colorScheme1.Append(accent2Color1);
            colorScheme1.Append(accent3Color1);
            colorScheme1.Append(accent4Color1);
            colorScheme1.Append(accent5Color1);
            colorScheme1.Append(accent6Color1);
            colorScheme1.Append(hyperlink1);
            colorScheme1.Append(followedHyperlinkColor1);

            A.FontScheme fontScheme1 = new A.FontScheme() { Name = "Office" };

            A.MajorFont majorFont1 = new A.MajorFont();
            A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "Calibri Light", Panose = "020F0302020204030204" };
            A.EastAsianFont eastAsianFont1 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont1 = new A.SupplementalFont() { Script = "Jpan", Typeface = "游ゴシック Light" };
            A.SupplementalFont supplementalFont2 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont3 = new A.SupplementalFont() { Script = "Hans", Typeface = "等线 Light" };
            A.SupplementalFont supplementalFont4 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont5 = new A.SupplementalFont() { Script = "Arab", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont6 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont7 = new A.SupplementalFont() { Script = "Thai", Typeface = "Angsana New" };
            A.SupplementalFont supplementalFont8 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont9 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont10 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont11 = new A.SupplementalFont() { Script = "Khmr", Typeface = "MoolBoran" };
            A.SupplementalFont supplementalFont12 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont13 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont14 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont15 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont16 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont17 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont18 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont19 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont20 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont21 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont22 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont23 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont24 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont25 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont26 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont27 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont28 = new A.SupplementalFont() { Script = "Viet", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont29 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont30 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };

            majorFont1.Append(latinFont1);
            majorFont1.Append(eastAsianFont1);
            majorFont1.Append(complexScriptFont1);
            majorFont1.Append(supplementalFont1);
            majorFont1.Append(supplementalFont2);
            majorFont1.Append(supplementalFont3);
            majorFont1.Append(supplementalFont4);
            majorFont1.Append(supplementalFont5);
            majorFont1.Append(supplementalFont6);
            majorFont1.Append(supplementalFont7);
            majorFont1.Append(supplementalFont8);
            majorFont1.Append(supplementalFont9);
            majorFont1.Append(supplementalFont10);
            majorFont1.Append(supplementalFont11);
            majorFont1.Append(supplementalFont12);
            majorFont1.Append(supplementalFont13);
            majorFont1.Append(supplementalFont14);
            majorFont1.Append(supplementalFont15);
            majorFont1.Append(supplementalFont16);
            majorFont1.Append(supplementalFont17);
            majorFont1.Append(supplementalFont18);
            majorFont1.Append(supplementalFont19);
            majorFont1.Append(supplementalFont20);
            majorFont1.Append(supplementalFont21);
            majorFont1.Append(supplementalFont22);
            majorFont1.Append(supplementalFont23);
            majorFont1.Append(supplementalFont24);
            majorFont1.Append(supplementalFont25);
            majorFont1.Append(supplementalFont26);
            majorFont1.Append(supplementalFont27);
            majorFont1.Append(supplementalFont28);
            majorFont1.Append(supplementalFont29);
            majorFont1.Append(supplementalFont30);

            A.MinorFont minorFont1 = new A.MinorFont();
            A.LatinFont latinFont2 = new A.LatinFont() { Typeface = "Calibri", Panose = "020F0502020204030204" };
            A.EastAsianFont eastAsianFont2 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont31 = new A.SupplementalFont() { Script = "Jpan", Typeface = "游明朝" };
            A.SupplementalFont supplementalFont32 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont33 = new A.SupplementalFont() { Script = "Hans", Typeface = "等线" };
            A.SupplementalFont supplementalFont34 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont35 = new A.SupplementalFont() { Script = "Arab", Typeface = "Arial" };
            A.SupplementalFont supplementalFont36 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Arial" };
            A.SupplementalFont supplementalFont37 = new A.SupplementalFont() { Script = "Thai", Typeface = "Cordia New" };
            A.SupplementalFont supplementalFont38 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont39 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont40 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont41 = new A.SupplementalFont() { Script = "Khmr", Typeface = "DaunPenh" };
            A.SupplementalFont supplementalFont42 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont43 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont44 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont45 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont46 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont47 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont48 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont49 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont50 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont51 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont52 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont53 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont54 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont55 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont56 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont57 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont58 = new A.SupplementalFont() { Script = "Viet", Typeface = "Arial" };
            A.SupplementalFont supplementalFont59 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont60 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };

            minorFont1.Append(latinFont2);
            minorFont1.Append(eastAsianFont2);
            minorFont1.Append(complexScriptFont2);
            minorFont1.Append(supplementalFont31);
            minorFont1.Append(supplementalFont32);
            minorFont1.Append(supplementalFont33);
            minorFont1.Append(supplementalFont34);
            minorFont1.Append(supplementalFont35);
            minorFont1.Append(supplementalFont36);
            minorFont1.Append(supplementalFont37);
            minorFont1.Append(supplementalFont38);
            minorFont1.Append(supplementalFont39);
            minorFont1.Append(supplementalFont40);
            minorFont1.Append(supplementalFont41);
            minorFont1.Append(supplementalFont42);
            minorFont1.Append(supplementalFont43);
            minorFont1.Append(supplementalFont44);
            minorFont1.Append(supplementalFont45);
            minorFont1.Append(supplementalFont46);
            minorFont1.Append(supplementalFont47);
            minorFont1.Append(supplementalFont48);
            minorFont1.Append(supplementalFont49);
            minorFont1.Append(supplementalFont50);
            minorFont1.Append(supplementalFont51);
            minorFont1.Append(supplementalFont52);
            minorFont1.Append(supplementalFont53);
            minorFont1.Append(supplementalFont54);
            minorFont1.Append(supplementalFont55);
            minorFont1.Append(supplementalFont56);
            minorFont1.Append(supplementalFont57);
            minorFont1.Append(supplementalFont58);
            minorFont1.Append(supplementalFont59);
            minorFont1.Append(supplementalFont60);

            fontScheme1.Append(majorFont1);
            fontScheme1.Append(minorFont1);

            A.FormatScheme formatScheme1 = new A.FormatScheme() { Name = "Office" };

            A.FillStyleList fillStyleList1 = new A.FillStyleList();

            A.SolidFill solidFill1 = new A.SolidFill();
            A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill1.Append(schemeColor1);

            A.GradientFill gradientFill1 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList1 = new A.GradientStopList();

            A.GradientStop gradientStop1 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation1 = new A.LuminanceModulation() { Val = 110000 };
            A.SaturationModulation saturationModulation1 = new A.SaturationModulation() { Val = 105000 };
            A.Tint tint1 = new A.Tint() { Val = 67000 };

            schemeColor2.Append(luminanceModulation1);
            schemeColor2.Append(saturationModulation1);
            schemeColor2.Append(tint1);

            gradientStop1.Append(schemeColor2);

            A.GradientStop gradientStop2 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation2 = new A.LuminanceModulation() { Val = 105000 };
            A.SaturationModulation saturationModulation2 = new A.SaturationModulation() { Val = 103000 };
            A.Tint tint2 = new A.Tint() { Val = 73000 };

            schemeColor3.Append(luminanceModulation2);
            schemeColor3.Append(saturationModulation2);
            schemeColor3.Append(tint2);

            gradientStop2.Append(schemeColor3);

            A.GradientStop gradientStop3 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor4 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation3 = new A.LuminanceModulation() { Val = 105000 };
            A.SaturationModulation saturationModulation3 = new A.SaturationModulation() { Val = 109000 };
            A.Tint tint3 = new A.Tint() { Val = 81000 };

            schemeColor4.Append(luminanceModulation3);
            schemeColor4.Append(saturationModulation3);
            schemeColor4.Append(tint3);

            gradientStop3.Append(schemeColor4);

            gradientStopList1.Append(gradientStop1);
            gradientStopList1.Append(gradientStop2);
            gradientStopList1.Append(gradientStop3);
            A.LinearGradientFill linearGradientFill1 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill1.Append(gradientStopList1);
            gradientFill1.Append(linearGradientFill1);

            A.GradientFill gradientFill2 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList2 = new A.GradientStopList();

            A.GradientStop gradientStop4 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor5 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation4 = new A.SaturationModulation() { Val = 103000 };
            A.LuminanceModulation luminanceModulation4 = new A.LuminanceModulation() { Val = 102000 };
            A.Tint tint4 = new A.Tint() { Val = 94000 };

            schemeColor5.Append(saturationModulation4);
            schemeColor5.Append(luminanceModulation4);
            schemeColor5.Append(tint4);

            gradientStop4.Append(schemeColor5);

            A.GradientStop gradientStop5 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor6 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation5 = new A.SaturationModulation() { Val = 110000 };
            A.LuminanceModulation luminanceModulation5 = new A.LuminanceModulation() { Val = 100000 };
            A.Shade shade1 = new A.Shade() { Val = 100000 };

            schemeColor6.Append(saturationModulation5);
            schemeColor6.Append(luminanceModulation5);
            schemeColor6.Append(shade1);

            gradientStop5.Append(schemeColor6);

            A.GradientStop gradientStop6 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor7 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation6 = new A.LuminanceModulation() { Val = 99000 };
            A.SaturationModulation saturationModulation6 = new A.SaturationModulation() { Val = 120000 };
            A.Shade shade2 = new A.Shade() { Val = 78000 };

            schemeColor7.Append(luminanceModulation6);
            schemeColor7.Append(saturationModulation6);
            schemeColor7.Append(shade2);

            gradientStop6.Append(schemeColor7);

            gradientStopList2.Append(gradientStop4);
            gradientStopList2.Append(gradientStop5);
            gradientStopList2.Append(gradientStop6);
            A.LinearGradientFill linearGradientFill2 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill2.Append(gradientStopList2);
            gradientFill2.Append(linearGradientFill2);

            fillStyleList1.Append(solidFill1);
            fillStyleList1.Append(gradientFill1);
            fillStyleList1.Append(gradientFill2);

            A.LineStyleList lineStyleList1 = new A.LineStyleList();

            A.Outline outline1 = new A.Outline() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill2 = new A.SolidFill();
            A.SchemeColor schemeColor8 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill2.Append(schemeColor8);
            A.PresetDash presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter1 = new A.Miter() { Limit = 800000 };

            outline1.Append(solidFill2);
            outline1.Append(presetDash1);
            outline1.Append(miter1);

            A.Outline outline2 = new A.Outline() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill3 = new A.SolidFill();
            A.SchemeColor schemeColor9 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill3.Append(schemeColor9);
            A.PresetDash presetDash2 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter2 = new A.Miter() { Limit = 800000 };

            outline2.Append(solidFill3);
            outline2.Append(presetDash2);
            outline2.Append(miter2);

            A.Outline outline3 = new A.Outline() { Width = 19050, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill4 = new A.SolidFill();
            A.SchemeColor schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill4.Append(schemeColor10);
            A.PresetDash presetDash3 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter3 = new A.Miter() { Limit = 800000 };

            outline3.Append(solidFill4);
            outline3.Append(presetDash3);
            outline3.Append(miter3);

            lineStyleList1.Append(outline1);
            lineStyleList1.Append(outline2);
            lineStyleList1.Append(outline3);

            A.EffectStyleList effectStyleList1 = new A.EffectStyleList();

            A.EffectStyle effectStyle1 = new A.EffectStyle();
            A.EffectList effectList1 = new A.EffectList();

            effectStyle1.Append(effectList1);

            A.EffectStyle effectStyle2 = new A.EffectStyle();
            A.EffectList effectList2 = new A.EffectList();

            effectStyle2.Append(effectList2);

            A.EffectStyle effectStyle3 = new A.EffectStyle();

            A.EffectList effectList3 = new A.EffectList();

            A.OuterShadow outerShadow1 = new A.OuterShadow() { BlurRadius = 57150L, Distance = 19050L, Direction = 5400000, Alignment = A.RectangleAlignmentValues.Center, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex11 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha1 = new A.Alpha() { Val = 63000 };

            rgbColorModelHex11.Append(alpha1);

            outerShadow1.Append(rgbColorModelHex11);

            effectList3.Append(outerShadow1);

            effectStyle3.Append(effectList3);

            effectStyleList1.Append(effectStyle1);
            effectStyleList1.Append(effectStyle2);
            effectStyleList1.Append(effectStyle3);

            A.BackgroundFillStyleList backgroundFillStyleList1 = new A.BackgroundFillStyleList();

            A.SolidFill solidFill5 = new A.SolidFill();
            A.SchemeColor schemeColor11 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill5.Append(schemeColor11);

            A.SolidFill solidFill6 = new A.SolidFill();

            A.SchemeColor schemeColor12 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint5 = new A.Tint() { Val = 95000 };
            A.SaturationModulation saturationModulation7 = new A.SaturationModulation() { Val = 170000 };

            schemeColor12.Append(tint5);
            schemeColor12.Append(saturationModulation7);

            solidFill6.Append(schemeColor12);

            A.GradientFill gradientFill3 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList3 = new A.GradientStopList();

            A.GradientStop gradientStop7 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor13 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint6 = new A.Tint() { Val = 93000 };
            A.SaturationModulation saturationModulation8 = new A.SaturationModulation() { Val = 150000 };
            A.Shade shade3 = new A.Shade() { Val = 98000 };
            A.LuminanceModulation luminanceModulation7 = new A.LuminanceModulation() { Val = 102000 };

            schemeColor13.Append(tint6);
            schemeColor13.Append(saturationModulation8);
            schemeColor13.Append(shade3);
            schemeColor13.Append(luminanceModulation7);

            gradientStop7.Append(schemeColor13);

            A.GradientStop gradientStop8 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor14 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint7 = new A.Tint() { Val = 98000 };
            A.SaturationModulation saturationModulation9 = new A.SaturationModulation() { Val = 130000 };
            A.Shade shade4 = new A.Shade() { Val = 90000 };
            A.LuminanceModulation luminanceModulation8 = new A.LuminanceModulation() { Val = 103000 };

            schemeColor14.Append(tint7);
            schemeColor14.Append(saturationModulation9);
            schemeColor14.Append(shade4);
            schemeColor14.Append(luminanceModulation8);

            gradientStop8.Append(schemeColor14);

            A.GradientStop gradientStop9 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor15 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade5 = new A.Shade() { Val = 63000 };
            A.SaturationModulation saturationModulation10 = new A.SaturationModulation() { Val = 120000 };

            schemeColor15.Append(shade5);
            schemeColor15.Append(saturationModulation10);

            gradientStop9.Append(schemeColor15);

            gradientStopList3.Append(gradientStop7);
            gradientStopList3.Append(gradientStop8);
            gradientStopList3.Append(gradientStop9);
            A.LinearGradientFill linearGradientFill3 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill3.Append(gradientStopList3);
            gradientFill3.Append(linearGradientFill3);

            backgroundFillStyleList1.Append(solidFill5);
            backgroundFillStyleList1.Append(solidFill6);
            backgroundFillStyleList1.Append(gradientFill3);

            formatScheme1.Append(fillStyleList1);
            formatScheme1.Append(lineStyleList1);
            formatScheme1.Append(effectStyleList1);
            formatScheme1.Append(backgroundFillStyleList1);

            themeElements1.Append(colorScheme1);
            themeElements1.Append(fontScheme1);
            themeElements1.Append(formatScheme1);
            A.ObjectDefaults objectDefaults1 = new A.ObjectDefaults();
            A.ExtraColorSchemeList extraColorSchemeList1 = new A.ExtraColorSchemeList();

            A.OfficeStyleSheetExtensionList officeStyleSheetExtensionList1 = new A.OfficeStyleSheetExtensionList();

            A.OfficeStyleSheetExtension officeStyleSheetExtension1 = new A.OfficeStyleSheetExtension() { Uri = "{05A4C25C-085E-4340-85A3-A5531E510DB2}" };

            Thm15.ThemeFamily themeFamily1 = new Thm15.ThemeFamily() { Name = "Office Theme", Id = "{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}", Vid = "{4A3C46E8-61CC-4603-A589-7422A47A8E4A}" };
            themeFamily1.AddNamespaceDeclaration("thm15", "http://schemas.microsoft.com/office/thememl/2012/main");

            officeStyleSheetExtension1.Append(themeFamily1);

            officeStyleSheetExtensionList1.Append(officeStyleSheetExtension1);

            theme1.Append(themeElements1);
            theme1.Append(objectDefaults1);
            theme1.Append(extraColorSchemeList1);
            theme1.Append(officeStyleSheetExtensionList1);

            themePart1.Theme = theme1;
        }

        // Generates content of webSettingsPart1.
        private void GenerateWebSettingsPart1Content(WebSettingsPart webSettingsPart1)
        {
            WebSettings webSettings1 = new WebSettings() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se" } };
            webSettings1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            webSettings1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            webSettings1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            webSettings1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            webSettings1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            webSettings1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            OptimizeForBrowser optimizeForBrowser1 = new OptimizeForBrowser();

            webSettings1.Append(optimizeForBrowser1);

            webSettingsPart1.WebSettings = webSettings1;
        }

        // Generates content of wordprocessingPeoplePart1.
        private void GenerateWordprocessingPeoplePart1Content(WordprocessingPeoplePart wordprocessingPeoplePart1)
        {
            W15.People people1 = new W15.People() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se wp14" } };
            people1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            people1.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
            people1.AddNamespaceDeclaration("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
            people1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            people1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            people1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            people1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            people1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            people1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            people1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            people1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            people1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            people1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            people1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            people1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            people1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            people1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            people1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            people1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            W15.Person person1 = new W15.Person() { Author = "JIMENEZ RUBIA Luis (SG-EXT)" };
            W15.PresenceInfo presenceInfo1 = new W15.PresenceInfo() { ProviderId = "None", UserId = "JIMENEZ RUBIA Luis (SG-EXT)" };

            person1.Append(presenceInfo1);

            people1.Append(person1);

            wordprocessingPeoplePart1.People = people1;
        }

        // Generates content of documentSettingsPart1.
        private void GenerateDocumentSettingsPart1Content(DocumentSettingsPart documentSettingsPart1)
        {
            Settings settings1 = new Settings() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se" } };
            settings1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            settings1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            settings1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            settings1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            settings1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            settings1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            settings1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            settings1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            settings1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            settings1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            settings1.AddNamespaceDeclaration("sl", "http://schemas.openxmlformats.org/schemaLibrary/2006/main");
            Zoom zoom1 = new Zoom() { Percent = "100" };
            ActiveWritingStyle activeWritingStyle1 = new ActiveWritingStyle() { Language = "fr-BE", VendorID = (UInt16Value)64U, DllVersion = 131078, NaturalLanguageGrammarCheck = true, CheckStyle = false, ApplicationName = "MSWord" };
            ActiveWritingStyle activeWritingStyle2 = new ActiveWritingStyle() { Language = "en-GB", VendorID = (UInt16Value)64U, DllVersion = 131078, NaturalLanguageGrammarCheck = true, CheckStyle = true, ApplicationName = "MSWord" };
            ProofState proofState1 = new ProofState() { Spelling = ProofingStateValues.Clean, Grammar = ProofingStateValues.Clean };
            DefaultTabStop defaultTabStop1 = new DefaultTabStop() { Val = 720 };
            CharacterSpacingControl characterSpacingControl1 = new CharacterSpacingControl() { Val = CharacterSpacingValues.DoNotCompress };

            Compatibility compatibility1 = new Compatibility();
            CompatibilitySetting compatibilitySetting1 = new CompatibilitySetting() { Name = CompatSettingNameValues.CompatibilityMode, Uri = "http://schemas.microsoft.com/office/word", Val = "14" };
            CompatibilitySetting compatibilitySetting2 = new CompatibilitySetting() { Name = CompatSettingNameValues.OverrideTableStyleFontSizeAndJustification, Uri = "http://schemas.microsoft.com/office/word", Val = "1" };
            CompatibilitySetting compatibilitySetting3 = new CompatibilitySetting() { Name = CompatSettingNameValues.EnableOpenTypeFeatures, Uri = "http://schemas.microsoft.com/office/word", Val = "1" };
            CompatibilitySetting compatibilitySetting4 = new CompatibilitySetting() { Name = CompatSettingNameValues.DoNotFlipMirrorIndents, Uri = "http://schemas.microsoft.com/office/word", Val = "1" };
            CompatibilitySetting compatibilitySetting5 = new CompatibilitySetting() { Name = CompatSettingNameValues.DifferentiateMultirowTableHeaders, Uri = "http://schemas.microsoft.com/office/word", Val = "1" };

            compatibility1.Append(compatibilitySetting1);
            compatibility1.Append(compatibilitySetting2);
            compatibility1.Append(compatibilitySetting3);
            compatibility1.Append(compatibilitySetting4);
            compatibility1.Append(compatibilitySetting5);

            Rsids rsids1 = new Rsids();
            RsidRoot rsidRoot1 = new RsidRoot() { Val = "005F6638" };
            Rsid rsid1 = new Rsid() { Val = "005F6638" };
            Rsid rsid2 = new Rsid() { Val = "00C33253" };

            rsids1.Append(rsidRoot1);
            rsids1.Append(rsid1);
            rsids1.Append(rsid2);

            M.MathProperties mathProperties1 = new M.MathProperties();
            M.MathFont mathFont1 = new M.MathFont() { Val = "Cambria Math" };
            M.BreakBinary breakBinary1 = new M.BreakBinary() { Val = M.BreakBinaryOperatorValues.Before };
            M.BreakBinarySubtraction breakBinarySubtraction1 = new M.BreakBinarySubtraction() { Val = M.BreakBinarySubtractionValues.MinusMinus };
            M.SmallFraction smallFraction1 = new M.SmallFraction() { Val = M.BooleanValues.Zero };
            M.DisplayDefaults displayDefaults1 = new M.DisplayDefaults();
            M.LeftMargin leftMargin1 = new M.LeftMargin() { Val = (UInt32Value)0U };
            M.RightMargin rightMargin1 = new M.RightMargin() { Val = (UInt32Value)0U };
            M.DefaultJustification defaultJustification1 = new M.DefaultJustification() { Val = M.JustificationValues.CenterGroup };
            M.WrapIndent wrapIndent1 = new M.WrapIndent() { Val = (UInt32Value)1440U };
            M.IntegralLimitLocation integralLimitLocation1 = new M.IntegralLimitLocation() { Val = M.LimitLocationValues.SubscriptSuperscript };
            M.NaryLimitLocation naryLimitLocation1 = new M.NaryLimitLocation() { Val = M.LimitLocationValues.UnderOver };

            mathProperties1.Append(mathFont1);
            mathProperties1.Append(breakBinary1);
            mathProperties1.Append(breakBinarySubtraction1);
            mathProperties1.Append(smallFraction1);
            mathProperties1.Append(displayDefaults1);
            mathProperties1.Append(leftMargin1);
            mathProperties1.Append(rightMargin1);
            mathProperties1.Append(defaultJustification1);
            mathProperties1.Append(wrapIndent1);
            mathProperties1.Append(integralLimitLocation1);
            mathProperties1.Append(naryLimitLocation1);
            ThemeFontLanguages themeFontLanguages1 = new ThemeFontLanguages() { Val = "en-US" };
            ColorSchemeMapping colorSchemeMapping1 = new ColorSchemeMapping() { Background1 = ColorSchemeIndexValues.Light1, Text1 = ColorSchemeIndexValues.Dark1, Background2 = ColorSchemeIndexValues.Light2, Text2 = ColorSchemeIndexValues.Dark2, Accent1 = ColorSchemeIndexValues.Accent1, Accent2 = ColorSchemeIndexValues.Accent2, Accent3 = ColorSchemeIndexValues.Accent3, Accent4 = ColorSchemeIndexValues.Accent4, Accent5 = ColorSchemeIndexValues.Accent5, Accent6 = ColorSchemeIndexValues.Accent6, Hyperlink = ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = ColorSchemeIndexValues.FollowedHyperlink };

            ShapeDefaults shapeDefaults1 = new ShapeDefaults();
            Ovml.ShapeDefaults shapeDefaults2 = new Ovml.ShapeDefaults() { Extension = V.ExtensionHandlingBehaviorValues.Edit, MaxShapeId = 1026 };

            Ovml.ShapeLayout shapeLayout1 = new Ovml.ShapeLayout() { Extension = V.ExtensionHandlingBehaviorValues.Edit };
            Ovml.ShapeIdMap shapeIdMap1 = new Ovml.ShapeIdMap() { Extension = V.ExtensionHandlingBehaviorValues.Edit, Data = "1" };

            shapeLayout1.Append(shapeIdMap1);

            shapeDefaults1.Append(shapeDefaults2);
            shapeDefaults1.Append(shapeLayout1);
            DecimalSymbol decimalSymbol1 = new DecimalSymbol() { Val = "," };
            ListSeparator listSeparator1 = new ListSeparator() { Val = ";" };
            W14.DocumentId documentId1 = new W14.DocumentId() { Val = "654105D3" };
            W15.ChartTrackingRefBased chartTrackingRefBased1 = new W15.ChartTrackingRefBased();
            W15.PersistentDocumentId persistentDocumentId1 = new W15.PersistentDocumentId() { Val = "{3129E8BA-5C72-4431-B9BD-96662FFC00FE}" };

            settings1.Append(zoom1);
            settings1.Append(activeWritingStyle1);
            settings1.Append(activeWritingStyle2);
            settings1.Append(proofState1);
            settings1.Append(defaultTabStop1);
            settings1.Append(characterSpacingControl1);
            settings1.Append(compatibility1);
            settings1.Append(rsids1);
            settings1.Append(mathProperties1);
            settings1.Append(themeFontLanguages1);
            settings1.Append(colorSchemeMapping1);
            settings1.Append(shapeDefaults1);
            settings1.Append(decimalSymbol1);
            settings1.Append(listSeparator1);
            settings1.Append(documentId1);
            settings1.Append(chartTrackingRefBased1);
            settings1.Append(persistentDocumentId1);

            documentSettingsPart1.Settings = settings1;
        }

        // Generates content of styleDefinitionsPart1.
        private void GenerateStyleDefinitionsPart1Content(StyleDefinitionsPart styleDefinitionsPart1)
        {
            Styles styles1 = new Styles() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se" } };
            styles1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            styles1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            styles1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            styles1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            styles1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            styles1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");

            DocDefaults docDefaults1 = new DocDefaults();

            RunPropertiesDefault runPropertiesDefault1 = new RunPropertiesDefault();

            RunPropertiesBaseStyle runPropertiesBaseStyle1 = new RunPropertiesBaseStyle();
            RunFonts runFonts51 = new RunFonts() { AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi, EastAsiaTheme = ThemeFontValues.MinorHighAnsi, ComplexScriptTheme = ThemeFontValues.MinorBidi };
            FontSize fontSize38 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript32 = new FontSizeComplexScript() { Val = "22" };
            Languages languages16 = new Languages() { Val = "en-US", EastAsia = "en-US", Bidi = "ar-SA" };

            runPropertiesBaseStyle1.Append(runFonts51);
            runPropertiesBaseStyle1.Append(fontSize38);
            runPropertiesBaseStyle1.Append(fontSizeComplexScript32);
            runPropertiesBaseStyle1.Append(languages16);

            runPropertiesDefault1.Append(runPropertiesBaseStyle1);

            ParagraphPropertiesDefault paragraphPropertiesDefault1 = new ParagraphPropertiesDefault();

            ParagraphPropertiesBaseStyle paragraphPropertiesBaseStyle1 = new ParagraphPropertiesBaseStyle();
            SpacingBetweenLines spacingBetweenLines11 = new SpacingBetweenLines() { After = "160", Line = "259", LineRule = LineSpacingRuleValues.Auto };

            paragraphPropertiesBaseStyle1.Append(spacingBetweenLines11);

            paragraphPropertiesDefault1.Append(paragraphPropertiesBaseStyle1);

            docDefaults1.Append(runPropertiesDefault1);
            docDefaults1.Append(paragraphPropertiesDefault1);

            LatentStyles latentStyles1 = new LatentStyles() { DefaultLockedState = false, DefaultUiPriority = 99, DefaultSemiHidden = false, DefaultUnhideWhenUsed = false, DefaultPrimaryStyle = false, Count = 371 };
            LatentStyleExceptionInfo latentStyleExceptionInfo1 = new LatentStyleExceptionInfo() { Name = "Normal", UiPriority = 0, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo2 = new LatentStyleExceptionInfo() { Name = "heading 1", UiPriority = 0, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo3 = new LatentStyleExceptionInfo() { Name = "heading 2", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo4 = new LatentStyleExceptionInfo() { Name = "heading 3", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo5 = new LatentStyleExceptionInfo() { Name = "heading 4", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo6 = new LatentStyleExceptionInfo() { Name = "heading 5", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo7 = new LatentStyleExceptionInfo() { Name = "heading 6", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo8 = new LatentStyleExceptionInfo() { Name = "heading 7", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo9 = new LatentStyleExceptionInfo() { Name = "heading 8", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo10 = new LatentStyleExceptionInfo() { Name = "heading 9", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo11 = new LatentStyleExceptionInfo() { Name = "index 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo12 = new LatentStyleExceptionInfo() { Name = "index 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo13 = new LatentStyleExceptionInfo() { Name = "index 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo14 = new LatentStyleExceptionInfo() { Name = "index 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo15 = new LatentStyleExceptionInfo() { Name = "index 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo16 = new LatentStyleExceptionInfo() { Name = "index 6", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo17 = new LatentStyleExceptionInfo() { Name = "index 7", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo18 = new LatentStyleExceptionInfo() { Name = "index 8", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo19 = new LatentStyleExceptionInfo() { Name = "index 9", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo20 = new LatentStyleExceptionInfo() { Name = "toc 1", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo21 = new LatentStyleExceptionInfo() { Name = "toc 2", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo22 = new LatentStyleExceptionInfo() { Name = "toc 3", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo23 = new LatentStyleExceptionInfo() { Name = "toc 4", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo24 = new LatentStyleExceptionInfo() { Name = "toc 5", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo25 = new LatentStyleExceptionInfo() { Name = "toc 6", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo26 = new LatentStyleExceptionInfo() { Name = "toc 7", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo27 = new LatentStyleExceptionInfo() { Name = "toc 8", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo28 = new LatentStyleExceptionInfo() { Name = "toc 9", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo29 = new LatentStyleExceptionInfo() { Name = "Normal Indent", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo30 = new LatentStyleExceptionInfo() { Name = "footnote text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo31 = new LatentStyleExceptionInfo() { Name = "annotation text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo32 = new LatentStyleExceptionInfo() { Name = "header", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo33 = new LatentStyleExceptionInfo() { Name = "footer", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo34 = new LatentStyleExceptionInfo() { Name = "index heading", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo35 = new LatentStyleExceptionInfo() { Name = "caption", UiPriority = 35, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo36 = new LatentStyleExceptionInfo() { Name = "table of figures", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo37 = new LatentStyleExceptionInfo() { Name = "envelope address", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo38 = new LatentStyleExceptionInfo() { Name = "envelope return", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo39 = new LatentStyleExceptionInfo() { Name = "footnote reference", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo40 = new LatentStyleExceptionInfo() { Name = "annotation reference", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo41 = new LatentStyleExceptionInfo() { Name = "line number", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo42 = new LatentStyleExceptionInfo() { Name = "page number", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo43 = new LatentStyleExceptionInfo() { Name = "endnote reference", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo44 = new LatentStyleExceptionInfo() { Name = "endnote text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo45 = new LatentStyleExceptionInfo() { Name = "table of authorities", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo46 = new LatentStyleExceptionInfo() { Name = "macro", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo47 = new LatentStyleExceptionInfo() { Name = "toa heading", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo48 = new LatentStyleExceptionInfo() { Name = "List", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo49 = new LatentStyleExceptionInfo() { Name = "List Bullet", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo50 = new LatentStyleExceptionInfo() { Name = "List Number", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo51 = new LatentStyleExceptionInfo() { Name = "List 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo52 = new LatentStyleExceptionInfo() { Name = "List 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo53 = new LatentStyleExceptionInfo() { Name = "List 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo54 = new LatentStyleExceptionInfo() { Name = "List 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo55 = new LatentStyleExceptionInfo() { Name = "List Bullet 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo56 = new LatentStyleExceptionInfo() { Name = "List Bullet 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo57 = new LatentStyleExceptionInfo() { Name = "List Bullet 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo58 = new LatentStyleExceptionInfo() { Name = "List Bullet 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo59 = new LatentStyleExceptionInfo() { Name = "List Number 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo60 = new LatentStyleExceptionInfo() { Name = "List Number 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo61 = new LatentStyleExceptionInfo() { Name = "List Number 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo62 = new LatentStyleExceptionInfo() { Name = "List Number 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo63 = new LatentStyleExceptionInfo() { Name = "Title", UiPriority = 10, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo64 = new LatentStyleExceptionInfo() { Name = "Closing", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo65 = new LatentStyleExceptionInfo() { Name = "Signature", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo66 = new LatentStyleExceptionInfo() { Name = "Default Paragraph Font", UiPriority = 1, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo67 = new LatentStyleExceptionInfo() { Name = "Body Text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo68 = new LatentStyleExceptionInfo() { Name = "Body Text Indent", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo69 = new LatentStyleExceptionInfo() { Name = "List Continue", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo70 = new LatentStyleExceptionInfo() { Name = "List Continue 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo71 = new LatentStyleExceptionInfo() { Name = "List Continue 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo72 = new LatentStyleExceptionInfo() { Name = "List Continue 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo73 = new LatentStyleExceptionInfo() { Name = "List Continue 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo74 = new LatentStyleExceptionInfo() { Name = "Message Header", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo75 = new LatentStyleExceptionInfo() { Name = "Subtitle", UiPriority = 11, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo76 = new LatentStyleExceptionInfo() { Name = "Salutation", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo77 = new LatentStyleExceptionInfo() { Name = "Date", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo78 = new LatentStyleExceptionInfo() { Name = "Body Text First Indent", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo79 = new LatentStyleExceptionInfo() { Name = "Body Text First Indent 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo80 = new LatentStyleExceptionInfo() { Name = "Note Heading", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo81 = new LatentStyleExceptionInfo() { Name = "Body Text 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo82 = new LatentStyleExceptionInfo() { Name = "Body Text 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo83 = new LatentStyleExceptionInfo() { Name = "Body Text Indent 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo84 = new LatentStyleExceptionInfo() { Name = "Body Text Indent 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo85 = new LatentStyleExceptionInfo() { Name = "Block Text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo86 = new LatentStyleExceptionInfo() { Name = "Hyperlink", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo87 = new LatentStyleExceptionInfo() { Name = "FollowedHyperlink", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo88 = new LatentStyleExceptionInfo() { Name = "Strong", UiPriority = 22, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo89 = new LatentStyleExceptionInfo() { Name = "Emphasis", UiPriority = 20, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo90 = new LatentStyleExceptionInfo() { Name = "Document Map", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo91 = new LatentStyleExceptionInfo() { Name = "Plain Text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo92 = new LatentStyleExceptionInfo() { Name = "E-mail Signature", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo93 = new LatentStyleExceptionInfo() { Name = "HTML Top of Form", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo94 = new LatentStyleExceptionInfo() { Name = "HTML Bottom of Form", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo95 = new LatentStyleExceptionInfo() { Name = "Normal (Web)", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo96 = new LatentStyleExceptionInfo() { Name = "HTML Acronym", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo97 = new LatentStyleExceptionInfo() { Name = "HTML Address", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo98 = new LatentStyleExceptionInfo() { Name = "HTML Cite", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo99 = new LatentStyleExceptionInfo() { Name = "HTML Code", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo100 = new LatentStyleExceptionInfo() { Name = "HTML Definition", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo101 = new LatentStyleExceptionInfo() { Name = "HTML Keyboard", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo102 = new LatentStyleExceptionInfo() { Name = "HTML Preformatted", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo103 = new LatentStyleExceptionInfo() { Name = "HTML Sample", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo104 = new LatentStyleExceptionInfo() { Name = "HTML Typewriter", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo105 = new LatentStyleExceptionInfo() { Name = "HTML Variable", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo106 = new LatentStyleExceptionInfo() { Name = "Normal Table", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo107 = new LatentStyleExceptionInfo() { Name = "annotation subject", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo108 = new LatentStyleExceptionInfo() { Name = "No List", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo109 = new LatentStyleExceptionInfo() { Name = "Outline List 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo110 = new LatentStyleExceptionInfo() { Name = "Outline List 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo111 = new LatentStyleExceptionInfo() { Name = "Outline List 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo112 = new LatentStyleExceptionInfo() { Name = "Table Simple 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo113 = new LatentStyleExceptionInfo() { Name = "Table Simple 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo114 = new LatentStyleExceptionInfo() { Name = "Table Simple 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo115 = new LatentStyleExceptionInfo() { Name = "Table Classic 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo116 = new LatentStyleExceptionInfo() { Name = "Table Classic 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo117 = new LatentStyleExceptionInfo() { Name = "Table Classic 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo118 = new LatentStyleExceptionInfo() { Name = "Table Classic 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo119 = new LatentStyleExceptionInfo() { Name = "Table Colorful 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo120 = new LatentStyleExceptionInfo() { Name = "Table Colorful 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo121 = new LatentStyleExceptionInfo() { Name = "Table Colorful 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo122 = new LatentStyleExceptionInfo() { Name = "Table Columns 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo123 = new LatentStyleExceptionInfo() { Name = "Table Columns 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo124 = new LatentStyleExceptionInfo() { Name = "Table Columns 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo125 = new LatentStyleExceptionInfo() { Name = "Table Columns 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo126 = new LatentStyleExceptionInfo() { Name = "Table Columns 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo127 = new LatentStyleExceptionInfo() { Name = "Table Grid 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo128 = new LatentStyleExceptionInfo() { Name = "Table Grid 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo129 = new LatentStyleExceptionInfo() { Name = "Table Grid 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo130 = new LatentStyleExceptionInfo() { Name = "Table Grid 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo131 = new LatentStyleExceptionInfo() { Name = "Table Grid 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo132 = new LatentStyleExceptionInfo() { Name = "Table Grid 6", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo133 = new LatentStyleExceptionInfo() { Name = "Table Grid 7", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo134 = new LatentStyleExceptionInfo() { Name = "Table Grid 8", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo135 = new LatentStyleExceptionInfo() { Name = "Table List 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo136 = new LatentStyleExceptionInfo() { Name = "Table List 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo137 = new LatentStyleExceptionInfo() { Name = "Table List 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo138 = new LatentStyleExceptionInfo() { Name = "Table List 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo139 = new LatentStyleExceptionInfo() { Name = "Table List 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo140 = new LatentStyleExceptionInfo() { Name = "Table List 6", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo141 = new LatentStyleExceptionInfo() { Name = "Table List 7", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo142 = new LatentStyleExceptionInfo() { Name = "Table List 8", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo143 = new LatentStyleExceptionInfo() { Name = "Table 3D effects 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo144 = new LatentStyleExceptionInfo() { Name = "Table 3D effects 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo145 = new LatentStyleExceptionInfo() { Name = "Table 3D effects 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo146 = new LatentStyleExceptionInfo() { Name = "Table Contemporary", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo147 = new LatentStyleExceptionInfo() { Name = "Table Elegant", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo148 = new LatentStyleExceptionInfo() { Name = "Table Professional", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo149 = new LatentStyleExceptionInfo() { Name = "Table Subtle 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo150 = new LatentStyleExceptionInfo() { Name = "Table Subtle 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo151 = new LatentStyleExceptionInfo() { Name = "Table Web 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo152 = new LatentStyleExceptionInfo() { Name = "Table Web 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo153 = new LatentStyleExceptionInfo() { Name = "Table Web 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo154 = new LatentStyleExceptionInfo() { Name = "Balloon Text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo155 = new LatentStyleExceptionInfo() { Name = "Table Grid", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo156 = new LatentStyleExceptionInfo() { Name = "Table Theme", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo157 = new LatentStyleExceptionInfo() { Name = "Placeholder Text", SemiHidden = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo158 = new LatentStyleExceptionInfo() { Name = "No Spacing", UiPriority = 1, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo159 = new LatentStyleExceptionInfo() { Name = "Light Shading", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo160 = new LatentStyleExceptionInfo() { Name = "Light List", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo161 = new LatentStyleExceptionInfo() { Name = "Light Grid", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo162 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo163 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo164 = new LatentStyleExceptionInfo() { Name = "Medium List 1", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo165 = new LatentStyleExceptionInfo() { Name = "Medium List 2", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo166 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo167 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo168 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo169 = new LatentStyleExceptionInfo() { Name = "Dark List", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo170 = new LatentStyleExceptionInfo() { Name = "Colorful Shading", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo171 = new LatentStyleExceptionInfo() { Name = "Colorful List", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo172 = new LatentStyleExceptionInfo() { Name = "Colorful Grid", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo173 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 1", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo174 = new LatentStyleExceptionInfo() { Name = "Light List Accent 1", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo175 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 1", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo176 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 1", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo177 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 1", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo178 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 1", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo179 = new LatentStyleExceptionInfo() { Name = "Revision", SemiHidden = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo180 = new LatentStyleExceptionInfo() { Name = "List Paragraph", UiPriority = 34, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo181 = new LatentStyleExceptionInfo() { Name = "Quote", UiPriority = 29, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo182 = new LatentStyleExceptionInfo() { Name = "Intense Quote", UiPriority = 30, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo183 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 1", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo184 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 1", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo185 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 1", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo186 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 1", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo187 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 1", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo188 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 1", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo189 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 1", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo190 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 1", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo191 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 2", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo192 = new LatentStyleExceptionInfo() { Name = "Light List Accent 2", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo193 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 2", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo194 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 2", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo195 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 2", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo196 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 2", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo197 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 2", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo198 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 2", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo199 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 2", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo200 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 2", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo201 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 2", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo202 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 2", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo203 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 2", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo204 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 2", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo205 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 3", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo206 = new LatentStyleExceptionInfo() { Name = "Light List Accent 3", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo207 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 3", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo208 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 3", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo209 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 3", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo210 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 3", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo211 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 3", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo212 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 3", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo213 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 3", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo214 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 3", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo215 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 3", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo216 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 3", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo217 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 3", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo218 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 3", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo219 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 4", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo220 = new LatentStyleExceptionInfo() { Name = "Light List Accent 4", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo221 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 4", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo222 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 4", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo223 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 4", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo224 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 4", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo225 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 4", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo226 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 4", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo227 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 4", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo228 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 4", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo229 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 4", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo230 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 4", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo231 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 4", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo232 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 4", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo233 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 5", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo234 = new LatentStyleExceptionInfo() { Name = "Light List Accent 5", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo235 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 5", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo236 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 5", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo237 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 5", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo238 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 5", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo239 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 5", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo240 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 5", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo241 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 5", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo242 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 5", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo243 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 5", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo244 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 5", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo245 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 5", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo246 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 5", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo247 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 6", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo248 = new LatentStyleExceptionInfo() { Name = "Light List Accent 6", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo249 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 6", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo250 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 6", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo251 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 6", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo252 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 6", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo253 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 6", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo254 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 6", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo255 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 6", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo256 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 6", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo257 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 6", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo258 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 6", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo259 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 6", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo260 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 6", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo261 = new LatentStyleExceptionInfo() { Name = "Subtle Emphasis", UiPriority = 19, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo262 = new LatentStyleExceptionInfo() { Name = "Intense Emphasis", UiPriority = 21, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo263 = new LatentStyleExceptionInfo() { Name = "Subtle Reference", UiPriority = 31, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo264 = new LatentStyleExceptionInfo() { Name = "Intense Reference", UiPriority = 32, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo265 = new LatentStyleExceptionInfo() { Name = "Book Title", UiPriority = 33, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo266 = new LatentStyleExceptionInfo() { Name = "Bibliography", UiPriority = 37, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo267 = new LatentStyleExceptionInfo() { Name = "TOC Heading", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo268 = new LatentStyleExceptionInfo() { Name = "Plain Table 1", UiPriority = 41 };
            LatentStyleExceptionInfo latentStyleExceptionInfo269 = new LatentStyleExceptionInfo() { Name = "Plain Table 2", UiPriority = 42 };
            LatentStyleExceptionInfo latentStyleExceptionInfo270 = new LatentStyleExceptionInfo() { Name = "Plain Table 3", UiPriority = 43 };
            LatentStyleExceptionInfo latentStyleExceptionInfo271 = new LatentStyleExceptionInfo() { Name = "Plain Table 4", UiPriority = 44 };
            LatentStyleExceptionInfo latentStyleExceptionInfo272 = new LatentStyleExceptionInfo() { Name = "Plain Table 5", UiPriority = 45 };
            LatentStyleExceptionInfo latentStyleExceptionInfo273 = new LatentStyleExceptionInfo() { Name = "Grid Table Light", UiPriority = 40 };
            LatentStyleExceptionInfo latentStyleExceptionInfo274 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo275 = new LatentStyleExceptionInfo() { Name = "Grid Table 2", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo276 = new LatentStyleExceptionInfo() { Name = "Grid Table 3", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo277 = new LatentStyleExceptionInfo() { Name = "Grid Table 4", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo278 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo279 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo280 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo281 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 1", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo282 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 1", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo283 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 1", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo284 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 1", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo285 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 1", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo286 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 1", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo287 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 1", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo288 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 2", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo289 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 2", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo290 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 2", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo291 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 2", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo292 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 2", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo293 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 2", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo294 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 2", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo295 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 3", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo296 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 3", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo297 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 3", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo298 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 3", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo299 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 3", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo300 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 3", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo301 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 3", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo302 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 4", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo303 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 4", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo304 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 4", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo305 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 4", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo306 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 4", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo307 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 4", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo308 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 4", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo309 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 5", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo310 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 5", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo311 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 5", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo312 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 5", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo313 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 5", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo314 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 5", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo315 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 5", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo316 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 6", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo317 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 6", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo318 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 6", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo319 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 6", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo320 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 6", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo321 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 6", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo322 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 6", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo323 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo324 = new LatentStyleExceptionInfo() { Name = "List Table 2", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo325 = new LatentStyleExceptionInfo() { Name = "List Table 3", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo326 = new LatentStyleExceptionInfo() { Name = "List Table 4", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo327 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo328 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo329 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo330 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 1", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo331 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 1", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo332 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 1", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo333 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 1", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo334 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 1", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo335 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 1", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo336 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 1", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo337 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 2", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo338 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 2", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo339 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 2", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo340 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 2", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo341 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 2", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo342 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 2", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo343 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 2", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo344 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 3", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo345 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 3", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo346 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 3", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo347 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 3", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo348 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 3", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo349 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 3", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo350 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 3", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo351 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 4", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo352 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 4", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo353 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 4", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo354 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 4", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo355 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 4", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo356 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 4", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo357 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 4", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo358 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 5", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo359 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 5", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo360 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 5", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo361 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 5", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo362 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 5", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo363 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 5", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo364 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 5", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo365 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 6", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo366 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 6", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo367 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 6", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo368 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 6", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo369 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 6", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo370 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 6", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo371 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 6", UiPriority = 52 };

            latentStyles1.Append(latentStyleExceptionInfo1);
            latentStyles1.Append(latentStyleExceptionInfo2);
            latentStyles1.Append(latentStyleExceptionInfo3);
            latentStyles1.Append(latentStyleExceptionInfo4);
            latentStyles1.Append(latentStyleExceptionInfo5);
            latentStyles1.Append(latentStyleExceptionInfo6);
            latentStyles1.Append(latentStyleExceptionInfo7);
            latentStyles1.Append(latentStyleExceptionInfo8);
            latentStyles1.Append(latentStyleExceptionInfo9);
            latentStyles1.Append(latentStyleExceptionInfo10);
            latentStyles1.Append(latentStyleExceptionInfo11);
            latentStyles1.Append(latentStyleExceptionInfo12);
            latentStyles1.Append(latentStyleExceptionInfo13);
            latentStyles1.Append(latentStyleExceptionInfo14);
            latentStyles1.Append(latentStyleExceptionInfo15);
            latentStyles1.Append(latentStyleExceptionInfo16);
            latentStyles1.Append(latentStyleExceptionInfo17);
            latentStyles1.Append(latentStyleExceptionInfo18);
            latentStyles1.Append(latentStyleExceptionInfo19);
            latentStyles1.Append(latentStyleExceptionInfo20);
            latentStyles1.Append(latentStyleExceptionInfo21);
            latentStyles1.Append(latentStyleExceptionInfo22);
            latentStyles1.Append(latentStyleExceptionInfo23);
            latentStyles1.Append(latentStyleExceptionInfo24);
            latentStyles1.Append(latentStyleExceptionInfo25);
            latentStyles1.Append(latentStyleExceptionInfo26);
            latentStyles1.Append(latentStyleExceptionInfo27);
            latentStyles1.Append(latentStyleExceptionInfo28);
            latentStyles1.Append(latentStyleExceptionInfo29);
            latentStyles1.Append(latentStyleExceptionInfo30);
            latentStyles1.Append(latentStyleExceptionInfo31);
            latentStyles1.Append(latentStyleExceptionInfo32);
            latentStyles1.Append(latentStyleExceptionInfo33);
            latentStyles1.Append(latentStyleExceptionInfo34);
            latentStyles1.Append(latentStyleExceptionInfo35);
            latentStyles1.Append(latentStyleExceptionInfo36);
            latentStyles1.Append(latentStyleExceptionInfo37);
            latentStyles1.Append(latentStyleExceptionInfo38);
            latentStyles1.Append(latentStyleExceptionInfo39);
            latentStyles1.Append(latentStyleExceptionInfo40);
            latentStyles1.Append(latentStyleExceptionInfo41);
            latentStyles1.Append(latentStyleExceptionInfo42);
            latentStyles1.Append(latentStyleExceptionInfo43);
            latentStyles1.Append(latentStyleExceptionInfo44);
            latentStyles1.Append(latentStyleExceptionInfo45);
            latentStyles1.Append(latentStyleExceptionInfo46);
            latentStyles1.Append(latentStyleExceptionInfo47);
            latentStyles1.Append(latentStyleExceptionInfo48);
            latentStyles1.Append(latentStyleExceptionInfo49);
            latentStyles1.Append(latentStyleExceptionInfo50);
            latentStyles1.Append(latentStyleExceptionInfo51);
            latentStyles1.Append(latentStyleExceptionInfo52);
            latentStyles1.Append(latentStyleExceptionInfo53);
            latentStyles1.Append(latentStyleExceptionInfo54);
            latentStyles1.Append(latentStyleExceptionInfo55);
            latentStyles1.Append(latentStyleExceptionInfo56);
            latentStyles1.Append(latentStyleExceptionInfo57);
            latentStyles1.Append(latentStyleExceptionInfo58);
            latentStyles1.Append(latentStyleExceptionInfo59);
            latentStyles1.Append(latentStyleExceptionInfo60);
            latentStyles1.Append(latentStyleExceptionInfo61);
            latentStyles1.Append(latentStyleExceptionInfo62);
            latentStyles1.Append(latentStyleExceptionInfo63);
            latentStyles1.Append(latentStyleExceptionInfo64);
            latentStyles1.Append(latentStyleExceptionInfo65);
            latentStyles1.Append(latentStyleExceptionInfo66);
            latentStyles1.Append(latentStyleExceptionInfo67);
            latentStyles1.Append(latentStyleExceptionInfo68);
            latentStyles1.Append(latentStyleExceptionInfo69);
            latentStyles1.Append(latentStyleExceptionInfo70);
            latentStyles1.Append(latentStyleExceptionInfo71);
            latentStyles1.Append(latentStyleExceptionInfo72);
            latentStyles1.Append(latentStyleExceptionInfo73);
            latentStyles1.Append(latentStyleExceptionInfo74);
            latentStyles1.Append(latentStyleExceptionInfo75);
            latentStyles1.Append(latentStyleExceptionInfo76);
            latentStyles1.Append(latentStyleExceptionInfo77);
            latentStyles1.Append(latentStyleExceptionInfo78);
            latentStyles1.Append(latentStyleExceptionInfo79);
            latentStyles1.Append(latentStyleExceptionInfo80);
            latentStyles1.Append(latentStyleExceptionInfo81);
            latentStyles1.Append(latentStyleExceptionInfo82);
            latentStyles1.Append(latentStyleExceptionInfo83);
            latentStyles1.Append(latentStyleExceptionInfo84);
            latentStyles1.Append(latentStyleExceptionInfo85);
            latentStyles1.Append(latentStyleExceptionInfo86);
            latentStyles1.Append(latentStyleExceptionInfo87);
            latentStyles1.Append(latentStyleExceptionInfo88);
            latentStyles1.Append(latentStyleExceptionInfo89);
            latentStyles1.Append(latentStyleExceptionInfo90);
            latentStyles1.Append(latentStyleExceptionInfo91);
            latentStyles1.Append(latentStyleExceptionInfo92);
            latentStyles1.Append(latentStyleExceptionInfo93);
            latentStyles1.Append(latentStyleExceptionInfo94);
            latentStyles1.Append(latentStyleExceptionInfo95);
            latentStyles1.Append(latentStyleExceptionInfo96);
            latentStyles1.Append(latentStyleExceptionInfo97);
            latentStyles1.Append(latentStyleExceptionInfo98);
            latentStyles1.Append(latentStyleExceptionInfo99);
            latentStyles1.Append(latentStyleExceptionInfo100);
            latentStyles1.Append(latentStyleExceptionInfo101);
            latentStyles1.Append(latentStyleExceptionInfo102);
            latentStyles1.Append(latentStyleExceptionInfo103);
            latentStyles1.Append(latentStyleExceptionInfo104);
            latentStyles1.Append(latentStyleExceptionInfo105);
            latentStyles1.Append(latentStyleExceptionInfo106);
            latentStyles1.Append(latentStyleExceptionInfo107);
            latentStyles1.Append(latentStyleExceptionInfo108);
            latentStyles1.Append(latentStyleExceptionInfo109);
            latentStyles1.Append(latentStyleExceptionInfo110);
            latentStyles1.Append(latentStyleExceptionInfo111);
            latentStyles1.Append(latentStyleExceptionInfo112);
            latentStyles1.Append(latentStyleExceptionInfo113);
            latentStyles1.Append(latentStyleExceptionInfo114);
            latentStyles1.Append(latentStyleExceptionInfo115);
            latentStyles1.Append(latentStyleExceptionInfo116);
            latentStyles1.Append(latentStyleExceptionInfo117);
            latentStyles1.Append(latentStyleExceptionInfo118);
            latentStyles1.Append(latentStyleExceptionInfo119);
            latentStyles1.Append(latentStyleExceptionInfo120);
            latentStyles1.Append(latentStyleExceptionInfo121);
            latentStyles1.Append(latentStyleExceptionInfo122);
            latentStyles1.Append(latentStyleExceptionInfo123);
            latentStyles1.Append(latentStyleExceptionInfo124);
            latentStyles1.Append(latentStyleExceptionInfo125);
            latentStyles1.Append(latentStyleExceptionInfo126);
            latentStyles1.Append(latentStyleExceptionInfo127);
            latentStyles1.Append(latentStyleExceptionInfo128);
            latentStyles1.Append(latentStyleExceptionInfo129);
            latentStyles1.Append(latentStyleExceptionInfo130);
            latentStyles1.Append(latentStyleExceptionInfo131);
            latentStyles1.Append(latentStyleExceptionInfo132);
            latentStyles1.Append(latentStyleExceptionInfo133);
            latentStyles1.Append(latentStyleExceptionInfo134);
            latentStyles1.Append(latentStyleExceptionInfo135);
            latentStyles1.Append(latentStyleExceptionInfo136);
            latentStyles1.Append(latentStyleExceptionInfo137);
            latentStyles1.Append(latentStyleExceptionInfo138);
            latentStyles1.Append(latentStyleExceptionInfo139);
            latentStyles1.Append(latentStyleExceptionInfo140);
            latentStyles1.Append(latentStyleExceptionInfo141);
            latentStyles1.Append(latentStyleExceptionInfo142);
            latentStyles1.Append(latentStyleExceptionInfo143);
            latentStyles1.Append(latentStyleExceptionInfo144);
            latentStyles1.Append(latentStyleExceptionInfo145);
            latentStyles1.Append(latentStyleExceptionInfo146);
            latentStyles1.Append(latentStyleExceptionInfo147);
            latentStyles1.Append(latentStyleExceptionInfo148);
            latentStyles1.Append(latentStyleExceptionInfo149);
            latentStyles1.Append(latentStyleExceptionInfo150);
            latentStyles1.Append(latentStyleExceptionInfo151);
            latentStyles1.Append(latentStyleExceptionInfo152);
            latentStyles1.Append(latentStyleExceptionInfo153);
            latentStyles1.Append(latentStyleExceptionInfo154);
            latentStyles1.Append(latentStyleExceptionInfo155);
            latentStyles1.Append(latentStyleExceptionInfo156);
            latentStyles1.Append(latentStyleExceptionInfo157);
            latentStyles1.Append(latentStyleExceptionInfo158);
            latentStyles1.Append(latentStyleExceptionInfo159);
            latentStyles1.Append(latentStyleExceptionInfo160);
            latentStyles1.Append(latentStyleExceptionInfo161);
            latentStyles1.Append(latentStyleExceptionInfo162);
            latentStyles1.Append(latentStyleExceptionInfo163);
            latentStyles1.Append(latentStyleExceptionInfo164);
            latentStyles1.Append(latentStyleExceptionInfo165);
            latentStyles1.Append(latentStyleExceptionInfo166);
            latentStyles1.Append(latentStyleExceptionInfo167);
            latentStyles1.Append(latentStyleExceptionInfo168);
            latentStyles1.Append(latentStyleExceptionInfo169);
            latentStyles1.Append(latentStyleExceptionInfo170);
            latentStyles1.Append(latentStyleExceptionInfo171);
            latentStyles1.Append(latentStyleExceptionInfo172);
            latentStyles1.Append(latentStyleExceptionInfo173);
            latentStyles1.Append(latentStyleExceptionInfo174);
            latentStyles1.Append(latentStyleExceptionInfo175);
            latentStyles1.Append(latentStyleExceptionInfo176);
            latentStyles1.Append(latentStyleExceptionInfo177);
            latentStyles1.Append(latentStyleExceptionInfo178);
            latentStyles1.Append(latentStyleExceptionInfo179);
            latentStyles1.Append(latentStyleExceptionInfo180);
            latentStyles1.Append(latentStyleExceptionInfo181);
            latentStyles1.Append(latentStyleExceptionInfo182);
            latentStyles1.Append(latentStyleExceptionInfo183);
            latentStyles1.Append(latentStyleExceptionInfo184);
            latentStyles1.Append(latentStyleExceptionInfo185);
            latentStyles1.Append(latentStyleExceptionInfo186);
            latentStyles1.Append(latentStyleExceptionInfo187);
            latentStyles1.Append(latentStyleExceptionInfo188);
            latentStyles1.Append(latentStyleExceptionInfo189);
            latentStyles1.Append(latentStyleExceptionInfo190);
            latentStyles1.Append(latentStyleExceptionInfo191);
            latentStyles1.Append(latentStyleExceptionInfo192);
            latentStyles1.Append(latentStyleExceptionInfo193);
            latentStyles1.Append(latentStyleExceptionInfo194);
            latentStyles1.Append(latentStyleExceptionInfo195);
            latentStyles1.Append(latentStyleExceptionInfo196);
            latentStyles1.Append(latentStyleExceptionInfo197);
            latentStyles1.Append(latentStyleExceptionInfo198);
            latentStyles1.Append(latentStyleExceptionInfo199);
            latentStyles1.Append(latentStyleExceptionInfo200);
            latentStyles1.Append(latentStyleExceptionInfo201);
            latentStyles1.Append(latentStyleExceptionInfo202);
            latentStyles1.Append(latentStyleExceptionInfo203);
            latentStyles1.Append(latentStyleExceptionInfo204);
            latentStyles1.Append(latentStyleExceptionInfo205);
            latentStyles1.Append(latentStyleExceptionInfo206);
            latentStyles1.Append(latentStyleExceptionInfo207);
            latentStyles1.Append(latentStyleExceptionInfo208);
            latentStyles1.Append(latentStyleExceptionInfo209);
            latentStyles1.Append(latentStyleExceptionInfo210);
            latentStyles1.Append(latentStyleExceptionInfo211);
            latentStyles1.Append(latentStyleExceptionInfo212);
            latentStyles1.Append(latentStyleExceptionInfo213);
            latentStyles1.Append(latentStyleExceptionInfo214);
            latentStyles1.Append(latentStyleExceptionInfo215);
            latentStyles1.Append(latentStyleExceptionInfo216);
            latentStyles1.Append(latentStyleExceptionInfo217);
            latentStyles1.Append(latentStyleExceptionInfo218);
            latentStyles1.Append(latentStyleExceptionInfo219);
            latentStyles1.Append(latentStyleExceptionInfo220);
            latentStyles1.Append(latentStyleExceptionInfo221);
            latentStyles1.Append(latentStyleExceptionInfo222);
            latentStyles1.Append(latentStyleExceptionInfo223);
            latentStyles1.Append(latentStyleExceptionInfo224);
            latentStyles1.Append(latentStyleExceptionInfo225);
            latentStyles1.Append(latentStyleExceptionInfo226);
            latentStyles1.Append(latentStyleExceptionInfo227);
            latentStyles1.Append(latentStyleExceptionInfo228);
            latentStyles1.Append(latentStyleExceptionInfo229);
            latentStyles1.Append(latentStyleExceptionInfo230);
            latentStyles1.Append(latentStyleExceptionInfo231);
            latentStyles1.Append(latentStyleExceptionInfo232);
            latentStyles1.Append(latentStyleExceptionInfo233);
            latentStyles1.Append(latentStyleExceptionInfo234);
            latentStyles1.Append(latentStyleExceptionInfo235);
            latentStyles1.Append(latentStyleExceptionInfo236);
            latentStyles1.Append(latentStyleExceptionInfo237);
            latentStyles1.Append(latentStyleExceptionInfo238);
            latentStyles1.Append(latentStyleExceptionInfo239);
            latentStyles1.Append(latentStyleExceptionInfo240);
            latentStyles1.Append(latentStyleExceptionInfo241);
            latentStyles1.Append(latentStyleExceptionInfo242);
            latentStyles1.Append(latentStyleExceptionInfo243);
            latentStyles1.Append(latentStyleExceptionInfo244);
            latentStyles1.Append(latentStyleExceptionInfo245);
            latentStyles1.Append(latentStyleExceptionInfo246);
            latentStyles1.Append(latentStyleExceptionInfo247);
            latentStyles1.Append(latentStyleExceptionInfo248);
            latentStyles1.Append(latentStyleExceptionInfo249);
            latentStyles1.Append(latentStyleExceptionInfo250);
            latentStyles1.Append(latentStyleExceptionInfo251);
            latentStyles1.Append(latentStyleExceptionInfo252);
            latentStyles1.Append(latentStyleExceptionInfo253);
            latentStyles1.Append(latentStyleExceptionInfo254);
            latentStyles1.Append(latentStyleExceptionInfo255);
            latentStyles1.Append(latentStyleExceptionInfo256);
            latentStyles1.Append(latentStyleExceptionInfo257);
            latentStyles1.Append(latentStyleExceptionInfo258);
            latentStyles1.Append(latentStyleExceptionInfo259);
            latentStyles1.Append(latentStyleExceptionInfo260);
            latentStyles1.Append(latentStyleExceptionInfo261);
            latentStyles1.Append(latentStyleExceptionInfo262);
            latentStyles1.Append(latentStyleExceptionInfo263);
            latentStyles1.Append(latentStyleExceptionInfo264);
            latentStyles1.Append(latentStyleExceptionInfo265);
            latentStyles1.Append(latentStyleExceptionInfo266);
            latentStyles1.Append(latentStyleExceptionInfo267);
            latentStyles1.Append(latentStyleExceptionInfo268);
            latentStyles1.Append(latentStyleExceptionInfo269);
            latentStyles1.Append(latentStyleExceptionInfo270);
            latentStyles1.Append(latentStyleExceptionInfo271);
            latentStyles1.Append(latentStyleExceptionInfo272);
            latentStyles1.Append(latentStyleExceptionInfo273);
            latentStyles1.Append(latentStyleExceptionInfo274);
            latentStyles1.Append(latentStyleExceptionInfo275);
            latentStyles1.Append(latentStyleExceptionInfo276);
            latentStyles1.Append(latentStyleExceptionInfo277);
            latentStyles1.Append(latentStyleExceptionInfo278);
            latentStyles1.Append(latentStyleExceptionInfo279);
            latentStyles1.Append(latentStyleExceptionInfo280);
            latentStyles1.Append(latentStyleExceptionInfo281);
            latentStyles1.Append(latentStyleExceptionInfo282);
            latentStyles1.Append(latentStyleExceptionInfo283);
            latentStyles1.Append(latentStyleExceptionInfo284);
            latentStyles1.Append(latentStyleExceptionInfo285);
            latentStyles1.Append(latentStyleExceptionInfo286);
            latentStyles1.Append(latentStyleExceptionInfo287);
            latentStyles1.Append(latentStyleExceptionInfo288);
            latentStyles1.Append(latentStyleExceptionInfo289);
            latentStyles1.Append(latentStyleExceptionInfo290);
            latentStyles1.Append(latentStyleExceptionInfo291);
            latentStyles1.Append(latentStyleExceptionInfo292);
            latentStyles1.Append(latentStyleExceptionInfo293);
            latentStyles1.Append(latentStyleExceptionInfo294);
            latentStyles1.Append(latentStyleExceptionInfo295);
            latentStyles1.Append(latentStyleExceptionInfo296);
            latentStyles1.Append(latentStyleExceptionInfo297);
            latentStyles1.Append(latentStyleExceptionInfo298);
            latentStyles1.Append(latentStyleExceptionInfo299);
            latentStyles1.Append(latentStyleExceptionInfo300);
            latentStyles1.Append(latentStyleExceptionInfo301);
            latentStyles1.Append(latentStyleExceptionInfo302);
            latentStyles1.Append(latentStyleExceptionInfo303);
            latentStyles1.Append(latentStyleExceptionInfo304);
            latentStyles1.Append(latentStyleExceptionInfo305);
            latentStyles1.Append(latentStyleExceptionInfo306);
            latentStyles1.Append(latentStyleExceptionInfo307);
            latentStyles1.Append(latentStyleExceptionInfo308);
            latentStyles1.Append(latentStyleExceptionInfo309);
            latentStyles1.Append(latentStyleExceptionInfo310);
            latentStyles1.Append(latentStyleExceptionInfo311);
            latentStyles1.Append(latentStyleExceptionInfo312);
            latentStyles1.Append(latentStyleExceptionInfo313);
            latentStyles1.Append(latentStyleExceptionInfo314);
            latentStyles1.Append(latentStyleExceptionInfo315);
            latentStyles1.Append(latentStyleExceptionInfo316);
            latentStyles1.Append(latentStyleExceptionInfo317);
            latentStyles1.Append(latentStyleExceptionInfo318);
            latentStyles1.Append(latentStyleExceptionInfo319);
            latentStyles1.Append(latentStyleExceptionInfo320);
            latentStyles1.Append(latentStyleExceptionInfo321);
            latentStyles1.Append(latentStyleExceptionInfo322);
            latentStyles1.Append(latentStyleExceptionInfo323);
            latentStyles1.Append(latentStyleExceptionInfo324);
            latentStyles1.Append(latentStyleExceptionInfo325);
            latentStyles1.Append(latentStyleExceptionInfo326);
            latentStyles1.Append(latentStyleExceptionInfo327);
            latentStyles1.Append(latentStyleExceptionInfo328);
            latentStyles1.Append(latentStyleExceptionInfo329);
            latentStyles1.Append(latentStyleExceptionInfo330);
            latentStyles1.Append(latentStyleExceptionInfo331);
            latentStyles1.Append(latentStyleExceptionInfo332);
            latentStyles1.Append(latentStyleExceptionInfo333);
            latentStyles1.Append(latentStyleExceptionInfo334);
            latentStyles1.Append(latentStyleExceptionInfo335);
            latentStyles1.Append(latentStyleExceptionInfo336);
            latentStyles1.Append(latentStyleExceptionInfo337);
            latentStyles1.Append(latentStyleExceptionInfo338);
            latentStyles1.Append(latentStyleExceptionInfo339);
            latentStyles1.Append(latentStyleExceptionInfo340);
            latentStyles1.Append(latentStyleExceptionInfo341);
            latentStyles1.Append(latentStyleExceptionInfo342);
            latentStyles1.Append(latentStyleExceptionInfo343);
            latentStyles1.Append(latentStyleExceptionInfo344);
            latentStyles1.Append(latentStyleExceptionInfo345);
            latentStyles1.Append(latentStyleExceptionInfo346);
            latentStyles1.Append(latentStyleExceptionInfo347);
            latentStyles1.Append(latentStyleExceptionInfo348);
            latentStyles1.Append(latentStyleExceptionInfo349);
            latentStyles1.Append(latentStyleExceptionInfo350);
            latentStyles1.Append(latentStyleExceptionInfo351);
            latentStyles1.Append(latentStyleExceptionInfo352);
            latentStyles1.Append(latentStyleExceptionInfo353);
            latentStyles1.Append(latentStyleExceptionInfo354);
            latentStyles1.Append(latentStyleExceptionInfo355);
            latentStyles1.Append(latentStyleExceptionInfo356);
            latentStyles1.Append(latentStyleExceptionInfo357);
            latentStyles1.Append(latentStyleExceptionInfo358);
            latentStyles1.Append(latentStyleExceptionInfo359);
            latentStyles1.Append(latentStyleExceptionInfo360);
            latentStyles1.Append(latentStyleExceptionInfo361);
            latentStyles1.Append(latentStyleExceptionInfo362);
            latentStyles1.Append(latentStyleExceptionInfo363);
            latentStyles1.Append(latentStyleExceptionInfo364);
            latentStyles1.Append(latentStyleExceptionInfo365);
            latentStyles1.Append(latentStyleExceptionInfo366);
            latentStyles1.Append(latentStyleExceptionInfo367);
            latentStyles1.Append(latentStyleExceptionInfo368);
            latentStyles1.Append(latentStyleExceptionInfo369);
            latentStyles1.Append(latentStyleExceptionInfo370);
            latentStyles1.Append(latentStyleExceptionInfo371);

            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true };
            StyleName styleName1 = new StyleName() { Val = "Normal" };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();
            Rsid rsid3 = new Rsid() { Val = "005F6638" };

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines12 = new SpacingBetweenLines() { After = "200", Line = "276", LineRule = LineSpacingRuleValues.Auto };

            styleParagraphProperties1.Append(spacingBetweenLines12);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            Languages languages17 = new Languages() { Val = "en-GB" };

            styleRunProperties1.Append(languages17);

            style1.Append(styleName1);
            style1.Append(primaryStyle1);
            style1.Append(rsid3);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);

            Style style2 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading1" };
            StyleName styleName2 = new StyleName() { Val = "heading 1" };
            BasedOn basedOn1 = new BasedOn() { Val = "Normal" };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "Normal" };
            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "Heading1Char" };
            PrimaryStyle primaryStyle2 = new PrimaryStyle();
            Rsid rsid4 = new Rsid() { Val = "005F6638" };

            StyleParagraphProperties styleParagraphProperties2 = new StyleParagraphProperties();
            KeepNext keepNext1 = new KeepNext();
            KeepLines keepLines1 = new KeepLines();
            SpacingBetweenLines spacingBetweenLines13 = new SpacingBetweenLines() { Before = "480", After = "0" };
            OutlineLevel outlineLevel1 = new OutlineLevel() { Val = 0 };

            styleParagraphProperties2.Append(keepNext1);
            styleParagraphProperties2.Append(keepLines1);
            styleParagraphProperties2.Append(spacingBetweenLines13);
            styleParagraphProperties2.Append(outlineLevel1);

            StyleRunProperties styleRunProperties2 = new StyleRunProperties();
            RunFonts runFonts52 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Bold bold26 = new Bold();
            BoldComplexScript boldComplexScript15 = new BoldComplexScript();
            Color color20 = new Color() { Val = "2E74B5", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };
            FontSize fontSize39 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript33 = new FontSizeComplexScript() { Val = "28" };

            styleRunProperties2.Append(runFonts52);
            styleRunProperties2.Append(bold26);
            styleRunProperties2.Append(boldComplexScript15);
            styleRunProperties2.Append(color20);
            styleRunProperties2.Append(fontSize39);
            styleRunProperties2.Append(fontSizeComplexScript33);

            style2.Append(styleName2);
            style2.Append(basedOn1);
            style2.Append(nextParagraphStyle1);
            style2.Append(linkedStyle1);
            style2.Append(primaryStyle2);
            style2.Append(rsid4);
            style2.Append(styleParagraphProperties2);
            style2.Append(styleRunProperties2);

            Style style3 = new Style() { Type = StyleValues.Character, StyleId = "DefaultParagraphFont", Default = true };
            StyleName styleName3 = new StyleName() { Val = "Default Paragraph Font" };
            UIPriority uIPriority1 = new UIPriority() { Val = 1 };
            SemiHidden semiHidden1 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();

            style3.Append(styleName3);
            style3.Append(uIPriority1);
            style3.Append(semiHidden1);
            style3.Append(unhideWhenUsed1);

            Style style4 = new Style() { Type = StyleValues.Table, StyleId = "TableNormal", Default = true };
            StyleName styleName4 = new StyleName() { Val = "Normal Table" };
            UIPriority uIPriority2 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden2 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed2 = new UnhideWhenUsed();

            StyleTableProperties styleTableProperties1 = new StyleTableProperties();
            TableIndentation tableIndentation1 = new TableIndentation() { Width = 0, Type = TableWidthUnitValues.Dxa };

            TableCellMarginDefault tableCellMarginDefault1 = new TableCellMarginDefault();
            TopMargin topMargin1 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellLeftMargin tableCellLeftMargin1 = new TableCellLeftMargin() { Width = 108, Type = TableWidthValues.Dxa };
            BottomMargin bottomMargin1 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellRightMargin tableCellRightMargin1 = new TableCellRightMargin() { Width = 108, Type = TableWidthValues.Dxa };

            tableCellMarginDefault1.Append(topMargin1);
            tableCellMarginDefault1.Append(tableCellLeftMargin1);
            tableCellMarginDefault1.Append(bottomMargin1);
            tableCellMarginDefault1.Append(tableCellRightMargin1);

            styleTableProperties1.Append(tableIndentation1);
            styleTableProperties1.Append(tableCellMarginDefault1);

            style4.Append(styleName4);
            style4.Append(uIPriority2);
            style4.Append(semiHidden2);
            style4.Append(unhideWhenUsed2);
            style4.Append(styleTableProperties1);

            Style style5 = new Style() { Type = StyleValues.Numbering, StyleId = "NoList", Default = true };
            StyleName styleName5 = new StyleName() { Val = "No List" };
            UIPriority uIPriority3 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden3 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed3 = new UnhideWhenUsed();

            style5.Append(styleName5);
            style5.Append(uIPriority3);
            style5.Append(semiHidden3);
            style5.Append(unhideWhenUsed3);

            Style style6 = new Style() { Type = StyleValues.Character, StyleId = "Heading1Char", CustomStyle = true };
            StyleName styleName6 = new StyleName() { Val = "Heading 1 Char" };
            BasedOn basedOn2 = new BasedOn() { Val = "DefaultParagraphFont" };
            LinkedStyle linkedStyle2 = new LinkedStyle() { Val = "Heading1" };
            Rsid rsid5 = new Rsid() { Val = "005F6638" };

            StyleRunProperties styleRunProperties3 = new StyleRunProperties();
            RunFonts runFonts53 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Bold bold27 = new Bold();
            BoldComplexScript boldComplexScript16 = new BoldComplexScript();
            Color color21 = new Color() { Val = "2E74B5", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };
            FontSize fontSize40 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript34 = new FontSizeComplexScript() { Val = "28" };
            Languages languages18 = new Languages() { Val = "en-GB" };

            styleRunProperties3.Append(runFonts53);
            styleRunProperties3.Append(bold27);
            styleRunProperties3.Append(boldComplexScript16);
            styleRunProperties3.Append(color21);
            styleRunProperties3.Append(fontSize40);
            styleRunProperties3.Append(fontSizeComplexScript34);
            styleRunProperties3.Append(languages18);

            style6.Append(styleName6);
            style6.Append(basedOn2);
            style6.Append(linkedStyle2);
            style6.Append(rsid5);
            style6.Append(styleRunProperties3);

            Style style7 = new Style() { Type = StyleValues.Character, StyleId = "BriefingtextChar", CustomStyle = true };
            StyleName styleName7 = new StyleName() { Val = "Briefing text Char" };
            LinkedStyle linkedStyle3 = new LinkedStyle() { Val = "Briefingtext" };
            UIPriority uIPriority4 = new UIPriority() { Val = 99 };
            PrimaryStyle primaryStyle3 = new PrimaryStyle();
            Locked locked1 = new Locked();
            Rsid rsid6 = new Rsid() { Val = "005F6638" };

            StyleRunProperties styleRunProperties4 = new StyleRunProperties();
            RunFonts runFonts54 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSizeComplexScript fontSizeComplexScript35 = new FontSizeComplexScript() { Val = "24" };

            styleRunProperties4.Append(runFonts54);
            styleRunProperties4.Append(fontSizeComplexScript35);

            style7.Append(styleName7);
            style7.Append(linkedStyle3);
            style7.Append(uIPriority4);
            style7.Append(primaryStyle3);
            style7.Append(locked1);
            style7.Append(rsid6);
            style7.Append(styleRunProperties4);

            Style style8 = new Style() { Type = StyleValues.Paragraph, StyleId = "Briefingtext", CustomStyle = true };
            StyleName styleName8 = new StyleName() { Val = "Briefing text" };
            BasedOn basedOn3 = new BasedOn() { Val = "Normal" };
            LinkedStyle linkedStyle4 = new LinkedStyle() { Val = "BriefingtextChar" };
            UIPriority uIPriority5 = new UIPriority() { Val = 99 };
            PrimaryStyle primaryStyle4 = new PrimaryStyle();
            Rsid rsid7 = new Rsid() { Val = "005F6638" };

            StyleParagraphProperties styleParagraphProperties3 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines14 = new SpacingBetweenLines() { After = "240", Line = "240", LineRule = LineSpacingRuleValues.Auto };
            Justification justification2 = new Justification() { Val = JustificationValues.Both };

            styleParagraphProperties3.Append(spacingBetweenLines14);
            styleParagraphProperties3.Append(justification2);

            StyleRunProperties styleRunProperties5 = new StyleRunProperties();
            RunFonts runFonts55 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSizeComplexScript fontSizeComplexScript36 = new FontSizeComplexScript() { Val = "24" };
            Languages languages19 = new Languages() { Val = "en-US" };

            styleRunProperties5.Append(runFonts55);
            styleRunProperties5.Append(fontSizeComplexScript36);
            styleRunProperties5.Append(languages19);

            style8.Append(styleName8);
            style8.Append(basedOn3);
            style8.Append(linkedStyle4);
            style8.Append(uIPriority5);
            style8.Append(primaryStyle4);
            style8.Append(rsid7);
            style8.Append(styleParagraphProperties3);
            style8.Append(styleRunProperties5);

            Style style9 = new Style() { Type = StyleValues.Paragraph, StyleId = "Briefingcontact", CustomStyle = true };
            StyleName styleName9 = new StyleName() { Val = "Briefing contact" };
            BasedOn basedOn4 = new BasedOn() { Val = "Briefingtext" };
            NextParagraphStyle nextParagraphStyle2 = new NextParagraphStyle() { Val = "Briefingtext" };
            UIPriority uIPriority6 = new UIPriority() { Val = 99 };
            Rsid rsid8 = new Rsid() { Val = "005F6638" };

            StyleParagraphProperties styleParagraphProperties4 = new StyleParagraphProperties();

            Tabs tabs1 = new Tabs();
            TabStop tabStop1 = new TabStop() { Val = TabStopValues.Left, Position = 1260 };

            tabs1.Append(tabStop1);
            SpacingBetweenLines spacingBetweenLines15 = new SpacingBetweenLines() { Before = "240" };
            Indentation indentation10 = new Indentation() { Start = "1259", Hanging = "1259" };
            Justification justification3 = new Justification() { Val = JustificationValues.Left };

            styleParagraphProperties4.Append(tabs1);
            styleParagraphProperties4.Append(spacingBetweenLines15);
            styleParagraphProperties4.Append(indentation10);
            styleParagraphProperties4.Append(justification3);

            StyleRunProperties styleRunProperties6 = new StyleRunProperties();
            Bold bold28 = new Bold();

            styleRunProperties6.Append(bold28);

            style9.Append(styleName9);
            style9.Append(basedOn4);
            style9.Append(nextParagraphStyle2);
            style9.Append(uIPriority6);
            style9.Append(rsid8);
            style9.Append(styleParagraphProperties4);
            style9.Append(styleRunProperties6);

            Style style10 = new Style() { Type = StyleValues.Character, StyleId = "CommentReference" };
            StyleName styleName10 = new StyleName() { Val = "annotation reference" };
            BasedOn basedOn5 = new BasedOn() { Val = "DefaultParagraphFont" };
            UIPriority uIPriority7 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden4 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed4 = new UnhideWhenUsed();
            Rsid rsid9 = new Rsid() { Val = "005F6638" };

            StyleRunProperties styleRunProperties7 = new StyleRunProperties();
            FontSize fontSize41 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript37 = new FontSizeComplexScript() { Val = "16" };

            styleRunProperties7.Append(fontSize41);
            styleRunProperties7.Append(fontSizeComplexScript37);

            style10.Append(styleName10);
            style10.Append(basedOn5);
            style10.Append(uIPriority7);
            style10.Append(semiHidden4);
            style10.Append(unhideWhenUsed4);
            style10.Append(rsid9);
            style10.Append(styleRunProperties7);

            Style style11 = new Style() { Type = StyleValues.Paragraph, StyleId = "CommentText" };
            StyleName styleName11 = new StyleName() { Val = "annotation text" };
            BasedOn basedOn6 = new BasedOn() { Val = "Normal" };
            LinkedStyle linkedStyle5 = new LinkedStyle() { Val = "CommentTextChar" };
            UIPriority uIPriority8 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden5 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed5 = new UnhideWhenUsed();
            Rsid rsid10 = new Rsid() { Val = "005F6638" };

            StyleParagraphProperties styleParagraphProperties5 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines16 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Auto };

            styleParagraphProperties5.Append(spacingBetweenLines16);

            StyleRunProperties styleRunProperties8 = new StyleRunProperties();
            FontSize fontSize42 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript38 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties8.Append(fontSize42);
            styleRunProperties8.Append(fontSizeComplexScript38);

            style11.Append(styleName11);
            style11.Append(basedOn6);
            style11.Append(linkedStyle5);
            style11.Append(uIPriority8);
            style11.Append(semiHidden5);
            style11.Append(unhideWhenUsed5);
            style11.Append(rsid10);
            style11.Append(styleParagraphProperties5);
            style11.Append(styleRunProperties8);

            Style style12 = new Style() { Type = StyleValues.Character, StyleId = "CommentTextChar", CustomStyle = true };
            StyleName styleName12 = new StyleName() { Val = "Comment Text Char" };
            BasedOn basedOn7 = new BasedOn() { Val = "DefaultParagraphFont" };
            LinkedStyle linkedStyle6 = new LinkedStyle() { Val = "CommentText" };
            UIPriority uIPriority9 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden6 = new SemiHidden();
            Rsid rsid11 = new Rsid() { Val = "005F6638" };

            StyleRunProperties styleRunProperties9 = new StyleRunProperties();
            FontSize fontSize43 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript39 = new FontSizeComplexScript() { Val = "20" };
            Languages languages20 = new Languages() { Val = "en-GB" };

            styleRunProperties9.Append(fontSize43);
            styleRunProperties9.Append(fontSizeComplexScript39);
            styleRunProperties9.Append(languages20);

            style12.Append(styleName12);
            style12.Append(basedOn7);
            style12.Append(linkedStyle6);
            style12.Append(uIPriority9);
            style12.Append(semiHidden6);
            style12.Append(rsid11);
            style12.Append(styleRunProperties9);

            Style style13 = new Style() { Type = StyleValues.Paragraph, StyleId = "BalloonText" };
            StyleName styleName13 = new StyleName() { Val = "Balloon Text" };
            BasedOn basedOn8 = new BasedOn() { Val = "Normal" };
            LinkedStyle linkedStyle7 = new LinkedStyle() { Val = "BalloonTextChar" };
            UIPriority uIPriority10 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden7 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed6 = new UnhideWhenUsed();
            Rsid rsid12 = new Rsid() { Val = "005F6638" };

            StyleParagraphProperties styleParagraphProperties6 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines17 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };

            styleParagraphProperties6.Append(spacingBetweenLines17);

            StyleRunProperties styleRunProperties10 = new StyleRunProperties();
            RunFonts runFonts56 = new RunFonts() { Ascii = "Segoe UI", HighAnsi = "Segoe UI", ComplexScript = "Segoe UI" };
            FontSize fontSize44 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript40 = new FontSizeComplexScript() { Val = "18" };

            styleRunProperties10.Append(runFonts56);
            styleRunProperties10.Append(fontSize44);
            styleRunProperties10.Append(fontSizeComplexScript40);

            style13.Append(styleName13);
            style13.Append(basedOn8);
            style13.Append(linkedStyle7);
            style13.Append(uIPriority10);
            style13.Append(semiHidden7);
            style13.Append(unhideWhenUsed6);
            style13.Append(rsid12);
            style13.Append(styleParagraphProperties6);
            style13.Append(styleRunProperties10);

            Style style14 = new Style() { Type = StyleValues.Character, StyleId = "BalloonTextChar", CustomStyle = true };
            StyleName styleName14 = new StyleName() { Val = "Balloon Text Char" };
            BasedOn basedOn9 = new BasedOn() { Val = "DefaultParagraphFont" };
            LinkedStyle linkedStyle8 = new LinkedStyle() { Val = "BalloonText" };
            UIPriority uIPriority11 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden8 = new SemiHidden();
            Rsid rsid13 = new Rsid() { Val = "005F6638" };

            StyleRunProperties styleRunProperties11 = new StyleRunProperties();
            RunFonts runFonts57 = new RunFonts() { Ascii = "Segoe UI", HighAnsi = "Segoe UI", ComplexScript = "Segoe UI" };
            FontSize fontSize45 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript41 = new FontSizeComplexScript() { Val = "18" };
            Languages languages21 = new Languages() { Val = "en-GB" };

            styleRunProperties11.Append(runFonts57);
            styleRunProperties11.Append(fontSize45);
            styleRunProperties11.Append(fontSizeComplexScript41);
            styleRunProperties11.Append(languages21);

            style14.Append(styleName14);
            style14.Append(basedOn9);
            style14.Append(linkedStyle8);
            style14.Append(uIPriority11);
            style14.Append(semiHidden8);
            style14.Append(rsid13);
            style14.Append(styleRunProperties11);

            styles1.Append(docDefaults1);
            styles1.Append(latentStyles1);
            styles1.Append(style1);
            styles1.Append(style2);
            styles1.Append(style3);
            styles1.Append(style4);
            styles1.Append(style5);
            styles1.Append(style6);
            styles1.Append(style7);
            styles1.Append(style8);
            styles1.Append(style9);
            styles1.Append(style10);
            styles1.Append(style11);
            styles1.Append(style12);
            styles1.Append(style13);
            styles1.Append(style14);

            styleDefinitionsPart1.Styles = styles1;
        }

        // Generates content of fontTablePart1.
        private void GenerateFontTablePart1Content(FontTablePart fontTablePart1)
        {
            Fonts fonts1 = new Fonts() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se" } };
            fonts1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            fonts1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            fonts1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            fonts1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            fonts1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            fonts1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");

            Font font1 = new Font() { Name = "Calibri" };
            Panose1Number panose1Number1 = new Panose1Number() { Val = "020F0502020204030204" };
            FontCharSet fontCharSet1 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily1 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch1 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature1 = new FontSignature() { UnicodeSignature0 = "E0002AFF", UnicodeSignature1 = "C000247B", UnicodeSignature2 = "00000009", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

            font1.Append(panose1Number1);
            font1.Append(fontCharSet1);
            font1.Append(fontFamily1);
            font1.Append(pitch1);
            font1.Append(fontSignature1);

            Font font2 = new Font() { Name = "Times New Roman" };
            Panose1Number panose1Number2 = new Panose1Number() { Val = "02020603050405020304" };
            FontCharSet fontCharSet2 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily2 = new FontFamily() { Val = FontFamilyValues.Roman };
            Pitch pitch2 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature2 = new FontSignature() { UnicodeSignature0 = "E0002EFF", UnicodeSignature1 = "C000785B", UnicodeSignature2 = "00000009", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

            font2.Append(panose1Number2);
            font2.Append(fontCharSet2);
            font2.Append(fontFamily2);
            font2.Append(pitch2);
            font2.Append(fontSignature2);

            Font font3 = new Font() { Name = "Calibri Light" };
            Panose1Number panose1Number3 = new Panose1Number() { Val = "020F0302020204030204" };
            FontCharSet fontCharSet3 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily3 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch3 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature3 = new FontSignature() { UnicodeSignature0 = "E0002AFF", UnicodeSignature1 = "C000247B", UnicodeSignature2 = "00000009", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

            font3.Append(panose1Number3);
            font3.Append(fontCharSet3);
            font3.Append(fontFamily3);
            font3.Append(pitch3);
            font3.Append(fontSignature3);

            Font font4 = new Font() { Name = "Arial" };
            Panose1Number panose1Number4 = new Panose1Number() { Val = "020B0604020202020204" };
            FontCharSet fontCharSet4 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily4 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch4 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature4 = new FontSignature() { UnicodeSignature0 = "E0002EFF", UnicodeSignature1 = "C0007843", UnicodeSignature2 = "00000009", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

            font4.Append(panose1Number4);
            font4.Append(fontCharSet4);
            font4.Append(fontFamily4);
            font4.Append(pitch4);
            font4.Append(fontSignature4);

            Font font5 = new Font() { Name = "Segoe UI" };
            Panose1Number panose1Number5 = new Panose1Number() { Val = "020B0502040204020203" };
            FontCharSet fontCharSet5 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily5 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch5 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature5 = new FontSignature() { UnicodeSignature0 = "E4002EFF", UnicodeSignature1 = "C000E47F", UnicodeSignature2 = "00000009", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

            font5.Append(panose1Number5);
            font5.Append(fontCharSet5);
            font5.Append(fontFamily5);
            font5.Append(pitch5);
            font5.Append(fontSignature5);

            Font font6 = new Font() { Name = "Arial,Times New Roman" };
            AltName altName1 = new AltName() { Val = "Times New Roman" };
            Panose1Number panose1Number6 = new Panose1Number() { Val = "00000000000000000000" };
            FontCharSet fontCharSet6 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily6 = new FontFamily() { Val = FontFamilyValues.Roman };
            NotTrueType notTrueType1 = new NotTrueType();
            Pitch pitch6 = new Pitch() { Val = FontPitchValues.Default };

            font6.Append(altName1);
            font6.Append(panose1Number6);
            font6.Append(fontCharSet6);
            font6.Append(fontFamily6);
            font6.Append(notTrueType1);
            font6.Append(pitch6);

            fonts1.Append(font1);
            fonts1.Append(font2);
            fonts1.Append(font3);
            fonts1.Append(font4);
            fonts1.Append(font5);
            fonts1.Append(font6);

            fontTablePart1.Fonts = fonts1;
        }

        // Generates content of wordprocessingCommentsExPart1.
        private void GenerateWordprocessingCommentsExPart1Content(WordprocessingCommentsExPart wordprocessingCommentsExPart1)
        {
            W15.CommentsEx commentsEx1 = new W15.CommentsEx() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se wp14" } };
            commentsEx1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            commentsEx1.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
            commentsEx1.AddNamespaceDeclaration("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
            commentsEx1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            commentsEx1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            commentsEx1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            commentsEx1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            commentsEx1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            commentsEx1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            commentsEx1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            commentsEx1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            commentsEx1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            commentsEx1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            commentsEx1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            commentsEx1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            commentsEx1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            commentsEx1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            commentsEx1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            commentsEx1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");
            W15.CommentEx commentEx1 = new W15.CommentEx() { ParaId = "732814B6", Done = false };
            W15.CommentEx commentEx2 = new W15.CommentEx() { ParaId = "63AF7470", Done = false };
            W15.CommentEx commentEx3 = new W15.CommentEx() { ParaId = "5D12E34C", Done = false };
            W15.CommentEx commentEx4 = new W15.CommentEx() { ParaId = "3A61E21A", Done = false };
            W15.CommentEx commentEx5 = new W15.CommentEx() { ParaId = "582BC19B", Done = false };
            W15.CommentEx commentEx6 = new W15.CommentEx() { ParaId = "5B3956BD", Done = false };
            W15.CommentEx commentEx7 = new W15.CommentEx() { ParaId = "41C181F4", Done = false };
            W15.CommentEx commentEx8 = new W15.CommentEx() { ParaId = "37A339B2", Done = false };
            W15.CommentEx commentEx9 = new W15.CommentEx() { ParaId = "0BE92328", Done = false };

            commentsEx1.Append(commentEx1);
            commentsEx1.Append(commentEx2);
            commentsEx1.Append(commentEx3);
            commentsEx1.Append(commentEx4);
            commentsEx1.Append(commentEx5);
            commentsEx1.Append(commentEx6);
            commentsEx1.Append(commentEx7);
            commentsEx1.Append(commentEx8);
            commentsEx1.Append(commentEx9);

            wordprocessingCommentsExPart1.CommentsEx = commentsEx1;
        }

        // Generates content of wordprocessingCommentsPart1.
        private void GenerateWordprocessingCommentsPart1Content(WordprocessingCommentsPart wordprocessingCommentsPart1)
        {
            Comments comments1 = new Comments() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se wp14" } };
            comments1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            comments1.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
            comments1.AddNamespaceDeclaration("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
            comments1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            comments1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            comments1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            comments1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            comments1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            comments1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            comments1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            comments1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            comments1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            comments1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            comments1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            comments1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            comments1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            comments1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            comments1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            comments1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            Comment comment1 = new Comment() { Initials = "LJR", Author = "JIMENEZ RUBIA Luis (SG-EXT)", Date = System.Xml.XmlConvert.ToDateTime("2019-10-29T10:43:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "0" };

            Paragraph paragraph32 = new Paragraph() { RsidParagraphAddition = "005F6638", RsidParagraphProperties = "005F6638", RsidRunAdditionDefault = "005F6638", ParagraphId = "732814B6", TextId = "77777777" };

            ParagraphProperties paragraphProperties32 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId21 = new ParagraphStyleId() { Val = "CommentText" };

            paragraphProperties32.Append(paragraphStyleId21);

            Run run85 = new Run();
            Text text74 = new Text();
            text74.Text = "[";

            run85.Append(text74);
            ProofError proofError9 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

            Run run86 = new Run();
            Text text75 = new Text();
            text75.Text = "C1  Member";

            run86.Append(text75);
            ProofError proofError10 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

            Run run87 = new Run();
            Text text76 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text76.Text = " State] + [C24 Lead DG] + ";

            run87.Append(text76);

            Run run88 = new Run();

            RunProperties runProperties63 = new RunProperties();
            RunStyle runStyle9 = new RunStyle() { Val = "CommentReference" };

            runProperties63.Append(runStyle9);
            AnnotationReferenceMark annotationReferenceMark1 = new AnnotationReferenceMark();

            run88.Append(runProperties63);
            run88.Append(annotationReferenceMark1);

            Run run89 = new Run();
            Text text77 = new Text();
            text77.Text = "[C2 Title] + [C18 Infringement Reference]";

            run89.Append(text77);

            paragraph32.Append(paragraphProperties32);
            paragraph32.Append(run85);
            paragraph32.Append(proofError9);
            paragraph32.Append(run86);
            paragraph32.Append(proofError10);
            paragraph32.Append(run87);
            paragraph32.Append(run88);
            paragraph32.Append(run89);

            comment1.Append(paragraph32);

            Comment comment2 = new Comment() { Initials = "LJR", Author = "JIMENEZ RUBIA Luis (SG-EXT)", Date = System.Xml.XmlConvert.ToDateTime("2019-10-31T11:59:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "1" };

            Paragraph paragraph33 = new Paragraph() { RsidParagraphAddition = "005F6638", RsidParagraphProperties = "005F6638", RsidRunAdditionDefault = "005F6638", ParagraphId = "63AF7470", TextId = "77777777" };

            ParagraphProperties paragraphProperties33 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId22 = new ParagraphStyleId() { Val = "CommentText" };

            paragraphProperties33.Append(paragraphStyleId22);

            Run run90 = new Run();

            RunProperties runProperties64 = new RunProperties();
            RunStyle runStyle10 = new RunStyle() { Val = "CommentReference" };

            runProperties64.Append(runStyle10);
            AnnotationReferenceMark annotationReferenceMark2 = new AnnotationReferenceMark();

            run90.Append(runProperties64);
            run90.Append(annotationReferenceMark2);

            Run run91 = new Run();
            Text text78 = new Text();
            text78.Text = "Suggestion: Since for a non-sensitive case there is no “Reason for sensitivity” would it make sense to have the Incriminated Fact?”";

            run91.Append(text78);

            paragraph33.Append(paragraphProperties33);
            paragraph33.Append(run90);
            paragraph33.Append(run91);

            comment2.Append(paragraph33);

            Comment comment3 = new Comment() { Initials = "LJR", Author = "JIMENEZ RUBIA Luis (SG-EXT)", Date = System.Xml.XmlConvert.ToDateTime("2019-11-11T16:41:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "2" };

            Paragraph paragraph34 = new Paragraph() { RsidParagraphAddition = "005F6638", RsidParagraphProperties = "005F6638", RsidRunAdditionDefault = "005F6638", ParagraphId = "5D12E34C", TextId = "77777777" };

            ParagraphProperties paragraphProperties34 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId23 = new ParagraphStyleId() { Val = "CommentText" };

            paragraphProperties34.Append(paragraphStyleId23);

            Run run92 = new Run();

            RunProperties runProperties65 = new RunProperties();
            RunStyle runStyle11 = new RunStyle() { Val = "CommentReference" };

            runProperties65.Append(runStyle11);
            AnnotationReferenceMark annotationReferenceMark3 = new AnnotationReferenceMark();

            run92.Append(runProperties65);
            run92.Append(annotationReferenceMark3);

            Run run93 = new Run();
            Text text79 = new Text();
            text79.Text = "Will not appear if the case were closed.";

            run93.Append(text79);

            paragraph34.Append(paragraphProperties34);
            paragraph34.Append(run92);
            paragraph34.Append(run93);

            comment3.Append(paragraph34);

            Comment comment4 = new Comment() { Initials = "LJR", Author = "JIMENEZ RUBIA Luis (SG-EXT)", Date = System.Xml.XmlConvert.ToDateTime("2019-10-29T11:41:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "3" };

            Paragraph paragraph35 = new Paragraph() { RsidParagraphAddition = "005F6638", RsidParagraphProperties = "005F6638", RsidRunAdditionDefault = "005F6638", ParagraphId = "38A6C90E", TextId = "77777777" };

            ParagraphProperties paragraphProperties35 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId24 = new ParagraphStyleId() { Val = "CommentText" };

            paragraphProperties35.Append(paragraphStyleId24);

            Run run94 = new Run();

            RunProperties runProperties66 = new RunProperties();
            RunStyle runStyle12 = new RunStyle() { Val = "CommentReference" };

            runProperties66.Append(runStyle12);
            AnnotationReferenceMark annotationReferenceMark4 = new AnnotationReferenceMark();

            run94.Append(runProperties66);
            run94.Append(annotationReferenceMark4);

            Run run95 = new Run();
            Text text80 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text80.Text = "Current ";

            run95.Append(text80);

            Run run96 = new Run();

            RunProperties runProperties67 = new RunProperties();
            RunStyle runStyle13 = new RunStyle() { Val = "CommentReference" };

            runProperties67.Append(runStyle13);
            AnnotationReferenceMark annotationReferenceMark5 = new AnnotationReferenceMark();

            run96.Append(runProperties67);
            run96.Append(annotationReferenceMark5);

            Run run97 = new Run();
            Text text81 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text81.Text = "[P1 ";

            run97.Append(text81);

            Run run98 = new Run() { RsidRunProperties = "00B40B82" };

            RunProperties runProperties68 = new RunProperties();
            RunFonts runFonts58 = new RunFonts() { ComplexScript = "Arial" };
            FontSize fontSize46 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript42 = new FontSizeComplexScript() { Val = "16" };

            runProperties68.Append(runFonts58);
            runProperties68.Append(fontSize46);
            runProperties68.Append(fontSizeComplexScript42);
            Text text82 = new Text();
            text82.Text = "Decision Type";

            run98.Append(runProperties68);
            run98.Append(text82);

            Run run99 = new Run();
            Text text83 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text83.Text = ", if any] + [P22 ";

            run99.Append(text83);

            Run run100 = new Run() { RsidRunProperties = "00B40B82" };

            RunProperties runProperties69 = new RunProperties();
            RunFonts runFonts59 = new RunFonts() { ComplexScript = "Arial" };
            FontSize fontSize47 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript43 = new FontSizeComplexScript() { Val = "16" };

            runProperties69.Append(runFonts59);
            runProperties69.Append(fontSize47);
            runProperties69.Append(fontSizeComplexScript43);
            Text text84 = new Text();
            text84.Text = "Decision Sent to the MS";

            run100.Append(runProperties69);
            run100.Append(text84);

            Run run101 = new Run();
            Text text85 = new Text();
            text85.Text = "]";

            run101.Append(text85);

            paragraph35.Append(paragraphProperties35);
            paragraph35.Append(run94);
            paragraph35.Append(run95);
            paragraph35.Append(run96);
            paragraph35.Append(run97);
            paragraph35.Append(run98);
            paragraph35.Append(run99);
            paragraph35.Append(run100);
            paragraph35.Append(run101);

            Paragraph paragraph36 = new Paragraph() { RsidParagraphAddition = "005F6638", RsidParagraphProperties = "005F6638", RsidRunAdditionDefault = "005F6638", ParagraphId = "717BEDE2", TextId = "77777777" };

            ParagraphProperties paragraphProperties36 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId25 = new ParagraphStyleId() { Val = "CommentText" };

            paragraphProperties36.Append(paragraphStyleId25);

            paragraph36.Append(paragraphProperties36);

            Paragraph paragraph37 = new Paragraph() { RsidParagraphAddition = "005F6638", RsidParagraphProperties = "005F6638", RsidRunAdditionDefault = "005F6638", ParagraphId = "3A61E21A", TextId = "77777777" };

            ParagraphProperties paragraphProperties37 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId26 = new ParagraphStyleId() { Val = "CommentText" };

            paragraphProperties37.Append(paragraphStyleId26);

            paragraph37.Append(paragraphProperties37);

            comment4.Append(paragraph35);
            comment4.Append(paragraph36);
            comment4.Append(paragraph37);

            Comment comment5 = new Comment() { Initials = "LJR", Author = "JIMENEZ RUBIA Luis (SG-EXT)", Date = System.Xml.XmlConvert.ToDateTime("2019-10-29T11:00:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "4" };

            Paragraph paragraph38 = new Paragraph() { RsidParagraphAddition = "005F6638", RsidParagraphProperties = "005F6638", RsidRunAdditionDefault = "005F6638", ParagraphId = "582BC19B", TextId = "77777777" };

            ParagraphProperties paragraphProperties38 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId27 = new ParagraphStyleId() { Val = "CommentText" };

            paragraphProperties38.Append(paragraphStyleId27);

            Run run102 = new Run();
            Text text86 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text86.Text = "Last ";

            run102.Append(text86);

            Run run103 = new Run();

            RunProperties runProperties70 = new RunProperties();
            RunStyle runStyle14 = new RunStyle() { Val = "CommentReference" };

            runProperties70.Append(runStyle14);
            AnnotationReferenceMark annotationReferenceMark6 = new AnnotationReferenceMark();

            run103.Append(runProperties70);
            run103.Append(annotationReferenceMark6);

            Run run104 = new Run();
            Text text87 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text87.Text = "[P1 ";

            run104.Append(text87);

            Run run105 = new Run() { RsidRunProperties = "00B40B82" };

            RunProperties runProperties71 = new RunProperties();
            RunFonts runFonts60 = new RunFonts() { ComplexScript = "Arial" };
            FontSize fontSize48 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript44 = new FontSizeComplexScript() { Val = "16" };

            runProperties71.Append(runFonts60);
            runProperties71.Append(fontSize48);
            runProperties71.Append(fontSizeComplexScript44);
            Text text88 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text88.Text = "Decision ";

            run105.Append(runProperties71);
            run105.Append(text88);
            ProofError proofError11 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

            Run run106 = new Run() { RsidRunProperties = "00B40B82" };

            RunProperties runProperties72 = new RunProperties();
            RunFonts runFonts61 = new RunFonts() { ComplexScript = "Arial" };
            FontSize fontSize49 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript45 = new FontSizeComplexScript() { Val = "16" };

            runProperties72.Append(runFonts61);
            runProperties72.Append(fontSize49);
            runProperties72.Append(fontSizeComplexScript45);
            Text text89 = new Text();
            text89.Text = "Type";

            run106.Append(runProperties72);
            run106.Append(text89);

            Run run107 = new Run();
            Text text90 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text90.Text = " ]";

            run107.Append(text90);
            ProofError proofError12 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

            Run run108 = new Run();
            Text text91 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text91.Text = " +[P22 ";

            run108.Append(text91);

            Run run109 = new Run() { RsidRunProperties = "00B40B82" };

            RunProperties runProperties73 = new RunProperties();
            RunFonts runFonts62 = new RunFonts() { ComplexScript = "Arial" };
            FontSize fontSize50 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript46 = new FontSizeComplexScript() { Val = "16" };

            runProperties73.Append(runFonts62);
            runProperties73.Append(fontSize50);
            runProperties73.Append(fontSizeComplexScript46);
            Text text92 = new Text();
            text92.Text = "Decision Sent to the MS";

            run109.Append(runProperties73);
            run109.Append(text92);

            Run run110 = new Run();
            Text text93 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text93.Text = " ]";

            run110.Append(text93);

            paragraph38.Append(paragraphProperties38);
            paragraph38.Append(run102);
            paragraph38.Append(run103);
            paragraph38.Append(run104);
            paragraph38.Append(run105);
            paragraph38.Append(proofError11);
            paragraph38.Append(run106);
            paragraph38.Append(run107);
            paragraph38.Append(proofError12);
            paragraph38.Append(run108);
            paragraph38.Append(run109);
            paragraph38.Append(run110);

            comment5.Append(paragraph38);

            Comment comment6 = new Comment() { Initials = "LJR", Author = "JIMENEZ RUBIA Luis (SG-EXT)", Date = System.Xml.XmlConvert.ToDateTime("2019-10-29T11:40:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "5" };

            Paragraph paragraph39 = new Paragraph() { RsidParagraphAddition = "005F6638", RsidParagraphProperties = "005F6638", RsidRunAdditionDefault = "005F6638", ParagraphId = "21AF400F", TextId = "77777777" };

            ParagraphProperties paragraphProperties39 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId28 = new ParagraphStyleId() { Val = "CommentText" };

            paragraphProperties39.Append(paragraphStyleId28);

            Run run111 = new Run();

            RunProperties runProperties74 = new RunProperties();
            RunStyle runStyle15 = new RunStyle() { Val = "CommentReference" };

            runProperties74.Append(runStyle15);
            AnnotationReferenceMark annotationReferenceMark7 = new AnnotationReferenceMark();

            run111.Append(runProperties74);
            run111.Append(annotationReferenceMark7);

            Run run112 = new Run();
            Text text94 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text94.Text = "Historical of ";

            run112.Append(text94);

            Run run113 = new Run();

            RunProperties runProperties75 = new RunProperties();
            RunStyle runStyle16 = new RunStyle() { Val = "CommentReference" };

            runProperties75.Append(runStyle16);
            AnnotationReferenceMark annotationReferenceMark8 = new AnnotationReferenceMark();

            run113.Append(runProperties75);
            run113.Append(annotationReferenceMark8);

            Run run114 = new Run();
            Text text95 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text95.Text = "[P1 ";

            run114.Append(text95);

            Run run115 = new Run() { RsidRunProperties = "00B40B82" };

            RunProperties runProperties76 = new RunProperties();
            RunFonts runFonts63 = new RunFonts() { ComplexScript = "Arial" };
            FontSize fontSize51 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript47 = new FontSizeComplexScript() { Val = "16" };

            runProperties76.Append(runFonts63);
            runProperties76.Append(fontSize51);
            runProperties76.Append(fontSizeComplexScript47);
            Text text96 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text96.Text = "Decision ";

            run115.Append(runProperties76);
            run115.Append(text96);
            ProofError proofError13 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

            Run run116 = new Run() { RsidRunProperties = "00B40B82" };

            RunProperties runProperties77 = new RunProperties();
            RunFonts runFonts64 = new RunFonts() { ComplexScript = "Arial" };
            FontSize fontSize52 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript48 = new FontSizeComplexScript() { Val = "16" };

            runProperties77.Append(runFonts64);
            runProperties77.Append(fontSize52);
            runProperties77.Append(fontSizeComplexScript48);
            Text text97 = new Text();
            text97.Text = "Type";

            run116.Append(runProperties77);
            run116.Append(text97);

            Run run117 = new Run();
            Text text98 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text98.Text = " ]";

            run117.Append(text98);
            ProofError proofError14 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

            Run run118 = new Run();
            Text text99 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text99.Text = " +[P22 ";

            run118.Append(text99);

            Run run119 = new Run() { RsidRunProperties = "00B40B82" };

            RunProperties runProperties78 = new RunProperties();
            RunFonts runFonts65 = new RunFonts() { ComplexScript = "Arial" };
            FontSize fontSize53 = new FontSize() { Val = "14" };
            FontSizeComplexScript fontSizeComplexScript49 = new FontSizeComplexScript() { Val = "16" };

            runProperties78.Append(runFonts65);
            runProperties78.Append(fontSize53);
            runProperties78.Append(fontSizeComplexScript49);
            Text text100 = new Text();
            text100.Text = "Decision Sent to the MS";

            run119.Append(runProperties78);
            run119.Append(text100);

            Run run120 = new Run();
            Text text101 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text101.Text = " ]";

            run120.Append(text101);

            paragraph39.Append(paragraphProperties39);
            paragraph39.Append(run111);
            paragraph39.Append(run112);
            paragraph39.Append(run113);
            paragraph39.Append(run114);
            paragraph39.Append(run115);
            paragraph39.Append(proofError13);
            paragraph39.Append(run116);
            paragraph39.Append(run117);
            paragraph39.Append(proofError14);
            paragraph39.Append(run118);
            paragraph39.Append(run119);
            paragraph39.Append(run120);

            Paragraph paragraph40 = new Paragraph() { RsidParagraphAddition = "005F6638", RsidParagraphProperties = "005F6638", RsidRunAdditionDefault = "005F6638", ParagraphId = "5B3956BD", TextId = "77777777" };

            ParagraphProperties paragraphProperties40 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId29 = new ParagraphStyleId() { Val = "CommentText" };

            paragraphProperties40.Append(paragraphStyleId29);

            paragraph40.Append(paragraphProperties40);

            comment6.Append(paragraph39);
            comment6.Append(paragraph40);

            Comment comment7 = new Comment() { Initials = "LJR", Author = "JIMENEZ RUBIA Luis (SG-EXT)", Date = System.Xml.XmlConvert.ToDateTime("2019-10-29T11:42:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "6" };

            Paragraph paragraph41 = new Paragraph() { RsidParagraphMarkRevision = "000F4EC3", RsidParagraphAddition = "005F6638", RsidParagraphProperties = "005F6638", RsidRunAdditionDefault = "005F6638", ParagraphId = "41C181F4", TextId = "77777777" };

            ParagraphProperties paragraphProperties41 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId30 = new ParagraphStyleId() { Val = "CommentText" };

            ParagraphMarkRunProperties paragraphMarkRunProperties24 = new ParagraphMarkRunProperties();
            Languages languages22 = new Languages() { Val = "fr-BE" };

            paragraphMarkRunProperties24.Append(languages22);

            paragraphProperties41.Append(paragraphStyleId30);
            paragraphProperties41.Append(paragraphMarkRunProperties24);

            Run run121 = new Run();

            RunProperties runProperties79 = new RunProperties();
            RunStyle runStyle17 = new RunStyle() { Val = "CommentReference" };

            runProperties79.Append(runStyle17);
            AnnotationReferenceMark annotationReferenceMark9 = new AnnotationReferenceMark();

            run121.Append(runProperties79);
            run121.Append(annotationReferenceMark9);

            Run run122 = new Run() { RsidRunProperties = "000F4EC3" };

            RunProperties runProperties80 = new RunProperties();
            Languages languages23 = new Languages() { Val = "fr-BE" };

            runProperties80.Append(languages23);
            Text text102 = new Text();
            text102.Text = "[ED1 Etat de Dossier";

            run122.Append(runProperties80);
            run122.Append(text102);

            Run run123 = new Run() { RsidRunProperties = "000F4EC3" };

            RunProperties runProperties81 = new RunProperties();
            RunFonts runFonts66 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Languages languages24 = new Languages() { Val = "fr-BE" };

            runProperties81.Append(runFonts66);
            runProperties81.Append(languages24);
            Text text103 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text103.Text = "] + ";

            run123.Append(runProperties81);
            run123.Append(text103);

            Run run124 = new Run();

            RunProperties runProperties82 = new RunProperties();
            RunFonts runFonts67 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Languages languages25 = new Languages() { Val = "fr-BE" };

            runProperties82.Append(runFonts67);
            runProperties82.Append(languages25);
            Text text104 = new Text();
            text104.Text = "[C28 Last Update Date]";

            run124.Append(runProperties82);
            run124.Append(text104);

            paragraph41.Append(paragraphProperties41);
            paragraph41.Append(run121);
            paragraph41.Append(run122);
            paragraph41.Append(run123);
            paragraph41.Append(run124);

            comment7.Append(paragraph41);

            Comment comment8 = new Comment() { Initials = "LJR", Author = "JIMENEZ RUBIA Luis (SG-EXT)", Date = System.Xml.XmlConvert.ToDateTime("2019-10-29T11:44:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "7" };

            Paragraph paragraph42 = new Paragraph() { RsidParagraphAddition = "005F6638", RsidParagraphProperties = "005F6638", RsidRunAdditionDefault = "005F6638", ParagraphId = "37A339B2", TextId = "77777777" };

            ParagraphProperties paragraphProperties42 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId31 = new ParagraphStyleId() { Val = "CommentText" };

            paragraphProperties42.Append(paragraphStyleId31);

            Run run125 = new Run();

            RunProperties runProperties83 = new RunProperties();
            RunStyle runStyle18 = new RunStyle() { Val = "CommentReference" };

            runProperties83.Append(runStyle18);
            AnnotationReferenceMark annotationReferenceMark10 = new AnnotationReferenceMark();

            run125.Append(runProperties83);
            run125.Append(annotationReferenceMark10);

            Run run126 = new Run();
            Text text105 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text105.Text = "[C9 ";

            run126.Append(text105);

            Run run127 = new Run() { RsidRunProperties = "00F6002E" };

            RunProperties runProperties84 = new RunProperties();
            RunFonts runFonts68 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };

            runProperties84.Append(runFonts68);
            Text text106 = new Text();
            text106.Text = "DG Case Handler";

            run127.Append(runProperties84);
            run127.Append(text106);

            Run run128 = new Run();

            RunProperties runProperties85 = new RunProperties();
            RunFonts runFonts69 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };

            runProperties85.Append(runFonts69);
            Text text107 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text107.Text = "] + ";

            run128.Append(runProperties85);
            run128.Append(text107);

            Run run129 = new Run();
            Text text108 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text108.Text = "[C28 Last Updated ";

            run129.Append(text108);
            ProofError proofError15 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

            Run run130 = new Run();
            Text text109 = new Text();
            text109.Text = "Date ]";

            run130.Append(text109);
            ProofError proofError16 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

            paragraph42.Append(paragraphProperties42);
            paragraph42.Append(run125);
            paragraph42.Append(run126);
            paragraph42.Append(run127);
            paragraph42.Append(run128);
            paragraph42.Append(run129);
            paragraph42.Append(proofError15);
            paragraph42.Append(run130);
            paragraph42.Append(proofError16);

            comment8.Append(paragraph42);

            Comment comment9 = new Comment() { Initials = "LJR", Author = "JIMENEZ RUBIA Luis (SG-EXT)", Date = System.Xml.XmlConvert.ToDateTime("2019-11-11T17:12:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "8" };

            Paragraph paragraph43 = new Paragraph() { RsidParagraphAddition = "005F6638", RsidParagraphProperties = "005F6638", RsidRunAdditionDefault = "005F6638", ParagraphId = "0BE92328", TextId = "77777777" };

            ParagraphProperties paragraphProperties43 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId32 = new ParagraphStyleId() { Val = "CommentText" };

            paragraphProperties43.Append(paragraphStyleId32);

            Run run131 = new Run();
            Text text110 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text110.Text = "NOTE: ";

            run131.Append(text110);

            Run run132 = new Run();

            RunProperties runProperties86 = new RunProperties();
            RunStyle runStyle19 = new RunStyle() { Val = "CommentReference" };

            runProperties86.Append(runStyle19);
            AnnotationReferenceMark annotationReferenceMark11 = new AnnotationReferenceMark();

            run132.Append(runProperties86);
            run132.Append(annotationReferenceMark11);

            Run run133 = new Run();
            Text text111 = new Text();
            text111.Text = "This information is empty in several of the legacy infringements.";

            run133.Append(text111);

            paragraph43.Append(paragraphProperties43);
            paragraph43.Append(run131);
            paragraph43.Append(run132);
            paragraph43.Append(run133);

            comment9.Append(paragraph43);

            comments1.Append(comment1);
            comments1.Append(comment2);
            comments1.Append(comment3);
            comments1.Append(comment4);
            comments1.Append(comment5);
            comments1.Append(comment6);
            comments1.Append(comment7);
            comments1.Append(comment8);
            comments1.Append(comment9);

            wordprocessingCommentsPart1.Comments = comments1;
        }

        private void SetPackageProperties(OpenXmlPackage document)
        {
            document.PackageProperties.Creator = "HOXHA Imir (SG-EXT)";
            document.PackageProperties.Title = "";
            document.PackageProperties.Subject = "";
            document.PackageProperties.Keywords = "";
            document.PackageProperties.Description = "";
            document.PackageProperties.Revision = "1";
            document.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2019-11-14T11:29:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2019-11-14T11:31:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.LastModifiedBy = "HOXHA Imir (SG-EXT)";
        }


    }
}
