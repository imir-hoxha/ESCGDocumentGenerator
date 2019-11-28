using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using M = DocumentFormat.OpenXml.Math;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using V = DocumentFormat.OpenXml.Vml;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using A = DocumentFormat.OpenXml.Drawing;
using Thm15 = DocumentFormat.OpenXml.Office2013.Theme;

namespace ConsoleApp1
{
    public class GeneratedClassA
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
            //ExtendedFilePropertiesPart extendedFilePropertiesPart1 = document.AddNewPart<ExtendedFilePropertiesPart>("rId3");
            //GenerateExtendedFilePropertiesPart1Content(extendedFilePropertiesPart1);

            MainDocumentPart mainDocumentPart1 = document.AddMainDocumentPart();
            GenerateMainDocumentPart1Content(mainDocumentPart1);

            //WebSettingsPart webSettingsPart1 = mainDocumentPart1.AddNewPart<WebSettingsPart>("rId3");
            //GenerateWebSettingsPart1Content(webSettingsPart1);

            //DocumentSettingsPart documentSettingsPart1 = mainDocumentPart1.AddNewPart<DocumentSettingsPart>("rId2");
            //GenerateDocumentSettingsPart1Content(documentSettingsPart1);

            //StyleDefinitionsPart styleDefinitionsPart1 = mainDocumentPart1.AddNewPart<StyleDefinitionsPart>("rId1");
            //GenerateStyleDefinitionsPart1Content(styleDefinitionsPart1);

            //ThemePart themePart1 = mainDocumentPart1.AddNewPart<ThemePart>("rId6");
            //GenerateThemePart1Content(themePart1);

            //GlossaryDocumentPart glossaryDocumentPart1 = mainDocumentPart1.AddNewPart<GlossaryDocumentPart>("rId5");
            //GenerateGlossaryDocumentPart1Content(glossaryDocumentPart1);

            //WebSettingsPart webSettingsPart2 = glossaryDocumentPart1.AddNewPart<WebSettingsPart>("rId3");
            //GenerateWebSettingsPart2Content(webSettingsPart2);

            //DocumentSettingsPart documentSettingsPart2 = glossaryDocumentPart1.AddNewPart<DocumentSettingsPart>("rId2");
            //GenerateDocumentSettingsPart2Content(documentSettingsPart2);

            //StyleDefinitionsPart styleDefinitionsPart2 = glossaryDocumentPart1.AddNewPart<StyleDefinitionsPart>("rId1");
            //GenerateStyleDefinitionsPart2Content(styleDefinitionsPart2);

            //FontTablePart fontTablePart1 = glossaryDocumentPart1.AddNewPart<FontTablePart>("rId4");
            //GenerateFontTablePart1Content(fontTablePart1);

            //FontTablePart fontTablePart2 = mainDocumentPart1.AddNewPart<FontTablePart>("rId4");
            //GenerateFontTablePart2Content(fontTablePart2);

            //SetPackageProperties(document);
        }

    

        // Generates content of mainDocumentPart1.
        private void GenerateMainDocumentPart1Content(MainDocumentPart mainDocumentPart1)
        {
            Document document1 = new Document() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se wp14" } };
            document1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            document1.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
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

            Paragraph paragraph1 = new Paragraph() { 
                RsidParagraphAddition = "00EC73B1", RsidRunAdditionDefault = "00EC73B1" };

            Run run1 = new Run();
            Text text1 = new Text();
            text1.Text = "Test 1";

            run1.Append(text1);

            paragraph1.Append(run1);

            SdtBlock sdtBlock1 = new SdtBlock();

            SdtProperties sdtProperties1 = new SdtProperties();
            Tag tag1 = new Tag() { Val = "testa" };
            SdtId sdtId1 = new SdtId() { Val = -1218056249 };

            SdtPlaceholder sdtPlaceholder1 = new SdtPlaceholder();
            DocPartReference docPartReference1 = new DocPartReference() { Val = "DefaultPlaceholder_-1854013440" };

            sdtPlaceholder1.Append(docPartReference1);

            sdtProperties1.Append(tag1);
            sdtProperties1.Append(sdtId1);
            sdtProperties1.Append(sdtPlaceholder1);
            SdtEndCharProperties sdtEndCharProperties1 = new SdtEndCharProperties();

            SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

            Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "001E13C4", RsidRunAdditionDefault = "001E13C4" };

            Run run2 = new Run();
            Text text2 = new Text();
            text2.Text = "Test a";

            run2.Append(text2);

            paragraph2.Append(run2);

            sdtContentBlock1.Append(paragraph2);

            sdtBlock1.Append(sdtProperties1);
            sdtBlock1.Append(sdtEndCharProperties1);
            sdtBlock1.Append(sdtContentBlock1);

            Paragraph paragraph3 = new Paragraph() { RsidParagraphAddition = "00EC73B1", RsidRunAdditionDefault = "00EC73B1" };

            Run run3 = new Run();
            Break break1 = new Break() { Type = BreakValues.Page };

            run3.Append(break1);

            paragraph3.Append(run3);

            Paragraph paragraph4 = new Paragraph() { RsidParagraphAddition = "00A15A83", RsidRunAdditionDefault = "00EC73B1" };

            Run run4 = new Run();
            LastRenderedPageBreak lastRenderedPageBreak1 = new LastRenderedPageBreak();
            Text text3 = new Text();
            text3.Text = "Test 2";

            run4.Append(lastRenderedPageBreak1);
            run4.Append(text3);

            paragraph4.Append(run4);
            BookmarkStart bookmarkStart1 = new BookmarkStart() { Name = "_GoBack", DisplacedByCustomXml = DisplacedByCustomXmlValues.Next, Id = "0" };

            SdtBlock sdtBlock2 = new SdtBlock();

            SdtProperties sdtProperties2 = new SdtProperties();
            Tag tag2 = new Tag() { Val = "testb" };
            SdtId sdtId2 = new SdtId() { Val = -285745073 };

            SdtPlaceholder sdtPlaceholder2 = new SdtPlaceholder();
            DocPartReference docPartReference2 = new DocPartReference() { Val = "03E0ACA1F9B64F42892D8CC68D472C5B" };

            sdtPlaceholder2.Append(docPartReference2);

            sdtProperties2.Append(tag2);
            sdtProperties2.Append(sdtId2);
            sdtProperties2.Append(sdtPlaceholder2);
            SdtEndCharProperties sdtEndCharProperties2 = new SdtEndCharProperties();

            SdtContentBlock sdtContentBlock2 = new SdtContentBlock();

            Paragraph paragraph5 = new Paragraph() { RsidParagraphAddition = "001E13C4", RsidParagraphProperties = "001E13C4", RsidRunAdditionDefault = "001E13C4" };

            Run run5 = new Run();
            Text text4 = new Text();
            text4.Text = "Test b";

            run5.Append(text4);

            paragraph5.Append(run5);

            sdtContentBlock2.Append(paragraph5);

            sdtBlock2.Append(sdtProperties2);
            sdtBlock2.Append(sdtEndCharProperties2);
            sdtBlock2.Append(sdtContentBlock2);
            BookmarkEnd bookmarkEnd1 = new BookmarkEnd() { DisplacedByCustomXml = DisplacedByCustomXmlValues.Previous, Id = "0" };
            Paragraph paragraph6 = new Paragraph() { RsidParagraphAddition = "001E13C4", RsidRunAdditionDefault = "001E13C4" };

            SectionProperties sectionProperties1 = new SectionProperties() { RsidR = "001E13C4" };
            PageSize pageSize1 = new PageSize() { Width = (UInt32Value)11906U, Height = (UInt32Value)16838U };
            PageMargin pageMargin1 = new PageMargin() { Top = 1440, Right = (UInt32Value)1440U, Bottom = 1440, Left = (UInt32Value)1440U, Header = (UInt32Value)708U, Footer = (UInt32Value)708U, Gutter = (UInt32Value)0U };
            Columns columns1 = new Columns() { Space = "708" };
            DocGrid docGrid1 = new DocGrid() { LinePitch = 360 };

            sectionProperties1.Append(pageSize1);
            sectionProperties1.Append(pageMargin1);
            sectionProperties1.Append(columns1);
            sectionProperties1.Append(docGrid1);

            body1.Append(paragraph1);
            body1.Append(sdtBlock1);
            body1.Append(paragraph3);
            body1.Append(paragraph4);
            body1.Append(bookmarkStart1);
            body1.Append(sdtBlock2);
            body1.Append(bookmarkEnd1);
            body1.Append(paragraph6);
            body1.Append(sectionProperties1);

            document1.Append(body1);

            mainDocumentPart1.Document = document1;
        }



   
   
     

    }
}

