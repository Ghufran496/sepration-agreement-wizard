using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentGenerationApi.Models;
using System.Text;

namespace DocumentGenerationApi.Services
{
    public class DocumentGenerationService
    {
        // Color definitions
        private static string MainHeadingColor = "000000"; // Blue
        private static string SubHeadingColor = "000000";  // Lighter blue
        private static string AccentColor = "000000";      // Orange for accents
        private static string BorderColor = "4472C4";      // Border blue

        public byte[] GenerateSeparationAgreement(DocumentRequest request)
        {
            using (MemoryStream memoryStream = new MemoryStream())
            {
                using (WordprocessingDocument wordDocument = 
                    WordprocessingDocument.Create(memoryStream, WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                        mainPart.Document = new Document();
                    Body body = mainPart.Document.AppendChild(new Body());

                    SetDocumentSettings(wordDocument);
                    
                    AddDifferentFirstPageHeader(wordDocument);

                    AddHeading(body, "SEPARATION AGREEMENT", 1, true, true);
                    AddSpacing(body, 24);

                    string party1Name = $"{request.PartyInfo.Party1FirstName} {request.PartyInfo.Party1MiddleName} {request.PartyInfo.Party1LastName}".Trim();
                    string party2Name = $"{request.PartyInfo.Party2FirstName} {request.PartyInfo.Party2MiddleName} {request.PartyInfo.Party2LastName}".Trim();
                    
                    AddParagraphWithAllCaps(body, "THIS IS A SEPARATION AGREEMENT DATED ___________ DATE OF SIGNING", false);
                    AddSpacing(body, 12);
                    
                    AddParagraphWithSpacing(body, "Between", false);
                    AddSpacing(body, 12);
                    
                    string party1FullName = "*PARTY ONE FULL NAME*";
                    string party1Label = "(*PARTY 1*)";
                    party1FullName = FormatClauseText(party1FullName, request.PartyInfo);
                    
                    AddParagraphWithSpacing(body, party1FullName, false, false, JustificationValues.Center);
                    AddParagraphWithSpacing(body, party1Label, false, false, JustificationValues.Right);
                    AddSpacing(body, 12);
                    
                    AddParagraphWithSpacing(body, "AND", false, false, JustificationValues.Center);
                    AddSpacing(body, 12);
                    
                    string party2FullName = "*PARTY TWO FULL NAME*";
                    string party2Label = "(*PARTY 2*)";
                    party2FullName = FormatClauseText(party2FullName, request.PartyInfo);
                    
                    AddParagraphWithSpacing(body, party2FullName, false, false, JustificationValues.Center);
                    AddParagraphWithSpacing(body, party2Label, false, false, JustificationValues.Right);
                    AddSpacing(body, 24);
                    
                    var categorizedClauses = new Dictionary<string, List<Clause>>();
                    foreach (var clause in request.SelectedClauses)
                    {
                        if (!categorizedClauses.ContainsKey(clause.Category))
                        {
                            categorizedClauses[clause.Category] = new List<Clause>();
                        }
                        categorizedClauses[clause.Category].Add(clause);
                    }
                    
                    int categoryCounter = 0;
                    
                    foreach (var category in categorizedClauses.Keys)
                    {
                        categoryCounter++;
                        
                        var clausesInCategory = categorizedClauses[category];
                        
                        bool isSchedulesCategory = category.Equals("Schedules", StringComparison.OrdinalIgnoreCase);
                        bool needsPageBreak = false;
                        
                        if (isSchedulesCategory)
                        {
                            foreach (var scheduleClause in clausesInCategory)
                            {
                                if (scheduleClause.Text.Contains("===PAGE BREAK==="))
                                {
                                    needsPageBreak = true;
                                    break;
                                }
                            }
                            
                            if (needsPageBreak)
                            {
                                AddPageBreak(body);
                            }
                        }
                        
                        if (clausesInCategory.Count > 1)
                        {
                            if (isSchedulesCategory)
                            {
                                AddHeading(body, "Schedules", 3, true);
                            }
                            else
                            {
                                string categoryHeading = $"{categoryCounter}.         {category}";
                                AddHeading(body, categoryHeading, 3);
                            }
                        }
                        
                        for (int i = 0; i < clausesInCategory.Count; i++)
                        {
                            var clause = clausesInCategory[i];
                            
                            bool containsPageBreak = clause.Text.Contains("===PAGE BREAK===");
                            string processedText = clause.Text;
                            
                            if (containsPageBreak)
                            {
                                processedText = processedText.Replace("===PAGE BREAK===", "");
                            }
                            
                            string formattedId;
                            if (clausesInCategory.Count == 1)
                            {
                                formattedId = categoryCounter.ToString();
                                
                                string headingWithId;
                                
                                if (isSchedulesCategory)
                                {
                                    headingWithId = "Schedules";
                                    AddHeading(body, headingWithId, 3, true);
                                }
                                else
                                {
                                    headingWithId = $"{formattedId}.         {clause.Category}";
                                    AddHeading(body, headingWithId, 3);
                                }
                                
                                string formattedText = FormatClauseText(processedText, request.PartyInfo);
                                
                                if (containsPageBreak)
                                {
                                    if (clause.Category.Equals("Execution", StringComparison.OrdinalIgnoreCase))
                                    {
                                        HandleTextWithPageBreaks(body, formattedText, true, false);
                                    }
                                    else if (!isSchedulesCategory)
                                    {
                                        if (processedText.TrimStart().StartsWith("===PAGE BREAK==="))
                                        {
                                            AddPageBreak(body);
                                        }
                                        
                                        HandleTextWithPageBreaks(body, formattedText, true);
                                    }
                                    else
                                    {
                                        AddParagraphWithSpacing(body, formattedText, true);
                                    }
                                }
                                else
                                {
                                    AddParagraphWithSpacing(body, formattedText, true);
                                }
                            }
                            else
                            {
                                formattedId = isSchedulesCategory ? "" : $"{categoryCounter}.{i+1}";
                                
                                string formattedText = FormatClauseText(processedText, request.PartyInfo);
                                
                                if (isSchedulesCategory)
                                {
                                    if (i > 0 || !needsPageBreak)
                                    {
                                        AddPageBreak(body);
                                    }
                                    
                                    string[] lines = formattedText.Split(new[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
                                    
                                    if (lines.Length > 0)
                                    {
                                        int titleIndex = -1;
                                        for (int j = 0; j < lines.Length; j++)
                                        {
                                            if (lines[j].Trim().StartsWith("SCHEDULE"))
                                            {
                                                titleIndex = j;
                                                break;
                                            }
                                        }
                                                
                                        if (titleIndex >= 0)
                                        {
                                            string scheduleTitle = lines[titleIndex];
                                            
                                            AddScheduleTitle(body, scheduleTitle, true, true);
                                            
                                            for (int j = 0; j < lines.Length; j++)
                                            {
                                                if (j != titleIndex && !string.IsNullOrWhiteSpace(lines[j]))
                                                {
                                                    AddSingleParagraphWithSpacing(body, lines[j], true);
                                                }
                                            }
                                        }
                                        else
                                        {
                                            AddParagraphWithSpacing(body, formattedText, true);
                                        }
                                    }
                                    else
                                    {
                                        AddParagraphWithSpacing(body, "", true);
                                    }
                                }
                                else
                                {
                                    if (containsPageBreak && !isSchedulesCategory)
                                    {
                                        HandleTextWithPageBreaks(body, $"\t{formattedId}\t{formattedText}", true);
                                    }
                                    else
                                    {
                                        AddParagraphWithSpacing(body, $"\t{formattedId}\t{formattedText}", true);
                                    }
                                }
                            }
                        }
                    }
                    
                    bool executionCategoryExists = categorizedClauses.ContainsKey("Execution");
                    
                    if (!executionCategoryExists)
                    {
                        AddHeading(body, "EXECUTION", 2);
                        AddSpacing(body, 12);
                        AddParagraphWithSpacing(body, "IN WITNESS WHEREOF, the parties have executed this Agreement as of the date first written above.", true);
                        AddSpacing(body, 36);
                        
                        AddSignatureLine(body, party1Name);
                        AddSpacing(body, 48);
                        
                        AddSignatureLine(body, party2Name);
                        AddSpacing(body, 36);
                    }
                    else
                    {
                        AddSpacing(body, 36);
                    }
                    
                    AddHorizontalLine(body);
                    AddSpacing(body, 12);
                    AddFooter(body, $"This document was generated using Separation Agreement Wizard on {DateTime.Now:MMMM d, yyyy}");
                }

                return memoryStream.ToArray();
            }
        }
        
        private void AddDifferentFirstPageHeader(WordprocessingDocument wordDocument)
        {
            MainDocumentPart mainPart = wordDocument.MainDocumentPart;
            
            HeaderPart defaultHeaderPart = mainPart.AddNewPart<HeaderPart>();
            
            Header defaultHeader = new Header();
            
            Table headerTable = new Table();
            
            TableProperties tableProperties = new TableProperties();
            TableWidth tableWidth = new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct };
            TableBorders tableBorders = new TableBorders();
            tableBorders.AppendChild(new TopBorder() { Val = BorderValues.None });
            tableBorders.AppendChild(new LeftBorder() { Val = BorderValues.None });
            tableBorders.AppendChild(new BottomBorder() { Val = BorderValues.None });
            tableBorders.AppendChild(new RightBorder() { Val = BorderValues.None });
            tableBorders.AppendChild(new InsideHorizontalBorder() { Val = BorderValues.None });
            tableBorders.AppendChild(new InsideVerticalBorder() { Val = BorderValues.None });
            
            tableProperties.AppendChild(tableWidth);
            tableProperties.AppendChild(tableBorders);
            headerTable.AppendChild(tableProperties);
            
            TableRow row = new TableRow();
            
            TableCell leftCell = new TableCell();
            Paragraph leftPara = new Paragraph();
            Run leftRun = new Run();
            RunProperties leftRunProperties = new RunProperties();
            leftRunProperties.AppendChild(new FontSize() { Val = "22" });
            leftRunProperties.AppendChild(new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri" });
            leftRun.AppendChild(leftRunProperties);
            leftRun.AppendChild(new Text("Separation Agreement"));
            leftPara.AppendChild(leftRun);
            leftCell.AppendChild(leftPara);
            row.AppendChild(leftCell);
            
            TableCell rightCell = new TableCell();
            Paragraph rightPara = new Paragraph();
            ParagraphProperties rightParaProps = new ParagraphProperties();
            rightParaProps.AppendChild(new Justification() { Val = JustificationValues.Right });
            rightPara.AppendChild(rightParaProps);
            
            Run rightRun = new Run();
            RunProperties rightRunProperties = new RunProperties();
            rightRunProperties.AppendChild(new FontSize() { Val = "22" });
            rightRunProperties.AppendChild(new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri" });
            rightRun.AppendChild(rightRunProperties);
            rightRun.AppendChild(new Text("Page "));
            rightPara.AppendChild(rightRun);
            
            Run pageFieldRun = new Run();
            pageFieldRun.AppendChild(new RunProperties(rightRunProperties.OuterXml));
            pageFieldRun.AppendChild(new FieldChar() { FieldCharType = FieldCharValues.Begin });
            rightPara.AppendChild(pageFieldRun);
            
            Run pageFieldCodeRun = new Run();
            pageFieldCodeRun.AppendChild(new RunProperties(rightRunProperties.OuterXml));
            pageFieldCodeRun.AppendChild(new FieldCode(" PAGE "));
            rightPara.AppendChild(pageFieldCodeRun);
            
            Run pageFieldEndRun = new Run();
            pageFieldEndRun.AppendChild(new RunProperties(rightRunProperties.OuterXml));
            pageFieldEndRun.AppendChild(new FieldChar() { FieldCharType = FieldCharValues.End });
            rightPara.AppendChild(pageFieldEndRun);
            
            rightCell.AppendChild(rightPara);
            row.AppendChild(rightCell);
            
            headerTable.AppendChild(row);
            
            TableRow lineRow = new TableRow();
            
            TableCell lineCell = new TableCell();
            
            TableCellProperties cellProperties = new TableCellProperties();
            GridSpan gridSpan = new GridSpan() { Val = 2 }; // Span 2 columns
            cellProperties.AppendChild(gridSpan);
            
            TableCellBorders cellBorders = new TableCellBorders();
            cellBorders.AppendChild(new BottomBorder() 
            { 
                Val = BorderValues.Single,
                Size = 4,
                Color = "000000"  // Black color instead of BorderColor
            });
            cellProperties.AppendChild(cellBorders);
            lineCell.AppendChild(cellProperties);
            
            Paragraph linePara = new Paragraph();
            lineCell.AppendChild(linePara);
            
            lineRow.AppendChild(lineCell);
            
            headerTable.AppendChild(lineRow);
            
            defaultHeader.AppendChild(headerTable);
            
            Paragraph spacingParagraph = new Paragraph();
            ParagraphProperties spacingProperties = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines = new SpacingBetweenLines() 
            { 
                After = "240",  // 0.17 inch spacing (240 twips) - reduced from 720
            };
            spacingProperties.AppendChild(spacingBetweenLines);
            spacingParagraph.AppendChild(spacingProperties);
            defaultHeader.AppendChild(spacingParagraph);
            
            defaultHeaderPart.Header = defaultHeader;
            
            HeaderPart firstPageHeaderPart = mainPart.AddNewPart<HeaderPart>();
            Header firstPageHeader = new Header();
            firstPageHeader.AppendChild(new Paragraph());
            firstPageHeaderPart.Header = firstPageHeader;
            
            SectionProperties sectionProps = mainPart.Document.Body.Elements<SectionProperties>().FirstOrDefault();
            if (sectionProps == null)
            {
                sectionProps = new SectionProperties();
                mainPart.Document.Body.AppendChild(sectionProps);
            }
            
            var existingHeaderRefs = sectionProps.Elements<HeaderReference>().ToList();
            foreach (var headerRef in existingHeaderRefs)
            {
                headerRef.Remove();
            }
            
            HeaderReference firstPageHeaderReference = new HeaderReference() 
            { 
                Type = HeaderFooterValues.First,
                Id = mainPart.GetIdOfPart(firstPageHeaderPart)
            };
            
            HeaderReference defaultHeaderReference = new HeaderReference() 
            { 
                Type = HeaderFooterValues.Default,
                Id = mainPart.GetIdOfPart(defaultHeaderPart)
            };
            
            sectionProps.AppendChild(firstPageHeaderReference);
            sectionProps.AppendChild(defaultHeaderReference);
            
            var existingTitlePgElements = sectionProps.Elements<TitlePage>().ToList();
            foreach (var titlePgElement in existingTitlePgElements)
            {
                titlePgElement.Remove();
            }
            
            TitlePage titlePage = new TitlePage();
            sectionProps.AppendChild(titlePage);
        }

        private void SetDocumentSettings(WordprocessingDocument document)
        {
            var sectionProps = new SectionProperties();
            var pageMargin = new PageMargin
            {
                Top = 720,    // 0.5 inch
                Right = 1440, // 1.0 inch - increased for better spacing
                Bottom = 720, // 0.5 inch
                Left = 1440,  // 1.0 inch - increased for better spacing
                Header = 720, // 0.5 inch for header
                Footer = 720, // 0.5 inch for footer
                Gutter = 0
            };
            
            sectionProps.AppendChild(pageMargin);
            
            var docDefaults = new DocDefaults();
            var runPropertiesDefault = new RunPropertiesDefault();
            var runPropertiesBaseStyle = new RunPropertiesBaseStyle();
            
            // Set Calibri as the default font
            var runFonts = new RunFonts() 
            { 
                Ascii = "Calibri", 
                HighAnsi = "Calibri", 
                ComplexScript = "Calibri" 
            };
            
            runPropertiesBaseStyle.AppendChild(runFonts);
            runPropertiesDefault.AppendChild(runPropertiesBaseStyle);
            docDefaults.AppendChild(runPropertiesDefault);
            
            // Set default paragraph properties with 1.15 line spacing
            var paragraphPropertiesDefault = new ParagraphPropertiesDefault();
            var paragraphPropertiesBaseStyle = new ParagraphPropertiesBaseStyle();
            var defaultSpacing = new SpacingBetweenLines() 
            { 
                LineRule = LineSpacingRuleValues.Auto,
                Line = "276", // 1.15 line spacing (240 * 1.15)
            };
            paragraphPropertiesBaseStyle.AppendChild(defaultSpacing);
            paragraphPropertiesDefault.AppendChild(paragraphPropertiesBaseStyle);
            docDefaults.AppendChild(paragraphPropertiesDefault);
            
            // Apply settings
            document.MainDocumentPart.Document.Body.AppendChild(sectionProps);
            
            // If there's a style part, add the default font setting
            StyleDefinitionsPart stylesPart = document.MainDocumentPart.StyleDefinitionsPart;
            if (stylesPart == null)
            {
                stylesPart = document.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
                stylesPart.Styles = new Styles();
                stylesPart.Styles.AppendChild(docDefaults);
                stylesPart.Styles.Save();
            }
            else if (stylesPart.Styles.DocDefaults == null)
            {
                stylesPart.Styles.InsertAt(docDefaults, 0);
                stylesPart.Styles.Save();
            }
        }

        private void AddHorizontalLine(Body body)
        {
            var para = body.AppendChild(new Paragraph());
            
            ParagraphProperties paragraphProperties = new ParagraphProperties();
            ParagraphBorders borders = new ParagraphBorders();
            
            TopBorder topBorder = new TopBorder()
            {
                Val = new EnumValue<BorderValues>(BorderValues.Single),
                Size = 12,
                Color = BorderColor
            };
            
            borders.AppendChild(topBorder);
            paragraphProperties.AppendChild(borders);
            para.AppendChild(paragraphProperties);
        }
        
        private void AddSpacing(Body body, int spacing)
        {
            Paragraph para = body.AppendChild(new Paragraph());
            ParagraphProperties properties = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines = new SpacingBetweenLines() 
            { 
                After = spacing.ToString(),
                Line = "276", // 1.15 line spacing
                LineRule = LineSpacingRuleValues.Auto
            };
            properties.AppendChild(spacingBetweenLines);
            para.AppendChild(properties);
            para.AppendChild(new Run());
        }
        
        private void AddSignatureLine(Body body, string name)
        {
            Table table = body.AppendChild(new Table());
            
            // Set table properties
            TableProperties tableProperties = new TableProperties();
            TableWidth tableWidth = new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct };
            
            // No borders for the table
            TableBorders tableBorders = new TableBorders();
            tableBorders.AppendChild(new TopBorder() { Val = BorderValues.None });
            tableBorders.AppendChild(new LeftBorder() { Val = BorderValues.None });
            tableBorders.AppendChild(new BottomBorder() { Val = BorderValues.None });
            tableBorders.AppendChild(new RightBorder() { Val = BorderValues.None });
            tableBorders.AppendChild(new InsideHorizontalBorder() { Val = BorderValues.None });
            tableBorders.AppendChild(new InsideVerticalBorder() { Val = BorderValues.None });
            
            tableProperties.AppendChild(tableWidth);
            tableProperties.AppendChild(tableBorders);
            table.AppendChild(tableProperties);
            
            // Create a row for the name, signature line, and date
            TableRow row = new TableRow();
            
            // First cell for name
            TableCell nameCell = new TableCell();
            Paragraph namePara = new Paragraph();
            Run nameRun = new Run();
            RunProperties nameRunProps = new RunProperties();
            nameRunProps.AppendChild(new Bold());
            nameRunProps.AppendChild(new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri" });
            nameRunProps.AppendChild(new FontSize() { Val = "22" });
            nameRun.AppendChild(nameRunProps);
            nameRun.AppendChild(new Text(name));
            namePara.AppendChild(nameRun);
            nameCell.AppendChild(namePara);
            
            // Second cell for signature line
            TableCell signatureCell = new TableCell();
            Paragraph signaturePara = new Paragraph();
            Run signatureRun = new Run();
            RunProperties signatureRunProps = new RunProperties();
            signatureRunProps.AppendChild(new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri" });
            signatureRunProps.AppendChild(new FontSize() { Val = "22" });
            signatureRun.AppendChild(signatureRunProps);
            signatureRun.AppendChild(new Text("_____________________"));
            signaturePara.AppendChild(signatureRun);
            signatureCell.AppendChild(signaturePara);
            
            // Third cell for date
            TableCell dateCell = new TableCell();
            Paragraph datePara = new Paragraph();
            ParagraphProperties dateParaProps = new ParagraphProperties();
            dateParaProps.AppendChild(new Justification() { Val = JustificationValues.Right });
            datePara.AppendChild(dateParaProps);
            
            Run dateRun = new Run();
            RunProperties dateRunProps = new RunProperties();
            dateRunProps.AppendChild(new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri" });
            dateRunProps.AppendChild(new FontSize() { Val = "22" });
            dateRun.AppendChild(dateRunProps);
            dateRun.AppendChild(new Text("Date: _________________"));
            datePara.AppendChild(dateRun);
            dateCell.AppendChild(datePara);
            
            // Add cells to row
            row.AppendChild(nameCell);
            row.AppendChild(signatureCell);
            row.AppendChild(dateCell);
            
            // Add row to table
            table.AppendChild(row);
        }
        
        private void AddFooter(Body body, string text)
        {
            Paragraph para = body.AppendChild(new Paragraph());
            ParagraphProperties properties = new ParagraphProperties();
            Justification justification = new Justification() { Val = JustificationValues.Center };
            properties.AppendChild(justification);
            para.AppendChild(properties);
            
            Run run = para.AppendChild(new Run());
            RunProperties runProperties = run.AppendChild(new RunProperties());
            runProperties.AppendChild(new Color() { Val = "888888" });
            runProperties.AppendChild(new FontSize() { Val = "22" });
            runProperties.AppendChild(new Italic());
            
            // Set Calibri font
            runProperties.AppendChild(new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri" });
            
            run.AppendChild(new Text(text));
        }

        private string FormatClauseText(string text, PartyInfo partyInfo)
        {
            if (string.IsNullOrEmpty(text))
                return text;
            
            string party1Name = $"{partyInfo.Party1FirstName} {partyInfo.Party1MiddleName} {partyInfo.Party1LastName}".Trim();
            string party2Name = $"{partyInfo.Party2FirstName} {partyInfo.Party2MiddleName} {partyInfo.Party2LastName}".Trim();
            
            string result = text;
            result = result.Replace("*PARTY 1*", party1Name);
            result = result.Replace("*PARTY 2*", party2Name);
            result = result.Replace("*PARTY ONE FULL NAME*", party1Name);
            result = result.Replace("*PARTY TWO FULL NAME*", party2Name);
            result = result.Replace("*PARTY 1 FULL NAME*", party1Name);
            result = result.Replace("*PARTY 2 FULL NAME*", party2Name);
            
            result = result.Replace("*LAWYER NAME FOR PARTY 1*", "");
            result = result.Replace("*LAWYER NAME FOR PARTY 2*", "");
            
            return result;
        }

        private void AddHeading(Body body, string text, int level, bool centered = false, bool allCaps = false)
        {
            Paragraph para = body.AppendChild(new Paragraph());
            ParagraphProperties properties = new ParagraphProperties();
            
            // Add spacing after paragraph with 1.15 line spacing
            SpacingBetweenLines spacing = new SpacingBetweenLines() 
            { 
                After = "200", 
                Line = "276", // 1.15 line spacing
                LineRule = LineSpacingRuleValues.Auto 
            };
            properties.AppendChild(spacing);
            
            // Center text if requested
            if (centered)
            {
                Justification justification = new Justification() { Val = JustificationValues.Center };
                properties.AppendChild(justification);
            }
            
            para.AppendChild(properties);
            
            Run run = para.AppendChild(new Run());
            RunProperties runProperties = run.AppendChild(new RunProperties());
            
            runProperties.AppendChild(new Bold());
            
            runProperties.AppendChild(new FontSize() { Val = "22" });
            
            string colorVal = MainHeadingColor;
            
            switch (level)
            {
                case 1:
                    colorVal = MainHeadingColor;
                    break;
                case 2:
                    colorVal = SubHeadingColor;
                    break;
                case 3:
                    colorVal = AccentColor;
                    break;
                default:
                    colorVal = "000000";
                    break;
            }
            
            runProperties.AppendChild(new Color() { Val = colorVal });
            
            // Apply Calibri font
            runProperties.AppendChild(new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri" });
            
            // Convert to uppercase if requested
            string displayText = allCaps ? text.ToUpper() : text;
            run.AppendChild(new Text(displayText));
        }

        private void AddParagraphWithSpacing(Body body, string text, bool justify = false, bool withBorder = false, JustificationValues? justificationValue = null)
        {
            // Check for page break markers
            if (text.Contains("===PAGE BREAK==="))
            {
                string[] segments = text.Split(new[] { "===PAGE BREAK===" }, StringSplitOptions.None);
                
                for (int i = 0; i < segments.Length; i++)
                {
                    // Process each segment before the page break
                    if (!string.IsNullOrWhiteSpace(segments[i]))
                    {
                        // Handle line breaks within each segment
                        if (segments[i].Contains("\n"))
                        {
                            string[] lines = segments[i].Split(new[] { "\n" }, StringSplitOptions.None);
                            foreach (var line in lines)
                            {
                                AddSingleParagraphWithSpacing(body, line, justify, withBorder, justificationValue);
                            }
                        }
                        else
                        {
                            AddSingleParagraphWithSpacing(body, segments[i], justify, withBorder, justificationValue);
                        }
                    }
                    
                    // Add page break after each segment except the last one
                    if (i < segments.Length - 1)
                    {
                        AddPageBreak(body);
                    }
                }
                
                return;
            }
            
            if (text.Contains("\n"))
            {
                string[] lines = text.Split(new[] { "\n" }, StringSplitOptions.None);
                
                foreach (var line in lines)
                {
                    AddSingleParagraphWithSpacing(body, line, justify, withBorder, justificationValue);
                }
                
                return;
            }
            
            AddSingleParagraphWithSpacing(body, text, justify, withBorder, justificationValue);
        }
        
        private void AddPageBreak(Body body)
        {
            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Break() { Type = BreakValues.Page });
        }
        
        private void AddSingleParagraphWithSpacing(Body body, string text, bool justify = false, bool withBorder = false, JustificationValues? justificationValue = null)
        {
            Paragraph para = body.AppendChild(new Paragraph());
            ParagraphProperties properties = new ParagraphProperties();
            
            // Add spacing with 1.15 line spacing
            SpacingBetweenLines spacing = new SpacingBetweenLines() 
            { 
                After = "200", 
                Line = "276", // 1.15 line spacing (240 * 1.15)
                LineRule = LineSpacingRuleValues.Auto 
            };
            properties.AppendChild(spacing);
            
            // Add justification based on the parameter or use justify parameter as before
            if (justificationValue.HasValue)
            {
                Justification justification = new Justification() { Val = justificationValue.Value };
                properties.AppendChild(justification);
            }
            else if (justify)
            {
                Justification justification = new Justification() { Val = JustificationValues.Both };
                properties.AppendChild(justification);
            }
            
            // Handle indentation for tab characters
            if (text.StartsWith("\t"))
            {
                // Count leading tabs
                int tabCount = 0;
                while (tabCount < text.Length && text[tabCount] == '\t')
                {
                    tabCount++;
                }
                
                // Remove the tab characters from the text
                text = text.TrimStart('\t');
                
                // Add indentation (720 twips per tab = 0.5 inch)
                Indentation indentation = new Indentation() { Left = (720 * tabCount).ToString() };
                properties.AppendChild(indentation);
            }
            
            para.AppendChild(properties);
            Run run = para.AppendChild(new Run());
            RunProperties runProperties = run.AppendChild(new RunProperties());
            
            // Set font size to 11 (22 in half-points)
            runProperties.AppendChild(new FontSize() { Val = "22" });
            
            // Apply Calibri font
            runProperties.AppendChild(new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri" });
            
            run.AppendChild(new Text(text));
        }

        private void AddParagraphWithAllCaps(Body body, string text, bool justify = false, bool withBorder = false, JustificationValues? justificationValue = null)
        {
            Paragraph para = body.AppendChild(new Paragraph());
            ParagraphProperties properties = new ParagraphProperties();
            
            SpacingBetweenLines spacing = new SpacingBetweenLines() 
            { 
                After = "200", 
                Line = "276", // 1.15 line spacing (240 * 1.15)
                LineRule = LineSpacingRuleValues.Auto 
            };
            properties.AppendChild(spacing);
            
            if (justificationValue.HasValue)
            {
                Justification justification = new Justification() { Val = justificationValue.Value };
                properties.AppendChild(justification);
            }
            else if (justify)
            {
                Justification justification = new Justification() { Val = JustificationValues.Both };
                properties.AppendChild(justification);
            }
            
            para.AppendChild(properties);
            Run run = para.AppendChild(new Run());
            RunProperties runProperties = run.AppendChild(new RunProperties());
            
            runProperties.AppendChild(new FontSize() { Val = "22" });
            
            runProperties.AppendChild(new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri" });
            
            runProperties.AppendChild(new Bold());
            
            run.AppendChild(new Text(text.ToUpper()));
        }

        private string FormatClauseId(int id)
        {
            if (id <= 0)
                return string.Empty;
                
            string idStr = id.ToString();
            
            if (idStr.Length == 3)
            {
                return idStr.Substring(0, 1);
            }
            else if (idStr.Length == 4)
            {
                string mainPart = idStr.Substring(0, 1);
                string subPart = idStr.Substring(1);
                
                if (subPart == "000")
                {
                    return mainPart;
                }
                
                int subPartInt = int.Parse(subPart);
                
                return $"{mainPart}.{subPartInt}";
            }
            else if (idStr.Length == 5)
            {
                string mainPart = idStr.Substring(0, 2);
                string subPart = idStr.Substring(2);
                
                int subPartInt = int.Parse(subPart);
                
                return $"{mainPart}.{subPartInt}";
            }
            else
            {
                return idStr;
            }
        }

        private void HandleTextWithPageBreaks(Body body, string text, bool justify = false, bool addPageBreaks = true)
        {
            string[] lines = text.Split(new[] { "\n" }, StringSplitOptions.None);
            
            int i = 0;
            while (i < lines.Length)
            {
                string line = lines[i];
                
                if (line.Trim().StartsWith("SCHEDULE") && i > 0 && addPageBreaks)
                {
                    AddPageBreak(body);
                    
                    StringBuilder scheduleContent = new StringBuilder();
                    scheduleContent.AppendLine(line);
                    
                    i++;
                    while (i < lines.Length && !lines[i].Trim().StartsWith("SCHEDULE"))
                    {
                        scheduleContent.AppendLine(lines[i]);
                        i++;
                    }
                    
                    string scheduleText = scheduleContent.ToString().Trim();
                    if (!string.IsNullOrWhiteSpace(scheduleText))
                    {
                        string[] scheduleLines = scheduleText.Split(new[] { "\n" }, StringSplitOptions.None);
                        foreach (var scheduleLine in scheduleLines)
                        {
                            if (!string.IsNullOrWhiteSpace(scheduleLine))
                            {
                                AddSingleParagraphWithSpacing(body, scheduleLine, justify);
                            }
                        }
                    }
                    
                }
                else
                {
                    if (!string.IsNullOrWhiteSpace(line))
                    {
                        AddSingleParagraphWithSpacing(body, line, justify);
                    }
                    i++;
                }
            }
        }

        private void AddScheduleTitle(Body body, string text, bool justify = false, bool center = false)
        {
            Paragraph para = body.AppendChild(new Paragraph());
            ParagraphProperties properties = new ParagraphProperties();
            
            SpacingBetweenLines spacing = new SpacingBetweenLines() 
            { 
                After = "200", 
                Line = "276", // 1.15 line spacing
                LineRule = LineSpacingRuleValues.Auto 
            };
            properties.AppendChild(spacing);
            
            if (center)
            {
                Justification justification = new Justification() { Val = JustificationValues.Center };
                properties.AppendChild(justification);
            }
            else if (justify)
            {
                Justification justification = new Justification() { Val = JustificationValues.Both };
                properties.AppendChild(justification);
            }
            
            para.AppendChild(properties);
            
            Run run = para.AppendChild(new Run());
            RunProperties runProperties = run.AppendChild(new RunProperties());
            
            runProperties.AppendChild(new FontSize() { Val = "22" });
            
            runProperties.AppendChild(new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri" });
            
            run.AppendChild(new Text(text));
        }
    }

    public static class Extensions
    {
        public static string ToOrdinal(this int num)
        {
            if (num <= 0) return num.ToString();

            switch (num % 100)
            {
                case 11:
                case 12:
                case 13:
                    return num + "th";
            }

            switch (num % 10)
            {
                case 1:
                    return num + "st";
                case 2:
                    return num + "nd";
                case 3:
                    return num + "rd";
                default:
                    return num + "th";
            }
        }
    }
} 