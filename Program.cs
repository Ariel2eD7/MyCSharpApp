
using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using Wp = DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System.Text.RegularExpressions;

namespace WordProcessorConsole
{
    class Program
    {




static void Main(string[] args)
{
    Console.WriteLine("Enter path to folder containing Word files:");
    string inputFolder = @"C:\Users\Developer1\index1\WordProcessorConsole\a"; // Or use Console.ReadLine();
    if (!Directory.Exists(inputFolder))
    {
        Console.WriteLine("Folder does not exist!");
        return;
    }

    // Create the output folder inside the input folder
    string outputFolder = Path.Combine(inputFolder, "Processed");
    Directory.CreateDirectory(outputFolder);

    // Get all .docx files in the folder
    var files = Directory.GetFiles(inputFolder, "*.docx");

    foreach (var file in files)
    {
        try
        {
            Console.WriteLine("Processing: " + Path.GetFileName(file));

            // Generate a new path for the processed file
            string outputFile = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(file) + "_modified.docx");

            // Process and save to new path
            ProcessWordDocument(file, outputFile);

            Console.WriteLine("Saved processed file: " + outputFile);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing {Path.GetFileName(file)}: {ex.Message}");
        }
    }

    Console.WriteLine("All documents processed. Press any key to exit.");
    Console.ReadKey();
}


private static void ProcessWordDocument(string originalPath, string newPath)
{
    // Copy original to new location first
    File.Copy(originalPath, newPath, true);

    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(newPath, true))
    {
        var mainPart = wordDoc.MainDocumentPart;
        if (mainPart?.Document?.Body == null)
        {
            Console.WriteLine("Document body not found!");
            return;
        }

        var body = mainPart.Document.Body;

        // ===== 1️⃣ Do all removals and modifications =====
        MergeTwoPagesWithHeadline(body, "השוואה חודשית");

        // Remove collapsible section (supports two optional headlines)
        RemoveCollapsibleSection(
            body,
            "ביטויים המובילים לאתר מעמוד ראשון בגוגל",
            "ביטויים המובילים לאתר מעמוד ראשון ושני בגוגל"
        );

        // Remove date ranges
        RemoveDateRanges(body);

        // Move first image to top to prevent accidental removal
        MoveFirstImageToTop(body);

        // Insert header from text start
        InsertHeaderFromTextStart(mainPart, body);

        // Other modifications
        RemoveMonthlyReportParagraphs(body);
        CopyHeadlineUnderLogo(body, "תנועה כוללת");
        PopulatePerformanceTables(body);
        MoveDetailsTableToTop(body);

        mainPart.Document.Save();
    }

    Console.WriteLine($"New file created: {newPath}");
}


        private static void MoveDetailsTableToTop(Wp.Body body)
        {
            // ===== 1️⃣ FIND THE "פירוט" TABLE =====
            var targetTable = body.Elements<Wp.Table>()
                .FirstOrDefault(tbl =>
                {
                    var firstRow = tbl.Elements<Wp.TableRow>().FirstOrDefault();
                    return firstRow != null && firstRow.InnerText.Trim().Contains("פירוט");
                });

            Wp.Table? cutDetailsTable = null;

            if (targetTable != null)
            {
                var allElements = body.Elements<OpenXmlElement>().ToList();
                int tblIndex = allElements.IndexOf(targetTable);

                // Remove the paragraph above if it contains the same headline text
                if (tblIndex > 0)
                {
                    if (allElements[tblIndex - 1] is Wp.Paragraph pAbove &&
                        pAbove.InnerText.Trim() == "בחודש האחרון בוצעו הפעולות הבאות:")
                    {
                        pAbove.Remove();
                        Console.WriteLine("Removed old headline above 'פירוט' table.");
                    }
                }

                // Cut (clone + remove) the "פירוט" table
                cutDetailsTable = (Wp.Table)targetTable.CloneNode(true);
                targetTable.Remove();
                Console.WriteLine("Cut 'פירוט' table successfully.");
            }
            else
            {
                Console.WriteLine("No table with first row containing 'פירוט' found.");
            }

            // ===== 2️⃣ FIND "תנועה כוללת" HEADLINE =====
            Wp.Paragraph? clonedHeadline = null;
            if (cutDetailsTable != null) // ✅ Only if פירוט exists
            {
                var headlinePara = body.Elements<Wp.Paragraph>()
                    .FirstOrDefault(p => p.InnerText.Trim() == "תנועה כוללת");

                if (headlinePara != null)
                {
                    clonedHeadline = (Wp.Paragraph)headlinePara.CloneNode(true);
                    UpdateParagraphText(clonedHeadline, "בחודש האחרון בוצעו הפעולות הבאות:");
                    Console.WriteLine("Cloned 'תנועה כוללת' headline.");
                }
                else
                {
                    Console.WriteLine("Headline 'תנועה כוללת' not found.");
                }
            }

            // ===== 3️⃣ FIND AND CUT "צפייה בקישורים" + ITS TABLE =====
            var linkHeadline = body.Elements<Wp.Paragraph>()
                .FirstOrDefault(p => p.InnerText.Trim() == "צפייה בקישורים");

            Wp.Paragraph? cutLinkHeadline = null;
            Wp.Table? cutLinkTable = null;

            if (linkHeadline != null)
            {
                var nextTable = linkHeadline.ElementsAfter().OfType<Wp.Table>().FirstOrDefault();
                if (nextTable != null)
                {
                    cutLinkHeadline = (Wp.Paragraph)linkHeadline.CloneNode(true);
                    cutLinkTable = (Wp.Table)nextTable.CloneNode(true);

                    // Remove originals
                    nextTable.Remove();
                    linkHeadline.Remove();

                    Console.WriteLine("Cut 'צפייה בקישורים' headline and its table.");
                }
                else
                {
                    Console.WriteLine("Headline 'צפייה בקישורים' found, but no table after it.");
                }
            }
            else
            {
                Console.WriteLine("Headline 'צפייה בקישורים' not found.");
            }

            // ===== 4️⃣ DETERMINE WHERE TO INSERT =====
            var firstPara = body.Elements<Wp.Paragraph>().FirstOrDefault();
            var firstImagePara = body.Elements<Wp.Paragraph>().FirstOrDefault(p => p.Descendants<Wp.Drawing>().Any());
            OpenXmlElement insertAfter = firstImagePara ?? (OpenXmlElement)firstPara;

            // ===== 5️⃣ INSERT AT TOP =====
            if (insertAfter != null)
            {
                if (cutDetailsTable != null && clonedHeadline != null)
                {
                    // Case A: We have a "פירוט" section
                    body.InsertAfter(clonedHeadline, insertAfter);
                    body.InsertAfter(cutDetailsTable, clonedHeadline);

                    if (cutLinkHeadline != null && cutLinkTable != null)
                    {
                        body.InsertAfter(cutLinkHeadline, cutDetailsTable);
                        body.InsertAfter(cutLinkTable, cutLinkHeadline);
                    }
                }
                else if (cutLinkHeadline != null && cutLinkTable != null)
                {
                    // Case B: No "פירוט" → only insert "צפייה בקישורים" section
                    body.InsertAfter(cutLinkHeadline, insertAfter);
                    body.InsertAfter(cutLinkTable, cutLinkHeadline);
                }
            }
            else
            {
                Console.WriteLine("No valid insertion point found (empty document?).");
            }

            Console.WriteLine("Moved 'פירוט' and/or 'צפייה בקישורים' sections to top of document successfully.");
        }

        
        private static void RemoveDateRanges(Wp.Body body)
        {
            // Pattern: <number> <word> <number> עד <number> <word> <number>
            // Example: 01 יולי 2025 עד 31 יולי 2025
            var dateRangePattern = new Regex(@"\b\d{1,2}\s+\S+\s+\d{4}\s+עד\s+\d{1,2}\s+\S+\s+\d{4}\b");

            var paragraphs = body.Elements<Wp.Paragraph>().ToList();

            foreach (var p in paragraphs)
            {
                if (dateRangePattern.IsMatch(p.InnerText))
                {
                    p.Remove();
                }
            }

            // ✅ Remove first appearance of the headline "צפייה בקישורים"
            var firstLinkHeadline = body.Elements<Wp.Paragraph>()
                .FirstOrDefault(p => p.InnerText.Trim() == "צפייה בקישורים");

            if (firstLinkHeadline != null)
            {
                firstLinkHeadline.Remove();
                Console.WriteLine("Removed first 'צפייה בקישורים' headline.");
            }

            // Optional: clean leftover empty paragraphs
            CleanupEmptyParagraphs(body);

            Console.WriteLine("Date ranges and first 'צפייה בקישורים' headline removed successfully.");
        }
    
     
        private static void MoveFirstImageToTop(Wp.Body body)
        {
            var firstImage = body.Descendants<Wp.Drawing>().FirstOrDefault();
            if (firstImage != null)
            {
                var clonedImage = (Wp.Drawing)firstImage.CloneNode(true);

                // Max width in pixels
                const int maxWidthPx = 400;
                const long emuPerPixel = 9525;
                long maxWidthEmu = maxWidthPx * emuPerPixel;

                // Adjust width and preserve aspect ratio
                var inline = clonedImage.Inline;
                if (inline != null)
                {
                    long origCx = inline.Extent?.Cx ?? maxWidthEmu;
                    long origCy = inline.Extent?.Cy ?? maxWidthEmu;

                    if (origCx > maxWidthEmu)
                    {
                        double ratio = (double)maxWidthEmu / origCx;
                        inline.Extent.Cx = maxWidthEmu;
                        inline.Extent.Cy = (long)(origCy * ratio);
                    }
                }

                var imagePara = new Wp.Paragraph(
                    new Wp.ParagraphProperties(
                        new Wp.Justification() { Val = Wp.JustificationValues.Center }
                    ),
                    new Wp.Run(clonedImage)
                );

                var firstPara = body.Elements<Wp.Paragraph>().FirstOrDefault();
                if (firstPara != null)
                    body.InsertBefore(imagePara, firstPara);
                else
                    body.PrependChild(imagePara);

                firstImage.Remove();
            }
        }

        private static void InsertHeaderFromTextStart(MainDocumentPart mainPart, Wp.Body body)
    {
        // Find the first paragraph that starts with "דו\"ח תקופתי"
        var targetPara = body.Elements<Wp.Paragraph>()
                            .FirstOrDefault(p => p.InnerText.TrimStart().StartsWith("דו\"ח תקופתי"));

        if (targetPara == null) return;

        // Clone all runs from that paragraph
        var clonedRuns = targetPara.Elements<Wp.Run>().Select(r => (Wp.Run)r.CloneNode(true)).ToList();

        // Remove the original paragraph from the body
        targetPara.Remove();

        // Get or create header
        HeaderPart headerPart = mainPart.HeaderParts.FirstOrDefault() ?? mainPart.AddNewPart<HeaderPart>();
        if (headerPart.Header == null) headerPart.Header = new Wp.Header();
        headerPart.Header.RemoveAllChildren();

        // Create new header paragraph
        var para = new Wp.Paragraph(
            new Wp.ParagraphProperties(
                new Wp.SpacingBetweenLines() { After = "200" },
                new Wp.ParagraphBorders(
                    new Wp.BottomBorder() { Val = Wp.BorderValues.Single, Color = "000000", Size = 4, Space = 1 }
                ),
                new Wp.BiDi(),
                new Wp.Justification() { Val = Wp.JustificationValues.Right }
            )
        );

        foreach (var run in clonedRuns)
        {
            var rPr = run.RunProperties ?? (run.RunProperties = new Wp.RunProperties());
            rPr.FontSize = new Wp.FontSize() { Val = "52" };
            rPr.FontSizeComplexScript = new Wp.FontSizeComplexScript() { Val = "52" };
            rPr.AppendChild(new Wp.RightToLeftText());
            rPr.Color = new Wp.Color() { Val = "000000" };
            para.Append(run);
        }

        headerPart.Header.AppendChild(para);
        headerPart.Header.Save();

        // Link the header to the section
        var sectProps = body.Elements<Wp.SectionProperties>().LastOrDefault() ?? new Wp.SectionProperties();
        if (!body.Elements<Wp.SectionProperties>().Any()) body.Append(sectProps);
        sectProps.RemoveAllChildren<Wp.HeaderReference>();

        string rId = mainPart.GetIdOfPart(headerPart);
        var headerReference = new Wp.HeaderReference() { Type = Wp.HeaderFooterValues.Default, Id = rId };
        sectProps.Append(headerReference);
    }

        private static void MergeTwoPagesWithHeadline(Wp.Body body, string headline)
        {
            var paragraphs = body.Elements<Wp.Paragraph>().ToList();

            var headlineIndexes = paragraphs
                .Select((p, idx) => new { Paragraph = p, Index = idx })
                .Where(x => x.Paragraph.InnerText.Trim() == headline)
                .Select(x => x.Index)
                .ToList();

            if (headlineIndexes.Count < 2) return;

            int firstIndex = headlineIndexes[0];
            int secondIndex = headlineIndexes[1];

            var contentToMove = paragraphs
                .Skip(firstIndex + 1)
                .Take(secondIndex - firstIndex - 1)
                .Select(p => (Wp.Paragraph)p.CloneNode(true))
                .ToList();

            for (int i = firstIndex + 1; i < secondIndex; i++)
                paragraphs[i].Remove();

            paragraphs[secondIndex].Remove();

            Wp.Paragraph insertAfterPara = paragraphs[firstIndex];
            foreach (var p in contentToMove)
            {
                if (!string.IsNullOrWhiteSpace(p.InnerText))
                {
                    body.InsertAfter(p, insertAfterPara);
                    insertAfterPara = p;
                }
            }
        }


        private static void RemoveCollapsibleSection(Wp.Body body, params string[] possibleHeadlines)
        {
            if (possibleHeadlines == null || possibleHeadlines.Length == 0)
                return;

            var elements = body.Elements<OpenXmlElement>().ToList();
            bool inTargetSection = false;
            int i = 0;

            while (i < elements.Count)
            {
                var elem = elements[i];

                // Check if we reached one of the possible headlines
                if (!inTargetSection && elem is Wp.Paragraph para)
                {
                    string text = para.InnerText.Trim();
                    if (possibleHeadlines.Contains(text))
                    {
                        inTargetSection = true;
                        elem.Remove();
                        elements.RemoveAt(i);
                        Console.WriteLine($"Found and started removing section: {text}");
                        continue;
                    }
                }

                if (inTargetSection)
                {
                    if (elem is Wp.Paragraph para2 && para2.InnerText.Trim() == "מילים מובילות")
                    {
                        para2.Remove();
                        elements.RemoveAt(i);

                        // Remove following tables
                        while (i < elements.Count && elements[i] is Wp.Table table)
                        {
                            table.Remove();
                            elements.RemoveAt(i);
                        }

                        CleanupEmptyParagraphs(body);
                        Console.WriteLine("Finished removing collapsible section.");
                        break;
                    }
                    else
                    {
                        elem.Remove();
                        elements.RemoveAt(i);
                    }
                }
                else
                {
                    i++;
                }
            }
        }


        private static void CleanupEmptyParagraphs(Wp.Body body)
        {
            var paragraphs = body.Elements<Wp.Paragraph>().ToList();
            foreach (var p in paragraphs)
            {
                if (!p.Descendants<Wp.Drawing>().Any() && string.IsNullOrWhiteSpace(p.InnerText))
                {
                    p.Remove();
                }
            }

            // Remove empty section properties at the end of the body
            var sectProps = body.Elements<Wp.SectionProperties>().ToList();
            foreach (var sp in sectProps)
            {
                if (!sp.HasChildren)
                    sp.Remove();
            }
        }

        private static void RemoveMonthlyReportParagraphs(Wp.Body body)
        {
            string startText = "פעולות שוטפות";
            string endText = "בעקבות פעולות הקידום:";

            var paragraphs = body.Elements<Wp.Paragraph>().ToList();

            foreach (var p in paragraphs)
            {
                string text = p.InnerText.Trim();
                if (text.StartsWith(startText) || text.StartsWith(endText))
                {
                    p.Remove();
                }
            }

            // Optional: clean leftover empty paragraphs
            CleanupEmptyParagraphs(body);
        }

        private static void CopyHeadlineUnderLogo(Wp.Body body, string headline)
        {
            var sourcePara = body.Elements<Wp.Paragraph>()
                                 .FirstOrDefault(p => p.InnerText.Trim() == headline);

            if (sourcePara == null)
            {
                Console.WriteLine($"Headline '{headline}' not found.");
                return;
            }

            var clonedPara = (Wp.Paragraph)sourcePara.CloneNode(true);

            // === Change headline text ===
            UpdateParagraphText(clonedPara, "בעקבות פעולות הקידום:");

            // === Add another line in same style ===
            var firstRun = clonedPara.GetFirstChild<Wp.Run>();
            if (firstRun != null)
            {
                var newRun = (Wp.Run)firstRun.CloneNode(true);
                newRun.RemoveAllChildren<Wp.Break>();
                newRun.PrependChild(new Wp.Break());

                var textElement = newRun.GetFirstChild<Wp.Text>();
                if (textElement != null)
                    textElement.Text = "הגענו למקום מעולה (עמוד 1 שורה 1)";
                else
                    newRun.AppendChild(new Wp.Text("הגענו למקום מעולה (עמוד 1 שורה 1)"));

                clonedPara.AppendChild(newRun);
            }

            var imagePara = body.Elements<Wp.Paragraph>()
                                .FirstOrDefault(p => p.Descendants<Wp.Drawing>().Any());

            if (imagePara == null)
            {
                Console.WriteLine("Logo image not found — cannot insert headline below it.");
                return;
            }

            // Insert headline under logo
            body.InsertAfter(clonedPara, imagePara);

            // === Add invisible full-width table ===
            var table1 = CreateInvisibleFullWidthTable();
            body.InsertAfter(table1, clonedPara);

            // Second headline
            var para2 = (Wp.Paragraph)sourcePara.CloneNode(true);
            UpdateParagraphText(para2, "שמרנו על מקום מעולה (עמוד 1 שורה 1)");
            body.InsertAfter(para2, table1);

            // Second invisible table
            var table2 = CreateInvisibleFullWidthTable();
            body.InsertAfter(table2, para2);

            // Third headline
            var para3 = (Wp.Paragraph)sourcePara.CloneNode(true);
            UpdateParagraphText(para3, "התקדמנו במיקומי מילות המפתח הבאות:");
            body.InsertAfter(para3, table2);

            // Third invisible table
            var table3 = CreateInvisibleFullWidthTable();
            body.InsertAfter(table3, para3);

            Console.WriteLine("Inserted headlines and invisible full-width tables successfully under logo.");
        }


        private static Wp.Table CreateInvisibleFullWidthTable()
        {
            var table = new Wp.Table();

            // Set full width and visible borders
            var tblProps = new Wp.TableProperties(
                new Wp.TableWidth { Type = Wp.TableWidthUnitValues.Pct, Width = "100%" }, // 100%
                new Wp.TableLayout { Type = Wp.TableLayoutValues.Fixed }, // FIXED LAYOUT!
                new Wp.TableBorders(
                    new Wp.TopBorder { Val = Wp.BorderValues.Single, Color = "000000", Size = 4 },
                    new Wp.BottomBorder { Val = Wp.BorderValues.Single, Color = "000000", Size = 4 },
                    new Wp.LeftBorder { Val = Wp.BorderValues.Single, Color = "000000", Size = 4 },
                    new Wp.RightBorder { Val = Wp.BorderValues.Single, Color = "000000", Size = 4 },
                    new Wp.InsideHorizontalBorder { Val = Wp.BorderValues.Single, Color = "000000", Size = 4 },
                    new Wp.InsideVerticalBorder { Val = Wp.BorderValues.Single, Color = "000000", Size = 4 }
                )
            );

            table.AppendChild(tblProps);

            // --- Do NOT create first row with placeholder text ---
            // We'll add rows dynamically when inserting actual content

            return table;
        }

        private static void UpdateParagraphText(Wp.Paragraph paragraph, string newText)
        {
            foreach (var run in paragraph.Descendants<Wp.Run>())
            {
                var textElement = run.GetFirstChild<Wp.Text>();
                if (textElement != null)
                    textElement.Text = newText;
            }
        }
     
     private static void PopulatePerformanceTables(Wp.Body body)
{
    var keywordHeadline = body.Elements<Wp.Paragraph>()
        .FirstOrDefault(p => p.InnerText.Trim() == "ביטויים בקידום");

    if (keywordHeadline == null)
    {
        Console.WriteLine("Headline 'ביטויים בקידום' not found.");
        return;
    }

    var sourceTable = keywordHeadline.ElementsAfter().OfType<Wp.Table>().FirstOrDefault();
    if (sourceTable == null)
    {
        Console.WriteLine("Source table under 'ביטויים בקידום' not found.");
        return;
    }

    // find destination tables
    var reachedPara = body.Elements<Wp.Paragraph>().FirstOrDefault(p => p.InnerText.Contains("הגענו למקום מעולה"));
    var keptPara = body.Elements<Wp.Paragraph>().FirstOrDefault(p => p.InnerText.Contains("שמרנו על מקום מעולה"));
    var progressedPara = body.Elements<Wp.Paragraph>().FirstOrDefault(p => p.InnerText.Contains("התקדמנו במיקומי מילות"));

    if (reachedPara == null || keptPara == null || progressedPara == null)
    {
        Console.WriteLine("One or more destination headlines not found.");
        return;
    }

    var reachedTable = reachedPara.ElementsAfter().OfType<Wp.Table>().FirstOrDefault();
    var keptTable = keptPara.ElementsAfter().OfType<Wp.Table>().FirstOrDefault();
    var progressedTable = progressedPara.ElementsAfter().OfType<Wp.Table>().FirstOrDefault();

    if (reachedTable == null || keptTable == null || progressedTable == null)
    {
        Console.WriteLine("One or more destination tables not found.");
        return;
    }

    var rows = sourceTable.Elements<Wp.TableRow>().Skip(1); // skip header row

    foreach (var row in rows)
    {
        var cells = row.Elements<Wp.TableCell>().ToList();
        if (cells.Count < 8) continue;

        string keyword = cells[0].InnerText.Trim();
        if (string.IsNullOrWhiteSpace(keyword)) continue;

        // --- Parse numeric values safely ---
        var values = cells.Skip(2).Take(6)
            .Select(c =>
            {
                var text = c.InnerText.Trim();
                if (string.IsNullOrWhiteSpace(text))
                    return (double?)null;

                // Try parse using invariant culture to handle "." or ","
                if (double.TryParse(text, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double v))
                    return (double?)v;

                return null;
            })
            .Where(v => v.HasValue)
            .Select(v => v.Value)
            .ToList();

        // Skip rows with too few numeric values
        if (values.Count < 2)
            continue;

        double last = values.Last();
        double prev = values[values.Count - 2];
        var previous5 = values.Take(values.Count - 1).ToList();

        if (last == 1)
        {
            if (previous5.All(v => v > 1))
                InsertIntoExistingTableRTL(reachedTable, keyword, "\u200B");
            else if (previous5.Any(v => v == 1))
                InsertIntoExistingTableRTL(keptTable, keyword, "\u200B", alternateCells: true);
        }
        else if (previous5.All(z => z > 10) && (last <= 10))
        {
            InsertIntoExistingTableRTL(progressedTable, keyword, $"!ממקום {prev} למקום {last} וכניסה לעמוד הראשון");
        }
        else if (previous5.All(v => last < v))
        {
            InsertIntoExistingTableRTL(progressedTable, keyword, $"ממקום {prev} למקום {last}");
        }
    }

    Console.WriteLine("Performance tables populated (skipping invalid/empty cells).");
}

        private static void InsertIntoExistingTableRTL(Wp.Table table, string keyword, string? progress = null, bool alternateCells = false)
        {
            bool isProgressedTable = !string.IsNullOrEmpty(progress);

            Wp.Paragraph CreateParagraph(string text, bool bold = false)
            {
                var run = new Wp.Run();
                var rPr = new Wp.RunProperties
                {
                    RunFonts = new Wp.RunFonts { Ascii = "Arial", HighAnsi = "Arial" },
                    FontSize = new Wp.FontSize { Val = "24" } // 12pt
                };
                if (bold) rPr.Bold = new Wp.Bold();
                run.RunProperties = rPr;
                run.Append(new Wp.Text(text));

                return new Wp.Paragraph(
                    new Wp.ParagraphProperties(
                        new Wp.SpacingBetweenLines { Before = "0", After = "0" },
                        new Wp.Indentation { Left = "0", Right = "0" }
                    ),
                    run
                );
            }

            if (alternateCells)
            {
                // Fill last row if there is an empty cell
                Wp.TableRow? lastRow = table.Elements<Wp.TableRow>().LastOrDefault();
                if (lastRow != null)
                {
                    var cells = lastRow.Elements<Wp.TableCell>().ToList();
                    bool added = false;

                    if (cells.Count >= 2)
                    {
                        for (int c = 1; c >= 0; c--) // Fill RIGHT first, then LEFT
                        {
                            var cell = cells[c];
                            if (!cell.Elements<Wp.Paragraph>().Any() || cell.Elements<Wp.Paragraph>().All(p => string.IsNullOrWhiteSpace(p.InnerText)))
                            {
                                cell.RemoveAllChildren();
                                cell.Append(CreateParagraph(keyword));
                                added = true;
                                break;
                            }
                        }
                    }

                    if (added) return;
                }

                // If last row is full or doesn't exist, create a new row with RIGHT cell filled
                var newRow = new Wp.TableRow(
                    new Wp.TableCell(new Wp.Paragraph()),      // LEFT empty
                    new Wp.TableCell(CreateParagraph(keyword)) // RIGHT filled
                );
                table.Append(newRow);
            }
            else
            {
                Wp.TableRow newRow;
                if (isProgressedTable)
                {
                    newRow = new Wp.TableRow(
                        new Wp.TableCell(CreateParagraph(progress)), // LEFT
                        new Wp.TableCell(CreateParagraph(keyword))   // RIGHT
                    );
                }
                else
                {
                    newRow = new Wp.TableRow(
                        new Wp.TableCell(new Wp.Paragraph()),        // LEFT empty
                        new Wp.TableCell(CreateParagraph(keyword))   // RIGHT
                    );
                }
                table.Append(newRow);
            }
        }

    }  
}