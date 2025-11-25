
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

    // Ask the user for the text to insert into the 'פירוט' table
    Console.WriteLine("Enter the text to insert in the 'פירוט' table:");
    string detailsText = Console.ReadLine();

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

            string outputFile = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(file) + "_modified.docx");

            // Pass detailsText to the processing function
            ProcessWordDocument(file, outputFile, detailsText);

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

private static void ProcessWordDocument(string originalPath, string newPath, string detailsText)
{
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

        // Your full processing pipeline
        MergeTwoPagesWithHeadline(body, "השוואה חודשית");
        RemoveCollapsibleSection(
            body,
            "ביטויים המובילים לאתר מעמוד ראשון בגוגל",
            "ביטויים המובילים לאתר מעמוד ראשון ושני בגוגל"
        );
        RemoveDateRanges(body);
        MoveFirstImageToTop(body);
        InsertHeaderFromTextStart(mainPart, body);
        RemoveMonthlyReportParagraphs(body);
        CopyHeadlineUnderLogo(body, "תנועה כוללת");
        PopulatePerformanceTables(body);

        // ✅ Pass the user text to the פירוט table function
        MoveDetailsTableToTop(body, detailsText);

        ReplaceLowSearchVolumeWithNumber(body);

        mainPart.Document.Save();
    }

    Console.WriteLine($"New file created: {newPath}");
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
                new Wp.Justification() { Val = Wp.JustificationValues.Left }
            )
        );

        foreach (var run in clonedRuns)
        {
            var rPr = run.RunProperties ?? (run.RunProperties = new Wp.RunProperties());
            rPr.FontSize = new Wp.FontSize() { Val = "52" };
            rPr.FontSizeComplexScript = new Wp.FontSizeComplexScript() { Val = "52" };
            rPr.AppendChild(new Wp.RightToLeftText());
            rPr.Color = new Wp.Color() { Val = "#17365D" };






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

        // --- Clone the paragraph to preserve style ---
        var styleTemplate = (Wp.Paragraph)sourcePara.CloneNode(true);

        // --- Create the new intro headline ---
        // --- Create the new intro headline with a line break ---
        var introHeadline = (Wp.Paragraph)styleTemplate.CloneNode(true);
        introHeadline.RemoveAllChildren<Wp.Run>(); // clear old text runs

        var run1 = new Wp.Run(new Wp.Text("פעולות שוטפות"));
        var lineBreak = new Wp.Run(new Wp.Break());
        var run2 = new Wp.Run(new Wp.Text(":מידי חודש מבוצעות פעולות בדיקה ובקרה הכוללות"));

        introHeadline.Append(run1, lineBreak, run2);


        // --- Create the plain text bullet list ---
        string[] introBullets =
        {
    "בדיקות כפילות תוכן •",
    "בדיקות תקינות קוד •",
    "בדיקות מהירות וזמינות שרת •",
    "בדיקות תקינות אופטימיזציה כולל TITLE, H1, H2, META •",
    "בדיקות תקינות לינקים נכנסים •",
    "בדיקת מיקומי ביטויי המפתח •",
    "בדיקת התנהגות גולשים באתר הנייח ובמובייל •",
    ", A1, זפו, console,  לצורך בקרה על תקינות האתר והתאמת אופטימיזציה לגוגל, שימוש בכלי בקרה אנליטיקס, פרוג",
    ".וכלי עבודה נוספים ייעודיים SEOQUAKE , מג'סטיק, HOTJAR, "
};

        var plainPara = new Wp.Paragraph();

        // optional: make sure it's RTL for Hebrew text
        plainPara.ParagraphProperties = new Wp.ParagraphProperties(
            new Wp.BiDi()
        );

        foreach (var line in introBullets)
        {
            var run = new Wp.Run();

            // --- Apply font size (24pt = 48 half-points) ---
            var runProps = new Wp.RunProperties(
                new Wp.FontSize() { Val = "22" },
                new Wp.FontSizeComplexScript() { Val = "22" } // needed for Hebrew
            );

            run.Append(runProps);
            run.Append(new Wp.Text(line.Trim())
            {
                Space = SpaceProcessingModeValues.Preserve
            });

            plainPara.Append(run);
            plainPara.Append(new Wp.Run(new Wp.Break())); // new line
        }



        // --- Create your existing main headline section ---
        var mainHeadlinePara = (Wp.Paragraph)sourcePara.CloneNode(true);
        UpdateParagraphText(mainHeadlinePara, "בעקבות פעולות הקידום:");

        var reachedPara = (Wp.Paragraph)sourcePara.CloneNode(true);
        UpdateParagraphText(reachedPara, "הגענו למקום מעולה (עמוד 1 שורה 1)");

        // --- Find logo paragraph ---
        var imagePara = body.Elements<Wp.Paragraph>()
                            .FirstOrDefault(p => p.Descendants<Wp.Drawing>().Any());

        if (imagePara == null)
        {
            Console.WriteLine("Logo image not found — cannot insert headline below it.");
            return;
        }

        // --- Insert the new intro section above 'בעקבות פעולות הקידום:' ---
        body.InsertAfter(introHeadline, imagePara);
        body.InsertAfter(plainPara, introHeadline);

        // --- Insert the main headline below the intro ---
        body.InsertAfter(mainHeadlinePara, plainPara);

        // --- Then continue with your three subsections ---
        body.InsertAfter(reachedPara, mainHeadlinePara);

        var table1 = CreateInvisibleFullWidthTable();
        body.InsertAfter(table1, reachedPara);

        var keptPara = (Wp.Paragraph)sourcePara.CloneNode(true);
        UpdateParagraphText(keptPara, "שמרנו על מקום מעולה (עמוד 1 שורה 1)");
        body.InsertAfter(keptPara, table1);

        var table2 = CreateInvisibleFullWidthTable();
        body.InsertAfter(table2, keptPara);

        var progressedPara = (Wp.Paragraph)sourcePara.CloneNode(true);
        UpdateParagraphText(progressedPara, "התקדמנו במיקומי מילות המפתח הבאות:");
        body.InsertAfter(progressedPara, table2);

        var table3 = CreateInvisibleFullWidthTable();
        body.InsertAfter(table3, progressedPara);

        Console.WriteLine("Inserted intro section, headlines, and tables successfully under logo.");
    }


    private static Wp.Table CreateInvisibleFullWidthTable()
    {
        var table = new Wp.Table();

        // Set full width and visible borders
        var tblProps = new Wp.TableProperties(
            new Wp.TableWidth { Type = Wp.TableWidthUnitValues.Pct, Width = "100%" }, // 100%
            new Wp.TableLayout { Type = Wp.TableLayoutValues.Fixed }, // FIXED LAYOUT!
            new Wp.TableBorders(
                new Wp.TopBorder { Val = Wp.BorderValues.Single, Color = "ffffff", Size = 4 },
                new Wp.BottomBorder { Val = Wp.BorderValues.Single, Color = "ffffff", Size = 4 },
                new Wp.LeftBorder { Val = Wp.BorderValues.Single, Color = "ffffff", Size = 4 },
                new Wp.RightBorder { Val = Wp.BorderValues.Single, Color = "ffffff", Size = 4 },
                new Wp.InsideHorizontalBorder { Val = Wp.BorderValues.Single, Color = "ffffff", Size = 4 },
                new Wp.InsideVerticalBorder { Val = Wp.BorderValues.Single, Color = "ffffff", Size = 4 }
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
            if (cells.Count < 3) continue; // must have at least keyword + local searches + 1 month

            string keyword = cells[0].InnerText.Trim(); // מילת מפתח
            if (string.IsNullOrWhiteSpace(keyword)) continue;

            // The rest of the columns after column 2 are monthly positions
            var positionCells = cells.Skip(2).ToList();
            if (positionCells.Count == 0) continue;

            // Parse last column (leftmost position)
            string lastText = positionCells.Last().InnerText.Trim();

            // Skip if empty or dash
            if (string.IsNullOrWhiteSpace(lastText) || lastText == "-")
                continue;

            if (!double.TryParse(lastText, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out double lastValue))
                continue;

            // Parse previous cells
            var prevCells = positionCells.Take(positionCells.Count - 1).ToList();

            var previousValues = prevCells
                .Select(c =>
                {
                    var t = c.InnerText.Trim();
                    if (double.TryParse(t, System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out double v))
                        return (double?)v;
                    return null;
                })
                .Where(v => v.HasValue)
                .Select(v => v.Value)
                .ToList();

            // Track if all previous cells are "-" or empty
            bool allPrevEmpty = prevCells.All(c =>
            {
                var t = c.InnerText.Trim();
                return string.IsNullOrWhiteSpace(t) || t == "-";
            });

            // ---- CASE 1: last position = 1 ----
            if (lastValue == 1)
            {
                if (previousValues.Any(v => v == 1))
                {
                    // kept top position
                    InsertIntoExistingTableRTL(keptTable, keyword, "\u200B", alternateCells: true);
                }
                else
                {
                    // newly reached top position
                    InsertIntoExistingTableRTL(reachedTable, keyword, "\u200B");
                }
            }
            // ---- CASE 2: last position != 1 ----
            else
            {
                // improved if last is smaller than all previous numeric values
                if (previousValues.Count == 0 || previousValues.All(v => lastValue < v))
                {
                    var prevCellText = positionCells.Count >= 2
                        ? positionCells[positionCells.Count - 2].InnerText.Trim()
                        : string.Empty;

                    bool prevIsEmpty = string.IsNullOrWhiteSpace(prevCellText) || prevCellText == "-";

                    // Check if all previous numeric values are >10 or empty
                    bool allPrevAbove10OrEmpty = prevCells.All(c =>
                    {
                        var t = c.InnerText.Trim();
                        if (string.IsNullOrWhiteSpace(t) || t == "-") return true;
                        if (double.TryParse(t, System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture, out double v))
                            return v > 10;
                        return true;
                    }); 

                    // Decide text to insert
                    if (allPrevEmpty)
                    {
                        if (lastValue <= 10)
                            InsertIntoExistingTableRTL(progressedTable, keyword, $"!כניסה למקום {lastValue} וכניסה לעמוד הראשון");
                        else
                            InsertIntoExistingTableRTL(progressedTable, keyword, $"כניסה למקום {lastValue}");
                    }
                    else if (prevIsEmpty)
                    {
                        if (lastValue <= 10 && allPrevAbove10OrEmpty)
                            InsertIntoExistingTableRTL(progressedTable, keyword, $"!כניסה למקום {lastValue} וכניסה לעמוד הראשון");
                        else
                            InsertIntoExistingTableRTL(progressedTable, keyword, $"כניסה למקום {lastValue}");
                    }
                    else
                    {
                        double prevValue = previousValues.LastOrDefault();
                        if (lastValue <= 10 && allPrevAbove10OrEmpty)
                            InsertIntoExistingTableRTL(progressedTable, keyword, $"!ממקום {prevValue} למקום {lastValue} וכניסה לעמוד הראשון");
                        else
                            InsertIntoExistingTableRTL(progressedTable, keyword, $"ממקום {prevValue} למקום {lastValue}");
                    }
                }
            }
        }

        // --- Cleanup: remove empty tables and their headlines ---
        void RemoveHeadlineAndTableIfEmpty(Wp.Paragraph headline)
        {
            var nextTable = headline.ElementsAfter().OfType<Wp.Table>().FirstOrDefault();
            if (nextTable != null)
            {
                // Check if table has any non-empty cells
                var hasData = nextTable.Elements<Wp.TableRow>().Any(r =>
                    r.Elements<Wp.TableCell>().Any(c => !string.IsNullOrWhiteSpace(c.InnerText)));

                if (!hasData)
                {
                    Console.WriteLine($"Removing empty section: {headline.InnerText}");
                    nextTable.Remove();
                    headline.Remove();
                }
            }
        }

        // Remove empty subsections
        RemoveHeadlineAndTableIfEmpty(reachedPara);
        RemoveHeadlineAndTableIfEmpty(keptPara);
        RemoveHeadlineAndTableIfEmpty(progressedPara);

        // --- Remove main headline ONLY if all subsections are gone ---
        bool anySubsectionExists = body.Elements<Wp.Paragraph>().Any(p =>
            p.InnerText.Contains("הגענו למקום מעולה") ||
            p.InnerText.Contains("שמרנו על מקום מעולה") ||
            p.InnerText.Contains("התקדמנו במיקומי מילות"));

        if (!anySubsectionExists)
        {
            var mainHeadline = body.Elements<Wp.Paragraph>()
                .FirstOrDefault(p => p.InnerText.Contains("בעקבות פעולות הקידום"));

            if (mainHeadline != null)
            {
                mainHeadline.Remove();
                Console.WriteLine("Removed 'בעקבות פעולות הקידום' because no performance tables contained data.");
            }
        }

        Console.WriteLine("Performance tables populated and cleaned up (main headline removed only if all subsections are empty).");

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






    private static void MoveDetailsTableToTop(Wp.Body body, string? customText = null)
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

            // ✅ Add new row with plain customText or default message
            var messageText = customText ?? "אין נתונים זמינים";

            var messageRow = new Wp.TableRow(
                new Wp.TableCell(
                    new Wp.Paragraph(
                        new Wp.Run(
                            new Wp.Text(messageText)
                        )
                    )
                )
            );

            cutDetailsTable.AppendChild(messageRow);
            Console.WriteLine("Appended plain customText row to existing 'פירוט' table.");

        }
        else
        {
            Console.WriteLine("No table with first row containing 'פירוט' found — creating a new one.");

            // ===== Create "פירוט" table =====
            cutDetailsTable = new Wp.Table(
                new Wp.TableProperties(
                    new Wp.TableWidth { Type = Wp.TableWidthUnitValues.Pct, Width = "100%" },
                    new Wp.TableBorders(
                        new Wp.TopBorder() { Val = Wp.BorderValues.Single, Size = 4, Color = "D3D3D3" },
                        new Wp.BottomBorder() { Val = Wp.BorderValues.Single, Size = 4, Color = "D3D3D3" },
                        new Wp.LeftBorder() { Val = Wp.BorderValues.Single, Size = 4, Color = "D3D3D3" },
                        new Wp.RightBorder() { Val = Wp.BorderValues.Single, Size = 4, Color = "D3D3D3" },
                        new Wp.InsideHorizontalBorder() { Val = Wp.BorderValues.Single, Size = 4, Color = "D3D3D3" },
                        new Wp.InsideVerticalBorder() { Val = Wp.BorderValues.Single, Size = 4, Color = "D3D3D3" }
                    )
                ),

                // Header Row
                new Wp.TableRow(
                    new Wp.TableRowProperties(
                        new Wp.TableRowHeight { Val = 400, HeightType = Wp.HeightRuleValues.AtLeast }
                    ),
                    new Wp.TableCell(
                        new Wp.TableCellProperties(
                            new Wp.Shading() { Color = "auto", Fill = "EEECE1", Val = Wp.ShadingPatternValues.Clear },
                            new Wp.TableCellVerticalAlignment() { Val = Wp.TableVerticalAlignmentValues.Center }
                        ),
                        new Wp.Paragraph(
                            new Wp.ParagraphProperties(
                                new Wp.Justification() { Val = Wp.JustificationValues.Left },
                                new Wp.SpacingBetweenLines() { Before = "0", After = "0" }
                            ),
                            new Wp.Run(
                                new Wp.RunProperties(
                                    new Wp.Bold(),
                                    new Wp.FontSize() { Val = "24" },
                                    new Wp.RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }
                                ),
                                new Wp.Text("פירוט")
                            )
                        )
                    )
                ),

                // Data Row (with customText)
                new Wp.TableRow(
                    new Wp.TableRowProperties(
                        new Wp.TableRowHeight { Val = 400, HeightType = Wp.HeightRuleValues.AtLeast }
                    ),
                    new Wp.TableCell(
                        new Wp.TableCellProperties(
                            new Wp.TableCellVerticalAlignment() { Val = Wp.TableVerticalAlignmentValues.Center }
                        ),
                        new Wp.Paragraph(
                            new Wp.ParagraphProperties(
                                new Wp.Justification() { Val = Wp.JustificationValues.Left },
                                new Wp.SpacingBetweenLines() { Before = "0", After = "0" }
                            ),
                            new Wp.Run(
                                new Wp.RunProperties(
                                    new Wp.Bold(),
                                    new Wp.FontSize() { Val = "24" },
                                    new Wp.RunFonts() { Ascii = "Arial", HighAnsi = "Arial" }
                                ),
                                new Wp.Text(customText ?? "אין נתונים זמינים")
                            )
                        )
                    )
                )
            );
        }

        // ===== 2️⃣ FIND "תנועה כוללת" HEADLINE =====
        Wp.Paragraph? clonedHeadline = null;
        if (cutDetailsTable != null)
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
                Console.WriteLine("Headline 'תנועה כוללת' not found — creating new one.");
                clonedHeadline = new Wp.Paragraph(
                    new Wp.Run(new Wp.Text("בחודש האחרון בוצעו הפעולות הבאות:"))
                );
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

        Console.WriteLine("Moved or created 'פירוט' and/or 'צפייה בקישורים' sections at top successfully.");
    }




    private static void ReplaceLowSearchVolumeWithNumber(Wp.Body body)
    {
        // 1️⃣ Find the "ביטויים בקידום" headline paragraph
        var keywordHeadline = body.Elements<Wp.Paragraph>()
            .FirstOrDefault(p => p.InnerText.Trim() == "ביטויים בקידום");

        if (keywordHeadline == null)
        {
           // Console.WriteLine("[DEBUG] Headline 'ביטויים בקידום' not found.");
            return;
        }

        // 2️⃣ Find the first table that comes after that headline
        var table = keywordHeadline.ElementsAfter().OfType<Wp.Table>().FirstOrDefault();
        if (table == null)
        {
           // Console.WriteLine("[DEBUG] Table under 'ביטויים בקידום' not found.");
            return;
        }

        // 3️⃣ Get all rows
        var allRows = table.Elements<Wp.TableRow>().ToList();
        if (allRows.Count < 2)
        {
            //Console.WriteLine("[DEBUG] Table has no data rows.");
            return;
        }

        // Header row
        var headerRow = allRows[0];
        var headerCells = headerRow.Elements<Wp.TableCell>().ToList();

        // Find the column index for "חיפושים מקומיים"
        int targetColumnIndex = -1;
        for (int i = 0; i < headerCells.Count; i++)
        {
            var headerText = headerCells[i].InnerText.Trim();
            if (headerText == "חיפושים מקומיים")
            {
                targetColumnIndex = i;
                break;
            }
        }

        if (targetColumnIndex == -1)
        {
           // Console.WriteLine("[DEBUG] Column 'חיפושים מקומיים' not found.");
            return;
        }

        // 4️⃣ Iterate all data rows (skip header)
        for (int rowIndex = 1; rowIndex < allRows.Count; rowIndex++)
        {
            var row = allRows[rowIndex];
            var cells = row.Elements<Wp.TableCell>().ToList();

            if (targetColumnIndex >= cells.Count)
                {

                    // Console.WriteLine($"[DEBUG] Row {rowIndex} does not have enough cells.");
                  
                continue;
            }

            var cell = cells[targetColumnIndex];
            string cellText = cell.InnerText.Trim();

           // Console.WriteLine($"[DEBUG] Row {rowIndex} content before: '{cellText}'");

            if (cellText == "מעט חיפושים")
            {
                // Clear existing content
                cell.RemoveAllChildren<Wp.Paragraph>();

                // Add new paragraph with number 10 and font size 10pt
                var newParagraph = new Wp.Paragraph(
                    new Wp.Run(
                        new Wp.RunProperties(
                            new Wp.RunFonts { Ascii = "Arial", HighAnsi = "Arial" },
                            new Wp.FontSize { Val = "20" } // 10pt
                        ),
                        new Wp.Text("10")
                    )
                );

                cell.Append(newParagraph);
               // Console.WriteLine($"[DEBUG] Replaced 'מעט חיפושים' with 10 in row {rowIndex}.");
            }
            else
            {
              //  Console.WriteLine($"[DEBUG] No replacement needed in row {rowIndex}.");
            }

           // Console.WriteLine($"[DEBUG] Row {rowIndex} content after: '{cell.InnerText.Trim()}'");
        }

        Console.WriteLine("✅ Finished replacing 'מעט חיפושים' with 10 in all applicable rows.");
    }





    }  
}