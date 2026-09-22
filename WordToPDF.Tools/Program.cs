using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordToPDF.Tools;

internal class Program
{
    private const string Font = "Times New Roman";

    private static int Main(string[] args)
    {
        var mode = args.Length > 0 ? args[0].ToLowerInvariant() : "generate-template";

        return mode switch
        {
            "generate-template" => GenerateTemplate(),
            "test" => RunTests(),
            _ => Unknown(mode)
        };
    }

    private static int Unknown(string mode)
    {
        Console.WriteLine($"Unknown mode '{mode}'. Use 'generate-template' or 'test'.");
        return 1;
    }

    private static int RunTests()
    {
        return TestHarness.Run();
    }

    private static int GenerateTemplate()
    {
        string outDir = Path.Combine(Directory.GetCurrentDirectory(), "WordTemplates");
        Directory.CreateDirectory(outDir);
        string outPath = Path.Combine(outDir, "DigitalInvoiceTemplate.docx");

        try
        {
            WriteTemplate(outPath);
        }
        catch (IOException)
        {
            outPath = Path.Combine(outDir, "DigitalInvoiceTemplate_new.docx");
            WriteTemplate(outPath);
        }

        Console.WriteLine($"Template written: {outPath}");
        return 0;
    }

    private static void WriteTemplate(string outPath)
    {

        using (var doc = WordprocessingDocument.Create(outPath, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document();
            var body = new Body();

            body.Append(new SectionProperties(
                new PageSize { Width = 11906, Height = 16838 },
                new PageMargin { Top = 720, Right = 720, Bottom = 720, Left = 720, Header = 576, Footer = 576, Gutter = 0 }));

            body.Append(P("", align: JustificationValues.Right, sizeHalf: 18));
            body.Append(P("SALES TAX INVOICE", bold: true, align: JustificationValues.Center, sizeHalf: 40));
            body.Append(Spacer());

            body.Append(MetaTable());
            body.Append(Spacer());

            body.Append(PartyTable());
            body.Append(Spacer());

            body.Append(TermsTable());
            body.Append(Spacer());

            body.Append(ItemsTable());
            body.Append(Spacer());

            body.Append(FooterTable());

            main.Document.Append(body);
            main.Document.Save();
        }
    }

    private static Table MetaTable()
    {
        return Borderless(
            Row(
                Label("Invoice No :", 1250),
                Box("{InvoiceNo}", 1500),
                Label("Fbr Invoice No :", 1600),
                Box("{FbrInvoiceNo}", 2500),
                Label("Date :", 850),
                Box("{InvoiceDate:dd-MMM-yy}", 1350)));
    }

    private static Table PartyTable()
    {
        string[][] rows =
        {
            new[] { "Name :", "{CompanyName}", "Name :", "{CustomerName}" },
            new[] { "Ph No & Address :", "{CompanyAddress}", "Address :", "{CustomerAddress}" },
            new[] { "ST Reg No.:", "{CompanySTN}", "ST Reg No. :", "{CustomerPhNo}" },
            new[] { "NTN #", "{CompanyNTN}", "NTN #", "{BuyerRegistrationNo}" },
            new[] { "Province :", "{SellerProvince}", "Province :", "{BuyerProvince}" }
        };

        int[] widths = { 1600, 3100, 500, 1600, 3100 };
        var table = new Table(
            new TableProperties(
                new TableWidth { Width = "9900", Type = TableWidthUnitValues.Dxa },
                new TableLayout { Type = TableLayoutValues.Fixed }),
            new TableGrid(widths.Select(w => new GridColumn { Width = w.ToString() })));

        var caption = new TableRow(
            Label("Supplier", 1600),
            Label("", 3100),
            Gap(500),
            Label("Buyer", 1600),
            Label("", 3100));
        table.Append(caption);

        foreach (var r in rows)
        {
            table.Append(new TableRow(
                Label(r[0], 1600, borders: true),
                Value(r[1], 3100, borders: true),
                Gap(500),
                Label(r[2], 1600, borders: true),
                Value(r[3], 3100, borders: true)));
        }

        return table;
    }

    private static Table TermsTable()
    {
        return Borderless(
            Row(
                Label("Terms Of Sale :", 1500),
                Box("{SalesType}", 1600),
                Label("OGPNo:", 1000),
                Box("{OGPNo}", 1500),
                Label("Remarks :", 1100),
                Box("{MasterRemarks}", 3200)));
    }

    private static Table ItemsTable()
    {
        int[] widths = { 650, 1750, 850, 850, 650, 750, 750, 1000, 650, 1000, 1000 };

        string[] headers =
        {
            "Qty", "Description", "PO NO", "Lot Number", "UOM", "HS Code", "Rate",
            "Value Excluding Sales Tax (Rs)", "Rate of Sales Tax", "Total Sales Tax Payable (Rs)",
            "Value Including Sales Tax (Rs)"
        };

        string[] detail =
        {
            "{Qty:N2}", "{ItemDescription}", "{CustomerPono}", "{SRONo}", "{UOM}", "{HsCode}", "{Rate:N2}",
            "{LineTotal:N0}", "{SalesTaxPercentage:N0}", "{SaleTaxAmount:N0}", "{Total:N0}"
        };

        var grid = new TableGrid(widths.Select(w => new GridColumn { Width = w.ToString() }));

        var headerRow = new TableRow();
        headerRow.Append(new TableRowProperties(new CantSplit()));
        headerRow.Append(headers.Select((h, i) => Cell(h, widths[i], bold: true, borders: true,
            align: JustificationValues.Center, sizeHalf: 16)).ToArray());

        var detailRow = new TableRow();
        detailRow.Append(new TableRowProperties(new CantSplit()));
        detailRow.Append(detail.Select((p, i) => Cell(p, widths[i], borders: true,
            align: i == 0 ? JustificationValues.Right : JustificationValues.Left, sizeHalf: 18)).ToArray());

        var totalsRow = new TableRow();
        totalsRow.Append(new TableRowProperties(new CantSplit()));
        totalsRow.Append(Enumerable.Range(0, 6).Select(i => Cell("", widths[i], bold: true, borders: true, sizeHalf: 18))
            .Concat(new[]
            {
                Cell("Total :", widths[6], bold: true, borders: true,
                    align: JustificationValues.Right, sizeHalf: 18),
                Cell("", widths[7], bold: true, borders: true, sizeHalf: 18),
                Cell("", widths[8], bold: true, borders: true, sizeHalf: 18),
                Cell("", widths[9], bold: true, borders: true, sizeHalf: 18),
                Cell("{GrandTotal:N0}", widths[10], bold: true, borders: true,
                    align: JustificationValues.Right, sizeHalf: 18)
            }).ToArray());

        return new Table(
            new TableProperties(
                new TableWidth { Width = "9900", Type = TableWidthUnitValues.Dxa },
                new TableLayout { Type = TableLayoutValues.Fixed },
                new TableBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 6 },
                    new BottomBorder { Val = BorderValues.Single, Size = 6 },
                    new LeftBorder { Val = BorderValues.Single, Size = 6 },
                    new RightBorder { Val = BorderValues.Single, Size = 6 },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 6 },
                    new InsideVerticalBorder { Val = BorderValues.Single, Size = 6 })),
            grid,
            headerRow,
            detailRow,
            totalsRow);
    }

    private static Table FooterTable()
    {
        int[] widths = { 2400, 1800, 5700 };
        var grid = new TableGrid(widths.Select(w => new GridColumn { Width = w.ToString() }));

        var note = new TableCell(
            new TableCellProperties(
                new TableCellWidth { Width = widths[2].ToString(), Type = TableWidthUnitValues.Dxa },
                new VerticalMerge { Val = MergedCellValues.Restart }),
            P("This is a system-generated invoice and does not require any stamp or signature.",
                sizeHalf: 20));

        var row1 = new TableRow(
            Label("Sales Tax :", widths[0]),
            Box("{SaleTaxAmount:N0}", widths[1]),
            note);

        var row2 = new TableRow(
            Label("Further Tax 0% :", widths[0]),
            Box("{FurtherTax:N0}", widths[1]),
            new TableCell(
                new TableCellProperties(
                    new TableCellWidth { Width = widths[2].ToString(), Type = TableWidthUnitValues.Dxa },
                    new VerticalMerge { Val = MergedCellValues.Continue }),
                P("", sizeHalf: 20)));

        var row3 = new TableRow(
            Label("Net Tax Inclusive Value :", widths[0]),
            Box("{GrandTotal:N0}", widths[1]),
            new TableCell(
                new TableCellProperties(
                    new TableCellWidth { Width = widths[2].ToString(), Type = TableWidthUnitValues.Dxa },
                    new VerticalMerge { Val = MergedCellValues.Continue }),
                P("", sizeHalf: 20)));

        return Borderless(new[] { row1, row2, row3 }, widths);
    }

    private static Table Borderless(params TableRow[] rows)
    {
        return Borderless(rows, new[] { 1250, 1500, 1600, 2500, 850, 1350 });
    }

    private static Table Borderless(TableRow[] rows, int[] widths)
    {
        var table = new Table(
            new TableProperties(
                new TableWidth { Width = widths.Sum().ToString(), Type = TableWidthUnitValues.Dxa },
                new TableLayout { Type = TableLayoutValues.Fixed }),
            new TableGrid(widths.Select(w => new GridColumn { Width = w.ToString() })));

        foreach (var row in rows)
        {
            table.Append(row);
        }

        return table;
    }

    private static TableCell Label(string text, int width, bool borders = false)
    {
        return Cell(text, width, bold: false, borders: borders, sizeHalf: 18);
    }

    private static TableCell Value(string text, int width, bool borders = false)
    {
        return Cell(text, width, bold: false, borders: borders, sizeHalf: 18);
    }

    private static TableCell Box(string text, int width)
    {
        return Cell(text, width, bold: false, borders: true, sizeHalf: 18);
    }

    private static TableCell Gap(int width)
    {
        var props = new TableCellProperties(
            new TableCellWidth { Width = width.ToString(), Type = TableWidthUnitValues.Dxa });
        return new TableCell(props, P("", sizeHalf: 18));
    }

    private static TableCell Cell(
        string text,
        int width,
        bool bold = false,
        bool borders = false,
        JustificationValues? align = null,
        int sizeHalf = 18)
    {
        var props = new TableCellProperties(
            new TableCellWidth { Width = width.ToString(), Type = TableWidthUnitValues.Dxa },
            new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center });

        if (borders)
        {
            props.Append(new TableCellBorders(
                new TopBorder { Val = BorderValues.Single, Size = 6 },
                new BottomBorder { Val = BorderValues.Single, Size = 6 },
                new LeftBorder { Val = BorderValues.Single, Size = 6 },
                new RightBorder { Val = BorderValues.Single, Size = 6 }));
        }

        return new TableCell(props, P(text, bold, align, sizeHalf));
    }

    private static Paragraph P(
        string text,
        bool bold = false,
        JustificationValues? align = null,
        int sizeHalf = 18)
    {
        var runProperties = new RunProperties(
            new RunFonts { Ascii = Font, HighAnsi = Font },
            new FontSize { Val = sizeHalf.ToString() });
        if (bold)
        {
            runProperties.AppendChild(new Bold());
        }

        var paragraphProperties = new ParagraphProperties(
            new SpacingBetweenLines { Before = "20", After = "20", Line = "240", LineRule = LineSpacingRuleValues.Auto });
        if (align != null)
        {
            paragraphProperties.AppendChild(new Justification { Val = align.Value });
        }

        return new Paragraph(
            paragraphProperties,
            new Run(runProperties, new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
    }

    private static Paragraph Spacer()
    {
        return new Paragraph(
            new ParagraphProperties(new SpacingBetweenLines { Before = "0", After = "80" }),
            new Run(
                new RunProperties(new RunFonts { Ascii = Font, HighAnsi = Font }, new FontSize { Val = "12" }),
                new Text("")));
    }

    private static TableRow Row(params TableCell[] cells)
    {
        return new TableRow(cells);
    }
}
