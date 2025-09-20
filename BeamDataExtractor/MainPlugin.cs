// AutoCAD .NET API namespaces
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Colors;

// Standard .NET namespaces
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

// Helper class to store all data for a single beam segment (used by multi-beam command)
public class BeamSegment
{
    public MText IdText { get; set; }
    public string BeamMark { get; set; }
    public double Width { get; set; }
    public double Depth { get; set; }
    public Polyline BeamBox { get; set; }
    public double MinX { get; set; }
    public double MaxX { get; set; }
    public double CenterY { get; set; }
    public string Level { get; set; }
    public string LeftBottom { get; set; }
    public string MidBottom { get; set; }
    public string RightBottom { get; set; }
    public string LeftTop { get; set; }
    public string MidTop { get; set; }
    public string RightTop { get; set; }
    public string LeftStirrupDia { get; set; }
    public string LeftStirrupSpace { get; set; }
    public string MidStirrupDia { get; set; }
    public string MidStirrupSpace { get; set; }
    public string RightStirrupDia { get; set; }
    public string RightStirrupSpace { get; set; }
    public string LeftAtDist { get; set; }
    public string RightAtDist { get; set; }

    public BeamSegment()
    {
        BeamMark = Level = LeftBottom = MidBottom = RightBottom = LeftTop = MidTop = RightTop = "";
        LeftStirrupDia = LeftStirrupSpace = MidStirrupDia = MidStirrupSpace = RightStirrupDia = RightStirrupSpace = "";
        LeftAtDist = RightAtDist = "";
    }
}

public class MainPlugin
{
    // ---------- small helpers ----------
    private const double XTol = 1e-3;
    private const double YTol = 1e-3;

    private static bool ContainsCI(string src, string value)
    {
        if (src == null || value == null) return false;
        return src.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private static string ReplaceCI(string src, string oldValue, string newValue)
    {
        if (string.IsNullOrEmpty(src) || string.IsNullOrEmpty(oldValue)) return src ?? "";
        return Regex.Replace(src, Regex.Escape(oldValue), newValue ?? "", RegexOptions.IgnoreCase);
    }

    private static bool IsVertical(Line l)
        => Math.Abs(l.StartPoint.X - l.EndPoint.X) < XTol && Math.Abs(l.StartPoint.Y - l.EndPoint.Y) > YTol;

    private static bool IsHorizontal(Line l)
        => Math.Abs(l.StartPoint.Y - l.EndPoint.Y) < YTol && Math.Abs(l.StartPoint.X - l.EndPoint.X) > XTol;

    private static IEnumerable<double> UniqueXs(IEnumerable<double> xs, int round = 3)
    {
        return xs.GroupBy(x => Math.Round(x, round))
                 .Select(g => g.Average())
                 .OrderBy(v => v);
    }

    private static double Mid(double a, double b) => (a + b) * 0.5;

    private static string MergeBars(string existing, IEnumerable<string> extras)
    {
        var set = new HashSet<string>(SplitBars(existing), StringComparer.OrdinalIgnoreCase);
        foreach (var s in extras)
        {
            foreach (var t in SplitBars(s))
                set.Add(t);
        }
        return string.Join(", ", set.Where(z => !string.IsNullOrWhiteSpace(z)));
    }

    private static IEnumerable<string> SplitBars(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) yield break;
        foreach (var part in text.Split(new[] { ',', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries))
        {
            var p = part.Trim();
            if (!string.IsNullOrEmpty(p)) yield return p;
        }
    }

    // Span containers (avoid tuple types so it compiles everywhere)
    private struct PolySpan
    {
        public double MinX, MaxX, CenterY;
        public Polyline Poly;
    }
    private struct Span
    {
        public double MinX, MaxX, CenterY;
    }

    // ====================================================================================================
    // === YOUR ORIGINAL, WORKING SINGLE-BEAM COMMAND (UNCHANGED) ===
    // ====================================================================================================
    [CommandMethod("BEAMTABLE")]
    public void ExtractBeamData()
    {
        Document doc = Application.DocumentManager.MdiActiveDocument;
        if (doc == null) return;
        Database db = doc.Database;
        Editor ed = doc.Editor;

        var pso = new PromptSelectionOptions { MessageForAdding = "\nSelect all objects for one beam diagram: " };
        PromptSelectionResult psr = ed.GetSelection(pso);
        if (psr.Status != PromptStatus.OK) return;

        string beamMark = "";
        double width = 0.0, depth = 0.0;
        string level = "";
        var reinforcementTexts = new List<MText>();
        var stirrupTexts = new List<MText>();
        var dimensionObjects = new List<Dimension>();
        var steelLines = new List<Line>();
        var boxLines = new List<Line>();

        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
            foreach (SelectedObject so in psr.Value)
            {
                if (so == null) continue;
                Entity ent = tr.GetObject(so.ObjectId, OpenMode.ForRead) as Entity;
                if (ent == null) continue;

                if (ent.Layer == "B_NO" && ent is MText beamNoText)
                {
                    string s = beamNoText.Text ?? "";
                    var tokens = s.Replace("X", "x").Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                    beamMark = tokens.FirstOrDefault() ?? "";
                    int xIdx = Array.IndexOf(tokens, "x");
                    if (xIdx > 0 && xIdx < tokens.Length - 1)
                    {
                        double.TryParse(tokens[xIdx - 1], NumberStyles.Any, CultureInfo.InvariantCulture, out width);
                        double.TryParse(tokens[xIdx + 1], NumberStyles.Any, CultureInfo.InvariantCulture, out depth);
                    }
                    else
                    {
                        string t2 = tokens.FirstOrDefault(t => (t ?? "").ToLowerInvariant().Contains("x"));
                        if (!string.IsNullOrEmpty(t2))
                        {
                            var parts = t2.Split(new[] { 'x', 'X' }, StringSplitOptions.RemoveEmptyEntries);
                            if (parts.Length == 2)
                            {
                                double.TryParse(parts[0], NumberStyles.Any, CultureInfo.InvariantCulture, out width);
                                double.TryParse(parts[1], NumberStyles.Any, CultureInfo.InvariantCulture, out depth);
                            }
                        }
                    }
                }
                else if (ent.Layer == "B_TEXT" && ent is MText reinText) { reinforcementTexts.Add(reinText); }
                else if (ent.Layer == "ring text" && ent is MText stirrupText) { stirrupTexts.Add(stirrupText); }
                else if (ent.Layer == "B_DIM" && ent is Dimension dim) { dimensionObjects.Add(dim); }
                else if (ent.Layer == "LEVEL" && ent is MText levelText) { level = levelText.Text ?? ""; }
                else if (ent.Layer == "B_BOX" && ent is Line boxLine) { boxLines.Add(boxLine); }
                else if (ent.Layer == "B_STEEL" && ent is Line steelLine) { steelLines.Add(steelLine); }
            }
            tr.Commit();
        }

        double beamMinX = 0, beamMaxX = 0, beamCenterY = 0;
        var allBeamLines = new List<Line>();
        allBeamLines.AddRange(boxLines);
        allBeamLines.AddRange(steelLines);

        if (allBeamLines.Any())
        {
            var allXCoords = allBeamLines.SelectMany(l => new[] { l.StartPoint.X, l.EndPoint.X });
            var allYCoords = allBeamLines.SelectMany(l => new[] { l.StartPoint.Y, l.EndPoint.Y });
            beamMinX = allXCoords.Min();
            beamMaxX = allXCoords.Max();
            beamCenterY = allYCoords.Average();
        }

        double beamLength = beamMaxX - beamMinX;
        double leftZoneMaxX = beamMinX + beamLength / 3.0;
        double rightZoneMinX = beamMaxX - beamLength / 3.0;

        var leftBottomBars = new List<string>();
        var midBottomBars = new List<string>();
        var rightBottomBars = new List<string>();
        var allBottomBarTexts = reinforcementTexts.Where(t => GetVisualCenter(t).Y < beamCenterY).ToList();
        var throughoutBottomBars = allBottomBarTexts.Where(t => ContainsCI(t.Text, "(T)")).Select(t => t.Text.Trim()).ToList();
        var curtailBottomBars = allBottomBarTexts.Where(t => ContainsCI(t.Text, "(C)") || ContainsCI(t.Text, "EXTRA")).ToList();
        leftBottomBars.AddRange(throughoutBottomBars);
        midBottomBars.AddRange(throughoutBottomBars);
        rightBottomBars.AddRange(throughoutBottomBars);
        foreach (var curtailText in curtailBottomBars)
        {
            double textX = GetVisualCenter(curtailText).X;
            if (textX <= leftZoneMaxX) { leftBottomBars.Add(curtailText.Text.Trim()); }
            else if (textX >= rightZoneMinX) { rightBottomBars.Add(curtailText.Text.Trim()); }
            else { midBottomBars.Add(curtailText.Text.Trim()); }
        }
        string finalLeftBottom = string.Join(", ", leftBottomBars.Distinct());
        string finalMidBottom = string.Join(", ", midBottomBars.Distinct());
        string finalRightBottom = string.Join(", ", rightBottomBars.Distinct());

        var leftTopBars = new List<string>();
        var midTopBars = new List<string>();
        var rightTopBars = new List<string>();
        var allTopBarTexts = reinforcementTexts.Where(t => GetVisualCenter(t).Y >= beamCenterY).ToList();
        var throughoutTopBars = allTopBarTexts.Where(t => ContainsCI(t.Text, "(T)")).Select(t => t.Text.Trim()).ToList();
        var curtailTopBars = allTopBarTexts.Where(t => ContainsCI(t.Text, "(C)") || ContainsCI(t.Text, "EXTRA")).ToList();
        leftTopBars.AddRange(throughoutTopBars);
        midTopBars.AddRange(throughoutTopBars);
        rightTopBars.AddRange(throughoutTopBars);
        foreach (var curtailText in curtailTopBars)
        {
            double textX = GetVisualCenter(curtailText).X;
            if (textX <= leftZoneMaxX) { leftTopBars.Add(curtailText.Text.Trim()); }
            else if (textX >= rightZoneMinX) { rightTopBars.Add(curtailText.Text.Trim()); }
            else { midTopBars.Add(curtailText.Text.Trim()); }
        }
        string finalLeftTop = string.Join(", ", leftTopBars.Distinct());
        string finalMidTop = string.Join(", ", midTopBars.Distinct());
        string finalRightTop = string.Join(", ", rightTopBars.Distinct());

        string leftStirrupDia = "", midStirrupDia = "", rightStirrupDia = "";
        string leftStirrupSpace = "", midStirrupSpace = "", rightStirrupSpace = "";
        foreach (var st in stirrupTexts)
        {
            var parts = (st.Text ?? "").Split('@');
            if (parts.Length == 2)
            {
                if (GetVisualCenter(st).X <= leftZoneMaxX) { leftStirrupDia = parts[0].Trim(); leftStirrupSpace = parts[1].Trim(); }
                else if (GetVisualCenter(st).X >= rightZoneMinX) { rightStirrupDia = parts[0].Trim(); rightStirrupSpace = parts[1].Trim(); }
                else { midStirrupDia = parts[0].Trim(); midStirrupSpace = parts[1].Trim(); }
            }
        }

        string leftAtDist = "", rightAtDist = "";
        foreach (var dim in dimensionObjects)
        {
            if (dim.TextPosition.X <= leftZoneMaxX) { leftAtDist = (dim.Measurement / 1000.0).ToString("F1", CultureInfo.InvariantCulture); }
            else if (dim.TextPosition.X >= rightZoneMinX) { rightAtDist = (dim.Measurement / 1000.0).ToString("F1", CultureInfo.InvariantCulture); }
        }

        string[] rowData = new string[24];
        rowData[0] = beamMark; rowData[1] = (width / 1000.0).ToString("F2", CultureInfo.InvariantCulture); rowData[2] = (depth / 1000.0).ToString("F2", CultureInfo.InvariantCulture); rowData[3] = level;
        rowData[4] = finalLeftBottom; rowData[5] = ""; rowData[6] = finalMidBottom; rowData[7] = ""; rowData[8] = finalRightBottom; rowData[9] = ""; rowData[10] = "";
        rowData[11] = finalLeftTop; rowData[12] = leftAtDist; rowData[13] = finalMidTop; rowData[14] = finalRightTop; rowData[15] = rightAtDist;
        rowData[16] = ""; rowData[17] = "2";
        rowData[18] = leftStirrupDia; rowData[19] = leftStirrupSpace; rowData[20] = midStirrupDia; rowData[21] = midStirrupSpace; rowData[22] = rightStirrupDia; rowData[23] = rightStirrupSpace;

        PromptPointResult ppr = ed.GetPoint("\nPick an insertion point for the table: ");
        if (ppr.Status != PromptStatus.OK) return;

        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
            BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
            Table table = new Table();
            table.SetDatabaseDefaults(db);
            table.Position = ppr.Value;
            int numRows = 3; int numCols = 24;
            table.SetSize(numRows, numCols);
            Color headerBgColor = Color.FromRgb(220, 230, 241);

            for (int r = 0; r < numRows; r++)
            {
                for (int c = 0; c < numCols; c++)
                {
                    table.Cells[r, c].Alignment = CellAlignment.MiddleCenter;
                    table.Cells[r, c].TextHeight = 2.5;
                    if (r < 2) table.Cells[r, c].BackgroundColor = headerBgColor;
                }
            }

            double[] colWidths = { 18, 15, 15, 18, 25, 25, 25, 25, 25, 25, 20, 25, 25, 25, 25, 25, 15, 20, 25, 25, 25, 25, 25, 25 };
            for (int i = 0; i < numCols; i++) table.Columns[i].Width = colWidths[i];
            for (int i = 0; i < numRows; i++) table.Rows[i].Height = 12;

            table.Cells[0, 0].TextString = "Beam Details"; table.Cells[0, 0].TextHeight = 3.5;
            table.Cells[0, 4].TextString = "Bottom"; table.Cells[0, 4].TextHeight = 3.5;
            table.Cells[0, 11].TextString = "Top"; table.Cells[0, 11].TextHeight = 3.5;
            table.Cells[0, 16].TextString = "Stirrups"; table.Cells[0, 16].TextHeight = 3.5;

            table.MergeCells(CellRange.Create(table, 0, 0, 0, 3));
            table.MergeCells(CellRange.Create(table, 0, 4, 0, 10));
            table.MergeCells(CellRange.Create(table, 0, 11, 0, 15));
            table.MergeCells(CellRange.Create(table, 0, 16, 0, 23));

            string[] headers = { "BeamId", "Width", "Depth", "Level", "Left_bottom", "Bottom left at(dist)", "Mid_bottom", "Curtail at(dist)", "Right_Bottom", "Bottom right at(dist)", "bent up", "Left_top", "Left at(dist)", "Mid_top", "Right_top", "Right at(dist)", "SFR", "Shear Stirrups Leg", "Shear Stirrups dia(L)", "Left Space Stirrups", "Shear Stirrups dia(M)", "Mid Space Stirrups", "Shear Stirrups dia(R)", "Right Space Stirrups" };
            for (int i = 0; i < headers.Length; i++) { table.Cells[1, i].TextString = headers[i]; }
            for (int i = 0; i < rowData.Length; i++) { table.Cells[2, i].TextString = rowData[i]; }

            table.GenerateLayout();
            btr.AppendEntity(table);
            tr.AddNewlyCreatedDBObject(table, true);
            tr.Commit();
        }
        ed.WriteMessage("\nStyled table with full top and bottom bar logic generated successfully.");
    }

    // ====================================================================================================
    // === NEW, SEPARATE MULTI-BEAM COMMAND (FIXED: shared boundary band) ===
    // ====================================================================================================
    [CommandMethod("MULTIBEAMTABLE")]
    public void ExtractMultiBeamData()
    {
        Document doc = Application.DocumentManager.MdiActiveDocument;
        if (doc == null) return;
        Database db = doc.Database;
        Editor ed = doc.Editor;

        var pso = new PromptSelectionOptions { MessageForAdding = "\nSelect all objects for the multi-beam diagram: " };
        PromptSelectionResult psr = ed.GetSelection(pso);
        if (psr.Status != PromptStatus.OK) return;

        var beamIdTexts = new List<MText>();
        var reinforcementTexts = new List<MText>();
        var stirrupTexts = new List<MText>();
        var beamBoxPolys = new List<Polyline>(); // red boxes as polylines
        var beamBoxLines = new List<Line>();     // red boxes as lines (vertical/horizontal)
        var dimensionObjects = new List<Dimension>();
        string level = "";

        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
            foreach (SelectedObject so in psr.Value)
            {
                if (so == null) continue;
                Entity ent = tr.GetObject(so.ObjectId, OpenMode.ForRead) as Entity;
                if (ent == null) continue;

                if (ent.Layer == "B_NO" && ent is MText beamNoText) beamIdTexts.Add(beamNoText);
                else if (ent.Layer == "B_TEXT" && ent is MText reinText) reinforcementTexts.Add(reinText);
                else if (ent.Layer == "ring text" && ent is MText stirrupText) stirrupTexts.Add(stirrupText);
                else if (ent.Layer == "B_BOX" && ent is Polyline p) beamBoxPolys.Add(p);
                else if (ent.Layer == "B_BOX" && ent is Line bl) beamBoxLines.Add(bl);
                else if (ent.Layer == "B_DIM" && ent is Dimension dim) dimensionObjects.Add(dim);
                else if (ent.Layer == "LEVEL" && ent is MText levelText) level = levelText.Text ?? "";
            }
            tr.Commit();
        }

        if (!beamIdTexts.Any()) { ed.WriteMessage("\nNo Beam IDs found on layer 'B_NO'."); return; }

        // Build spans from polylines (preferred)
        var polySpans = new List<PolySpan>();
        foreach (var pl in beamBoxPolys)
        {
            if (!pl.Bounds.HasValue) continue;
            var b = pl.Bounds.Value;
            polySpans.Add(new PolySpan
            {
                MinX = b.MinPoint.X,
                MaxX = b.MaxPoint.X,
                CenterY = Mid(b.MinPoint.Y, b.MaxPoint.Y),
                Poly = pl
            });
        }

        // Fallback spans from vertical/horizontal red lines
        var lineSpans = new List<Span>();
        if (polySpans.Count == 0 && beamBoxLines.Any())
        {
            var vXs = UniqueXs(beamBoxLines.Where(IsVertical).Select(l => l.StartPoint.X)).ToList();
            if (vXs.Count >= 2)
            {
                var horiz = beamBoxLines.Where(IsHorizontal).ToList();
                for (int i = 0; i < vXs.Count - 1; i++)
                {
                    double minX = vXs[i], maxX = vXs[i + 1];
                    if (maxX - minX < 10 * XTol) continue;

                    var hYs = new List<double>();
                    foreach (var hl in horiz)
                    {
                        double x1 = Math.Min(hl.StartPoint.X, hl.EndPoint.X);
                        double x2 = Math.Max(hl.StartPoint.X, hl.EndPoint.X);
                        if (x2 >= minX && x1 <= maxX) hYs.Add(hl.StartPoint.Y);
                    }

                    double centerY;
                    if (hYs.Count >= 2) centerY = Mid(hYs.Min(), hYs.Max());
                    else
                    {
                        var reinInSpanYs = reinforcementTexts
                            .Select(t => GetVisualCenter(t))
                            .Where(c => c.X >= minX - XTol && c.X <= maxX + XTol)
                            .Select(c => c.Y)
                            .OrderBy(y => y)
                            .ToList();
                        centerY = reinInSpanYs.Count > 0 ? reinInSpanYs[reinInSpanYs.Count / 2] : 0.0;
                    }

                    lineSpans.Add(new Span { MinX = minX, MaxX = maxX, CenterY = centerY });
                }
            }
        }

        // Map each Beam ID to a span
        var beamSegments = new List<BeamSegment>();
        foreach (var idText in beamIdTexts)
        {
            var seg = new BeamSegment { IdText = idText, Level = level };

            // Parse BeamId/Width/Depth like single-beam
            string s = idText.Text ?? "";
            var tokens = s.Replace("X", "x").Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            seg.BeamMark = tokens.FirstOrDefault() ?? "";
            int xIdx = Array.IndexOf(tokens, "x");
            if (xIdx > 0 && xIdx < tokens.Length - 1)
            {
                double.TryParse(tokens[xIdx - 1], NumberStyles.Any, CultureInfo.InvariantCulture, out double w);
                double.TryParse(tokens[xIdx + 1], NumberStyles.Any, CultureInfo.InvariantCulture, out double d);
                seg.Width = w; seg.Depth = d;
            }
            else
            {
                string combined = string.Join("", tokens);
                var match = Regex.Match(combined, @"(\d+)[xX](\d+)");
                if (match.Success)
                {
                    double.TryParse(match.Groups[1].Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double w);
                    double.TryParse(match.Groups[2].Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double d);
                    seg.Width = w; seg.Depth = d;
                }
            }

            // Choose span: prefer poly, else line spans; fallback to nearest
            var idC = GetVisualCenter(idText);
            bool mapped = false;

            if (polySpans.Count > 0)
            {
                foreach (var sp in polySpans)
                {
                    if (idC.X >= sp.MinX - XTol && idC.X <= sp.MaxX + XTol)
                    {
                        seg.BeamBox = sp.Poly;
                        seg.MinX = sp.MinX; seg.MaxX = sp.MaxX; seg.CenterY = sp.CenterY;
                        mapped = true;
                        break;
                    }
                }
            }

            if (!mapped && lineSpans.Count > 0)
            {
                foreach (var sp in lineSpans)
                {
                    if (idC.X >= sp.MinX - XTol && idC.X <= sp.MaxX + XTol)
                    {
                        seg.MinX = sp.MinX; seg.MaxX = sp.MaxX; seg.CenterY = sp.CenterY;
                        mapped = true;
                        break;
                    }
                }
            }

            if (!mapped)
            {
                if (polySpans.Count > 0)
                {
                    var nearest = polySpans.OrderBy(sp => Math.Abs(Mid(sp.MinX, sp.MaxX) - idC.X)).First();
                    seg.BeamBox = nearest.Poly;
                    seg.MinX = nearest.MinX; seg.MaxX = nearest.MaxX; seg.CenterY = nearest.CenterY;
                }
                else if (beamBoxLines.Count > 0)
                {
                    double minX = beamBoxLines.Min(l => Math.Min(l.StartPoint.X, l.EndPoint.X));
                    double maxX = beamBoxLines.Max(l => Math.Max(l.StartPoint.X, l.EndPoint.X));
                    double minY = beamBoxLines.Min(l => Math.Min(l.StartPoint.Y, l.EndPoint.Y));
                    double maxY = beamBoxLines.Max(l => Math.Max(l.StartPoint.Y, l.EndPoint.Y));
                    seg.MinX = minX; seg.MaxX = maxX; seg.CenterY = Mid(minY, maxY);
                }
                else
                {
                    seg.MinX = idC.X - 1000; seg.MaxX = idC.X + 1000; seg.CenterY = idC.Y;
                }
            }

            beamSegments.Add(seg);
        }

        // Left-to-right for shared-bar logic
        beamSegments = beamSegments.OrderBy(sg => sg.MinX).ToList();

        // Extract per-segment bars/stirrups (same rules as single-beam)
        foreach (var segment in beamSegments)
        {
            if (segment.MaxX <= segment.MinX) continue;

            double beamLength = segment.MaxX - segment.MinX;
            double leftZoneMaxX = segment.MinX + beamLength / 3.0;
            double rightZoneMinX = segment.MaxX - beamLength / 3.0;

            var relevantReinf = reinforcementTexts
                .Where(t =>
                {
                    var c = GetVisualCenter(t);
                    return c.X >= segment.MinX - XTol && c.X <= segment.MaxX + XTol;
                })
                .ToList();

            // Bottom bars
            var bottomTexts = relevantReinf.Where(t => GetVisualCenter(t).Y < segment.CenterY).ToList();
            var throughBottom = bottomTexts.Where(t => ContainsCI(t.Text, "(T)")).Select(t => t.Text.Trim()).ToList();
            var curtailBottom = bottomTexts.Where(t => ContainsCI(t.Text, "(C)") || ContainsCI(t.Text, "EXTRA")).ToList();

            var leftBottom = new List<string>(throughBottom);
            var midBottom = new List<string>(throughBottom);
            var rightBottom = new List<string>(throughBottom);

            foreach (var ct in curtailBottom)
            {
                double x = GetVisualCenter(ct).X;
                if (x <= leftZoneMaxX) leftBottom.Add(ct.Text.Trim());
                else if (x >= rightZoneMinX) rightBottom.Add(ct.Text.Trim());
                else midBottom.Add(ct.Text.Trim());
            }
            segment.LeftBottom = string.Join(", ", leftBottom.Distinct());
            segment.MidBottom = string.Join(", ", midBottom.Distinct());
            segment.RightBottom = string.Join(", ", rightBottom.Distinct());

            // Top bars
            var topTexts = relevantReinf.Where(t => GetVisualCenter(t).Y >= segment.CenterY).ToList();
            var throughTop = topTexts.Where(t => ContainsCI(t.Text, "(T)")).Select(t => t.Text.Trim()).ToList();
            var curtailTop = topTexts.Where(t => ContainsCI(t.Text, "(C)") || ContainsCI(t.Text, "EXTRA")).ToList();

            var leftTop = new List<string>(throughTop);
            var midTop = new List<string>(throughTop);
            var rightTop = new List<string>(throughTop);

            foreach (var ct in curtailTop)
            {
                double x = GetVisualCenter(ct).X;
                if (x <= leftZoneMaxX) leftTop.Add(ct.Text.Trim());
                else if (x >= rightZoneMinX) rightTop.Add(ct.Text.Trim());
                else midTop.Add(ct.Text.Trim());
            }
            segment.LeftTop = string.Join(", ", leftTop.Distinct());
            segment.MidTop = string.Join(", ", midTop.Distinct());
            segment.RightTop = string.Join(", ", rightTop.Distinct());

            // Stirrups
            foreach (var st in stirrupTexts)
            {
                var c = GetVisualCenter(st);
                if (c.X < segment.MinX - XTol || c.X > segment.MaxX + XTol) continue;

                var parts = (st.Text ?? "").Split('@');
                if (parts.Length != 2) continue;

                if (c.X <= leftZoneMaxX) { segment.LeftStirrupDia = parts[0].Trim(); segment.LeftStirrupSpace = parts[1].Trim(); }
                else if (c.X >= rightZoneMinX) { segment.RightStirrupDia = parts[0].Trim(); segment.RightStirrupSpace = parts[1].Trim(); }
                else { segment.MidStirrupDia = parts[0].Trim(); segment.MidStirrupSpace = parts[1].Trim(); }
            }

            // Distances (left/right at top)
            foreach (var dim in dimensionObjects)
            {
                if (dim.TextPosition.X < segment.MinX - XTol || dim.TextPosition.X > segment.MaxX + XTol) continue;

                if (dim.TextPosition.X <= leftZoneMaxX)
                    segment.LeftAtDist = (dim.Measurement / 1000.0).ToString("F1", CultureInfo.InvariantCulture);
                else if (dim.TextPosition.X >= rightZoneMinX)
                    segment.RightAtDist = (dim.Measurement / 1000.0).ToString("F1", CultureInfo.InvariantCulture);
            }
        }

        // Common / boundary “green bar”: copy anything in the gap OR either side's boundary band into both sides
        for (int i = 0; i < beamSegments.Count - 1; i++)
        {
            var a = beamSegments[i];
            var b = beamSegments[i + 1];

            // zone thresholds
            double aLen = a.MaxX - a.MinX;
            double bLen = b.MaxX - b.MinX;
            double aRightZoneMinX = a.MaxX - aLen / 3.0;
            double bLeftZoneMaxX = b.MinX + bLen / 3.0;

            // Anything top-side either in the literal gap OR within each side's boundary zone
            var sharedCandidates = reinforcementTexts.Where(t =>
            {
                var c = GetVisualCenter(t);
                bool isTop = c.Y >= Math.Min(a.CenterY, b.CenterY);
                bool inGap = c.X > a.MaxX && c.X < b.MinX;
                bool inRightBoundaryOfA = c.X >= aRightZoneMinX && c.X <= a.MaxX + XTol;
                bool inLeftBoundaryOfB = c.X <= bLeftZoneMaxX && c.X >= b.MinX - XTol;
                return isTop && (inGap || inRightBoundaryOfA || inLeftBoundaryOfB);
            })
            .Select(t => t.Text.Trim())
            .ToList();

            if (sharedCandidates.Any())
            {
                // Add to both with proper de-duplication
                a.RightTop = MergeBars(a.RightTop, sharedCandidates);
                b.LeftTop = MergeBars(b.LeftTop, sharedCandidates);

                // write back (class reference not required, but explicitness doesn’t hurt readability)
                beamSegments[i] = a;
                beamSegments[i + 1] = b;
            }
        }

        // Build table
        PromptPointResult ppr = ed.GetPoint("\nPick an insertion point for the table: ");
        if (ppr.Status != PromptStatus.OK) return;

        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
            BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
            Table table = new Table();
            table.SetDatabaseDefaults(db);
            table.Position = ppr.Value;
            int numDataRows = beamSegments.Count;
            int numHeaderRows = 2;
            int numCols = 24;
            table.SetSize(numHeaderRows + numDataRows, numCols);
            Color headerBgColor = Color.FromRgb(220, 230, 241);

            for (int r = 0; r < table.Rows.Count; r++)
            {
                for (int c = 0; c < numCols; c++)
                {
                    table.Cells[r, c].Alignment = CellAlignment.MiddleCenter;
                    table.Cells[r, c].TextHeight = 2.5;
                    if (r < numHeaderRows) table.Cells[r, c].BackgroundColor = headerBgColor;
                }
            }

            double[] colWidths = { 18, 15, 15, 18, 25, 25, 25, 25, 25, 25, 20, 25, 25, 25, 25, 25, 15, 20, 25, 25, 25, 25, 25, 25 };
            for (int i = 0; i < numCols; i++) table.Columns[i].Width = colWidths[i];
            for (int i = 0; i < table.Rows.Count; i++) table.Rows[i].Height = 12;

            table.Cells[0, 0].TextString = "Beam Details"; table.Cells[0, 0].TextHeight = 3.5;
            table.Cells[0, 4].TextString = "Bottom"; table.Cells[0, 4].TextHeight = 3.5;
            table.Cells[0, 11].TextString = "Top"; table.Cells[0, 11].TextHeight = 3.5;
            table.Cells[0, 16].TextString = "Stirrups"; table.Cells[0, 16].TextHeight = 3.5;
            table.MergeCells(CellRange.Create(table, 0, 0, 0, 3));
            table.MergeCells(CellRange.Create(table, 0, 4, 0, 10));
            table.MergeCells(CellRange.Create(table, 0, 11, 0, 15));
            table.MergeCells(CellRange.Create(table, 0, 16, 0, 23));

            string[] headers = {
                "BeamId", "Width", "Depth", "Level",
                "Left_bottom", "Bottom left at(dist)", "Mid_bottom", "Curtail at(dist)", "Right_Bottom", "Bottom right at(dist)", "bent up",
                "Left_top", "Left at(dist)", "Mid_top", "Right_top", "Right at(dist)",
                "SFR", "Shear Stirrups Leg", "Shear Stirrups dia(L)", "Left Space Stirrups", "Shear Stirrups dia(M)", "Mid Space Stirrups", "Shear Stirrups dia(R)", "Right Space Stirrups"
            };
            for (int i = 0; i < headers.Length; i++) table.Cells[1, i].TextString = headers[i];

            for (int i = 0; i < numDataRows; i++)
            {
                var seg = beamSegments[i];
                int row = i + numHeaderRows;

                table.Cells[row, 0].TextString = seg.BeamMark;
                table.Cells[row, 1].TextString = (seg.Width / 1000.0).ToString("F2", CultureInfo.InvariantCulture);
                table.Cells[row, 2].TextString = (seg.Depth / 1000.0).ToString("F2", CultureInfo.InvariantCulture);
                table.Cells[row, 3].TextString = seg.Level;

                // Bottom
                table.Cells[row, 4].TextString = seg.LeftBottom;
                table.Cells[row, 6].TextString = seg.MidBottom;
                table.Cells[row, 8].TextString = seg.RightBottom;

                // Top
                table.Cells[row, 11].TextString = seg.LeftTop;
                table.Cells[row, 13].TextString = seg.MidTop;
                table.Cells[row, 14].TextString = seg.RightTop;

                // Distances
                table.Cells[row, 12].TextString = seg.LeftAtDist;
                table.Cells[row, 15].TextString = seg.RightAtDist;

                // Stirrups
                table.Cells[row, 17].TextString = "2";
                table.Cells[row, 18].TextString = seg.LeftStirrupDia;
                table.Cells[row, 19].TextString = seg.LeftStirrupSpace;
                table.Cells[row, 20].TextString = seg.MidStirrupDia;
                table.Cells[row, 21].TextString = seg.MidStirrupSpace;
                table.Cells[row, 22].TextString = seg.RightStirrupDia;
                table.Cells[row, 23].TextString = seg.RightStirrupSpace;
            }

            table.GenerateLayout();
            btr.AppendEntity(table);
            tr.AddNewlyCreatedDBObject(table, true);
            tr.Commit();
        }

        ed.WriteMessage($"\nMulti-beam table with {beamSegments.Count} segments generated successfully.");
    }

    // ---------- centers ----------
    private Point3d GetVisualCenter(MText text)
    {
        if (text == null || !text.Bounds.HasValue) return Point3d.Origin;
        Point3d minPoint = text.Bounds.Value.MinPoint;
        Point3d maxPoint = text.Bounds.Value.MaxPoint;
        return new Point3d((minPoint.X + maxPoint.X) / 2.0, (minPoint.Y + maxPoint.Y) / 2.0, (minPoint.Z + maxPoint.Z) / 2.0);
    }

    private Point3d GetVisualCenter(Polyline pline)
    {
        if (pline == null || !pline.Bounds.HasValue) return Point3d.Origin;
        Point3d minPoint = pline.Bounds.Value.MinPoint;
        Point3d maxPoint = pline.Bounds.Value.MaxPoint;
        return new Point3d((minPoint.X + maxPoint.X) / 2.0, (minPoint.Y + maxPoint.Y) / 2.0, (minPoint.Z + maxPoint.Z) / 2.0);
    }
}
