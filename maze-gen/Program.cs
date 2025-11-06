using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using OfficeOpenXml;

namespace MazeFromExcel
{
    class Cell
    {
        public int X { get; set; }
        public int Y { get; set; }

        public override string ToString() => $"cell({X},{Y}).";
    }

    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length != 2)
            {
                Console.Error.WriteLine("Bad arguments o7");
                return;
            }
            string filePath = args[0];
            int sheetNum;
            try
            {
                sheetNum = int.Parse(args[1]);
            }
            catch (Exception)
            {
                Console.Error.WriteLine("Bad sheet number");
                return;
            }

            ExcelPackage.License.SetNonCommercialPersonal("Miami Student");
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                if (package.Workbook.Worksheets.Count <= sheetNum || sheetNum < 0)
                {
                    Console.Error.WriteLine("Bad sheet number");
                    return;
                }
                var ws = package.Workbook.Worksheets[sheetNum];

                int maxRow = ws.Dimension.End.Row;
                int maxCol = ws.Dimension.End.Column;
                Console.WriteLine($"Max Row: {maxRow}");
                Console.WriteLine($"Max Col: {maxCol}");

                var moveableCells = new HashSet<(int, int)>();
                var doors = new HashSet<(int, int)>();
                var walls = new HashSet<(int, int)>();
                var freezeTraps = new HashSet<(int, int)>();
                var levers = new HashSet<(int, int)>();

                // If I wanted to start my index at 1 I would have used lua....
                // Sadly x sheets start at 1.
                for (int row = 1; row <= maxRow; row++)
                {
                    for (int col = 1; col <= maxCol; col++)
                    {
                        var cell = ws.Cells[row, col];
                        var fillColor = cell.Style.Fill.BackgroundColor.Rgb;
                        var text = cell.Text;

                        if (fillColor == null)
                        {
                            fillColor = "FFFFFFFF";
                        }

                        // ARGB format.....ew
                        bool isWall = fillColor.ToUpper() == "FF000000"; // black
                        bool isFreezeTrap = fillColor.ToUpper() == "FF0000FF"; // blue
                        bool isDoor = fillColor.ToUpper() == "FF5A0E00"; // brown
                        bool isLever = fillColor.ToUpper() == "FFFF00FF"; // magenta

                        bool isMoveable =
                            fillColor.ToUpper() == "FFFFFFFF" || isFreezeTrap || isDoor || isLever; // white

                        if (isMoveable)
                            moveableCells.Add((col, row));
                        if (isLever)
                            levers.Add((col, row));
                        if (isDoor)
                            doors.Add((col, row));
                        if (isWall)
                            walls.Add((col, row));
                        if (isFreezeTrap)
                            freezeTraps.Add((col, row));
                    }
                }

                var adjacents = new List<(int, int, int, int)>();

                foreach (var (x, y) in moveableCells)
                {
                    var neighbors = new (int, int)[]
                    {
                        (x + 1, y),
                        (x - 1, y),
                        (x, y + 1),
                        (x, y - 1),
                    };

                    foreach (var (nx, ny) in neighbors)
                    {
                        if (moveableCells.Contains((nx, ny)))
                            adjacents.Add((x, y, nx, ny));
                    }
                }

                Console.WriteLine($"% {moveableCells.Count} Cells:");
                foreach (var (x, y) in moveableCells)
                    Console.WriteLine($"cell({x},{y}).");

                Console.WriteLine($"\n% {adjacents.Count} Adjacents:");
                foreach (var (x1, y1, x2, y2) in adjacents)
                    Console.WriteLine($"adjacent(({x1},{y1}),({x2},{y2})).");

                Console.WriteLine($"\n% {freezeTraps.Count} Freeze Traps:");
                foreach (var (x, y) in freezeTraps)
                    Console.WriteLine($"freeze_trap({x},{y}).");

                Console.WriteLine($"\n% {doors.Count} Doors:");
                foreach (var (x, y) in doors)
                    Console.WriteLine($"door({x},{y}).");

                Console.WriteLine($"\n% {walls.Count} Walls:");
                foreach (var (x, y) in walls)
                    Console.WriteLine($"wall({x},{y}).");

                Console.WriteLine($"\n% {levers.Count} Levers:");
                foreach (var (x, y) in levers)
                    Console.WriteLine($"lever({x},{y}).");
            }
        }
    }
}
