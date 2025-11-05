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
            if (args.Length != 1)
            {
                Console.Error.WriteLine("Bad arguments o7");
                return;
            }
            string filePath = args[0];

            ExcelPackage.License.SetNonCommercialPersonal("Miami Student");
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var ws = package.Workbook.Worksheets[0];
                
                int maxRow = ws.Dimension.End.Row;
                int maxCol = ws.Dimension.End.Column;
                Console.WriteLine($"Max Row: {maxRow}");
                Console.WriteLine($"Max Col: {maxCol}");
                var openCells = new HashSet<(int, int)>();

                for (int row = 1; row <= maxRow; row++)
                {
                    for (int col = 1; col <= maxCol; col++)
                    {
                        var cell = ws.Cells[row, col];
                        var fill = cell.Style.Fill.BackgroundColor;

                        bool isWhite = fill.Rgb == null || fill.Rgb.ToUpper() == "FFFFFFFF";
                        bool isBlack = fill.Rgb != null && fill.Rgb.ToUpper() == "FF000000";

                        int x = col;
                        int y = row;

                        if (isWhite && !isBlack)
                            openCells.Add((x, y));
                    }
                }

                var adjacents = new List<(int, int, int, int)>();

                foreach (var (x, y) in openCells)
                {
                    var neighbors = new (int, int)[]
                    {
                        (x+1, y), (x-1, y), (x, y+1), (x, y-1)
                    };

                    foreach (var (nx, ny) in neighbors)
                    {
                        if (openCells.Contains((nx, ny)))
                            adjacents.Add((x, y, nx, ny));
                    }
                }

                Console.WriteLine("Cells:");
                foreach (var (x, y) in openCells)
                    Console.WriteLine($"cell({x},{y}).");

                Console.WriteLine("\nAdjacents:");
                foreach (var (x1, y1, x2, y2) in adjacents)
                    Console.WriteLine($"adjacent(({x1},{y1}),({x2},{y2})).");
            }
        }
    }
}
