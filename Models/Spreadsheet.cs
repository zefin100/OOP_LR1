using System;
using System.Collections.Generic;

namespace SpreadsheetApp.Models
{
    public class Cell
    {
        public string Name { get; }
        public string Expression { get; set; } = string.Empty;
        public object? Value { get; set; }
        public string Error { get; set; } = string.Empty;

        // Список клітинок, на які посилається ця клітинка (за вимогою викладача)
        public List<string> References { get; set; } = new();

        public Cell(string name)
        {
            Name = name;
        }
    }

    public class Spreadsheet
    {
        private readonly Dictionary<string, Cell> _cells = new(StringComparer.OrdinalIgnoreCase);

        public int RowCount { get; private set; }
        public int ColumnCount { get; private set; }

        public Spreadsheet(int rows, int columns)
        {
            RowCount = rows;
            ColumnCount = columns;

            for (int r = 1; r <= rows; r++)
            {
                for (int c = 0; c < columns; c++)
                {
                    string name = GetColumnName(c) + r;
                    _cells[name] = new Cell(name);
                }
            }
        }

        public bool AddRow()
        {
            RowCount++;
            for (int c = 0; c < ColumnCount; c++)
            {
                string name = GetColumnName(c) + RowCount;
                if (!_cells.ContainsKey(name))
                {
                    _cells[name] = new Cell(name);
                }
            }
            return true;
        }

        public bool DeleteRow()
        {
            if (RowCount <= 1) return false; // Захист: залишаємо мінімум 1 рядок

            for (int c = 0; c < ColumnCount; c++)
            {
                string name = GetColumnName(c) + RowCount;
                _cells.Remove(name);
            }
            RowCount--;
            return true;
        }

        public bool AddColumn()
        {
            ColumnCount++;
            for (int r = 1; r <= RowCount; r++)
            {
                string name = GetColumnName(ColumnCount - 1) + r;
                if (!_cells.ContainsKey(name))
                {
                    _cells[name] = new Cell(name);
                }
            }
            return true;
        }

        public bool DeleteColumn()
        {
            if (ColumnCount <= 1) return false; // Захист: залишаємо мінімум 1 стовпчик

            for (int r = 1; r <= RowCount; r++)
            {
                string name = GetColumnName(ColumnCount - 1) + r;
                _cells.Remove(name);
            }
            ColumnCount--;
            return true;
        }

        public Cell GetCell(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return null!;
            string normalized = NormalizeName(name);
            if (_cells.TryGetValue(normalized, out var cell))
                return cell;
            return null!;
        }

        public static string NormalizeName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return name;
            char first = char.ToUpperInvariant(name[0]);
            first = first switch
            {
                'А' => 'A',
                'В' => 'B',
                'С' => 'C',
                'Е' => 'E',
                'Н' => 'H',
                _ => first
            };
            return first + name.Substring(1).ToUpperInvariant();
        }

        public IEnumerable<Cell> GetAllCells() => _cells.Values;

        public static string GetColumnName(int index)
        {
            int dividend = index + 1;
            string columnName = string.Empty;
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }
            return columnName;
        }
    }
}

