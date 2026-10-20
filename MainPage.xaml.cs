using Microsoft.Maui.Controls;
using System;
using SpreadsheetApp.Models;
using SpreadsheetApp.Parsers;
using SpreadsheetApp.Services;

namespace SpreadsheetApp
{
    public partial class MainPage : ContentPage
    {
        private const int InitialRowCount = 20;
        private const int InitialColumnCount = 7; // A-G

        private Spreadsheet _spreadsheet;
        private GoogleDriveManager _driveManager;
        private Entry[,] _entries;
        private Entry? _selectedEntry;

        public MainPage()
        {
            InitializeComponent();
            _spreadsheet = new Spreadsheet(InitialRowCount, InitialColumnCount);
            _driveManager = new GoogleDriveManager();
            _entries = new Entry[_spreadsheet.RowCount, _spreadsheet.ColumnCount];

            CreateGrid();
        }

        private void CreateGrid()
        {
            SpreadsheetGrid.Children.Clear();
            SpreadsheetGrid.RowDefinitions.Clear();
            SpreadsheetGrid.ColumnDefinitions.Clear();

            int rows = _spreadsheet.RowCount;
            int cols = _spreadsheet.ColumnCount;
            _entries = new Entry[rows, cols];

            // Створення визначень колонок (0 - для номерів рядків, решта - для даних)
            SpreadsheetGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            for (int col = 0; col < cols; col++)
            {
                SpreadsheetGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(100) });
            }

            // Створення визначень рядків (0 - для заголовків колонок, решта - для даних)
            SpreadsheetGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            for (int row = 0; row < rows; row++)
            {
                SpreadsheetGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            }

            // Додавання заголовків колонок (A, B, C...)
            for (int col = 0; col < cols; col++)
            {
                var label = new Label
                {
                    Text = Spreadsheet.GetColumnName(col),
                    HorizontalOptions = LayoutOptions.Center,
                    FontAttributes = FontAttributes.Bold
                };
                Grid.SetRow(label, 0);
                Grid.SetColumn(label, col + 1);
                SpreadsheetGrid.Children.Add(label);
            }

            // Додавання номерів рядків та клітинок
            for (int row = 0; row < rows; row++)
            {
                var label = new Label
                {
                    Text = (row + 1).ToString(),
                    VerticalOptions = LayoutOptions.Center,
                    HorizontalOptions = LayoutOptions.Center,
                    FontAttributes = FontAttributes.Bold,
                    Margin = new Thickness(0, 0, 10, 0)
                };
                Grid.SetRow(label, row + 1);
                Grid.SetColumn(label, 0);
                SpreadsheetGrid.Children.Add(label);

                for (int col = 0; col < cols; col++)
                {
                    string cellName = Spreadsheet.GetColumnName(col) + (row + 1);
                    var cell = _spreadsheet.GetCell(cellName);

                    var entry = new Entry
                    {
                        Text = cell != null ? FormatCellValue(cell.Value) : "",
                        TextColor = (cell != null && cell.Value == null && !string.IsNullOrWhiteSpace(cell.Expression)) ? Colors.Red : Colors.White,
                        Margin = new Thickness(1)
                    };

                    int r = row;
                    int c = col;
                    entry.Focused += (s, e) => Entry_Focused(r, c, entry);
                    entry.Unfocused += (s, e) => Entry_Unfocused(r, c, entry);

                    Grid.SetRow(entry, row + 1);
                    Grid.SetColumn(entry, col + 1);
                    SpreadsheetGrid.Children.Add(entry);
                    _entries[row, col] = entry;
                }
            }
        }

        private void Entry_Focused(int row, int col, Entry entry)
        {
            _selectedEntry = entry;
            string cellName = Spreadsheet.GetColumnName(col) + (row + 1);
            var cell = _spreadsheet.GetCell(cellName);
            FormulaBar.Text = cell?.Expression ?? "";

            // Під час редагування показуємо сирий вираз (формулу)
            entry.Text = cell?.Expression ?? "";
        }

        private void Entry_Unfocused(int row, int col, Entry entry)
        {
            string cellName = Spreadsheet.GetColumnName(col) + (row + 1);
            var cell = _spreadsheet.GetCell(cellName);
            if (cell != null)
            {
                cell.Expression = entry.Text;
            }

            // Автоматичний миттєвий перерахунок усієї таблиці при виході з клітинки
            RecalculateAll();
        }

        private void FormulaBar_Unfocused(object sender, FocusEventArgs e)
        {
            if (_selectedEntry != null)
            {
                _selectedEntry.Text = FormulaBar.Text;
                int row = Grid.GetRow(_selectedEntry) - 1;
                int col = Grid.GetColumn(_selectedEntry) - 1;
                string cellName = Spreadsheet.GetColumnName(col) + (row + 1);
                var cell = _spreadsheet.GetCell(cellName);
                if (cell != null)
                {
                    cell.Expression = FormulaBar.Text;
                }

                // Автоматичний миттєвий перерахунок усієї таблиці
                RecalculateAll();
            }
        }

        private void EvaluateCell(SpreadsheetApp.Models.Cell cell, Entry entry)
        {
            if (string.IsNullOrWhiteSpace(cell.Expression))
            {
                cell.Value = null;
                cell.Error = string.Empty;
                cell.References.Clear();
                entry.Text = "";
                return;
            }

            try
            {
                var parser = new ExpressionParser(_spreadsheet);
                cell.Value = parser.Evaluate(cell.Expression, cell);
                cell.Error = string.Empty;
                entry.Text = FormatCellValue(cell.Value);
                entry.TextColor = Colors.White;
            }
            catch (Exception ex)
            {
                cell.Value = null;
                cell.Error = ex.Message;
                entry.Text = "#ERROR";
                entry.TextColor = Colors.Red;
            }
        }

        private static string FormatCellValue(object? value)
        {
            if (value is double d)
            {
                // Округлення до 2 знаків після коми для охайного показу (напр. 3.33 або 10)
                return Math.Round(d, 2).ToString(System.Globalization.CultureInfo.InvariantCulture);
            }
            return value?.ToString() ?? string.Empty;
        }

        private void RecalculateAll()
        {
            // Скидаємо значення перед новим циклом обчислення
            foreach (var cell in _spreadsheet.GetAllCells())
            {
                cell.Value = null;
                cell.Error = string.Empty;
            }

            // Точкове обчислення клітинок з рекурсивним розв'язанням залежностей
            int rows = _spreadsheet.RowCount;
            int cols = _spreadsheet.ColumnCount;

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    string cellName = Spreadsheet.GetColumnName(c) + (r + 1);
                    var cell = _spreadsheet.GetCell(cellName);
                    if (cell != null && r < _entries.GetLength(0) && c < _entries.GetLength(1))
                    {
                        EvaluateCell(cell, _entries[r, c]);
                    }
                }
            }

            // Оновлення тексту та кольору для всіх клітинок сітки
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    string cellName = Spreadsheet.GetColumnName(c) + (r + 1);
                    var cell = _spreadsheet.GetCell(cellName);
                    if (cell != null && r < _entries.GetLength(0) && c < _entries.GetLength(1))
                    {
                        var entry = _entries[r, c];
                        if (string.IsNullOrWhiteSpace(cell.Expression))
                        {
                            entry.Text = "";
                        }
                        else if (cell.Value != null)
                        {
                            entry.Text = FormatCellValue(cell.Value);
                            entry.TextColor = Colors.White;
                        }
                        else
                        {
                            entry.Text = "#ERROR";
                            entry.TextColor = Colors.Red;
                        }
                    }
                }
            }
        }

        private void AddRowButton_Clicked(object sender, EventArgs e)
        {
            _spreadsheet.AddRow();
            CreateGrid();
            RecalculateAll();
        }

        private async void DeleteRowButton_Clicked(object sender, EventArgs e)
        {
            if (!_spreadsheet.DeleteRow())
            {
                await DisplayAlert("Попередження", "Не можна видалити останній рядок. Таблиця повинна містити хоча б один рядок!", "OK");
                return;
            }
            CreateGrid();
            RecalculateAll();
        }

        private void AddColumnButton_Clicked(object sender, EventArgs e)
        {
            _spreadsheet.AddColumn();
            CreateGrid();
            RecalculateAll();
        }

        private async void DeleteColumnButton_Clicked(object sender, EventArgs e)
        {
            if (!_spreadsheet.DeleteColumn())
            {
                await DisplayAlert("Попередження", "Не можна видалити останній стовпчик. Таблиця повинна містити хоча б один стовпчик!", "OK");
                return;
            }
            CreateGrid();
            RecalculateAll();
        }

        private async void SaveButton_Clicked(object sender, EventArgs e)
        {
            try
            {
                await _driveManager.SaveSpreadsheetAsync(_spreadsheet);
                await DisplayAlert("Успіх", "Таблицю успішно збережено на Google Drive!", "OK");
            }
            catch (Exception ex)
            {
                await DisplayAlert("Помилка збереження", ex.ToString(), "OK");
            }
        }

        private async void LoadButton_Clicked(object sender, EventArgs e)
        {
            try
            {
                await _driveManager.LoadSpreadsheetAsync(_spreadsheet);
                CreateGrid(); // Оновлюємо сітку під завантажені розміри
                RecalculateAll();
                await DisplayAlert("Успіх", "Таблицю завантажено з Google Drive!", "OK");
            }
            catch (Exception ex)
            {
                await DisplayAlert("Помилка завантаження", ex.ToString(), "OK");
            }
        }

        private void HelpButton_Clicked(object sender, EventArgs e)
        {
            DisplayAlert("Довідка: Варіант 37", 
                "Розробник: Черненко Ігор Ігорович\n" +
                "Група: К-28\n" +
                "GitHub: https://github.com/zefin100/OOP_LR1\n\n" +
                "Підтримувані операції (Варіант 37):\n" +
                "• Арифметичні: +, -, *, /\n" +
                "• Функції: max(x, y), min(x, y)\n" +
                "• Порівняння: <=, >=, <>\n" +
                "• Логічні: not\n\n" +
                "Можливості програми:\n" +
                "• Автоматичний перерахунок у реальному часі при виході з клітинки\n" +
                "• Динамічне додавання та видалення рядків і стовпчиків\n" +
                "• Захист від видалення останньої клітинки (мінімум 1 рядок і 1 стовпчик)\n" +
                "• Синхронізація з хмарою Google Drive (OAuth 2.0)", "OK");
        }
    }
}
