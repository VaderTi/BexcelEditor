using System;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using BexcelEditor.Definitions;

namespace BexcelEditor.Forms
{
    public partial class MainWindow
    {
        private bool _bexcelLoaded;
        private static Bexcel _bexcel;
        public MainWindow()
        {
            _bexcelLoaded = false;
            InitializeComponent();
        }

        private void OpenBexcelFile_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dlg = new Microsoft.Win32.OpenFileDialog
                {
                    DefaultExt = ".bexcel",
                    Filter = "Bexcel File (*.bexcel)|*.bexcel"
                };

                if (dlg.ShowDialog() != true) return;
                _bexcel = BexcelConverter.Read(dlg.FileName);

                _bexcelLoaded = true;
                SearchBox.IsEnabled = true;
                UpdateSheets();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }

        private void SaveBexcelFile_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_bexcel == null)
                    return;

                var dlg = new Microsoft.Win32.SaveFileDialog
                {
                    DefaultExt = ".bexcel",
                    Filter = "Bexcel file (*.bexcel)|*.bexcel"
                };

                if (dlg.ShowDialog() == true)
                {
                    BexcelConverter.Save(_bexcel, dlg.FileName);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }

        private void UpdateSheets(string search = "")
        {
            if(!_bexcelLoaded) return;
            Sheets.ItemsSource = (string.IsNullOrEmpty(search) ? _bexcel.Sheets.OrderBy(x => x.Name) : _bexcel.Sheets.Where(x => x.Name.IndexOf(search, StringComparison.InvariantCultureIgnoreCase) >= 0).OrderBy(x => x.Name));
        }

        private void SheetChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!(Sheets.SelectedItem is Bexcel.Sheet sheet))
                return;

            Sheet.ItemsSource = sheet.Dt.DefaultView;

            sheet.Dt.RowChanged += Dt_RowChanged;
            sheet.Dt.RowDeleting += Dt_RowDeleting;
        }

        private static void Dt_RowChanged(object sender, DataRowChangeEventArgs e)
        { // Change & Add
            try
            {
                var sheet = _bexcel.Sheets.First(x => x.Name == ((dynamic)sender).TableName);

                Debug.WriteLine(e.Action);

                if (e.Action == DataRowAction.Add)
                {
                    var bexcelRow = new Bexcel.Row
                    {
                        Index1 = -1,
                        Index2 = 1
                    };

                    for (var i = 0; i < e.Row.ItemArray.Length; i++)
                    {
                        bexcelRow.Cells.Add(new Bexcel.Cell { Index = i, Name = Convert.ToString(e.Row.ItemArray[i]) });
                    }

                    sheet.Rows.Add(bexcelRow);
                    Debug.WriteLine($"New row added to {((dynamic)sender).TableName} table");
                }
                else
                {
                    var rowIndex = sheet.Dt.Rows.IndexOf(e.Row);

                    for (var i = 0; i < e.Row.ItemArray.Length; i++)
                    {
                        if (sheet.Rows[rowIndex].Cells.Find(x => x.Index == i) != null)
                        {
                            sheet.Rows[rowIndex].Cells[i].Name = Convert.ToString(e.Row.ItemArray[i]);
                        }
                        else
                        {
                            sheet.Rows[rowIndex].Cells.Add(new Bexcel.Cell
                            {
                                Index = i,
                                Name = Convert.ToString(e.Row.ItemArray[i])
                            });
                        }
                    }

                    Debug.WriteLine($"Row updated on {((dynamic)sender).TableName} table");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }

        private void MenuItem_SaveSheetAsSQL(object sender, RoutedEventArgs e)
        {
            if (!(sender is MenuItem mi)) return;

            var sheet = (Bexcel.Sheet)mi.DataContext;

            var file = new System.IO.StreamWriter($"{sheet.Name}.sql");
            file.Write($"DELETE FROM \"{sheet.Name}\";\n");
            foreach (var r in sheet.Rows)
            {
                var line = $"INSERT INTO \"{sheet.Name}\"(";

                line = sheet.Columns.Select(t => t.Name).Aggregate(line, (current, s) => current + $"\"{s}\", ");

                line = new StringBuilder(line.ToString()) {[line.Length-2] = Convert.ToChar(")")}.ToString();

                line += "VALUES (";

                line = r.Cells.Aggregate(line, (current, c) => current + $"'{c.Name}', ");

                line = new StringBuilder(line.ToString()) {[line.Length-2] = Convert.ToChar(")")}.ToString();
                line = new StringBuilder(line.ToString()) {[line.Length-1] = Convert.ToChar(";")}.ToString();
                line += "\n";
                Debug.WriteLine(line);
                file.Write(line);
            }
            file.Close();
        }

        private void MenuItem_DeleteSheet(object sender, RoutedEventArgs e)
        {
            if (!(sender is MenuItem mi)) return;

            var sheet = (Bexcel.Sheet)mi.DataContext;

            if(MessageBox.Show("Delete sheet "+sheet.Name+"?", "Bexcel Editor - Delete?", MessageBoxButton.YesNo) == MessageBoxResult.No) return;

            _bexcel.Sheets.Remove(_bexcel.Sheets.First(x => x.Name == sheet.Name));

            UpdateSheets();

            MessageBox.Show("Sheet "+sheet.Name+ " deleted!", "Bexcel Editor - Delete");
        }

        private static void Dt_RowDeleting(object sender, DataRowChangeEventArgs e)
        {
            var sheet = _bexcel.Sheets.First(x => x.Name == ((dynamic)sender).TableName);
            var rowIndex = sheet.Dt.Rows.IndexOf(e.Row);

            sheet.Rows.RemoveAt(rowIndex);

            Debug.WriteLine($"Row deleted on {((dynamic)sender).TableName} table");
        }

        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            UpdateSheets(SearchBox.Text);
        }
    }
}
