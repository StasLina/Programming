using System;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Data;

namespace Task1
{
    internal class SmartController
    {
        class Data
        {
            private  readonly struct Threshold {
                readonly double Number;
                readonly double TaxesRate;
                public Threshold(double number, double taxesRate)
                {
                    Number = number;
                    TaxesRate = taxesRate;
                }
                public static readonly Threshold Threshold1 = new Threshold(20000.0, 0.12);
                public static readonly Threshold Threshold2 = new Threshold(40000.0, 0.20);
                public static readonly Threshold ThresholdFinaly = new Threshold(double.MaxValue, 0.35);
                public static readonly List<Threshold> ThresholdList = new List<Threshold> { Threshold1, Threshold2, ThresholdFinaly };
                public static double GetTaxRate(double amountMoney)
                {
                    var enumer = ThresholdList.GetEnumerator();

                    if (enumer.Current.Number > amountMoney)
                    {
                        while (enumer.MoveNext())
                        {
                            if (enumer.Current.Number < amountMoney)
                            {
                                return enumer.Current.TaxesRate;
                            }
                        }
                    }
                    else
                    {
                        return enumer.Current.TaxesRate;
                    }

                    return 0;
                }
            }

            private struct Record
            {
                string _lastName;
                string _firstName;
                double _annualIncomeNumber;
                double _taxesPaid;

                public Record(string lastName, string firstName, double annualIncomeNumber) {
                    this._lastName = lastName;
                    this._firstName = firstName;
                    this._annualIncomeNumber = annualIncomeNumber;
                    this._taxesPaid = _annualIncomeNumber * Threshold.GetTaxRate(this._annualIncomeNumber);
                }

                public string LastName{
                    get
                    {
                        return _lastName;
                    } 
                }

                public string FirstName
                {
                    get
                    {
                        return _firstName;
                    } 
                }

                public double AnnualIncomeNumber
                {
                    get
                    {
                        return _annualIncomeNumber;
                    } 
                }

                public double TaxesPaid
                {
                    get
                    {
                        return _taxesPaid;
                    } 
                }
            }

            private List<Record> listRecords;

            public void Add(string lastName, string firstName, double annualIncomeNumber)
            {
                listRecords.Add(new Record(lastName, firstName, annualIncomeNumber));
            }

            public System.Data.DataTable GetDataTable()
            {
                System.Data.DataTable table = new System.Data.DataTable("Уплачиваемый доход");

                DataRow row;

                // Определяем столбцы
                DataColumn column;
                column = new DataColumn();
                column.DataType = System.Type.GetType("System.String");
                column.ColumnName = "Фамилия";
                column.ReadOnly = true;
                table.Columns.Add(column);

                column = new DataColumn();
                column.DataType = System.Type.GetType("System.String");
                column.ColumnName = "Имя";
                column.ReadOnly = true;
                table.Columns.Add(column);

                column = new DataColumn();
                column.DataType = System.Type.GetType("System.Double");
                column.ColumnName = "Годовой доход";
                column.ReadOnly = true;
                table.Columns.Add(column);

                column = new DataColumn();
                column.DataType = System.Type.GetType("System.Double");
                column.ColumnName = "Налог";
                column.ReadOnly = true;
                table.Columns.Add(column);

                // Добавляем записи
                var RecordsIterator = listRecords.GetEnumerator();

                do{
                    row = table.NewRow();
                    row["Фамилия"] = RecordsIterator.Current.FirstName;
                    row["Имя"] = RecordsIterator.Current.LastName;
                    row["Годовой доход"] = RecordsIterator.Current.AnnualIncomeNumber;
                    row["Налог"] = RecordsIterator.Current.TaxesPaid;
                } while(RecordsIterator.MoveNext());

                return table;
            }
        }

        Data _model = new Data();
        Exception _lastException;

        public string GetLastException()
        {
            if(_lastException == null)
            {
                return "";
            }
            else
            {
                return _lastException.Message;
            }
        }

        public System.Data.DataTable GetDataTable()
        {
            return _model.GetDataTable();
        }

        // Атрибут необходим для корректной работы OpenFileDialog
        static void Test()
        {
            // Вызываем функцию, которая открывает диалог и читает файл
            string sqlContent = ReadSqlFile();
            if (!string.IsNullOrEmpty(sqlContent))
            {
                Console.WriteLine("Содержимое файла:");
                Console.WriteLine(sqlContent);
            }
        }

        [STAThread]
        static string ReadSqlFile()
        {
            // Создаем экземпляр OpenFileDialog
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "SQL files (*.sql)|*.sql|All files (*.*)|*.*", // Фильтр для отображения SQL файлов
                Title = "Выберите файл SQL" // Заголовок окна выбора файла
            };

            // Открываем диалоговое окно для выбора файла
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    // Читаем содержимое выбранного файла
                    string filePath = openFileDialog.FileName;
                    return File.ReadAllText(filePath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при чтении файла: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return null;
                }
            }

            // Если пользователь не выбрал файл
            return null;
        }

        public int ReadExcelFile()
        {
            int retCode = 0;
            // Открытие диалогового окна выбора файла
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files (*.xls;*.xlsx)|*.xls;*.xlsx",
                Title = "Выберите файл Excel"
                //"C:\\Users\\sthoz\\OneDrive\\Рабочий стол\\test.xlsx";
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;

                // Создание экземпляра приложения Excel
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook workbook = null;
                Microsoft.Office.Interop.Excel.Worksheet worksheet = null;
                
                try
                {
                    // Открытие книги Excel
                    workbook = excelApp.Workbooks.Open(filePath);
                    worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1]; // Чтение первого листа

                    // Чтение данных из ячейки A1
                    //MessageBox.Show($"Значение в ячейке A1: {cellValue}");

                    // Обработка данных (например, чтение нескольких ячеек или строк)
                    // Фамилия, Имя, годовой дох
                    while (true)
                    {
                        string lastName = worksheet.Cells[2, 1].Value.ToString();
                        string firstName = worksheet.Cells[2, 1].Value.ToString();
                        string annualIncome = worksheet.Cells[2, 1].Value.ToString();

                        double annualIncomeNumber = double.Parse(annualIncome);
                        _model.Add(lastName, firstName, annualIncomeNumber);

                        if (lastName == "")
                        {
                            break;
                        }
                    }

                    // Пример чтения данных из диапазона A1:B2
                    Microsoft.Office.Interop.Excel.Range range = worksheet.get_Range("A1", "B2");
                    object[,] values = (object[,])range.Value2;

                    for (int i = 1; i <= values.GetLength(0); i++)
                    {
                        for (int j = 1; j <= values.GetLength(1); j++)
                        {

                            Debug.Write(values[i, j] + "\t");
                        }
                        Debug.WriteLine("");
                    }
                }
                catch (Exception excpetion)
                {
                    _lastException = excpetion;
                    retCode = 1;
                }
                finally
                {
                    // Закрытие книги и приложения Excel
                    if (workbook != null)
                    {
                        workbook.Close(false);
                    }
                    excelApp.Quit();

                    // Освобождение ресурсов
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                }
            }
            return retCode;
        }
    }
}