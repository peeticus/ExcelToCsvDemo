using System.Globalization;
using System.IO;
using System.Text;
using System.Windows;

using CsvHelper;

using ExcelDataReader;

namespace ExcelToCsvDemo
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private const string XlsFilePath = "Input 2003.xls";
        private const string XlsxFilePath = "Input 2007.xlsx";

        public MainWindow()
        {
            InitializeComponent();

            // Allows reading excel files with .Net Core. <see cref="https://github.com/ExcelDataReader/ExcelDataReader#important-note-on-net-core" />
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        /// <summary>
        /// Converts an exel file to a csv file. The excel source file must represent
        /// a collection of <see cref="Person"/> objects, and the resulting csv file will represent this
        /// same collection of objects.
        /// </summary>
        /// <param name="excelFileName">The name of the excel file to convert.</param>
        private static void ConvertExcelFileToCsv(string excelFileName)
        {
            var csvFileName = excelFileName.Replace(".xlsx", ".csv").Replace(".xls", ".csv").Replace("Input", "Output");
            using var stream = File.Open(excelFileName, FileMode.Open, FileAccess.Read);
            using var reader = ExcelReaderFactory.CreateReader(stream);
            do
            {
                var isHeaderRow = true;

                using var writer = new StreamWriter(csvFileName);
                using var csv = new CsvWriter(writer, CultureInfo.InvariantCulture);
                csv.WriteHeader<Person>();
                csv.NextRecord();

                while (reader.Read())
                {
                    if (!isHeaderRow)
                    {
                        var person = new Person
                        {
                            Id = (int)reader.GetDouble(0),
                            Name = reader.GetString(1),
                            Occupation = reader.GetString(2),
                            DateEntered = reader.GetDateTime(3),
                        };

                        csv.WriteRecord(person);
                        csv.NextRecord();
                    }
                    isHeaderRow = false;
                }

            } while (reader.NextResult());
        }

        private void ConvertButton_Click(object sender, RoutedEventArgs e)
        {
            ConvertExcelFileToCsv(XlsFilePath);
            ConvertExcelFileToCsv(XlsxFilePath);
            MessageBox.Show("Successfully converted xlsx and xls files to csv.");
        }
    }
}
