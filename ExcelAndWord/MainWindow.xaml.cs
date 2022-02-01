using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace ExcelAndWord
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnToExcel_Click(object sender, RoutedEventArgs e)
        {
            var application = new Excel.Application();
            application.SheetsInNewWorkbook = 1;

            Excel.Workbook workbook = application.Workbooks.Add(Type.Missing);

            int startRowIndex = 1;

            Excel.Worksheet worksheet = (Excel.Worksheet)application.Worksheets.Item[1];
            worksheet.Name = "Order";

            worksheet.Cells[1][startRowIndex] = "Number";
            worksheet.Cells[2][startRowIndex] = "Table";
            worksheet.Cells[3][startRowIndex] = "Product";

            startRowIndex++;
            worksheet.Cells[1][startRowIndex] = tbxNum.Text;
            worksheet.Cells[2][startRowIndex] = tbxTable.Text;
            worksheet.Cells[3][startRowIndex] = tbxProduct.Text;

            worksheet.Columns.AutoFit();
            worksheet.Rows.AutoFit();

            application.Visible = true;
        }

        private void btnToWord_Click(object sender, RoutedEventArgs e)
        {
            var application = new Word.Application();

            Word.Document document = application.Documents.Add();

            Word.Paragraph orderParagraph = document.Paragraphs.Add();
            Word.Range orderRange = orderParagraph.Range;
            orderRange.Text = "Order";
            orderRange.InsertParagraphAfter();

            Word.Paragraph tableParagraph = document.Paragraphs.Add();
            Word.Range tableRange = tableParagraph.Range;
            Word.Table orderTable = document.Tables.Add(tableRange, 2, 3);
            orderTable.Borders.InsideLineStyle = orderTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

            Word.Range cellRange;

            cellRange = orderTable.Cell(1, 1).Range;
            cellRange.Text = "Number";
            cellRange = orderTable.Cell(1, 2).Range;
            cellRange.Text = "Table";
            cellRange = orderTable.Cell(1, 3).Range;
            cellRange.Text = "Product";
            cellRange = orderTable.Cell(2, 1).Range;
            cellRange.Text = tbxNum.Text;
            cellRange = orderTable.Cell(2, 2).Range;
            cellRange.Text = tbxTable.Text;
            cellRange = orderTable.Cell(2, 3).Range;
            cellRange.Text = tbxProduct.Text;

            application.Visible = true;
        }
    }
}
