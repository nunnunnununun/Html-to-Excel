using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using AngleSharp.Html.Parser;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;

namespace WpfApp1
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        public IWebDriver webDriver;

        public MainWindow() {
            InitializeComponent();
        }

        private async void Html_parse() {
            webDriver = new ChromeDriver();
            webDriver.Url = "https://www2.jica.go.jp/ja/announce/index.php?contract=1&p=1";

            // 指定したサイトのHTMLをストリームで取得する
            var doc = default(IHtmlDocument);
            using (var client = new HttpClient())
            using (var stream = await client.GetStreamAsync(webDriver.Url.ToString())) {
                // AngleSharp.Parser.Html.HtmlParserオブジェクトにHTMLをパースさせる
                var parser = new HtmlParser();
                doc = await parser.ParseDocumentAsync(stream);
            }
            
            // tagを指定してテーブル部分を取得する
            var items = doc.QuerySelectorAll("#mainColumn > div.border > table > tbody > tr")
                .Skip(2)
                .Select(item => {
            // td単位で複数のデータを取得する
             var row = item.GetElementsByTagName("td");
                    
                    var subject1 = row.ElementAt(0).TextContent;
                    var subject2 = row.ElementAt(1).TextContent;
                    var subject3 = row.ElementAt(2).TextContent;
                    var subject4 = row.ElementAt(3).TextContent;
                    var subject5 = row.ElementAt(4).TextContent;
                    var subject6 = row.ElementAt(5).TextContent;
                    var subject7 = row.ElementAt(6).TextContent;
                    var subject8 = row.ElementAt(7).TextContent;
                    return new {Subject1 = subject1, Subject2 = subject2, Subject3 = subject3, Subject4 = subject4, Subject5 = subject5, Subject6 = subject6, Subject7 = subject7, Subject8 = subject8 };
        });


            //Excel出力
            string filePath = @"C:\test\sample.xlsx";

            //ブック作成
            var book = CreateNewBook(filePath);

            //シート無しのexcelファイルは保存は出来るが、開くとエラーが発生する
            book.CreateSheet("newSheet");

            //シート名からシート取得
            var sheet = book.GetSheet("newSheet");

            // 取得した情報を出力する
            int num = 0;
            items.ToList().ForEach(item => {
                // 行を作成する
                IRow rowObj = sheet.CreateRow(num);
                // セルを作成する
                ICell cellObj1 = rowObj.CreateCell(0);
                ICell cellObj2 = rowObj.CreateCell(1);
                ICell cellObj3 = rowObj.CreateCell(2);
                ICell cellObj4 = rowObj.CreateCell(3);
                ICell cellObj5 = rowObj.CreateCell(4);
                ICell cellObj6 = rowObj.CreateCell(5);
                ICell cellObj7 = rowObj.CreateCell(6);
                ICell cellObj8 = rowObj.CreateCell(7);
                // 作成したセルに値を設定する
                cellObj1.SetCellValue(item.Subject1);
                cellObj2.SetCellValue(item.Subject2);
                cellObj3.SetCellValue(item.Subject3);
                cellObj4.SetCellValue(item.Subject4);
                cellObj5.SetCellValue(item.Subject5);
                cellObj6.SetCellValue(item.Subject6);
                cellObj7.SetCellValue(item.Subject7);
                cellObj8.SetCellValue(item.Subject8);
                num += 1;
            });

            //ブックを保存
            using (var fs = new FileStream(filePath, FileMode.Create)) {
                book.Write(fs);
            }
            webDriver.Quit();
        }

        //ブック作成
        static IWorkbook CreateNewBook(string filePath) {
            IWorkbook book;
            var extension = Path.GetExtension(filePath);

            // HSSF => Microsoft Excel(xls形式)(excel 97-2003)
            // XSSF => Office Open XML Workbook形式(xlsx形式)(excel 2007以降)
            if (extension == ".xls") {
                book = new HSSFWorkbook();
            } else if (extension == ".xlsx") {
                book = new XSSFWorkbook();
            } else {
                throw new ApplicationException("CreateNewBook: invalid extension");
            }

            return book;
        }

        private void Button_Click(object sender, RoutedEventArgs e) {
            Html_parse();
        }
    }
}
