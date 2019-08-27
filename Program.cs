using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.WindowsAPICodePack.Dialogs;
using System.IO;
using System.Collections;
using NPOI.XSSF;
using NPOI.XSSF.UserModel;
using NPOI;
using NPOI.SS.UserModel;
using NPOI.HSSF;
using NPOI.HSSF.UserModel;
using System.Xml.Linq;
using System.Xml.XPath;
using System.Windows.Forms;

namespace XMLtoExcel
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            string filePath = string.Empty;
            if (args.Length == 1)
            {
                filePath = args[0];
            }
            
            else
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("Select language folder. example : Languages/Korean");
                CommonOpenFileDialog dialog1 = new CommonOpenFileDialog();
                dialog1.IsFolderPicker = true;
                dialog1.ShowDialog();
                filePath = dialog1.FileName;
            }

            DirectoryInfo dirs = new DirectoryInfo(filePath); // 루트 폴더 경로 가지고있음
            IEnumerable files = dirs.GetFiles("*.xml", SearchOption.AllDirectories);

            var version = "xlsx";
            IWorkbook workbook = CreateWorkbook(version);
            var sheet = workbook.CreateSheet("Exported");
            int i = 1;
            GetCell(sheet, 0, 0).SetCellValue("상위폴더");
            GetCell(sheet, 0, 1).SetCellValue("변수");
            GetCell(sheet, 0, 2).SetCellValue("내용물");

            foreach(FileInfo XMLFile in files)
            {
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine(string.Format("XML 파일 {0}을 분석하는 중...", XMLFile.Name));
                string ExcelFilePath = Path.Combine(XMLFile.DirectoryName, XMLFile.FileNameWithoutExtension() + ".xlsx");
                XDocument doc = XDocument.Load(XMLFile.FullName);
                IEnumerator enumer = doc.Descendants().GetEnumerator();
                while(enumer.MoveNext())
                {
                    //0번째 열 = 상위폴더
                    //1번째 열 = 노드이름
                    //2번쨰 열 = InnerText
                    XElement item = enumer.Current as XElement;
                    if (item.Name.LocalName == "LanguageData")
                        continue;
                    GetCell(sheet, i, 0).SetCellValue(XMLFile.DirectoryName.Split('\\').Last());
                    GetCell(sheet, i, 1).SetCellValue(item.Name.LocalName);
                    if(string.IsNullOrEmpty(GetCell(sheet, i, 2).StringCellValue))
                        GetCell(sheet, i, 2).SetCellValue(item.Value);
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.Write(string.Format("  Path : {0} ", item.Name.LocalName));
                    Console.ForegroundColor = ConsoleColor.White;
                    Console.Write("| ");
                    Console.ForegroundColor = ConsoleColor.Cyan;
                    Console.Write(string.Format("Value : {0}", item.Value));
                    Console.Write("\n");
                    i++;
                }

                sheet.AutoSizeColumn(0);
                sheet.AutoSizeColumn(1);
                sheet.AutoSizeColumn(2);
            }

            SaveFileDialog dialog = new SaveFileDialog();
            dialog.FileName = "*.xlsx";
            dialog.Filter = ("(*.xlsx)|*.xlsx");
            dialog.OverwritePrompt = true;
            dialog.Title = "Save file As..";
            dialog.ShowDialog();

            WriteExcel(workbook, dialog.FileName);
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine(@"Made by Madeline. https://github.com/zzzz465");
            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.WriteLine("you may now close this tab. by clicking X on right top on the screen, or press alt+F4");
            while(true)
            {

            }
        }

        public static IWorkbook GetWorkbook(string filename, string version)
        {
            if (!File.Exists(filename))
                File.Create(filename);
            using (var stream = new FileStream(filename, FileMode.Open, FileAccess.Read))
            {
                //표준 xls 버젼
                if ("xls".Equals(version))
                {
                    return new HSSFWorkbook(stream);
                }
                //확장 xlsx 버젼
                else if ("xlsx".Equals(version))
                {
                    return new XSSFWorkbook(stream);
                }
                throw new NotSupportedException();
            }
        }

        public static IWorkbook CreateWorkbook(string version)
        {
            //표준 xls 버젼
            if ("xls".Equals(version))
            {
                return new HSSFWorkbook();
            }
            //확장 xlsx 버젼
            else if ("xlsx".Equals(version))
            {
                return new XSSFWorkbook();
            }
            throw new NotSupportedException();
        }

        public static IRow GetRow(ISheet sheet, int rownum)
        {
            var row = sheet.GetRow(rownum);
            if (row == null)
            {
                row = sheet.CreateRow(rownum);
            }
            return row;
        }
        // Row로 부터 Cell를 취득, 생성하기
        public static ICell GetCell(IRow row, int cellnum)
        {
            var cell = row.GetCell(cellnum);
            if (cell == null)
            {
                cell = row.CreateCell(cellnum);
            }
            return cell;
        }
        public static ICell GetCell(ISheet sheet, int rownum, int cellnum)
        {
            var row = GetRow(sheet, rownum);
            return GetCell(row, cellnum);
        }
        public static void WriteExcel(IWorkbook workbook, string filepath)
        {
            using (var file = new FileStream(filepath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(file);
            }
        }
    }
}
