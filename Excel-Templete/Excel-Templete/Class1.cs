using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;
using DocumentFormat.OpenXml.Packaging;

namespace DocumentFormat.OpenXml
{
    public class ExcelTemplate
    {
        public int CreateExcel (DataTable dt, string TemplatePath, string FileName)
        {
            
            //открыть шаблон

            //byte[] byteArray = System.IO.File.ReadAllBytes(Server.MapPath(model.template.path));
            string TempFileName = FileName;
            System.IO.File.Copy(TemplatePath, TempFileName, true);

            SpreadsheetDocument spreadsheetDoc = SpreadsheetDocument.Open(TempFileName, true);
            WorkbookPart wp = spreadsheetDoc.WorkbookPart;
            Sheet sheet = wp.Workbook.Sheets.GetFirstChild<Sheet>();
            WorksheetPart worksheetPart = wp.WorksheetParts.First();
            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
            //поиск столбцов для всех сущностей таблицы

            int rownum = -1;
            int login = -1;
            int date = -1;
            int documentNumber = -1;
            int debtorFio = -1;
            int isCopy = -1;

            foreach (Row r in sheetData.Elements<Row>())
            {
                int indx = 0;
                foreach (Cell c in r.Elements<Cell>())
                {

                    //Получение значения ячейки 
                    int id = -1;
                    string cell = c.ToString();

                    if (Int32.TryParse(c.InnerText, out id))
                    {
                        SharedStringItem item = GetSharedStringItemById(wp, id);
                        if (item.Text != null)
                        {
                            cell = item.Text.Text;
                        }
                        else if (item.InnerText != null)
                        {
                            cell = item.InnerText;
                        }
                        else if (item.InnerXml != null)
                        {
                            cell = item.InnerXml;
                        }
                    }

                    switch (cell)
                    {
                        case "{rownum}":
                            rownum = indx;
                            break;
                        case "{login}":
                            login = indx;
                            break;
                        case "{date}":
                            date = indx;
                            break;
                        case "{documentNumber}":
                            documentNumber = indx;
                            break;
                        case "{debtorFio}":
                            debtorFio = indx;
                            break;
                        case "{isCopy}":
                            isCopy = indx;
                            break;
                    }
                    indx++;

                }

            }
            //копировать строку 2
            Row row2 = sheetData.Elements<Row>().ElementAt(1);
            row2.Remove();
            //удалить строку 2
            //   sheetData.Elements<Row>().ElementAt(1).Remove();
            int max = 1000;
            long rn = 1;
            long thisIndex = sheetData.Elements<Row>().Count() + 1;
            foreach (var req in dt.Rows)
            {
                Row newRow = new Row();
                newRow.RowIndex = (UInt32)thisIndex;
                //заполнить согласно данным мапинга
                if (rownum >= 0)
                {
                    Cell c = textCell(rn.ToString(), rownum, thisIndex);
                    newRow.AppendChild(c);
                    rn++;
                }
                //вставить
                sheetData.Append(newRow);
                thisIndex++;
                if (max-- == 0) break;
            }
            int cnt = sheetData.Elements<Row>().Count<Row>();
            //сохранить список в реестр
            //Вернуть файл
            wp.Workbook.Save();
            spreadsheetDoc.Close();
            return 0;
        }
        public static SharedStringItem GetSharedStringItemById(WorkbookPart workbookPart, int id)
        {
            return workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
        }
        private Cell textCell(string text, int column, long row)
        {
            string[] headerColumns = new string[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
            Cell c = new Cell();
            c.CellReference = headerColumns[column] + row.ToString();
            CellValue v = new CellValue();
            v.Text = text;
            c.AppendChild(v);
            return c;
        }
        private Cell dateCell(DateTime datetext, int column, long row)
        {
            string[] headerColumns = new string[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
            Cell c = new Cell();
            c.StyleIndex = 1;
            c.CellReference = headerColumns[column] + row.ToString();
            CellValue v = new CellValue();
            v.Text = datetext.ToOADate().ToString();
            c.AppendChild(v);
            return c;
        }
    }
}
