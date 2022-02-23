using System;
using System.Linq;
using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using ClosedXML.Report.Options;
using MoreLinq;

namespace ClosedXML.Report.Footer {
    /// <summary>
    /// Тег для шаблона ClosedXML.Reports позволяющий размещать в подвале на каждой странице изображение
    /// Порядок использования:
    ///     1. В шаблоне необходимо создать лист Footer
    ///     2. Добавить в A2 <<Delete>>
    ///     3. Скопировать структуру колонок с исходного листа
    ///     4. Разместить изображение с именем Footer
    ///     5. Выделить область которая должна быть подвалом и задать ей имя Footer
    ///     6. На основном листе в столбце A добавить <<Footer>>
    /// </summary>
    public class FooterOptionTag: OptionTag {
        public override void Execute(ProcessingContext context){
            var xlCell = Cell.GetXlCell(context.Range);
            var cellAddr = xlCell.Address.ToStringRelative(false);
            var ws = Range.Worksheet;

            string imageAlign = HasParameter("Align") ? GetParameter("Align").ToLower() : "center";
            string footerText = HasParameter("Text") ? GetParameter("Text") : "";
            // имя шаблона подвала, по умолчанию Footer
            string footerName = HasParameter("Name") ? GetParameter("Name") : "Footer";
            // последняя строка на странице
            int lastRowOnFirstPage = HasParameter("Row") ? Convert.ToInt32(GetParameter("Row")) : ws.PageSetup.LastRowToRepeatAtTop;
            
            // размещение текста в подвале
            if (!String.IsNullOrWhiteSpace(footerText)) {
                var footerPosition = imageAlign switch {
                    "left" => ws.PageSetup.Footer.Left,
                    "center" => ws.PageSetup.Footer.Center,
                    "right" => ws.PageSetup.Footer.Right,
                    _ => ws.PageSetup.Footer.Center
                };
                footerPosition.AddText(footerText, XLHFOccurrence.AllPages);
            }

            if (ws.Workbook.Worksheets.Contains(footerName)) {
                var footerWorksheet = ws.Workbook.Worksheets.Worksheet(footerName);
                IXLPicture footerPicture;
                if (footerWorksheet.Pictures.Contains(footerName)) {
                    footerPicture = footerWorksheet.Picture(footerName);
                }
                else {
                    footerPicture = footerWorksheet.Pictures.First();
                }
                var footerRange = footerWorksheet.NamedRange(footerName);
                footerWorksheet.Hide();

                for (var page = 1; page < (ws.LastRowUsed().RowNumber() / lastRowOnFirstPage) + 2; page++) {
                    
                    var picture = ws.AddPicture(footerPicture.ImageStream, $"footer_{page}");
                    picture.WithPlacement(XLPicturePlacement.FreeFloating);
                    picture.Width = footerPicture.Width;
                    picture.Height = footerPicture.Height;
                    
                    // Вставить над последней строкой на странице столько же строк как и в шаблоне подвала
                    var lastRowOnPage = lastRowOnFirstPage * page;
                    var startRowOnPage = lastRowOnPage - footerRange.Ranges.First().RowCount();
                    var newFooterRows = ws.Row(startRowOnPage).InsertRowsBelow(footerRange.Ranges.First().RowCount());
                    newFooterRows.ForEach(x => x.Clear());
                    
                    // Разместить изображение с аналогичным смещением как в шаблоне
                    var topRow = newFooterRows.First();
                    picture.MoveTo(topRow.Cell(1), footerPicture.Left, footerPicture.Top);
                    newFooterRows.Last().AddHorizontalPageBreak();
                }
            }
        }
        public static void Register(string tag="Footer") {
            TagsRegister.Add<FooterOptionTag>(tag, 0);
        }
    }
}