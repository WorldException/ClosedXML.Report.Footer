using System.Collections.Generic;
using System.IO;
using ClosedXML.Report;
using ClosedXML.Report.Footer;
using Xunit;

namespace Test {
    public class UnitTestFooter {

        public ReportsModel GenData(int lines, int reports_count=1) {
            var reports = new ReportsModel() {
                Reports = new List<Report>()
            };
            for (var report_i = 1; report_i <= reports_count; report_i++) {
                var report = new Report() {
                    Title = $"Report N{report_i}",
                    Rows = new List<Row>()
                };
                reports.Reports.Add(report);
                for (var n = 1; n < lines; n++) {
                    report.Rows.Add(new Row() {
                        Name = $"Row {n}",
                        Period = "aaaa",
                        Cost = n + 10
                    });
                }
            }

            return reports;
        }
        
        [Fact]
        public void Test1() {
            FooterOptionTag.Register();
            var data = GenData(100, 1); 
            using var tmpl = new XLTemplate("template_1.xlsx");
            tmpl.AddVariable(data);
            var result = tmpl.Generate();
            tmpl.SaveAs("test1.xlsx");
        }
    }
}