using System.Collections.Generic;

namespace Test {
    public class Row {
        public string Name { get; set; }
        public string Period { get; set; }
        public double Cost { get; set; }
    }
    
    public class Report {
        public string Title { get; set; }
        public List<Row> Rows { get; set; }
    }

    public class ReportsModel {
        public List<Report> Reports { get; set; }
    }
}