using CsvHelper;
using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;

namespace ExtractCouncilTaxBase
{
    class Program
    {
        static void Main(string[] args)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            List<ExtractionInstruction> ExtractionInstructions = new List<ExtractionInstruction>();
            ExtractionInstruction instruction2019 = new ExtractionInstruction()
            {
                SpreadsheetPath = @"Assets/Local_Authorities_Council_Taxbase_2019_Drop_down.xlsx",
                Year = 2019,
                SheetName = "Data",
                OldIDColumn = 0,
                LANameColumn = 3,
                ModernIDColumn = 1,
                SecondHomesColumn = 107,
                MinRow = 6,
                MaxRow = 322
            };
            ExtractionInstructions.Add(instruction2019);
            ExtractionInstruction instruction2018 = new ExtractionInstruction()
            {
                SpreadsheetPath = @"Assets/LA_Drop_down_2018_rev.xlsx",
                Year = 2018,
                SheetName = "Data",
                OldIDColumn = 0,
                LANameColumn = 3,
                ModernIDColumn = -1,
                SecondHomesColumn = 107,
                MinRow = 6,
                MaxRow = 331
            };
            ExtractionInstructions.Add(instruction2018);
            ExtractionInstruction instruction2017 = new ExtractionInstruction()
            {
                SpreadsheetPath = @"Assets/LA_Drop_down_2017_web__rev_.xlsx",
                Year = 2017,
                SheetName = "Data",
                OldIDColumn = 0,
                LANameColumn = 3,
                ModernIDColumn = -1,
                SecondHomesColumn = 107,
                MinRow = 6,
                MaxRow = 331
            };
            ExtractionInstructions.Add(instruction2017);
            ExtractionInstruction instruction2016 = new ExtractionInstruction()
            {
                SpreadsheetPath = @"Assets/LA_Drop_down_2016_revised_Jan.xlsx",
                Year = 2016,
                SheetName = "Data",
                OldIDColumn = 0,
                LANameColumn = 3,
                ModernIDColumn = -1,
                SecondHomesColumn = 107,
                MinRow = 6,
                MaxRow = 331
            };
            ExtractionInstructions.Add(instruction2016);
            ExtractionInstruction instruction2015 = new ExtractionInstruction()
            {
                SpreadsheetPath = @"Assets/CTB_Form_October_2015-_drop_down_-_revised_Feb_2016.xlsx",
                Year = 2015,
                SheetName = "Data",
                OldIDColumn = 0,
                LANameColumn = 3,
                ModernIDColumn = -1,
                SecondHomesColumn = 107,
                MinRow = 6,
                MaxRow = 331
            };
            ExtractionInstructions.Add(instruction2015);
            ExtractionInstruction instruction2014 = new ExtractionInstruction()
            {
                SpreadsheetPath = @"Assets/Revised_CTB_Form_October_2014-_drop_down_sv.xlsx",
                Year = 2014,
                SheetName = "Data",
                OldIDColumn = 0,
                LANameColumn = 3,
                ModernIDColumn = -1,
                SecondHomesColumn = 107,
                MinRow = 6,
                MaxRow = 331
            };
            ExtractionInstructions.Add(instruction2014);
            ExtractionInstruction instruction2013 = new ExtractionInstruction()
            {
                SpreadsheetPath = @"Assets/Council_Taxbase_local_authority_level_data_2013.xlsx",
                Year = 2013,
                SheetName = "Data",
                OldIDColumn = 0,
                LANameColumn = 3,
                ModernIDColumn = -1,
                SecondHomesColumn = 107,
                MinRow = 6,
                MaxRow = 331
            };
            ExtractionInstructions.Add(instruction2013);
            ExtractionInstruction instruction2012 = new ExtractionInstruction()
            {
                SpreadsheetPath = @"Assets/Revised_2012_Local_Authority_level_data.xls",
                Year = 2012,
                SheetName = "Data",
                OldIDColumn = 0,
                LANameColumn = 3,
                ModernIDColumn = -1,
                SecondHomesColumn = 107,
                MinRow = 6,
                MaxRow = 331
            };
            ExtractionInstructions.Add(instruction2012);
            ExtractionInstruction instruction2011 = new ExtractionInstruction()
            {
                SpreadsheetPath = @"Assets/2011_Local_Authority_level_data.xls",
                Year = 2011,
                SheetName = "Data",
                OldIDColumn = 0,
                LANameColumn = 3,
                ModernIDColumn = -1,
                SecondHomesColumn = 107,
                MinRow = 6,
                MaxRow = 331
            };
            ExtractionInstructions.Add(instruction2011);
            ExtractionInstruction instruction2010 = new ExtractionInstruction()
            {
                SpreadsheetPath = @"Assets/2010_Local_Authority_level_data.xls",
                Year = 2010,
                SheetName = "Data",
                OldIDColumn = 0,
                LANameColumn = 3,
                ModernIDColumn = -1,
                SecondHomesColumn = 107,
                MinRow = 6,
                MaxRow = 331
            };
            ExtractionInstructions.Add(instruction2010);

            List<CouncilTaxBaseLine> CTBaseLines = new List<CouncilTaxBaseLine>();
            foreach (ExtractionInstruction extractioninstruction in ExtractionInstructions)
            {
                string ExcelFilePath = extractioninstruction.SpreadsheetPath;
                DataTable CURRENTTABLE;
                using (var stream = File.Open(ExcelFilePath, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = reader.AsDataSet();
                        DataTableCollection Worksheets = result.Tables;
                        CURRENTTABLE = Worksheets[Worksheets.IndexOf(extractioninstruction.SheetName)];
                    }
                }
                for (int i = extractioninstruction.MinRow - 1; i < extractioninstruction.MaxRow - 1; i++)
                {
                    CouncilTaxBaseLine councilTaxBaseLine = new CouncilTaxBaseLine();
                    councilTaxBaseLine.LA_OldID = CURRENTTABLE.Rows[i].ItemArray[extractioninstruction.OldIDColumn] as string;
                    if (extractioninstruction.ModernIDColumn != -1)
                    {
                        councilTaxBaseLine.LA_NewID = CURRENTTABLE.Rows[i].ItemArray[extractioninstruction.ModernIDColumn] as string;
                    }
                    councilTaxBaseLine.LA_name = CURRENTTABLE.Rows[i].ItemArray[extractioninstruction.LANameColumn] as string;
                    councilTaxBaseLine.SecondHomes = CURRENTTABLE.Rows[i].ItemArray[extractioninstruction.SecondHomesColumn] as double?;
                    councilTaxBaseLine.Year = extractioninstruction.Year;

                    CTBaseLines.Add(councilTaxBaseLine);
                }
            }


            using (var writer = new StreamWriter("SecondHomesCombined.csv"))
            {
                using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
                {
                    csv.WriteRecords(CTBaseLines);
                }
            }
        }
    }

    public class CouncilTaxBaseLine
    {
        public string LA_OldID { get; set; }
        public string LA_NewID { get; set; }
        public string LA_name { get; set; }
        public double? SecondHomes { get; set; }
        public int Year { get; set; }
    }

    public class ExtractionInstruction
    {
        public int Year { get; set; }
        public int OldIDColumn { get; set; }
        public int ModernIDColumn { get; set; }
        public int LANameColumn { get; set; }
        public int SecondHomesColumn { get; set; }
        public string SpreadsheetPath { get; set; }
        public string SheetName { get; set; }
        public int MinRow { get; set; }
        public int MaxRow { get; set; }
    }
}
