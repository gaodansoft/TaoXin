using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Man
{
    public class Report : IDisposable
    {
        public int CurRow = 0;
        public Regex RegexPara = new Regex(@"\[[^/]+\]");

        public readonly string NULLValue = "NA";
        protected SLDocument slDocument;
        public Report()
        {
            slDocument = new SLDocument(GetSPath());
        }
        public void Create(ReportData report)
        {
            slDocument.SetCellValue(2, 2, report.TemplateName);
            for (int i = 0; i < report.BomData.Count; i++)
            {
                CurRow = i + 4;
                slDocument.CopyRow(4, CurRow);
                AddRow(report.BomData[i]);
            }

        }
        public void AddRow(BomData bomData)
        {
            slDocument.SetCellValue(CurRow, 2, bomData.Level);
            slDocument.SetCellValue(CurRow, 3, bomData.TemplateName);
            slDocument.SetCellValue(CurRow, 4, bomData.SeriesID);
            slDocument.SetCellValue(CurRow, 5, bomData.TemplateID);
            slDocument.SetCellValue(CurRow, 6, bomData.TemplateState);
            slDocument.SetCellValue(CurRow, 7, bomData.Identifier);
        }
        public static string GetSPath()
        {
            var plugInpath = System.IO.Path.GetDirectoryName(typeof(Report).Assembly.CodeBase.Replace("file:///", ""));
            return Path.Combine(plugInpath, "Template.xlsx");
        }
        public void Save(string fileName)
        {
            MemoryStream ms = new MemoryStream();
            slDocument.SaveAs(ms);
            var t = ms.ToArray();
            File.WriteAllBytes(fileName, t);
        }


        public byte[] GetFile()
        {
            MemoryStream ms = new MemoryStream();
            slDocument.SaveAs(ms);
            return ms.ToArray();
        }
        public void Dispose()
        {
            if (slDocument != null)
            {
                slDocument.Dispose();
            }

        }
    }
    public class ReportData
    {
        public string TemplateName { get; set; }
        public List<BomData> BomData = new List<BomData>();
    }
    public class BomData
    {
        public string Level { get; set; }
        public string TemplateName { get; set; }
        public string SeriesID { get; set; }
        public string TemplateID { get; set; }
        public string TemplateState { get; set; }
        public string Identifier { get; set; }
    }
}
