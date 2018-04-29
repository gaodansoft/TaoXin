﻿using SpreadsheetLight;
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
           // slDocument.SetCellValue(2, 2, report.TemplateName);
            for (int i = 0; i < report.BomData.Count; i++)
            {
              
               // slDocument.CopyRow(CurRow+1, CurRow + 2);
                CurRow = i + 2;
                slDocument.InsertRow(CurRow, 1);
                slDocument.CopyRow(CurRow+1, CurRow);
                AddRow(report.BomData[i]);
            }

        }
        public void AddRow(BomData bomData)
        {
            int start = bomData.Level * 3;
            slDocument.SetCellValue(CurRow, start+1, bomData.Bom.Name);
            slDocument.SetCellValue(CurRow, start + 2, bomData.Bom.ID );
            slDocument.SetCellValue(CurRow, start + 3, bomData.Bom.Node);


        }
        public static string GetSPath()
        {
            var plugInpath = System.IO.Path.GetDirectoryName(typeof(Report).Assembly.CodeBase.Replace("file:///", ""));
            return Path.Combine(plugInpath, "家系图表模板.xlsx");
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
        public int  Level { get; set; }
        public BOM Bom { get; set; }
    }
}
