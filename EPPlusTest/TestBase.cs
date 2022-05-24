using System;
using System.IO;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest
{
    [TestClass]
    public abstract class TestBase
    {
        protected ExcelPackage _pck;
        protected string _clipartPath = "";
        protected string _worksheetPath = "F:\\Workbooks";
        protected string _testInputPath = ".\\Workbooks\\";
        public TestContext TestContext { get; set; }

        [TestInitialize]
        public void InitBase()
        {
            _clipartPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources");
            if (!Directory.Exists(_clipartPath))
            {
                Directory.CreateDirectory(_clipartPath);
            }

            var di = new DirectoryInfo(_worksheetPath);
            _worksheetPath = di.FullName + "\\";

            _pck = new ExcelPackage();
        }

        protected ExcelPackage OpenPackage(string name, bool delete = false)
        {
            var fi = new FileInfo(_worksheetPath + name);
            if (delete && fi.Exists)
            {
                fi.Delete();
            }
            _pck = new ExcelPackage(fi);
            return _pck;
        }
        protected ExcelPackage OpenTemplatePackage(string name)
        {
            var t = new FileInfo(_testInputPath + name);
            if (t.Exists)
            {
                var fi = new FileInfo(_worksheetPath + name);
                _pck = new ExcelPackage(fi, t);
            }
            else
            {
                Assert.Inconclusive($"Template {name} does not exist in path {_testInputPath}");
            }
            return _pck;
        }

        protected void SaveWorksheet(string name)
        {
            if (_pck.Workbook.Worksheets.Count == 0) return;
            var fi = new FileInfo(_worksheetPath + name);
            if (fi.Exists)
            {
                fi.Delete();
            }
            _pck.SaveAs(fi);
        }
    }
}
