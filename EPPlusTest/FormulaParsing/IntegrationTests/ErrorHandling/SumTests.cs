using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace EPPlusTest.FormulaParsing.IntegrationTests.ErrorHandling
{
    /// <summary>
    /// Summary description for SumTests
    /// </summary>
    [TestClass]
    public class SumTests : FormulaErrorHandlingTestBase
    {
        [TestInitialize]
        public void ClassInitialize()
        {
            BaseInitialize();
        }

        [TestCleanup]
        public void ClassCleanup()
        {
            BaseCleanup();
        }

        private TestContext testContextInstance;

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }
        [TestMethod]
        public void SingleCell()
        {
            Assert.AreEqual(3d, Worksheet.Cells["B3"].Value);
        }

        [TestMethod]
        public void MultiCell()
        {
            Assert.AreEqual(40d, Worksheet.Cells["C10"].Value);
        }

        [TestMethod]
        public void Name()
        {
            Assert.AreEqual(10d, Worksheet.Cells["E10"].Value);
        }

        [TestMethod]
        public void ReferenceError()
        {
            Assert.AreEqual("#REF!", Worksheet.Cells["H10"].Value.ToString());
        }

        [TestMethod]
        public void NameOnOtherSheet()
        {
            Assert.AreEqual(130d, Worksheet.Cells["I10"].Value);
        }

        [TestMethod]
        public void ArrayInclText()
        {
            Assert.AreEqual(7d, Worksheet.Cells["J10"].Value);
        }

        [TestMethod]
        public void NameError()
        {
            Assert.AreEqual("#NAME?", Worksheet.Cells["A41"].Value.ToString());
        }

        [TestMethod]
        public void DivByZeroError()
        {
            Assert.AreEqual("#DIV/0!", Worksheet.Cells["C4"].Value.ToString());
        }

        [TestMethod]
        public void NAError()
        {
            Assert.AreEqual("#N/A", Worksheet.Cells["E30"].Value.ToString());
        }

        [TestMethod, Ignore]
        public void NulError()
        {
            // TODO : Bug during debug returns #NULL! and during tests returns #NAME!
            Assert.AreEqual("#NULL!", Worksheet.Cells["G39"].Value.ToString());
        }

        [TestMethod]
        public void NumberError()
        {
            Assert.AreEqual("#NUM!", Worksheet.Cells["F39"].Value.ToString());
        }
    }
}
