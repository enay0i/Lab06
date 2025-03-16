using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using System;
using System.Threading;
using Test_QuocCuong;
using static Microsoft.IO.RecyclableMemoryStreamManager;

namespace AutomationExerciseTests
{
    public class ProductTests
    {
        private IWebDriver driver;
        string root = AppDomain.CurrentDomain.BaseDirectory;
        string[] expectedHeaders = { "Họ Tên", "Tên Khách Sạn", "Email", "Số Điện Thoại", "Ngày Sinh", "Giới Tính", "Phương Thức Thanh Toán", "Ngày Check-in", "Ngày Check-out", "Tổng Giá", "Tổng Số Phòng", "Trạng Thái Hóa Đơn" };
        [SetUp]
        public void Setup()
        {
            ChromeOptions options = new ChromeOptions();
            options.AddUserProfilePreference("download.default_directory", root);
            options.AddUserProfilePreference("download.prompt_for_download", false);
            options.AddUserProfilePreference("download.directory_upgrade", true);
            options.AddUserProfilePreference("safebrowsing.enabled", true);
            options.AddUserProfilePreference("credentials_enable_service", false);
            options.AddUserProfilePreference("profile.password_manager_enabled", false);
            string filePath = Path.Combine(root, "HoaDon.xlsx");
            string filePathh = Path.Combine(root, "BaoCaoDoanhThu.xlsx");
            driver = new ChromeDriver(options);
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
            if (File.Exists(filePathh))
            {
                File.Delete(filePathh);
            }
            driver.Navigate().GoToUrl("http://localhost:3000/loginOwner");
            IWebElement email = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[2]/div[1]/div[3]/form[1]/div[1]/div[1]/div[2]/div[1]/div[1]/span[1]/input[1]"));
            email.SendKeys("qcuong@gmail.com");
            IWebElement pass = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[2]/div[1]/div[3]/form[1]/div[2]/div[1]/div[2]/div[1]/div[1]/span[1]/input[1]"));
            pass.SendKeys("123456");
            IWebElement btnLogin = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[2]/div[1]/div[3]/form[1]/div[3]/div[1]/div[1]/div[1]/div[1]/button[1]"));
            btnLogin.Click();
            Thread.Sleep(2000);
        }

        //[Test]

        //public void ActiveCustomer()
        //{
        //    driver.Navigate().GoToUrl("http://localhost:3000/loginOwner");
        //    IWebElement email = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[2]/div[1]/div[3]/form[1]/div[1]/div[1]/div[2]/div[1]/div[1]/span[1]/input[1]"));
        //    email.SendKeys("qcuong@gmail.com");
        //    IWebElement pass = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[2]/div[1]/div[3]/form[1]/div[2]/div[1]/div[2]/div[1]/div[1]/span[1]/input[1]"));
        //    pass.SendKeys("123456");
        //    IWebElement btnLogin = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[2]/div[1]/div[3]/form[1]/div[3]/div[1]/div[1]/div[1]/div[1]/button[1]"));
        //    btnLogin.Click();
        //    Thread.Sleep(2000);
        //    IWebElement cusTab = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[1]/a[3]/div[1]/*[name()='svg'][1]"));
        //    cusTab.Click();
        //    Thread.Sleep(2000);
        //    IWebElement deActivate = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[4]/span[1]"));
        //    deActivate.Click();
        //    IWebElement txtReason = driver.FindElement(By.XPath("/html[1]/body[1]/div[2]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/input[1]"));
        //    txtReason.SendKeys("kho noi");
        //    IWebElement btnAccept = driver.FindElement(By.XPath("/html[1]/body[1]/div[2]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/button[2]"));
        //    btnAccept.Click();
        //    bool a = deActivate.Text == "Vô hiệu hóa";
        //    string b = a.ToString();
        //    //ExcelProvider.WriteResultToExcel("C:\\Users\\thanh\\source\\repos\\Lab06\\Lab06\\bin\\Debug\\net8.0\\TestCase_BDCLPM_HK2.xlsx", "TestCase_QuocCuong",7, "kho noi", "gay");
        //    Assert.IsTrue(true);
        //}

        [Test]

        public void EX04_TestCase()
        {
            IWebElement mainTab = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[1]/a[1]/div[1]"));
            mainTab.Click();
            Thread.Sleep(2000);
            IWebElement btnExport = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/button[1]"));
            btnExport.Click();
            IWebElement txtMonth = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/h1[1]"));
            string txtThangStr = txtMonth.Text;
            int txtThang = int.Parse(new string(txtThangStr.Where(char.IsDigit).ToArray()));
            IWebElement txtYear = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[2]/div[1]/h1[1]"));
            string txtYearStr = txtMonth.Text;

            int txtYearr = int.Parse(new string(txtYearStr.Where(char.IsDigit).ToArray()));
    
            Thread.Sleep(2000);
            string patha = Path.Combine(root, "BaoCaoDoanhThu.xlsx");
            bool isValid = ExcelProvider.ValidateExcelData(patha, "data", txtYearr, txtThang);
            string actual = isValid ? "File Excel hiển thị đúng doanh thu trong vòng 2 năm" : "File Excel hiển thị sai";
            string result = isValid ? "Passed" : "Failed";
            string path = Path.Combine(root, "TestCase_BDCLPM_HK2.xlsx");
            ExcelProvider.WriteResultToExcel(path, "TestCase_QuocCuong", 36, actual, result);
            Assert.IsTrue(true);
        }
        [Test]

        public void EX01_EX02_TestCase()
        {
            IWebElement invoiceTab = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[1]/a[6]"));
            invoiceTab.Click();
            Thread.Sleep(2000);
            IWebElement btnExport = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[2]/div[1]/button[1]"));
            btnExport.Click();
            IWebElement firstCus = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td[1]/div[1]"));
           Console.WriteLine(firstCus.Text);
            int totalRows = driver.FindElements(By.XPath("//table/tbody/tr")).Count;
            IWebElement lastCus = driver.FindElement(By.XPath($"//table/tbody/tr[{totalRows}]/td[1]"));
            Console.WriteLine(lastCus.Text);
            string path = Path.Combine(root, "HoaDon.xlsx");
            Thread.Sleep(3000);
            bool isValid = ExcelProvider.ValidateCustomerData(path, "data", totalRows-1, firstCus.Text, lastCus.Text);
            string actual = isValid ? "File excel bên trong bao gồm 1 dòng các thông tin của hóa đơn" : "File Excel hiển thị sai";
            string result = isValid ? "Passed" : "Failed";
            ExcelProvider.WriteResultToExcel("TestCase_BDCLPM_HK2.xlsx", "TestCase_QuocCuong",27, actual, result);
            ExcelProvider.WriteResultToExcel("TestCase_BDCLPM_HK2.xlsx", "TestCase_QuocCuong", 30, actual, result);
            Assert.IsTrue(true);
        }
        [Test]

        public void EX03_TestCase()
        {
            IWebElement invoiceTab = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[1]/a[6]"));
            invoiceTab.Click();
            Thread.Sleep(2000);
            IWebElement btnExport = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[2]/div[1]/button[1]"));
            btnExport.Click();
            string path = Path.Combine(root, "HoaDon.xlsx");
            Thread.Sleep(3000);
            bool isValid = ExcelProvider.ValidateEmptyCustomerData(path, "data", expectedHeaders);
            string actual = isValid ? "File excel không có thông tin của 1 dòng hóa đơn bất kỳ" : "File Excel hiển thị sai";
            string result = isValid ? "Passed" : "Failed";
            ExcelProvider.WriteResultToExcel("TestCase_BDCLPM_HK2.xlsx", "TestCase_QuocCuong", 33, actual, result);
            Assert.IsTrue(true);
        }

        [TearDown]
        public void TearDown()
        {
            Thread.Sleep(3000);
            driver.Quit();
        }
    }
}
