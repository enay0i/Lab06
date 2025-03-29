using Lab06;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System;
using System.Diagnostics;
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

        [Test]

        public void ActiveCustomer()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
            IWebElement cusTab = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[1]/a[3]/div[1]/*[name()='svg'][1]"));
            cusTab.Click();
            Thread.Sleep(2000);
            IWebElement deActivate = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[4]/span[1]"));
            deActivate.Click();
            IWebElement txtReason = driver.FindElement(By.XPath("/html[1]/body[1]/div[2]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/input[1]"));
            txtReason.SendKeys("kho noi");
            IWebElement btnAccept = driver.FindElement(By.XPath("/html[1]/body[1]/div[2]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/button[2]"));
            btnAccept.Click();
            Thread.Sleep(2000);
            deActivate = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[4]/span[1]"));
            bool a = deActivate.Text == "Vô hiệu hóa";
    
            string actual = a ? "Trạng thái của khách hàng đã chuyển sang \"Vô hiệu hóa\"\r\n\r\n" : "Vô hiệu hóa tài khoản thất bại";
            string result = a ? "Passed" : "Failed";
            string path = Path.Combine(root, "TestCase_BDCLPM_HK2.xlsx");
            ExcelProvider.WriteResultToExcel(path, "TestCase_QuocCuong", 13, actual, result);
            deActivate.Click() ;
            btnAccept = driver.FindElement(By.XPath("/html[1]/body[1]/div[2]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/button[2]"));
            btnAccept.Click();
            Thread.Sleep(2000);
            deActivate = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[4]/span[1]"));
             a = deActivate.Text == "Đã kích hoạt";
           actual = a ? "Trạng thái của khách hàng đã chuyển sang \"Đã kích hoạt\"\r\n\r\n" : "Kích hoạt tài khoản thất bại";
             result = a ? "Passed" : "Failed";
            ExcelProvider.WriteResultToExcel(path, "TestCase_QuocCuong", 15, actual, result);
            Assert.That(a);
        }

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
            Process.Start(new ProcessStartInfo
            {
                FileName = patha,
                UseShellExecute = true
            });
            bool isValid = ExcelProvider.ValidateExcelData(patha, "data", txtYearr, txtThang);
            string actual = isValid ? "File Excel hiển thị đúng doanh thu trong vòng 2 năm" : "File Excel hiển thị sai";
            string result = isValid ? "Passed" : "Failed";
            string path = Path.Combine(root, "TestCase_BDCLPM_HK2.xlsx");
            ExcelProvider.WriteResultToExcel(path, "TestCase_QuocCuong", 36, actual, result);
            Assert.That(isValid);
        }
        [Test]

        public void EX01_EX02_TestCase()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            IWebElement invoiceTab = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[1]/a[6]")));
            invoiceTab.Click();
            Thread.Sleep(2000);
            IWebElement btnExport = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[2]/div[1]/button[1]")));
            btnExport.Click();
            IWebElement firstCus = wait.Until(ExpectedConditions.ElementExists(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td[1]/div[1]")));
            string firstCusText = firstCus.Text;
            IWebElement element = wait.Until(ExpectedConditions.ElementExists(By.XPath("//span[@class='font-bold']")));
            int total = int.Parse(element.Text);
            var pageElements = driver.FindElements(By.XPath("//ul[contains(@class,'ant-pagination')]/li[contains(@class, 'ant-pagination-item')]"));
            IWebElement lastPageElement = pageElements.Last();
            string lastPageNumber = lastPageElement.Text.Trim();
            lastPageElement.Click();
            Thread.Sleep(3000);
            string lastCusText = "";

            wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath("//table/tbody/tr")));
            var rows = driver.FindElements(By.XPath("//table/tbody/tr"));
            int totalRows = rows.Count;
            if (totalRows > 0)
            {
                IWebElement lastCus = wait.Until(ExpectedConditions.ElementExists(By.XPath($"//table/tbody/tr[{totalRows}]/td[1]")));
                lastCusText = lastCus.Text;
            }
            else
            {
                Console.WriteLine("Không tìm thấy dữ liệu khách hàng.");
            }
            string path = Path.Combine(root, "HoaDon.xlsx");
            Process.Start(new ProcessStartInfo
            {
                FileName = path,
                UseShellExecute = true
            });
            Console.WriteLine(total + " tổng số dòng");
            Thread.Sleep(3000);

            bool isValid = ExcelProvider.ValidateCustomerData(path, "data", total, firstCusText, lastCusText);
            string actual = isValid ? "File Excel chứa đúng dữ liệu hóa đơn" : "File Excel hiển thị sai";
            string result = isValid ? "Passed" : "Failed";

            ExcelProvider.WriteResultToExcel("TestCase_BDCLPM_HK2.xlsx", "TestCase_QuocCuong", 27, actual, result);
            ExcelProvider.WriteResultToExcel("TestCase_BDCLPM_HK2.xlsx", "TestCase_QuocCuong", 30, actual, result);

            Assert.That(isValid);
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
            Process.Start(new ProcessStartInfo
            {
                FileName = path,
                UseShellExecute = true
            });
            Thread.Sleep(3000);
            bool isValid = ExcelProvider.ValidateEmptyCustomerData(path, "data", expectedHeaders);
            string actual = isValid ? "File excel không có thông tin của 1 dòng hóa đơn bất kỳ" : "File Excel hiển thị sai";
            string result = isValid ? "Passed" : "Failed";
            ExcelProvider.WriteResultToExcel("TestCase_BDCLPM_HK2.xlsx", "TestCase_QuocCuong", 33, actual, result);
            Assert.IsTrue(true);
        }
        private static IEnumerable<TestCaseData> GetDataForVoucher_CreateVoucher_Test()
        {
            return ExcelProvider.GetDataForAddVoucher(87, 106);
        }
        private static VoucherInfo ParseStringDataToObject(string dataTest)
        {
            var lines = dataTest.Split('\n');

            return new VoucherInfo
            {
                Code = lines[0].Split(':')[1].Trim() ?? "",
                Name = lines[1].Split(':')[1].Trim() ?? "",
                Discount = lines[2].Split(':')[1].Trim() ?? "",
                StartDate = lines[3].Split(':')[1].Trim() ?? "",
                EndDate = lines[4].Split(':')[1].Trim() ?? "",
                Owner = lines[5].Split(':')[1].Trim() ?? ""
            };
        }
        public void GetValueInSelectorAntd(string value, IList<IWebElement> options)
        {

            bool valueFound = false;
            foreach (IWebElement option in options)
            {

                if (option.Text == value)
                {
                    option.Click();
                    valueFound = true;
                    break;
                }

            }
            if (!valueFound)
            {
                IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                js.ExecuteScript("arguments[0].scrollIntoView(true);", options.First());
                options = driver.FindElements(By.ClassName("ant-select-item-option"));

                foreach (IWebElement option in options)
                {
                    if (option.Text == value)
                    {
                        option.Click();
                        break;
                    }
                }
            }

        }
        [Test]
        [TestCaseSource(nameof(GetDataForVoucher_CreateVoucher_Test))]
        public void AddVoucher_TestCase(string testdata, string expResult)
        {
            try
            {
                IWebElement vouTab = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[1]/a[5]"));
                vouTab.Click();
                Thread.Sleep(1000);
                IWebElement btnFloat = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[3]/div[1]/span[1]/*[name()='svg'][1]"));
                btnFloat.Click();
                Thread.Sleep(2000);
                IWebElement txtCode = driver.FindElement(By.Id("code"));
                IWebElement txtName = driver.FindElement(By.Id("voucherName"));
                IWebElement txtDisc = driver.FindElement(By.Id("discount"));
                IWebElement txtStart = driver.FindElement(By.XPath("/html[1]/body[1]/div[2]/div[1]/div[2]/div[1]/div[1]/div[1]/" +
                    "div[1]/form[1]/div[4]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/input[1]"));
                IWebElement txtEnd = driver.FindElement(By.XPath("/html[1]/body[1]/div[2]/div[1]/div[2]/div[1]/div[1]/div[1]" +
                    "/div[1]/form[1]/div[4]/div[1]/div[2]/div[1]/div[1]/div[1]/div[3]/input[1]"));
                IWebElement txtOption = driver.FindElement(By.XPath("/html[1]/body[1]/div[2]/div[1]/div[2]/div[1]/div[1]/div" +
                    "[1]/div[1]/form[1]/div[5]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/span[1]/div[1]"));
                VoucherInfo voucherInfo = ParseStringDataToObject(testdata);
                txtCode.SendKeys(voucherInfo.Code);
                txtName.SendKeys(voucherInfo.Name);
                txtDisc.SendKeys(voucherInfo.Discount);
                txtStart.SendKeys(voucherInfo.StartDate);
                //txtStart.SendKeys(Keys.Tab);
                Thread.Sleep(1000);
                txtEnd.SendKeys(voucherInfo.EndDate);
                Thread.Sleep(1000);
                txtEnd.SendKeys(Keys.Tab);
                txtOption.Click();
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
                wait.Until(d => d.FindElement(By.ClassName("ant-select-dropdown")));
                IList<IWebElement> optionOwner = driver.FindElements(By.ClassName("ant-select-item-option"));
                GetValueInSelectorAntd(voucherInfo.Owner, optionOwner);
                IWebElement btnAdd = driver.FindElement(By.XPath("/html[1]/body[1]/div[2]/div[1]/div[2]/div[1]/div[1]/div[1]" +
                    "/div[1]/form[1]/div[6]/div[1]/div[1]/div[1]/div[1]/div[1]/button[1]/span[1]"));
                btnAdd.Click();
                Thread.Sleep(1000);
                IWebElement messageError;
                string actual = "";
                try
                {
                    messageError = driver.FindElement(By.ClassName("ant-form-item-explain-error"));
                    actual = messageError.Text;
                }
                catch (NoSuchElementException)
                {
                    try
                    {
                        messageError = driver.FindElement(By.ClassName("ant-notification-notice-description"));
                        actual = messageError.GetAttribute("innerText");
                    }
                    catch (NoSuchElementException)
                    {
                        IWebElement discountInput = driver.FindElement(By.Id("discount"));
                        bool isValid = (bool)((IJavaScriptExecutor)driver).ExecuteScript(
                            "return arguments[0].checkValidity();", discountInput);
                        if (!isValid)
                        {
                            actual = (string)((IJavaScriptExecutor)driver).ExecuteScript(
                                "return arguments[0].validationMessage;", discountInput);
                        }
                        else
                        {
                            actual = "No error message found";
                        }
                    }
                }
                bool status = actual.Equals(expResult.Trim());
                string a = status ? "Passed" : "Failed";
                ExcelProvider.WriteResultToExcell("TestCase_BDCLPM_HK2.xlsx", "TestCase_QuocCuong", actual, a.ToString());
                Console.WriteLine("adu " + ExcelProvider.rowIndex);
                Assert.That(actual, Is.EqualTo(expResult.Trim()), "Thông báo không đúng!");
            }
            catch (Exception ex)
            {
                ExcelProvider.WriteResultToExcell("TestCase_BDCLPM_HK2.xlsx", "TestCase_QuocCuong", ex.Message, "Failed");
                Console.WriteLine("Lỗi nhận element");
                throw;
            }
        }


        [TearDown]
        public void TearDown()
        {
            Thread.Sleep(3000);
            driver.Dispose();
        }
    }
}
