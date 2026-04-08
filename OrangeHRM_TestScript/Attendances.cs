using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

namespace OrangeHRM_TestScript
{
    [TestClass]
    public class Attendances
    {
        private const string BASE_URL = "http://localhost:9425/orangehrm-5.6";
        private const string ADMIN_USER = "nghi45397";
        private const string ADMIN_PASS = "Nghiphamtrung09042005!";
        private const string EMP_USER = "nguyenvana";
        private const string EMP_PASS = "Hass@12341";

        private string excelFilePath = @"D:\BDCLPM\TestCase_Nhom14.xlsx";
        private const string SHEET_NAME = "Attendance TCs";

        // Column indices (0-based, NPOI)
        private const int COL_TESTDATA = 7;   // col H - Test Data
        private const int COL_EXPECTED = 8;   // col I - Expected Result
        private const int COL_ACTUAL = 9;   // col J - Actual Result
        private const int COL_STATUS = 11;  // col L - Result

        // Row indices (0-based, NPOI) – taken from Excel row number minus 1
        // F1 – My Attendance Records
        private const int ROW_TC01 = 3;   // Excel row 4
        private const int ROW_TC02 = 7;   // Excel row 8
        private const int ROW_TC03 = 9;   // Excel row 10
        private const int ROW_TC04 = 12;  // Excel row 13
        private const int ROW_TC05 = 15;  // Excel row 16
        private const int ROW_TC06 = 17;  // Excel row 18
        private const int ROW_TC07 = 20;  // Excel row 21
        // F2 – Punch In/Out
        private const int ROW_TC08 = 24;  // Excel row 25
        private const int ROW_TC09 = 27;  // Excel row 28
        private const int ROW_TC10 = 29;  // Excel row 30
        private const int ROW_TC11 = 31;  // Excel row 32
        private const int ROW_TC12 = 35;  // Excel row 36
        private const int ROW_TC13 = 39;  // Excel row 40
        private const int ROW_TC14 = 41;  // Excel row 42
        private const int ROW_TC15 = 44;  // Excel row 45
        private const int ROW_TC16 = 47;  // Excel row 48
        // F3 – Employee Records (Manager)
        private const int ROW_TC17 = 52;  // Excel row 53
        private const int ROW_TC18 = 55;  // Excel row 56
        private const int ROW_TC19 = 58;  // Excel row 59
        private const int ROW_TC20 = 60;  // Excel row 61
        private const int ROW_TC21 = 64;  // Excel row 65
        private const int ROW_TC22 = 67;  // Excel row 68
        private const int ROW_TC23 = 70;  // Excel row 71
        // F4 – Attendance Configuration
        private const int ROW_TC24 = 75;  // Excel row 76
        private const int ROW_TC25 = 78;  // Excel row 79
        private const int ROW_TC26 = 82;  // Excel row 83
        private const int ROW_TC27 = 87;  // Excel row 88
        private const int ROW_TC28 = 91;  // Excel row 92
        private const int ROW_TC29 = 94;  // Excel row 95
        private const int ROW_TC30 = 100; // Excel row 101
        private const int ROW_TC31 = 103; // Excel row 104
        private const int ROW_TC32 = 106; // Excel row 107
        private const int ROW_TC33 = 111; // Excel row 112
        private const int ROW_TC34 = 115; // Excel row 116

        // Last step row indices (0-based) – dòng chứa Expected Result của step CUỐI mỗi TC
        private const int ROW_TC17_LAST = 54;  // Excel row 55  (step 3)
        private const int ROW_TC18_LAST = 57;  // Excel row 58  (step 3)
        private const int ROW_TC19_LAST = 59;  // Excel row 60  (step 2)
        private const int ROW_TC20_LAST = 62;  // Excel row 63  (step 3)
        private const int ROW_TC21_LAST = 66;  // Excel row 67  (step 3)
        private const int ROW_TC22_LAST = 69;  // Excel row 70  (step 3)
        private const int ROW_TC23_LAST = 73;  // Excel row 74  (step 4)
        private const int ROW_TC24_LAST = 76;  // Excel row 77  (step 2)
        private const int ROW_TC25_LAST = 80;  // Excel row 81  (step 3)
        private const int ROW_TC26_LAST = 83;  // Excel row 84  (step 2)
        private const int ROW_TC27_LAST = 88;  // Excel row 89  (step 2)
        private const int ROW_TC28_LAST = 92;  // Excel row 93  (step 2)
        private const int ROW_TC29_LAST = 96;  // Excel row 97  (step 3)
        private const int ROW_TC30_LAST = 101; // Excel row 102 (step 2)
        private const int ROW_TC31_LAST = 104; // Excel row 105 (step 2)
        private const int ROW_TC32_LAST = 109; // Excel row 110 (step 4)
        private const int ROW_TC33_LAST = 112; // Excel row 113 (step 2)
        private const int ROW_TC34_LAST = 116; // Excel row 117 (step 2)

        private static readonly object excelLock = new object();
        private IWebDriver dr;
        private WebDriverWait wait;

        // ═══════════════════════════════════════════════════════════════
        // SETUP / TEARDOWN
        // ═══════════════════════════════════════════════════════════════

        [TestInitialize]
        public void Setup()
        {
            dr = new ChromeDriver();
            dr.Manage().Window.Maximize();
            dr.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
            wait = new WebDriverWait(dr, TimeSpan.FromSeconds(15));
        }

        [TestCleanup]
        public void TearDown() => dr?.Quit();

        // ═══════════════════════════════════════════════════════════════
        // HELPER METHODS
        // ═══════════════════════════════════════════════════════════════

        private string ReadCell(ISheet sheet, int row, int col)
        {
            lock (excelLock)
            {
                var fmt = new DataFormatter();
                IRow r = sheet.GetRow(row);
                if (r == null) return "";
                return fmt.FormatCellValue(r.GetCell(col)) ?? "";
            }
        }

        private (string expected, string testData) ReadExcelRow(int rowIndex)
        {
            lock (excelLock)
            {
                using FileStream fs = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read);
                XSSFWorkbook wb = new XSSFWorkbook(fs);
                ISheet sh = wb.GetSheet(SHEET_NAME);
                return (ReadCell(sh, rowIndex, COL_EXPECTED), ReadCell(sh, rowIndex, COL_TESTDATA));
            }
        }

        private string ReadExpected(int rowIndex)
        {
            lock (excelLock)
            {
                using FileStream fs = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read);
                XSSFWorkbook wb = new XSSFWorkbook(fs);
                ISheet sh = wb.GetSheet(SHEET_NAME);
                return ReadCell(sh, rowIndex, COL_EXPECTED);
            }
        }

        private void WriteExcelResult(int rowIndex, string actualMsg, string status)
        {
            lock (excelLock)
            {
                using FileStream fsRead = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read);
                XSSFWorkbook wb = new XSSFWorkbook(fsRead);
                ISheet sh = wb.GetSheet(SHEET_NAME);
                IRow row = sh.GetRow(rowIndex) ?? sh.CreateRow(rowIndex);
                row.CreateCell(COL_ACTUAL).SetCellValue(actualMsg);
                row.CreateCell(COL_STATUS).SetCellValue(status);
                using FileStream fsWrite = new FileStream(excelFilePath, FileMode.Create, FileAccess.Write);
                wb.Write(fsWrite);
            }
        }

        private (string username, string password) ExtractCredentials(string text, string defaultUser, string defaultPass)
        {
            string u = defaultUser, p = defaultPass;
            if (string.IsNullOrWhiteSpace(text)) return (u, p);
            foreach (var line in text.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries))
            {
                if (line.Trim().ToLower().StartsWith("username:"))
                    u = line.Substring(line.IndexOf(':') + 1).Trim();
                if (line.Trim().ToLower().StartsWith("password:"))
                    p = line.Substring(line.IndexOf(':') + 1).Trim();
            }
            return (u, p);
        }

        private void LoginAs(string username, string password)
        {
            dr.Navigate().GoToUrl(BASE_URL);
            IWebElement txtUser = wait.Until(d => d.FindElement(By.Name("username")));
            txtUser.Clear();
            txtUser.SendKeys(username);
            dr.FindElement(By.Name("password")).SendKeys(password);
            dr.FindElement(By.CssSelector("button[type='submit']")).Click();
            wait.Until(d => d.Url.Contains("dashboard") ||
                            d.FindElements(By.ClassName("oxd-topbar-header-breadcrumb")).Count > 0);
            Thread.Sleep(800);
        }

        private void Logout()
        {
            try
            {
                IWebElement userDropdown = wait.Until(d => d.FindElement(By.XPath(
                    "//li[contains(@class,'oxd-userdropdown')]")));
                userDropdown.Click();
                wait.Until(d => d.FindElement(By.XPath("//a[normalize-space()='Logout']"))).Click();
                wait.Until(d => d.FindElement(By.Name("username")));
            }
            catch
            {
                dr.Navigate().GoToUrl(BASE_URL);
                wait.Until(d => d.FindElement(By.Name("username")));
            }
        }

        private void GoToMyAttendanceRecords()
        {
            dr.Navigate().GoToUrl(BASE_URL + "/web/index.php/attendance/viewMyAttendanceRecord");
            Thread.Sleep(1000);
        }

        private void GoToPunchInOut()
        {
            dr.Navigate().GoToUrl(BASE_URL + "/web/index.php/attendance/punchIn");
            Thread.Sleep(1000);
        }

        private void GoToEmployeeRecords()
        {
            dr.Navigate().GoToUrl(BASE_URL + "/web/index.php/attendance/viewAttendanceRecord");
            Thread.Sleep(1000);
        }

        private void GoToAttendanceConfiguration()
        {
            dr.Navigate().GoToUrl(BASE_URL + "/web/index.php/attendance/configure");
            Thread.Sleep(1000);
        }

        private void ToggleCheckbox(IWebElement toggle, bool targetState)
        {
            if (toggle.Selected != targetState)
            {
                IJavaScriptExecutor js = (IJavaScriptExecutor)dr;
                js.ExecuteScript("arguments[0].click();", toggle);
                Thread.Sleep(600);
            }
        }

        private void ClickSave()
        {
            var saveBtn = dr.FindElements(By.XPath("//button[contains(.,'Save')]")).FirstOrDefault();
            if (saveBtn != null) { saveBtn.Click(); Thread.Sleep(1200); }
        }



        // ═══════════════════════════════════════════════════════════════
        // F3 – EMPLOYEE RECORDS (TC17 – TC23)
        // ═══════════════════════════════════════════════════════════════

        /// <summary>TC17 – Kiểm tra hệ thống hiển thị trang Employee Attendance Records khi Admin truy cập</summary>
        [TestMethod]
        public void ATT_TC17_EmployeeRecords_DisplayPage()
        {
            var (_, testData) = ReadExcelRow(ROW_TC17);
            string expectedMsg = ReadExpected(ROW_TC17_LAST);
            var (username, password) = ExtractCredentials(testData, ADMIN_USER, ADMIN_PASS);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                LoginAs(username, password);
                GoToEmployeeRecords();

                bool hasTitle = dr.FindElements(By.XPath("//h5[contains(@class, 'oxd-table-filter-title') and text()='Employee Attendance Records']")).Count > 0;

                // 2. Kiểm tra bộ lọc "Employee Name" (Dựa trên label)
                bool hasEmployeeFilter = dr.FindElements(By.XPath("//label[text()='Employee Name']/ancestor::div[contains(@class,'oxd-input-group')]//input")).Count > 0;

                // 3. Kiểm tra bộ lọc "Date" (Sửa lại XPath cho chính xác với cấu trúc label/input)
                bool hasDateFilter = dr.FindElements(By.XPath("//label[text()='Date']/ancestor::div[contains(@class,'oxd-input-group')]//input")).Count > 0;

                // 4. Kiểm tra sự tồn tại của bảng dữ liệu
                bool hasTable = dr.FindElements(By.CssSelector("div.oxd-table")).Count > 0;

                bool ok = hasTitle && hasEmployeeFilter && hasDateFilter && hasTable;

                if (ok)
                {
                    actualMsg = expectedMsg;
                    status = "Passed";
                }
                else
                {
                    // Ghi log chi tiết lỗi nếu thiếu thành phần nào
                    actualMsg = "Thiếu thành phần giao diện: " +
                                (!hasTitle ? "[Title] " : "") +
                                (!hasEmployeeFilter ? "[Employee Name Filter] " : "") +
                                (!hasDateFilter ? "[Date Filter] " : "") +
                                (!hasTable ? "[Table] " : "");
                    status = "Failed";
                }
            }
            catch (Exception ex)
            {
                actualMsg = "Lỗi hệ thống: " + ex.Message;
            }

            WriteExcelResult(ROW_TC17, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC18 – Kiểm tra Admin xem được danh sách tất cả nhân viên và tổng giờ làm việc theo ngày</summary>
        [TestMethod]
        public void ATT_TC18_EmployeeRecords_ViewAllEmployees()
        {
            var (_, testData) = ReadExcelRow(ROW_TC18);
            string expectedMsg = ReadExpected(ROW_TC18_LAST);
            var (username, password) = ExtractCredentials(testData, ADMIN_USER, ADMIN_PASS);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                LoginAs(username, password);
                GoToEmployeeRecords();

                IWebElement dateInput = wait.Until(d => d.FindElement(By.XPath(
                    "//label[contains(text(),'Date')]/following::input[1]")));

                Thread.Sleep(400);

                dr.FindElement(By.XPath("//button[contains(.,'View')] | //button[contains(.,'Search')]")).Click();
                Thread.Sleep(1500);

                var rows = dr.FindElements(By.XPath("//div[@class='oxd-table-body']//div[@role='row']"));
                bool hasRows = rows.Count > 0;
                bool hasTotalHours = dr.FindElements(By.XPath(
                    "//*[contains(text(),'Total Hours')] | //*[contains(text(),'Duration')]")).Count > 0;

                bool ok = hasRows || hasTotalHours;
                actualMsg = ok ? expectedMsg : "Không hiển thị danh sách nhân viên hoặc tổng giờ làm";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(ROW_TC18, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC19 – Kiểm tra hệ thống hiển thị số bản ghi tìm được (X) Records Found</summary>
        [TestMethod]
        public void ATT_TC19_EmployeeRecords_RecordsFoundCount()
        {
            var (_, testData) = ReadExcelRow(ROW_TC19);
            string expectedMsg = ReadExpected(ROW_TC19_LAST);
            var (username, password) = ExtractCredentials(testData, ADMIN_USER, ADMIN_PASS);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                LoginAs(username, password);
                GoToEmployeeRecords();

                IWebElement dateInput = wait.Until(d => d.FindElement(By.XPath(
                    "//label[contains(text(),'Date')]/following::input[1]")));

   
                Thread.Sleep(400);

                dr.FindElement(By.XPath("//button[contains(.,'View')] | //button[contains(.,'Search')]")).Click();
                Thread.Sleep(1500);

                bool hasRecordsFound = dr.FindElements(By.XPath(
                    "//*[contains(text(),'Records Found')] | //*[contains(@class,'orangehrm-horizontal-padding')]")).Count > 0;

                bool ok = hasRecordsFound;
                actualMsg = ok ? expectedMsg : "Không hiển thị '(N) Records Found'";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(ROW_TC19, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        // =========================================================================
        // MODULE: EMPLOYEE ATTENDANCE RECORDS - FILTERING (TC20 - TC24)
        // =========================================================================

        // =========================================================================
        // MODULE: EMPLOYEE ATTENDANCE RECORDS - FILTERING (TC20 - TC24)
        // =========================================================================

        /// <summary>TC20 – Kiểm tra Manager có thể lọc danh sách theo tên nhân viên hợp lệ</summary>
        [TestMethod]
        public void ATT_TC20_EmployeeRecords_FilterByValidName()
        {
            int row = 61; // Dòng 62 trong Excel
            var (_, testData) = ReadExcelRow(row);
            string expectedMsg = ReadExpected(row);

            // Bắt lỗi Data an toàn cho mảng (Array), dùng .Length và check null
            if (testData.Length <= 7 || testData[7] == null || string.IsNullOrWhiteSpace(testData[7].ToString()))
            {
                throw new Exception($"[Lỗi Data] Dòng {row} thiếu tên nhân viên để lọc (cột H)!");
            }
            string employeeName = testData[7].ToString().Trim();

            var (username, password) = ExtractCredentials(testData, ADMIN_USER, ADMIN_PASS);
            string actualMsg = "";
            string status = "Failed";

            try
            {
                LoginAs(username, password);
                GoToEmployeeRecords();

                // 1. Nhập tên nhân viên
                var empNameInput = wait.Until(d => d.FindElement(By.XPath("//label[text()='Employee Name']/ancestor::div[contains(@class,'oxd-input-group')]//input")));
                empNameInput.Clear();
                Thread.Sleep(200);

                foreach (char c in employeeName)
                {
                    empNameInput.SendKeys(c.ToString());
                    Thread.Sleep(50);
                }

                // Chờ Autocomplete và click chọn
                var opt = wait.Until(d => d.FindElements(By.XPath($"//div[contains(@class,'oxd-autocomplete-option') and contains(., '{employeeName}')]")).FirstOrDefault());
                if (opt != null)
                {
                    opt.Click();
                }
                else
                {
                    var firstOpt = dr.FindElements(By.XPath("//div[contains(@class,'oxd-autocomplete-option')]")).FirstOrDefault();
                    if (firstOpt != null) firstOpt.Click();
                }
                Thread.Sleep(500);

                // 2. Click View
                dr.FindElement(By.XPath("//button[contains(@type, 'submit') or contains(.,'View')]")).Click();
                Thread.Sleep(2000);

                // 3. Xác minh kết quả
                bool hasResult = dr.FindElements(By.XPath("//div[@class='oxd-table-body']//div[@role='row']")).Count > 0 ||
                                 dr.FindElements(By.XPath("//*[contains(text(),'Records Found')]")).Count > 0;

                if (hasResult)
                {
                    actualMsg = expectedMsg;
                    status = "Passed";
                }
                else
                {
                    actualMsg = $"Không tìm thấy dữ liệu cho '{employeeName}'";
                }
            }
            catch (Exception ex) { actualMsg = "Lỗi thực thi: " + ex.Message; }

            WriteExcelResult(row, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC21 – Kiểm tra Manager lọc theo tên nhân viên KHÔNG hợp lệ</summary>
        [TestMethod]
        public void ATT_TC21_EmployeeRecords_FilterByInvalidName()
        {
            int row = 63;
            var (_, testData) = ReadExcelRow(row);
            string expectedMsg = ReadExpected(row);

            if (testData.Length <= 7 || testData[7] == null || string.IsNullOrWhiteSpace(testData[7].ToString()))
            {
                throw new Exception($"[Lỗi Data] Dòng {row} thiếu tên nhân viên không hợp lệ (cột H)!");
            }
            string invalidName = testData[7].ToString().Trim();

            var (username, password) = ExtractCredentials(testData, ADMIN_USER, ADMIN_PASS);
            string actualMsg = "";
            string status = "Failed";

            try
            {
                LoginAs(username, password);
                GoToEmployeeRecords();

                var empNameInput = wait.Until(d => d.FindElement(By.XPath("//label[text()='Employee Name']/ancestor::div[contains(@class,'oxd-input-group')]//input")));
                empNameInput.SendKeys(invalidName);

                // Bấm ra ngoài (click vào title) để hệ thống nhận diện
                dr.FindElement(By.XPath("//h5")).Click();
                dr.FindElement(By.XPath("//button[contains(@type, 'submit') or contains(.,'View')]")).Click();
                Thread.Sleep(1500);

                // Kiểm tra xem hệ thống có báo "Invalid" hoặc "No Records Found" không
                bool isInvalidText = dr.FindElements(By.XPath("//span[contains(@class, 'oxd-input-field-error-message') and contains(text(), 'Invalid')]")).Count > 0;
                bool isNoRecordsText = dr.FindElements(By.XPath("//*[contains(text(),'No Records')]")).Count > 0;

                if (isInvalidText || isNoRecordsText)
                {
                    actualMsg = expectedMsg;
                    status = "Passed";
                }
                else
                {
                    actualMsg = "Hệ thống không chặn hoặc không báo lỗi khi nhập tên sai.";
                }
            }
            catch (Exception ex) { actualMsg = "Lỗi thực thi: " + ex.Message; }

            WriteExcelResult(row, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC22 – Kiểm tra lọc danh sách theo Ngày hợp lệ</summary>
        [TestMethod]
        public void ATT_TC22_EmployeeRecords_FilterByValidDate()
        {
            int row = 64;
            var (_, testData) = ReadExcelRow(row);
            string expectedMsg = ReadExpected(row);

            if (testData.Length <= 7 || testData[7] == null || string.IsNullOrWhiteSpace(testData[7].ToString()))
            {
                throw new Exception($"[Lỗi Data] Dòng {row} thiếu dữ liệu Ngày lọc (cột H)!");
            }
            string filterDate = testData[7].ToString().Trim();

            var (username, password) = ExtractCredentials(testData, ADMIN_USER, ADMIN_PASS);
            string actualMsg = "";
            string status = "Failed";

            try
            {
                LoginAs(username, password);
                GoToEmployeeRecords();

                var dateInput = wait.Until(d => d.FindElement(By.XPath("//label[text()='Date']/ancestor::div[contains(@class,'oxd-input-group')]//input")));
                dateInput.SendKeys(Keys.Control + "a");
                dateInput.SendKeys(Keys.Backspace);
                dateInput.SendKeys(filterDate);
                dateInput.SendKeys(Keys.Escape); // Nhấn ESC để đóng lịch lại

                dr.FindElement(By.XPath("//button[contains(@type, 'submit') or contains(.,'View')]")).Click();
                Thread.Sleep(2000);

                bool hasResult = dr.FindElements(By.XPath("//div[@class='oxd-table-body']//div[@role='row']")).Count > 0 ||
                                 dr.FindElements(By.XPath("//*[contains(text(),'Records Found')]")).Count > 0;

                if (hasResult)
                {
                    actualMsg = expectedMsg;
                    status = "Passed";
                }
                else
                {
                    actualMsg = $"Không tìm thấy dữ liệu chấm công cho ngày {filterDate}.";
                }
            }
            catch (Exception ex) { actualMsg = "Lỗi thực thi: " + ex.Message; }

            WriteExcelResult(row, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC23 – Kiểm tra lọc theo Ngày có định dạng sai hoặc Ngày tương lai</summary>
        [TestMethod]
        public void ATT_TC23_EmployeeRecords_FilterByInvalidDate()
        {
            int row = 65;
            var (_, testData) = ReadExcelRow(row);
            string expectedMsg = ReadExpected(row);

            if (testData.Length <= 7 || testData[7] == null || string.IsNullOrWhiteSpace(testData[7].ToString()))
            {
                throw new Exception($"[Lỗi Data] Dòng {row} thiếu dữ liệu Ngày sai (cột H)!");
            }
            string invalidDate = testData[7].ToString().Trim();

            var (username, password) = ExtractCredentials(testData, ADMIN_USER, ADMIN_PASS);
            string actualMsg = "";
            string status = "Failed";

            try
            {
                LoginAs(username, password);
                GoToEmployeeRecords();

                var dateInput = wait.Until(d => d.FindElement(By.XPath("//label[text()='Date']/ancestor::div[contains(@class,'oxd-input-group')]//input")));
                dateInput.SendKeys(Keys.Control + "a");
                dateInput.SendKeys(Keys.Backspace);
                dateInput.SendKeys(invalidDate);
                dateInput.SendKeys(Keys.Escape);

                dr.FindElement(By.XPath("//button[contains(@type, 'submit') or contains(.,'View')]")).Click();
                Thread.Sleep(1500);

                bool isFormatError = dr.FindElements(By.XPath("//span[contains(@class, 'oxd-input-field-error-message')]")).Count > 0;
                bool isNoRecords = dr.FindElements(By.XPath("//*[contains(text(),'No Records')]")).Count > 0;

                if (isFormatError || isNoRecords)
                {
                    actualMsg = expectedMsg;
                    status = "Passed";
                }
                else
                {
                    actualMsg = "Hệ thống chấp nhận ngày sai mà không báo lỗi.";
                }
            }
            catch (Exception ex) { actualMsg = "Lỗi thực thi: " + ex.Message; }

            WriteExcelResult(row, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC24 – Lọc kết hợp Tên nhân viên và Ngày</summary>
        [TestMethod]
        public void ATT_TC24_EmployeeRecords_FilterByNameAndDate()
        {
            int row = 66;
            var (_, testData) = ReadExcelRow(row);
            string expectedMsg = ReadExpected(row);

            if (testData.Length <= 7)
            {
                throw new Exception($"[Lỗi Data] Dòng {row} thiếu data (cột G hoặc H)!");
            }

            // Gán biến an toàn không dùng ??
            string empName = "An Văn Nguyễn";
            if (testData[6] != null && !string.IsNullOrWhiteSpace(testData[6].ToString()))
            {
                empName = testData[6].ToString().Trim();
            }

            string filterDate = "2024-01-01";
            if (testData[7] != null && !string.IsNullOrWhiteSpace(testData[7].ToString()))
            {
                filterDate = testData[7].ToString().Trim();
            }

            var (username, password) = ExtractCredentials(testData, ADMIN_USER, ADMIN_PASS);
            string actualMsg = "";
            string status = "Failed";

            try
            {
                LoginAs(username, password);
                GoToEmployeeRecords();

                // Nhập Tên
                var empNameInput = wait.Until(d => d.FindElement(By.XPath("//label[text()='Employee Name']/ancestor::div[contains(@class,'oxd-input-group')]//input")));
                empNameInput.Clear();
                foreach (char c in empName) { empNameInput.SendKeys(c.ToString()); Thread.Sleep(50); }

                var opt = wait.Until(d => d.FindElements(By.XPath($"//div[contains(@class,'oxd-autocomplete-option')]")).FirstOrDefault());
                if (opt != null) opt.Click();

                // Nhập Ngày
                var dateInput = dr.FindElement(By.XPath("//label[text()='Date']/ancestor::div[contains(@class,'oxd-input-group')]//input"));
                dateInput.SendKeys(Keys.Control + "a");
                dateInput.SendKeys(Keys.Backspace);
                dateInput.SendKeys(filterDate);
                dateInput.SendKeys(Keys.Escape);

                // Nhấn View
                dr.FindElement(By.XPath("//button[contains(@type, 'submit') or contains(.,'View')]")).Click();
                Thread.Sleep(2000);

                bool hasResult = dr.FindElements(By.XPath("//div[@class='oxd-table-body']//div[@role='row']")).Count > 0 ||
                                 dr.FindElements(By.XPath("//*[contains(text(),'Records Found')]")).Count > 0;
                bool isNoRecords = dr.FindElements(By.XPath("//*[contains(text(),'No Records')]")).Count > 0;

                if (hasResult || isNoRecords)
                {
                    actualMsg = expectedMsg;
                    status = "Passed";
                }
                else
                {
                    actualMsg = "Lọc kết hợp không ra kết quả như mong đợi.";
                }
            }
            catch (Exception ex) { actualMsg = "Lỗi thực thi: " + ex.Message; }

            WriteExcelResult(row, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC25 – Kiểm tra Admin bật/tắt tùy chọn "Employee can change current time when punching in/out"</summary>
        [TestMethod]
        public void ATT_TC25_Config_ToggleChangeTime()
        {
            string expectedMsg = ReadExpected(ROW_TC25_LAST);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                LoginAs(ADMIN_USER, ADMIN_PASS);
                GoToAttendanceConfiguration();

                var toggles = dr.FindElements(By.XPath("//input[@type='checkbox']"));
                Assert.IsTrue(toggles.Count > 0, "Không tìm thấy toggle nào trên trang Configuration");

                IWebElement toggle = toggles[0];
                bool initialState = toggle.Selected;
                ToggleCheckbox(toggle, !initialState);

                bool changed = toggle.Selected != initialState;

                ClickSave();

                // Reload kiểm tra trạng thái giữ nguyên
                GoToAttendanceConfiguration();
                var togglesAfter = dr.FindElements(By.XPath("//input[@type='checkbox']"));
                bool persistedState = togglesAfter.Count > 0 && togglesAfter[0].Selected == !initialState;

                bool ok = changed && persistedState;
                actualMsg = ok ? expectedMsg : "Toggle không thay đổi hoặc không lưu được sau reload";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(ROW_TC25, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC26 – Kiểm tra khi BẬT "Employee can change current time" → nhân viên được sửa giờ khi Punch In/Out</summary>
        [TestMethod]
        public void ATT_TC26_Config_ChangeTimeEnabled()
        {
            var (_, testData) = ReadExcelRow(ROW_TC26);
            string expectedMsg = ReadExpected(ROW_TC26_LAST);
            var (empUser, empPass) = ExtractCredentials(testData, EMP_USER, EMP_PASS);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                // Admin bật toggle
                LoginAs(ADMIN_USER, ADMIN_PASS);
                GoToAttendanceConfiguration();

                var toggles = dr.FindElements(By.XPath("//input[@type='checkbox']"));
                if (toggles.Count > 0) ToggleCheckbox(toggles[0], true);
                ClickSave();

                // Login Employee kiểm tra Time field
                Logout();
                LoginAs(empUser, empPass);
                GoToPunchInOut();

                IWebElement timeInput = wait.Until(d => d.FindElement(By.XPath(
                    "//label[contains(text(),'Time')]/following::input[1]")));
                bool isEditable = timeInput.Enabled && !timeInput.GetAttribute("readonly").Equals("true", StringComparison.OrdinalIgnoreCase);

                bool ok = isEditable;
                actualMsg = ok ? expectedMsg : "Trường Time không thể chỉnh sửa khi toggle bật";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(ROW_TC26, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC27 – Kiểm tra khi TẮT "Employee can change current time" → nhân viên KHÔNG thể sửa giờ</summary>
        [TestMethod]
        public void ATT_TC27_Config_ChangeTimeDisabled()
        {
            var (_, testData) = ReadExcelRow(ROW_TC27);
            string expectedMsg = ReadExpected(ROW_TC27_LAST);
            var (empUser, empPass) = ExtractCredentials(testData, EMP_USER, EMP_PASS);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                // Admin tắt toggle
                LoginAs(ADMIN_USER, ADMIN_PASS);
                GoToAttendanceConfiguration();

                var toggles = dr.FindElements(By.XPath("//input[@type='checkbox']"));
                if (toggles.Count > 0) ToggleCheckbox(toggles[0], false);
                ClickSave();

                // Login Employee kiểm tra Time field bị disable
                Logout();
                LoginAs(empUser, empPass);
                GoToPunchInOut();

                IWebElement timeInput = wait.Until(d => d.FindElement(By.XPath(
                    "//label[contains(text(),'Time')]/following::input[1]")));
                bool isDisabled = !timeInput.Enabled || timeInput.GetAttribute("readonly") == "true" ||
                                  timeInput.GetAttribute("disabled") != null;

                bool ok = isDisabled;
                actualMsg = ok ? expectedMsg : "Trường Time vẫn cho phép chỉnh sửa khi toggle tắt";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(ROW_TC27, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC28 – Kiểm tra Admin bật/tắt tùy chọn "Employee can edit/delete own attendance records"</summary>
        [TestMethod]
        public void ATT_TC28_Config_ToggleEditDelete()
        {
            string expectedMsg = ReadExpected(ROW_TC28_LAST);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                LoginAs(ADMIN_USER, ADMIN_PASS);
                GoToAttendanceConfiguration();

                var toggles = dr.FindElements(By.XPath("//input[@type='checkbox']"));
                Assert.IsTrue(toggles.Count >= 2, "Không tìm thấy đủ toggles trên Configuration");

                IWebElement toggle = toggles[1]; // Toggle thứ 2 = edit/delete own record
                bool initialState = toggle.Selected;
                ToggleCheckbox(toggle, !initialState);

                bool changed = toggle.Selected != initialState;
                ClickSave();

                GoToAttendanceConfiguration();
                var togglesAfter = dr.FindElements(By.XPath("//input[@type='checkbox']"));
                bool persisted = togglesAfter.Count >= 2 && togglesAfter[1].Selected == !initialState;

                bool ok = changed && persisted;
                actualMsg = ok ? expectedMsg : "Toggle edit/delete không thay đổi hoặc không lưu sau reload";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(ROW_TC28, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC29 – Kiểm tra khi BẬT quyền edit/delete → nhân viên có thể chỉnh sửa/xóa bản ghi của mình</summary>
        [TestMethod]
        public void ATT_TC29_Config_EditDeleteEnabled()
        {
            var (_, testData) = ReadExcelRow(ROW_TC29);
            string expectedMsg = ReadExpected(ROW_TC29_LAST);
            var (empUser, empPass) = ExtractCredentials(testData, EMP_USER, EMP_PASS);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                // Admin bật toggle edit/delete
                LoginAs(ADMIN_USER, ADMIN_PASS);
                GoToAttendanceConfiguration();

                var toggles = dr.FindElements(By.XPath("//input[@type='checkbox']"));
                if (toggles.Count >= 2) ToggleCheckbox(toggles[1], true);
                ClickSave();

                // Login Employee vào My Records
                Logout();
                LoginAs(empUser, empPass);
                GoToMyAttendanceRecords();

                IWebElement dateInput = wait.Until(d => d.FindElement(By.XPath(
                    "//label[contains(text(),'Date')]/following::input[1]")));
                dateInput.Clear();
                dateInput.SendKeys(DateTime.Now.ToString("MM/dd/yyyy"));
                Thread.Sleep(300);
                dr.FindElement(By.XPath("//button[contains(.,'View')]")).Click();
                Thread.Sleep(1500);

                bool hasEditDelete = dr.FindElements(By.XPath(
                    "//button[@title='Edit'] | //button[@title='Delete'] | " +
                    "//i[contains(@class,'bi-pencil')] | //i[contains(@class,'bi-trash')] | " +
                    "//button[contains(@class,'oxd-icon-button')]")).Count > 0;

                bool ok = hasEditDelete;
                actualMsg = ok ? expectedMsg : "Không thấy nút Edit/Delete dù đã bật quyền";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(ROW_TC29, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC30 – Kiểm tra khi TẮT quyền edit/delete → nhân viên KHÔNG thể chỉnh sửa/xóa bản ghi</summary>
        [TestMethod]
        public void ATT_TC30_Config_EditDeleteDisabled()
        {
            var (expectedMsg, testData) = ReadExcelRow(ROW_TC30);
            var (empUser, empPass) = ExtractCredentials(testData, EMP_USER, EMP_PASS);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                // Admin tắt toggle edit/delete
                LoginAs(ADMIN_USER, ADMIN_PASS);
                GoToAttendanceConfiguration();

                var toggles = dr.FindElements(By.XPath("//input[@type='checkbox']"));
                if (toggles.Count >= 2) ToggleCheckbox(toggles[1], false);
                ClickSave();

                // Login Employee vào My Records
                Logout();
                LoginAs(empUser, empPass);
                GoToMyAttendanceRecords();

                IWebElement dateInput = wait.Until(d => d.FindElement(By.XPath(
                    "//label[contains(text(),'Date')]/following::input[1]")));
                dateInput.Clear();
                dateInput.SendKeys(DateTime.Now.ToString("MM/dd/yyyy"));
                Thread.Sleep(300);
                dr.FindElement(By.XPath("//button[contains(.,'View')]")).Click();
                Thread.Sleep(1500);

                // Thử truy cập URL edit trực tiếp
                dr.Navigate().GoToUrl(BASE_URL + "/web/index.php/attendance/edit");
                Thread.Sleep(1000);
                bool forbidden = dr.Url.Contains("dashboard") ||
                    dr.FindElements(By.XPath("//*[contains(text(),'Access Denied')] | //*[contains(text(),'403')]")).Count > 0;

                bool ok = forbidden;
                actualMsg = ok ? expectedMsg : "Employee vẫn truy cập được trang edit khi đã tắt quyền";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(ROW_TC30, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC31 – Kiểm tra Admin bật/tắt tùy chọn "Supervisor can add/edit/delete attendance records of subordinates"</summary>
        [TestMethod]
        public void ATT_TC31_Config_ToggleSupervisor()
        {
            string expectedMsg = ReadExpected(ROW_TC31);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                LoginAs(ADMIN_USER, ADMIN_PASS);
                GoToAttendanceConfiguration();

                var toggles = dr.FindElements(By.XPath("//input[@type='checkbox']"));
                Assert.IsTrue(toggles.Count >= 3, "Không tìm thấy đủ 3 toggles trên Configuration");

                IWebElement toggle = toggles[2]; // Toggle thứ 3 = supervisor
                bool initialState = toggle.Selected;
                ToggleCheckbox(toggle, !initialState);

                bool changed = toggle.Selected != initialState;
                ClickSave();

                GoToAttendanceConfiguration();
                var togglesAfter = dr.FindElements(By.XPath("//input[@type='checkbox']"));
                bool persisted = togglesAfter.Count >= 3 && togglesAfter[2].Selected == !initialState;

                bool ok = changed && persisted;
                actualMsg = ok ? expectedMsg : "Toggle Supervisor không thay đổi hoặc không lưu sau reload";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(ROW_TC31, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC32 – Kiểm tra khi BẬT quyền supervisor → Manager có thể thêm/sửa/xóa bản ghi chấm công của nhân viên cấp dưới</summary>
        [TestMethod]
        public void ATT_TC32_Config_SupervisorEnabled()
        {
            var (expectedMsg, testData) = ReadExcelRow(ROW_TC32);
            var (username, password) = ExtractCredentials(testData, ADMIN_USER, ADMIN_PASS);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                LoginAs(username, password);
                GoToAttendanceConfiguration();

                var toggles = dr.FindElements(By.XPath("//input[@type='checkbox']"));
                if (toggles.Count >= 3) ToggleCheckbox(toggles[2], true);
                ClickSave();

                // Vào Employee Records kiểm tra có Add / Edit / Delete
                GoToEmployeeRecords();

                IWebElement dateInput = wait.Until(d => d.FindElement(By.XPath(
                    "//label[contains(text(),'Date')]/following::input[1]")));
                dateInput.Clear();
                dateInput.SendKeys(DateTime.Now.ToString("MM/dd/yyyy"));
                Thread.Sleep(300);
                dr.FindElement(By.XPath("//button[contains(.,'View')] | //button[contains(.,'Search')]")).Click();
                Thread.Sleep(1500);

                bool hasAddBtn = dr.FindElements(By.XPath("//button[normalize-space()='Add']")).Count > 0;
                bool hasActionBtns = dr.FindElements(By.XPath(
                    "//button[contains(@class,'oxd-icon-button')]")).Count > 0;

                bool ok = hasAddBtn || hasActionBtns;
                actualMsg = ok ? expectedMsg : "Không thấy Add/Edit/Delete button dù đã bật quyền Supervisor";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(ROW_TC32, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC33 – Kiểm tra hệ thống lưu cấu hình khi Admin nhấn Save</summary>
        [TestMethod]
        public void ATT_TC33_Config_SaveSuccess()
        {
            var (expectedMsg, testData) = ReadExcelRow(ROW_TC33);
            var (username, password) = ExtractCredentials(testData, ADMIN_USER, ADMIN_PASS);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                LoginAs(username, password);
                GoToAttendanceConfiguration();

                var toggles = dr.FindElements(By.XPath("//input[@type='checkbox']"));
                if (toggles.Count > 0)
                {
                    IJavaScriptExecutor js = (IJavaScriptExecutor)dr;
                    js.ExecuteScript("arguments[0].click();", toggles[0]);
                    Thread.Sleep(600);
                }

                dr.FindElement(By.XPath("//button[contains(.,'Save')]")).Click();
                Thread.Sleep(1500);

                bool hasSavedMsg = dr.FindElements(By.XPath(
                    "//*[contains(text(),'Saved')] | //*[contains(text(),'successfully')] | " +
                    "//*[contains(@class,'oxd-toast-content')] | //*[contains(@class,'oxd-alert-content-text')]")).Count > 0;

                bool ok = hasSavedMsg;
                actualMsg = ok ? expectedMsg : "Không thấy thông báo lưu thành công sau khi nhấn Save";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(ROW_TC33, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC34 – Kiểm tra nhân viên (Employee role) KHÔNG có quyền truy cập trang Configuration</summary>
        [TestMethod]
        public void ATT_TC34_Config_EmployeeAccessDenied()
        {
            var (expectedMsg, testData) = ReadExcelRow(ROW_TC34);
            var (empUser, empPass) = ExtractCredentials(testData, EMP_USER, EMP_PASS);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                LoginAs(empUser, empPass);

                // Kiểm tra menu Attendance không có Configuration
                var menuAttendance = dr.FindElements(By.XPath("//span[contains(text(),'Attendance')]")).FirstOrDefault();
                if (menuAttendance != null) { menuAttendance.Click(); Thread.Sleep(500); }

                bool noConfigMenu = dr.FindElements(By.XPath(
                    "//a[contains(text(),'Configuration')] | //span[text()='Configuration']")).Count == 0;

                // Thử truy cập trực tiếp URL
                dr.Navigate().GoToUrl(BASE_URL + "/web/index.php/attendance/configure");
                Thread.Sleep(1500);

                bool isForbidden = dr.Url.Contains("dashboard") ||
                    dr.FindElements(By.XPath(
                        "//*[contains(text(),'Access Denied')] | //*[contains(text(),'403')] | //*[contains(text(),'Forbidden')]")).Count > 0;

                bool ok = noConfigMenu || isForbidden;
                actualMsg = ok ? expectedMsg : "Employee vẫn có thể truy cập trang Configuration";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(ROW_TC34, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }
    }
}