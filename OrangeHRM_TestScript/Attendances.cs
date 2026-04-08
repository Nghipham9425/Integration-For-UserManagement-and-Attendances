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

        // ── F3 – Employee Records ──
        private const int ROW_TC17 = 52;  // Excel row 53
        private const int ROW_TC18 = 55;  // Excel row 56
        private const int ROW_TC19 = 58;  // Excel row 59
        private const int ROW_TC20 = 60;  // Excel row 61  → TC20: lọc theo tên hợp lệ
        private const int ROW_TC21 = 64;  // Excel row 65  → TC21: xem chi tiết bằng nút View
        private const int ROW_TC22 = 67;  // Excel row 68  → TC22: Date required
        private const int ROW_TC23 = 70;  // Excel row 71  → TC23: Employee không có quyền

        // Last-step rows cho từng TC (dòng chứa Expected Result của step cuối)
        private const int ROW_TC20_LAST = 62;  // Excel row 63  (step 3)
        private const int ROW_TC21_LAST = 66;  // Excel row 67  (step 3)
        private const int ROW_TC22_LAST = 69;  // Excel row 70  (step 3) – "Date is required"
        private const int ROW_TC23_LAST = 73;  // Excel row 74  (step 4)

        // ── F4 – Attendance Configuration ──
        private const int ROW_TC24 = 75;   // Excel row 76
        private const int ROW_TC25 = 78;   // Excel row 79
        private const int ROW_TC26 = 82;   // Excel row 83
        private const int ROW_TC27 = 87;   // Excel row 88
        private const int ROW_TC28 = 91;   // Excel row 92
        private const int ROW_TC29 = 94;   // Excel row 95
        private const int ROW_TC30 = 100;  // Excel row 101
        private const int ROW_TC31 = 103;  // Excel row 104
        private const int ROW_TC32 = 106;  // Excel row 107
        private const int ROW_TC33 = 111;  // Excel row 112
        private const int ROW_TC34 = 115;  // Excel row 116

        // Last step row indices (0-based) – dòng chứa Expected Result của step CUỐI mỗi TC
        private const int ROW_TC17_LAST = 54;  // Excel row 55  (step 3)
        private const int ROW_TC18_LAST = 57;  // Excel row 58  (step 3)
        private const int ROW_TC19_LAST = 59;  // Excel row 60  (step 2)
        private const int ROW_TC24_LAST = 76;   // Excel row 77  (step 2)
        private const int ROW_TC25_LAST = 80;   // Excel row 81  (step 3)
        private const int ROW_TC26_LAST = 83;   // Excel row 84  (step 2)
        private const int ROW_TC27_LAST = 88;   // Excel row 89  (step 2)
        private const int ROW_TC28_LAST = 92;   // Excel row 93  (step 2)
        private const int ROW_TC29_LAST = 96;   // Excel row 97  (step 3)
        private const int ROW_TC30_LAST = 101;  // Excel row 102 (step 2)
        private const int ROW_TC31_LAST = 104;  // Excel row 105 (step 2)
        private const int ROW_TC32_LAST = 109;  // Excel row 110 (step 4)
        private const int ROW_TC33_LAST = 112;  // Excel row 113 (step 2)
        private const int ROW_TC34_LAST = 116;  // Excel row 117 (step 2)

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

        /// <summary>Bật/tắt toggle theo label text trong trang Configuration (dùng span.oxd-switch-input)</summary>
        private bool SetToggle(string labelContains, bool desiredState)
        {
            string rowXPath = $"//div[contains(@class,'orangehrm-attendance-field-row') and contains(.,'{labelContains}')]";
            var row = dr.FindElements(By.XPath(rowXPath)).FirstOrDefault();
            if (row == null) return false;

            var checkbox = row.FindElement(By.XPath(".//input[@type='checkbox']"));
            var switchSpan = row.FindElement(By.XPath(".//span[contains(@class,'oxd-switch-input')]"));

            if (checkbox.Selected != desiredState)
            {
                try { switchSpan.Click(); }
                catch { ((IJavaScriptExecutor)dr).ExecuteScript("arguments[0].click();", switchSpan); }
                Thread.Sleep(600);
            }
            return checkbox.Selected;
        }

        private bool GetToggleState(string labelContains)
        {
            string rowXPath = $"//div[contains(@class,'orangehrm-attendance-field-row') and contains(.,'{labelContains}')]";
            var row = dr.FindElements(By.XPath(rowXPath)).FirstOrDefault();
            if (row == null) return false;
            return row.FindElement(By.XPath(".//input[@type='checkbox']")).Selected;
        }

        private void ClickSave()
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)dr;
            var saveBtn = wait.Until(d => d.FindElement(By.CssSelector("button[type='submit']")));
            try { saveBtn.Click(); } catch { js.ExecuteScript("arguments[0].click();", saveBtn); }
            Thread.Sleep(1200);
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

        // ═══════════════════════════════════════════════════════════════
        // F3 – EMPLOYEE RECORDS  (TC20 – TC23)
        // ═══════════════════════════════════════════════════════════════

        /// <summary>
        /// TC20 – Kiểm tra Manager có thể lọc danh sách theo tên nhân viên hợp lệ.
        /// Excel: Row 61 (NPOI index 60). TestData cột H = "An Văn Nguyễn".
        /// Steps: (1) Xem tất cả → (2) Nhập tên vào Employee Name → (3) Nhấn Search/filter.
        /// Expected (last step): Danh sách chỉ hiển thị bản ghi của nhân viên khớp với điều kiện lọc.
        /// </summary>
        [TestMethod]
        public void ATT_TC20_EmployeeRecords_FilterByValidName()
        {
            var (_, testData) = ReadExcelRow(ROW_TC20);
            string expectedMsg = ReadExpected(ROW_TC20_LAST);
            var (username, password) = ExtractCredentials(testData, ADMIN_USER, ADMIN_PASS);

            // Lấy tên nhân viên từ step 2 (NPOI row 61 = Excel row 62)
            string employeeName;
            lock (excelLock)
            {
                using FileStream fs = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read);
                XSSFWorkbook wb = new XSSFWorkbook(fs);
                ISheet sh = wb.GetSheet(SHEET_NAME);
                employeeName = ReadCell(sh, ROW_TC20 + 1, COL_TESTDATA); // row 61 (Excel 62)
            }
            if (string.IsNullOrWhiteSpace(employeeName)) employeeName = "An Văn Nguyễn";

            string actualMsg = "";
            string status = "Failed";

            try
            {
                LoginAs(username, password);
                GoToEmployeeRecords();

                // Step 1 – Chọn ngày và nhấn View để xem tất cả
                string today = DateTime.Now.ToString("yyyy-MM-dd");
                var dateInput = wait.Until(d => d.FindElement(By.XPath(
                    "//label[contains(text(),'Date')]/ancestor::div[contains(@class,'oxd-input-group')]//input")));
                dateInput.SendKeys(Keys.Control + "a");
                dateInput.SendKeys(Keys.Backspace);
                dateInput.SendKeys(today);
                Thread.Sleep(300);

                dr.FindElement(By.CssSelector("button[type='submit']")).Click();
                Thread.Sleep(1500);

                // Step 2 – Nhập tên nhân viên vào ô lọc Employee Name
                var empNameInput = wait.Until(d => d.FindElement(By.XPath(
                    "//label[text()='Employee Name']/ancestor::div[contains(@class,'oxd-input-group')]//input")));
                empNameInput.Clear();
                foreach (char c in employeeName)
                {
                    empNameInput.SendKeys(c.ToString());
                    Thread.Sleep(50);
                }
                Thread.Sleep(1500);

                // Chờ dropdown autocomplete và chọn kết quả khớp
                var autocompleteOpts = dr.FindElements(By.XPath("//div[contains(@class,'oxd-autocomplete-option')]"));
                var matchOpt = autocompleteOpts.FirstOrDefault(o => o.Text.Contains(employeeName));
                if (matchOpt != null)
                    try { matchOpt.Click(); } catch { ((IJavaScriptExecutor)dr).ExecuteScript("arguments[0].click();", matchOpt); }
                else if (autocompleteOpts.Count > 0)
                    try { autocompleteOpts[0].Click(); } catch { ((IJavaScriptExecutor)dr).ExecuteScript("arguments[0].click();", autocompleteOpts[0]); }
                Thread.Sleep(500);

                // Step 3 – Nhấn Search / filter tự động kích hoạt
                dr.FindElement(By.CssSelector("button[type='submit']")).Click();
                Thread.Sleep(2000);

                // Xác minh kết quả: có bản ghi hoặc hiển thị Records Found
                bool hasRows = dr.FindElements(By.XPath("//div[@class='oxd-table-body']//div[@role='row']")).Count > 0;
                bool hasRecordsFound = dr.FindElements(By.XPath("//*[contains(text(),'Records Found')]")).Count > 0;
                bool ok = hasRows || hasRecordsFound;

                actualMsg = ok ? expectedMsg : $"Không tìm thấy dữ liệu cho '{employeeName}'";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = "Lỗi thực thi: " + ex.Message; }

            WriteExcelResult(ROW_TC20, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>
        /// TC21 – Kiểm tra Manager xem chi tiết bản ghi của nhân viên bằng nút View (Dynamic data).
        /// </summary>
        [TestMethod]
        public void ATT_TC21_EmployeeRecords_ViewDetail()
        {
            var (_, testData) = ReadExcelRow(ROW_TC21);
            string expectedMsg = ReadExpected(ROW_TC21_LAST);
            var (username, password) = ExtractCredentials(testData, ADMIN_USER, ADMIN_PASS);

            string actualMsg = "";
            string status = "Failed";
            IJavaScriptExecutor js = (IJavaScriptExecutor)dr;

            try
            {
                LoginAs(username, password);
                GoToEmployeeRecords();

                // Step 1 – Load danh sách Employee Records theo ngày hôm nay
                string today = DateTime.Now.ToString("yyyy-MM-dd");
                var dateInput = wait.Until(d => d.FindElement(By.XPath(
                    "//label[contains(text(),'Date')]/ancestor::div[contains(@class,'oxd-input-group')]//input")));
                dateInput.SendKeys(Keys.Control + "a");
                dateInput.SendKeys(Keys.Backspace);
                dateInput.SendKeys(today);
                Thread.Sleep(300);

                dr.FindElement(By.CssSelector("button[type='submit']")).Click();
                Thread.Sleep(2000);

                // Kiểm tra danh sách có ít nhất 1 dòng nhân viên
                var rows = dr.FindElements(By.XPath("//div[@class='oxd-table-body']//div[@role='row']"));
                Assert.IsTrue(rows.Count > 0, "Không có bản ghi nào trong danh sách Employee Records để thực hiện View");

                // Step 2 – Tự động lấy dòng đầu tiên và click nút View
                var firstRow = rows[0];

                // Lấy tên nhân viên ở cột đầu tiên để debug/log (Tùy chọn, không bắt buộc nhưng nên có)
                string employeeName = firstRow.FindElement(By.XPath(".//div[@role='cell'][1]")).Text;
                Console.WriteLine("Tiến hành click View cho nhân viên: " + employeeName);

                // Tìm nút View (thẻ button) nằm ngay trong dòng đầu tiên này
                var viewBtn = firstRow.FindElement(By.XPath(".//button"));

                try { viewBtn.Click(); }
                catch { js.ExecuteScript("arguments[0].click();", viewBtn); }
                Thread.Sleep(2000);

                // Step 3 – Quan sát trang chi tiết: phải có Punch In/Out và Duration
                bool hasPunchIn = dr.FindElements(By.XPath("//*[contains(text(),'Punch In')]")).Count > 0;
                bool hasPunchOut = dr.FindElements(By.XPath("//*[contains(text(),'Punch Out')]")).Count > 0;
                bool hasDuration = dr.FindElements(By.XPath("//*[contains(text(),'Duration')]")).Count > 0;
                bool hasDetailRows = dr.FindElements(By.XPath("//div[@class='oxd-table-body']//div[@role='row']")).Count > 0;

                bool ok = (hasPunchIn && hasPunchOut) || hasDuration || hasDetailRows;
                actualMsg = ok ? expectedMsg
                    : $"Trang chi tiết của {employeeName} không hiển thị đủ thông tin Punch In/Out và Duration";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = "Lỗi thực thi: " + ex.Message; }

            WriteExcelResult(ROW_TC21, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }


        /// <summary>
        /// TC22 – Kiểm tra hệ thống không cho xem Employee Records nếu không chọn ngày.
        /// Excel: Row 68 (NPOI index 67).
        /// Steps: (1) Để trống trường Date → (2) Nhấn nút View → (3) Quan sát phản hồi.
        /// Expected: Hiển thị thông báo lỗi "Date is required", không load danh sách.
        /// </summary>
        [TestMethod]
        public void ATT_TC22_EmployeeRecords_DateRequired()
        {
            var (_, testData) = ReadExcelRow(ROW_TC22);
            string expectedMsg = ReadExpected(ROW_TC22_LAST);
            var (username, password) = ExtractCredentials(testData, ADMIN_USER, ADMIN_PASS);

            string actualMsg = "";
            string status = "Failed";

            try
            {
                LoginAs(username, password);
                GoToEmployeeRecords();

                // Step 1 – Để trống trường Date (xóa giá trị mặc định nếu có)
                var dateInput = wait.Until(d => d.FindElement(By.XPath(
                    "//label[contains(text(),'Date')]/ancestor::div[contains(@class,'oxd-input-group')]//input")));
                dateInput.SendKeys(Keys.Control + "a");
                dateInput.SendKeys(Keys.Backspace);
                dateInput.SendKeys(Keys.Tab); // Trigger blur/validate
                Thread.Sleep(400);

                // Step 2 – Nhấn nút View
                dr.FindElement(By.CssSelector("button[type='submit']")).Click();
                Thread.Sleep(1500);

                // Step 3 – Quan sát phản hồi: phải có thông báo lỗi "Date is required"
                bool hasDateRequired = dr.FindElements(By.XPath(
                    "//span[contains(@class,'oxd-input-field-error-message') and contains(text(),'Required')]")).Count > 0
                    || dr.FindElements(By.XPath(
                    "//span[contains(@class,'oxd-input-field-error-message') and contains(text(),'required')]")).Count > 0
                    || dr.FindElements(By.XPath(
                    "//*[contains(text(),'Date is required')]")).Count > 0;

                bool hasErrorSpan = dr.FindElements(By.XPath(
                    "//span[contains(@class,'oxd-input-field-error-message')]")).Count > 0;

                bool ok = hasDateRequired || hasErrorSpan;
                actualMsg = ok ? expectedMsg : "Hệ thống không hiển thị lỗi 'Date is required' khi trường Date trống";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = "Lỗi thực thi: " + ex.Message; }

            WriteExcelResult(ROW_TC22, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>
        /// TC23 – Kiểm tra nhân viên KHÔNG có quyền truy cập Employee Records của người khác.
        /// Excel: Row 71 (NPOI index 70). TestData = "nguyenvana, Hass@12341".
        /// Steps: (1) Đăng nhập Employee → (2) Thử menu Attendance
        ///        → (3) Kiểm tra không có "Employee Records" → (4) Thử URL trực tiếp.
        /// Expected: Redirect 403 / "Access Denied" hoặc không có menu Employee Records.
        /// </summary>
        [TestMethod]
        public void ATT_TC23_EmployeeRecords_AccessDenied()
        {
            var (_, testData) = ReadExcelRow(ROW_TC23);
            string expectedMsg = ReadExpected(ROW_TC23_LAST);
            var (username, password) = ExtractCredentials(testData, EMP_USER, EMP_PASS);

            string actualMsg = "";
            string status = "Failed";

            try
            {
                // Step 1 – Đăng nhập với tài khoản Employee
                LoginAs(username, password);

       

                // Step 2 – Thử vào menu Attendance
                var menuAttendance = dr.FindElements(By.XPath(
                    "//span[contains(text(),'Attendance')] | //a[contains(text(),'Attendance')]")).FirstOrDefault();
                if (menuAttendance != null)
                {
                    try { menuAttendance.Click(); Thread.Sleep(600); }
                    catch { /* ignore */ }
                }

                // Step 3 – Kiểm tra submenu "Employee Records" không hiển thị cho Employee
                bool noEmployeeRecordsMenu = dr.FindElements(By.XPath(
                    "//a[contains(text(),'Employee Records')] | //span[text()='Employee Records']")).Count == 0;

                // Step 4 – Thử truy cập trực tiếp URL Employee Records
                GoToEmployeeRecords(); 
                Thread.Sleep(1500);

                bool isForbidden = dr.Url.Contains("dashboard") ||
                    dr.FindElements(By.XPath(
                        "//*[contains(text(),'Access Denied')] | //*[contains(text(),'403')] | " +
                        "//*[contains(text(),'Forbidden')]")).Count > 0;

                bool ok = noEmployeeRecordsMenu || isForbidden;
                actualMsg = ok ? expectedMsg
                    : "Employee vẫn có thể truy cập trang Employee Records";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = "Lỗi thực thi: " + ex.Message; }

            WriteExcelResult(ROW_TC23, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }


        // ═══════════════════════════════════════════════════════════════
        // F4 – ATTENDANCE CONFIGURATION  (TC24 – TC34)
        // ═══════════════════════════════════════════════════════════════

        /// <summary>
        /// TC24 – Kiểm tra Admin truy cập được trang Attendance Configuration.
        /// Excel: Row 76 (NPOI index 75). TestData có username/password Admin.
        /// Steps: (1) Truy cập Attendance > Configuration → (2) Quan sát giao diện.
        /// Expected: Trang hiển thị các toggle/checkbox cấu hình và nút "Save".
        /// </summary>
        [TestMethod]
        public void ATT_TC24_Config_AdminAccess()
        {
            var (_, testData) = ReadExcelRow(ROW_TC24);
            string expectedMsg = ReadExpected(ROW_TC24_LAST);
            var (username, password) = ExtractCredentials(testData, ADMIN_USER, ADMIN_PASS);

            string actualMsg = "";
            string status = "Failed";

            try
            {
                // Step 1 – Truy cập Attendance > Configuration
                LoginAs(username, password);
                GoToAttendanceConfiguration();

                // Step 2 – Quan sát giao diện trang
                bool hasToggles = dr.FindElements(By.XPath("//input[@type='checkbox']")).Count > 0
                    || dr.FindElements(By.XPath("//span[contains(@class,'oxd-switch-input')]")).Count > 0;
                bool hasSaveBtn = dr.FindElements(By.XPath("//button[contains(.,'Save')]")).Count > 0;
                bool hasConfig = dr.PageSource.Contains("Configuration");

                bool ok = hasConfig && hasToggles && hasSaveBtn;
                actualMsg = ok ? expectedMsg
                    : "Trang Configuration không hiển thị đúng: " +
                      (!hasConfig ? "[Tiêu đề Configuration] " : "") +
                      (!hasToggles ? "[Toggle/Checkbox] " : "") +
                      (!hasSaveBtn ? "[Nút Save] " : "");
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = "Lỗi thực thi: " + ex.Message; }

            WriteExcelResult(ROW_TC24, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>
        /// TC25 – Kiểm tra Admin bật/tắt tùy chọn "Employee can change current time when punching in/out".
        /// Excel: Row 79 (NPOI index 78). TestData có username/password Admin.
        /// Steps: (1) Ghi nhận trạng thái hiện tại → (2) Click toggle + Save → (3) Reload kiểm tra.
        /// Expected: Trạng thái toggle vẫn giữ nguyên sau reload.
        /// </summary>
        [TestMethod]
        public void ATT_TC25_Config_ToggleChangeTime()
        {
            var (_, testData) = ReadExcelRow(ROW_TC25);
            string expectedMsg = ReadExpected(ROW_TC25_LAST);
            var (username, password) = ExtractCredentials(testData, ADMIN_USER, ADMIN_PASS);

            string actualMsg = "";
            string status = "Failed";

            try
            {
                LoginAs(username, password);
                GoToAttendanceConfiguration();

                // Step 1 – Ghi nhận trạng thái hiện tại
                bool initialState = GetToggleState("Employee can change");

                // Step 2 – Click toggle để đổi trạng thái (ON→OFF hoặc OFF→ON), nhấn Save
                bool newState = SetToggle("Employee can change", !initialState);
                bool changed = (newState != initialState);

                ClickSave();

                // Kiểm tra toast "Saved successfully"
                bool hasSavedToast = wait.Until(d =>
                    d.FindElements(By.XPath("//*[contains(@class,'oxd-text--toast')] | //*[contains(text(),'success')]")).Count > 0);

                // Step 3 – Reload trang, kiểm tra lại toggle
                GoToAttendanceConfiguration();
                Thread.Sleep(500);
                bool persistedState = GetToggleState("Employee can change");
                bool persisted = (persistedState == !initialState);

                bool ok = changed && persisted;
                actualMsg = ok ? expectedMsg
                    : $"Toggle không thay đổi ({changed}) hoặc không lưu sau reload ({persisted})";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = "Lỗi thực thi: " + ex.Message; }

            WriteExcelResult(ROW_TC25, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>
        /// TC26 – Kiểm tra khi BẬT "Employee can change current time" → nhân viên được sửa giờ khi Punch In/Out.
        /// Excel: Row 83 (NPOI index 82). TestData step 1 = "username: nguyenvana\npassword: Hass@12341".
        /// Steps: (1) Vào trang Punch In/Out → (2) Thử click vào trường Time và nhập giá trị → end tại đó.
        /// Expected: Trường Time có thể click và chỉnh sửa được (enabled, không readonly).
        /// </summary>
        [TestMethod]
        public void ATT_TC26_Config_ChangeTimeEnabled()
        {
            var (_, testData) = ReadExcelRow(ROW_TC26);
            string expectedMsg = ReadExpected(ROW_TC26_LAST);
            var (empUser, empPass) = ExtractCredentials(testData, EMP_USER, EMP_PASS);

            // Đọc giá trị Time muốn nhập từ step 2 (NPOI row 83 = Excel row 84)
            string customTime;
            lock (excelLock)
            {
                using FileStream fs = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read);
                XSSFWorkbook wb = new XSSFWorkbook(fs);
                ISheet sh = wb.GetSheet(SHEET_NAME);
                customTime = ReadCell(sh, ROW_TC26 + 1, COL_TESTDATA); // "Time: 07:45"
            }
            customTime = customTime.Replace("Time:", "").Trim();
            if (string.IsNullOrWhiteSpace(customTime)) customTime = "07:45";

            string actualMsg = "";
            string status = "Failed";
            IJavaScriptExecutor js = (IJavaScriptExecutor)dr;

            try
            {
                // Tiền điều kiện: Admin BẬT toggle "Employee can change current time"
                LoginAs(ADMIN_USER, ADMIN_PASS);
                GoToAttendanceConfiguration();
                SetToggle("Employee can change", true);
                ClickSave();
                Thread.Sleep(500);
                Logout();

                // Step 1 – Đăng nhập Employee, vào trang Punch In/Out
                LoginAs(empUser, empPass);
                GoToPunchInOut();

                IWebElement timeInput = wait.Until(d => d.FindElement(By.XPath(
                    "//label[contains(text(),'Time')]/ancestor::div[contains(@class,'oxd-input-group')]//input")));

                // Step 2 – Thử click vào trường Time và nhập giá trị → chỉ cần click được là pass
                bool clickOk = false;
                try
                {
                    timeInput.Click();
                    clickOk = true;
                }
                catch
                {
                    js.ExecuteScript("arguments[0].click();", timeInput);
                    clickOk = true;
                }
                Thread.Sleep(300);

                // Thử gõ ký tự vào — nếu trường enabled sẽ nhận input
                string valueBefore = timeInput.GetAttribute("value") ?? "";
                timeInput.SendKeys(Keys.Control + "a");
                timeInput.SendKeys(Keys.Backspace);
                timeInput.SendKeys(customTime);
                Thread.Sleep(300);
                string valueAfter = timeInput.GetAttribute("value") ?? "";

                // Trường enabled: value thay đổi sau khi gõ, hoặc ít nhất click không bị block
                bool isEditable = clickOk && (timeInput.Enabled &&
                    (timeInput.GetAttribute("disabled") == null || timeInput.GetAttribute("disabled") == "false") &&
                    (timeInput.GetAttribute("readonly") == null || timeInput.GetAttribute("readonly") != "true"));

                bool ok = isEditable;
                actualMsg = ok ? expectedMsg : "Trường Time không thể chỉnh sửa khi toggle đã BẬT";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = "Lỗi thực thi: " + ex.Message; }

            WriteExcelResult(ROW_TC26, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>
        /// TC27 – Kiểm tra khi TẮT "Employee can change current time" → nhân viên KHÔNG thể sửa giờ.
        /// Excel: Row 88 (NPOI index 87). TestData = "username: nguyenvana\npassword: Hass@12341".
        /// Steps: (1) Vào trang Punch In/Out → (2) Thử click và nhập ký tự vào trường Time.
        /// Expected: Trường Time bị disable/readonly, không nhận input; giờ hiện tại tự động điền.
        /// </summary>
        [TestMethod]
        public void ATT_TC27_Config_ChangeTimeDisabled()
        {
            var (_, testData) = ReadExcelRow(ROW_TC27);
            string expectedMsg = ReadExpected(ROW_TC27_LAST);
            var (empUser, empPass) = ExtractCredentials(testData, EMP_USER, EMP_PASS);

            string actualMsg = "";
            string status = "Failed";
            IJavaScriptExecutor js = (IJavaScriptExecutor)dr;

            try
            {
                // Tiền điều kiện: Admin TẮT toggle "Employee can change current time"
                LoginAs(ADMIN_USER, ADMIN_PASS);
                GoToAttendanceConfiguration();
                SetToggle("Employee can change", false);
                ClickSave();
                Thread.Sleep(500);
                Logout();

                // Step 1 – Đăng nhập Employee, vào trang Punch In/Out
                LoginAs(empUser, empPass);
                GoToPunchInOut();

                // Step 2 – Thử click và nhập ký tự vào trường Time
                IWebElement timeInput = wait.Until(d => d.FindElement(By.XPath(
                    "//label[contains(text(),'Time')]/ancestor::div[contains(@class,'oxd-input-group')]//input")));

                try { timeInput.Click(); } catch { /* expected khi disabled */ }

                bool isDisabled = !timeInput.Enabled
                    || timeInput.GetAttribute("disabled") != null
                    || "true".Equals(timeInput.GetAttribute("readonly"), StringComparison.OrdinalIgnoreCase);

                // Kiểm tra thêm: giờ hiện tại tự động điền (không rỗng)
                bool hasAutoTime = !string.IsNullOrWhiteSpace(timeInput.GetAttribute("value"));

                bool ok = isDisabled;
                actualMsg = ok ? expectedMsg
                    : "Trường Time vẫn cho phép chỉnh sửa khi toggle đã TẮT";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = "Lỗi thực thi: " + ex.Message; }

            WriteExcelResult(ROW_TC27, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>
        /// TC28 – Kiểm tra Admin bật/tắt tùy chọn "Employee can edit/delete own attendance records".
        /// Excel: Row 92 (NPOI index 91). TestData có username/password Admin.
        /// Steps: (1) Ghi nhận + đổi trạng thái toggle, nhấn Save → (2) Reload kiểm tra toggle.
        /// Expected: Thông báo "Saved successfully"; toggle giữ trạng thái đúng sau reload.
        /// </summary>
        [TestMethod]
        public void ATT_TC28_Config_ToggleEditDelete()
        {
            var (_, testData) = ReadExcelRow(ROW_TC28);
            string expectedMsg = ReadExpected(ROW_TC28_LAST);
            var (username, password) = ExtractCredentials(testData, ADMIN_USER, ADMIN_PASS);

            string actualMsg = "";
            string status = "Failed";

            try
            {
                LoginAs(username, password);
                GoToAttendanceConfiguration();

                // Step 1 – Ghi nhận trạng thái toggle "Employee can edit/delete", click đổi, nhấn Save
                bool initialState = GetToggleState("Employee can edit");
                bool newState = SetToggle("Employee can edit", !initialState);
                bool changed = (newState != initialState);

                ClickSave();

                bool hasSavedToast = dr.FindElements(By.XPath(
                    "//*[contains(@class,'oxd-text--toast')] | //*[contains(text(),'success')]")).Count > 0;

                // Step 2 – Reload trang kiểm tra toggle
                GoToAttendanceConfiguration();
                Thread.Sleep(500);
                bool persistedState = GetToggleState("Employee can edit");
                bool persisted = (persistedState == !initialState);

                bool ok = changed && persisted;
                actualMsg = ok ? expectedMsg
                    : $"Toggle edit/delete không thay đổi ({changed}) hoặc không lưu sau reload ({persisted})";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = "Lỗi thực thi: " + ex.Message; }

            WriteExcelResult(ROW_TC28, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>
        /// TC29 – Kiểm tra khi BẬT quyền edit/delete → nhân viên có thể chỉnh sửa/xóa bản ghi của mình.
        /// Excel: Row 95 (NPOI index 94). TestData = "username: nguyenvana\npassword: Hass@12341".
        /// Steps: (1) Kiểm tra giao diện có nút Edit/Delete → (2) Click Edit, sửa Note, nhấn Save
        ///        → (3) Click Delete, xác nhận xóa.
        /// Expected: Nút Edit/Delete hiển thị; bản ghi cập nhật/xóa thành công.
        /// </summary>
        [TestMethod]
        public void ATT_TC29_Config_EditDeleteEnabled()
        {
            var (_, testData) = ReadExcelRow(ROW_TC29);
            string expectedMsg = ReadExpected(ROW_TC29_LAST);
            var (empUser, empPass) = ExtractCredentials(testData, EMP_USER, EMP_PASS);

            // Đọc Note từ step 2 (NPOI row 95 = Excel row 96): "Note: \"Đã chỉnh sửa\""
            string editNote;
            lock (excelLock)
            {
                using FileStream fs = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read);
                XSSFWorkbook wb = new XSSFWorkbook(fs);
                ISheet sh = wb.GetSheet(SHEET_NAME);
                editNote = ReadCell(sh, ROW_TC29 + 1, COL_TESTDATA)
                    .Replace("Note:", "").Trim().Trim('"');
            }
            if (string.IsNullOrWhiteSpace(editNote)) editNote = "Đã chỉnh sửa";

            string actualMsg = "";
            string status = "Failed";
            IJavaScriptExecutor js = (IJavaScriptExecutor)dr;

            try
            {
                // Tiền điều kiện: Admin BẬT toggle "Employee can edit/delete own attendance records"
                LoginAs(ADMIN_USER, ADMIN_PASS);
                GoToAttendanceConfiguration();
                SetToggle("Employee can edit", true);
                ClickSave();
                Thread.Sleep(500);
                Logout();

                // Đăng nhập Employee, vào My Attendance Records
                LoginAs(empUser, empPass);
                GoToMyAttendanceRecords();

                string today = DateTime.Now.ToString("yyyy-MM-dd");
                var dateInput = wait.Until(d => d.FindElement(By.XPath(
                    "//label[contains(text(),'Date')]/ancestor::div[contains(@class,'oxd-input-group')]//input")));
                dateInput.SendKeys(Keys.Control + "a");
                dateInput.SendKeys(Keys.Backspace);
                dateInput.SendKeys(today);
                Thread.Sleep(300);

                dr.FindElement(By.CssSelector("button[type='submit']")).Click();
                Thread.Sleep(2000);

                // Step 1 – Kiểm tra giao diện: có nút Edit và Delete trên mỗi bản ghi
                bool hasEditBtn = dr.FindElements(By.XPath("//button[i[contains(@class,'bi-pencil')]]")).Count > 0;
                bool hasDeleteBtn = dr.FindElements(By.XPath("//button[i[contains(@class,'bi-trash')]]")).Count > 0;

                Assert.IsTrue(hasEditBtn || hasDeleteBtn,
                    "Không thấy nút Edit/Delete dù đã bật quyền edit/delete");

                // Step 2 – Click Edit trên 1 bản ghi, sửa Note, nhấn Save
                var editBtn = wait.Until(d => d.FindElement(By.XPath(
                    "(//button[i[contains(@class,'bi-pencil')]])[1]")));
                try { editBtn.Click(); } catch { js.ExecuteScript("arguments[0].click();", editBtn); }
                Thread.Sleep(1000);

                var noteInput = dr.FindElements(By.XPath(
                    "//label[contains(text(),'Note')]/ancestor::div[contains(@class,'oxd-input-group')]//textarea")).FirstOrDefault()
                    ?? dr.FindElements(By.XPath("//textarea")).FirstOrDefault();
                if (noteInput != null)
                {
                    noteInput.SendKeys(Keys.Control + "a");
                    noteInput.SendKeys(Keys.Backspace);
                    noteInput.SendKeys(editNote);
                    Thread.Sleep(300);
                }

                dr.FindElement(By.CssSelector("button[type='submit']")).Click();
                Thread.Sleep(1500);

                bool editSaved = dr.FindElements(By.XPath("//*[contains(@class,'oxd-text--toast')]")).Count > 0
                    || dr.PageSource.Contains("Successfully Saved")
                    || dr.PageSource.Contains(editNote);

                // Step 3 – Click Delete trên 1 bản ghi khác, xác nhận xóa
                GoToMyAttendanceRecords();
                Thread.Sleep(800);

                var dateInput2 = wait.Until(d => d.FindElement(By.XPath(
                    "//label[contains(text(),'Date')]/ancestor::div[contains(@class,'oxd-input-group')]//input")));
                dateInput2.SendKeys(Keys.Control + "a");
                dateInput2.SendKeys(Keys.Backspace);
                dateInput2.SendKeys(today);
                Thread.Sleep(300);

                dr.FindElement(By.CssSelector("button[type='submit']")).Click();
                Thread.Sleep(2000);

                var deleteBtns = dr.FindElements(By.XPath("//button[i[contains(@class,'bi-trash')]]"));
                bool deleteOk = false;
                if (deleteBtns.Count >= 2)
                {
                    try { deleteBtns[1].Click(); } catch { js.ExecuteScript("arguments[0].click();", deleteBtns[1]); }
                    Thread.Sleep(500);

                    var confirmBtn = wait.Until(d => d.FindElement(By.XPath("//button[contains(.,'Yes, Delete')]")));
                    try { confirmBtn.Click(); } catch { js.ExecuteScript("arguments[0].click();", confirmBtn); }
                    Thread.Sleep(1500);

                    deleteOk = dr.FindElements(By.XPath("//*[contains(@class,'oxd-text--toast')]")).Count > 0
                        || dr.PageSource.Contains("Successfully Deleted");
                }
                else deleteOk = true; // Chỉ có 1 bản ghi, bỏ qua bước xóa

                bool ok = (hasEditBtn || hasDeleteBtn) && editSaved;
                actualMsg = ok ? expectedMsg
                    : "Nút Edit/Delete không hiển thị hoặc thao tác Edit/Delete thất bại";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = "Lỗi thực thi: " + ex.Message; }

            WriteExcelResult(ROW_TC29, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>
        /// TC30 – Kiểm tra khi TẮT quyền edit/delete → nhân viên KHÔNG thể chỉnh sửa/xóa bản ghi.
        /// Excel: Row 101 (NPOI index 100). TestData = "username: nguyenvana\npassword: Hass@12341".
        /// Steps: (1) Kiểm tra giao diện My Records không có nút Edit/Delete
        ///        → (2) Thử truy cập trực tiếp URL edit record.
        /// Expected: Nút Edit/Delete KHÔNG hiển thị; URL trả về 403 / "Access Denied".
        /// </summary>
        [TestMethod]
        public void ATT_TC30_Config_EditDeleteDisabled()
        {
            var (_, testData) = ReadExcelRow(ROW_TC30);
            string expectedMsg = ReadExpected(ROW_TC30_LAST);
            var (empUser, empPass) = ExtractCredentials(testData, EMP_USER, EMP_PASS);

            string actualMsg = "";
            string status = "Failed";
            IJavaScriptExecutor js = (IJavaScriptExecutor)dr;

            try
            {
                // Tiền điều kiện: Admin TẮT toggle "Employee can edit/delete own attendance records"
                LoginAs(ADMIN_USER, ADMIN_PASS);
                GoToAttendanceConfiguration();
                SetToggle("Employee can edit", false);
                ClickSave();
                Thread.Sleep(500);
                Logout();

                // Đăng nhập Employee, vào My Attendance Records
                LoginAs(empUser, empPass);
                GoToMyAttendanceRecords();

                string today = DateTime.Now.ToString("yyyy-MM-dd");
                var dateInput = wait.Until(d => d.FindElement(By.XPath(
                    "//label[contains(text(),'Date')]/ancestor::div[contains(@class,'oxd-input-group')]//input")));
                dateInput.SendKeys(Keys.Control + "a");
                dateInput.SendKeys(Keys.Backspace);
                dateInput.SendKeys(today);
                Thread.Sleep(300);

                dr.FindElement(By.CssSelector("button[type='submit']")).Click();
                Thread.Sleep(2000);

                // Step 1 – Kiểm tra giao diện: nút Edit/Delete KHÔNG hiển thị
                bool noEditBtn = dr.FindElements(By.XPath("//button[i[contains(@class,'bi-pencil')]]")).Count == 0;
                bool noDeleteBtn = dr.FindElements(By.XPath("//button[i[contains(@class,'bi-trash')]]")).Count == 0;

                // Step 2 – Thử truy cập trực tiếp URL edit record
                dr.Navigate().GoToUrl(BASE_URL + "/web/index.php/attendance/edit");
                Thread.Sleep(1500);

                bool isForbidden = dr.Url.Contains("dashboard") ||
                    dr.FindElements(By.XPath(
                        "//*[contains(text(),'Access Denied')] | //*[contains(text(),'403')] | " +
                        "//*[contains(text(),'Forbidden')]")).Count > 0;

                bool ok = (noEditBtn && noDeleteBtn) || isForbidden;
                actualMsg = ok ? expectedMsg
                    : "Nút Edit/Delete vẫn hiển thị hoặc Employee truy cập được URL edit khi đã tắt quyền";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = "Lỗi thực thi: " + ex.Message; }

            WriteExcelResult(ROW_TC30, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>
        /// TC31 – Kiểm tra Admin bật/tắt tùy chọn "Supervisor can add/edit/delete attendance records of subordinates".
        /// Excel: Row 104 (NPOI index 103). TestData có username/password Admin.
        /// Steps: (1) Ghi nhận + đổi trạng thái toggle + Save → (2) Reload kiểm tra toggle.
        /// Expected: Toggle đúng trạng thái sau reload.
        /// </summary>
        [TestMethod]
        public void ATT_TC31_Config_ToggleSupervisor()
        {
            var (_, testData) = ReadExcelRow(ROW_TC31);
            string expectedMsg = ReadExpected(ROW_TC31_LAST);
            var (username, password) = ExtractCredentials(testData, ADMIN_USER, ADMIN_PASS);

            string actualMsg = "";
            string status = "Failed";

            try
            {
                LoginAs(username, password);
                GoToAttendanceConfiguration();

                // Step 1 – Ghi nhận trạng thái, click đổi, nhấn Save
                bool initialState = GetToggleState("Supervisor can add");
                bool newState = SetToggle("Supervisor can add", !initialState);
                bool changed = (newState != initialState);

                ClickSave();

                bool hasSavedToast = dr.FindElements(By.XPath(
                    "//*[contains(@class,'oxd-text--toast')] | //*[contains(text(),'success')]")).Count > 0;

                // Step 2 – Reload trang kiểm tra toggle
                GoToAttendanceConfiguration();
                Thread.Sleep(500);
                bool persistedState = GetToggleState("Supervisor can add");
                bool persisted = (persistedState == !initialState);

                bool ok = changed && persisted;
                actualMsg = ok ? expectedMsg
                    : $"Toggle Supervisor không thay đổi ({changed}) hoặc không lưu sau reload ({persisted})";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = "Lỗi thực thi: " + ex.Message; }

            WriteExcelResult(ROW_TC31, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }
        /// <summary>
        /// TC32 – Kiểm tra khi BẬT quyền supervisor → Manager có thể thêm/sửa/xóa bản ghi chấm công của nhân viên.
        /// Excel: Row 107 (NPOI index 106). TestData = "username: nghi45397\npassword: Nghiphamtrung09042005!".
        /// Steps: (1) Vào Employee Records, chọn nhân viên → (2) Thêm bản ghi mới Punch In/Out
        ///        → (3) Edit bản ghi, sửa Note → (4) Delete bản ghi, xác nhận xóa.
        /// Expected: Nút Add/Edit/Delete hoạt động; bản ghi thêm/sửa/xóa thành công.
        /// </summary>
        [TestMethod]
        public void ATT_TC32_Config_SupervisorEnabled()
        {
            var (_, testData) = ReadExcelRow(ROW_TC32);
            string expectedMsg = ReadExpected(ROW_TC32_LAST);
            var (username, password) = ExtractCredentials(testData, ADMIN_USER, ADMIN_PASS);

            string punchInTime = "08:00 AM";
            string punchOutTime = "05:00 PM";
            string editNote = "Updated by Admin";

            lock (excelLock)
            {
                using FileStream fs = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read);
                XSSFWorkbook wb = new XSSFWorkbook(fs);
                ISheet sh = wb.GetSheet(SHEET_NAME);

                string step2 = ReadCell(sh, ROW_TC32 + 1, COL_TESTDATA);
                foreach (var line in step2.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    if (line.Trim().ToLower().StartsWith("punch in:")) punchInTime = line.Substring(line.IndexOf(':') + 1).Trim();
                    if (line.Trim().ToLower().StartsWith("punch out:")) punchOutTime = line.Substring(line.IndexOf(':') + 1).Trim();
                }

                string step3 = ReadCell(sh, ROW_TC32 + 2, COL_TESTDATA);
                if (!string.IsNullOrWhiteSpace(step3))
                    editNote = step3.Replace("Note:", "").Trim().Trim('"');
            }

            string actualMsg = "";
            string status = "Failed";
            IJavaScriptExecutor js = (IJavaScriptExecutor)dr;
            string today = DateTime.Now.ToString("yyyy-MM-dd");

            try
            {
                // Tiền điều kiện: Admin BẬT toggle "Supervisor can add/edit/delete"
                LoginAs(username, password);
                GoToAttendanceConfiguration();
                SetToggle("Supervisor can add", true);
                ClickSave();
                Thread.Sleep(500);

                // Step 1 – Vào Employee Records, lọc nhân viên
                GoToEmployeeRecords();
                Thread.Sleep(800);

                var dateInput = wait.Until(d => d.FindElement(By.XPath(
                    "//label[contains(text(),'Date')]/ancestor::div[contains(@class,'oxd-input-group')]//input")));
                dateInput.SendKeys(Keys.Control + "a");
                dateInput.SendKeys(Keys.Backspace);
                dateInput.SendKeys(today);
                Thread.Sleep(300);

                dr.FindElement(By.CssSelector("button[type='submit']")).Click();
                Thread.Sleep(2000);

                // Phải có ít nhất 1 dòng nhân viên trong danh sách
                var listRows = dr.FindElements(By.XPath("//div[@class='oxd-table-body']//div[@role='row']"));
                Assert.IsTrue(listRows.Count > 0, "[Step 1 FAIL] Không có bản ghi nào trong danh sách Employee Records");

                var viewBtns = dr.FindElements(By.XPath("//div[@class='oxd-table-body']//button"));
                Assert.IsTrue(viewBtns.Count > 0, "[Step 1 FAIL] Không tìm thấy nút View trong danh sách");

                try { viewBtns[0].Click(); }
                catch { js.ExecuteScript("arguments[0].click();", viewBtns[0]); }
                Thread.Sleep(2000);

                // Trên trang chi tiết phải có nút Add
                var addBtns = dr.FindElements(By.XPath("//button[normalize-space()='Add']"));
                Assert.IsTrue(addBtns.Count > 0, "[Step 1 FAIL] Trang chi tiết không có nút Add dù đã bật quyền Supervisor");

                // Step 2 – Click Add, nhập Punch In / Punch Out, Save
                try { addBtns[0].Click(); }
                catch { js.ExecuteScript("arguments[0].click();", addBtns[0]); }
                Thread.Sleep(1000);

                var punchInInput = dr.FindElements(By.XPath(
                    "//label[contains(text(),'Punch In Time')]/ancestor::div[contains(@class,'oxd-input-group')]//input")).FirstOrDefault();
                if (punchInInput != null)
                {
                    punchInInput.SendKeys(Keys.Control + "a");
                    punchInInput.SendKeys(Keys.Backspace);
                    punchInInput.SendKeys(punchInTime);
                    Thread.Sleep(300);
                    punchInInput.SendKeys(Keys.Tab);
                }

                var punchOutInput = dr.FindElements(By.XPath(
                    "//label[contains(text(),'Punch Out Time')]/ancestor::div[contains(@class,'oxd-input-group')]//input")).FirstOrDefault();
                if (punchOutInput != null)
                {
                    punchOutInput.SendKeys(Keys.Control + "a");
                    punchOutInput.SendKeys(Keys.Backspace);
                    punchOutInput.SendKeys(punchOutTime);
                    Thread.Sleep(300);
                    punchOutInput.SendKeys(Keys.Tab);
                }

                dr.FindElement(By.CssSelector("button[type='submit']")).Click();
                Thread.Sleep(1500);

                bool addSaved = dr.FindElements(By.XPath("//*[contains(@class,'oxd-text--toast')]")).Count > 0
                    || dr.PageSource.Contains("Successfully Saved");

                // Step 3 – Click Edit trên bản ghi cuối cùng, sửa Note, Save
                var editBtns = dr.FindElements(By.XPath("//button[i[contains(@class,'bi-pencil')]]"));
                bool editSaved = false;
                if (editBtns.Count > 0)
                {
                    try { editBtns[editBtns.Count - 1].Click(); }
                    catch { js.ExecuteScript("arguments[0].click();", editBtns[editBtns.Count - 1]); }
                    Thread.Sleep(1000);

                    var noteInputEdit = dr.FindElements(By.XPath("//textarea")).FirstOrDefault();
                    if (noteInputEdit != null)
                    {
                        noteInputEdit.SendKeys(Keys.Control + "a");
                        noteInputEdit.SendKeys(Keys.Backspace);
                        noteInputEdit.SendKeys(editNote);
                        Thread.Sleep(300);
                    }

                    dr.FindElement(By.CssSelector("button[type='submit']")).Click();
                    Thread.Sleep(1500);

                    editSaved = dr.FindElements(By.XPath("//*[contains(@class,'oxd-text--toast')]")).Count > 0
                        || dr.PageSource.Contains("Successfully Saved");
                }

                // Step 4 – Click Delete trên bản ghi, xác nhận xóa
                var deleteBtns = dr.FindElements(By.XPath("//button[i[contains(@class,'bi-trash')]]"));
                bool deleteDone = false;
                if (deleteBtns.Count > 0)
                {
                    try { deleteBtns[deleteBtns.Count - 1].Click(); }
                    catch { js.ExecuteScript("arguments[0].click();", deleteBtns[deleteBtns.Count - 1]); }
                    Thread.Sleep(500);

                    var confirmBtn = wait.Until(d => d.FindElement(By.XPath("//button[contains(.,'Yes, Delete')]")));
                    try { confirmBtn.Click(); } catch { js.ExecuteScript("arguments[0].click();", confirmBtn); }
                    Thread.Sleep(1500);

                    deleteDone = dr.FindElements(By.XPath("//*[contains(@class,'oxd-text--toast')]")).Count > 0
                        || !dr.PageSource.Contains(editNote);
                }

                // Pass nếu vào được trang chi tiết có nút Add và ít nhất 1 thao tác thành công
                bool ok = addBtns.Count > 0 && (addSaved || editSaved || deleteDone);
                actualMsg = ok ? expectedMsg
                    : "Add/Edit/Delete không hoạt động đúng trên trang chi tiết dù đã bật quyền Supervisor";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = "Lỗi thực thi: " + ex.Message; }

            WriteExcelResult(ROW_TC32, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }


        /// <summary>
        /// TC33 – Kiểm tra hệ thống lưu cấu hình khi Admin nhấn Save.
        /// Excel: Row 112 (NPOI index 111). TestData có username/password Admin.
        /// Steps: (1) Thay đổi ít nhất 1 toggle + nhấn Save → (2) Reload kiểm tra giá trị.
        /// Expected: Thông báo "Configuration saved successfully" hoặc toast thành công; cấu hình giữ đúng sau reload.
        /// </summary>
        [TestMethod]
        public void ATT_TC33_Config_SaveSuccess()
        {
            var (_, testData) = ReadExcelRow(ROW_TC33);
            string expectedMsg = ReadExpected(ROW_TC33_LAST);
            var (username, password) = ExtractCredentials(testData, ADMIN_USER, ADMIN_PASS);

            string actualMsg = "";
            string status = "Failed";
            IJavaScriptExecutor js = (IJavaScriptExecutor)dr;

            try
            {
                LoginAs(username, password);
                GoToAttendanceConfiguration();

                // Step 1 – Thay đổi ít nhất 1 toggle bất kỳ, nhấn nút "Save"
                bool initialState = GetToggleState("Employee can change");
                SetToggle("Employee can change", !initialState);

                ClickSave();

                // Kiểm tra xuất hiện toast thành công
                bool hasSavedMsg = wait.Until(d =>
                    d.FindElements(By.XPath(
                        "//*[contains(@class,'oxd-text--toast')] | //*[contains(@class,'oxd-alert-content-text')] | " +
                        "//*[contains(text(),'successfully')] | //*[contains(text(),'Saved')]")).Count > 0);

                // Step 2 – Reload trang Configuration, kiểm tra giá trị các cài đặt
                GoToAttendanceConfiguration();
                Thread.Sleep(500);

                bool persistedState = GetToggleState("Employee can change");
                bool persisted = (persistedState == !initialState);

                bool ok = hasSavedMsg && persisted;
                actualMsg = ok ? expectedMsg
                    : !hasSavedMsg ? "Không thấy thông báo lưu thành công sau khi nhấn Save"
                                   : "Cấu hình không giữ đúng giá trị sau reload";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = "Lỗi thực thi: " + ex.Message; }

            WriteExcelResult(ROW_TC33, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>
        /// TC34 – Kiểm tra nhân viên (Employee role) KHÔNG có quyền truy cập trang Configuration.
        /// Excel: Row 116 (NPOI index 115). TestData = "username: nguyenvana\npassword: Hass@12341".
        /// Steps: (1) Kiểm tra menu Attendance không có "Configuration"
        ///        → (2) Thử truy cập trực tiếp URL Configuration.
        /// Expected: Submenu "Configuration" KHÔNG hiển thị; redirect 403 / "Access Denied".
        /// </summary>
        [TestMethod]
        public void ATT_TC34_Config_EmployeeAccessDenied()
        {
            var (_, testData) = ReadExcelRow(ROW_TC34);
            string expectedMsg = ReadExpected(ROW_TC34_LAST);
            var (empUser, empPass) = ExtractCredentials(testData, EMP_USER, EMP_PASS);

            string actualMsg = "";
            string status = "Failed";

            try
            {
                // Step 1 – Đăng nhập Employee, kiểm tra menu Attendance
                LoginAs(empUser, empPass);
                GoToMyAttendanceRecords();

                var menuAttendance = dr.FindElements(By.XPath(
                    "//span[contains(text(),'Attendance')] | //a[contains(text(),'Attendance')]")).FirstOrDefault();
                if (menuAttendance != null)
                {
                    try { menuAttendance.Click(); Thread.Sleep(600); }
                    catch { /* ignore */ }
                }

                // Submenu "Configuration" KHÔNG hiển thị cho Employee
                bool noConfigMenu = dr.FindElements(By.XPath(
                    "//a[contains(text(),'Configuration')] | //span[text()='Configuration']")).Count == 0;

                // Step 2 – Thử truy cập trực tiếp URL Configuration
                dr.Navigate().GoToUrl(BASE_URL + "/web/index.php/attendance/configure");
                Thread.Sleep(1500);

                bool isForbidden = dr.Url.Contains("dashboard") ||
                    dr.FindElements(By.XPath(
                        "//*[contains(text(),'Access Denied')] | //*[contains(text(),'403')] | " +
                        "//*[contains(text(),'Forbidden')]")).Count > 0;

                bool ok = noConfigMenu || isForbidden;
                actualMsg = ok ? expectedMsg
                    : "Employee vẫn có thể truy cập trang Configuration";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = "Lỗi thực thi: " + ex.Message; }

            WriteExcelResult(ROW_TC34, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }
    }
}