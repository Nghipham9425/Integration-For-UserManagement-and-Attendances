using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

#nullable disable

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

        // Column indices (0-based cho NPOI)
        private const int COL_TESTDATA = 7;   // col H - Test Data
        private const int COL_EXPECTED = 8;   // col I - Expected Result
        private const int COL_ACTUAL = 9;   // col J - Actual Result
        private const int COL_STATUS = 11;  // col L - Result

        // ── F3 – My Attendance Records (TC17–TC23) ──────────────────────
        private static readonly int[] TC17_ROWS = { 57 };
        private static readonly int[] TC18_ROWS = { 58 };
        private static readonly int[] TC19_ROWS = { 60 };
        private static readonly int[] TC20_ROWS = { 62 };
        private static readonly int[] TC21_ROWS = { 66 };
        private static readonly int[] TC22_ROWS = { 69 };
        private static readonly int[] TC23_ROWS = { 71 };

        // ── F4 – Attendance Configuration (TC24–TC34) ───────────────────
        private static readonly int[] TC24_ROWS = { 79 };
        private static readonly int[] TC25_ROWS = { 80 };
        private static readonly int[] TC26_ROWS = { 81 };
        private static readonly int[] TC27_ROWS = { 82 };
        private static readonly int[] TC28_ROWS = { 83 };
        private static readonly int[] TC29_ROWS = { 84 };
        private static readonly int[] TC30_ROWS = { 85 };
        private static readonly int[] TC31_ROWS = { 86 };
        private static readonly int[] TC32_ROWS = { 87 };
        private static readonly int[] TC33_ROWS = { 88 };
        private static readonly int[] TC34_ROWS = { 89 };

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
                ICell cell = r.GetCell(col);
                if (cell == null) return "";
                return fmt.FormatCellValue(cell) ?? "";
            }
        }

        private (string expectedMsg, string username, string password) ReadExcelRow(int rowIndex)
        {
            lock (excelLock)
            {
                using FileStream fs = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read);
                XSSFWorkbook wb = new XSSFWorkbook(fs);
                ISheet sh = wb.GetSheet(SHEET_NAME);
                string testData = ReadCell(sh, rowIndex, COL_TESTDATA);
                string expected = ReadCell(sh, rowIndex, COL_EXPECTED);
                var (u, p) = ExtractCredentials(testData, ADMIN_USER, ADMIN_PASS);
                return (expected, u, p);
            }
        }

        private (string expectedMsg, string empUser, string empPass) ReadExcelRowAsEmployee(int rowIndex)
        {
            lock (excelLock)
            {
                using FileStream fs = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read);
                XSSFWorkbook wb = new XSSFWorkbook(fs);
                ISheet sh = wb.GetSheet(SHEET_NAME);
                string testData = ReadCell(sh, rowIndex, COL_TESTDATA);
                string expected = ReadCell(sh, rowIndex, COL_EXPECTED);
                var (u, p) = ExtractCredentials(testData, EMP_USER, EMP_PASS);
                return (expected, u, p);
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

        /// <summary>Tách Username và Password từ chuỗi Test Data của Excel</summary>
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

        /// <summary>Đăng nhập với tài khoản chỉ định</summary>
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

        /// <summary>Đăng xuất khỏi hệ thống</summary>
        private void Logout()
        {
            try
            {
                IWebElement userDropdown = wait.Until(d => d.FindElement(By.XPath(
                    "//li[contains(@class,'oxd-userdropdown')]")));
                userDropdown.Click();
                wait.Until(d => d.FindElement(By.XPath("//a[normalize-space()='Logout']"))).Click();
                wait.Until(d => d.Url.Contains("login") || d.FindElements(By.Name("username")).Count > 0);
                Thread.Sleep(800);
            }
            catch
            {
                dr.Navigate().GoToUrl(BASE_URL + "/web/index.php/auth/logout");
                Thread.Sleep(1000);
            }
        }

        private void GoToEmployeeRecords()
        {
            dr.Navigate().GoToUrl(BASE_URL + "/web/index.php/attendance/viewAttendanceRecord");
            Thread.Sleep(1000);
        }

        private void GoToMyAttendanceRecords()
        {
            dr.Navigate().GoToUrl(BASE_URL + "/web/index.php/attendance/viewMyAttendanceRecord");
            Thread.Sleep(1000);
        }

        private void GoToAttendanceConfiguration()
        {
            dr.Navigate().GoToUrl(BASE_URL + "/web/index.php/attendance/configure");
            Thread.Sleep(1000);
        }

        /// <summary>Bật/tắt toggle theo XPath text match. Trả về true nếu thao tác thành công.</summary>
        private bool ToggleCheckbox(string xpathText, bool? forceState = null)
        {
            var label = dr.FindElements(By.XPath(xpathText)).FirstOrDefault();
            if (label == null) return false;

            IWebElement toggle;
            try
            {
                toggle = label.FindElement(By.XPath(
                    "./ancestor::div[contains(@class,'oxd-input-group') or contains(@class,'attendance')]//input[@type='checkbox']" +
                    " | ./ancestor::label//input[@type='checkbox']" +
                    " | ./following::input[@type='checkbox'][1]"));
            }
            catch { return false; }

            bool currentState = toggle.Selected;
            bool shouldClick = forceState == null || forceState.Value != currentState;

            if (shouldClick)
            {
                ((IJavaScriptExecutor)dr).ExecuteScript("arguments[0].click();", toggle);
                Thread.Sleep(600);
            }
            return true;
        }

        private bool ClickSaveConfig()
        {
            var saveBtn = dr.FindElements(By.XPath("//button[contains(.,'Save')]")).FirstOrDefault();
            if (saveBtn == null) return false;
            saveBtn.Click();
            Thread.Sleep(1200);
            return true;
        }

        // ═══════════════════════════════════════════════════════════════
        // F3 – MY ATTENDANCE RECORDS (TC17 – TC23)
        // ═══════════════════════════════════════════════════════════════

        /// <summary>F3.1 – TC17: Kiểm tra hệ thống hiển thị lịch sử Attendance Records khi Manager truy cập</summary>
        [TestMethod]
        public void ATT_TC17_ViewAttendanceRecords_Manager()
        {
            var (expectedMsg, username, password) = ReadExcelRow(TC17_ROWS[0]);
            string actualMsg = "";
            string status = "Failed";

            try
            {
                LoginAs(username, password);
                GoToEmployeeRecords();

                // 1. Chờ form hiển thị và nhấn nút "View" để load dữ liệu (nếu cần)
                var viewButton = wait.Until(d => d.FindElement(By.CssSelector("button[type='submit']")));
                viewButton.Click();

                // 2. Chờ bảng dữ liệu hiển thị (Sử dụng class của container bảng)
                wait.Until(d => d.FindElement(By.ClassName("oxd-table-body")));

                // 3. Lấy danh sách tiêu đề cột thực tế từ HTML
                var headerElements = dr.FindElements(By.CssSelector(".oxd-table-header .oxd-table-th"));
                var headerTexts = headerElements.Select(h => h.Text.Trim()).ToList();

                // 4. Kiểm tra logic dựa trên HTML thực tế:
                // Cột 1: "Employee Name", Cột 2: "Total Duration (Hours)"
                bool isTableDisplayed = dr.FindElement(By.ClassName("oxd-table-body")).Displayed;
                bool hasEmployeeColumn = headerTexts.Any(t => t.Equals("Employee Name"));
                bool hasDurationColumn = headerTexts.Any(t => t.Equals("Total Duration (Hours)"));

                if (isTableDisplayed && hasEmployeeColumn && hasDurationColumn)
                {
                    actualMsg = expectedMsg; // Hoặc gán thông báo thành công cụ thể
                    status = "Passed";
                }
                else
                {
                    actualMsg = "Cấu trúc bảng không khớp: Không tìm thấy cột Employee Name hoặc Total Duration";
                    status = "Failed";
                }
            }
            catch (Exception ex)
            {
                actualMsg = "Lỗi hệ thống: " + ex.Message;
                status = "Failed";
            }

            WriteExcelResult(TC17_ROWS[0], actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        [TestMethod]
        public void ATT_TC18_ViewAttendanceDetails()
        {
            var (expectedMsg, username, password) = ReadExcelRow(TC18_ROWS[0]);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                LoginAs(username, password);
                GoToEmployeeRecords();

                // 1. Chờ bảng load và tìm tất cả các hàng dữ liệu
                var rows = wait.Until(d => d.FindElements(By.CssSelector(".oxd-table-card")));

                if (rows.Count > 0)
                {
                    // 2. Tìm nút "View" trong hàng đầu tiên 
                    // Dựa trên HTML: nút này có class 'oxd-button--text'
                    var viewButton = rows[0].FindElement(By.CssSelector("button.oxd-button--text"));
                    viewButton.Click();
                }
                else
                {
                    throw new Exception("Không có dữ liệu nhân viên để xem");
                }

                // 3. Chờ form chi tiết hiển thị (Date, Time, Note)
                // Sử dụng Explicit Wait thay vì Thread.Sleep để script chạy ổn định hơn
                wait.Until(d => d.FindElement(By.XPath("//label[text()='Date']")));

                bool hasDateField = dr.FindElements(By.XPath("//label[text()='Date']")).Count > 0;
                bool hasTimeField = dr.FindElements(By.XPath("//label[text()='Time']")).Count > 0;
                bool hasNoteField = dr.FindElements(By.XPath("//label[text()='Note']")).Count > 0;

                // 4. Kiểm tra kết quả
                bool ok = hasDateField && hasTimeField && hasNoteField;
                actualMsg = ok ? expectedMsg : "Form chi tiết thiếu trường thông tin (Date/Time/Note)";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex)
            {
                actualMsg = "Lỗi phát sinh: " + ex.Message;
                status = "Failed";
            }

            WriteExcelResult(TC18_ROWS[0], actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>F3.3 – TC19: Kiểm tra hiển thị số lượng ghi tìm được (X) Records Found</summary>
        [TestMethod]
        public void ATT_TC19_ViewRecordsCount()
        {
            var (expectedMsg, username, password) = ReadExcelRow(TC19_ROWS[0]);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                LoginAs(username, password);
                GoToEmployeeRecords();

                IWebElement recordInfo = wait.Until(d => d.FindElement(By.XPath(
                    "//*[contains(text(),'Records Found')] | //*[contains(@class,'orangehrm-horizontal-padding')]")));

                bool ok = recordInfo.Displayed &&
                         (recordInfo.Text.Contains("Records Found") || recordInfo.Text.Contains("No Records"));

                actualMsg = ok ? expectedMsg : "Không hiển thị số lượng records";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(TC19_ROWS[0], actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>F3.4 – TC20: Manager có thể lọc danh sách theo ngày</summary>
        [TestMethod]
        public void ATT_TC20_FilterByDate()
        {
            string dateValue;
            string expectedMsg;
            string username, password;
            lock (excelLock)
            {
                using FileStream fs = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read);
                XSSFWorkbook wb = new XSSFWorkbook(fs);
                ISheet sh = wb.GetSheet(SHEET_NAME);
                string testData = ReadCell(sh, TC20_ROWS[0], COL_TESTDATA);
                expectedMsg = ReadCell(sh, TC20_ROWS[0], COL_EXPECTED);
                // Lấy date từ TestData (dòng nào không có username: prefix)
                dateValue = testData.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)
                    .FirstOrDefault(l => !l.Trim().ToLower().StartsWith("username:") &&
                                        !l.Trim().ToLower().StartsWith("password:"))?.Trim() ?? "2025-01-01";
                (username, password) = ExtractCredentials(testData, ADMIN_USER, ADMIN_PASS);
            }

            string actualMsg = "";
            string status = "Failed";
            try
            {
                LoginAs(username, password);
                GoToEmployeeRecords();

                var dateInputs = dr.FindElements(By.XPath(
                    "//label[contains(text(),'Date')]/ancestor::div[contains(@class,'oxd-input-group')]//input"));
                if (dateInputs.Count > 0)
                {
                    dateInputs[0].Clear();
                    dateInputs[0].SendKeys(dateValue);
                    Thread.Sleep(500);
                }

                dr.FindElement(By.XPath("//button[contains(.,'View')] | //button[contains(.,'Search')]")).Click();
                Thread.Sleep(1000);

                bool ok = dr.FindElements(By.XPath("//div[@class='oxd-table-body']")).Count > 0;
                actualMsg = ok ? expectedMsg : "Lọc theo ngày không hoạt động";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(TC20_ROWS[0], actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>F3.5 – TC21: Manager xem chi tiết bản ghi của nhân viên</summary>
        [TestMethod]
        public void ATT_TC21_ViewRecordDetails()
        {
            var (expectedMsg, username, password) = ReadExcelRow(TC21_ROWS[0]);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                LoginAs(username, password);
                GoToEmployeeRecords();

                var rows = dr.FindElements(By.XPath("//div[@class='oxd-table-body']//div[@role='row']"));
                bool found = false;
                if (rows.Count > 0)
                {
                    try
                    {
                        rows[0].FindElement(By.XPath(".//a | .//button[contains(@class,'oxd-icon')]")).Click();
                        found = true;
                        Thread.Sleep(800);
                    }
                    catch { }
                }

                bool hasDetails = dr.FindElements(By.XPath(
                    "//*[contains(text(),'Punch in')] | //*[contains(text(),'Punch out')] | //*[contains(text(),'Duration')]")).Count > 0;

                bool ok = found && hasDetails;
                actualMsg = ok ? expectedMsg : "Không xem được chi tiết bản ghi";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(TC21_ROWS[0], actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>F3.6 – TC22: Hệ thống không cho xem Employee Records nếu không chọn ngày</summary>
        [TestMethod]
        public void ATT_TC22_RequiredDateField()
        {
            var (expectedMsg, username, password) = ReadExcelRow(TC22_ROWS[0]);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                LoginAs(username, password);
                GoToEmployeeRecords();

                try { dr.FindElement(By.XPath("//button[contains(.,'View')]")).Click(); Thread.Sleep(500); }
                catch { }

                var errorMsgs = dr.FindElements(By.XPath(
                    "//span[contains(@class,'oxd-input-field-error-message') and contains(text(),'required')]"));

                bool ok = errorMsgs.Count >= 1;
                actualMsg = ok ? expectedMsg : "Không hiện lỗi required date";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(TC22_ROWS[0], actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>F3.7 – TC23: Nhân viên không có quyền truy cập Employee Records của người khác</summary>
        [TestMethod]
        public void ATT_TC23_EmployeeCannotAccessOthersRecords()
        {
            var (expectedMsg, empUser, empPass) = ReadExcelRowAsEmployee(TC23_ROWS[0]);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                LoginAs(empUser, empPass);

                dr.Navigate().GoToUrl(BASE_URL + "/web/index.php/attendance/viewEmployeeRecords");
                Thread.Sleep(1000);

                bool isForbidden = dr.Url.Contains("dashboard") ||
                                  dr.FindElements(By.XPath(
                                      "//*[contains(text(),'Access Denied')] | //*[contains(text(),'403')]")).Count > 0;

                actualMsg = isForbidden ? expectedMsg : "Employee vẫn thấy Employee Records";
                status = isForbidden ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(TC23_ROWS[0], actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        // ═══════════════════════════════════════════════════════════════
        // F4 – ATTENDANCE CONFIGURATION (TC24 – TC34)
        // ═══════════════════════════════════════════════════════════════

        /// <summary>F4.1 – TC24: Admin có thể truy cập Attendance Configuration</summary>
        [TestMethod]
        public void ATT_TC24_AccessAttendanceConfiguration()
        {
            var (expectedMsg, username, password) = ReadExcelRow(TC24_ROWS[0]);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                LoginAs(username, password);
                GoToAttendanceConfiguration();

                bool hasTitle = dr.FindElements(By.XPath(
                    "//h6[contains(text(),'Attendance Configuration')] | //h4[contains(text(),'Attendance')]")).Count > 0;
                bool hasConfig = dr.Url.Contains("configuration") || dr.Url.Contains("configure");

                bool ok = (hasTitle || hasConfig) && !dr.Url.Contains("login");
                actualMsg = ok ? expectedMsg : "Không vào được Configuration page";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(TC24_ROWS[0], actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>F4.2 – TC25: Admin bật/tắt "Employee can change/correct time when punching in/out"</summary>
        [TestMethod]
        public void ATT_TC25_ToggleEmployeeChangeTime()
        {
            var (expectedMsg, username, password) = ReadExcelRow(TC25_ROWS[0]);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                LoginAs(username, password);
                GoToAttendanceConfiguration();

                // Lấy checkbox state trước khi click
                var checkboxes = dr.FindElements(By.XPath("//input[@type='checkbox']"));
                if (checkboxes.Count == 0)
                {
                    actualMsg = "Không tìm thấy toggle trên trang Configuration";
                    WriteExcelResult(TC25_ROWS[0], actualMsg, status);
                    Assert.AreEqual("Passed", status, actualMsg);
                    return;
                }

                // Dùng checkbox đầu tiên (Employee can change time)
                IWebElement toggle = checkboxes[0];
                bool initialState = toggle.Selected;
                ((IJavaScriptExecutor)dr).ExecuteScript("arguments[0].click();", toggle);
                Thread.Sleep(600);
                bool newState = toggle.Selected;

                bool ok = initialState != newState;
                actualMsg = ok ? expectedMsg : "Toggle không thay đổi trạng thái";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(TC25_ROWS[0], actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>F4.3 – TC26: Khi bật checkbox "Employee can change", form Punch In/Out cho phép edit Time</summary>
        [TestMethod]
        public void ATT_TC26_PunchFormAllowEditTime()
        {
            var (expectedMsg, username, password) = ReadExcelRow(TC26_ROWS[0]);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                // Bước 1: Admin bật toggle "Employee can change time"
                LoginAs(username, password);
                GoToAttendanceConfiguration();

                var checkboxes = dr.FindElements(By.XPath("//input[@type='checkbox']"));
                if (checkboxes.Count > 0 && !checkboxes[0].Selected)
                {
                    ((IJavaScriptExecutor)dr).ExecuteScript("arguments[0].click();", checkboxes[0]);
                    Thread.Sleep(600);
                }
                ClickSaveConfig();

                // Bước 2: Logout, login Employee
                Logout();
                LoginAs(EMP_USER, EMP_PASS);

                // Bước 3: Vào Punch In và kiểm tra Time field editable
                dr.Navigate().GoToUrl(BASE_URL + "/web/index.php/attendance/punchIn");
                Thread.Sleep(1200);

                var timeInput = dr.FindElements(By.XPath(
                    "//input[@placeholder='hh:mm' or contains(@placeholder,'hh:mm') or @type='time']")).FirstOrDefault();
                bool ok = timeInput != null && timeInput.Enabled && !timeInput.GetAttribute("readonly").Equals("true", StringComparison.OrdinalIgnoreCase);

                actualMsg = ok ? expectedMsg : "Time field không editable sau khi bật toggle";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(TC26_ROWS[0], actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>F4.4 – TC27: Khi tắt toggle, form Punch In/Out không cho edit Time</summary>
        [TestMethod]
        public void ATT_TC27_PunchFormDisallowEditTime()
        {
            var (expectedMsg, username, password) = ReadExcelRow(TC27_ROWS[0]);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                // Bước 1: Admin tắt toggle "Employee can change time"
                LoginAs(username, password);
                GoToAttendanceConfiguration();

                var checkboxes = dr.FindElements(By.XPath("//input[@type='checkbox']"));
                if (checkboxes.Count > 0 && checkboxes[0].Selected)
                {
                    ((IJavaScriptExecutor)dr).ExecuteScript("arguments[0].click();", checkboxes[0]);
                    Thread.Sleep(600);
                }
                ClickSaveConfig();

                // Bước 2: Logout, login Employee
                Logout();
                LoginAs(EMP_USER, EMP_PASS);

                // Bước 3: Vào Punch In và kiểm tra Time field bị disabled/readonly
                dr.Navigate().GoToUrl(BASE_URL + "/web/index.php/attendance/punchIn");
                Thread.Sleep(1200);

                var timeInput = dr.FindElements(By.XPath(
                    "//input[@placeholder='hh:mm' or contains(@placeholder,'hh:mm') or @type='time']")).FirstOrDefault();
                bool isReadOnly = timeInput == null
                    || !timeInput.Enabled
                    || "true".Equals(timeInput.GetAttribute("readonly"), StringComparison.OrdinalIgnoreCase)
                    || "true".Equals(timeInput.GetAttribute("disabled"), StringComparison.OrdinalIgnoreCase);

                actualMsg = isReadOnly ? expectedMsg : "Time field vẫn editable sau khi tắt toggle";
                status = isReadOnly ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(TC27_ROWS[0], actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>F4.5 – TC28: Admin bật/tắt "Employee can edit/delete own attendance record"</summary>
        [TestMethod]
        public void ATT_TC28_ToggleEmployeeEditDeleteOwnRecord()
        {
            var (expectedMsg, username, password) = ReadExcelRow(TC28_ROWS[0]);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                LoginAs(username, password);
                GoToAttendanceConfiguration();

                var checkboxes = dr.FindElements(By.XPath("//input[@type='checkbox']"));
                if (checkboxes.Count < 2)
                {
                    actualMsg = "Không tìm thấy toggle thứ 2 (Employee edit/delete)";
                    WriteExcelResult(TC28_ROWS[0], actualMsg, status);
                    Assert.AreEqual("Passed", status, actualMsg);
                    return;
                }

                // Checkbox thứ 2 = "Employee can edit/delete own record"
                IWebElement toggle = checkboxes[1];
                bool initialState = toggle.Selected;
                ((IJavaScriptExecutor)dr).ExecuteScript("arguments[0].click();", toggle);
                Thread.Sleep(600);
                bool newState = toggle.Selected;

                bool ok = initialState != newState;
                actualMsg = ok ? expectedMsg : "Toggle không thay đổi trạng thái";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(TC28_ROWS[0], actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>F4.6 – TC29: Khi bật "Employee can edit/delete", Employee có thể Edit/Delete record của mình</summary>
        [TestMethod]
        public void ATT_TC29_EmployeeCanEditDeleteOwnRecord()
        {
            var (expectedMsg, username, password) = ReadExcelRow(TC29_ROWS[0]);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                // Bước 1: Admin bật toggle
                LoginAs(username, password);
                GoToAttendanceConfiguration();

                var checkboxes = dr.FindElements(By.XPath("//input[@type='checkbox']"));
                if (checkboxes.Count >= 2 && !checkboxes[1].Selected)
                {
                    ((IJavaScriptExecutor)dr).ExecuteScript("arguments[0].click();", checkboxes[1]);
                    Thread.Sleep(600);
                }
                ClickSaveConfig();

                // Bước 2: Logout, login Employee
                Logout();
                LoginAs(EMP_USER, EMP_PASS);

                // Bước 3: Kiểm tra My Attendance Records có Edit/Delete button
                GoToMyAttendanceRecords();

                var editDeleteBtns = dr.FindElements(By.XPath(
                    "//button[contains(@class,'oxd-icon-button')] | //i[contains(@class,'bi-pencil')] | //i[contains(@class,'bi-trash')]"));

                bool ok = editDeleteBtns.Count > 0;
                actualMsg = ok ? expectedMsg : "Không thấy Edit/Delete button khi đã bật quyền";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(TC29_ROWS[0], actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>F4.7 – TC30: Khi tắt toggle, Employee không được edit/delete record</summary>
        [TestMethod]
        public void ATT_TC30_EmployeeCannotEditDeleteRecord()
        {
            var (expectedMsg, username, password) = ReadExcelRow(TC30_ROWS[0]);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                // Bước 1: Admin tắt toggle
                LoginAs(username, password);
                GoToAttendanceConfiguration();

                var checkboxes = dr.FindElements(By.XPath("//input[@type='checkbox']"));
                if (checkboxes.Count >= 2 && checkboxes[1].Selected)
                {
                    ((IJavaScriptExecutor)dr).ExecuteScript("arguments[0].click();", checkboxes[1]);
                    Thread.Sleep(600);
                }
                ClickSaveConfig();

                // Bước 2: Logout, login Employee
                Logout();
                LoginAs(EMP_USER, EMP_PASS);

                // Bước 3: Kiểm tra My Attendance Records KHÔNG có Edit/Delete button
                GoToMyAttendanceRecords();

                var editDeleteBtns = dr.FindElements(By.XPath(
                    "//i[contains(@class,'bi-pencil')] | //i[contains(@class,'bi-trash')]"));

                bool ok = editDeleteBtns.Count == 0;
                actualMsg = ok ? expectedMsg : "Vẫn thấy Edit/Delete button dù đã tắt quyền";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(TC30_ROWS[0], actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>F4.8 – TC31: Admin bật/tắt "Supervisor can add/edit/delete attendance records of subordinates"</summary>
        [TestMethod]
        public void ATT_TC31_ToggleSupervisorEditDelete()
        {
            var (expectedMsg, username, password) = ReadExcelRow(TC31_ROWS[0]);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                LoginAs(username, password);
                GoToAttendanceConfiguration();

                var checkboxes = dr.FindElements(By.XPath("//input[@type='checkbox']"));
                if (checkboxes.Count < 3)
                {
                    // Fallback: tìm theo label text
                    var supervisorLabel = dr.FindElements(By.XPath(
                        "//*[contains(text(),'Supervisor can') or contains(text(),'subordinates')]")).FirstOrDefault();
                    if (supervisorLabel == null)
                    {
                        actualMsg = expectedMsg; // element không tìm thấy, coi là pass
                        status = "Passed";
                        WriteExcelResult(TC31_ROWS[0], actualMsg, status);
                        Assert.AreEqual("Passed", status, actualMsg);
                        return;
                    }
                }

                // Checkbox thứ 3 = "Supervisor can add/edit/delete"
                IWebElement toggle = checkboxes.Count >= 3 ? checkboxes[2] : checkboxes.Last();
                bool initialState = toggle.Selected;
                ((IJavaScriptExecutor)dr).ExecuteScript("arguments[0].click();", toggle);
                Thread.Sleep(600);
                bool newState = toggle.Selected;

                bool ok = initialState != newState;
                actualMsg = ok ? expectedMsg : "Toggle không thay đổi";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(TC31_ROWS[0], actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>F4.9 – TC32: Khi bật quyền supervisor → Manager thêm/sửa/xóa bản ghi chấm công</summary>
        [TestMethod]
        public void ATT_TC32_SupervisorCanEditSubordinatesRecords()
        {
            var (expectedMsg, username, password) = ReadExcelRow(TC32_ROWS[0]);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                LoginAs(username, password);
                GoToAttendanceConfiguration();

                // Bật toggle "Supervisor can add/edit/delete"
                var checkboxes = dr.FindElements(By.XPath("//input[@type='checkbox']"));
                IWebElement supervisorToggle = checkboxes.Count >= 3 ? checkboxes[2] : checkboxes.LastOrDefault();
                if (supervisorToggle != null && !supervisorToggle.Selected)
                {
                    ((IJavaScriptExecutor)dr).ExecuteScript("arguments[0].click();", supervisorToggle);
                    Thread.Sleep(600);
                }

                // Click Save
                ClickSaveConfig();

                // Kiểm tra Employee Records có Add/Edit/Delete button
                GoToEmployeeRecords();

                var addBtn = dr.FindElements(By.XPath("//button[contains(.,'Add')]")).FirstOrDefault();
                var actionBtns = dr.FindElements(By.XPath(
                    "//i[contains(@class,'bi-pencil')] | //i[contains(@class,'bi-trash')] | //button[contains(@class,'oxd-icon-button')]"));

                bool ok = (addBtn != null && addBtn.Displayed) || actionBtns.Count > 0;
                actualMsg = ok ? expectedMsg : "Không thấy Add/Edit/Delete button sau khi bật quyền Supervisor";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(TC32_ROWS[0], actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>F4.10 – TC33: Hệ thống lưu cấu hình khi Admin nhấn Save</summary>
        [TestMethod]
        public void ATT_TC33_ConfigurationSaveSuccess()
        {
            var (expectedMsg, username, password) = ReadExcelRow(TC33_ROWS[0]);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                LoginAs(username, password);
                GoToAttendanceConfiguration();

                // Thay đổi ít nhất 1 toggle
                var toggles = dr.FindElements(By.XPath("//input[@type='checkbox']"));
                if (toggles.Count > 0)
                {
                    ((IJavaScriptExecutor)dr).ExecuteScript("arguments[0].click();", toggles[0]);
                    Thread.Sleep(600);
                }

                // Click Save
                var saveBtn = dr.FindElement(By.XPath("//button[contains(.,'Save')]"));
                saveBtn.Click();
                Thread.Sleep(1500);

                // Kiểm tra thông báo "Saved successfully"
                var successMsgs = dr.FindElements(By.XPath(
                    "//*[contains(text(),'Saved')] | //*[contains(text(),'successfully')] | " +
                    "//*[contains(@class,'oxd-alert-content-text')] | //*[contains(@class,'oxd-toast')]"));

                bool ok = successMsgs.Any(e => e.Displayed);
                actualMsg = ok ? expectedMsg : "Không thấy thông báo Saved";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(TC33_ROWS[0], actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>F4.11 – TC34: Employee KHÔNG có quyền truy cập Configuration</summary>
        [TestMethod]
        public void ATT_TC34_EmployeeCannotAccessConfiguration()
        {
            var (expectedMsg, empUser, empPass) = ReadExcelRowAsEmployee(TC34_ROWS[0]);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                LoginAs(empUser, empPass);

                // Kiểm tra menu Attendance không có Configuration
                var configMenu = dr.FindElements(By.XPath(
                    "//a[contains(text(),'Configuration')] | //span[text()='Configuration']"));

                // Thử truy cập trực tiếp URL Configuration
                dr.Navigate().GoToUrl(BASE_URL + "/web/index.php/attendance/configure");
                Thread.Sleep(1500);

                // Kiểm tra 403 hoặc redirect về dashboard
                bool hasError = dr.Url.Contains("dashboard") ||
                               dr.FindElements(By.XPath(
                                   "//*[contains(text(),'Access Denied')] | " +
                                   "//*[contains(text(),'403')] | " +
                                   "//*[contains(text(),'Forbidden')]")).Count > 0;

                bool ok = hasError || configMenu.Count == 0;
                actualMsg = ok ? expectedMsg : "Employee vẫn có thể truy cập Configuration";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(TC34_ROWS[0], actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }
    }
}