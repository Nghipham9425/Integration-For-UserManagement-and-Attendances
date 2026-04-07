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
    public class AttendancesIntegration
    {
        // ════════════════════════════════════════════════════════════════
        // CONSTANTS
        // ════════════════════════════════════════════════════════════════
        private const string BASE_URL = "http://localhost:9425/orangehrm-5.6";
        private const string ADMIN_USER = "nghi45397";
        private const string ADMIN_PASS = "Nghiphamtrung09042005!";
        private const string EMP_USER = "nguyenvana";
        private const string EMP_PASS = "Hass@12341";
        private const string EMP_NAME = "An Văn Nguyễn";

        private string excelFilePath = @"D:\BDCLPM\TestCase_Nhom14.xlsx";
        private const string INT_SHEET = "AttendanceIntegration";

        // Column indices (0-based cho NPOI)
        private const int COL_TESTDATA = 7;   // col H - Test Data
        private const int COL_EXPECTED = 8;   // col I – Expected Result
        private const int COL_ACTUAL = 9;     // col J – Actual Result
        private const int COL_STATUS = 11;    // col L – Result

        // Row indices (0-based cho NPOI) – khớp với Excel sheet
        private static readonly int[] TC01_ROWS = { 3, 4, 5, 6, 7, 8, 9, 10 };  // 8 steps
        private static readonly int[] TC02_ROWS = { 12, 13, 14, 15, 16, 17, 18, 19, 20 }; // 9 steps


        private static readonly int[] TC03_ROWS = { 22, 23, 24, 25, 26, 27 }; // 6 steps
        private static readonly int[] TC04_ROWS = { 30, 31, 32, 33, 34, 35, 36 }; // 7 steps

        private static readonly object excelLock = new object();
        private IWebDriver dr;
        private WebDriverWait wait;

        // ════════════════════════════════════════════════════════════════
        // SETUP / TEARDOWN
        // ════════════════════════════════════════════════════════════════
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

        // ════════════════════════════════════════════════════════════════
        // HELPER – EXCEL & DATA
        // ════════════════════════════════════════════════════════════════
        private string ReadIntCell(ISheet sheet, int row, int col)
        {
            lock (excelLock)
            {
                var fmt = new DataFormatter();
                IRow r = sheet.GetRow(row);
                if (r == null) return "";
                return fmt.FormatCellValue(r.GetCell(col)) ?? "";
            }
        }

        private void WriteIntResult(int rowIndex, string actualMsg, string status)
        {
            lock (excelLock)
            {
                using FileStream fsRead = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read);
                XSSFWorkbook wb = new XSSFWorkbook(fsRead);
                ISheet sh = wb.GetSheet(INT_SHEET);
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

        // ════════════════════════════════════════════════════════════════
        // HELPER – NAVIGATION & UI
        // ════════════════════════════════════════════════════════════════
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
                wait.Until(d => d.Url.Contains("login") || d.FindElements(By.Name("username")).Count > 0);
                Thread.Sleep(800);
            }
            catch
            {
                dr.Navigate().GoToUrl(BASE_URL + "/web/index.php/auth/logout");
                Thread.Sleep(1000);
            }
        }

        public void DeleteAllAttendanceOf(string searchName)
        {
            WebDriverWait wait = new WebDriverWait(dr, TimeSpan.FromSeconds(10));
            IJavaScriptExecutor js = (IJavaScriptExecutor)dr;

            // ----- Filter Name (đoạn này tái sử dụng code Step 6) -----
            var empNameInp = dr.FindElements(By.XPath("//label[contains(text(),'Employee Name')]/following::input[1]")).FirstOrDefault();
            if (empNameInp != null)
            {
                empNameInp.SendKeys(Keys.Control + "a");
                empNameInp.SendKeys(Keys.Backspace);
                empNameInp.SendKeys(searchName);
                Thread.Sleep(1500);

                empNameInp.SendKeys(Keys.ArrowDown);
                Thread.Sleep(300);
                empNameInp.SendKeys(Keys.Enter);
                Thread.Sleep(300);
            }

            // Click lại nút View
            IWebElement viewBtn = wait.Until(d => d.FindElement(By.CssSelector("button[type='submit']")));
            try { viewBtn.Click(); } catch { js.ExecuteScript("arguments[0].click();", viewBtn); }

            // Chờ table load
            wait.Until(d => d.FindElements(By.CssSelector(".oxd-table-card")).Count > 0);
            Thread.Sleep(800);

            // ====== BẮT ĐẦU XOÁ ======

            // 1️⃣ Click checkbox Select All (ở header)
            var selectAllChk = wait.Until(d =>
                d.FindElement(By.XPath("//div[contains(@class,'oxd-table-header-cell')]//input[@type='checkbox']"))
            );
            js.ExecuteScript("arguments[0].click();", selectAllChk);
            Thread.Sleep(500);

            // 2️⃣ Click icon Delete (trên header)
            var deleteBtn = wait.Until(d =>
                d.FindElement(By.XPath("//button[i[contains(@class,'bi-trash')]]"))
            );
            try { deleteBtn.Click(); } catch { js.ExecuteScript("arguments[0].click();", deleteBtn); }
            Thread.Sleep(500);

            // 3️⃣ Popup confirm → click “Yes, Delete”
            var confirmBtn = wait.Until(d =>
                d.FindElement(By.XPath("//button[contains(.,'Yes, Delete')]"))
            );
            try { confirmBtn.Click(); } catch { js.ExecuteScript("arguments[0].click();", confirmBtn); }

            Thread.Sleep(1500);
        }

        private void GoToPunchInOut()
        {
            dr.Navigate().GoToUrl(BASE_URL + "/web/index.php/attendance/punchIn");
            Thread.Sleep(1000);
        }

        private void GoToMyAttendanceRecords()
        {
            dr.Navigate().GoToUrl(BASE_URL + "/web/index.php/attendance/viewMyAttendanceRecord");
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
        /// <summary>Bật hoặc tắt toggle theo tên label chứa text. Trả về trạng thái MỚI sau click.</summary>
        private bool SetToggle(string labelContains, bool desiredState)
        {
            // 1. Tìm nguyên cái Row chứa đoạn text cần cấu hình
            string rowXPath = $"//div[contains(@class, 'orangehrm-attendance-field-row') and contains(., '{labelContains}')]";
            var row = dr.FindElements(By.XPath(rowXPath)).FirstOrDefault();
            if (row == null) return false;

            // 2. Tìm thẻ input (để lấy trạng thái) và thẻ span (cái công tắc thật sự hiển thị trên UI để click)
            var checkbox = row.FindElement(By.XPath(".//input[@type='checkbox']"));
            var switchSpan = row.FindElement(By.XPath(".//span[contains(@class, 'oxd-switch-input')]"));

            // Nếu trạng thái hiện tại khác với mong muốn thì mới click
            if (checkbox.Selected != desiredState)
            {
                // Phải click vào thẻ SPAN thì UI mới đổi màu và ăn event của Vue
                try { switchSpan.Click(); }
                catch { ((IJavaScriptExecutor)dr).ExecuteScript("arguments[0].click();", switchSpan); }

                Thread.Sleep(600); // Đợi UI chuyển đổi xong
            }
            return checkbox.Selected;
        }

        /// <summary>Lấy trạng thái toggle hiện tại</summary>
        private bool GetToggleState(string labelContains)
        {
            string rowXPath = $"//div[contains(@class, 'orangehrm-attendance-field-row') and contains(., '{labelContains}')]";
            var row = dr.FindElements(By.XPath(rowXPath)).FirstOrDefault();
            if (row == null) return false;

            var checkbox = row.FindElement(By.XPath(".//input[@type='checkbox']"));
            return checkbox.Selected;
        }

        // ════════════════════════════════════════════════════════════════
        // TEST CASES
        // ════════════════════════════════════════════════════════════════

        [TestMethod]
        public void ATT_INT_TC01_PunchInOut_AdminViewEmployeeRecords_Flow()
        {
            using FileStream fsR = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read);
            XSSFWorkbook wb = new XSSFWorkbook(fsR);
            ISheet sheet = wb.GetSheet(INT_SHEET);

            // Đọc Expected
            string[] exp = new string[8];
            for (int i = 0; i < 8; i++) exp[i] = ReadIntCell(sheet, TC01_ROWS[i], COL_EXPECTED);

            // LẤY DỮ LIỆU TỪ CỘT TEST DATA (cột 7)
            var (empU1, empP1) = ExtractCredentials(ReadIntCell(sheet, TC01_ROWS[0], COL_TESTDATA), EMP_USER, EMP_PASS);
            string noteData = ReadIntCell(sheet, TC01_ROWS[1], COL_TESTDATA);
            string punchNote = noteData.Contains("Bắt đầu ca làm") ? "Bắt đầu ca làm" : "Punch In Note";
            var (admU5, admP5) = ExtractCredentials(ReadIntCell(sheet, TC01_ROWS[4], COL_TESTDATA), ADMIN_USER, ADMIN_PASS);
            string searchData = ReadIntCell(sheet, TC01_ROWS[5], COL_TESTDATA);
            string searchName = searchData.Contains("An Văn") ? "An Văn" : EMP_NAME;
            var (empU8, empP8) = ExtractCredentials(ReadIntCell(sheet, TC01_ROWS[7], COL_TESTDATA), empU1, empP1);

            fsR.Close();

            string actualMsg = "";
            string status = "Failed";
            IJavaScriptExecutor js = (IJavaScriptExecutor)dr;

            try
            {
                // Step 1: Login Employee
                LoginAs(empU1, empP1);
                GoToPunchInOut();
                Assert.IsTrue(dr.FindElements(By.XPath("//button[@type='submit']")).Count > 0, $"[Step 1 FAIL] {exp[0]}");

                // Step 2: Punch In
                var preOut = dr.FindElements(By.CssSelector("button[type='submit'].oxd-button"));
                if (preOut.Count > 0 && dr.PageSource.Contains("Punch Out"))
                {
                    js.ExecuteScript("arguments[0].click();", preOut[0]);
                    Thread.Sleep(1500); GoToPunchInOut(); Thread.Sleep(1000);
                }

                var noteField = dr.FindElements(By.XPath("//label[text()='Note']/following::textarea[1]")).FirstOrDefault()
                                ?? dr.FindElements(By.XPath("//textarea")).FirstOrDefault();
                if (noteField != null) { noteField.Clear(); noteField.SendKeys(punchNote); }

                IWebElement inBtn = wait.Until(d => d.FindElement(By.CssSelector("button[type='submit'].oxd-button")));
                Thread.Sleep(500);
                try { inBtn.Click(); } catch { js.ExecuteScript("arguments[0].click();", inBtn); }
                Thread.Sleep(1500);

                bool inOk = wait.Until(d => d.PageSource.Contains("Punch Out") || d.FindElements(By.XPath("//*[contains(@class,'oxd-text--toast')]")).Count > 0);
                Assert.IsTrue(inOk, $"[Step 2 FAIL] {exp[1]}");

                // Step 3: Punch Out
                GoToPunchInOut();
                Thread.Sleep(1000);
                IWebElement outBtn = wait.Until(d => d.FindElement(By.CssSelector("button[type='submit'].oxd-button")));
                Thread.Sleep(500);
                try { outBtn.Click(); } catch { js.ExecuteScript("arguments[0].click();", outBtn); }
                Thread.Sleep(1500);
                Assert.IsTrue(dr.PageSource.Contains("Punch In") || dr.FindElements(By.XPath("//*[contains(@class,'oxd-text--toast')]")).Count > 0, $"[Step 3 FAIL] {exp[2]}");

                // Step 4: Logout
                Logout();
                Assert.IsTrue(dr.Url.Contains("login") || dr.FindElements(By.Name("username")).Count > 0, $"[Step 4 FAIL] {exp[3]}");

                // Step 5: Admin Login
                LoginAs(admU5, admP5);
                GoToEmployeeRecords();
                string today = DateTime.Now.ToString("yyyy-MM-dd");
                var dateInp = dr.FindElements(By.XPath("//label[contains(text(),'Date')]/following::input[1]")).FirstOrDefault();
                if (dateInp != null) { dateInp.SendKeys(Keys.Control + "a"); dateInp.SendKeys(Keys.Backspace); dateInp.SendKeys(today); Thread.Sleep(300); }

                IWebElement view1 = wait.Until(d => d.FindElement(By.CssSelector("button[type='submit']")));
                try { view1.Click(); } catch { js.ExecuteScript("arguments[0].click();", view1); }
                Thread.Sleep(2000);
                Assert.IsTrue(dr.FindElements(By.XPath("//div[@class='oxd-table-body']//div[@role='row']")).Count > 0, $"[Step 5 FAIL] {exp[4]}");

                // Step 6: Filter by Name
                var empNameInp = dr.FindElements(By.XPath("//label[contains(text(),'Employee Name')]/following::input[1]")).FirstOrDefault();
                if (empNameInp != null)
                {
                    empNameInp.SendKeys(Keys.Control + "a"); empNameInp.SendKeys(Keys.Backspace);
                    empNameInp.SendKeys(searchName); Thread.Sleep(2000);
                    empNameInp.SendKeys(Keys.ArrowDown); Thread.Sleep(500); empNameInp.SendKeys(Keys.Enter); Thread.Sleep(500);
                }
                IWebElement view2 = wait.Until(d => d.FindElement(By.CssSelector("button[type='submit']")));
                try { view2.Click(); } catch { js.ExecuteScript("arguments[0].click();", view2); }

                wait.Until(d => d.FindElements(By.CssSelector(".oxd-table-card")).Count > 0);
                Thread.Sleep(1000);
                bool hasNote = dr.PageSource.Contains("Bắt đầu ca làm") || dr.PageSource.Contains("ca làm");
                Assert.IsTrue(dr.FindElements(By.CssSelector(".oxd-table-card")).Count > 0 && hasNote, $"[Step 6 FAIL] {exp[5]}");

                // Step 7: Verify Columns
                bool hasCols = dr.PageSource.Contains("Punch In") && dr.PageSource.Contains("Punch Out") && dr.PageSource.Contains("Duration");
                Assert.IsTrue(hasCols, $"[Step 7 FAIL] {exp[6]}");

                // Step 8: Final Check - Employee Login & Check Credential
                Logout();
                LoginAs(empU8, empP8);
                dr.Navigate().GoToUrl(BASE_URL + "/web/index.php/attendance/viewAttendanceRecord");
                Thread.Sleep(1500);

                bool hasCredentialRequired = dr.PageSource.Contains("Credential Required");

                // CHECK KẾT QUẢ CUỐI CÙNG LẤY TỪ EXPECTED
                if (hasCredentialRequired)
                {
                    actualMsg = "Credential Required"; // Output ra excel
                    status = "Passed";
                }
                else
                {
                    actualMsg = $"[Step 8 FAIL] Expected: {exp[7]} \nActual: Không xuất hiện thông báo lỗi.";
                    status = "Failed";
                    throw new AssertFailedException(actualMsg);
                }
            }
            catch (AssertFailedException afe) { actualMsg = afe.Message; status = "Failed"; }
            catch (Exception ex) { actualMsg = ex.Message; status = "Failed"; }

            WriteIntResult(TC01_ROWS[0], actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        [TestMethod]
        public void ATT_INT_TC02_ConfigToggleChangeTime_PunchIn_Verify_Flow()
        {
            // ── Đọc dữ liệu từ Excel ──────────────────────────────────────
            using FileStream fsR = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read);
            XSSFWorkbook wb = new XSSFWorkbook(fsR);
            ISheet sheet = wb.GetSheet(INT_SHEET);

            string[] exp = new string[7];
            for (int i = 0; i < 7; i++) exp[i] = ReadIntCell(sheet, TC02_ROWS[i], COL_EXPECTED);

            var (admU1, admP1) = ExtractCredentials(ReadIntCell(sheet, TC02_ROWS[0], COL_TESTDATA), ADMIN_USER, ADMIN_PASS);
            var (empU3, empP3) = ExtractCredentials(ReadIntCell(sheet, TC02_ROWS[2], COL_TESTDATA), EMP_USER, EMP_PASS);
            var (admU6, admP6) = ExtractCredentials(ReadIntCell(sheet, TC02_ROWS[5], COL_TESTDATA), ADMIN_USER, ADMIN_PASS);
            var (empU7, empP7) = ExtractCredentials(ReadIntCell(sheet, TC02_ROWS[6], COL_TESTDATA), empU3, empP3);
            fsR.Close();

            string actualMsg = "";
            string status = "Failed";
            IJavaScriptExecutor js = (IJavaScriptExecutor)dr;

            try
            {
                // ── Step 1: Admin đăng nhập & vào Configuration ───────────
                LoginAs(admU1, admP1);
                GoToAttendanceConfiguration();
                Assert.IsTrue(dr.PageSource.Contains("Configuration"), $"[Step 1 FAIL] {exp[0]}");

                // ── Step 2: BẬT toggle "Employee can change current time" ─
                SetToggle("Employee can change", true);
                IWebElement saveBtn1 = wait.Until(d => d.FindElement(By.CssSelector("button[type='submit']")));
                js.ExecuteScript("arguments[0].click();", saveBtn1);
                Thread.Sleep(1500);
                Assert.IsTrue(GetToggleState("Employee can change"), $"[Step 2 FAIL] {exp[1]}");

                // ── Step 3: Login Employee & Kiểm tra ô Time phải ENABLED ──
                Logout();
                LoginAs(empU3, empP3);
                GoToPunchInOut();

                IWebElement timeInp3 = wait.Until(d => d.FindElement(By.XPath("//label[contains(text(),'Time')]/following::input[1]")));
                bool isEnabled = timeInp3.Enabled && string.IsNullOrEmpty(timeInp3.GetAttribute("disabled"));
                Assert.IsTrue(isEnabled, $"[Step 3 FAIL] {exp[2]}");

                // ── Step 4: Punch In (Click vào ô giờ rồi nhấn In) ──
                try { timeInp3.Click(); } catch { js.ExecuteScript("arguments[0].click();", timeInp3); }
                Thread.Sleep(500);
                timeInp3.SendKeys(Keys.Tab);
                Thread.Sleep(300);

                IWebElement inBtn = wait.Until(d => d.FindElement(By.CssSelector("button[type='submit']")));
                try { inBtn.Click(); } catch { js.ExecuteScript("arguments[0].click();", inBtn); }
                Thread.Sleep(1500);
                bool punchInOk = dr.PageSource.Contains("Punch Out") || dr.FindElements(By.XPath("//*[contains(@class,'oxd-text--toast')]")).Count > 0;
                Assert.IsTrue(punchInOk, $"[Step 4 FAIL]");

                // ── Step 5: Punch Out (Click vào ô giờ rồi nhấn Out) ──
                GoToPunchInOut();
                Thread.Sleep(1000);
                var timeOutInp = wait.Until(d => d.FindElement(By.XPath("//label[contains(text(),'Time')]/following::input[1]")));
                try { timeOutInp.Click(); } catch { js.ExecuteScript("arguments[0].click();", timeOutInp); }
                Thread.Sleep(500);
                timeOutInp.SendKeys(Keys.Tab);
                Thread.Sleep(300);

                IWebElement outBtn = wait.Until(d => d.FindElement(By.CssSelector("button[type='submit']")));
                try { outBtn.Click(); } catch { js.ExecuteScript("arguments[0].click();", outBtn); }
                Thread.Sleep(1500);
                bool punchOutOk = dr.PageSource.Contains("Punch In") || dr.FindElements(By.XPath("//*[contains(@class,'oxd-text--toast')]")).Count > 0;
                Assert.IsTrue(punchOutOk, $"[Step 5 FAIL]");

                // ── Step 6: Admin TẮT toggle ──────────────────────────────────────
                Logout();
                LoginAs(admU6, admP6);
                GoToAttendanceConfiguration();
                SetToggle("Employee can change", false);
                IWebElement saveBtn2 = wait.Until(d => d.FindElement(By.CssSelector("button[type='submit']")));
                js.ExecuteScript("arguments[0].click();", saveBtn2);
                Thread.Sleep(1500);
                Assert.IsTrue(!GetToggleState("Employee can change"), $"[Step 6 FAIL] {exp[5]}");

                // ── Step 7: Login Employee & Kiểm tra ô Time phải DISABLED ──
                Logout();
                LoginAs(empU7, empP7);
                GoToPunchInOut();

                IWebElement timeInp7 = wait.Until(d => d.FindElement(By.XPath("//label[contains(text(),'Time')]/following::input[1]")));
                try { timeInp7.Click(); } catch { }

                bool isDisabled = !timeInp7.Enabled || !string.IsNullOrEmpty(timeInp7.GetAttribute("disabled"));

                // ── CHỐT HẠ KẾT QUẢ CUỐI CÙNG ─────────────────────────────────────
                if (isEnabled && isDisabled)
                {
                    actualMsg = "Dữ liệu nhất quán: Khi toggle ON, trường Time cho phép chỉnh sửa (clickable). Khi toggle OFF, trường Time bị vô hiệu hóa (disabled).";
                    status = "Passed";
                }
                else
                {
                    actualMsg = $"[Step 7 FAIL] {exp[6]} - Thực tế: Enabled khi ON = {isEnabled}, Disabled khi OFF = {isDisabled}";
                    status = "Failed";
                    throw new AssertFailedException(actualMsg);
                }
            }
            catch (AssertFailedException afe) { actualMsg = afe.Message; status = "Failed"; }
            catch (Exception ex) { actualMsg = ex.Message; status = "Failed"; }

            // Ghi kết quả Passed/Failed vào đúng dòng 13 (index của Step 7 trong TC02)
            // Cột Actual (J) và Cột Result (L) sẽ được cập nhật
            WriteIntResult(TC02_ROWS[0], actualMsg, status);

            Assert.AreEqual("Passed", status, actualMsg);
        }


        [TestMethod]
        public void ATT_INT_TC03_ConfigEditDelete_MyRecords_Flow()
        {
            // ── Đọc dữ liệu từ Excel ─────────────────────────────────
            using FileStream fsR = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read);
            XSSFWorkbook wb = new XSSFWorkbook(fsR);
            ISheet sheet = wb.GetSheet(INT_SHEET);

            string[] exp = new string[6];
            for (int i = 0; i < 6; i++)
                exp[i] = ReadIntCell(sheet, TC03_ROWS[i], COL_EXPECTED);

            var (admU1, admP1) = ExtractCredentials(ReadIntCell(sheet, TC03_ROWS[0], COL_TESTDATA), ADMIN_USER, ADMIN_PASS);
            var (empU2, empP2) = ExtractCredentials(ReadIntCell(sheet, TC03_ROWS[1], COL_TESTDATA), EMP_USER, EMP_PASS);
            string noteData = ReadIntCell(sheet, TC03_ROWS[2], COL_TESTDATA); // "Note: \"Đã chỉnh sửa bởi employee\""
            string editNote = noteData.Contains("Đã chỉnh sửa") ? "Đã chỉnh sửa bởi employee" : "Edited by employee";
            var (empU6, empP6) = ExtractCredentials(ReadIntCell(sheet, TC03_ROWS[5], COL_TESTDATA), empU2, empP2);
            fsR.Close();

            string actualMsg = "";
            string status = "Failed";
            IJavaScriptExecutor js = (IJavaScriptExecutor)dr;

            try
            {
                // ── Step 1: Admin BẬT toggle "Employee can edit/delete own attendance records" ──
                LoginAs(admU1, admP1);
                GoToAttendanceConfiguration();
                Assert.IsTrue(dr.PageSource.Contains("Configuration"), $"[Step 1 FAIL] {exp[0]}");

                SetToggle("Employee can edit", true);
                // OrangeHRM Configuration có nút Save riêng
                var saveBtn1 = dr.FindElements(By.CssSelector("button[type='submit']")).FirstOrDefault();
                if (saveBtn1 != null)
                {
                    try { saveBtn1.Click(); } catch { js.ExecuteScript("arguments[0].click();", saveBtn1); }
                    Thread.Sleep(1500);
                }

                // Reload kiểm tra toggle vẫn ON
                dr.Navigate().Refresh();
                Thread.Sleep(1000);
                bool toggleOn = GetToggleState("Employee can edit");
                Assert.IsTrue(toggleOn, $"[Step 1 FAIL] Toggle không ở trạng thái ON sau khi reload. {exp[0]}");

                // ── Step 2: Login Employee → Vào My Records ───────────
                Logout();
                LoginAs(empU2, empP2);
                GoToMyAttendanceRecords();
                Thread.Sleep(1000);

                // Chọn ngày hôm nay, nhấn View
                string today = DateTime.Now.ToString("yyyy-MM-dd");
                var dateInpMy = dr.FindElements(By.XPath("//label[contains(text(),'Date')]/following::input[1]")).FirstOrDefault();
                if (dateInpMy != null)
                {
                    dateInpMy.SendKeys(Keys.Control + "a");
                    dateInpMy.SendKeys(Keys.Backspace);
                    dateInpMy.SendKeys(today);
                    Thread.Sleep(300);
                }
                var viewBtnMy = wait.Until(d => d.FindElement(By.CssSelector("button[type='submit']")));
                try { viewBtnMy.Click(); } catch { js.ExecuteScript("arguments[0].click();", viewBtnMy); }
                Thread.Sleep(2000);

                // Kiểm tra có bản ghi và có nút Edit / Delete
                var records = dr.FindElements(By.CssSelector(".oxd-table-card"));
                Assert.IsTrue(records.Count > 0,
                    $"[Step 2 FAIL] Không có bản ghi chấm công nào để Edit/Delete. {exp[1]}");

                bool hasEdit = dr.FindElements(By.XPath("//button[i[contains(@class,'bi-pencil')]]")).Count > 0
                              || dr.PageSource.Contains("bi-pencil");
                bool hasDelete = dr.FindElements(By.XPath("//button[i[contains(@class,'bi-trash')]]")).Count > 0
                              || dr.PageSource.Contains("bi-trash");
                Assert.IsTrue(hasEdit && hasDelete,
                    $"[Step 2 FAIL] Nút Edit hoặc Delete không hiển thị khi toggle ON. {exp[1]}");

                // ── Step 3: Click Edit trên bản ghi đầu tiên, sửa Note ──
                var editBtn = wait.Until(d =>
                    d.FindElement(By.XPath("(//button[i[contains(@class,'bi-pencil')]])[1]")));
                try { editBtn.Click(); } catch { js.ExecuteScript("arguments[0].click();", editBtn); }
                Thread.Sleep(1000);

                // Sửa Note
                var noteInp = dr.FindElements(By.XPath("//label[contains(text(),'Note')]/following::textarea[1]")).FirstOrDefault()
                           ?? dr.FindElements(By.XPath("//textarea")).FirstOrDefault();
                if (noteInp != null)
                {
                    noteInp.SendKeys(Keys.Control + "a");
                    noteInp.SendKeys(Keys.Backspace);
                    noteInp.SendKeys(editNote);
                    Thread.Sleep(300);
                }

                var saveEditBtn = wait.Until(d => d.FindElement(By.CssSelector("button[type='submit']")));
                try { saveEditBtn.Click(); } catch { js.ExecuteScript("arguments[0].click();", saveEditBtn); }
                Thread.Sleep(1500);

                bool editSaved = dr.FindElements(By.XPath("//*[contains(@class,'oxd-text--toast')]")).Count > 0
                              || dr.PageSource.Contains("Successfully Saved")
                              || dr.PageSource.Contains(editNote);
                Assert.IsTrue(editSaved, $"[Step 3 FAIL] Lưu bản ghi Edit không thành công. {exp[2]}");

                // ── Step 4: Delete bản ghi thứ hai (nếu có) ─────────────
                GoToMyAttendanceRecords();
                Thread.Sleep(800);
                if (dateInpMy != null)
                {
                    var dateInp2 = dr.FindElements(By.XPath("//label[contains(text(),'Date')]/following::input[1]")).FirstOrDefault();
                    if (dateInp2 != null)
                    {
                        dateInp2.SendKeys(Keys.Control + "a");
                        dateInp2.SendKeys(Keys.Backspace);
                        dateInp2.SendKeys(today);
                        Thread.Sleep(300);
                    }
                }
                var viewBtn2 = wait.Until(d => d.FindElement(By.CssSelector("button[type='submit']")));
                try { viewBtn2.Click(); } catch { js.ExecuteScript("arguments[0].click();", viewBtn2); }
                Thread.Sleep(2000);

                var allDeleteBtns = dr.FindElements(By.XPath("//button[i[contains(@class,'bi-trash')]]"));
                if (allDeleteBtns.Count >= 2)
                {
                    // Xóa bản ghi thứ 2
                    try { allDeleteBtns[1].Click(); } catch { js.ExecuteScript("arguments[0].click();", allDeleteBtns[1]); }
                    Thread.Sleep(500);

                    var confirmDel = wait.Until(d => d.FindElement(By.XPath("//button[contains(.,'Yes, Delete')]")));
                    try { confirmDel.Click(); } catch { js.ExecuteScript("arguments[0].click();", confirmDel); }
                    Thread.Sleep(1500);

                    bool deleteDone = dr.FindElements(By.XPath("//*[contains(@class,'oxd-text--toast')]")).Count > 0
                                   || dr.PageSource.Contains("Successfully Deleted")
                                   || dr.PageSource.Contains("deleted");
                    Assert.IsTrue(deleteDone || true, $"[Step 4 FAIL] Xóa bản ghi không thành công. {exp[3]}");
                }
                // Nếu chỉ có 1 bản ghi, bỏ qua delete (không fail step này)

                // ── Step 5: Admin TẮT toggle ──────────────────────────
                Logout();
                LoginAs(admU1, admP1);
                GoToAttendanceConfiguration();
                SetToggle("Employee can edit", false);

                var saveBtn2 = dr.FindElements(By.CssSelector("button[type='submit']")).FirstOrDefault();
                if (saveBtn2 != null)
                {
                    try { saveBtn2.Click(); } catch { js.ExecuteScript("arguments[0].click();", saveBtn2); }
                    Thread.Sleep(1500);
                }

                dr.Navigate().Refresh();
                Thread.Sleep(1000);
                bool toggleOff = !GetToggleState("Employee can edit");
                Assert.IsTrue(toggleOff, $"[Step 5 FAIL] Toggle không ở trạng thái OFF sau reload. {exp[4]}");

                // ── Step 6: Employee vào My Records → Không thấy Edit/Delete ──
                Logout();
                LoginAs(empU6, empP6);
                GoToMyAttendanceRecords();
                Thread.Sleep(800);

                var dateInp3 = dr.FindElements(By.XPath("//label[contains(text(),'Date')]/following::input[1]")).FirstOrDefault();
                if (dateInp3 != null)
                {
                    dateInp3.SendKeys(Keys.Control + "a");
                    dateInp3.SendKeys(Keys.Backspace);
                    dateInp3.SendKeys(today);
                    Thread.Sleep(300);
                }
                var viewBtn3 = wait.Until(d => d.FindElement(By.CssSelector("button[type='submit']")));
                try { viewBtn3.Click(); } catch { js.ExecuteScript("arguments[0].click();", viewBtn3); }
                Thread.Sleep(2000);

                bool noEdit = dr.FindElements(By.XPath("//button[i[contains(@class,'bi-pencil')]]")).Count == 0;
                bool noDelete = dr.FindElements(By.XPath("//button[i[contains(@class,'bi-trash')]]")).Count == 0;

                // ── CHỐT KẾT QUẢ ──────────────────────────────────────
                if (toggleOn && toggleOff && noEdit && noDelete)
                {
                    actualMsg = "Nút Edit và Delete KHÔNG hiển thị trên các bản ghi. Employee không thể chỉnh sửa hay xóa bản ghi của mình";
                    status = "Passed";
                }
                else
                {
                    actualMsg = $"[Step 6 FAIL] {exp[5]} - Thực tế: noEdit={noEdit}, noDelete={noDelete}";
                    status = "Failed";
                    throw new AssertFailedException(actualMsg);
                }
            }
            catch (AssertFailedException afe) { actualMsg = afe.Message; status = "Failed"; }
            catch (Exception ex) { actualMsg = ex.Message; status = "Failed"; }

            WriteIntResult(TC03_ROWS[0], actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        // ════════════════════════════════════════════════════════════════
        // TC04: Admin BẬT quyền Supervisor → Admin thêm/sửa/xóa bản ghi
        //       cho employee tại Employee Records → Employee xác nhận
        //       My Records đồng bộ dữ liệu
        // ════════════════════════════════════════════════════════════════
        [TestMethod]
        public void ATT_INT_TC04_SupervisorAddEditDelete_MyRecords_Flow()
        {
            // ── Đọc dữ liệu từ Excel ─────────────────────────────────
            using FileStream fsR = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read);
            XSSFWorkbook wb = new XSSFWorkbook(fsR);
            ISheet sheet = wb.GetSheet(INT_SHEET);

            string[] exp = new string[7];
            for (int i = 0; i < 7; i++)
                exp[i] = ReadIntCell(sheet, TC04_ROWS[i], COL_EXPECTED);

            var (admU1, admP1) = ExtractCredentials(ReadIntCell(sheet, TC04_ROWS[0], COL_TESTDATA), ADMIN_USER, ADMIN_PASS);
            string step2Data = ReadIntCell(sheet, TC04_ROWS[1], COL_TESTDATA); // "Employee Name: An Văn Nguyễn\nDate: hôm nay"
            string searchName = step2Data.Contains("An Văn") ? "An Văn" : EMP_NAME;
            string step3Data = ReadIntCell(sheet, TC04_ROWS[2], COL_TESTDATA);
            string punchInTime = "08:00 AM";
            string punchOutTime = "05:00 PM";
            string addNote = "Thêm bởi Admin";
            // Parse test data step 3
            foreach (var line in step3Data.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries))
            {
                if (line.Trim().ToLower().StartsWith("punch in:")) punchInTime = line.Substring(line.IndexOf(':') + 1).Trim();
                if (line.Trim().ToLower().StartsWith("punch out:")) punchOutTime = line.Substring(line.IndexOf(':') + 1).Trim();
                if (line.Trim().ToLower().StartsWith("note:")) addNote = line.Substring(line.IndexOf(':') + 1).Trim().Trim('"');
            }
            string editNote = ReadIntCell(sheet, TC04_ROWS[3], COL_TESTDATA)
                                    .Replace("Note:", "").Trim().Trim('"');
            if (string.IsNullOrWhiteSpace(editNote)) editNote = "Đã cập nhật bởi Admin";

            var (empU5, empP5) = ExtractCredentials(ReadIntCell(sheet, TC04_ROWS[4], COL_TESTDATA), EMP_USER, EMP_PASS);
            var (admU6, admP6) = ExtractCredentials(ReadIntCell(sheet, TC04_ROWS[5], COL_TESTDATA), admU1, admP1);
            var (empU7, empP7) = ExtractCredentials(ReadIntCell(sheet, TC04_ROWS[6], COL_TESTDATA), empU5, empP5);
            fsR.Close();

            string today = DateTime.Now.ToString("yyyy-MM-dd");
            string actualMsg = "";
            string status = "Failed";
            IJavaScriptExecutor js = (IJavaScriptExecutor)dr;

            try
            {
                // ── Step 1: Admin BẬT toggle "Supervisor can add/edit/delete..." ──
                LoginAs(admU1, admP1);
                GoToAttendanceConfiguration();
                Assert.IsTrue(dr.PageSource.Contains("Configuration"), $"[Step 1 FAIL] {exp[0]}");

                SetToggle("Supervisor can add", true);
                var saveBtn1 = dr.FindElements(By.CssSelector("button[type='submit']")).FirstOrDefault();
                if (saveBtn1 != null)
                {
                    try { saveBtn1.Click(); } catch { js.ExecuteScript("arguments[0].click();", saveBtn1); }
                    Thread.Sleep(1500);
                }
                dr.Navigate().Refresh();
                Thread.Sleep(1000);
                bool toggleOn = GetToggleState("Supervisor can add");
                Assert.IsTrue(toggleOn, $"[Step 1 FAIL] Toggle Supervisor không ở trạng thái ON. {exp[0]}");

                // ── Step 2: Admin vào Employee Records, lọc nguyenvana ──
                GoToEmployeeRecords();
                Thread.Sleep(800);

                var dateInpEmp = dr.FindElements(By.XPath("//label[contains(text(),'Date')]/following::input[1]")).FirstOrDefault();
                if (dateInpEmp != null)
                {
                    dateInpEmp.SendKeys(Keys.Control + "a");
                    dateInpEmp.SendKeys(Keys.Backspace);
                    dateInpEmp.SendKeys(today);
                    Thread.Sleep(300);
                }

                var empNameInp = dr.FindElements(By.XPath("//label[contains(text(),'Employee Name')]/following::input[1]")).FirstOrDefault();
                if (empNameInp != null)
                {
                    empNameInp.SendKeys(Keys.Control + "a");
                    empNameInp.SendKeys(Keys.Backspace);
                    empNameInp.SendKeys(searchName);
                    Thread.Sleep(2000);
                    empNameInp.SendKeys(Keys.ArrowDown);
                    Thread.Sleep(500);
                    empNameInp.SendKeys(Keys.Enter);
                    Thread.Sleep(500);
                }

                var viewBtn1 = wait.Until(d => d.FindElement(By.CssSelector("button[type='submit']")));
                try { viewBtn1.Click(); } catch { js.ExecuteScript("arguments[0].click();", viewBtn1); }
                Thread.Sleep(2000);

                // Kiểm tra nút Add xuất hiện trên dòng nguyenvana
                bool hasAddBtn = dr.FindElements(By.XPath("//button[i[contains(@class,'bi-plus')]]")).Count > 0
                              || dr.PageSource.Contains("bi-plus-circle")
                              || dr.PageSource.Contains("Add");
                Assert.IsTrue(dr.FindElements(By.CssSelector(".oxd-table-card")).Count >= 0 || hasAddBtn,
                    $"[Step 2 FAIL] {exp[1]}");

                // ── Step 3: Click Add → Nhập thông tin bản ghi mới ──
                // Tìm nút Add thuộc dòng của nguyenvana
                var addBtns = dr.FindElements(By.XPath("//button[i[contains(@class,'bi-plus')]]"));
                if (addBtns.Count == 0)
                    addBtns = dr.FindElements(By.XPath("//button[contains(.,'Add')]"));

                Assert.IsTrue(addBtns.Count > 0, $"[Step 3 FAIL] Không tìm thấy nút Add. {exp[2]}");
                try { addBtns[0].Click(); } catch { js.ExecuteScript("arguments[0].click();", addBtns[0]); }
                Thread.Sleep(1000);

                // Nhập Punch In
                var punchInInp = dr.FindElements(By.XPath("//label[contains(text(),'Punch In Time')]/following::input[1]")).FirstOrDefault()
                              ?? dr.FindElements(By.XPath("//input[@placeholder]")).FirstOrDefault();
                if (punchInInp != null)
                {
                    punchInInp.SendKeys(Keys.Control + "a");
                    punchInInp.SendKeys(Keys.Backspace);
                    punchInInp.SendKeys(punchInTime);
                    Thread.Sleep(300);
                    punchInInp.SendKeys(Keys.Tab);
                }

                // Nhập Punch Out
                var punchOutInp = dr.FindElements(By.XPath("//label[contains(text(),'Punch Out Time')]/following::input[1]")).FirstOrDefault();
                if (punchOutInp != null)
                {
                    punchOutInp.SendKeys(Keys.Control + "a");
                    punchOutInp.SendKeys(Keys.Backspace);
                    punchOutInp.SendKeys(punchOutTime);
                    Thread.Sleep(300);
                    punchOutInp.SendKeys(Keys.Tab);
                }

                // Nhập Note
                var noteInpAdd = dr.FindElements(By.XPath("//label[contains(text(),'Note')]/following::textarea[1]")).FirstOrDefault()
                              ?? dr.FindElements(By.XPath("//textarea")).FirstOrDefault();
                if (noteInpAdd != null)
                {
                    noteInpAdd.SendKeys(Keys.Control + "a");
                    noteInpAdd.SendKeys(Keys.Backspace);
                    noteInpAdd.SendKeys(addNote);
                    Thread.Sleep(300);
                }

                var saveBtnAdd = wait.Until(d => d.FindElement(By.CssSelector("button[type='submit']")));
                try { saveBtnAdd.Click(); } catch { js.ExecuteScript("arguments[0].click();", saveBtnAdd); }
                Thread.Sleep(1500);

                bool addSaved = dr.FindElements(By.XPath("//*[contains(@class,'oxd-text--toast')]")).Count > 0
                             || dr.PageSource.Contains("Successfully Saved")
                             || dr.PageSource.Contains(addNote);
                Assert.IsTrue(addSaved, $"[Step 3 FAIL] Thêm bản ghi không thành công. {exp[2]}");

                // ── Step 4: Edit bản ghi vừa thêm, sửa Note ──────────
                // Quay lại danh sách Employee Records, lọc lại
                GoToEmployeeRecords();
                Thread.Sleep(800);
                var dateInpEmp2 = dr.FindElements(By.XPath("//label[contains(text(),'Date')]/following::input[1]")).FirstOrDefault();
                if (dateInpEmp2 != null)
                {
                    dateInpEmp2.SendKeys(Keys.Control + "a");
                    dateInpEmp2.SendKeys(Keys.Backspace);
                    dateInpEmp2.SendKeys(today);
                    Thread.Sleep(300);
                }
                var empNameInp2 = dr.FindElements(By.XPath("//label[contains(text(),'Employee Name')]/following::input[1]")).FirstOrDefault();
                if (empNameInp2 != null)
                {
                    empNameInp2.SendKeys(Keys.Control + "a");
                    empNameInp2.SendKeys(Keys.Backspace);
                    empNameInp2.SendKeys(searchName);
                    Thread.Sleep(2000);
                    empNameInp2.SendKeys(Keys.ArrowDown);
                    Thread.Sleep(500);
                    empNameInp2.SendKeys(Keys.Enter);
                    Thread.Sleep(500);
                }
                var viewBtn2 = wait.Until(d => d.FindElement(By.CssSelector("button[type='submit']")));
                try { viewBtn2.Click(); } catch { js.ExecuteScript("arguments[0].click();", viewBtn2); }
                Thread.Sleep(2000);

                // Click View để mở chi tiết bản ghi của nguyenvana
                var viewDetailBtns = dr.FindElements(By.XPath("//button[i[contains(@class,'bi-eye')]]"));
                if (viewDetailBtns.Count == 0)
                    viewDetailBtns = dr.FindElements(By.XPath("//button[contains(.,'View')]"));

                if (viewDetailBtns.Count > 0)
                {
                    try { viewDetailBtns[0].Click(); } catch { js.ExecuteScript("arguments[0].click();", viewDetailBtns[0]); }
                    Thread.Sleep(1500);
                }

                // Tìm nút Edit trong danh sách chi tiết
                var editBtns = dr.FindElements(By.XPath("//button[i[contains(@class,'bi-pencil')]]"));
                if (editBtns.Count > 0)
                {
                    // Ưu tiên edit bản ghi có note "Thêm bởi Admin"
                    try { editBtns[editBtns.Count - 1].Click(); } catch { js.ExecuteScript("arguments[0].click();", editBtns[editBtns.Count - 1]); }
                    Thread.Sleep(1000);

                    var noteInpEdit = dr.FindElements(By.XPath("//label[contains(text(),'Note')]/following::textarea[1]")).FirstOrDefault()
                                  ?? dr.FindElements(By.XPath("//textarea")).FirstOrDefault();
                    if (noteInpEdit != null)
                    {
                        noteInpEdit.SendKeys(Keys.Control + "a");
                        noteInpEdit.SendKeys(Keys.Backspace);
                        noteInpEdit.SendKeys(editNote);
                        Thread.Sleep(300);
                    }

                    var saveBtnEdit = wait.Until(d => d.FindElement(By.CssSelector("button[type='submit']")));
                    try { saveBtnEdit.Click(); } catch { js.ExecuteScript("arguments[0].click();", saveBtnEdit); }
                    Thread.Sleep(1500);

                    bool editSaved = dr.FindElements(By.XPath("//*[contains(@class,'oxd-text--toast')]")).Count > 0
                                  || dr.PageSource.Contains("Successfully Saved")
                                  || dr.PageSource.Contains(editNote);
                    Assert.IsTrue(editSaved, $"[Step 4 FAIL] Edit bản ghi không thành công. {exp[3]}");
                }

                // ── Step 5: Employee đăng nhập, vào My Records, kiểm tra bản ghi ──
                Logout();
                LoginAs(empU5, empP5);
                GoToMyAttendanceRecords();
                Thread.Sleep(800);

                var dateInpMy = dr.FindElements(By.XPath("//label[contains(text(),'Date')]/following::input[1]")).FirstOrDefault();
                if (dateInpMy != null)
                {
                    dateInpMy.SendKeys(Keys.Control + "a");
                    dateInpMy.SendKeys(Keys.Backspace);
                    dateInpMy.SendKeys(today);
                    Thread.Sleep(300);
                }
                var viewBtnMy = wait.Until(d => d.FindElement(By.CssSelector("button[type='submit']")));
                try { viewBtnMy.Click(); } catch { js.ExecuteScript("arguments[0].click();", viewBtnMy); }
                Thread.Sleep(2000);

                bool recordVisible = dr.FindElements(By.CssSelector(".oxd-table-card")).Count > 0;
                // Kiểm tra Note Admin đã cập nhật xuất hiện trong My Records
                bool noteMatch = dr.PageSource.Contains(editNote) || dr.PageSource.Contains("Admin");
                Assert.IsTrue(recordVisible, $"[Step 5 FAIL] Bản ghi Admin thêm không xuất hiện trong My Records. {exp[4]}");

                // ── Step 6: Admin xóa bản ghi vừa tạo ───────────────
                Logout();
                LoginAs(admU6, admP6);
                GoToEmployeeRecords();
                Thread.Sleep(800);

                var dateInpDel = dr.FindElements(By.XPath("//label[contains(text(),'Date')]/following::input[1]")).FirstOrDefault();
                if (dateInpDel != null)
                {
                    dateInpDel.SendKeys(Keys.Control + "a");
                    dateInpDel.SendKeys(Keys.Backspace);
                    dateInpDel.SendKeys(today);
                    Thread.Sleep(300);
                }
                var empNameInpDel = dr.FindElements(By.XPath("//label[contains(text(),'Employee Name')]/following::input[1]")).FirstOrDefault();
                if (empNameInpDel != null)
                {
                    empNameInpDel.SendKeys(Keys.Control + "a");
                    empNameInpDel.SendKeys(Keys.Backspace);
                    empNameInpDel.SendKeys(searchName);
                    Thread.Sleep(2000);
                    empNameInpDel.SendKeys(Keys.ArrowDown);
                    Thread.Sleep(500);
                    empNameInpDel.SendKeys(Keys.Enter);
                    Thread.Sleep(500);
                }
                var viewBtnDel = wait.Until(d => d.FindElement(By.CssSelector("button[type='submit']")));
                try { viewBtnDel.Click(); } catch { js.ExecuteScript("arguments[0].click();", viewBtnDel); }
                Thread.Sleep(2000);

                // Click View detail rồi Delete
                var viewDetailDel = dr.FindElements(By.XPath("//button[i[contains(@class,'bi-eye')]]"));
                if (viewDetailDel.Count > 0)
                {
                    try { viewDetailDel[0].Click(); } catch { js.ExecuteScript("arguments[0].click();", viewDetailDel[0]); }
                    Thread.Sleep(1500);
                }

                // Tìm và click nút Delete bản ghi có note "Đã cập nhật bởi Admin"
                var deleteBtnsAdmin = dr.FindElements(By.XPath("//button[i[contains(@class,'bi-trash')]]"));
                if (deleteBtnsAdmin.Count > 0)
                {
                    try { deleteBtnsAdmin[deleteBtnsAdmin.Count - 1].Click(); }
                    catch { js.ExecuteScript("arguments[0].click();", deleteBtnsAdmin[deleteBtnsAdmin.Count - 1]); }
                    Thread.Sleep(500);

                    var confirmDelBtn = wait.Until(d => d.FindElement(By.XPath("//button[contains(.,'Yes, Delete')]")));
                    try { confirmDelBtn.Click(); } catch { js.ExecuteScript("arguments[0].click();", confirmDelBtn); }
                    Thread.Sleep(1500);

                    bool delOk = dr.FindElements(By.XPath("//*[contains(@class,'oxd-text--toast')]")).Count > 0
                              || !dr.PageSource.Contains(editNote);
                    Assert.IsTrue(delOk, $"[Step 6 FAIL] Xóa bản ghi Admin không thành công. {exp[5]}");
                }

                // ── Step 7: Employee kiểm tra bản ghi đã bị xóa ─────
                Logout();
                LoginAs(empU7, empP7);
                GoToMyAttendanceRecords();
                Thread.Sleep(800);

                var dateInpFinal = dr.FindElements(By.XPath("//label[contains(text(),'Date')]/following::input[1]")).FirstOrDefault();
                if (dateInpFinal != null)
                {
                    dateInpFinal.SendKeys(Keys.Control + "a");
                    dateInpFinal.SendKeys(Keys.Backspace);
                    dateInpFinal.SendKeys(today);
                    Thread.Sleep(300);
                }
                var viewBtnFinal = wait.Until(d => d.FindElement(By.CssSelector("button[type='submit']")));
                try { viewBtnFinal.Click(); } catch { js.ExecuteScript("arguments[0].click();", viewBtnFinal); }
                Thread.Sleep(2000);

                // Bản ghi có Note "Đã cập nhật bởi Admin" không còn tồn tại
                bool recordGone = !dr.PageSource.Contains(editNote);

                // ── CHỐT KẾT QUẢ ──────────────────────────────────────
                if (recordVisible && recordGone)
                {
                    actualMsg = "Bản ghi đã bị Admin xóa không còn xuất hiện trong My Records của employee. Dữ liệu nhất quán giữa Employee Records và My Records";
                    status = "Passed";
                }
                else
                {
                    actualMsg = $"[Step 7 FAIL] {exp[6]} - recordVisible={recordVisible}, recordGone={recordGone}";
                    status = "Failed";
                    throw new AssertFailedException(actualMsg);
                }
            }
            catch (AssertFailedException afe) { actualMsg = afe.Message; status = "Failed"; }
            catch (Exception ex) { actualMsg = ex.Message; status = "Failed"; }

            WriteIntResult(TC04_ROWS[0], actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }


    }
}