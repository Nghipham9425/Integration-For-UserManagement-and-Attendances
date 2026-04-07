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
    public class UserManagement
    {
        private const string BASE_URL = "http://localhost:9425/orangehrm-5.6";
        private const string ADMIN_USER = "nghi45397";
        private const string ADMIN_PASS = "Nghiphamtrung09042005!";
        private const string NEW_USERNAME = "nguyenvana";
        private const string NEW_PASSWORD = "Hass@12341";
        private const string EDIT_USERNAME = "emp_renamed";

        private string excelFilePath = @"D:\BDCLPM\TestCase_Nhom14.xlsx";
        private const string SHEET_NAME = "User Management TCs";

        // Column indices (0-based, NPOI)
        private const int COL_TESTDATA = 7;   // col H
        private const int COL_EXPECTED = 8;   // col I
        private const int COL_ACTUAL = 9;   // col J
        private const int COL_STATUS = 11;  // col L

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

            dr.Navigate().GoToUrl(BASE_URL);
            IWebElement txtUser = wait.Until(d => d.FindElement(By.Name("username")));
            txtUser.SendKeys(ADMIN_USER);
            dr.FindElement(By.Name("password")).SendKeys(ADMIN_PASS);
            dr.FindElement(By.CssSelector("button[type='submit']")).Click();
            wait.Until(d => d.Url.Contains("dashboard") ||
                            d.FindElements(By.ClassName("oxd-topbar-header-breadcrumb")).Count > 0);
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

        private void GoToUserManagement()
        {
            try
            {
                IWebElement menuAdmin = wait.Until(d => d.FindElement(By.XPath(
                    "//span[text()='ADMIN'] | //a[normalize-space()='Admin']")));
                menuAdmin.Click();
                wait.Until(d => d.Url.Contains("viewSystemUsers") || d.Url.Contains("admin"));
            }
            catch
            {
                dr.Navigate().GoToUrl(BASE_URL + "/web/index.php/admin/viewSystemUsers");
                wait.Until(d => d.Url.Contains("admin"));
            }
            Thread.Sleep(800);
        }

        private void OpenAddUserForm()
        {
            GoToUserManagement();
            wait.Until(d => d.FindElement(By.XPath("//button[normalize-space()='Add']"))).Click();
            wait.Until(d => d.FindElement(By.XPath("//h6[text()='Add User']")));
        }

        private void SelectOxdOption(IWebElement wrapper, string optionText)
        {
            wrapper.Click();
            Thread.Sleep(300);
            wait.Until(d => d.FindElement(By.XPath(
                $"//div[contains(@class,'oxd-select-dropdown')]//span[text()='{optionText}']"))).Click();
        }

        private IWebElement FindRowByUsername(string username)
        {
            foreach (IWebElement row in dr.FindElements(
                By.XPath("//div[@class='oxd-table-body']//div[@role='row']")))
            {
                var cells = row.FindElements(By.XPath(".//div[@role='cell']"));
                if (cells.Count >= 2 && cells[1].Text.Trim() == username) return row;
            }
            return null;
        }

        private void OpenEditForm(string username)
        {
            IWebElement row = FindRowByUsername(username);
            Assert.IsNotNull(row, $"Không tìm thấy user '{username}' để Edit");
            row.FindElement(By.XPath(".//button[contains(@class,'oxd-icon-button')][2]")).Click();
            wait.Until(d => d.FindElement(By.XPath("//h6[text()='Edit User']")));
        }

        private void SortByColumn(string columnHeaderText, string sortOption)
        {
            IWebElement sortIcon = wait.Until(d => d.FindElement(By.XPath(
                $"//div[contains(@class,'oxd-table-header')]" +
                $"//div[@role='columnheader' and contains(.,'{columnHeaderText}')]" +
                $"//i[contains(@class,'oxd-table-header-sort-icon')]")));
            sortIcon.Click();
            Thread.Sleep(600);
            var deadline = DateTime.Now.AddSeconds(6);
            IWebElement found = null;
            while (DateTime.Now < deadline && found == null)
            {
                try
                {
                    foreach (IWebElement item in dr.FindElements(By.CssSelector(
                        "li.oxd-table-header-sort-dropdown-item span.oxd-text")))
                    {
                        if (item.Displayed && item.Text.Trim() == sortOption) { found = item; break; }
                    }
                }
                catch { }
                if (found == null) Thread.Sleep(150);
            }
            if (found == null) throw new Exception($"Không tìm thấy sort option '{sortOption}'");
            found.Click();
            Thread.Sleep(800);
        }

        private List<string> GetColumnValues(int cellIndex)
            => dr.FindElements(By.XPath(
                    $"//div[@class='oxd-table-body']//div[@role='row']//div[@role='cell'][{cellIndex}]"))
               .Select(c => c.Text.Trim())
               .Where(t => !string.IsNullOrWhiteSpace(t))
               .ToList();

        private void FillEmployeeName(string empHint)
        {
            IWebElement empInput = wait.Until(d => d.FindElement(By.XPath(
                "//label[text()='Employee Name']/following::input[1]")));
            empInput.Click();
            empInput.SendKeys(Keys.Control + "a");
            empInput.SendKeys(Keys.Backspace);
            foreach (char c in empHint) { empInput.SendKeys(c.ToString()); Thread.Sleep(80); }
            Thread.Sleep(2500);
            try
            {
                IWebElement opt = wait.Until(d => d.FindElement(By.XPath(
                    "//div[contains(@class,'oxd-autocomplete-option') or @role='option']")));
                opt.Click();
            }
            catch
            {
                empInput.SendKeys(Keys.ArrowDown);
                Thread.Sleep(400);
                empInput.SendKeys(Keys.Enter);
            }
        }

        private int GetRecordsFound()
        {
            try
            {
                IWebElement info = wait.Until(d => d.FindElement(By.ClassName("orangehrm-horizontal-padding")));
                string text = info.Text;
                int start = text.IndexOf('(') + 1;
                int end = text.IndexOf(')');
                if (start > 0 && end > start)
                    return int.Parse(text.Substring(start, end - start).Trim());
            }
            catch { }
            return -1;
        }

        // ═══════════════════════════════════════════════════════════════
        // F1 – ĐĂNG NHẬP (TC01 – TC03)
        // ═══════════════════════════════════════════════════════════════

        /// <summary>TC01 – Đăng nhập Admin thành công với thông tin hợp lệ.</summary>
        [TestMethod]
        public void UM_TC01_AdminLogin_Success()
        {
            string expectedMsg = ReadExpected(7);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                GoToUserManagement();
                bool ok = dr.Url.Contains("admin") || dr.Url.Contains("viewSystemUsers");
                actualMsg = ok ? expectedMsg : "Không vào được User Management sau login";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(7, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC02 – Đăng nhập thất bại với sai mật khẩu.</summary>
        [TestMethod]
        public void UM_TC02_AdminLogin_WrongPassword()
        {
            string expectedMsg = ReadExpected(10);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                Logout();

                IWebElement txtUser = wait.Until(d => d.FindElement(By.Name("username")));
                txtUser.Clear();
                txtUser.SendKeys(ADMIN_USER);
                dr.FindElement(By.Name("password")).SendKeys("WrongPass");
                dr.FindElement(By.CssSelector("button[type='submit']")).Click();

                IWebElement errEl = wait.Until(d => d.FindElement(By.XPath(
                    "//*[contains(@class,'oxd-alert-content-text')" +
                    " or contains(@class,'oxd-input-field-error-message')]")));

                bool ok = errEl.Displayed && !dr.Url.Contains("dashboard");
                actualMsg = ok ? expectedMsg : "Vẫn vào dashboard dù sai mật khẩu";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(10, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC03 – Đăng nhập thất bại khi để trống username/password.</summary>
        [TestMethod]
        public void UM_TC03_AdminLogin_EmptyFields()
        {
            string expectedMsg = ReadExpected(12);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                Logout();
                dr.FindElement(By.CssSelector("button[type='submit']")).Click();
                wait.Until(d => d.FindElement(By.CssSelector(".oxd-input-field-error-message")));

                var requiredMsgs = dr.FindElements(By.XPath(
                    "//span[contains(@class,'oxd-input-field-error-message')" +
                    " and normalize-space()='Required']"));

                bool ok = requiredMsgs.Count >= 1;
                actualMsg = ok ? expectedMsg : $"Chỉ thấy {requiredMsgs.Count} 'Required'";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(12, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        // ═══════════════════════════════════════════════════════════════
        // F2.1 – XEM DANH SÁCH USER (TC04 – TC05)
        // ═══════════════════════════════════════════════════════════════

        /// <summary>TC04 – Admin xem được danh sách toàn bộ user.</summary>
        [TestMethod]
        public void UM_TC04_ViewUserList()
        {
            string expectedMsg = ReadExpected(16);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                GoToUserManagement();
                IWebElement tableBody = wait.Until(d => d.FindElement(By.CssSelector("div.oxd-table-body")));
                ((IJavaScriptExecutor)dr).ExecuteScript("window.scrollBy(0, -200);");
                IWebElement recordInfo = wait.Until(d => d.FindElement(By.ClassName("orangehrm-horizontal-padding")));

                bool ok = tableBody.Displayed && recordInfo.Text.Contains("Records Found");
                actualMsg = ok ? expectedMsg : "Bảng hoặc Records Found không hiển thị";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(16, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC05 – Danh sách user hiển thị đủ các cột: Username, User Role, Employee Name, Status, Actions.</summary>
        [TestMethod]
        public void UM_TC05_UserList_Columns()
        {
            string expectedMsg = ReadExpected(18);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                GoToUserManagement();
                var headerTexts = dr.FindElements(By.XPath(
                    "//div[@class='oxd-table-header']//div[@role='columnheader']"))
                    .Select(h => h.Text.Trim()).ToList();

                bool allCols = headerTexts.Any(t => t.Contains("Username"))
                            && headerTexts.Any(t => t.Contains("User Role"))
                            && headerTexts.Any(t => t.Contains("Employee Name"))
                            && headerTexts.Any(t => t.Contains("Status"));

                IWebElement adminRow = FindRowByUsername(ADMIN_USER);
                bool ok = allCols && adminRow != null;
                actualMsg = ok ? expectedMsg : "Thiếu cột hoặc không thấy user trong bảng";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(18, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        // ═══════════════════════════════════════════════════════════════
        // F2.2 – TÌM KIẾM (TC06 – TC12)
        // ═══════════════════════════════════════════════════════════════

        /// <summary>TC06 – Tìm kiếm user theo Username.</summary>
        [TestMethod]
        public void UM_TC06_Search_ByUsername()
        {
            string searchValue, expectedMsg;
            lock (excelLock)
            {
                using FileStream fs = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read);
                XSSFWorkbook wb = new XSSFWorkbook(fs);
                ISheet sh = wb.GetSheet(SHEET_NAME);
                searchValue = ReadCell(sh, 20, COL_TESTDATA);
                expectedMsg = ReadCell(sh, 22, COL_EXPECTED);
            }

            string actualMsg = "";
            string status = "Failed";
            try
            {
                GoToUserManagement();
                IWebElement inputUser = wait.Until(d => d.FindElement(By.XPath(
                    "//label[text()='Username']/ancestor::div[contains(@class,'oxd-input-group')]//input")));
                inputUser.Clear();
                inputUser.SendKeys(searchValue);
                dr.FindElement(By.XPath("//button[contains(.,'Search')]")).Click();
                Thread.Sleep(1500);

                var rows = dr.FindElements(By.CssSelector(".oxd-table-body .oxd-table-card"));
                bool allMatch = rows.All(row =>
                {
                    var cells = row.FindElements(By.CssSelector(".oxd-table-cell"));
                    return cells.Count < 2 || cells[1].Text.Contains(searchValue, StringComparison.OrdinalIgnoreCase);
                });

                actualMsg = allMatch ? expectedMsg : $"Kết quả không khớp username '{searchValue}'";
                status = allMatch ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(22, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC07 – Tìm kiếm theo User Role = "ESS".</summary>
        [TestMethod]
        public void UM_TC07_Search_ByRole()
        {
            string roleValue, expectedMsg;
            lock (excelLock)
            {
                using FileStream fs = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read);
                XSSFWorkbook wb = new XSSFWorkbook(fs);
                ISheet sh = wb.GetSheet(SHEET_NAME);
                roleValue = ReadCell(sh, 23, COL_TESTDATA);
                expectedMsg = ReadCell(sh, 25, COL_EXPECTED);
            }

            string actualMsg = "";
            string status = "Failed";
            try
            {
                GoToUserManagement();
                IWebElement roleDropdown = wait.Until(d => d.FindElement(By.XPath(
                    "//label[text()='User Role']/ancestor::div[contains(@class,'oxd-input-group')]" +
                    "//div[contains(@class,'oxd-select-wrapper')]")));
                SelectOxdOption(roleDropdown, roleValue);
                dr.FindElement(By.XPath("//button[contains(.,'Search')]")).Click();
                Thread.Sleep(1500);

                var rows = dr.FindElements(By.CssSelector(".oxd-table-body .oxd-table-card"));
                bool allMatch = rows.All(row =>
                {
                    var cells = row.FindElements(By.CssSelector(".oxd-table-cell"));
                    return cells.Count < 3 || cells[2].Text.Trim() == roleValue;
                });

                actualMsg = allMatch ? expectedMsg : $"Có role không phải '{roleValue}'";
                status = allMatch ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(25, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC08 – Tìm kiếm theo Employee Name.</summary>
        [TestMethod]
        public void UM_TC08_Search_ByEmployeeName()
        {
            string empName, expectedMsg;
            lock (excelLock)
            {
                using FileStream fs = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read);
                XSSFWorkbook wb = new XSSFWorkbook(fs);
                ISheet sh = wb.GetSheet(SHEET_NAME);
                empName = ReadCell(sh, 26, COL_TESTDATA);
                expectedMsg = ReadCell(sh, 28, COL_EXPECTED);
            }

            string actualMsg = "";
            string status = "Failed";
            try
            {
                GoToUserManagement();
                IWebElement empInput = wait.Until(d => d.FindElement(By.XPath(
                    "//label[text()='Employee Name']/ancestor::div[contains(@class,'oxd-input-group')]//input")));
                empInput.Clear();
                empInput.SendKeys(empName);
                Thread.Sleep(2000);
                try
                {
                    IWebElement suggestion = wait.Until(d => d.FindElement(
                        By.CssSelector(".oxd-autocomplete-dropdown .oxd-autocomplete-option")));
                    suggestion.Click();
                }
                catch { }

                dr.FindElement(By.XPath("//button[contains(.,'Search')]")).Click();
                Thread.Sleep(1500);

                var rows = dr.FindElements(By.CssSelector(".oxd-table-body .oxd-table-card"));
                string[] parts = empName.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                string first = parts.FirstOrDefault() ?? empName;
                string last = parts.LastOrDefault() ?? empName;

                bool allMatch = rows.All(row =>
                {
                    var cells = row.FindElements(By.CssSelector(".oxd-table-cell"));
                    if (cells.Count < 4) return true;
                    string name = cells[3].Text.Trim();
                    return name.Contains(first, StringComparison.OrdinalIgnoreCase)
                        || name.Contains(last, StringComparison.OrdinalIgnoreCase);
                });

                actualMsg = allMatch ? expectedMsg : "Employee Name không khớp";
                status = allMatch ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(28, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC09 – Tìm kiếm theo Status = "Enabled".</summary>
        [TestMethod]
        public void UM_TC09_Search_ByStatus()
        {
            string statusValue, expectedMsg;
            lock (excelLock)
            {
                using FileStream fs = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read);
                XSSFWorkbook wb = new XSSFWorkbook(fs);
                ISheet sh = wb.GetSheet(SHEET_NAME);
                statusValue = ReadCell(sh, 29, COL_TESTDATA);
                expectedMsg = ReadCell(sh, 31, COL_EXPECTED);
            }

            string actualMsg = "";
            string status = "Failed";
            try
            {
                GoToUserManagement();
                IWebElement statusDrop = wait.Until(d => d.FindElement(By.XPath(
                    "//label[text()='Status']/ancestor::div[contains(@class,'oxd-input-group')]" +
                    "//div[contains(@class,'oxd-select-wrapper')]")));
                SelectOxdOption(statusDrop, statusValue);
                dr.FindElement(By.XPath("//button[contains(.,'Search')]")).Click();
                Thread.Sleep(1500);

                var rows = dr.FindElements(By.CssSelector(".oxd-table-body .oxd-table-card"));
                bool allMatch = rows.All(row =>
                {
                    var cells = row.FindElements(By.CssSelector(".oxd-table-cell"));
                    return cells.Count < 5 || cells[4].Text.Trim() == statusValue;
                });

                actualMsg = allMatch ? expectedMsg : $"Có status không phải '{statusValue}'";
                status = allMatch ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(31, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC10 – Tìm kiếm kết hợp Username + Role.</summary>
        [TestMethod]
        public void UM_TC10_Search_MultiCondition()
        {
            string expectedMsg = ReadExpected(34);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                GoToUserManagement();
                IWebElement inputUser = wait.Until(d => d.FindElement(By.XPath(
                    "//label[text()='Username']/following::input[1]")));
                inputUser.Clear();
                inputUser.SendKeys("AnNguyen");

                IWebElement roleDropdown = dr.FindElement(By.XPath(
                    "//label[text()='User Role']/following::div[contains(@class,'oxd-select-wrapper')][1]"));
                SelectOxdOption(roleDropdown, "ESS");

                dr.FindElement(By.XPath("//button[contains(.,'Search')]")).Click();
                Thread.Sleep(1200);

                var rows = dr.FindElements(By.XPath("//div[@class='oxd-table-body']//div[@role='row']"));
                bool allMatch = rows.All(row =>
                {
                    var cells = row.FindElements(By.XPath(".//div[@role='cell']"));
                    if (cells.Count < 3) return true;
                    return cells[1].Text.Contains("AnNguyen", StringComparison.OrdinalIgnoreCase)
                        && cells[2].Text.Trim() == "ESS";
                });

                bool ok = rows.Count > 0 && allMatch;
                actualMsg = ok ? expectedMsg : "Kết quả không khớp điều kiện kết hợp";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(34, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC11 – Tìm kiếm không có kết quả → "No Records Found".</summary>
        [TestMethod]
        public void UM_TC11_Search_NoResult()
        {
            string noExistKeyword, expectedMsg;
            lock (excelLock)
            {
                using FileStream fs = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read);
                XSSFWorkbook wb = new XSSFWorkbook(fs);
                ISheet sh = wb.GetSheet(SHEET_NAME);
                noExistKeyword = ReadCell(sh, 35, COL_TESTDATA);
                expectedMsg = ReadCell(sh, 37, COL_EXPECTED);
            }

            string actualMsg = "";
            string status = "Failed";
            try
            {
                GoToUserManagement();
                IWebElement inputUser = wait.Until(d => d.FindElement(By.XPath(
                    "//label[text()='Username']/following::input[1]")));
                inputUser.Clear();
                inputUser.SendKeys(noExistKeyword);
                dr.FindElement(By.XPath("//button[contains(.,'Search')]")).Click();
                Thread.Sleep(1000);

                bool noRecords = dr.FindElements(By.XPath(
                    "//*[contains(text(),'No Records Found')]")).Count > 0
                    || dr.FindElements(By.XPath(
                    "//div[@class='oxd-table-body']//div[@role='row']")).Count == 0;

                actualMsg = noRecords ? expectedMsg : "Vẫn có kết quả dù keyword không tồn tại";
                status = noRecords ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(37, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC12 – Reset bộ lọc → danh sách trở về toàn bộ.</summary>
        [TestMethod]
        public void UM_TC12_Search_Reset()
        {
            string expectedMsg = ReadExpected(40);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                GoToUserManagement();
                IWebElement inputUser = wait.Until(d => d.FindElement(By.XPath(
                    "//label[text()='Username']/following::input[1]")));
                inputUser.Clear();
                inputUser.SendKeys("xyznotexist123");
                dr.FindElement(By.XPath("//button[contains(.,'Search')]")).Click();
                Thread.Sleep(800);

                dr.FindElement(By.XPath("//button[contains(.,'Reset')]")).Click();
                Thread.Sleep(1000);

                string inputVal = dr.FindElement(By.XPath(
                    "//label[text()='Username']/following::input[1]")).GetAttribute("value");
                var rows = dr.FindElements(By.XPath("//div[@class='oxd-table-body']//div[@role='row']"));

                bool ok = inputVal == "" && rows.Count >= 1;
                actualMsg = ok ? expectedMsg : $"Input='{inputVal}', rows={rows.Count}";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(40, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        // ═══════════════════════════════════════════════════════════════
        // F2.3 – SẮP XẾP (TC13 – TC17)
        // ═══════════════════════════════════════════════════════════════

        /// <summary>TC13 – Sắp xếp Username tăng dần A→Z.</summary>
        [TestMethod]
        public void UM_TC13_Sort_Username_Asc()
        {
            string expectedMsg = ReadExpected(43);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                GoToUserManagement();
                SortByColumn("Username", "Ascending");
                var values = GetColumnValues(2);
                var expected = values.OrderBy(u => u, StringComparer.OrdinalIgnoreCase).ToList();
                bool ok = values.SequenceEqual(expected);
                actualMsg = ok ? expectedMsg : "Username không sắp xếp A→Z";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(43, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC14 – Sắp xếp Username giảm dần Z→A.</summary>
        [TestMethod]
        public void UM_TC14_Sort_Username_Desc()
        {
            string expectedMsg = ReadExpected(45);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                GoToUserManagement();
                SortByColumn("Username", "Descending");
                var values = GetColumnValues(2);
                var expected = values.OrderByDescending(u => u, StringComparer.OrdinalIgnoreCase).ToList();
                bool ok = values.SequenceEqual(expected);
                actualMsg = ok ? expectedMsg : "Username không sắp xếp Z→A";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(45, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC15 – Sắp xếp User Role (Asc + Desc).</summary>
        [TestMethod]
        [DataRow("Ascending", 46)]
        [DataRow("Descending", 47)]
        public void UM_TC15_Sort_UserRole(string direction, int rowIdx)
        {
            string expectedMsg = ReadExpected(rowIdx);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                GoToUserManagement();
                SortByColumn("User Role", direction);
                var values = GetColumnValues(3);
                var expected = direction == "Ascending"
                    ? values.OrderBy(r => r, StringComparer.OrdinalIgnoreCase).ToList()
                    : values.OrderByDescending(r => r, StringComparer.OrdinalIgnoreCase).ToList();
                bool ok = values.SequenceEqual(expected);
                actualMsg = ok ? expectedMsg : $"User Role không sort {direction}";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(rowIdx, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC16 – Sắp xếp Employee Name (Asc + Desc).</summary>
        [TestMethod]
        [DataRow("Ascending", 48)]
        [DataRow("Descending", 49)]
        public void UM_TC16_Sort_EmployeeName(string direction, int rowIdx)
        {
            string expectedMsg = ReadExpected(rowIdx);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                GoToUserManagement();
                var before = dr.FindElements(By.XPath(
                    "//div[@class='oxd-table-body']//div[@role='row']//div[@role='cell'][3]"))
                    .Select(c => c.Text.Trim().Split('\n').First().Trim())
                    .Where(t => !string.IsNullOrWhiteSpace(t)).ToList();

                SortByColumn("Employee Name", direction);

                var after = dr.FindElements(By.XPath(
                    "//div[@class='oxd-table-body']//div[@role='row']//div[@role='cell'][3]"))
                    .Select(c => c.Text.Trim().Split('\n').First().Trim())
                    .Where(t => !string.IsNullOrWhiteSpace(t)).ToList();

                bool ok = after.Count > 1 && !before.SequenceEqual(after);
                actualMsg = ok ? expectedMsg : $"Employee Name không đổi thứ tự khi sort {direction}";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(rowIdx, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC17 – Sắp xếp Status (Asc + Desc).</summary>
        [TestMethod]
        [DataRow("Ascending", 50)]
        [DataRow("Descending", 51)]
        public void UM_TC17_Sort_Status(string direction, int rowIdx)
        {
            string expectedMsg = ReadExpected(rowIdx);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                GoToUserManagement();
                SortByColumn("Status", direction);
                var values = GetColumnValues(5);
                var expected = direction == "Ascending"
                    ? values.OrderBy(s => s, StringComparer.OrdinalIgnoreCase).ToList()
                    : values.OrderByDescending(s => s, StringComparer.OrdinalIgnoreCase).ToList();
                bool ok = values.SequenceEqual(expected);
                actualMsg = ok ? expectedMsg : $"Status không sort {direction}";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(rowIdx, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        // ═══════════════════════════════════════════════════════════════
        // F2.4 – THÊM USER (TC18 – TC23)
        // ═══════════════════════════════════════════════════════════════

        /// <summary>TC18 – Thêm user mới đầy đủ thông tin hợp lệ.</summary>
        [TestMethod]
        public void UM_TC18_AddUser_Success()
        {
            string role, empName, statusVal, username, password, confirmPw, expectedMsg;
            lock (excelLock)
            {
                using FileStream fs = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read);
                XSSFWorkbook wb = new XSSFWorkbook(fs);
                ISheet sh = wb.GetSheet(SHEET_NAME);
                role = ReadCell(sh, 54, COL_TESTDATA);
                empName = ReadCell(sh, 55, COL_TESTDATA);
                statusVal = ReadCell(sh, 56, COL_TESTDATA);
                username = ReadCell(sh, 57, COL_TESTDATA);
                password = ReadCell(sh, 58, COL_TESTDATA);
                confirmPw = ReadCell(sh, 59, COL_TESTDATA);
                expectedMsg = ReadCell(sh, 60, COL_EXPECTED);
            }

            string actualMsg = "";
            string status = "Failed";
            try
            {
                OpenAddUserForm();

                IWebElement roleDrop = wait.Until(d => d.FindElement(By.XPath(
                    "//label[text()='User Role']/following::div[contains(@class,'oxd-select-wrapper')][1]")));
                SelectOxdOption(roleDrop, role.Length > 0 ? role : "ESS");

                string nameToType = empName.Length > 0 ? empName : "An Văn";
                FillEmployeeName(nameToType.Contains(" ") ? nameToType.Substring(0, nameToType.LastIndexOf(' ')) : nameToType);

                IWebElement statusDrop = dr.FindElement(By.XPath(
                    "//label[text()='Status']/following::div[contains(@class,'oxd-select-wrapper')][1]"));
                SelectOxdOption(statusDrop, statusVal.Length > 0 ? statusVal : "Enabled");

                IWebElement txtUser = dr.FindElement(By.XPath("//label[text()='Username']/following::input[1]"));
                txtUser.Clear(); txtUser.SendKeys(username.Length > 0 ? username : NEW_USERNAME);

                IWebElement txtPass = dr.FindElement(By.XPath("//label[text()='Password']/following::input[1]"));
                txtPass.Clear(); txtPass.SendKeys(password.Length > 0 ? password : NEW_PASSWORD);

                IWebElement txtConfirm = dr.FindElement(By.XPath("//label[text()='Confirm Password']/following::input[1]"));
                txtConfirm.Clear(); txtConfirm.SendKeys(confirmPw.Length > 0 ? confirmPw : NEW_PASSWORD);

                dr.FindElement(By.XPath("//button[@type='submit' and contains(.,'Save')]")).Click();

                IWebElement toastMsg = wait.Until(d => d.FindElement(By.XPath(
                    "//p[contains(@class,'oxd-text--toast-message')]")));
                actualMsg = toastMsg.Text.Trim();
                status = actualMsg == expectedMsg ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(60, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC19 – Validation: để trống Username → "Username is required".</summary>
        [TestMethod]
        public void UM_TC19_AddUser_UsernameRequired()
        {
            string expectedMsg = ReadExpected(62);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                OpenAddUserForm();
                IWebElement roleDrop = wait.Until(d => d.FindElement(By.XPath(
                    "//label[text()='User Role']/following::div[contains(@class,'oxd-select-wrapper')][1]")));
                SelectOxdOption(roleDrop, "ESS");
                IWebElement statusDrop = dr.FindElement(By.XPath(
                    "//label[text()='Status']/following::div[contains(@class,'oxd-select-wrapper')][1]"));
                SelectOxdOption(statusDrop, "Enabled");

                // Bỏ trống Username, điền Password
                IWebElement txtPass = dr.FindElement(By.XPath("//label[text()='Password']/following::input[1]"));
                txtPass.Clear(); txtPass.SendKeys(NEW_PASSWORD);
                IWebElement txtConfirm = dr.FindElement(By.XPath("//label[text()='Confirm Password']/following::input[1]"));
                txtConfirm.Clear(); txtConfirm.SendKeys(NEW_PASSWORD);

                dr.FindElement(By.XPath("//button[@type='submit' and contains(.,'Save')]")).Click();

                IWebElement errMsg = wait.Until(d => d.FindElement(By.XPath(
                    "//label[text()='Username']/following::span[contains(@class,'error')][1]")));

                actualMsg = errMsg.Text.Trim();
                status = actualMsg == expectedMsg ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(62, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC20 – Validation: để trống Password → "Password is required".</summary>
        [TestMethod]
        public void UM_TC20_AddUser_PasswordRequired()
        {
            string expectedMsg = ReadExpected(64);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                OpenAddUserForm();
                IWebElement roleDrop = wait.Until(d => d.FindElement(By.XPath(
                    "//label[text()='User Role']/following::div[contains(@class,'oxd-select-wrapper')][1]")));
                SelectOxdOption(roleDrop, "ESS");
                IWebElement statusDrop = dr.FindElement(By.XPath(
                    "//label[text()='Status']/following::div[contains(@class,'oxd-select-wrapper')][1]"));
                SelectOxdOption(statusDrop, "Enabled");

                IWebElement txtUser = dr.FindElement(By.XPath("//label[text()='Username']/following::input[1]"));
                txtUser.Clear(); txtUser.SendKeys(NEW_USERNAME + "_nopwd");

                dr.FindElement(By.XPath("//button[@type='submit' and contains(.,'Save')]")).Click();

                IWebElement errMsg = wait.Until(d => d.FindElement(By.XPath(
                    "//label[text()='Password']/following::span[contains(@class,'error')][1]")));

                actualMsg = errMsg.Text.Trim();
                status = actualMsg == expectedMsg ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(64, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC21 – Validation: Username trùng → "Already exists".</summary>
        [TestMethod]
        public void UM_TC21_AddUser_DuplicateUsername()
        {
            string dupUser, expectedMsg;
            lock (excelLock)
            {
                using FileStream fs = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read);
                XSSFWorkbook wb = new XSSFWorkbook(fs);
                ISheet sh = wb.GetSheet(SHEET_NAME);
                dupUser = ReadCell(sh, 65, COL_TESTDATA);
                expectedMsg = ReadCell(sh, 66, COL_EXPECTED);
            }
            if (string.IsNullOrWhiteSpace(dupUser)) dupUser = ADMIN_USER;

            string actualMsg = "";
            string status = "Failed";
            try
            {
                OpenAddUserForm();
                IWebElement roleDrop = wait.Until(d => d.FindElement(By.XPath(
                    "//label[text()='User Role']/following::div[contains(@class,'oxd-select-wrapper')][1]")));
                SelectOxdOption(roleDrop, "ESS");
                IWebElement statusDrop = dr.FindElement(By.XPath(
                    "//label[text()='Status']/following::div[contains(@class,'oxd-select-wrapper')][1]"));
                SelectOxdOption(statusDrop, "Enabled");

                IWebElement empInput = dr.FindElement(By.XPath(
                    "//label[text()='Employee Name']/following::input[1]"));
                empInput.SendKeys("a");
                Thread.Sleep(800);
                empInput.SendKeys(Keys.ArrowDown);
                empInput.SendKeys(Keys.Enter);

                IWebElement txtUser = dr.FindElement(By.XPath("//label[text()='Username']/following::input[1]"));
                txtUser.Clear(); txtUser.SendKeys(dupUser);

                IWebElement txtPass = dr.FindElement(By.XPath("//label[text()='Password']/following::input[1]"));
                txtPass.Clear(); txtPass.SendKeys(NEW_PASSWORD);
                IWebElement txtConfirm = dr.FindElement(By.XPath("//label[text()='Confirm Password']/following::input[1]"));
                txtConfirm.Clear(); txtConfirm.SendKeys(NEW_PASSWORD);

                dr.FindElement(By.XPath("//button[@type='submit' and contains(.,'Save')]")).Click();

                IWebElement errEl = wait.Until(d => d.FindElement(By.XPath(
                    "//span[contains(@class,'oxd-input-field-error-message') and text()='Already exists']")));

                actualMsg = errEl.Text.Trim();
                status = actualMsg == expectedMsg ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(66, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC22 – Validation: Password quá ngắn → lỗi độ phức tạp.</summary>
        [TestMethod]
        public void UM_TC22_AddUser_WeakPassword()
        {
            string weakPwd, expectedMsg;
            lock (excelLock)
            {
                using FileStream fs = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read);
                XSSFWorkbook wb = new XSSFWorkbook(fs);
                ISheet sh = wb.GetSheet(SHEET_NAME);
                weakPwd = ReadCell(sh, 67, COL_TESTDATA);
                expectedMsg = ReadCell(sh, 68, COL_EXPECTED);
            }
            if (string.IsNullOrWhiteSpace(weakPwd)) weakPwd = "123";

            string actualMsg = "";
            string status = "Failed";
            try
            {
                OpenAddUserForm();
                IWebElement roleDrop = wait.Until(d => d.FindElement(By.XPath(
                    "//label[text()='User Role']/following::div[contains(@class,'oxd-select-wrapper')][1]")));
                SelectOxdOption(roleDrop, "ESS");
                IWebElement statusDrop = dr.FindElement(By.XPath(
                    "//label[text()='Status']/following::div[contains(@class,'oxd-select-wrapper')][1]"));
                SelectOxdOption(statusDrop, "Enabled");

                IWebElement txtUser = dr.FindElement(By.XPath("//label[text()='Username']/following::input[1]"));
                txtUser.Clear(); txtUser.SendKeys(NEW_USERNAME + "_weak");

                IWebElement txtPass = dr.FindElement(By.XPath("//label[text()='Password']/following::input[1]"));
                txtPass.Clear(); txtPass.SendKeys(weakPwd);
                IWebElement txtConfirm = dr.FindElement(By.XPath("//label[text()='Confirm Password']/following::input[1]"));
                txtConfirm.Clear(); txtConfirm.SendKeys(weakPwd);

                dr.FindElement(By.XPath("//button[@type='submit' and contains(.,'Save')]")).Click();

                IWebElement errMsg = wait.Until(d => d.FindElement(By.XPath(
                    "//span[contains(@class,'oxd-input-field-error-message') " +
                    "and normalize-space()='Should have at least 8 characters']")));

                actualMsg = errMsg.Text.Trim();
                status = actualMsg == expectedMsg ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(68, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC23 – Cancel khi thêm user → form đóng, không tạo user mới.</summary>
        [TestMethod]
        public void UM_TC23_AddUser_Cancel()
        {
            string cancelUser, expectedMsg;
            lock (excelLock)
            {
                using FileStream fs = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read);
                XSSFWorkbook wb = new XSSFWorkbook(fs);
                ISheet sh = wb.GetSheet(SHEET_NAME);
                cancelUser = ReadCell(sh, 69, COL_TESTDATA);
                expectedMsg = ReadCell(sh, 70, COL_EXPECTED);
            }
            if (string.IsNullOrWhiteSpace(cancelUser)) cancelUser = "testcancel";

            string actualMsg = "";
            string status = "Failed";
            try
            {
                OpenAddUserForm();
                IWebElement txtUser = wait.Until(d => d.FindElement(By.XPath(
                    "//label[text()='Username']/following::input[1]")));
                txtUser.Clear(); txtUser.SendKeys(cancelUser);

                dr.FindElement(By.XPath(
                    "//button[@type='button' and normalize-space()='Cancel']")).Click();
                wait.Until(d => d.FindElement(By.XPath("//h5[normalize-space()='System Users']")));

                bool exists = dr.FindElements(By.XPath(
                    $"//div[@class='oxd-table-card']//div[normalize-space()='{cancelUser}']")).Count > 0;

                actualMsg = !exists ? expectedMsg : $"User '{cancelUser}' vẫn xuất hiện sau Cancel";
                status = !exists ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(70, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        // ═══════════════════════════════════════════════════════════════
        // F2.5 – SỬA USER (TC24 – TC29)
        // ═══════════════════════════════════════════════════════════════

        /// <summary>TC24 – Sửa Role (ESS → Admin) và Status (Enabled → Disabled) thành công.</summary>
        [TestMethod]
        public void UM_TC24_EditUser_ChangeRoleStatus()
        {
            string expectedMsg = ReadExpected(75);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                GoToUserManagement();
                OpenEditForm(NEW_USERNAME);

                IWebElement roleDrop = wait.Until(d => d.FindElement(By.XPath(
                    "//label[text()='User Role']/following::div[contains(@class,'oxd-select-wrapper')][1]")));
                SelectOxdOption(roleDrop, "Admin");

                IWebElement statusDrop = dr.FindElement(By.XPath(
                    "//label[text()='Status']/following::div[contains(@class,'oxd-select-wrapper')][1]"));
                SelectOxdOption(statusDrop, "Disabled");

                dr.FindElement(By.XPath("//button[@type='submit' and contains(.,'Save')]")).Click();

                IWebElement toastMsg = wait.Until(d => d.FindElement(By.XPath(
                    "//p[contains(@class,'oxd-text--toast-message')]")));
                actualMsg = toastMsg.Text.Trim();
                status = actualMsg == expectedMsg ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(75, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC25 – Sửa Username thành công.</summary>
        [TestMethod]
        public void UM_TC25_EditUser_ChangeUsername()
        {
            string newUsernameFromExcel, expectedMsg;
            lock (excelLock)
            {
                using FileStream fs = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read);
                XSSFWorkbook wb = new XSSFWorkbook(fs);
                ISheet sh = wb.GetSheet(SHEET_NAME);
                newUsernameFromExcel = ReadCell(sh, 76, COL_TESTDATA);
                expectedMsg = ReadCell(sh, 77, COL_EXPECTED);
            }
            string editTo = !string.IsNullOrWhiteSpace(newUsernameFromExcel) ? newUsernameFromExcel : EDIT_USERNAME;

            string actualMsg = "";
            string status = "Failed";
            try
            {
                GoToUserManagement();
                string target = FindRowByUsername(NEW_USERNAME) != null ? NEW_USERNAME : EDIT_USERNAME;
                OpenEditForm(target);

                IWebElement txtUser = wait.Until(d => d.FindElement(By.XPath(
                    "//label[text()='Username']/following::input[1]")));
                txtUser.Clear();
                txtUser.SendKeys(editTo);

                IWebElement saveBtn = dr.FindElement(By.XPath("//button[@type='submit' and contains(.,'Save')]"));
                ((IJavaScriptExecutor)dr).ExecuteScript("arguments[0].scrollIntoView(true);", saveBtn);
                Thread.Sleep(500);
                ((IJavaScriptExecutor)dr).ExecuteScript("arguments[0].click();", saveBtn);

                IWebElement toastMsg = wait.Until(d => d.FindElement(By.XPath(
                    "//p[contains(@class,'oxd-text--toast-message')]")));
                actualMsg = toastMsg.Text.Trim();

                GoToUserManagement();
                bool found = FindRowByUsername(editTo) != null;
                status = (actualMsg == expectedMsg && found) ? "Passed" : "Failed";
                if (!found) actualMsg = $"Toast OK nhưng không tìm thấy '{editTo}' trong danh sách";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(77, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC26 – Sửa Username và đổi Password.</summary>
        [TestMethod]
        public void UM_TC26_EditUser_ChangeUsername_Password()
        {
            string newUsernameFromExcel, newPasswordFromExcel, expectedMsg;
            lock (excelLock)
            {
                using FileStream fs = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read);
                XSSFWorkbook wb = new XSSFWorkbook(fs);
                ISheet sh = wb.GetSheet(SHEET_NAME);
                newUsernameFromExcel = ReadCell(sh, 76, COL_TESTDATA);
                newPasswordFromExcel = ReadCell(sh, 76, COL_ACTUAL);  // col J – new password nếu có
                expectedMsg = ReadCell(sh, 77, COL_EXPECTED);
            }
            string editTo = !string.IsNullOrWhiteSpace(newUsernameFromExcel) ? newUsernameFromExcel : EDIT_USERNAME;
            string newPass = !string.IsNullOrWhiteSpace(newPasswordFromExcel) ? newPasswordFromExcel : NEW_PASSWORD;

            string actualMsg = "";
            string status = "Failed";
            try
            {
                GoToUserManagement();
                string target = FindRowByUsername(NEW_USERNAME) != null ? NEW_USERNAME : EDIT_USERNAME;
                OpenEditForm(target);

                IWebElement txtUser = wait.Until(d => d.FindElement(By.XPath(
                    "//label[text()='Username']/following::input[1]")));
                txtUser.Clear();
                txtUser.SendKeys(editTo);

                // Tick "Change Password?"
                IWebElement chkChangePass = dr.FindElement(By.XPath(
                    "//label[text()='Change Password ?']/following::input[1]"));
                if (!chkChangePass.Selected)
                    ((IJavaScriptExecutor)dr).ExecuteScript("arguments[0].click();", chkChangePass);
                Thread.Sleep(500);

                IWebElement txtPass = wait.Until(d => d.FindElement(By.XPath(
                    "//label[text()='Password']/following::input[1]")));
                txtPass.Clear(); txtPass.SendKeys(newPass);
                IWebElement txtConfirm = dr.FindElement(By.XPath(
                    "//label[text()='Confirm Password']/following::input[1]"));
                txtConfirm.Clear(); txtConfirm.SendKeys(newPass);

                IWebElement btnSave = dr.FindElement(By.XPath("//button[@type='submit' and contains(.,'Save')]"));
                ((IJavaScriptExecutor)dr).ExecuteScript("arguments[0].scrollIntoView(true);", btnSave);
                ((IJavaScriptExecutor)dr).ExecuteScript("arguments[0].click();", btnSave);

                IWebElement toastMsg = wait.Until(d => d.FindElement(By.XPath(
                    "//p[contains(@class,'oxd-text--toast-message')]")));
                actualMsg = toastMsg.Text.Trim();

                GoToUserManagement();
                bool found = FindRowByUsername(editTo) != null;
                status = (actualMsg == expectedMsg && found) ? "Passed" : "Failed";
                if (!found) actualMsg = $"Toast OK nhưng không tìm thấy '{editTo}' trong danh sách";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(77, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC27 – Không nhập Password khi Edit → password cũ không đổi, lưu thành công.</summary>
        [TestMethod]
        public void UM_TC27_EditUser_PasswordUnchanged()
        {
            string expectedMsg = ReadExpected(83);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                GoToUserManagement();
                string target = FindRowByUsername(EDIT_USERNAME) != null ? EDIT_USERNAME : NEW_USERNAME;
                OpenEditForm(target);

                IWebElement statusDrop = wait.Until(d => d.FindElement(By.XPath(
                    "//label[text()='Status']/following::div[contains(@class,'oxd-select-wrapper')][1]")));
                SelectOxdOption(statusDrop, "Enabled");

                IWebElement saveBtn = dr.FindElement(By.XPath("//button[@type='submit' and contains(.,'Save')]"));
                ((IJavaScriptExecutor)dr).ExecuteScript("arguments[0].scrollIntoView(true);", saveBtn);
                Thread.Sleep(500);
                ((IJavaScriptExecutor)dr).ExecuteScript("arguments[0].click();", saveBtn);

                IWebElement toastMsg = wait.Until(d => d.FindElement(By.XPath(
                    "//p[contains(@class,'oxd-text--toast-message')]")));
                actualMsg = toastMsg.Text.Trim();
                status = actualMsg == expectedMsg ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(83, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC28 – Sửa username trùng → "Already exists".</summary>
        [TestMethod]
        public void UM_TC28_EditUser_DuplicateUsername()
        {
            string dupUsername, expectedMsg;
            lock (excelLock)
            {
                using FileStream fs = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read);
                XSSFWorkbook wb = new XSSFWorkbook(fs);
                ISheet sh = wb.GetSheet(SHEET_NAME);
                dupUsername = ReadCell(sh, 84, COL_TESTDATA);
                expectedMsg = ReadCell(sh, 85, COL_EXPECTED);
            }
            if (string.IsNullOrWhiteSpace(dupUsername)) dupUsername = ADMIN_USER;

            string actualMsg = "";
            string status = "Failed";
            try
            {
                GoToUserManagement();
                string target = FindRowByUsername(EDIT_USERNAME) != null ? EDIT_USERNAME : NEW_USERNAME;
                OpenEditForm(target);
                Thread.Sleep(3000);

                IWebElement txtUser = wait.Until(d => d.FindElement(
                    By.XPath("//label[text()='Username']/following::input[1]")));
                txtUser.Click();
                txtUser.SendKeys(Keys.Control + "a");
                txtUser.SendKeys(Keys.Backspace);
                Thread.Sleep(500);
                txtUser.SendKeys(dupUsername);
                txtUser.SendKeys(Keys.Tab);
                Thread.Sleep(1000);

                IWebElement saveBtn = wait.Until(d => d.FindElement(By.XPath("//button[@type='submit']")));
                ((IJavaScriptExecutor)dr).ExecuteScript("arguments[0].click();", saveBtn);

                IWebElement errEl = wait.Until(d => d.FindElement(By.XPath(
                    "//span[contains(@class,'oxd-input-field-error-message') and text()='Already exists']")));

                actualMsg = errEl.Text.Trim();
                status = actualMsg == expectedMsg ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(85, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC29 – Cancel Edit → không lưu thay đổi.</summary>
        [TestMethod]
        public void UM_TC29_EditUser_Cancel()
        {
            string expectedMsg = ReadExpected(87);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                GoToUserManagement();
                string target = FindRowByUsername(EDIT_USERNAME) != null ? EDIT_USERNAME : NEW_USERNAME;

                IWebElement row = FindRowByUsername(target);
                var cells = row.FindElements(By.XPath(".//div[@role='cell']"));
                string oldStatus = cells.Count >= 5 ? cells[4].Text.Trim() : "";

                row.FindElement(By.XPath(".//button[contains(@class,'oxd-icon-button')][2]")).Click();
                wait.Until(d => d.FindElement(By.XPath("//h6[text()='Edit User']")));
                Thread.Sleep(1500);

                IWebElement statusDrop = wait.Until(d => d.FindElement(By.XPath(
                    "//label[text()='Status']/following::div[contains(@class,'oxd-select-wrapper')][1]")));
                SelectOxdOption(statusDrop, oldStatus == "Enabled" ? "Disabled" : "Enabled");

                IWebElement btnCancel = wait.Until(d => d.FindElement(
                    By.XPath("//button[normalize-space()='Cancel']")));
                ((IJavaScriptExecutor)dr).ExecuteScript("arguments[0].click();", btnCancel);
                Thread.Sleep(2000);

                GoToUserManagement();
                IWebElement rowAfter = FindRowByUsername(target);
                string newStatus = "";
                if (rowAfter != null)
                {
                    var ca = rowAfter.FindElements(By.XPath(".//div[@role='cell']"));
                    newStatus = ca.Count >= 5 ? ca[4].Text.Trim() : "";
                }

                bool ok = oldStatus == newStatus;
                actualMsg = ok ? expectedMsg : $"Status thay đổi: {oldStatus} → {newStatus}";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(87, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        // ═══════════════════════════════════════════════════════════════
        // F2.6 – XÓA USER (TC30 – TC34)
        // ═══════════════════════════════════════════════════════════════

        /// <summary>TC30 – Xóa (Disable) user đơn lẻ.</summary>
        [TestMethod]
        public void UM_TC30_DeleteUser_Single()
        {
            string targetUser, expectedMsg;
            lock (excelLock)
            {
                using FileStream fs = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read);
                XSSFWorkbook wb = new XSSFWorkbook(fs);
                ISheet sh = wb.GetSheet(SHEET_NAME);
                targetUser = ReadCell(sh, 89, COL_TESTDATA).Trim();
                expectedMsg = ReadCell(sh, 90, COL_EXPECTED);
            }

            string actualMsg = "";
            string status = "Failed";
            try
            {
                GoToUserManagement();
                IWebElement rowBefore = FindRowByUsername(targetUser);
                Assert.IsNotNull(rowBefore, $"Không tìm thấy user '{targetUser}'");

                rowBefore.FindElement(By.XPath(".//button[contains(@class,'oxd-icon-button')][1]")).Click();
                IWebElement confirmBtn = wait.Until(d => d.FindElement(By.XPath("//button[contains(., 'Yes')]")));
                ((IJavaScriptExecutor)dr).ExecuteScript("arguments[0].click();", confirmBtn);
                Thread.Sleep(3000);

                bool stillExists = FindRowByUsername(targetUser) != null;
                actualMsg = stillExists ? expectedMsg : "LỖI: User đã bị xóa mất (sai nghiệp vụ Disable)";
                status = stillExists ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(90, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC31 – Cancel xóa → user vẫn còn.</summary>
        [TestMethod]
        public void UM_TC31_DeleteUser_Cancel()
        {
            string expectedMsg = ReadExpected(93);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                GoToUserManagement();
                IWebElement targetRow = null;
                foreach (IWebElement r in dr.FindElements(By.XPath("//div[@class='oxd-table-body']//div[@role='row']")))
                {
                    var cells = r.FindElements(By.XPath(".//div[@role='cell']"));
                    if (cells.Count >= 2 && cells[1].Text.Trim() != ADMIN_USER) { targetRow = r; break; }
                }
                Assert.IsNotNull(targetRow, "Không có user khác để test");
                string targetUsername = targetRow.FindElements(By.XPath(".//div[@role='cell']"))[1].Text.Trim();

                targetRow.FindElement(By.XPath(".//button[contains(@class,'oxd-icon-button')][1]")).Click();
                IWebElement cancelBtn = wait.Until(d => d.FindElement(By.XPath(
                    "//button[contains(normalize-space(), 'No, Cancel')]")));
                ((IJavaScriptExecutor)dr).ExecuteScript("arguments[0].click();", cancelBtn);
                Thread.Sleep(1000);

                bool ok = FindRowByUsername(targetUsername) != null;
                actualMsg = ok ? expectedMsg : $"User '{targetUsername}' bị xóa dù Cancel";
                status = ok ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(93, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC32 – Xóa nhiều user cùng lúc (Bulk Delete).</summary>
        [TestMethod]
        public void UM_TC32_DeleteUser_Bulk()
        {
            string expectedMsg = ReadExpected(96);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                GoToUserManagement();
                Thread.Sleep(1500);

                int ticked = 0;
                List<string> targetUsers = new List<string>();
                foreach (IWebElement r in dr.FindElements(By.XPath("//div[@class='oxd-table-body']//div[@role='row']")))
                {
                    var cells = r.FindElements(By.XPath(".//div[@role='cell']"));
                    if (cells.Count >= 2 && cells[1].Text.Trim() != ADMIN_USER)
                    {
                        targetUsers.Add(cells[1].Text.Trim());
                        ((IJavaScriptExecutor)dr).ExecuteScript("arguments[0].click();",
                            r.FindElement(By.XPath(".//label")));
                        ticked++;
                        if (ticked == 2) break;
                    }
                }
                Assert.IsTrue(ticked == 2, $"Không tick đủ 2 checkbox (tick được: {ticked})");
                Thread.Sleep(500);

                IWebElement bulkDeleteBtn = wait.Until(d => d.FindElement(By.XPath(
                    "//button[contains(@class,'oxd-button--label-danger') and contains(normalize-space(), 'Delete Selected')]")));
                ((IJavaScriptExecutor)dr).ExecuteScript("arguments[0].click();", bulkDeleteBtn);

                IWebElement confirmBtn = wait.Until(d => d.FindElement(By.XPath(
                    "//button[contains(@class,'oxd-button--label-danger') and contains(normalize-space(), 'Yes')]")));
                ((IJavaScriptExecutor)dr).ExecuteScript("arguments[0].click();", confirmBtn);
                Thread.Sleep(3000);

                bool allExist = targetUsers.All(u => FindRowByUsername(u) != null);
                actualMsg = allExist ? expectedMsg : "LỖI: Một số user bị xóa mất (sai nghiệp vụ Disable)";
                status = allExist ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(96, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC33 – Select All và Bulk Delete (không xóa admin).</summary>
        [TestMethod]
        public void UM_TC33_DeleteUser_SelectAll()
        {
            string expectedMsg = ReadExpected(99);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                GoToUserManagement();
                Thread.Sleep(1500);

                List<string> sampleUsers = new List<string>();
                foreach (IWebElement r in dr.FindElements(By.XPath("//div[@class='oxd-table-body']//div[@role='row']")))
                {
                    var cells = r.FindElements(By.XPath(".//div[@role='cell']"));
                    if (cells.Count >= 2 && cells[1].Text.Trim() != ADMIN_USER)
                    {
                        sampleUsers.Add(cells[1].Text.Trim());
                        if (sampleUsers.Count == 2) break;
                    }
                }
                Assert.IsTrue(sampleUsers.Count > 0, "Không có user nào (ngoài admin) để test Select All");

                IWebElement headerCb = wait.Until(d => d.FindElement(By.XPath(
                    "//div[@class='oxd-table-header']//label")));
                ((IJavaScriptExecutor)dr).ExecuteScript("arguments[0].click();", headerCb);
                Thread.Sleep(1000);

                // Đảm bảo admin không bị tick
                IWebElement adminRow = FindRowByUsername(ADMIN_USER);
                if (adminRow != null)
                {
                    IWebElement adminCb = adminRow.FindElement(By.XPath(".//input[@type='checkbox']"));
                    Assert.IsFalse(adminCb.Selected, "LỖI BẢO MẬT: Checkbox admin bị tick!");
                }

                IWebElement bulkDeleteBtn = wait.Until(d => d.FindElement(By.XPath(
                    "//button[contains(@class,'oxd-button--label-danger') and contains(normalize-space(), 'Delete Selected')]")));
                ((IJavaScriptExecutor)dr).ExecuteScript("arguments[0].click();", bulkDeleteBtn);

                IWebElement confirmBtn = wait.Until(d => d.FindElement(By.XPath(
                    "//button[contains(@class,'oxd-button--label-danger') and contains(normalize-space(), 'Yes')]")));
                ((IJavaScriptExecutor)dr).ExecuteScript("arguments[0].click();", confirmBtn);
                Thread.Sleep(3000);

                bool allExist = sampleUsers.All(u => FindRowByUsername(u) != null);
                actualMsg = allExist ? expectedMsg : "LỖI: Các user bị xóa sạch (sai nghiệp vụ Disable)";
                status = allExist ? "Passed" : "Failed";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(99, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        /// <summary>TC34 – Admin không thể tự xóa bản thân.</summary>
        [TestMethod]
        public void UM_TC34_DeleteUser_CannotDeleteSelf()
        {
            string expectedMsg = ReadExpected(101);
            string actualMsg = "";
            string status = "Failed";
            try
            {
                GoToUserManagement();
                Thread.Sleep(1500);

                IWebElement adminRow = FindRowByUsername(ADMIN_USER);
                Assert.IsNotNull(adminRow, $"Không tìm thấy '{ADMIN_USER}'");

                IWebElement deleteBtn = adminRow.FindElement(By.XPath(".//button[.//i[contains(@class,'bi-trash')]]"));
                ((IJavaScriptExecutor)dr).ExecuteScript("arguments[0].click();", deleteBtn);

                IWebElement toastMsg = wait.Until(d => d.FindElement(
                    By.XPath("//div[contains(@id,'oxd-toaster')]//p[contains(@class,'oxd-text--toast-message')]")));

                string toastText = ((IJavaScriptExecutor)dr).ExecuteScript(
                    "return arguments[0].textContent;", toastMsg) as string ?? "";
                toastText = toastText.Trim();

                bool ok = toastText.Contains("Cannot be deleted")
                       || expectedMsg.Contains(toastText)
                       || toastText.Contains(expectedMsg);
                actualMsg = ok ? expectedMsg : (string.IsNullOrEmpty(toastText) ? "Toast rỗng!" : toastText);
                status = ok ? "Passed" : "Failed";
            }
            catch (WebDriverTimeoutException)
            {
                actualMsg = "Timeout: Không bắt được Toast.";
            }
            catch (Exception ex) { actualMsg = ex.Message; }

            WriteExcelResult(101, actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }
    }
}