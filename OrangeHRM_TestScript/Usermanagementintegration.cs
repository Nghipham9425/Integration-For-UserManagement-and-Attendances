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
    /// <summary>
    /// Integration tests cho User Management.
    /// Mỗi TC kiểm tra một luồng nghiệp vụ nhiều bước end-to-end.
    /// Sheet Excel: "UserManagementIntegration"
    /// </summary>
    [TestClass]
    public class UserManagementIntegration
    {
        private const string BASE_URL    = "http://localhost:9425/orangehrm-5.6";
        private const string ADMIN_USER  = "nghi45397";
        private const string ADMIN_PASS  = "Nghiphamtrung09042005!";

        // ── Excel ────────────────────────────────────────────────────────
        private string excelFilePath = @"D:\BDCLPM\TestCase_Nhom14.xlsx";
        private const string INT_SHEET = "UserManagementIntegration";

        private const int COL_INT_TESTDATA = 7;   // col H – Test Data
        private const int COL_INT_EXPECTED = 8;   // col I – Expected Result
        private const int COL_INT_ACTUAL   = 9;   // col J – Actual Result
        private const int COL_INT_STATUS   = 11;  // col L – Result

        // Row indices (0-based) cho từng TC
        private static readonly int[] TC01_ROWS = { 3, 4, 5, 6, 7, 8, 9, 10, 11 };
        private static readonly int[] TC02_ROWS = { 13, 14, 15, 16, 17, 18, 19 };

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

        private string ReadIntCell(ISheet sheet, int rowIdx, int colIdx)
        {
            lock (excelLock)
            {
                var fmt = new DataFormatter();
                IRow r = sheet.GetRow(rowIdx);
                if (r == null) return "";
                return fmt.FormatCellValue(r.GetCell(colIdx)) ?? "";
            }
        }

        private void WriteIntResult(int rowIdx, string actualMsg, string status)
        {
            lock (excelLock)
            {
                using FileStream fsRead = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read);
                XSSFWorkbook wb = new XSSFWorkbook(fsRead);
                ISheet sh = wb.GetSheet(INT_SHEET);
                IRow row = sh.GetRow(rowIdx) ?? sh.CreateRow(rowIdx);
                row.CreateCell(COL_INT_ACTUAL).SetCellValue(actualMsg);
                row.CreateCell(COL_INT_STATUS).SetCellValue(status);
                using FileStream fsWrite = new FileStream(excelFilePath, FileMode.Create, FileAccess.Write);
                wb.Write(fsWrite);
            }
        }

        /// <summary>Parse chuỗi "key: value\n..." thành Dictionary.</summary>
        private Dictionary<string, string> ParseTestData(string raw)
        {
            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrWhiteSpace(raw)) return dict;
            foreach (string line in raw.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries))
            {
                int idx = line.IndexOf(':');
                if (idx < 0) continue;
                dict[line.Substring(0, idx).Trim()] = line.Substring(idx + 1).Trim();
            }
            return dict;
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
                dr.Navigate().GoToUrl(BASE_URL + "/web/index.php/auth/logout");
                Thread.Sleep(1000);
            }
        }

        private void GoToUserManagement()
        {
            try
            {
                dr.FindElement(By.XPath("//span[text()='ADMIN'] | //a[normalize-space()='Admin']")).Click();
                wait.Until(d => d.Url.Contains("viewSystemUsers") || d.Url.Contains("admin"));
            }
            catch
            {
                dr.Navigate().GoToUrl(BASE_URL + "/web/index.php/admin/viewSystemUsers");
                wait.Until(d => d.Url.Contains("admin"));
            }
            Thread.Sleep(800);
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

        private void OpenAddUserForm()
        {
            GoToUserManagement();
            wait.Until(d => d.FindElement(By.XPath("//button[normalize-space()='Add']"))).Click();
            wait.Until(d => d.FindElement(By.XPath("//h6[text()='Add User']")));
        }

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

        private void OpenEditForm(string username)
        {
            IWebElement row = FindRowByUsername(username);
            Assert.IsNotNull(row, $"Không tìm thấy user '{username}' để Edit");
            row.FindElement(By.XPath(".//button[contains(@class,'oxd-icon-button')][2]")).Click();
            wait.Until(d => d.FindElement(By.XPath("//h6[text()='Edit User']")));
        }

        // ═══════════════════════════════════════════════════════════════
        // UM_INT_TC01
        // Luồng: Login không tồn tại → Admin tạo → User login OK
        //        → Admin disable → User login bị từ chối
        // ═══════════════════════════════════════════════════════════════

        [TestMethod]
        public void UM_INT_TC01_CreateLoginDisable_Flow()
        {
            // ── Đọc TestData & Expected từ Excel ─────────────────────────
            string actualMsg = "";
            string status = "Failed";

            string s1User, s1Pass, adminUser, adminPass;
            string newUser, newPass, newRole, newEmp, newStatus;
            string s5User, s5Pass, s7AdminUser, s7AdminPass, disableStatus;
            string s9User, s9Pass;
            string exp1, exp3, exp5, exp7, exp9;

            lock (excelLock)
            {
                using FileStream fs = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read);
                XSSFWorkbook wb = new XSSFWorkbook(fs);
                ISheet sheet = wb.GetSheet(INT_SHEET);

                var td1 = ParseTestData(ReadIntCell(sheet, TC01_ROWS[0], COL_INT_TESTDATA));
                var td2 = ParseTestData(ReadIntCell(sheet, TC01_ROWS[1], COL_INT_TESTDATA));
                var td3 = ParseTestData(ReadIntCell(sheet, TC01_ROWS[2], COL_INT_TESTDATA));
                var td5 = ParseTestData(ReadIntCell(sheet, TC01_ROWS[4], COL_INT_TESTDATA));
                var td7 = ParseTestData(ReadIntCell(sheet, TC01_ROWS[6], COL_INT_TESTDATA));
                var td9 = ParseTestData(ReadIntCell(sheet, TC01_ROWS[8], COL_INT_TESTDATA));

                exp1 = ReadIntCell(sheet, TC01_ROWS[0], COL_INT_EXPECTED);
                exp3 = ReadIntCell(sheet, TC01_ROWS[2], COL_INT_EXPECTED);
                exp5 = ReadIntCell(sheet, TC01_ROWS[4], COL_INT_EXPECTED);
                exp7 = ReadIntCell(sheet, TC01_ROWS[6], COL_INT_EXPECTED);
                exp9 = ReadIntCell(sheet, TC01_ROWS[8], COL_INT_EXPECTED);

                s1User = td1.GetValueOrDefault("username", "newuser01");
                s1Pass = td1.GetValueOrDefault("password", "AnyPass@1");

                adminUser = td2.GetValueOrDefault("username", ADMIN_USER);
                adminPass = td2.GetValueOrDefault("password", ADMIN_PASS);

                newUser    = td3.GetValueOrDefault("Username", s1User);
                newPass    = td3.GetValueOrDefault("Password", "NewPass@123!");
                newRole    = td3.GetValueOrDefault("Role", "ESS");
                newEmp     = td3.GetValueOrDefault("Employee", "An Văn");
                newStatus  = td3.GetValueOrDefault("Status", "Enabled");

                s5User = td5.GetValueOrDefault("username", newUser);
                s5Pass = td5.GetValueOrDefault("password", newPass);

                s7AdminUser  = td7.GetValueOrDefault("username", adminUser);
                s7AdminPass  = td7.GetValueOrDefault("password", adminPass);
                disableStatus = td7.GetValueOrDefault("Status", "Disabled");

                s9User = td9.GetValueOrDefault("username", newUser);
                s9Pass = td9.GetValueOrDefault("password", newPass);
            }

            try
            {
                // ── Step 1: Login với tài khoản chưa tồn tại ─────────────
                dr.Navigate().GoToUrl(BASE_URL);
                wait.Until(d => d.FindElement(By.Name("username"))).SendKeys(s1User);
                dr.FindElement(By.Name("password")).SendKeys(s1Pass);
                dr.FindElement(By.CssSelector("button[type='submit']")).Click();

                IWebElement errEl = wait.Until(d => d.FindElement(By.XPath(
                    "//*[contains(@class,'oxd-alert-content-text')" +
                    " or contains(@class,'oxd-input-field-error-message')]")));
                Assert.IsTrue(errEl.Displayed && !dr.Url.Contains("dashboard"),
                    $"[Step 1 FAIL] Expected: {exp1}\nActual: Hệ thống vào dashboard dù tài khoản chưa tồn tại");

                // ── Step 2: Admin đăng nhập ───────────────────────────────
                LoginAs(adminUser, adminPass);
                Assert.IsTrue(dr.Url.Contains("dashboard") ||
                              dr.FindElements(By.ClassName("oxd-topbar-header-breadcrumb")).Count > 0,
                    $"[Step 2 FAIL] Admin '{adminUser}' đăng nhập không thành công");

                // ── Step 3: Admin tạo tài khoản mới ──────────────────────
                OpenAddUserForm();

                IWebElement roleDrop = wait.Until(d => d.FindElement(By.XPath(
                    "//label[text()='User Role']/following::div[contains(@class,'oxd-select-wrapper')][1]")));
                SelectOxdOption(roleDrop, newRole);

                string empHint = newEmp.Contains(" ")
                    ? newEmp.Substring(0, newEmp.LastIndexOf(' ')) : newEmp;
                FillEmployeeName(empHint);

                IWebElement statusDrop = dr.FindElement(By.XPath(
                    "//label[text()='Status']/following::div[contains(@class,'oxd-select-wrapper')][1]"));
                SelectOxdOption(statusDrop, newStatus);

                IWebElement txtUser = dr.FindElement(By.XPath("//label[text()='Username']/following::input[1]"));
                txtUser.Clear(); txtUser.SendKeys(newUser);
                IWebElement txtPass = dr.FindElement(By.XPath("//label[text()='Password']/following::input[1]"));
                txtPass.Clear(); txtPass.SendKeys(newPass);
                IWebElement txtConfirm = dr.FindElement(By.XPath("//label[text()='Confirm Password']/following::input[1]"));
                txtConfirm.Clear(); txtConfirm.SendKeys(newPass);

                dr.FindElement(By.XPath("//button[@type='submit' and contains(.,'Save')]")).Click();

                IWebElement toastSave = wait.Until(d => d.FindElement(By.XPath(
                    "//p[contains(@class,'oxd-text--toast-message')]")));
                string toastSaveText = toastSave.Text.Trim();

                GoToUserManagement();
                bool userCreated = FindRowByUsername(newUser) != null;
                Assert.IsTrue(toastSaveText.Contains("Successfully Saved") && userCreated,
                    $"[Step 3 FAIL] Expected: {exp3}\ntoast='{toastSaveText}', userFound={userCreated}");

                // ── Step 4: Admin đăng xuất ───────────────────────────────
                Logout();
                Assert.IsTrue(dr.FindElements(By.Name("username")).Count > 0,
                    "[Step 4 FAIL] Admin đăng xuất không thành công");

                // ── Step 5: Đăng nhập bằng tài khoản vừa tạo ─────────────
                LoginAs(s5User, s5Pass);
                Assert.IsTrue(dr.Url.Contains("dashboard") ||
                              dr.FindElements(By.ClassName("oxd-topbar-header-breadcrumb")).Count > 0,
                    $"[Step 5 FAIL] Expected: {exp5}\n'{s5User}' không đăng nhập được");

                // ── Step 6: Đăng xuất tài khoản vừa tạo ──────────────────
                Logout();
                Assert.IsTrue(dr.FindElements(By.Name("username")).Count > 0,
                    $"[Step 6 FAIL] '{s5User}' đăng xuất không thành công");

                // ── Step 7: Admin đăng nhập lại → Disable tài khoản ──────
                LoginAs(s7AdminUser, s7AdminPass);
                GoToUserManagement();
                OpenEditForm(newUser);

                IWebElement editStatusDrop = wait.Until(d => d.FindElement(By.XPath(
                    "//label[text()='Status']/following::div[contains(@class,'oxd-select-wrapper')][1]")));
                SelectOxdOption(editStatusDrop, disableStatus);

                IWebElement saveBtn = dr.FindElement(By.XPath("//button[@type='submit' and contains(.,'Save')]"));
                ((IJavaScriptExecutor)dr).ExecuteScript("arguments[0].scrollIntoView(true);", saveBtn);
                Thread.Sleep(400);
                ((IJavaScriptExecutor)dr).ExecuteScript("arguments[0].click();", saveBtn);

                IWebElement toastEdit = wait.Until(d => d.FindElement(By.XPath(
                    "//p[contains(@class,'oxd-text--toast-message')]")));
                string toastEditText = toastEdit.Text.Trim();

                GoToUserManagement();
                IWebElement disabledRow = FindRowByUsername(newUser);
                bool statusIsDisabled = false;
                if (disabledRow != null)
                {
                    var cells = disabledRow.FindElements(By.XPath(".//div[@role='cell']"));
                    statusIsDisabled = cells.Count >= 5 && cells[4].Text.Trim() == disableStatus;
                }
                Assert.IsTrue(toastEditText.Contains("Successfully Updated") && statusIsDisabled,
                    $"[Step 7 FAIL] Expected: {exp7}\ntoast='{toastEditText}', statusDisabled={statusIsDisabled}");

                // ── Step 8: Admin đăng xuất ───────────────────────────────
                Logout();
                Assert.IsTrue(dr.FindElements(By.Name("username")).Count > 0,
                    "[Step 8 FAIL] Admin đăng xuất lần 2 không thành công");

                // ── Step 9: Login với tài khoản bị Disable → báo lỗi ─────
                dr.FindElement(By.Name("username")).SendKeys(s9User);
                dr.FindElement(By.Name("password")).SendKeys(s9Pass);
                dr.FindElement(By.CssSelector("button[type='submit']")).Click();

                IWebElement loginErrEl = wait.Until(d => d.FindElement(By.XPath(
                    "//*[contains(@class,'oxd-alert-content-text')" +
                    " or contains(@class,'oxd-input-field-error-message')]")));
                string step9ActualMsg = loginErrEl.Text.Trim();

                bool step9Ok = loginErrEl.Displayed
                               && !dr.Url.Contains("dashboard")
                               && step9ActualMsg == exp9.Trim();
                Assert.IsTrue(step9Ok,
                    $"[Step 9 FAIL] Expected: '{exp9}'\nActual: '{step9ActualMsg}'");

                actualMsg = step9ActualMsg;
                status = "Passed";
            }
            catch (AssertFailedException afe) { actualMsg = afe.Message; status = "Failed"; }
            catch (Exception ex)              { actualMsg = ex.Message;  status = "Failed"; }

            WriteIntResult(TC01_ROWS[0], actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }

        // ═══════════════════════════════════════════════════════════════
        // UM_INT_TC02
        // Luồng: Admin thêm user (ESS) → Search xác nhận → Edit đổi Role → Admin → Search verify
        // ═══════════════════════════════════════════════════════════════

        [TestMethod]
        public void UM_INT_TC02_AddUser_EditRole_SearchVerify_Flow()
        {
            string actualMsg = "";
            string status = "Failed";

            string adminUser, adminPass;
            string newUser, newPass, newRole, empHint, newStatus, searchUser, newRoleAfterEdit;
            string exp1, exp2, exp3, exp4, exp5, exp6, exp7;

            lock (excelLock)
            {
                using FileStream fs = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read);
                XSSFWorkbook wb = new XSSFWorkbook(fs);
                ISheet sheet = wb.GetSheet(INT_SHEET);

                var td1 = ParseTestData(ReadIntCell(sheet, TC02_ROWS[0], COL_INT_TESTDATA));
                var td2 = ParseTestData(ReadIntCell(sheet, TC02_ROWS[1], COL_INT_TESTDATA));
                var td3 = ParseTestData(ReadIntCell(sheet, TC02_ROWS[2], COL_INT_TESTDATA));
                var td6 = ParseTestData(ReadIntCell(sheet, TC02_ROWS[5], COL_INT_TESTDATA));

                exp1 = ReadIntCell(sheet, TC02_ROWS[0], COL_INT_EXPECTED);
                exp2 = ReadIntCell(sheet, TC02_ROWS[1], COL_INT_EXPECTED);
                exp3 = ReadIntCell(sheet, TC02_ROWS[2], COL_INT_EXPECTED);
                exp4 = ReadIntCell(sheet, TC02_ROWS[3], COL_INT_EXPECTED);
                exp5 = ReadIntCell(sheet, TC02_ROWS[4], COL_INT_EXPECTED);
                exp6 = ReadIntCell(sheet, TC02_ROWS[5], COL_INT_EXPECTED);
                exp7 = ReadIntCell(sheet, TC02_ROWS[6], COL_INT_EXPECTED);

                adminUser = td1.GetValueOrDefault("username", ADMIN_USER);
                adminPass = td1.GetValueOrDefault("password", ADMIN_PASS);

                newUser   = td2.GetValueOrDefault("Username", "edituser01");
                newPass   = td2.GetValueOrDefault("Password", "EditUser@01!");
                newRole   = td2.GetValueOrDefault("Role", "ESS");
                empHint   = td2.GetValueOrDefault("Employee", "An Văn");
                newStatus = td2.GetValueOrDefault("Status", "Enabled");

                searchUser       = td3.GetValueOrDefault("Username search", newUser);
                newRoleAfterEdit = td6.GetValueOrDefault("User Role filter", "Admin");
            }

            try
            {
                // ── Step 1: Admin đăng nhập, ghi Records Found ───────────
                LoginAs(adminUser, adminPass);
                GoToUserManagement();
                int recBefore = GetRecordsFound();
                Assert.IsTrue(recBefore >= 0, $"[Step 1 FAIL] {exp1}\nKhông lấy được Records Found");

                // ── Step 2: Thêm user mới với Role ESS ───────────────────
                wait.Until(d => d.FindElement(By.XPath("//button[normalize-space()='Add']"))).Click();
                wait.Until(d => d.FindElement(By.XPath("//h6[text()='Add User']")));

                IWebElement roleDrop = wait.Until(d => d.FindElement(By.XPath(
                    "//label[text()='User Role']/following::div[contains(@class,'oxd-select-wrapper')][1]")));
                SelectOxdOption(roleDrop, newRole);

                FillEmployeeName(empHint.Contains(" ") ? empHint.Substring(0, empHint.LastIndexOf(' ')) : empHint);

                SelectOxdOption(dr.FindElement(By.XPath(
                    "//label[text()='Status']/following::div[contains(@class,'oxd-select-wrapper')][1]")), newStatus);

                IWebElement txtUser = dr.FindElement(By.XPath("//label[text()='Username']/following::input[1]"));
                txtUser.Clear(); txtUser.SendKeys(newUser);
                IWebElement txtPass = dr.FindElement(By.XPath("//label[text()='Password']/following::input[1]"));
                txtPass.Clear(); txtPass.SendKeys(newPass);
                dr.FindElement(By.XPath("//label[text()='Confirm Password']/following::input[1]")).SendKeys(newPass);
                dr.FindElement(By.XPath("//button[@type='submit' and contains(.,'Save')]")).Click();
                Thread.Sleep(1500);

                GoToUserManagement();
                int recAfterAdd = GetRecordsFound();
                bool userCreated = FindRowByUsername(newUser) != null;
                Assert.IsTrue(userCreated && recAfterAdd == recBefore + 1,
                    $"[Step 2 FAIL] {exp2}\nfound={userCreated}, records {recBefore}→{recAfterAdd}");

                // ── Step 3: Search xác nhận user mới ─────────────────────
                IWebElement searchInp = wait.Until(d => d.FindElement(By.XPath(
                    "//label[text()='Username']/following::input[1]")));
                searchInp.Clear(); searchInp.SendKeys(searchUser);
                dr.FindElement(By.XPath("//button[contains(.,'Search')]")).Click();
                Thread.Sleep(1500);

                var rows3 = dr.FindElements(By.XPath("//div[@class='oxd-table-body']//div[@role='row']"));
                IWebElement foundRow3 = FindRowByUsername(searchUser);
                Assert.IsTrue(rows3.Count == 1 && foundRow3 != null,
                    $"[Step 3 FAIL] {exp3}\nrows={rows3.Count}, found={foundRow3 != null}");

                var cells3 = foundRow3.FindElements(By.XPath(".//div[@role='cell']"));
                Assert.IsTrue(cells3.Count >= 3 && cells3[2].Text.Trim() == "ESS",
                    $"[Step 3 FAIL] Role mong đợi ESS, thực tế: {(cells3.Count >= 3 ? cells3[2].Text.Trim() : "N/A")}");

                // ── Step 4: Mở Edit form từ kết quả search ───────────────
                IWebElement editBtn = foundRow3.FindElement(By.XPath(
                    ".//button[.//i[contains(@class,'bi-pencil')] or contains(@class,'oxd-icon-button')][2]"));
                editBtn.Click();
                wait.Until(d => d.FindElement(By.XPath("//h6[text()='Edit User']")));
                Assert.IsTrue(true, $"[Step 4] {exp4}"); // Mở Edit thành công

                // ── Step 5: Đổi Role ESS → Admin, nhấn Save ──────────────
                IWebElement editRoleDrop = wait.Until(d => d.FindElement(By.XPath(
                    "//label[text()='User Role']/following::div[contains(@class,'oxd-select-wrapper')][1]")));
                SelectOxdOption(editRoleDrop, newRoleAfterEdit);

                IWebElement saveBtn = dr.FindElement(By.XPath("//button[@type='submit' and contains(.,'Save')]"));
                ((IJavaScriptExecutor)dr).ExecuteScript("arguments[0].scrollIntoView(true);", saveBtn);
                Thread.Sleep(400);
                ((IJavaScriptExecutor)dr).ExecuteScript("arguments[0].click();", saveBtn);
                Thread.Sleep(1500);

                GoToUserManagement();

                // ── Step 6: Search lại, xác nhận Role = Admin ────────────
                IWebElement searchInp2 = wait.Until(d => d.FindElement(By.XPath(
                    "//label[text()='Username']/following::input[1]")));
                searchInp2.Clear(); searchInp2.SendKeys(searchUser);
                dr.FindElement(By.XPath("//button[contains(.,'Search')]")).Click();
                Thread.Sleep(1500);

                IWebElement updatedRow = FindRowByUsername(searchUser);
                Assert.IsNotNull(updatedRow, $"[Step 6 FAIL] {exp6}\nKhông tìm thấy user sau Edit");

                var cells6 = updatedRow.FindElements(By.XPath(".//div[@role='cell']"));
                string roleAfterEdit = cells6.Count >= 3 ? cells6[2].Text.Trim() : "";
                Assert.IsTrue(roleAfterEdit == newRoleAfterEdit,
                    $"[Step 6 FAIL] {exp6}\nRole mong đợi '{newRoleAfterEdit}', thực tế: '{roleAfterEdit}'");

                // ── Step 7: Filter danh sách theo Role = Admin ────────────
                dr.FindElement(By.XPath("//button[contains(.,'Reset')]")).Click();
                Thread.Sleep(800);

                IWebElement roleFilter = wait.Until(d => d.FindElement(By.XPath(
                    "//label[text()='User Role']/ancestor::div[contains(@class,'oxd-input-group')]" +
                    "//div[contains(@class,'oxd-select-wrapper')]")));
                SelectOxdOption(roleFilter, newRoleAfterEdit);
                dr.FindElement(By.XPath("//button[contains(.,'Search')]")).Click();
                Thread.Sleep(1500);

                bool inAdminList = FindRowByUsername(newUser) != null;
                Assert.IsTrue(inAdminList,
                    $"[Step 7 FAIL] {exp7}\n'{newUser}' không xuất hiện khi filter Role={newRoleAfterEdit}");

                actualMsg = $"Luồng INT_TC02 PASS: Add user ({newRole}) → Search xác nhận → Edit Role→{newRoleAfterEdit} → Search verify → Filter {newRoleAfterEdit}";
                status = "Passed";
            }
            catch (AssertFailedException afe) { actualMsg = afe.Message; status = "Failed"; }
            catch (Exception ex)              { actualMsg = ex.Message;  status = "Failed"; }

            WriteIntResult(TC02_ROWS[0], actualMsg, status);
            Assert.AreEqual("Passed", status, actualMsg);
        }
    }
}