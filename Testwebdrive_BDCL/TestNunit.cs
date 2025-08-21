using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.IO;
using ExcelDataReader;
using System.Data;
using System.Text;
using OpenQA.Selenium.Interactions;

namespace WebTrangSuc.Tests
{
    [TestFixture]
    public class Testcasequanlyadmin
    {
        private IWebDriver driver;
        private WebDriverWait wait;
        private string baseUrl = "http://localhost:55119";

        [SetUp]
        public void Setup()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            driver = new ChromeDriver();
            driver.Manage().Window.Maximize();
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1.5);
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(1.5));
        }
        private class TaiKhoanTestData
        {
            public string HoVaTen { get; set; }
            public string GioiTinh { get; set; }
            public string NamSinh { get; set; }
            public string SDT { get; set; }
            public string Email { get; set; }
            public string UserName { get; set; }
            public string Matkhau { get; set; }
            public string Avatar { get; set; }
            public string IDRole { get; set; }
            public string SearchKeyword { get; set; }
        }
        private TaiKhoanTestData GetTaiKhoanTestDataFromExcel(string testCaseId, string sheetName)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            string filePath = @"D:\TestData.xlsx";
            TaiKhoanTestData testData = new TaiKhoanTestData();

            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"Không tìm thấy file Excel tại đường dẫn: {filePath}.");
            }

            try
            {
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                        });
                        DataTable table = result.Tables[sheetName];
                        DataRow targetRow = table.AsEnumerable().FirstOrDefault(row => row["TestCaseId"].ToString() == testCaseId);

                        if (targetRow == null)
                        {
                            throw new Exception($"Không tìm thấy dữ liệu cho TestCaseId: {testCaseId}");
                        }

                        // Ánh xạ dữ liệu từ Excel vào đối tượng
                        testData.HoVaTen = targetRow["HoVaTen"].ToString();
                        testData.GioiTinh = targetRow["GioiTinh"].ToString();
                        testData.NamSinh = targetRow["NamSinh"].ToString();
                        testData.SDT = targetRow["SDT"].ToString();
                        testData.Email = targetRow["Email"].ToString();
                        testData.UserName = targetRow["UserName"].ToString();
                        testData.Matkhau = targetRow["Matkhau"].ToString();
                        testData.Avatar = targetRow["Avatar"].ToString();
                        testData.IDRole = targetRow["IDRole"].ToString();
                        testData.SearchKeyword = targetRow["SearchKeyword"].ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi khi đọc file Excel: {ex.Message}");
                throw;
            }

            return testData;
        }
        private class CategoryTestData
        {
            public string TenLoaiSanPham { get; set; }
            public string ImagePath { get; set; }
        }
        private CategoryTestData GetCategoryTestDataFromExcel(string testCaseId, string sheetName)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            string filePath = @"D:\TestData.xlsx";
            CategoryTestData testData = new CategoryTestData();

            Console.WriteLine($"Đang kiểm tra file tại: {filePath}");
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"Không tìm thấy file Excel tại đường dẫn: {filePath}. Vui lòng kiểm tra lại.");
            }

            try
            {
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    Console.WriteLine("Đã mở file thành công.");
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        Console.WriteLine("Đã tạo reader thành công.");
                        var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                        });
                        // Tìm sheet theo tên
                        DataTable table = null;
                        foreach (DataTable dt in result.Tables)
                        {
                            if (dt.TableName == sheetName)
                            {
                                table = dt;
                                break;
                            }
                        }
                        // Tìm hàng tương ứng với TestCaseId
                        DataRow targetRow = null;
                        foreach (DataRow row in table.Rows)
                        {
                            if (row["TestCaseId"].ToString() == testCaseId)
                            {
                                targetRow = row;
                                break;
                            }
                        }

                        if (targetRow == null)
                        {
                            throw new Exception($"Không tìm thấy dữ liệu cho TestCaseId: {testCaseId}");
                        }

                        // Ánh xạ dữ liệu từ hàng tìm được
                        testData.TenLoaiSanPham = targetRow["TenLoaiSanPham"].ToString();
                        testData.ImagePath = targetRow["ImagePath"].ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi khi đọc file Excel: {ex.Message}");
                throw;
            }

            return testData;
        }
        private class ThuongHieuTestData
        {
            public string TenThuongHieu { get; set; }
        }
        private ThuongHieuTestData GetThuongHieuTestDataFromExcel(string testCaseId, string sheetName)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            string filePath = @"D:\TestData.xlsx";
            ThuongHieuTestData testData = new ThuongHieuTestData();

            Console.WriteLine($"Đang kiểm tra file tại: {filePath}");
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"Không tìm thấy file Excel tại đường dẫn: {filePath}. Vui lòng kiểm tra lại.");
            }

            try
            {
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    Console.WriteLine("Đã mở file thành công.");
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        Console.WriteLine("Đã tạo reader thành công.");
                        var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                        });
                        // Tìm sheet theo tên
                        DataTable table = null;
                        foreach (DataTable dt in result.Tables)
                        {
                            if (dt.TableName == sheetName)
                            {
                                table = dt;
                                break;
                            }
                        }
                        // Tìm hàng tương ứng với TestCaseId
                        DataRow targetRow = null;
                        foreach (DataRow row in table.Rows)
                        {
                            if (row["TestCaseId"].ToString() == testCaseId)
                            {
                                targetRow = row;
                                break;
                            }
                        }

                        if (targetRow == null)
                        {
                            throw new Exception($"Không tìm thấy dữ liệu cho TestCaseId: {testCaseId}");
                        }

                        // Ánh xạ dữ liệu từ hàng tìm được
                        testData.TenThuongHieu = targetRow["TenThuongHieu"].ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi khi đọc file Excel: {ex.Message}");
                throw;
            }

            return testData;
        }
        private class ProductTestData
        {
            public string TenSanPham { get; set; }
            public string Gia { get; set; }
            public string MoTaSanPham { get; set; }
            public string SoLuongTonKho { get; set; }
            public string IDLoaiSanPham { get; set; }
            public string IDMauSac { get; set; }
            public string IDChatLieu { get; set; }
            public string IDThuongHieu { get; set; }
            public string TrangThaiSanPham { get; set; }
            public string ImagePath { get; set; }
            public string SearchKeyword { get; set; }
        }
        private ProductTestData GetProductTestDataFromExcel(string testCaseId, string sheetName)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            string filePath = @"D:\TestData.xlsx";
            ProductTestData testData = new ProductTestData();

            Console.WriteLine($"Đang kiểm tra file tại: {filePath}");
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"Không tìm thấy file Excel tại đường dẫn: {filePath}. Vui lòng kiểm tra lại.");
            }

            try
            {
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    Console.WriteLine("Đã mở file thành công.");
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        Console.WriteLine("Đã tạo reader thành công.");
                        var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                        });
                        // Tìm sheet theo tên
                        DataTable table = null;
                        foreach (DataTable dt in result.Tables)
                        {
                            if (dt.TableName == sheetName)
                            {
                                table = dt;
                                break;
                            }
                        }
                        // Tìm hàng tương ứng với TestCaseId
                        DataRow targetRow = null;
                        foreach (DataRow row in table.Rows)
                        {
                            if (row["TestCaseId"].ToString() == testCaseId)
                            {
                                targetRow = row;
                                break;
                            }
                        }

                        if (targetRow == null)
                        {
                            throw new Exception($"Không tìm thấy dữ liệu cho TestCaseId: {testCaseId}");
                        }

                        // Ánh xạ dữ liệu từ hàng tìm được
                        testData.TenSanPham = targetRow["TenSanPham"].ToString();
                        testData.Gia = targetRow["Gia"].ToString();
                        testData.MoTaSanPham = targetRow["MoTaSanPham"].ToString();
                        testData.SoLuongTonKho = targetRow["SoLuongTonKho"].ToString();
                        testData.IDLoaiSanPham = targetRow["IDLoaiSanPham"].ToString();
                        testData.IDMauSac = targetRow["IDMauSac"].ToString();
                        testData.IDChatLieu = targetRow["IDChatLieu"].ToString();
                        testData.IDThuongHieu = targetRow["IDThuongHieu"].ToString();
                        testData.TrangThaiSanPham = targetRow["TrangThaiSanPham"].ToString();
                        testData.ImagePath = targetRow["ImagePath"].ToString();
                        testData.SearchKeyword = targetRow["SearchKeyword"].ToString();

                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi khi đọc file Excel: {ex.Message}");
                throw;
            }

            return testData;
        }
        private void Login(string username, string password)
        {
            driver.Navigate().GoToUrl("http://localhost:55119/Auth");
            driver.FindElement(By.CssSelector("#email")).SendKeys(username);
            driver.FindElement(By.CssSelector("#password")).SendKeys(password);
            driver.FindElement(By.CssSelector("#dangnhap")).Click();
            Thread.Sleep(3000);
        }
        //---------------------------------Các test Quản lý tài khoản----------------------------------------------------------------------
        // testcase QLTK1_01 thêm tài khoản mới hợp lệ
        [Test]
        public void Test01_CreateNewAccountValid()
        {
            Login("tuan159", "Tuan@159159");
            wait.Until(d => d.FindElement(By.Id("userMenuButton")).Displayed);
            driver.FindElement(By.Id("userMenuButton")).Click();
            driver.FindElement(By.Id("manageLink")).Click();
            driver.FindElement(By.LinkText("Quản lý tài khoản")).Click();
            Thread.Sleep(1000);
            IWebElement createButton = wait.Until(d => d.FindElement(By.CssSelector("a.btn.btn-primary[href='/Admin/TaiKhoan/Create']")));
            createButton.Click();
            Thread.Sleep(1000);

            // Lấy dữ liệu TestCaseId = "QLTK1_01"
            TaiKhoanTestData testData = GetTaiKhoanTestDataFromExcel("QLTK1_01", "QLTK");
            driver.FindElement(By.Id("HoVaTen")).SendKeys(testData.HoVaTen);
            new SelectElement(driver.FindElement(By.Id("GioiTinh"))).SelectByText(testData.GioiTinh);
            driver.FindElement(By.Id("NamSinh")).SendKeys(testData.NamSinh);
            driver.FindElement(By.Id("SDT")).SendKeys(testData.SDT);
            driver.FindElement(By.Id("Email")).SendKeys(testData.Email);
            driver.FindElement(By.Id("UserName")).SendKeys(testData.UserName);
            driver.FindElement(By.Id("Matkhau")).SendKeys(testData.Matkhau);
            IWebElement avatarElement = wait.Until(d => d.FindElement(By.Name("Avatar")));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", avatarElement);
            Thread.Sleep(500);
            avatarElement.SendKeys(testData.Avatar);
            new SelectElement(driver.FindElement(By.Id("IDRole"))).SelectByValue(testData.IDRole);

            IWebElement submitButton = driver.FindElement(By.CssSelector("input[type='submit'].btn.btn-success"));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", submitButton);
            Thread.Sleep(500);
            submitButton.Click();
            Thread.Sleep(1000);

            bool found = false;
            while (!found)
            {
                // Tìm tài khoản trên trang hiện tại
                var accountElements = driver.FindElements(By.XPath($"//td[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{testData.HoVaTen.ToLower()}')]"));
                if (accountElements.Count > 0 && accountElements[0].Displayed)
                {
                    IWebElement accountName = accountElements[0];
                    Assert.That(accountName.Displayed, Is.True, "Tài khoản mới không hiển thị trong danh sách!");
                    Console.WriteLine($"Tài khoản '{testData.HoVaTen}' đã được thêm thành công!");
                    found = true; // Tìm thấy tài khoản thoát vòng lặp
                }
                else
                {
                    // Không tìm thấy tài khoản Kiểm tra xem có nút "Trang tiếp theo" không
                    var nextPageButtons = driver.FindElements(By.XPath("//li[@class='PagedList-skipToNext']/a[@rel='next']"));
                    if (nextPageButtons.Count > 0 && nextPageButtons[0].Displayed)
                    {
                        Console.WriteLine("Không tìm thấy tài khoản trên trang hiện tại, chuyển sang trang tiếp theo...");
                        nextPageButtons[0].Click();
                        Thread.Sleep(1000);
                        Console.WriteLine($"URL sau khi chuyển trang: {driver.Url}");
                    }
                    else
                    {
                        // Không còn trang tiếp theo để chuyển
                        var tableRows = driver.FindElements(By.XPath("//table//tr"));
                        Console.WriteLine("Nội dung bảng danh sách tài khoản:");
                        foreach (var row in tableRows)
                        {
                            Console.WriteLine($"Row content: {row.Text}");
                        }
                        Assert.Fail($"Không tìm thấy tài khoản '{testData.HoVaTen}' trong danh sách sau khi duyệt qua tất cả các trang!");
                        break;
                    }
                }
            }
        }
        //Test Đăng nhập với tài khoản vừa tạo
        [Test]
        public void Test02_Combined_CreateAndLoginNewAccount()
        {
            TaiKhoanTestData testData = GetTaiKhoanTestDataFromExcel("QLTK1_01", "QLTK");
            Login($"{testData.UserName}", $"{testData.Matkhau}");
            wait.Until(d => d.FindElement(By.Id("userMenuButton")).Displayed);
            driver.FindElement(By.Id("userMenuButton")).Click();
            driver.FindElement(By.XPath("//*[@id=\"userMenu\"]/div/a[1]")).Click();
            wait.Until(d => d.FindElement(By.Id("Hovaten")).Displayed);
            string actualFullName = driver.FindElement(By.Id("Hovaten")).Text;
            Assert.That(actualFullName, Is.EqualTo(testData.HoVaTen), $"Họ tên hiển thị '{actualFullName}' không khớp với dữ liệu mong muốn '{testData.HoVaTen}'");
        }
    
        // testcase QLTK1_02 thêm tài khoản mới mật khẩu dưới 6 kí tự
        [Test]
        public void Test02_CreateNewAccount_mksai_unValid()
        {
            Login("tuan159", "Tuan@159159");
            wait.Until(d => d.FindElement(By.Id("userMenuButton")).Displayed);
            driver.FindElement(By.Id("userMenuButton")).Click();
            driver.FindElement(By.Id("manageLink")).Click();
            driver.FindElement(By.LinkText("Quản lý tài khoản")).Click();
            Thread.Sleep(1000);
            IWebElement createButton = wait.Until(d => d.FindElement(By.CssSelector("a.btn.btn-primary[href='/Admin/TaiKhoan/Create']")));
            createButton.Click();
            Thread.Sleep(1000);

            // Lấy dữ liệu TestCaseId = "QLTK1_02
            TaiKhoanTestData testData = GetTaiKhoanTestDataFromExcel("QLTK1_02", "QLTK");
            driver.FindElement(By.Id("HoVaTen")).SendKeys(testData.HoVaTen);
            new SelectElement(driver.FindElement(By.Id("GioiTinh"))).SelectByText(testData.GioiTinh);
            driver.FindElement(By.Id("NamSinh")).SendKeys(testData.NamSinh);
            driver.FindElement(By.Id("SDT")).SendKeys(testData.SDT);
            driver.FindElement(By.Id("Email")).SendKeys(testData.Email);
            driver.FindElement(By.Id("UserName")).SendKeys(testData.UserName);
            driver.FindElement(By.Id("Matkhau")).SendKeys(testData.Matkhau);
            IWebElement avatarElement = wait.Until(d => d.FindElement(By.Name("Avatar")));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", avatarElement);
            Thread.Sleep(500);
            avatarElement.SendKeys(testData.Avatar);
            new SelectElement(driver.FindElement(By.Id("IDRole"))).SelectByValue(testData.IDRole);

            IWebElement submitButton = driver.FindElement(By.CssSelector("input[type='submit'].btn.btn-success"));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", submitButton);
            Thread.Sleep(500);
            submitButton.Click();
            Thread.Sleep(1000);

            bool found = false;
            while (!found)
            {
                // Tìm tài khoản trên trang hiện tại
                var accountElements = driver.FindElements(By.XPath($"//td[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{testData.HoVaTen.ToLower()}')]"));
                if (accountElements.Count > 0 && accountElements[0].Displayed)
                {
                    IWebElement accountName = accountElements[0];
                    Assert.That(accountName.Displayed, Is.True, "Tài khoản hiển thị trong danh sách!");
                    Console.WriteLine($"Tài khoản '{testData.HoVaTen}' đã được tìm thấy!");
                    Assert.Fail($"Tài khoản '{testData.HoVaTen}' không được tồn tại trong danh sách!"); 
                    found = true; 
                }
                else
                {
                    var nextPageButtons = driver.FindElements(By.XPath("//li[@class='PagedList-skipToNext']/a[@rel='next']"));
                    if (nextPageButtons.Count > 0 && nextPageButtons[0].Displayed)
                    {
                        Console.WriteLine("Không tìm thấy tài khoản trên trang hiện tại, chuyển sang trang tiếp theo...");
                        nextPageButtons[0].Click();
                        Thread.Sleep(1000);
                        Console.WriteLine($"URL sau khi chuyển trang: {driver.Url}");
                    }
                    else
                    {
                        var tableRows = driver.FindElements(By.XPath("//table//tr"));
                        Console.WriteLine("Nội dung bảng danh sách tài khoản:");
                        foreach (var row in tableRows)
                        {
                            Console.WriteLine($"Row content: {row.Text}");
                        }
                        Console.WriteLine($"Không tìm thấy tài khoản '{testData.HoVaTen}' trong danh sách sau khi duyệt qua tất cả các trang!");
                        break;
                    }
                }
            }
        }
        // testcase QLTK1_03 thêm tài khoản mới mật khẩu bị trùng sdt và email
        [Test]
        public void Test03_CreateNewAccount_trungSDTEMAILunValid()
        {
            Login("tuan159", "Tuan@159159");
            wait.Until(d => d.FindElement(By.Id("userMenuButton")).Displayed);
            driver.FindElement(By.Id("userMenuButton")).Click();
            driver.FindElement(By.Id("manageLink")).Click();
            driver.FindElement(By.LinkText("Quản lý tài khoản")).Click();
            Thread.Sleep(1000);
            IWebElement createButton = wait.Until(d => d.FindElement(By.CssSelector("a.btn.btn-primary[href='/Admin/TaiKhoan/Create']")));
            createButton.Click();
            Thread.Sleep(1000);

            // Lấy dữ liệu TestCaseId = "QLTK1_03"
            TaiKhoanTestData testData = GetTaiKhoanTestDataFromExcel("QLTK1_03", "QLTK");
            driver.FindElement(By.Id("HoVaTen")).SendKeys(testData.HoVaTen);
            new SelectElement(driver.FindElement(By.Id("GioiTinh"))).SelectByText(testData.GioiTinh);
            driver.FindElement(By.Id("NamSinh")).SendKeys(testData.NamSinh);
            driver.FindElement(By.Id("SDT")).SendKeys(testData.SDT);
            driver.FindElement(By.Id("Email")).SendKeys(testData.Email);
            driver.FindElement(By.Id("UserName")).SendKeys(testData.UserName);
            driver.FindElement(By.Id("Matkhau")).SendKeys(testData.Matkhau);
            IWebElement avatarElement = wait.Until(d => d.FindElement(By.Name("Avatar")));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", avatarElement);
            Thread.Sleep(500);
            avatarElement.SendKeys(testData.Avatar);
            new SelectElement(driver.FindElement(By.Id("IDRole"))).SelectByValue(testData.IDRole);

            IWebElement submitButton = driver.FindElement(By.CssSelector("input[type='submit'].btn.btn-success"));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", submitButton);
            Thread.Sleep(500);
            submitButton.Click();
            Thread.Sleep(1000);

            bool found = false;
            while (!found)
            {
                // Tìm tài khoản trên trang hiện tại
                var accountElements = driver.FindElements(By.XPath($"//td[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{testData.HoVaTen.ToLower()}')]"));
                if (accountElements.Count > 0 && accountElements[0].Displayed)
                {
                    IWebElement accountName = accountElements[0];
                    Assert.That(accountName.Displayed, Is.True, "Tài khoản hiển thị trong danh sách!");
                    Console.WriteLine($"Tài khoản '{testData.HoVaTen}' đã được tìm thấy!");
                    Assert.Fail($"Tài khoản '{testData.HoVaTen}' không được tồn tại trong danh sách!");
                    found = true;
                }
                else
                {
                    var nextPageButtons = driver.FindElements(By.XPath("//li[@class='PagedList-skipToNext']/a[@rel='next']"));
                    if (nextPageButtons.Count > 0 && nextPageButtons[0].Displayed)
                    {
                        Console.WriteLine("Không tìm thấy tài khoản trên trang hiện tại, chuyển sang trang tiếp theo...");
                        nextPageButtons[0].Click();
                        Thread.Sleep(1000);
                        Console.WriteLine($"URL sau khi chuyển trang: {driver.Url}");
                    }
                    else
                    {
                        var tableRows = driver.FindElements(By.XPath("//table//tr"));
                        Console.WriteLine("Nội dung bảng danh sách tài khoản:");
                        foreach (var row in tableRows)
                        {
                            Console.WriteLine($"Row content: {row.Text}");
                        }
                        Console.WriteLine($"Không tìm thấy tài khoản '{testData.HoVaTen}' trong danh sách sau khi duyệt qua tất cả các trang!");
                        break;
                    }
                }
            }
        }
        // testcase QLTK1_04 thêm tài khoản mới mật khẩu trùng tài khoản
        [Test]
        public void Test04_CreateNewAccount_TrungUsernameunValid()
        {
            Login("tuan159", "Tuan@159159");
            wait.Until(d => d.FindElement(By.Id("userMenuButton")).Displayed);
            driver.FindElement(By.Id("userMenuButton")).Click();
            driver.FindElement(By.Id("manageLink")).Click();
            driver.FindElement(By.LinkText("Quản lý tài khoản")).Click();
            Thread.Sleep(1000);
            IWebElement createButton = wait.Until(d => d.FindElement(By.CssSelector("a.btn.btn-primary[href='/Admin/TaiKhoan/Create']")));
            createButton.Click();
            Thread.Sleep(1000);

            // Lấy dữ liệu TestCaseId = "QLTK1_04"
            TaiKhoanTestData testData = GetTaiKhoanTestDataFromExcel("QLTK1_04", "QLTK");
            driver.FindElement(By.Id("HoVaTen")).SendKeys(testData.HoVaTen);
            new SelectElement(driver.FindElement(By.Id("GioiTinh"))).SelectByText(testData.GioiTinh);
            driver.FindElement(By.Id("NamSinh")).SendKeys(testData.NamSinh);
            driver.FindElement(By.Id("SDT")).SendKeys(testData.SDT);
            driver.FindElement(By.Id("Email")).SendKeys(testData.Email);
            driver.FindElement(By.Id("UserName")).SendKeys(testData.UserName);
            driver.FindElement(By.Id("Matkhau")).SendKeys(testData.Matkhau);
            IWebElement avatarElement = wait.Until(d => d.FindElement(By.Name("Avatar")));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", avatarElement);
            Thread.Sleep(500);
            avatarElement.SendKeys(testData.Avatar);
            new SelectElement(driver.FindElement(By.Id("IDRole"))).SelectByValue(testData.IDRole);

            IWebElement submitButton = driver.FindElement(By.CssSelector("input[type='submit'].btn.btn-success"));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", submitButton);
            Thread.Sleep(500);
            submitButton.Click();
            Thread.Sleep(1000);

            bool found = false;
            while (!found)
            {
                // Tìm tài khoản trên trang hiện tại
                var accountElements = driver.FindElements(By.XPath($"//td[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{testData.HoVaTen.ToLower()}')]"));
                if (accountElements.Count > 0 && accountElements[0].Displayed)
                {
                    IWebElement accountName = accountElements[0];
                    Assert.That(accountName.Displayed, Is.True, "Tài khoản hiển thị trong danh sách!");
                    Console.WriteLine($"Tài khoản '{testData.HoVaTen}' đã được tìm thấy!");
                    Assert.Fail($"Tài khoản '{testData.HoVaTen}' không được tồn tại trong danh sách!");
                    found = true;
                }
                else
                {
                    var nextPageButtons = driver.FindElements(By.XPath("//li[@class='PagedList-skipToNext']/a[@rel='next']"));
                    if (nextPageButtons.Count > 0 && nextPageButtons[0].Displayed)
                    {
                        Console.WriteLine("Không tìm thấy tài khoản trên trang hiện tại, chuyển sang trang tiếp theo...");
                        nextPageButtons[0].Click();
                        Thread.Sleep(1000);
                        Console.WriteLine($"URL sau khi chuyển trang: {driver.Url}");
                    }
                    else
                    {
                        var tableRows = driver.FindElements(By.XPath("//table//tr"));
                        Console.WriteLine("Nội dung bảng danh sách tài khoản:");
                        foreach (var row in tableRows)
                        {
                            Console.WriteLine($"Row content: {row.Text}");
                        }
                        Console.WriteLine($"Không tìm thấy tài khoản '{testData.HoVaTen}' trong danh sách sau khi duyệt qua tất cả các trang!");
                        break;
                    }
                }
            }
        }
        // testcase QLTK1_05 thêm tài khoản mới sai định dạng mail
        [Test]
        public void Test05_CreateNewAccount_GmailsaidinhdangunValid()
        {
            Login("tuan159", "Tuan@159159");
            wait.Until(d => d.FindElement(By.Id("userMenuButton")).Displayed);
            driver.FindElement(By.Id("userMenuButton")).Click();
            driver.FindElement(By.Id("manageLink")).Click();
            driver.FindElement(By.LinkText("Quản lý tài khoản")).Click();
            Thread.Sleep(1000);
            IWebElement createButton = wait.Until(d => d.FindElement(By.CssSelector("a.btn.btn-primary[href='/Admin/TaiKhoan/Create']")));
            createButton.Click();
            Thread.Sleep(1000);

            // Lấy dữ liệu TestCaseId = "QLTK1_05"
            TaiKhoanTestData testData = GetTaiKhoanTestDataFromExcel("QLTK1_05", "QLTK");
            driver.FindElement(By.Id("HoVaTen")).SendKeys(testData.HoVaTen);
            new SelectElement(driver.FindElement(By.Id("GioiTinh"))).SelectByText(testData.GioiTinh);
            driver.FindElement(By.Id("NamSinh")).SendKeys(testData.NamSinh);
            driver.FindElement(By.Id("SDT")).SendKeys(testData.SDT);
            driver.FindElement(By.Id("Email")).SendKeys(testData.Email);
            driver.FindElement(By.Id("UserName")).SendKeys(testData.UserName);
            driver.FindElement(By.Id("Matkhau")).SendKeys(testData.Matkhau);
            IWebElement avatarElement = wait.Until(d => d.FindElement(By.Name("Avatar")));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", avatarElement);
            Thread.Sleep(500);
            avatarElement.SendKeys(testData.Avatar);
            new SelectElement(driver.FindElement(By.Id("IDRole"))).SelectByValue(testData.IDRole);

            IWebElement submitButton = driver.FindElement(By.CssSelector("input[type='submit'].btn.btn-success"));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", submitButton);
            Thread.Sleep(500);
            submitButton.Click();
            Thread.Sleep(1000);

            bool found = false;
            while (!found)
            {
                // Tìm tài khoản trên trang hiện tại
                var accountElements = driver.FindElements(By.XPath($"//td[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{testData.HoVaTen.ToLower()}')]"));
                if (accountElements.Count > 0 && accountElements[0].Displayed)
                {
                    IWebElement accountName = accountElements[0];
                    Assert.That(accountName.Displayed, Is.True, "Tài khoản hiển thị trong danh sách!");
                    Console.WriteLine($"Tài khoản '{testData.HoVaTen}' đã được tìm thấy!");
                    Assert.Fail($"Tài khoản '{testData.HoVaTen}' không được tồn tại trong danh sách!");
                    found = true;
                }
                else
                {
                    var nextPageButtons = driver.FindElements(By.XPath("//li[@class='PagedList-skipToNext']/a[@rel='next']"));
                    if (nextPageButtons.Count > 0 && nextPageButtons[0].Displayed)
                    {
                        Console.WriteLine("Không tìm thấy tài khoản trên trang hiện tại, chuyển sang trang tiếp theo...");
                        nextPageButtons[0].Click();
                        Thread.Sleep(1000);
                        Console.WriteLine($"URL sau khi chuyển trang: {driver.Url}");
                    }
                    else
                    {
                        var tableRows = driver.FindElements(By.XPath("//table//tr"));
                        Console.WriteLine("Nội dung bảng danh sách tài khoản:");
                        foreach (var row in tableRows)
                        {
                            Console.WriteLine($"Row content: {row.Text}");
                        }
                        Console.WriteLine($"Không tìm thấy tài khoản '{testData.HoVaTen}' trong danh sách sau khi duyệt qua tất cả các trang!");
                        break;
                    }
                }
            }
        }
       
        // testcase QLTK1_06 thêm tài khoản mới sai định dạng SDT
        [Test]
        public void Test06_CreateNewAccount_SDTsaidinhdangunValid()
        {
            Login("tuan159", "Tuan@159159");
            wait.Until(d => d.FindElement(By.Id("userMenuButton")).Displayed);
            driver.FindElement(By.Id("userMenuButton")).Click();
            driver.FindElement(By.Id("manageLink")).Click();
            driver.FindElement(By.LinkText("Quản lý tài khoản")).Click();
            Thread.Sleep(1000);
            IWebElement createButton = wait.Until(d => d.FindElement(By.CssSelector("a.btn.btn-primary[href='/Admin/TaiKhoan/Create']")));
            createButton.Click();
            Thread.Sleep(1000);

            // Lấy dữ liệu TestCaseId = "QLTK1_06"
            TaiKhoanTestData testData = GetTaiKhoanTestDataFromExcel("QLTK1_06", "QLTK");
            driver.FindElement(By.Id("HoVaTen")).SendKeys(testData.HoVaTen);
            new SelectElement(driver.FindElement(By.Id("GioiTinh"))).SelectByText(testData.GioiTinh);
            driver.FindElement(By.Id("NamSinh")).SendKeys(testData.NamSinh);
            driver.FindElement(By.Id("SDT")).SendKeys(testData.SDT);
            driver.FindElement(By.Id("Email")).SendKeys(testData.Email);
            driver.FindElement(By.Id("UserName")).SendKeys(testData.UserName);
            driver.FindElement(By.Id("Matkhau")).SendKeys(testData.Matkhau);
            IWebElement avatarElement = wait.Until(d => d.FindElement(By.Name("Avatar")));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", avatarElement);
            Thread.Sleep(500);
            avatarElement.SendKeys(testData.Avatar);
            new SelectElement(driver.FindElement(By.Id("IDRole"))).SelectByValue(testData.IDRole);

            IWebElement submitButton = driver.FindElement(By.CssSelector("input[type='submit'].btn.btn-success"));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", submitButton);
            Thread.Sleep(500);
            submitButton.Click();
            Thread.Sleep(1000);

            bool found = false;
            while (!found)
            {
                // Tìm tài khoản trên trang hiện tại
                var accountElements = driver.FindElements(By.XPath($"//td[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{testData.HoVaTen.ToLower()}')]"));
                if (accountElements.Count > 0 && accountElements[0].Displayed)
                {
                    IWebElement accountName = accountElements[0];
                    Assert.That(accountName.Displayed, Is.True, "Tài khoản hiển thị trong danh sách!");
                    Console.WriteLine($"Tài khoản '{testData.HoVaTen}' đã được tìm thấy!");
                    Assert.Fail($"Tài khoản '{testData.HoVaTen}' không được tồn tại trong danh sách!");
                    found = true;
                }
                else
                {
                    var nextPageButtons = driver.FindElements(By.XPath("//li[@class='PagedList-skipToNext']/a[@rel='next']"));
                    if (nextPageButtons.Count > 0 && nextPageButtons[0].Displayed)
                    {
                        Console.WriteLine("Không tìm thấy tài khoản trên trang hiện tại, chuyển sang trang tiếp theo...");
                        nextPageButtons[0].Click();
                        Thread.Sleep(1000);
                        Console.WriteLine($"URL sau khi chuyển trang: {driver.Url}");
                    }
                    else
                    {
                        var tableRows = driver.FindElements(By.XPath("//table//tr"));
                        Console.WriteLine("Nội dung bảng danh sách tài khoản:");
                        foreach (var row in tableRows)
                        {
                            Console.WriteLine($"Row content: {row.Text}");
                        }
                        Console.WriteLine($"Không tìm thấy tài khoản '{testData.HoVaTen}' trong danh sách sau khi duyệt qua tất cả các trang!");
                        break;
                    }
                }
            }
        }
        // testcase QLTK1_07 thêm tài khoản mới để trống
        [Test]
        public void Test07_CreateNewAccount_nullunValid()
        {
            Login("tuan159", "Tuan@159159");
            wait.Until(d => d.FindElement(By.Id("userMenuButton")).Displayed);
            driver.FindElement(By.Id("userMenuButton")).Click();
            driver.FindElement(By.Id("manageLink")).Click();
            driver.FindElement(By.LinkText("Quản lý tài khoản")).Click();
            Thread.Sleep(1000);
            IWebElement createButton = wait.Until(d => d.FindElement(By.CssSelector("a.btn.btn-primary[href='/Admin/TaiKhoan/Create']")));
            createButton.Click();
            Thread.Sleep(1000);

            // Lấy dữ liệu TestCaseId = "QLTK1_07"
            TaiKhoanTestData testData = GetTaiKhoanTestDataFromExcel("QLTK1_07", "QLTK");
            new SelectElement(driver.FindElement(By.Id("IDRole"))).SelectByValue(testData.IDRole);

            IWebElement submitButton = driver.FindElement(By.CssSelector("input[type='submit'].btn.btn-success"));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", submitButton);
            Thread.Sleep(500);
            submitButton.Click();
            Thread.Sleep(1000);

            bool found = false;
            while (!found)
            {
                // Tìm tài khoản trên trang hiện tại
                var accountElements = driver.FindElements(By.XPath($"//td[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{testData.HoVaTen.ToLower()}')]"));
                if (accountElements.Count > 0 && accountElements[0].Displayed)
                {
                    IWebElement accountName = accountElements[0];
                    Assert.That(accountName.Displayed, Is.True, "Tài khoản hiển thị trong danh sách!");
                    Console.WriteLine($"Tài khoản '{testData.HoVaTen}' đã được tìm thấy!");
                    Assert.Fail($"Tài khoản '{testData.HoVaTen}' không được tồn tại trong danh sách!");
                    found = true;
                }
                else
                {
                    var nextPageButtons = driver.FindElements(By.XPath("//li[@class='PagedList-skipToNext']/a[@rel='next']"));
                    if (nextPageButtons.Count > 0 && nextPageButtons[0].Displayed)
                    {
                        Console.WriteLine("Không tìm thấy tài khoản trên trang hiện tại, chuyển sang trang tiếp theo...");
                        nextPageButtons[0].Click();
                        Thread.Sleep(1000);
                        Console.WriteLine($"URL sau khi chuyển trang: {driver.Url}");
                    }
                    else
                    {
                        var tableRows = driver.FindElements(By.XPath("//table//tr"));
                        Console.WriteLine("Nội dung bảng danh sách tài khoản:");
                        foreach (var row in tableRows)
                        {
                            Console.WriteLine($"Row content: {row.Text}");
                        }
                        Console.WriteLine($"Không tìm thấy tài khoản '{testData.HoVaTen}' trong danh sách sau khi duyệt qua tất cả các trang!");
                        break;
                    }
                }
            }
        }
       // testcase QLTK2_01 sửa tài khoản hợp lệ
        [Test]
        public void Test08_EditAccountValid()
        {
            Login("tuan159", "Tuan@159159");
            wait.Until(d => d.FindElement(By.Id("userMenuButton")).Displayed);
            driver.FindElement(By.Id("userMenuButton")).Click();
            driver.FindElement(By.Id("manageLink")).Click();
            driver.FindElement(By.LinkText("Quản lý tài khoản")).Click();
            Thread.Sleep(1000);

            // Tìm tài khoản từ test case QLTK1_01 trong danh sách
            TaiKhoanTestData originalData = GetTaiKhoanTestDataFromExcel("QLTK1_01", "QLTK");
            bool found = false;
            IWebElement editButton = null;
            while (!found)
            {
                // Tìm hàng chứa tài khoản nút "Sửa"
                var editButtons = driver.FindElements(By.XPath($"//td[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{originalData.HoVaTen.ToLower()}')]/following-sibling::td/a[@class='btn btn-warning' and text()='Sửa']"));
                if (editButtons.Count > 0 && editButtons[0].Displayed)
                {
                    editButton = editButtons[0];
                    found = true;// thoát vòng lặp
                }
                else
                {
                    // Không tìm thấy tài khoản trên trang hiện tại, kiểm tra xem có nút "Trang tiếp theo" không
                    var nextPageButtons = driver.FindElements(By.XPath("//li[@class='PagedList-skipToNext']/a[@rel='next']"));
                    if (nextPageButtons.Count > 0 && nextPageButtons[0].Displayed)
                    {
                        Console.WriteLine("Không tìm thấy tài khoản trên trang hiện tại, chuyển sang trang tiếp theo...");
                        nextPageButtons[0].Click();
                        Thread.Sleep(1000); 
                        Console.WriteLine($"URL sau khi chuyển trang: {driver.Url}");
                    }
                    else
                    {
                        var tableRows = driver.FindElements(By.XPath("//table//tr"));
                        Console.WriteLine("Nội dung bảng danh sách tài khoản:");
                        foreach (var row in tableRows)
                        {
                            Console.WriteLine($"Row content: {row.Text}");
                        }
                        Assert.Fail($"Không tìm thấy tài khoản '{originalData.HoVaTen}' trong danh sách để chỉnh sửa!");
                        break;
                    }
                }
            }
            editButton.Click();
            Thread.Sleep(1000);
            Console.WriteLine($"URL sau khi nhấn Sửa: {driver.Url}");

            // Lấy dữ liệu TestCaseId = "QLTK2_01"
            TaiKhoanTestData updatedData = GetTaiKhoanTestDataFromExcel("QLTK2_01", "QLTK");
            driver.FindElement(By.Id("HoVaTen")).Clear();
            driver.FindElement(By.Id("HoVaTen")).SendKeys(updatedData.HoVaTen);
            new SelectElement(driver.FindElement(By.Id("GioiTinh"))).SelectByText(updatedData.GioiTinh);
            driver.FindElement(By.Id("NamSinh")).Clear();
            driver.FindElement(By.Id("NamSinh")).SendKeys(updatedData.NamSinh);
            driver.FindElement(By.Id("SDT")).Clear();
            driver.FindElement(By.Id("SDT")).SendKeys(updatedData.SDT);
            driver.FindElement(By.Id("Email")).Clear();
            driver.FindElement(By.Id("Email")).SendKeys(updatedData.Email);
            driver.FindElement(By.Id("UserName")).Clear();
            driver.FindElement(By.Id("UserName")).SendKeys(updatedData.UserName);

            IWebElement avatarElement = wait.Until(d => d.FindElement(By.Name("Avatar")));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", avatarElement);
            Thread.Sleep(500);
            avatarElement.SendKeys(updatedData.Avatar);
            new SelectElement(driver.FindElement(By.Id("IDRole"))).SelectByValue(updatedData.IDRole);

            IWebElement submitButton = wait.Until(d => d.FindElement(By.CssSelector("input[type='submit'].btn.btn-success")));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", submitButton);
            Thread.Sleep(500);
            submitButton.Click();

            Thread.Sleep(1000);

            bool foundUpdated = false;
            while (!foundUpdated)
            {
                // Tìm tài khoản đã sửa trên trang hiện tại
                var accountElements = driver.FindElements(By.XPath($"//tr/td[1][contains(text(), '{updatedData.HoVaTen}')]"));
                if (accountElements.Count > 0 && accountElements[0].Displayed)
                {
                    IWebElement accountName = accountElements[0];
                    Assert.That(accountName.Displayed, Is.True, "Tài khoản đã sửa không hiển thị trong danh sách!");
                    Console.WriteLine($"Tài khoản '{updatedData.HoVaTen}' đã được sửa thành công!");
                    foundUpdated = true; // Tìm thấy tài khoản, thoát vòng lặp
                }
                else
                {
                    // Không tìm thấy tài khoản trên trang hiện tại, kiểm tra xem có nút "Trang tiếp theo" không
                    var nextPageButtons = driver.FindElements(By.XPath("//li[@class='PagedList-skipToNext']/a[@rel='next']"));
                    if (nextPageButtons.Count > 0 && nextPageButtons[0].Displayed)
                    {
                        Console.WriteLine("Không tìm thấy tài khoản đã sửa trên trang hiện tại, chuyển sang trang tiếp theo...");
                        nextPageButtons[0].Click();
                        Thread.Sleep(1000); // Chờ trang tải
                        Console.WriteLine($"URL sau khi chuyển trang: {driver.Url}");
                    }
                    else
                    {
                        // Không còn trang tiếp theo để chuyển
                        var tableRows = driver.FindElements(By.XPath("//table//tr"));
                        Console.WriteLine("Nội dung bảng danh sách tài khoản:");
                        foreach (var row in tableRows)
                        {
                            Console.WriteLine($"Row content: {row.Text}");
                        }
                        Assert.Fail($"Không tìm thấy tài khoản đã sửa '{updatedData.HoVaTen}' trong danh sách sau khi duyệt qua tất cả các trang!");
                        break;
                    }
                }
            }
        }
        // testcase QLTK4_01 tìm kiếm tài khoản hợp lệ
        [Test]
        public void Test09_SearchAccountValid()
        {
            Login("tuan159", "Tuan@159159");
            wait.Until(d => d.FindElement(By.Id("userMenuButton")).Displayed);
            driver.FindElement(By.Id("userMenuButton")).Click();
            driver.FindElement(By.Id("manageLink")).Click();
            driver.FindElement(By.LinkText("Quản lý tài khoản")).Click();
            Thread.Sleep(1000);

            // Lấy dữ liệu TestCaseId = "QLTK4_01"
            TaiKhoanTestData searchData = GetTaiKhoanTestDataFromExcel("QLTK4_01", "QLTK");
            IWebElement searchInput = wait.Until(d => d.FindElement(By.Name("SearchString")));
            searchInput.Clear();
            searchInput.SendKeys(searchData.SearchKeyword);
     
            IWebElement searchButton = wait.Until(d => d.FindElement(By.CssSelector("input[value='Search']"))); searchButton.Click();
            Thread.Sleep(1000);
            Console.WriteLine($"URL sau khi tìm kiếm: {driver.Url}");

            bool found = false;
            while (!found)
            {
                // Tìm tài khoản trên trang hiện tại
                var accountElements = driver.FindElements(By.XPath($"//tr/td[1][contains(text(), '{searchData.HoVaTen}')]"));
                if (accountElements.Count > 0 && accountElements[0].Displayed)
                {
                    IWebElement accountName = accountElements[0];
                    Assert.That(accountName.Displayed, Is.True, "Tài khoản không hiển thị trong kết quả tìm kiếm!");
                    Console.WriteLine($"Tài khoản '{searchData.HoVaTen}' đã được tìm thấy thành công!");
                    found = true; // thoát vòng lặp
                }
                else
                {
                    var nextPageButtons = driver.FindElements(By.XPath("//li[@class='PagedList-skipToNext']/a[@rel='next']"));
                    if (nextPageButtons.Count > 0 && nextPageButtons[0].Displayed)
                    {
                        Console.WriteLine("Không tìm thấy tài khoản trên trang hiện tại, chuyển sang trang tiếp theo...");
                        nextPageButtons[0].Click();
                        Thread.Sleep(1000); 
                        Console.WriteLine($"URL sau khi chuyển trang: {driver.Url}");
                    }
                    else
                    {
                        var tableRows = driver.FindElements(By.XPath("//table//tr"));
                        Console.WriteLine("Nội dung bảng danh sách tài khoản:");
                        foreach (var row in tableRows)
                        {
                            Console.WriteLine($"Row content: {row.Text}");
                        }
                        Assert.Fail($"Không tìm thấy tài khoản '{searchData.HoVaTen}' trong kết quả tìm kiếm sau khi duyệt qua tất cả các trang!");
                        break;
                    }
                }
            }
        }
        // testcase QLTK4_02 tìm kiếm sản phẩm hợp lệ bằng role
        [Test]
        public void Test10_SearchAccountroleValid()
        {
            Login("tuan159", "Tuan@159159");
            wait.Until(d => d.FindElement(By.Id("userMenuButton")).Displayed);
            driver.FindElement(By.Id("userMenuButton")).Click();
            driver.FindElement(By.Id("manageLink")).Click();
            driver.FindElement(By.LinkText("Quản lý tài khoản")).Click();
            Thread.Sleep(1000);

            // Lấy dữ liệu TestCaseId = "QLTK4_01"
            TaiKhoanTestData searchData = GetTaiKhoanTestDataFromExcel("QLTK4_02", "QLTK");

            IWebElement searchInput = wait.Until(d => d.FindElement(By.Name("SearchString")));
            searchInput.Clear();
            searchInput.SendKeys(searchData.SearchKeyword);
           
            IWebElement searchButton = wait.Until(d => d.FindElement(By.CssSelector("input[value='Search']"))); searchButton.Click();
            Thread.Sleep(1000);
            Console.WriteLine($"URL sau khi tìm kiếm: {driver.Url}");


            // Kiểm tra kết quả 
            bool found = false;
            while (!found)
            {
                // Tìm tài khoản trên trang hiện tại
                var accountElements = driver.FindElements(By.XPath($"//tr/td[1][contains(text(), '{searchData.HoVaTen}')]"));
                if (accountElements.Count > 0 && accountElements[0].Displayed)
                {
                    IWebElement accountName = accountElements[0];
                    Assert.That(accountName.Displayed, Is.True, "Tài khoản không hiển thị trong kết quả tìm kiếm!");
                    Console.WriteLine($"Tài khoản '{searchData.HoVaTen}' đã được tìm thấy thành công!");
                    found = true; // thoát vòng lặp
                }
                else
                {
                    var nextPageButtons = driver.FindElements(By.XPath("//li[@class='PagedList-skipToNext']/a[@rel='next']"));
                    if (nextPageButtons.Count > 0 && nextPageButtons[0].Displayed)
                    {
                        Console.WriteLine("Không tìm thấy tài khoản trên trang hiện tại, chuyển sang trang tiếp theo...");
                        nextPageButtons[0].Click();
                        Thread.Sleep(1000);
                        Console.WriteLine($"URL sau khi chuyển trang: {driver.Url}");
                    }
                    else
                    {
                        var tableRows = driver.FindElements(By.XPath("//table//tr"));
                        Console.WriteLine("Nội dung bảng danh sách tài khoản:");
                        foreach (var row in tableRows)
                        {
                            Console.WriteLine($"Row content: {row.Text}");
                        }
                        Assert.Fail($"Không tìm thấy tài khoản '{searchData.HoVaTen}' trong kết quả tìm kiếm sau khi duyệt qua tất cả các trang!");
                        break;
                    }
                }
            }
        }
        //testcase QLTK3_01 xóa tài khoản hợp lệ
        [Test]
        public void Test11_DeleteAccountValid()
        {
            Login("tuan159", "Tuan@159159");
            wait.Until(d => d.FindElement(By.Id("userMenuButton")).Displayed);
            driver.FindElement(By.Id("userMenuButton")).Click();

            driver.FindElement(By.Id("manageLink")).Click();
            driver.FindElement(By.LinkText("Quản lý tài khoản")).Click();
            Thread.Sleep(1000);

            //dữ liệu xóa TestCaseId = "QLTK2_01"
            TaiKhoanTestData account = GetTaiKhoanTestDataFromExcel("QLTK2_01", "QLTK");
            bool found = false;
            IWebElement deleteButton = null;
            while (!found)
            {
                // Tìm hàng chứa tài khoản và nút "Xóa"
                var deleteButtons = driver.FindElements(By.XPath($"//td[contains(text(), '{account.HoVaTen}')]/following-sibling::td/a[@class='btn btn-danger' and text()='Xóa']"));
                if (deleteButtons.Count > 0 && deleteButtons[0].Displayed)
                {
                    deleteButton = deleteButtons[0];
                    found = true; // Tìm thấy nút "Xóa", thoát vòng lặp
                }
                else
                {
                    // Không tìm thấy tài khoản trên trang hiện tại, kiểm tra xem có nút "Trang tiếp theo" không
                    var nextPageButtons = driver.FindElements(By.XPath("//li[@class='PagedList-skipToNext']/a[@rel='next']"));
                    if (nextPageButtons.Count > 0 && nextPageButtons[0].Displayed)
                    {
                        Console.WriteLine("Không tìm thấy tài khoản trên trang hiện tại, chuyển sang trang tiếp theo...");
                        nextPageButtons[0].Click();
                        Thread.Sleep(1000); 
                        Console.WriteLine($"URL sau khi chuyển trang: {driver.Url}");
                    }
                    else
                    {
                        // Không còn trang tiếp theo để chuyển
                        var tableRows = driver.FindElements(By.XPath("//table//tr"));
                        Console.WriteLine("Nội dung bảng danh sách tài khoản:");
                        foreach (var row in tableRows)
                        {
                            Console.WriteLine($"Row content: {row.Text}");
                        }
                        Assert.Fail($"Không tìm thấy tài khoản '{account.HoVaTen}' trong danh sách để xóa!");
                        break;
                    }
                }
            }
            deleteButton.Click();
            Thread.Sleep(1000);
            Console.WriteLine($"URL sau khi nhấn Xóa: {driver.Url}");

            // Xác nhận xóa
            IWebElement confirmDeleteButton = wait.Until(d => d.FindElement(By.CssSelector("input[type='submit'].btn.btn-danger")));
            confirmDeleteButton.Click();
            Thread.Sleep(1000);
            Console.WriteLine($"URL sau khi xác nhận xóa: {driver.Url}");

            // Kiểm tra tài khoản
            bool accountStillExists = false;
            driver.Navigate().Refresh();
            Thread.Sleep(1000);
            while (true)
            {
                // Tìm tài khoản trong danh sách
                var accountElements = driver.FindElements(By.XPath($"//tr/td[1][contains(text(), '{account.HoVaTen}')]"));
                if (accountElements.Count > 0 && accountElements[0].Displayed)
                {
                    accountStillExists = true;
                    break; 
                }
                else
                {
                    var nextPageButtons = driver.FindElements(By.XPath("//li[@class='PagedList-skipToNext']/a[@rel='next']"));
                    if (nextPageButtons.Count > 0 && nextPageButtons[0].Displayed)
                    {
                        Console.WriteLine("Không tìm thấy tài khoản trên trang hiện tại, chuyển sang trang tiếp theo...");
                        nextPageButtons[0].Click();
                        Thread.Sleep(1000);
                        Console.WriteLine($"URL sau khi chuyển trang: {driver.Url}");
                    }
                    else
                    {
                        break;
                    }
                }
            }

            // Kiểm tra kết quả
            if (accountStillExists)
            {
                var tableRows = driver.FindElements(By.XPath("//table//tr"));
                Console.WriteLine("Nội dung bảng danh sách tài khoản:");
                foreach (var row in tableRows)
                {
                    Console.WriteLine($"Row content: {row.Text}");
                }
                Assert.Fail($"Tài khoản '{account.HoVaTen}' vẫn tồn tại trong danh sách sau khi xóa!");
            }
            else
            {
                Console.WriteLine($"Tài khoản '{account.HoVaTen}' đã được xóa thành công!");
            }
        }
        
       
        //----------------------------------------------------------testcase quản lý danh mục------------------------------------
        //testcase QLDM1_01 Tạo danh mục mới hợp lệ
        [Test]
        public void Test12_CreateCategory_Valid()
        {
            Login("tuan159", "Tuan@159159");
            wait.Until(d => d.FindElement(By.Id("userMenuButton")).Displayed);
            driver.FindElement(By.Id("userMenuButton")).Click();
            driver.FindElement(By.Id("manageLink")).Click();
            driver.FindElement(By.LinkText("Quản lý danh mục")).Click();
            driver.FindElement(By.CssSelector("a.btn.btn-primary[href='/Admin/LoaiSanPham/Create']")).Click();
            wait.Until(d => d.FindElement(By.Id("TenLoaiSanPham")).Displayed);

            // Lấy dữ liệu từ Excel
            var testData = GetCategoryTestDataFromExcel("QLDM1_01", "QLDM"); //TestCaseId

            driver.FindElement(By.Id("TenLoaiSanPham")).SendKeys(testData.TenLoaiSanPham);
            driver.FindElement(By.Name("file")).SendKeys(testData.ImagePath);

            IWebElement submitButton = driver.FindElement(By.CssSelector("button.btn.btn-success"));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", submitButton);
            Thread.Sleep(500);
            submitButton.Click();
            bool found = false;
            while (!found)
            {
                // Tìm danh mục trên trang hiện tại
                var categoryElements = driver.FindElements(By.XPath($"//td[contains(text(), '{testData.TenLoaiSanPham}')]"));
                if (categoryElements.Count > 0 && categoryElements[0].Displayed)
                {
                    IWebElement categoryName = categoryElements[0];
                    Assert.That(categoryName.Displayed, Is.True, $"Danh mục mới '{testData.TenLoaiSanPham}' không hiển thị trong danh sách!");
                    Console.WriteLine($"Danh mục '{testData.TenLoaiSanPham}' đã được thêm thành công!");
                    found = true;
                }
                else
                {
                    var nextPageButtons = driver.FindElements(By.XPath("//li[@class='PagedList-skipToNext']/a[@rel='next']"));
                    if (nextPageButtons.Count > 0 && nextPageButtons[0].Displayed)
                    {
                        Console.WriteLine("Không tìm thấy danh mục trên trang hiện tại, chuyển sang trang tiếp theo...");
                        ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", nextPageButtons[0]);
                        Thread.Sleep(500);
                        nextPageButtons[0].Click();
                        Thread.Sleep(1000);
                        Console.WriteLine($"URL sau khi chuyển trang: {driver.Url}");
                    }
                    else
                    {
                        var tableRows = driver.FindElements(By.XPath("//table//tr"));
                        Console.WriteLine("Nội dung bảng danh sách danh mục:");
                        foreach (var row in tableRows)
                        {
                            Console.WriteLine($"Row content: {row.Text}");
                        }
                        Assert.Fail($"Không tìm thấy danh mục '{testData.TenLoaiSanPham}' trong danh sách sau khi duyệt qua tất cả các trang!");
                        break;
                    }
                }
            }

        }
        //testcase QLDM1_02 Tạo danh mục mới ten trùng
        [Test]
        public void Test13_CreateCategorytrunglapunValid()
        {
            Login("tuan159", "Tuan@159159");
            wait.Until(d => d.FindElement(By.Id("userMenuButton")).Displayed);
            driver.FindElement(By.Id("userMenuButton")).Click();
            driver.FindElement(By.Id("manageLink")).Click();
            driver.FindElement(By.LinkText("Quản lý danh mục")).Click();
            driver.FindElement(By.CssSelector("a.btn.btn-primary[href='/Admin/LoaiSanPham/Create']")).Click();
            wait.Until(d => d.FindElement(By.Id("TenLoaiSanPham")).Displayed);

            // Lấy dữ liệu từ Excel
            var testData = GetCategoryTestDataFromExcel("QLDM1_02", "QLDM"); //TestCaseId

            driver.FindElement(By.Id("TenLoaiSanPham")).SendKeys(testData.TenLoaiSanPham);
            driver.FindElement(By.Name("file")).SendKeys(testData.ImagePath);

            IWebElement submitButton = driver.FindElement(By.CssSelector("button.btn.btn-success"));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", submitButton);
            Thread.Sleep(500);
            submitButton.Click();
            var categoryElements = driver.FindElements(By.XPath($"//td[contains(text(), '{testData.TenLoaiSanPham}')]"));
            IWebElement categoryName = categoryElements[0];
            Assert.That(categoryName.Displayed, Is.True, $"Danh mục '{testData.TenLoaiSanPham}' không hiển thị trong danh sách!");

        }
        //testcase QLDM1_03 Tạo danh mục trống tên
        [Test]
        public void Test14_CreateCategorynamenullunValid()
        {
            Login("tuan159", "Tuan@159159");
            wait.Until(d => d.FindElement(By.Id("userMenuButton")).Displayed);
            driver.FindElement(By.Id("userMenuButton")).Click();
            driver.FindElement(By.Id("manageLink")).Click();
            driver.FindElement(By.LinkText("Quản lý danh mục")).Click();
            driver.FindElement(By.CssSelector("a.btn.btn-primary[href='/Admin/LoaiSanPham/Create']")).Click();
            wait.Until(d => d.FindElement(By.Id("TenLoaiSanPham")).Displayed);

            // Lấy dữ liệu từ Excel
            var testData = GetCategoryTestDataFromExcel("QLDM1_03", "QLDM"); //TestCaseId

            driver.FindElement(By.Id("TenLoaiSanPham")).SendKeys(testData.TenLoaiSanPham);
            driver.FindElement(By.Name("file")).SendKeys(testData.ImagePath);

            IWebElement submitButton = driver.FindElement(By.CssSelector("button.btn.btn-success"));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", submitButton);
            Thread.Sleep(500);
            submitButton.Click();
            Thread.Sleep(1500);
            //tìm kiếm qua các trang
            bool categoryFound = false;
            while (!categoryFound)
            {
                var rows = driver.FindElements(By.XPath("//table[@class='table table-hover']/tbody/tr"));
                foreach (var row in rows)
                {
                    var nameElement = row.FindElement(By.XPath("./td[2]"));
                    Console.WriteLine($"- {nameElement.Text}");
                }

                categoryFound = rows.Any(row =>
                    row.FindElement(By.XPath("./td[2]")).Text.Trim().Equals(testData.TenLoaiSanPham.Trim(), StringComparison.OrdinalIgnoreCase)); if (categoryFound)
                {
                    break;
                }

                var nextButton = driver.FindElements(By.CssSelector("li.PagedList-skipToNext a[rel='next']"));
                if (nextButton.Count == 0 || !nextButton[0].Displayed)
                {
                    Assert.Fail($"Không tìm thấy danh mục '{testData.TenLoaiSanPham}' sau khi duyệt qua tất cả các trang.");
                    break;
                }
                IWebElement nextButtonElement = nextButton[0];
                new Actions(driver)
                    .MoveToElement(nextButtonElement)
                    .Pause(TimeSpan.FromMilliseconds(500))
                    .Click()
                    .Perform();

                wait.Until(d => d.FindElements(By.XPath("//table//td")).Count > 0);
            }
            Assert.That(categoryFound, Is.False, $"Danh mục '{testData.TenLoaiSanPham}' tạo thành công.");
        }
        //testcase QLDM1_04 Tạo danh mục trống hình
        [Test]
        public void Test15_CreateCategorynopicunValid()
        {
            Login("tuan159", "Tuan@159159");
            wait.Until(d => d.FindElement(By.Id("userMenuButton")).Displayed);
            driver.FindElement(By.Id("userMenuButton")).Click();
            driver.FindElement(By.Id("manageLink")).Click();
            driver.FindElement(By.LinkText("Quản lý danh mục")).Click();
            driver.FindElement(By.CssSelector("a.btn.btn-primary[href='/Admin/LoaiSanPham/Create']")).Click();
            wait.Until(d => d.FindElement(By.Id("TenLoaiSanPham")).Displayed);

            // Lấy dữ liệu từ Excel
            var testData = GetCategoryTestDataFromExcel("QLDM1_04", "QLDM");

            driver.FindElement(By.Id("TenLoaiSanPham")).SendKeys(testData.TenLoaiSanPham);

            IWebElement submitButton = driver.FindElement(By.CssSelector("button.btn.btn-success"));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", submitButton);
            Thread.Sleep(500);
            submitButton.Click();
            Thread.Sleep(1500);
            //tìm kiếm qua các trang
            bool categoryFound = false;
            while (!categoryFound)
            {
                var rows = driver.FindElements(By.XPath("//table[@class='table table-hover']/tbody/tr"));
                foreach (var row in rows)
                {
                    var nameElement = row.FindElement(By.XPath("./td[2]"));
                    Console.WriteLine($"- {nameElement.Text}");
                }

                categoryFound = rows.Any(row =>
                    row.FindElement(By.XPath("./td[2]")).Text.Trim().Equals(testData.TenLoaiSanPham.Trim(), StringComparison.OrdinalIgnoreCase)); if (categoryFound)
                {
                    break;
                }

                var nextButton = driver.FindElements(By.CssSelector("li.PagedList-skipToNext a[rel='next']"));
                if (nextButton.Count == 0 || !nextButton[0].Displayed)
                {
                    Assert.Fail($"Không tìm thấy danh mục '{testData.TenLoaiSanPham}' sau khi duyệt qua tất cả các trang.");
                    break;
                }
                IWebElement nextButtonElement = nextButton[0];
                new Actions(driver)
                    .MoveToElement(nextButtonElement)
                    .Pause(TimeSpan.FromMilliseconds(500))
                    .Click()
                    .Perform();

                wait.Until(d => d.FindElements(By.XPath("//table//td")).Count > 0);
            }
            Assert.That(categoryFound, Is.False, $"Danh mục '{testData.TenLoaiSanPham}' tạo thành công.");
        }
        // Test case QLDM2_01: Sửa thông tin danh mục hợp lệ
        [Test]
        public void Test16_EditCategory_Valid()
        {
            Login("tuan159", "Tuan@159159");
            wait.Until(d => d.FindElement(By.Id("userMenuButton")).Displayed);
            driver.FindElement(By.Id("userMenuButton")).Click();
            driver.FindElement(By.Id("manageLink")).Click();
            driver.FindElement(By.LinkText("Quản lý danh mục")).Click();

            // Danh mục cần sửa (QLDM1_01)
            var originalData = GetCategoryTestDataFromExcel("QLDM1_01", "QLDM");
            IWebElement categoryRow = wait.Until(d => d.FindElement(By.XPath($"//tr[td[contains(text(), '{originalData.TenLoaiSanPham}')]]")));
            IWebElement editButton = categoryRow.FindElement(By.CssSelector("a.btn.btn-primary"));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", editButton);
            Thread.Sleep(500);
            editButton.Click();
           

            wait.Until(d => d.FindElement(By.Id("TenLoaiSanPham")).Displayed);
            // Lấy dữ liệu mới 
            var updatedData = GetCategoryTestDataFromExcel("QLDM2_01", "QLDM");
            IWebElement tenLoaiSanPham = driver.FindElement(By.Id("TenLoaiSanPham"));
            tenLoaiSanPham.Clear();
            tenLoaiSanPham.SendKeys(updatedData.TenLoaiSanPham);
            driver.FindElement(By.Name("file")).SendKeys(updatedData.ImagePath);

            IWebElement submitButton = driver.FindElement(By.CssSelector("button.btn.btn-success"));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", submitButton);
            Thread.Sleep(500);
            submitButton.Click();

            var updatedCategory = wait.Until(d => d.FindElement(By.XPath($"//td[contains(text(), '{updatedData.TenLoaiSanPham}')]")));
            Assert.That(updatedCategory.Displayed, Is.True, $"Danh mục mới '{updatedData.TenLoaiSanPham}' Hiển thị trong danh sách sau khi sửa!");
        }
        //// Test case QLDM2_02: Sửa thông tin danh mục trùng tên
        //[Test]
        //public void TestunValid_EditCategory()
        //{
        //    Login("tuan159", "Tuan@159159");
        //    wait.Until(d => d.FindElement(By.Id("userMenuButton")).Displayed);
        //    driver.FindElement(By.Id("userMenuButton")).Click();
        //    driver.FindElement(By.Id("manageLink")).Click();
        //    driver.FindElement(By.LinkText("Quản lý danh mục")).Click();

        //    // Danh mục cần sửa (QLDM2_01)
        //    var originalData = GetCategoryTestDataFromExcel("QLDM2_01", "QLSP");

        //    // Tìm danh mục "Bộ Trang Sức" trong danh sách
        //    try
        //    {
        //        IWebElement categoryRow = wait.Until(d => d.FindElement(By.XPath($"//tr[td[contains(text(), '{originalData.TenLoaiSanPham}')]]")));
        //        IWebElement editButton = categoryRow.FindElement(By.CssSelector("a.btn.btn-primary"));
        //        ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", editButton);
        //        Thread.Sleep(500);
        //        editButton.Click();
        //    }
        //    catch (WebDriverTimeoutException)
        //    {
        //        Assert.Fail($"Không tìm thấy danh mục '{originalData.TenLoaiSanPham}' trong danh sách để sửa!");
        //    }

        //    wait.Until(d => d.FindElement(By.Id("TenLoaiSanPham")).Displayed);
        //    // Lấy dữ liệu mới từ Excel
        //    var updatedData = GetCategoryTestDataFromExcel("QLDM2_02", "QLSP");
        //    IWebElement tenLoaiSanPham = driver.FindElement(By.Id("TenLoaiSanPham"));
        //    tenLoaiSanPham.Clear();
        //    tenLoaiSanPham.SendKeys(updatedData.TenLoaiSanPham);

        //    // Kiểm tra file hình ảnh có tồn tại không
        //    if (!File.Exists(updatedData.ImagePath))
        //    {
        //        Assert.Fail($"File hình ảnh không tồn tại tại đường dẫn: {updatedData.ImagePath}. Vui lòng kiểm tra lại.");
        //    }
        //    driver.FindElement(By.Name("file")).SendKeys(updatedData.ImagePath);

        //    IWebElement submitButton = driver.FindElement(By.CssSelector("button.btn.btn-success"));
        //    ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", submitButton);
        //    Thread.Sleep(500);
        //    submitButton.Click();

        //    // Kiểm tra danh mục đã được cập nhật
        //    try
        //    {
        //        var updatedCategory = wait.Until(d => d.FindElement(By.XPath($"//td[contains(text(), '{updatedData.TenLoaiSanPham}')]")));
        //        Assert.That(updatedCategory.Displayed, Is.True, $"Danh mục mới '{updatedData.TenLoaiSanPham}' không hiển thị trong danh sách sau khi sửa!");
        //    }
        //    catch (WebDriverTimeoutException)
        //    {
        //        Assert.Fail($"Không tìm thấy danh mục '{updatedData.TenLoaiSanPham}' trong danh sách sau thời gian chờ!");
        //    }
        //}
        
        //// Test case QLDM2_05: Sửa thông tin danh mục dang có sản phẩm liên kết
        //[Test]
        //public void TestunValid_EditCategory_colienket()
        //{
        //    Login("tuan159", "Tuan@159159");
        //    wait.Until(d => d.FindElement(By.Id("userMenuButton")).Displayed);
        //    driver.FindElement(By.Id("userMenuButton")).Click();
        //    driver.FindElement(By.Id("manageLink")).Click();
        //    driver.FindElement(By.LinkText("Quản lý danh mục")).Click();

        //    // Danh mục cần sửa
        //    var originalData = GetCategoryTestDataFromExcel("QLDM1_00", "QLDM"); //TestCaseId

        //    // Tìm danh mục "Bộ Trang Sức" trong danh sách
        //    try
        //    {
        //        IWebElement categoryRow = wait.Until(d => d.FindElement(By.XPath($"//tr[td[contains(text(), '{originalData.TenLoaiSanPham}')]]")));
        //        IWebElement editButton = categoryRow.FindElement(By.CssSelector("a.btn.btn-primary"));
        //        ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", editButton);
        //        Thread.Sleep(500);
        //        editButton.Click();
        //    }
        //    catch (WebDriverTimeoutException)
        //    {
        //        Assert.Fail($"Không tìm thấy danh mục '{originalData.TenLoaiSanPham}' trong danh sách để sửa!");
        //    }

        //    wait.Until(d => d.FindElement(By.Id("TenLoaiSanPham")).Displayed);
        //    // Lấy dữ liệu mới từ Excel
        //    var updatedData = GetCategoryTestDataFromExcel("QLDM2_05", "QLDM"); //TestCaseId
        //    IWebElement tenLoaiSanPham = driver.FindElement(By.Id("TenLoaiSanPham"));
        //    tenLoaiSanPham.Clear();
        //    tenLoaiSanPham.SendKeys(updatedData.TenLoaiSanPham);

        //    // Kiểm tra file hình ảnh có tồn tại không
        //    if (!File.Exists(updatedData.ImagePath))
        //    {
        //        Assert.Fail($"File hình ảnh không tồn tại tại đường dẫn: {updatedData.ImagePath}. Vui lòng kiểm tra lại.");
        //    }
        //    driver.FindElement(By.Name("file")).SendKeys(updatedData.ImagePath);

        //    IWebElement submitButton = driver.FindElement(By.CssSelector("button.btn.btn-success"));
        //    ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", submitButton);
        //    Thread.Sleep(500);
        //    submitButton.Click();

        //    // Kiểm tra danh mục đã được cập nhật
        //    try
        //    {
        //        var updatedCategory = wait.Until(d => d.FindElement(By.XPath($"//td[contains(text(), '{updatedData.TenLoaiSanPham}')]")));
        //        Assert.That(updatedCategory.Displayed, Is.True, $"Danh mục mới '{updatedData.TenLoaiSanPham}' không hiển thị trong danh sách sau khi sửa!");
        //    }
        //    catch (WebDriverTimeoutException)
        //    {
        //        Assert.Fail($"Không tìm thấy danh mục '{updatedData.TenLoaiSanPham}' trong danh sách sau thời gian chờ!");
        //    }
        //}

       
        [Test]
        // Testcase QLDM3_01: Xóa danh mục hợp lệ
        public void Test17_DeleteCategory_Valid()
        {
            Login("tuan159", "Tuan@159159");
            wait.Until(d => d.FindElement(By.Id("userMenuButton")).Displayed);
            driver.FindElement(By.Id("userMenuButton")).Click();
            driver.FindElement(By.Id("manageLink")).Click();
            driver.FindElement(By.LinkText("Quản lý danh mục")).Click();

            // Lấy dữ liệu từ Excel cho danh mục cần xóa (QLDM1_02)
            var testData = GetCategoryTestDataFromExcel("QLDM1_02", "QLDM");
            Console.WriteLine($"Danh mục cần xóa: {testData.TenLoaiSanPham}");

            IWebElement categoryRow = null;
            bool categoryFound = false;
            while (!categoryFound)
            {
                wait.Until(d => d.FindElement(By.CssSelector("table.table-hover")).Displayed);
                var rows = driver.FindElements(By.XPath("//table[@class='table table-hover']/tbody/tr"));

                Console.WriteLine("Danh mục trên trang hiện tại:");
                foreach (var row in rows)
                {
                    var nameElement = row.FindElement(By.XPath("./td[2]"));
                    Console.WriteLine($"- {nameElement.Text}");
                }

                // Tìm hàng chứa danh mục
                var matchingRow = rows.FirstOrDefault(row =>
                    row.FindElement(By.XPath("./td[2]")).Text.Trim().Equals(testData.TenLoaiSanPham.Trim(), StringComparison.OrdinalIgnoreCase));

                if (matchingRow != null)
                {
                    categoryRow = matchingRow;
                    categoryFound = true;
                    break;
                }

                var nextButton = driver.FindElements(By.CssSelector("li.PagedList-skipToNext a[rel='next']"));
                if (nextButton.Count == 0 || !nextButton[0].Displayed)
                {
                    Assert.Fail($"Không tìm thấy danh mục '{testData.TenLoaiSanPham}' trong danh sách để xóa.");
                    break;
                }

                new Actions(driver)
                    .MoveToElement(nextButton[0])
                    .Pause(TimeSpan.FromMilliseconds(500))
                    .Click()
                    .Perform();

                wait.Until(d => d.FindElement(By.CssSelector("table.table-hover")).Displayed);
            }

            IWebElement deleteButton = categoryRow.FindElement(By.CssSelector("a.btn.btn-danger"));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", deleteButton);
            Thread.Sleep(500);
            deleteButton.Click();

            IAlert alert = wait.Until(d => driver.SwitchTo().Alert());
            alert.Accept();
            wait.Until(d => d.FindElement(By.XPath("//button[@class='btn btn-danger' and text()='Xóa']")).Displayed);
            IWebElement confirmDeleteButton = driver.FindElement(By.XPath("//button[@class='btn btn-danger' and text()='Xóa']"));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", confirmDeleteButton);
            Thread.Sleep(500);
            confirmDeleteButton.Click();
            wait.Until(d => !d.FindElements(By.XPath($"//td[contains(text(), '{testData.TenLoaiSanPham}')]")).Any());
            Assert.Pass($"Danh mục '{testData.TenLoaiSanPham}' đã được xóa thành công!");
        }

        //-----------------------------------Testcase Quản lý thương hiệu----------------------------------------------
        //testcase QLTH1_01 thêm thương hiệu mới hợp lệ
        [Test]
        public void Test18_ThemThuongHieuHopLe()
        {
            Login("tuan159", "Tuan@159159");
            wait.Until(d => d.FindElement(By.Id("userMenuButton")).Displayed);
            driver.FindElement(By.Id("userMenuButton")).Click();
            driver.FindElement(By.Id("manageLink")).Click();
            driver.FindElement(By.LinkText("Quản lý thương hiệu")).Click();
            Thread.Sleep(1000);
            Console.WriteLine($"URL hiện tại: {driver.Url}");

            IWebElement createButton = wait.Until(d => d.FindElement(By.LinkText("Thêm mới")));
            createButton.Click();
            Thread.Sleep(1000);
            Console.WriteLine($"URL sau khi nhấn Thêm mới: {driver.Url}");

            // Lấy dữ liệu test case QLTH1_01
            ThuongHieuTestData testData = GetThuongHieuTestDataFromExcel("QLTH1_01", "QLTH");

            IWebElement tenThuongHieuInput = wait.Until(d => d.FindElement(By.Id("TenThuongHieu")));
            tenThuongHieuInput.Clear();
            tenThuongHieuInput.SendKeys(testData.TenThuongHieu);
            IWebElement saveButton = driver.FindElement(By.CssSelector("input.btn.btn-primary[value='Lưu']"));
            saveButton.Click();
            Thread.Sleep(1000);
            Console.WriteLine($"URL sau khi submit: {driver.Url}");
            bool found = false;
            while (!found)
            {
                var resultElements = wait.Until(d => d.FindElements(By.XPath($"//td[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{testData.TenThuongHieu.ToLower()}')]")));
                if (resultElements.Count > 0 && resultElements[0].Displayed)
                {
                    Assert.That(resultElements[0].Displayed, Is.True, $"Không tìm thấy thương hiệu '{testData.TenThuongHieu}' trong danh sách!");
                    Console.WriteLine($"Thương hiệu '{testData.TenThuongHieu}' đã được thêm thành công: {resultElements[0].Text}");
                    found = true;
                }
                else
                {
                    var nextPageButton = driver.FindElements(By.XPath("//li[@class='PagedList-skipToNext']/a[@rel='next']"));
                    if (nextPageButton.Count > 0 && nextPageButton[0].Displayed)
                    {
                        Console.WriteLine("Không tìm thấy thương hiệu trên trang hiện tại, chuyển sang trang tiếp theo...");
                        nextPageButton[0].Click();
                        Thread.Sleep(1000);
                        Console.WriteLine($"URL sau khi chuyển trang: {driver.Url}");
                    }
                    else
                    {
                        var tableRows = driver.FindElements(By.XPath("//table//tr"));
                        Console.WriteLine("Nội dung bảng danh sách thương hiệu:");
                        foreach (var row in tableRows)
                        {
                            Console.WriteLine($"Row content: {row.Text}");
                        }
                        Assert.Fail($"Không tìm thấy thương hiệu '{testData.TenThuongHieu}' trong danh sách sau khi duyệt qua tất cả các trang!");
                        break;
                    }
                }
            }
        }
        [Test]
        public void Test19_ThemThuongHieu_TRungtenunValid()
        {
            Login("tuan159", "Tuan@159159");
            wait.Until(d => d.FindElement(By.Id("userMenuButton")).Displayed);
            driver.FindElement(By.Id("userMenuButton")).Click();
            driver.FindElement(By.Id("manageLink")).Click();
            driver.FindElement(By.LinkText("Quản lý thương hiệu")).Click();
            Thread.Sleep(1000);
            Console.WriteLine($"URL hiện tại: {driver.Url}");
            IWebElement createButton = wait.Until(d => d.FindElement(By.LinkText("Thêm mới")));
            createButton.Click();
            Thread.Sleep(1000);
            Console.WriteLine($"URL sau khi nhấn Thêm mới: {driver.Url}");
            ThuongHieuTestData testData = GetThuongHieuTestDataFromExcel("QLTH1_02", "QLTH");
            IWebElement tenThuongHieuInput = wait.Until(d => d.FindElement(By.Id("TenThuongHieu")));
            tenThuongHieuInput.Clear();
            tenThuongHieuInput.SendKeys(testData.TenThuongHieu);
            IWebElement saveButton = driver.FindElement(By.CssSelector("input.btn.btn-primary[value='Lưu']"));
            saveButton.Click();
            Thread.Sleep(1000);
            Console.WriteLine($"URL sau khi submit: {driver.Url}");
            bool found = false;
            while (!found)
            {
                var resultElements = wait.Until(d => d.FindElements(By.XPath($"//td[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{testData.TenThuongHieu.ToLower()}')]")));
                if (resultElements.Count > 0 && resultElements[0].Displayed)
                {
                    Assert.That(resultElements[0].Displayed, Is.True, $"Không tìm thấy thương hiệu '{testData.TenThuongHieu}' trong danh sách!");
                    Console.WriteLine($"Thương hiệu '{testData.TenThuongHieu}' đã được thêm thành công: {resultElements[0].Text}");
                    found = true;
                }
                else
                {
                    var nextPageButton = driver.FindElements(By.XPath("//li[@class='PagedList-skipToNext']/a[@rel='next']"));
                    if (nextPageButton.Count > 0 && nextPageButton[0].Displayed)
                    {
                        Console.WriteLine("Không tìm thấy thương hiệu trên trang hiện tại, chuyển sang trang tiếp theo...");
                        nextPageButton[0].Click();
                        Thread.Sleep(1000);
                        Console.WriteLine($"URL sau khi chuyển trang: {driver.Url}");
                    }
                    else
                    {
                        var tableRows = driver.FindElements(By.XPath("//table//tr"));
                        Console.WriteLine("Nội dung bảng danh sách thương hiệu:");
                        foreach (var row in tableRows)
                        {
                            Console.WriteLine($"Row content: {row.Text}");
                        }
                        Assert.Fail($"Không tìm thấy thương hiệu '{testData.TenThuongHieu}' trong danh sách sau khi duyệt qua tất cả các trang!");
                        break;
                    }
                }
            }
        }
        [Test]
        public void Test20_ThemThuongHieu_nullunValid()
        {
            Login("tuan159", "Tuan@159159");
            wait.Until(d => d.FindElement(By.Id("userMenuButton")).Displayed);
            driver.FindElement(By.Id("userMenuButton")).Click();
            driver.FindElement(By.Id("manageLink")).Click();
            driver.FindElement(By.LinkText("Quản lý thương hiệu")).Click();
            Thread.Sleep(1000);
            Console.WriteLine($"URL hiện tại: {driver.Url}");
            IWebElement createButton = wait.Until(d => d.FindElement(By.LinkText("Thêm mới")));
            createButton.Click();
            Thread.Sleep(1000);
            Console.WriteLine($"URL sau khi nhấn Thêm mới: {driver.Url}");
            ThuongHieuTestData testData = GetThuongHieuTestDataFromExcel("QLTH1_03", "QLTH");
            IWebElement tenThuongHieuInput = wait.Until(d => d.FindElement(By.Id("TenThuongHieu")));
            tenThuongHieuInput.Clear();
            tenThuongHieuInput.SendKeys(testData.TenThuongHieu);
            IWebElement saveButton = driver.FindElement(By.CssSelector("input.btn.btn-primary[value='Lưu']"));
            saveButton.Click();
            Thread.Sleep(1000);
            Console.WriteLine($"URL sau khi submit: {driver.Url}");
            bool found = false;
            while (!found)
            {
                var resultElements = wait.Until(d => d.FindElements(By.XPath($"//td[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{testData.TenThuongHieu.ToLower()}')]")));
                if (resultElements.Count > 0 && resultElements[0].Displayed)
                {
                    Assert.That(resultElements[0].Displayed, Is.True, $"Không tìm thấy thương hiệu '{testData.TenThuongHieu}' trong danh sách!");
                    Console.WriteLine($"Thương hiệu '{testData.TenThuongHieu}' đã được thêm thành công: {resultElements[0].Text}");
                    found = true;
                }
                else
                {
                    var nextPageButton = driver.FindElements(By.XPath("//li[@class='PagedList-skipToNext']/a[@rel='next']"));
                    if (nextPageButton.Count > 0 && nextPageButton[0].Displayed)
                    {
                        Console.WriteLine("Không tìm thấy thương hiệu trên trang hiện tại, chuyển sang trang tiếp theo...");
                        nextPageButton[0].Click();
                        Thread.Sleep(1000);
                        Console.WriteLine($"URL sau khi chuyển trang: {driver.Url}");
                    }
                    else
                    {
                        var tableRows = driver.FindElements(By.XPath("//table//tr"));
                        Console.WriteLine("Nội dung bảng danh sách thương hiệu:");
                        foreach (var row in tableRows)
                        {
                            Console.WriteLine($"Row content: {row.Text}");
                        }
                        Assert.Fail($"Không tìm thấy thương hiệu '{testData.TenThuongHieu}' trong danh sách sau khi duyệt qua tất cả các trang!");
                        break;
                    }
                }
            }
        }
        [Test]
        public void Test21_SuaThuongHieuHopLe()
        {
            Login("tuan159", "Tuan@159159");
            wait.Until(d => d.FindElement(By.Id("userMenuButton")).Displayed);
            driver.FindElement(By.Id("userMenuButton")).Click();
            driver.FindElement(By.Id("manageLink")).Click();
            driver.FindElement(By.LinkText("Quản lý thương hiệu")).Click();
            Thread.Sleep(1000);
            Console.WriteLine($"URL hiện tại: {driver.Url}");
            ThuongHieuTestData originalData = GetThuongHieuTestDataFromExcel("QLTH1_01", "QLTH");
            bool found = false;
            IWebElement editButton = null;
            while (!found)
            {
                var editButtons = wait.Until(d => d.FindElements(By.XPath($"//td[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{originalData.TenThuongHieu.ToLower()}')]/following-sibling::td/a[@class='btn btn-warning' and text()='Sửa']")));
                if (editButtons.Count > 0 && editButtons[0].Displayed)
                {
                    editButton = editButtons[0];
                    found = true;
                }
                else
                {
                    var nextPageButton = driver.FindElements(By.XPath("//li[@class='PagedList-skipToNext']/a[@rel='next']"));
                    if (nextPageButton.Count > 0 && nextPageButton[0].Displayed)
                    {
                        Console.WriteLine("Không tìm thấy thương hiệu trên trang hiện tại, chuyển sang trang tiếp theo...");
                        nextPageButton[0].Click();
                        Thread.Sleep(1000);
                        Console.WriteLine($"URL sau khi chuyển trang: {driver.Url}");
                    }
                    else
                    {
                        var tableRows = driver.FindElements(By.XPath("//table//tr"));
                        Console.WriteLine("Nội dung bảng danh sách thương hiệu:");
                        foreach (var row in tableRows)
                        {
                            Console.WriteLine($"Row content: {row.Text}");
                        }
                        Assert.Fail($"Không tìm thấy thương hiệu '{originalData.TenThuongHieu}' trong danh sách để chỉnh sửa!");
                        break;
                    }
                }
            }
            editButton.Click();
            Thread.Sleep(1000);
            Console.WriteLine($"URL sau khi nhấn Sửa: {driver.Url}");
            ThuongHieuTestData testData = GetThuongHieuTestDataFromExcel("QLTH2_01", "QLTH");
            IWebElement tenThuongHieuInput = wait.Until(d => d.FindElement(By.Id("TenThuongHieu")));
            tenThuongHieuInput.Clear();
            tenThuongHieuInput.SendKeys(testData.TenThuongHieu);
            IWebElement saveButton = driver.FindElement(By.CssSelector("button.btn.btn-success.mr-2"));
            saveButton.Click();
            Thread.Sleep(1000);
            Console.WriteLine($"URL sau khi submit: {driver.Url}");
            found = false;
            while (!found)
            {
                var resultElements = wait.Until(d => d.FindElements(By.XPath($"//td[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{testData.TenThuongHieu.ToLower()}')]")));
                if (resultElements.Count > 0 && resultElements[0].Displayed)
                {
                    Assert.That(resultElements[0].Displayed, Is.True, $"Không tìm thấy thương hiệu '{testData.TenThuongHieu}' trong danh sách sau khi sửa!");
                    Console.WriteLine($"Thương hiệu đã được sửa thành công thành '{testData.TenThuongHieu}': {resultElements[0].Text}");
                    found = true;
                }
                else
                {
                    var nextPageButton = driver.FindElements(By.XPath("//li[@class='PagedList-skipToNext']/a[@rel='next']"));
                    if (nextPageButton.Count > 0 && nextPageButton[0].Displayed)
                    {
                        Console.WriteLine("Không tìm thấy thương hiệu trên trang hiện tại, chuyển sang trang tiếp theo...");
                        nextPageButton[0].Click();
                        Thread.Sleep(1000);
                        Console.WriteLine($"URL sau khi chuyển trang: {driver.Url}");
                    }
                    else
                    {
                        var tableRows = driver.FindElements(By.XPath("//table//tr"));
                        Console.WriteLine("Nội dung bảng danh sách thương hiệu:");
                        foreach (var row in tableRows)
                        {
                            Console.WriteLine($"Row content: {row.Text}");
                        }
                        Assert.Fail($"Không tìm thấy thương hiệu '{testData.TenThuongHieu}' trong danh sách sau khi duyệt qua tất cả các trang!");
                        break;
                    }
                }
            }
        }
        [Test]
        public void Test22_ThêmThuongHieutrungtestkethop()
        {
            Login("tuan159", "Tuan@159159");
            wait.Until(d => d.FindElement(By.Id("userMenuButton")).Displayed);
            driver.FindElement(By.Id("userMenuButton")).Click();
            driver.FindElement(By.Id("manageLink")).Click();
            driver.FindElement(By.LinkText("Quản lý thương hiệu")).Click();
            Thread.Sleep(1000);
            Console.WriteLine($"URL hiện tại: {driver.Url}");

            IWebElement createButton = wait.Until(d => d.FindElement(By.LinkText("Thêm mới")));
            createButton.Click();
            Thread.Sleep(1000);
            Console.WriteLine($"URL sau khi nhấn Thêm mới: {driver.Url}");

            // Lấy dữ liệu test case QLTH1_01
            ThuongHieuTestData testData = GetThuongHieuTestDataFromExcel("QLTH2_05", "QLTH");

            IWebElement tenThuongHieuInput = wait.Until(d => d.FindElement(By.Id("TenThuongHieu")));
            tenThuongHieuInput.Clear();
            tenThuongHieuInput.SendKeys(testData.TenThuongHieu);
            IWebElement saveButton = driver.FindElement(By.CssSelector("input.btn.btn-primary[value='Lưu']"));
            saveButton.Click();
            Thread.Sleep(1000);
            Console.WriteLine($"URL sau khi submit: {driver.Url}");
            bool found = false;
            while (!found)
            {
                var resultElements = wait.Until(d => d.FindElements(By.XPath($"//td[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{testData.TenThuongHieu.ToLower()}')]")));
                if (resultElements.Count > 0 && resultElements[0].Displayed)
                {
                    Assert.That(resultElements[0].Displayed, Is.True, $"Không tìm thấy thương hiệu '{testData.TenThuongHieu}' trong danh sách!");
                    Console.WriteLine($"Thương hiệu '{testData.TenThuongHieu}' đã được thêm thành công: {resultElements[0].Text}");
                    found = true;
                }
                else
                {
                    var nextPageButton = driver.FindElements(By.XPath("//li[@class='PagedList-skipToNext']/a[@rel='next']"));
                    if (nextPageButton.Count > 0 && nextPageButton[0].Displayed)
                    {
                        Console.WriteLine("Không tìm thấy thương hiệu trên trang hiện tại, chuyển sang trang tiếp theo...");
                        nextPageButton[0].Click();
                        Thread.Sleep(1000);
                        Console.WriteLine($"URL sau khi chuyển trang: {driver.Url}");
                    }
                    else
                    {
                        var tableRows = driver.FindElements(By.XPath("//table//tr"));
                        Console.WriteLine("Nội dung bảng danh sách thương hiệu:");
                        foreach (var row in tableRows)
                        {
                            Console.WriteLine($"Row content: {row.Text}");
                        }
                        Assert.Fail($"Không tìm thấy thương hiệu '{testData.TenThuongHieu}' trong danh sách sau khi duyệt qua tất cả các trang!");
                        break;
                    }
                }
            }
        }
        [Test]
        public void Test23_XoaThuongHieuHopLe()
        {
            Login("tuan159", "Tuan@159159");
            wait.Until(d => d.FindElement(By.Id("userMenuButton")).Displayed);
            driver.FindElement(By.Id("userMenuButton")).Click();
            driver.FindElement(By.Id("manageLink")).Click();
            driver.FindElement(By.LinkText("Quản lý thương hiệu")).Click();
            Thread.Sleep(1000);
            Console.WriteLine($"URL hiện tại: {driver.Url}");
            ThuongHieuTestData testData = GetThuongHieuTestDataFromExcel("QLTH2_01", "QLTH");
            bool found = false;
            IWebElement deleteButton = null;
            while (!found)
            {
                var deleteButtons = wait.Until(d => d.FindElements(By.XPath($"//td[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{testData.TenThuongHieu.ToLower()}')]/following-sibling::td/a[@class='btn btn-danger' and text()='Xóa']")));
                if (deleteButtons.Count > 0 && deleteButtons[0].Displayed)
                {
                    deleteButton = deleteButtons[0];
                    found = true;
                }
                else
                {
                    var nextPageButton = driver.FindElements(By.XPath("//li[@class='PagedList-skipToNext']/a[@rel='next']"));
                    if (nextPageButton.Count > 0 && nextPageButton[0].Displayed)
                    {
                        Console.WriteLine("Không tìm thấy thương hiệu trên trang hiện tại, chuyển sang trang tiếp theo...");
                        nextPageButton[0].Click();
                        Thread.Sleep(1000);
                        Console.WriteLine($"URL sau khi chuyển trang: {driver.Url}");
                    }
                    else
                    {
                        var tableRows = driver.FindElements(By.XPath("//table//tr"));
                        Console.WriteLine("Nội dung bảng danh sách thương hiệu:");
                        foreach (var row in tableRows)
                        {
                            Console.WriteLine($"Row content: {row.Text}");
                        }
                        Assert.Fail($"Không tìm thấy thương hiệu '{testData.TenThuongHieu}' trong danh sách để xóa!");
                        break;
                    }
                }
            }
            deleteButton.Click();
            Thread.Sleep(1000);
            Console.WriteLine($"URL sau khi nhấn Xóa: {driver.Url}");
            IWebElement confirmDeleteButton = wait.Until(d => d.FindElement(By.CssSelector("button.btn.btn-danger")));
            confirmDeleteButton.Click();
            Thread.Sleep(1000);
            Console.WriteLine($"URL sau khi xác nhận xóa: {driver.Url}");
            bool stillExists = false;
            while (!stillExists)
            {
                var resultElements = wait.Until(d => d.FindElements(By.XPath($"//td[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{testData.TenThuongHieu.ToLower()}')]")));
                if (resultElements.Count > 0 && resultElements[0].Displayed)
                {
                    Assert.Fail($"Thương hiệu '{testData.TenThuongHieu}' vẫn còn trong danh sách sau khi xóa!");
                }
                else
                {
                    var nextPageButton = driver.FindElements(By.XPath("//li[@class='PagedList-skipToNext']/a[@rel='next']"));
                    if (nextPageButton.Count > 0 && nextPageButton[0].Displayed)
                    {
                        Console.WriteLine("Không tìm thấy thương hiệu trên trang hiện tại, chuyển sang trang tiếp theo...");
                        nextPageButton[0].Click();
                        Thread.Sleep(1000);
                        Console.WriteLine($"URL sau khi chuyển trang: {driver.Url}");
                    }
                    else
                    {
                        Console.WriteLine($"Thương hiệu '{testData.TenThuongHieu}' đã được xóa thành công!");
                        break;
                    }
                }
            }
        }
        //------------------------------TestCase Quản lý sản phẩm---------------------------------------------------------------
        // test case QLSP1_01 thêm sản phẩm mới hợp lệ
        [Test]
        public void Test24_CreateNewProductvalid()
        {
            Login("tuan159", "Tuan@159159");
            wait.Until(d => d.FindElement(By.Id("userMenuButton")).Displayed);
            driver.FindElement(By.Id("userMenuButton")).Click();
            driver.FindElement(By.Id("manageLink")).Click();
            driver.FindElement(By.LinkText("Quản lý sản phẩm")).Click();
            driver.FindElement(By.CssSelector("a.btn.btn-primary[href='/Admin/SanPham/Create']")).Click();
            wait.Until(d => d.FindElement(By.Id("TenSanPham")).Displayed);

            // Lấy dữ liệu từ Excel
            var testData = GetProductTestDataFromExcel("QLSP1_01", "QLSP"); //TestCaseId

            driver.FindElement(By.Id("TenSanPham")).SendKeys(testData.TenSanPham);
            driver.FindElement(By.Id("Gia")).SendKeys(testData.Gia);
            driver.FindElement(By.Id("MoTaSanPham")).SendKeys(testData.MoTaSanPham);
            driver.FindElement(By.Id("SoLuongTonKho")).SendKeys(testData.SoLuongTonKho);

            new SelectElement(driver.FindElement(By.Id("IDLoaiSanPham"))).SelectByText(testData.IDLoaiSanPham);
            new SelectElement(driver.FindElement(By.Id("IDMauSac"))).SelectByText(testData.IDMauSac);
            new SelectElement(driver.FindElement(By.Id("IDChatLieu"))).SelectByText(testData.IDChatLieu);
            new SelectElement(driver.FindElement(By.Id("IDThuongHieu"))).SelectByText(testData.IDThuongHieu);
            new SelectElement(driver.FindElement(By.Id("TrangThaiSanPham"))).SelectByText(testData.TrangThaiSanPham);

            driver.FindElement(By.Name("files")).SendKeys(testData.ImagePath);

            IWebElement submitButton = driver.FindElement(By.CssSelector("button.btn.btn-success"));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", submitButton);
            Thread.Sleep(500);

            driver.FindElement(By.CssSelector("button.btn.btn-success")).Click();
            // Kiểm tra sản phẩm có xuất hiện trong danh sách không
            var productElements = wait.Until(d => d.FindElements(By.XPath($"//td[contains(text(), '{testData.TenSanPham}')]")));
            if (productElements.Count > 0 && productElements[0].Displayed)
            {
                var productName = productElements[0];
                Assert.That(productName, Is.Not.Null, "Sản phẩm mới không được thêm vào danh sách!");
                Assert.That(productName.Displayed, Is.True, "Sản phẩm mới không hiển thị trong danh sách!");
            }
            else
            {
                Assert.Fail($"Không tìm thấy sản phẩm '{testData.TenSanPham}' trong danh sách sau thời gian chờ!");
            }
        }
        // testcase QLSP1_02 thêm sản phẩm mới có chứa thông tin trùng lặp
        [Test]
        public void Test25_CreateNewProductunvalid()
        {
            Login("tuan159", "Tuan@159159");
            wait.Until(d => d.FindElement(By.Id("userMenuButton")).Displayed);
            driver.FindElement(By.Id("userMenuButton")).Click();
            driver.FindElement(By.Id("manageLink")).Click();
            driver.FindElement(By.LinkText("Quản lý sản phẩm")).Click();
            driver.FindElement(By.CssSelector("a.btn.btn-primary[href='/Admin/SanPham/Create']")).Click();
            wait.Until(d => d.FindElement(By.Id("TenSanPham")).Displayed);

            // Lấy dữ liệu từ Excel
            var testData = GetProductTestDataFromExcel("QLSP1_02", "QLSP"); //TestCaseId

            driver.FindElement(By.Id("TenSanPham")).SendKeys(testData.TenSanPham);
            driver.FindElement(By.Id("Gia")).SendKeys(testData.Gia);
            driver.FindElement(By.Id("MoTaSanPham")).SendKeys(testData.MoTaSanPham);
            driver.FindElement(By.Id("SoLuongTonKho")).SendKeys(testData.SoLuongTonKho);

            new SelectElement(driver.FindElement(By.Id("IDLoaiSanPham"))).SelectByText(testData.IDLoaiSanPham);
            new SelectElement(driver.FindElement(By.Id("IDMauSac"))).SelectByText(testData.IDMauSac);
            new SelectElement(driver.FindElement(By.Id("IDChatLieu"))).SelectByText(testData.IDChatLieu);
            new SelectElement(driver.FindElement(By.Id("IDThuongHieu"))).SelectByText(testData.IDThuongHieu);
            new SelectElement(driver.FindElement(By.Id("TrangThaiSanPham"))).SelectByText(testData.TrangThaiSanPham);

            driver.FindElement(By.Name("files")).SendKeys(testData.ImagePath);

            IWebElement submitButton = driver.FindElement(By.CssSelector("button.btn.btn-success"));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", submitButton);
            Thread.Sleep(500);

            driver.FindElement(By.CssSelector("button.btn.btn-success")).Click();
            // Kiểm tra sản phẩm có xuất hiện trong danh sách không
            var productElements = wait.Until(d => d.FindElements(By.XPath($"//td[contains(text(), '{testData.TenSanPham}')]")));
            var productName = productElements[0];
            Assert.That(productName.Displayed, Is.False, "Sản phẩm mới hiển thị trong danh sách!");
        }
       
        //Testcase QLSP03_01 xóa sản phẩm hợp lệ
        [Test]
        public void Test28_DeleteProduct()
        {
            Login("tuan159", "Tuan@159159");
            wait.Until(d => d.FindElement(By.Id("userMenuButton")).Displayed);
            driver.FindElement(By.Id("userMenuButton")).Click();
            driver.FindElement(By.Id("manageLink")).Click();
            driver.FindElement(By.LinkText("Quản lý sản phẩm")).Click();

            Thread.Sleep(3000);

            // Lấy dữ liệu từ Excel cho TestCaseId = QLSP2_01
            var originalData = GetProductTestDataFromExcel("QLSP2_01", "QLSP");

            // Tìm sản phẩm để chỉnh sửa (dựa trên tên sản phẩm từ datatest)
            var productRows = wait.Until(d => d.FindElements(By.XPath($"//tr[td[contains(text(), '{originalData.TenSanPham}')]]")));
            if (productRows.Count > 0)
            {
                IWebElement productRow = productRows[0];
                IWebElement deleteButton = productRow.FindElement(By.CssSelector("a.btn.btn-danger"));
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", deleteButton);
                Thread.Sleep(1000);
                deleteButton.Click();
                Thread.Sleep(1000);
            }
            else
            {
                Assert.Fail($"Không tìm thấy sản phẩm '{originalData.TenSanPham}' trong danh sách để sửa!");
            }

            IAlert alert = wait.Until(d => driver.SwitchTo().Alert());
            alert.Accept();

            wait.Until(d => d.Url.Contains("/admin/sanpham/Delete/"));
            Console.WriteLine($"URL trang chi tiết xóa: {driver.Url}");

            IWebElement finalDeleteButton = wait.Until(d => d.FindElement(By.CssSelector("button.btn.btn-danger")));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", finalDeleteButton);
            Thread.Sleep(500);
            finalDeleteButton.Click();
            Thread.Sleep(1000);

            // Kiểm tra sản phẩm đã bị xóa
            var productElements = driver.FindElements(By.XPath("//td[contains(text(), 'Nhẫn Bạc 925 Cao Cấp')]"));
            if (productElements.Count > 0)
            {
                Assert.Fail("Sản phẩm 'Nhẫn Bạc 925 Cao Cấp' vẫn tồn tại trong danh sách sau khi xóa!");
            }
            else
            {
                Assert.Pass("Sản phẩm 'Nhẫn Bạc 925 Cao Cấp' đã được xóa thành công!");
            }
        }
        // Testcase QLSP2_01 sửa thông tin sản phẩm hợp lệ
        [Test]
        public void Test26_EditProduct()
        {
            // Đăng nhập
            Login("tuan159", "Tuan@159159");
            // Chờ và click vào userMenuButton
            wait.Until(d => d.FindElement(By.Id("userMenuButton")).Displayed);
            driver.FindElement(By.Id("userMenuButton")).Click();

            driver.FindElement(By.Id("manageLink")).Click();
            driver.FindElement(By.LinkText("Quản lý sản phẩm")).Click();

            Thread.Sleep(1000);
            Console.WriteLine($"URL hiện tại: {driver.Url}");

            // Lấy dữ liệu từ Excel cho TestCaseId = QLSP1_01
            var originalData = GetProductTestDataFromExcel("QLSP1_01", "QLSP");

            // Tìm sản phẩm để chỉnh sửa (dựa trên tên sản phẩm từ datatest)
            var productRows = wait.Until(d => d.FindElements(By.XPath($"//tr[td[contains(text(), '{originalData.TenSanPham}')]]")));
            if (productRows.Count > 0)
            {
                IWebElement productRow = productRows[0];
                IWebElement editButton = productRow.FindElement(By.CssSelector("a.btn.btn-primary"));
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", editButton);
                Thread.Sleep(500);
                editButton.Click();
            }
            else
            {
                Assert.Fail($"Không tìm thấy sản phẩm '{originalData.TenSanPham}' trong danh sách để sửa!");
            }

            // Đảm bảo form chỉnh sửa đã hiển thị
            wait.Until(d => d.FindElement(By.Id("TenSanPham")).Displayed);
            Console.WriteLine($"URL trang chỉnh sửa: {driver.Url}");

            // Lấy dữ liệu từ Excel cho TestCaseId = QLSP2_01
            var testData = GetProductTestDataFromExcel("QLSP2_01", "QLSP");

            // Điền thông tin vào form chỉnh sửa
            IWebElement tenSanPham = driver.FindElement(By.Id("TenSanPham"));
            tenSanPham.Clear();
            tenSanPham.SendKeys(testData.TenSanPham);

            IWebElement gia = driver.FindElement(By.Id("Gia"));
            gia.Clear();
            gia.SendKeys(testData.Gia);

            IWebElement moTa = driver.FindElement(By.Id("MoTaSanPham"));
            moTa.Clear();
            moTa.SendKeys(testData.MoTaSanPham);

            IWebElement soLuong = driver.FindElement(By.Id("SoLuongTonKho"));
            soLuong.Clear();
            soLuong.SendKeys(testData.SoLuongTonKho);

            new SelectElement(driver.FindElement(By.Id("IDLoaiSanPham"))).SelectByText(testData.IDLoaiSanPham);
            new SelectElement(driver.FindElement(By.Id("IDMauSac"))).SelectByText(testData.IDMauSac);
            new SelectElement(driver.FindElement(By.Id("IDChatLieu"))).SelectByText(testData.IDChatLieu);
            new SelectElement(driver.FindElement(By.Id("IDThuongHieu"))).SelectByText(testData.IDThuongHieu);
            new SelectElement(driver.FindElement(By.Id("TrangThaiSanPham"))).SelectByText(testData.TrangThaiSanPham);

            // Upload hình ảnh
            driver.FindElement(By.Name("files")).SendKeys(testData.ImagePath);

            // Submit form
            IWebElement submitButton = driver.FindElement(By.CssSelector("button.btn.btn-success"));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", submitButton);
            Thread.Sleep(500);
            submitButton.Click();

            Thread.Sleep(1000);

            // Kiểm tra sản phẩm đã được cập nhật đúng
            var updatedProductElements = wait.Until(d => d.FindElements(By.XPath($"//td[contains(text(), '{testData.TenSanPham}')]")));
            if (updatedProductElements.Count > 0 && updatedProductElements[0].Displayed)
            {
                IWebElement updatedProduct = updatedProductElements[0];
                var updatedPriceElements = driver.FindElements(By.XPath($"//tr[td[contains(text(), '{testData.TenSanPham}')]]/td[contains(text(), '{Convert.ToInt32(testData.Gia).ToString("N0")}')]"));
                if (updatedPriceElements.Count > 0 && updatedPriceElements[0].Displayed)
                {
                    Assert.That(updatedProduct.Displayed, Is.True, $"Tên sản phẩm mới '{testData.TenSanPham}' không hiển thị trong danh sách!");
                    Assert.That(updatedPriceElements[0].Displayed, Is.True, $"Giá sản phẩm mới '{testData.Gia}' không hiển thị trong danh sách!");
                }
                else
                {
                    Assert.Fail($"Không tìm thấy giá sản phẩm đã cập nhật '{testData.Gia}' trong danh sách sau thời gian chờ!");
                }
            }
            else
            {
                Assert.Fail($"Không tìm thấy sản phẩm đã cập nhật '{testData.TenSanPham}' trong danh sách sau thời gian chờ!");
            }
        }
        // Testcase QLSP4_01 tìm kiếm thông tin sản phẩm hợp lệ
        [Test]
        public void Test27_timkiemProduct()
        {
            Login("tuan159", "Tuan@159159");
            // Chờ và click vào userMenuButton
            wait.Until(d => d.FindElement(By.Id("userMenuButton")).Displayed);
            driver.FindElement(By.Id("userMenuButton")).Click();

            driver.FindElement(By.Id("manageLink")).Click();
            driver.FindElement(By.LinkText("Quản lý sản phẩm")).Click();
            Thread.Sleep(1000);
            Console.WriteLine($"URL hiện tại: {driver.Url}");

            // Lấy dữ liệu test case QLSP4_01
            ProductTestData testData = GetProductTestDataFromExcel("QLSP4_01", "QLSP");
            string searchKeyword = testData.SearchKeyword;
            Console.WriteLine($"Từ khóa tìm kiếm từ file Excel: {searchKeyword}");
            IWebElement searchInput = wait.Until(d => d.FindElement(By.Id("SearchString")));
            searchInput.SendKeys(searchKeyword);
            IWebElement searchButton = driver.FindElement(By.CssSelector("input.btn.btn-secondary[value='Search']"));
            searchButton.Click();

            Thread.Sleep(1000);
            Console.WriteLine($"URL sau khi tìm kiếm: {driver.Url}");

            // Kiểm tra kết quả
            var tableRows = driver.FindElements(By.XPath("//table//tr"));
            Console.WriteLine("Nội dung bảng kết quả tìm kiếm:");
            foreach (var row in tableRows)
            {
                Console.WriteLine($"Row content: {row.Text}");
            }

            // Điều chỉnh XPath để tìm kiếm không phân biệt hoa thường, hỗ trợ tiếng Việt
            var resultElements = wait.Until(d => d.FindElements(By.XPath($"//td[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZÀÁÂÃÈÉÊÌÍÒÓÔÕÙÚÝĂĐĨŨƠƯ', 'abcdefghijklmnopqrstuvwxyzàáâãèéêìíòóôõùúýăđĩũơư'), '{searchKeyword.ToLower()}')]")));
            if (resultElements.Count > 0 && resultElements[0].Displayed)
            {
                IWebElement result = resultElements[0];
                Assert.That(result.Displayed, Is.True, $"Không tìm thấy sản phẩm nào liên quan đến '{searchKeyword}' trong kết quả!");
                Console.WriteLine($"Sản phẩm tìm thấy: {result.Text}");
            }
            else
            {
                var noResultElements = driver.FindElements(By.XPath("//td[contains(text(), 'Không có bản ghi nào')]"));
                if (noResultElements.Count > 0 && noResultElements[0].Displayed)
                {
                    Assert.That(noResultElements[0].Displayed, Is.True, "Danh sách trống nhưng không hiển thị thông báo 'Không có bản ghi nào'!");
                    Console.WriteLine($"Không tìm thấy sản phẩm nào liên quan đến '{searchKeyword}'.");
                }
                else
                {
                    Assert.Fail("Tìm kiếm thất bại: Không có kết quả và không hiển thị thông báo 'Không có bản ghi nào'!");
                }
            }
        }
        // Testcase QLSP4_02 tìm kiếm thông tin sản phẩm đã xóa
        [Test]
        public void Test29_timkiemDeleteProductunValid()
        {
            Login("tuan159", "Tuan@159159");
            // Chờ và click vào userMenuButton
            wait.Until(d => d.FindElement(By.Id("userMenuButton")).Displayed);
            driver.FindElement(By.Id("userMenuButton")).Click();

            driver.FindElement(By.Id("manageLink")).Click();
            driver.FindElement(By.LinkText("Quản lý sản phẩm")).Click();
            Thread.Sleep(1000);
            Console.WriteLine($"URL hiện tại: {driver.Url}");

            // Lấy dữ liệu test case QLSP4_02
            ProductTestData testData = GetProductTestDataFromExcel("QLSP4_02", "QLSP");
            string searchKeyword = testData.SearchKeyword;
            Console.WriteLine($"Từ khóa tìm kiếm từ file Excel: {searchKeyword}");
            IWebElement searchInput = wait.Until(d => d.FindElement(By.Id("SearchString")));
            searchInput.SendKeys(searchKeyword);
            IWebElement searchButton = driver.FindElement(By.CssSelector("input.btn.btn-secondary[value='Search']"));
            searchButton.Click();

            Thread.Sleep(1000);
            Console.WriteLine($"URL sau khi tìm kiếm: {driver.Url}");

            // Kiểm tra kết quả
            var tableRows = driver.FindElements(By.XPath("//table//tr"));
            Console.WriteLine("Nội dung bảng kết quả tìm kiếm:");
            foreach (var row in tableRows)
            {
                Console.WriteLine($"Row content: {row.Text}");
            }

            var resultElements = wait.Until(d => d.FindElements(By.XPath($"//td[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZÀÁÂÃÈÉÊÌÍÒÓÔÕÙÚÝĂĐĨŨƠƯ', 'abcdefghijklmnopqrstuvwxyzàáâãèéêìíòóôõùúýăđĩũơư'), '{searchKeyword.ToLower()}')]")));
            if (resultElements.Count > 0 && resultElements[0].Displayed)
            {
                IWebElement result = resultElements[0];
                Assert.That(result.Displayed, Is.True, $"Không tìm thấy sản phẩm nào liên quan đến '{searchKeyword}' trong kết quả!");
                Console.WriteLine($"Sản phẩm tìm thấy: {result.Text}");
            }
            else
            {
                var noResultElements = driver.FindElements(By.XPath("//td[contains(text(), 'Không có bản ghi nào')]"));
                if (noResultElements.Count > 0 && noResultElements[0].Displayed)
                {
                    Assert.That(noResultElements[0].Displayed, Is.True, "Danh sách trống nhưng không hiển thị thông báo 'Không có bản ghi nào'!");
                    Console.WriteLine($"Không tìm thấy sản phẩm nào liên quan đến '{searchKeyword}'.");
                }
                else
                {
                    Assert.Fail("Tìm kiếm thất bại: Không có kết quả và không hiển thị thông báo 'Không có bản ghi nào'!");
                }
            }
        }

        [Test]
        // Testcase QLDM3_02: Xóa danh mục có sản phẩm gắn liền
        public void Test30_DeleteCategory_unValid()
        {
            Login("tuan159", "Tuan@159159");
            wait.Until(d => d.FindElement(By.Id("userMenuButton")).Displayed);
            driver.FindElement(By.Id("userMenuButton")).Click();
            driver.FindElement(By.Id("manageLink")).Click();
            driver.FindElement(By.LinkText("Quản lý danh mục")).Click();

            // Lấy dữ liệu từ Excel cho danh mục cần xóa (QLDM2_01)
            var testData = GetCategoryTestDataFromExcel("QLDM2_01", "QLDM");
            Console.WriteLine($"Danh mục cần xóa: {testData.TenLoaiSanPham}");

            IWebElement categoryRow = null;
            bool categoryFound = false;
            while (!categoryFound)
            {
                wait.Until(d => d.FindElement(By.CssSelector("table.table-hover")).Displayed);
                var rows = driver.FindElements(By.XPath("//table[@class='table table-hover']/tbody/tr"));

                Console.WriteLine("Danh mục trên trang hiện tại:");
                foreach (var row in rows)
                {
                    var nameElement = row.FindElement(By.XPath("./td[2]"));
                    Console.WriteLine($"- {nameElement.Text}");
                }

                // Tìm hàng chứa danh mục
                var matchingRow = rows.FirstOrDefault(row =>
                    row.FindElement(By.XPath("./td[2]")).Text.Trim().Equals(testData.TenLoaiSanPham.Trim(), StringComparison.OrdinalIgnoreCase));

                if (matchingRow != null)
                {
                    categoryRow = matchingRow;
                    categoryFound = true;
                    break;
                }

                var nextButton = driver.FindElements(By.CssSelector("li.PagedList-skipToNext a[rel='next']"));
                if (nextButton.Count == 0 || !nextButton[0].Displayed)
                {
                    Assert.Fail($"Không tìm thấy danh mục '{testData.TenLoaiSanPham}' trong danh sách để xóa.");
                    break;
                }

                new Actions(driver)
                    .MoveToElement(nextButton[0])
                    .Pause(TimeSpan.FromMilliseconds(500))
                    .Click()
                    .Perform();

                wait.Until(d => d.FindElement(By.CssSelector("table.table-hover")).Displayed);
            }

            IWebElement deleteButton = categoryRow.FindElement(By.CssSelector("a.btn.btn-danger"));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", deleteButton);
            Thread.Sleep(500);
            deleteButton.Click();

            IAlert alert = wait.Until(d => driver.SwitchTo().Alert());
            alert.Accept();
            wait.Until(d => d.FindElement(By.XPath("//button[@class='btn btn-danger' and text()='Xóa']")).Displayed);
            IWebElement confirmDeleteButton = driver.FindElement(By.XPath("//button[@class='btn btn-danger' and text()='Xóa']"));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", confirmDeleteButton);
            Thread.Sleep(500);
            confirmDeleteButton.Click();
            wait.Until(d => !d.FindElements(By.XPath($"//td[contains(text(), '{testData.TenLoaiSanPham}')]")).Any());
            Assert.Fail($"Danh mục '{testData.TenLoaiSanPham}' đã được xóa thành công!");
        }

        // Testcase QLTH3_02: Xóa thuong hiệu có sản phẩm gắn liền
        [Test]
        public void Test31_XoaThuongHieuunvalid()
        {
            Login("tuan159", "Tuan@159159");
            wait.Until(d => d.FindElement(By.Id("userMenuButton")).Displayed);
            driver.FindElement(By.Id("userMenuButton")).Click();
            driver.FindElement(By.Id("manageLink")).Click();
            driver.FindElement(By.LinkText("Quản lý thương hiệu")).Click();
            Thread.Sleep(1000);
            Console.WriteLine($"URL hiện tại: {driver.Url}");
            ThuongHieuTestData testData = GetThuongHieuTestDataFromExcel("QLTH2_05", "QLTH");
            bool found = false;
            IWebElement deleteButton = null;
            while (!found)
            {
                var deleteButtons = wait.Until(d => d.FindElements(By.XPath($"//td[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{testData.TenThuongHieu.ToLower()}')]/following-sibling::td/a[@class='btn btn-danger' and text()='Xóa']")));
                if (deleteButtons.Count > 0 && deleteButtons[0].Displayed)
                {
                    deleteButton = deleteButtons[0];
                    found = true;
                }
                else
                {
                    var nextPageButton = driver.FindElements(By.XPath("//li[@class='PagedList-skipToNext']/a[@rel='next']"));
                    if (nextPageButton.Count > 0 && nextPageButton[0].Displayed)
                    {
                        Console.WriteLine("Không tìm thấy thương hiệu trên trang hiện tại, chuyển sang trang tiếp theo...");
                        nextPageButton[0].Click();
                        Thread.Sleep(1000);
                        Console.WriteLine($"URL sau khi chuyển trang: {driver.Url}");
                    }
                    else
                    {
                        var tableRows = driver.FindElements(By.XPath("//table//tr"));
                        Console.WriteLine("Nội dung bảng danh sách thương hiệu:");
                        foreach (var row in tableRows)
                        {
                            Console.WriteLine($"Row content: {row.Text}");
                        }
                        Assert.Fail($"Không tìm thấy thương hiệu '{testData.TenThuongHieu}' trong danh sách để xóa!");
                        break;
                    }
                }
            }
            deleteButton.Click();
            Thread.Sleep(1000);
            IWebElement confirmDeleteButton = wait.Until(d => d.FindElement(By.CssSelector("button.btn.btn-danger")));
            confirmDeleteButton.Click();
            Thread.Sleep(1000);
            bool stillExists = false;
            while (!stillExists)
            {
                var resultElements = wait.Until(d => d.FindElements(By.XPath($"//td[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{testData.TenThuongHieu.ToLower()}')]")));
                if (resultElements.Count > 0 && resultElements[0].Displayed)
                {
                    Console.WriteLine($"Thương hiệu '{testData.TenThuongHieu}' vẫn tồn tại trong danh sách!");
                    Assert.Pass($"Thương hiệu '{testData.TenThuongHieu}' được tìm thấy thành công trong danh sách!");
                }
                else
                {
                    var nextPageButton = driver.FindElements(By.XPath("//li[@class='PagedList-skipToNext']/a[@rel='next']"));
                    if (nextPageButton.Count > 0 && nextPageButton[0].Displayed)
                    {
                        Console.WriteLine("Không tìm thấy thương hiệu trên trang hiện tại, chuyển sang trang tiếp theo...");
                        nextPageButton[0].Click();
                        Thread.Sleep(1000);
                    }
                    else
                    {
                        Console.WriteLine($"Không tìm thấy thương hiệu '{testData.TenThuongHieu}' sau khi duyệt qua tất cả các trang!");
                        Assert.Fail($"Thương hiệu '{testData.TenThuongHieu}' không được tìm thấy trong danh sách!");
                        break;
                    }
                }
            }
        }

        //testcase QLSP1_03 Thêm sản phẩm bỏ trống thông tin
        [Test]
        public void Test32_CreateNewProduct_nullunvalid()
        {
            Login("tuan159", "Tuan@159159");
            // Chờ và click vào userMenuButton
            wait.Until(d => d.FindElement(By.Id("userMenuButton")).Displayed);
            driver.FindElement(By.Id("userMenuButton")).Click();

            driver.FindElement(By.Id("manageLink")).Click();
            driver.FindElement(By.LinkText("Quản lý sản phẩm")).Click();

            Thread.Sleep(1000);
            driver.FindElement(By.CssSelector("a.btn.btn-primary[href='/Admin/SanPham/Create']")).Click();

            wait.Until(d => d.Url.Contains("/Admin/SanPham/Create"));
            Console.WriteLine($"URL trang thêm mới: {driver.Url}");
            wait.Until(d => d.FindElement(By.Id("TenSanPham")).Displayed);

            IWebElement submitButton = driver.FindElement(By.CssSelector("button.btn.btn-success"));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", submitButton);
            Thread.Sleep(500);
            submitButton.Click();

            // Kiểm tra thông báo lỗi validation
            var errorMessages = wait.Until(d => d.FindElements(By.CssSelector(".text-danger")));
            if (errorMessages.Count > 0 && errorMessages.Any(e => e.Displayed && !string.IsNullOrWhiteSpace(e.Text)))
            {
                Assert.That(errorMessages.Count, Is.GreaterThan(0), "Không hiển thị thông báo lỗi nào khi để trống thông tin!");
                Console.WriteLine("Các thông báo lỗi validation:");
                foreach (var error in errorMessages)
                {
                    if (!string.IsNullOrWhiteSpace(error.Text))
                    {
                        Console.WriteLine(error.Text);
                    }
                }
                Assert.That(driver.Url, Does.Contain("/Admin/SanPham/Create"), "Trang chuyển hướng không mong muốn khi thông tin trống!");
            }
            else
            {
                Assert.Fail("Không hiển thị thông báo lỗi validation khi để trống thông tin bắt buộc!");
            }
        }
        [TearDown]
        public void TearDown()
        {
            if (driver != null)
            {
                driver.Quit();
                driver.Dispose();
            }
        }
    }
}