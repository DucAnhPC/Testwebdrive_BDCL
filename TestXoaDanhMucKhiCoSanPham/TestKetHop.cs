using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.IO;
using ExcelDataReader;
using System.Data;
using System.Text;
using NUnit.Framework.Interfaces;

namespace TestXoaDanhMucKhiCoSanPham
{
    public class Tests
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
        private void Login(string username, string password)
        {
            driver.Navigate().GoToUrl("http://localhost:55119/Auth");
            driver.FindElement(By.CssSelector("#email")).SendKeys(username);
            driver.FindElement(By.CssSelector("#password")).SendKeys(password);
            driver.FindElement(By.CssSelector("#dangnhap")).Click();
            Thread.Sleep(3000);
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
            // Đăng ký encoding để đọc file Excel
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // Đường dẫn tới file Excel (thay đổi theo thực tế)
            string filePath = @"D:\TestData.xlsx";
            TaiKhoanTestData testData = new TaiKhoanTestData();

            // Kiểm tra file có tồn tại không
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

            string filePath = @"D:\TestData.xlsx"; // Đường dẫn tới file Excel
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
        //testcase TESTDM_01 Kiểm Tra xóa danh mục khi có sản phẩm liên kết
        [Test]
        public void Test_kethopxoadanhmuccosanpham()
        {
            Login("tuan159", "Tuan@159159");
            wait.Until(d => d.FindElement(By.Id("userMenuButton")).Displayed);
            driver.FindElement(By.Id("userMenuButton")).Click();
            driver.FindElement(By.Id("manageLink")).Click();
            wait.Until(d => d.FindElement(By.XPath("//a[@href='/admin/loaisanpham/index']")).Displayed);
            driver.FindElement(By.XPath("//a[@href='/admin/loaisanpham/index']")).Click();
            driver.FindElement(By.CssSelector("a.btn.btn-primary[href='/Admin/LoaiSanPham/Create']")).Click();
            wait.Until(d => d.FindElement(By.Id("TenLoaiSanPham")).Displayed);

            var categoryData = GetCategoryTestDataFromExcel("QLDM1_01", "QLDM");
            string categoryName = categoryData.TenLoaiSanPham;
            driver.FindElement(By.Id("TenLoaiSanPham")).SendKeys(categoryName);
            driver.FindElement(By.Name("file")).SendKeys(categoryData.ImagePath);

            IWebElement submitCategoryButton = driver.FindElement(By.CssSelector("button.btn.btn-success"));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", submitCategoryButton);
            Thread.Sleep(500);
            submitCategoryButton.Click();
            try
            {
                var addedCategory = wait.Until(d => d.FindElement(By.XPath($"//td[contains(text(), '{categoryName}')]")));
                Assert.That(addedCategory, Is.Not.Null, $"Danh mục mới '{categoryName}' không được thêm vào danh sách!");
                Assert.That(addedCategory.Displayed, Is.True, $"Danh mục mới '{categoryName}' không hiển thị trong danh sách!");
            }
            catch (WebDriverTimeoutException)
            {
                Assert.Fail($"Không tìm thấy danh mục '{categoryName}' trong danh sách sau khi thêm!");
            }

            //thêm sản phẩm mới sử dụng danh mục vừa tạo
            wait.Until(d => d.FindElement(By.XPath("//a[@href='/admin/sanpham/index']")).Displayed);
            driver.FindElement(By.XPath("//a[@href='/admin/sanpham/index']")).Click();
            driver.FindElement(By.CssSelector("a.btn.btn-primary[href='/Admin/SanPham/Create']")).Click();
            wait.Until(d => d.FindElement(By.Id("TenSanPham")).Displayed);

            var productData = GetProductTestDataFromExcel("TESTDM_01", "QLSP");
            driver.FindElement(By.Id("TenSanPham")).SendKeys(productData.TenSanPham);
            driver.FindElement(By.Id("Gia")).SendKeys(productData.Gia);
            driver.FindElement(By.Id("MoTaSanPham")).SendKeys(productData.MoTaSanPham);
            driver.FindElement(By.Id("SoLuongTonKho")).SendKeys(productData.SoLuongTonKho);

            new SelectElement(driver.FindElement(By.Id("IDLoaiSanPham"))).SelectByText(categoryName);
            new SelectElement(driver.FindElement(By.Id("IDMauSac"))).SelectByText(productData.IDMauSac);
            new SelectElement(driver.FindElement(By.Id("IDChatLieu"))).SelectByText(productData.IDChatLieu);
            new SelectElement(driver.FindElement(By.Id("IDThuongHieu"))).SelectByText(productData.IDThuongHieu);
            new SelectElement(driver.FindElement(By.Id("TrangThaiSanPham"))).SelectByText(productData.TrangThaiSanPham);

            driver.FindElement(By.Name("files")).SendKeys(productData.ImagePath);

            IWebElement submitProductButton = driver.FindElement(By.CssSelector("button.btn.btn-success"));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", submitProductButton);
            Thread.Sleep(500);
            submitProductButton.Click();

            // Kiểm tra sản phẩm đã được thêm
            try
            {
                var addedProduct = wait.Until(d => d.FindElement(By.XPath($"//td[contains(text(), '{productData.TenSanPham}')]")));
                Assert.That(addedProduct, Is.Not.Null, $"Sản phẩm mới '{productData.TenSanPham}' không được thêm vào danh sách!");
                Assert.That(addedProduct.Displayed, Is.True, $"Sản phẩm mới '{productData.TenSanPham}' không hiển thị trong danh sách!");
            }
            catch (WebDriverTimeoutException)
            {
                Assert.Fail($"Không tìm thấy sản phẩm '{productData.TenSanPham}' trong danh sách sau khi thêm!");
            }

            //Thử xóa danh mục vừa tạo
            wait.Until(d => d.FindElement(By.XPath("//a[@href='/admin/loaisanpham/index']")).Displayed);
            driver.FindElement(By.XPath("//a[@href='/admin/loaisanpham/index']")).Click();

            try
            {
                IWebElement categoryRow = wait.Until(d => d.FindElement(By.XPath($"//tr[td[contains(text(), '{categoryName}')]]")));
                IWebElement deleteButton = categoryRow.FindElement(By.CssSelector("a.btn.btn-danger"));
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", deleteButton);
                Thread.Sleep(500);
                deleteButton.Click();

                try
                {
                    IAlert alert = wait.Until(d => driver.SwitchTo().Alert());
                    alert.Accept();
                }
                catch (WebDriverTimeoutException)
                {
                    Console.WriteLine("Không có popup xác nhận xóa.");
                }

                try
                {
                    wait.Until(d => d.FindElement(By.XPath("//button[@class='btn btn-danger' and text()='Xóa']")).Displayed);
                    IWebElement confirmDeleteButton = driver.FindElement(By.XPath("//button[@class='btn btn-danger' and text()='Xóa']"));
                    ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", confirmDeleteButton);
                    Thread.Sleep(500);
                    confirmDeleteButton.Click();
                }
                catch (WebDriverTimeoutException)
                {
                    Assert.Fail("Không tìm thấy nút 'Xóa' trên trang xác nhận sau khi xử lý popup!");
                }

                //Kiểm tra xem danh mục có bị xóa hay không
                try
                {
                    wait.Until(d => !d.FindElements(By.XPath($"//td[contains(text(), '{categoryName}')]")).Any());
                    Assert.Fail($"Danh mục '{categoryName}' đã bị xóa dù có sản phẩm liên kết!");
                }
                catch (WebDriverTimeoutException)
                {
                    Assert.Pass($"Danh mục '{categoryName}' không bị xóa khi có sản phẩm liên kết!");
                }
            }
            catch (WebDriverTimeoutException)
            {
                Assert.Fail($"Không tìm thấy danh mục '{categoryName}' trong danh sách để xóa!");
            }

        }

        //testcase TESTTH_01 Kiểm Tra xóa thương hiệu khi có sản phẩm liên kết
        [Test]
        public void Test_kethopxoathuonghieucosanpham()
        {
            Login("tuan159", "Tuan@159159");
            wait.Until(d => d.FindElement(By.Id("userMenuButton")).Displayed);

            driver.FindElement(By.Id("userMenuButton")).Click();
            driver.FindElement(By.Id("manageLink")).Click();
            wait.Until(d => d.FindElement(By.XPath("//a[@href='/admin/thuonghieu/index']")).Displayed);
            driver.FindElement(By.XPath("//a[@href='/admin/thuonghieu/index']")).Click();
            IWebElement createButton = wait.Until(d => d.FindElement(By.LinkText("Thêm mới")));
            createButton.Click();

            ThuongHieuTestData thuonghieuData = GetThuongHieuTestDataFromExcel("QLTH2_05", "QLTH");
            IWebElement tenThuongHieuInput = wait.Until(d => d.FindElement(By.Id("TenThuongHieu")));
            tenThuongHieuInput.Clear();
            tenThuongHieuInput.SendKeys(thuonghieuData.TenThuongHieu);
            IWebElement saveButton = driver.FindElement(By.CssSelector("input.btn.btn-primary[value='Lưu']"));
            saveButton.Click();

            // Kiểm tra xem thương hiệu mới có xuất hiện trong danh sách không
            bool thuongHieuFound = false;
            while (!thuongHieuFound)
            {
                try
                {
                    IWebElement result = wait.Until(d => d.FindElement(By.XPath($"//td[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{thuonghieuData.TenThuongHieu.ToLower()}')]")));
                    Assert.That(result.Displayed, Is.True, $"Không tìm thấy thương hiệu '{thuonghieuData.TenThuongHieu}' trong danh sách!");
                    Console.WriteLine($"Thương hiệu '{thuonghieuData.TenThuongHieu}' đã được thêm thành công: {result.Text}");
                    thuongHieuFound = true;
                }
                catch (WebDriverTimeoutException)
                {
                    try
                    {
                        IWebElement nextPageButton = driver.FindElement(By.XPath("//li[@class='PagedList-skipToNext']/a[@rel='next']"));
                        if (nextPageButton != null && nextPageButton.Displayed)
                        {
                            Console.WriteLine("Không tìm thấy thương hiệu trên trang hiện tại, chuyển sang trang tiếp theo...");
                            nextPageButton.Click();
                            Thread.Sleep(1000);
                        }
                        else
                        {
                            Assert.Fail($"Không tìm thấy thương hiệu '{thuonghieuData.TenThuongHieu}' trong danh sách sau khi duyệt qua tất cả các trang!");
                            break;
                        }
                    }
                    catch (NoSuchElementException)
                    {
                        Assert.Fail($"Không tìm thấy thương hiệu '{thuonghieuData.TenThuongHieu}' trong danh sách và không còn trang để chuyển!");
                        break;
                    }
                }
            }

            //Thêm sản phẩm mới sử dụng thương hiệu vừa tạo
            wait.Until(d => d.FindElement(By.XPath("//a[@href='/admin/sanpham/index']")).Displayed);
            driver.FindElement(By.XPath("//a[@href='/admin/sanpham/index']")).Click();
            driver.FindElement(By.CssSelector("a.btn.btn-primary[href='/Admin/SanPham/Create']")).Click();
            wait.Until(d => d.FindElement(By.Id("TenSanPham")).Displayed);

            var productData = GetProductTestDataFromExcel("TESTTH_01", "QLSP");
            driver.FindElement(By.Id("TenSanPham")).SendKeys(productData.TenSanPham);
            driver.FindElement(By.Id("Gia")).SendKeys(productData.Gia);
            driver.FindElement(By.Id("MoTaSanPham")).SendKeys(productData.MoTaSanPham);
            driver.FindElement(By.Id("SoLuongTonKho")).SendKeys(productData.SoLuongTonKho);

            new SelectElement(driver.FindElement(By.Id("IDLoaiSanPham"))).SelectByText(productData.IDLoaiSanPham);
            new SelectElement(driver.FindElement(By.Id("IDMauSac"))).SelectByText(productData.IDMauSac);
            new SelectElement(driver.FindElement(By.Id("IDChatLieu"))).SelectByText(productData.IDChatLieu);
            new SelectElement(driver.FindElement(By.Id("IDThuongHieu"))).SelectByText(thuonghieuData.TenThuongHieu);
            new SelectElement(driver.FindElement(By.Id("TrangThaiSanPham"))).SelectByText(productData.TrangThaiSanPham);

            driver.FindElement(By.Name("files")).SendKeys(productData.ImagePath);

            IWebElement submitProductButton = driver.FindElement(By.CssSelector("button.btn.btn-success"));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", submitProductButton);
            Thread.Sleep(500);
            submitProductButton.Click();

            // Kiểm tra sản phẩm đã được thêm
            try
            {
                var addedProduct = wait.Until(d => d.FindElement(By.XPath($"//td[contains(text(), '{productData.TenSanPham}')]")));
                Assert.That(addedProduct, Is.Not.Null, $"Sản phẩm mới '{productData.TenSanPham}' không được thêm vào danh sách!");
                Assert.That(addedProduct.Displayed, Is.True, $"Sản phẩm mới '{productData.TenSanPham}' không hiển thị trong danh sách!");
            }
            catch (WebDriverTimeoutException)
            {
                Assert.Fail($"Không tìm thấy sản phẩm '{productData.TenSanPham}' trong danh sách sau khi thêm!");
            }

            //Thử xóa thương hiệu vừa tạo
            wait.Until(d => d.FindElement(By.XPath("//a[@href='/admin/thuonghieu/index']")).Displayed);
            driver.FindElement(By.XPath("//a[@href='/admin/thuonghieu/index']")).Click();

            var testData = GetThuongHieuTestDataFromExcel("QLTH2_05", "QLTH");
            bool thuongHieuFoundForDelete = false;
            IWebElement deleteButton = null;
            while (!thuongHieuFoundForDelete)
            {
                try
                {
                    IWebElement row = wait.Until(d => d.FindElement(By.XPath($"//td[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{testData.TenThuongHieu.ToLower()}')]/following-sibling::td/a[@class='btn btn-danger' and text()='Xóa']")));
                    deleteButton = row;
                    thuongHieuFoundForDelete = true;
                }
                catch (WebDriverTimeoutException)
                {
                    try
                    {
                        IWebElement nextPageButton = driver.FindElement(By.XPath("//li[@class='PagedList-skipToNext']/a[@rel='next']"));
                        if (nextPageButton != null && nextPageButton.Displayed)
                        {
                            Console.WriteLine("Không tìm thấy thương hiệu trên trang hiện tại, chuyển sang trang tiếp theo...");
                            nextPageButton.Click();
                            Thread.Sleep(1000);
                        }
                        else
                        {
                            Assert.Fail($"Không tìm thấy thương hiệu '{testData.TenThuongHieu}' trong danh sách để xóa!");
                            break;
                        }
                    }
                    catch (NoSuchElementException)
                    {
                        Assert.Fail($"Không tìm thấy thương hiệu '{testData.TenThuongHieu}' trong danh sách và không còn trang để chuyển!");
                        break;
                    }
                }
            }

            deleteButton.Click();

            // Xác nhận xóa (trang xác nhận)
            try
            {
                wait.Until(d => d.FindElement(By.CssSelector("button.btn.btn-danger")).Displayed);
                IWebElement confirmDeleteButton = driver.FindElement(By.CssSelector("button.btn.btn-danger"));
                confirmDeleteButton.Click();
            }
            catch (WebDriverTimeoutException)
            {
                Assert.Fail("Không tìm thấy nút xác nhận xóa trên trang xác nhận!");
            }

            //Kiểm tra xem thương hiệu có bị xóa hay không (nếu xóa được thì fail)
            try
            {
                wait.Until(d => !d.FindElements(By.XPath($"//td[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{testData.TenThuongHieu.ToLower()}')]")).Any());
                Assert.Fail($"Thương hiệu '{testData.TenThuongHieu}' đã bị xóa dù có sản phẩm liên kết!");
            }
            catch (WebDriverTimeoutException)
            {
                Assert.Pass($"Thương hiệu '{testData.TenThuongHieu}' không bị xóa khi có sản phẩm liên kết, đúng như kỳ vọng!");
            }
        }

        [Test]
        public void Test_Combined_CreateAndLoginNewAccount()
        {
            Login("tuan159", "Tuan@159159");
            wait.Until(d => d.FindElement(By.Id("userMenuButton")).Displayed);
            driver.FindElement(By.Id("userMenuButton")).Click();
            driver.FindElement(By.Id("manageLink")).Click();
            driver.FindElement(By.LinkText("Quản lý tài khoản")).Click();

            IWebElement createButton = wait.Until(d => d.FindElement(By.CssSelector("a.btn.btn-primary[href='/Admin/TaiKhoan/Create']")));
            createButton.Click();

            TaiKhoanTestData testData = GetTaiKhoanTestDataFromExcel("QLTKDN_01", "QLTK");
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
            bool found = false;
            while (!found)
            {
                var accountElements = driver.FindElements(By.XPath($"//td[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{testData.HoVaTen.ToLower()}')]"));
                if (accountElements.Count > 0 && accountElements[0].Displayed)
                {
                    IWebElement accountName = accountElements[0];
                    Assert.That(accountName.Displayed, Is.True, "Tài khoản mới không hiển thị trong danh sách!");
                    Console.WriteLine($"Tài khoản '{testData.HoVaTen}' đã được thêm thành công!");
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
                    }
                    else
                    {
                        Assert.Fail($"Không tìm thấy tài khoản '{testData.HoVaTen}' trong danh sách sau khi duyệt qua tất cả các trang!");
                        break;
                    }
                }
            }
            wait.Until(d => d.FindElement(By.XPath("//a[@href='/']")).Displayed);
                IWebElement backToHomeLink = driver.FindElement(By.XPath("//a[@href='/']"));
                backToHomeLink.Click();
                wait.Until(d => d.FindElement(By.Id("userMenuButton")).Displayed);
                driver.FindElement(By.Id("userMenuButton")).Click();
                wait.Until(d => d.FindElement(By.Id("logoutButton")).Displayed);
                driver.FindElement(By.Id("logoutButton")).Click();
                wait.Until(d => d.FindElement(By.Id("email")).Displayed);
                driver.FindElement(By.Id("email")).SendKeys(testData.UserName);
                driver.FindElement(By.Id("password")).SendKeys(testData.Matkhau);
                IWebElement loginButton = driver.FindElement(By.Id("dangnhap"));
                loginButton.Click();

                Console.WriteLine($"URL hiện tại sau khi đăng nhập: {driver.Url}");
                Assert.Fail($"Đăng nhập thất bại với tài khoản '{testData.UserName}'! Không chuyển đến trang chủ.");
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