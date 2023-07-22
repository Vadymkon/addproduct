using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace addproduct
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            //readXLS(@"D:\install\Downloads\r2.xlsx");
        }

        public void readXLS(string FilePath)
        {
            using (ExcelPackage package = new ExcelPackage(new FileInfo(FilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
                    {
                        Console.WriteLine(" Row:" + row + " column:" + col + " Value:" + worksheet.Cells[row, col].Value?.ToString().Trim());
                    }
                }
            }
        }
    void makeXLfile()
        {
            using (ExcelPackage excel = new ExcelPackage())
            {
                excel.Workbook.Worksheets.Add("MainList");

                var ws = excel.Workbook.Worksheets["MainList"];
                var headerRow = new List<string[]>()
                  {
                    new string[] { "Назва", "Опис", "Код товару", "Цена", "Категория" }
                  };
                // заголовок
                ws.Cells["A1:E1"].LoadFromArrays(headerRow);

                //заполнение
                string[] links = File.ReadAllLines(@"D:\vadymkon\Проги\TRY\addproduct\datalinks.txt");
                
                IWebDriver driver = new EdgeDriver();
                string path = "";
                for (int i = 0; i<links.Length; i++)
                {
                    path = links[i];
                    driver.Navigate().GoToUrl(path);

                    //описание
                    string name = driver.FindElement(By.CssSelector(".entry-title")).Text;
                    string description = driver.FindElement(By.CssSelector(".woocommerce-product-details__short-description")).Text;
                    string articul = driver.FindElement(By.CssSelector(".sku")).Text;
                    string price = driver.FindElement(By.CssSelector(".amount")).Text.OnlyNumbers();
                    // label1.Text = $"{name}\n{description}\n{articul}\n{price}";

                    //картинка (пока что на рабочий стол)
                    string getlinkofimage = driver.FindElement(By.CssSelector(".wp-post-image")).GetAttribute("data-src");
                    using (WebClient wc = new WebClient())
                        wc.DownloadFileAsync(new Uri(getlinkofimage), $"{Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)}/photos/{articul}.jpg");

                    //+2 потому что 0й и 1й елемент заняты
                    ws.Cells[$"A{i + 2}:E{i + 2}"].LoadFromArrays(new List<string[]> { new string[] { name, description, articul, price, "" } });
                }

                // ws.Cells["D1"].Value = "Inactive Agencies";
                //save
                ws.Protection.IsProtected = false;
                ws.Protection.AllowSelectLockedCells = false;
                FileInfo excelFile = new FileInfo($"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}/test.xlsx");
                excel.SaveAs(excelFile);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            makeXLfile();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            //http://www.blagodar.kiev.ua/admin/index.php?route=catalog/product/add&token=EC3TGRllW4p8LsmjX0kJB8O8hfYRgMuM
            string path = "http://www.blagodar.kiev.ua/admin/index.php?route=catalog/product&token=XbOG1jLzih5gTlvGPnajZmdN9HGQtm3y";


            // create a new instance of the Edge driver
            IWebDriver driver = new EdgeDriver();
            driver.Navigate().GoToUrl(path);

            // вход в аккаунт
            driver.FindElement(By.Id("input-username")).SendKeys("admin");
            driver.FindElement(By.Id("input-password")).SendKeys("20658184");
            driver.FindElement(By.ClassName("text-right")).FindElement(By.TagName("button")).Click();

            //Тут обьявляется товар
            string nameofJPGdirectory = "demo";
            //data
            bool IsThere = false;
            string name = "ТОВАР ТОВАРСКИЙ";
            string description = "Привет\n\rКак дела?\n\rДавно не созванивались\n\rНапиши!";
            string articul = "apple_logo";
            string categorya = "Тваринт";
            string price = "499";

            //проверка
            driver.FindElement(By.Id("input-model")).Clear();
            driver.FindElement(By.Id("input-model")).SendKeys(articul);
            driver.FindElement(By.Id("button-filter")).Click();
            var we = driver.FindElement(By.CssSelector(".table-responsive")).FindElements(By.CssSelector(".text-left"));
            foreach (IWebElement elem in we) if(elem.Text==articul) IsThere = true;
            
            //если нету то добавить
            if (!IsThere)
            {
                //Запись товара
                driver.FindElement(By.CssSelector("a[data-original-title=\"Додати\"]")).Click();
                //Запись имени товара и описания
                driver.FindElement(By.Name("product_description[3][name]")).SendKeys(name);

                driver.FindElement(By.CssSelector(".note-icon-code")).Click();
                driver.FindElement(By.CssSelector(".note-codable")).Clear();
                driver.FindElement(By.CssSelector(".note-codable")).SendKeys(description.Oformlator());
                //driver.FindElement(By.CssSelector(".note-editable.panel-body")).FindElement(By.TagName("p")).SendKeys("<b>Описание</b>\nАзаза");

                //Переключение вкладки и запись иного свойства
                driver.FindElement(By.CssSelector("a[href=\"#tab-data\"]")).Click();
                driver.FindElement(By.Id("input-model")).SendKeys(articul);
                driver.FindElement(By.Id("input-price")).SendKeys(price);

                //Категории
                driver.FindElement(By.CssSelector("a[href=\"#tab-links\"]")).Click();
                new SelectElement(driver.FindElement(By.Id("input-manufacturer"))).SelectByValue("11");
                driver.FindElement(By.CssSelector(".table-striped")).FindElements(By.TagName("label")).First(x => x.Text.Contains("Золота")).Click();
                driver.FindElement(By.CssSelector(".table-striped")).FindElements(By.TagName("label")).First(x => x.Text.Contains(categorya.Remove(categorya.Length - 1))).Click();



                //картинка
                driver.FindElement(By.CssSelector("a[href=\"#tab-image\"]")).Click();
                driver.FindElement(By.Id("thumb-image")).Click();
                driver.FindElement(By.Id("button-image")).Click();
                //тут открывается динамическое окно выбора, потому нужно время подождать
                IWebElement element = new WebDriverWait(driver, TimeSpan.FromSeconds(3)).Until(x => x.FindElement(By.Id("button-upload")));
                //Так ищется и открывается папка
                driver.FindElement(By.Id("modal-image")).FindElements(By.TagName("a")).First(an => an.GetAttribute("href") != null && an.GetAttribute("href").Contains($"directory={nameofJPGdirectory}")).Click();
                //Поиск картинки
                new WebDriverWait(driver, TimeSpan.FromSeconds(3)).Until(x => x.FindElement(By.Name("search")));
                new WebDriverWait(driver, TimeSpan.FromSeconds(3)).Until(x => x.FindElement(By.Name("search")));
                driver.FindElement(By.Name("search")).Click(); driver.FindElement(By.Name("search")).SendKeys(articul);
                driver.FindElement(By.Id("button-search")).Click();
                //Выбор картинки
                new WebDriverWait(driver, TimeSpan.FromSeconds(3)).Until(x => x.FindElement(By.CssSelector($"img[title=\"{articul}.jpg\"]")));
                new WebDriverWait(driver, TimeSpan.FromSeconds(3)).Until(x => x.FindElement(By.CssSelector($"img[title=\"{articul}.jpg\"]")));
                driver.FindElement(By.CssSelector($"img[title=\"{articul}.jpg\"]")).Click();

                //Сохранить товар (кнопка)
               // driver.FindElement(By.CssSelector(".btn-primary")).Click();
            }

            /*
             * Двойные wait чтобы не сбивалось. чёто сбивается. Но с двойными вейтами уже вроде нет
            //Закрыть брузер
            driver.Close();
            */
        }

        List <string> getinfo (string path = "https://www.m-opt.com/shop-ua/kartina-zpk-024-ua/")
        { //закидываешь ссылку, получаешь все нужные данные о товаре и скачиваешь картинку
            IWebDriver driver = new EdgeDriver();
            driver.Navigate().GoToUrl(path);

            //описание
            string name = driver.FindElement(By.CssSelector(".entry-title")).Text;
            string description = driver.FindElement(By.CssSelector(".woocommerce-product-details__short-description")).Text;
            string articul = driver.FindElement(By.CssSelector(".sku")).Text;
            string price = driver.FindElement(By.CssSelector(".amount")).Text.OnlyNumbers();
           // label1.Text = $"{name}\n{description}\n{articul}\n{price}";

            //картинка (пока что на рабочий стол)
            string getlinkofimage = driver.FindElement(By.CssSelector(".wp-post-image")).GetAttribute("data-src");
            using (WebClient wc = new WebClient())
                wc.DownloadFileAsync(new Uri(getlinkofimage), $"{Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)}/{articul}.jpg");
            

            return new List<string>{ name, description, articul, price, "" };
        }

        private void button2_Click(object sender, EventArgs e)
        {
            getinfo();
        }

        void button3_Click(object sender, EventArgs e)
        {
            List<string> links = new List<string>();
            IWebDriver driver = new EdgeDriver();
            string path = "";
            int i = 1;
            while (true)
            {
                path = $"https://www.m-opt.com/shop-ua/page-ua/{i}/?filter_brand=zolota-pidkova-ua&query_type_brand=or";
                ++i; //переходим по страницам
                driver.Navigate().GoToUrl(path);
                if (driver.FindElements(By.Id("Error_404")).Count != 0) { driver.Dispose(); break; }
                var elems = driver.FindElements(By.CssSelector(".product-type-simple")); //находит все список-пункты
                foreach (IWebElement am in elems)
                    if (!links.Contains(am.FindElement(By.TagName("a")).GetAttribute("href"))) //если там уже нету этой ссылки
                        links.Add(am.FindElement(By.TagName("a")).GetAttribute("href")); //берёт ссылки на товар
            }
            //в файлик всё
            using (StreamWriter writer = new StreamWriter($"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}/datalinks.txt", false))
                links.ForEach(x => writer.WriteLine(x));
        }

        private void button5_Click(object sender, EventArgs e)
        {
                int col = 2;
                List<string> category1 = new List<string>();
                List<string> code1 = new List<string>();
                List<string> code2 = new List<string>();
                List<string> category2 = new List<string>();
            //читаем категории
            using (ExcelPackage package = new ExcelPackage(new FileInfo(@"D:\vadymkon\Проги\TRY\addproduct\GoldPrice.xlsx")))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                    for (int row = 4; row <= worksheet.Dimension.End.Row; row++)
                    {
                        category1.Add(worksheet.Cells[row, col].Value?.ToString());
                        if (category1[category1.Count - 1] == null) { category1.RemoveAt(category1.Count - 1); break; }
                    Console.WriteLine(" Row:" + row + " column:" + col + " Value:" + worksheet.Cells[row, col].Value?.ToString().Trim());
                    }
                col = 3;
                for (int row = 4; row <= worksheet.Dimension.End.Row; row++)
                {
                    code1.Add(worksheet.Cells[row, col].Value?.ToString().Replace("г", ""));
                    if (code1[code1.Count - 1] == null) { code1.RemoveAt(code1.Count - 1); break;}
                    Console.WriteLine(" Row:" + row + " column:" + col + " Value:" + worksheet.Cells[row, col].Value?.ToString().Trim());
                }
            }
            //читаем оригинальный порядок кодов
            using (ExcelPackage package = new ExcelPackage(new FileInfo(@"C:\Users\Вадюшка\Desktop\test.xlsx")))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                col = 3;
                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                    {
                        code2.Add(worksheet.Cells[row, col].Value?.ToString());
                    if (code2[code2.Count - 1] == null) { code2.RemoveAt(code2.Count - 1); break; }
                    category2.Add("");
                        Console.WriteLine(" Row:" + row + " column:" + col + " Value:" + worksheet.Cells[row, col].Value?.ToString().Trim());
                    }
            }

            //пересортировка
            Console.WriteLine($"{category1.Count} {code1.Count} {category2.Count} {code2.Count}");
            for (int i =0; i<code1.Count;i++)
            {
                int index = code2.IndexOf(code1[i]);
                category2[index] = category1[i];
            }
            //запись
            using (ExcelPackage excel = new ExcelPackage(new FileInfo(@"C:\Users\Вадюшка\Desktop\test.xlsx")))
            {
                var ws = excel.Workbook.Worksheets[1];
                for (int i = 0; i < category2.Count; i++)
                    ws.Cells[$"E{i + 2}"].Value = category2[i];
            FileInfo excelFile = new FileInfo($"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}/test.xlsx");
            excel.SaveAs(excelFile);
            }
            MessageBox.Show("Done!");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            int col = 1;
            List<string> names = new List<string>();
            List<string> descs = new List<string>();
            List<string> codes = new List<string>();
            List<string> prices = new List<string>();
            List<string> categorys = new List<string>();
            //читаем коды
            using (ExcelPackage package = new ExcelPackage(new FileInfo(@"D:\vadymkon\Проги\TRY\addproduct\test.xlsx")))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                col = 1;
                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    names.Add(worksheet.Cells[row, col].Value?.ToString());
                    if (names[names.Count - 1] == null) { names.RemoveAt(names.Count - 1); break; }
                    Console.WriteLine("NAME:  Row:" + row + " column:" + col + " Value:" + worksheet.Cells[row, col].Value?.ToString().Trim());
                }
                col = 2;
                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    descs.Add(worksheet.Cells[row, col].Value?.ToString());
                    if (descs[descs.Count - 1] == null) { descs.RemoveAt(descs.Count - 1); break; }
                    Console.WriteLine("DESCS:  Row:" + row + " column:" + col + " Value:" + worksheet.Cells[row, col].Value?.ToString().Trim());
                }
                col = 3;
                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    codes.Add(worksheet.Cells[row, col].Value?.ToString());
                    if (codes[codes.Count - 1] == null) { codes.RemoveAt(codes.Count - 1); break; }
                    Console.WriteLine("CODE:  Row:" + row + " column:" + col + " Value:" + worksheet.Cells[row, col].Value?.ToString().Trim());
                }
                col = 4;
                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    prices.Add(worksheet.Cells[row, col].Value?.ToString());
                    if (prices[prices.Count - 1] == null) { prices.RemoveAt(prices.Count - 1); break; }
                    Console.WriteLine("PRICES:  Row:" + row + " column:" + col + " Value:" + worksheet.Cells[row, col].Value?.ToString().Trim());
                }
                col = 5;
                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    categorys.Add(worksheet.Cells[row, col].Value?.ToString());
                    if (categorys[categorys.Count - 1] == null) { categorys.RemoveAt(categorys.Count - 1); break; }
                    Console.WriteLine("CATEGORYS:  Row:" + row + " column:" + col + " Value:" + worksheet.Cells[row, col].Value?.ToString().Trim());
                }
            }
            //Ручное изменение категорий
            for (int i = 0; i<categorys.Count;i++)
                if(categorys[i].Contains("Новорічний мотив"))
                    categorys[i] = "Новорічн";

            Console.WriteLine($"There: {names.Count} {descs.Count} {codes.Count} {prices.Count} {categorys.Count}");
            
            //http://www.blagodar.kiev.ua/admin/index.php?route=catalog/product/add&token=EC3TGRllW4p8LsmjX0kJB8O8hfYRgMuM
            string path = "http://www.blagodar.kiev.ua/admin/index.php?route=catalog/product&token=XbOG1jLzih5gTlvGPnajZmdN9HGQtm3y";


            // create a new instance of the Edge driver
            IWebDriver driver = new EdgeDriver();
            driver.Navigate().GoToUrl(path);

            // вход в аккаунт
            driver.FindElement(By.Id("input-username")).SendKeys("admin");
            driver.FindElement(By.Id("input-password")).SendKeys("20658184");
            driver.FindElement(By.ClassName("text-right")).FindElement(By.TagName("button")).Click();


            string nameofJPGdirectory = "photoszp";
            //data
            bool   IsThere = false;
            string name = "ТОВАР ТОВАРСКИЙ";
            string description = "Привет\n\rКак дела?\n\rДавно не созванивались\n\rНапиши!";
            string articul = "ЗПТ-022";
            string categorya = "Тваринт";
            string price = "499";

            //kk то с какого товара начнётся заполнение (номер)
            for (int kk = 0; kk < names.Count; kk++)
            {
                //Тут обьявляется товар
                IsThere = false;
                name = names[kk];
                description = descs[kk];
                articul = codes[kk];
                categorya = categorys[kk];
                price = prices[kk];

                //проверка
                driver.FindElement(By.Id("input-model")).Clear();
                driver.FindElement(By.Id("input-model")).SendKeys(articul);
                driver.FindElement(By.Id("button-filter")).Click();
                var we = driver.FindElement(By.CssSelector(".table-responsive")).FindElements(By.CssSelector(".text-left"));
                foreach (IWebElement elem in we) if (elem.Text == articul) IsThere = true;

                //если нету то добавить
                if (!IsThere)
                {
                    //Запись товара
                    driver.FindElement(By.CssSelector("a[data-original-title=\"Додати\"]")).Click();
                    //Запись имени товара и описания
                    driver.FindElement(By.Name("product_description[3][name]")).SendKeys(name);

                    driver.FindElement(By.CssSelector(".note-icon-code")).Click();
                    driver.FindElement(By.CssSelector(".note-codable")).Clear();
                    driver.FindElement(By.CssSelector(".note-codable")).SendKeys(description.Oformlator());
                    //driver.FindElement(By.CssSelector(".note-editable.panel-body")).FindElement(By.TagName("p")).SendKeys("<b>Описание</b>\nАзаза");

                    //Переключение вкладки и запись иного свойства
                    driver.FindElement(By.CssSelector("a[href=\"#tab-data\"]")).Click();
                    driver.FindElement(By.Id("input-model")).SendKeys(articul);
                    driver.FindElement(By.Id("input-price")).SendKeys(price);

                    //Категории
                    driver.FindElement(By.CssSelector("a[href=\"#tab-links\"]")).Click();
                    new SelectElement(driver.FindElement(By.Id("input-manufacturer"))).SelectByValue("11");
                    driver.FindElement(By.CssSelector(".table-striped")).FindElements(By.TagName("label")).First(x => x.Text.Contains("Золота")).Click();
                    driver.FindElement(By.CssSelector(".table-striped")).FindElements(By.TagName("label")).First(x => x.Text.Contains(categorya.Remove(categorya.Length - 1))).Click();



                    //картинка
                    driver.FindElement(By.CssSelector("a[href=\"#tab-image\"]")).Click();
                    driver.FindElement(By.Id("thumb-image")).Click();
                    driver.FindElement(By.Id("button-image")).Click();
                    //тут открывается динамическое окно выбора, потому нужно время подождать
                    IWebElement element = new WebDriverWait(driver, TimeSpan.FromSeconds(3)).Until(x => x.FindElement(By.Id("button-upload")));
                    //Так ищется и открывается папка
                    driver.FindElement(By.Id("modal-image")).FindElements(By.TagName("a")).First(an => an.GetAttribute("href") != null && an.GetAttribute("href").Contains($"directory={nameofJPGdirectory}")).Click();
                    //Поиск картинки
                    new WebDriverWait(driver, TimeSpan.FromSeconds(3)).Until(x => x.FindElement(By.Name("search")));
                    new WebDriverWait(driver, TimeSpan.FromSeconds(3)).Until(x => x.FindElement(By.Name("search")));
                    driver.FindElement(By.Name("search")).Click(); driver.FindElement(By.Name("search")).SendKeys(articul);
                    driver.FindElement(By.Id("button-search")).Click();
                    //Выбор картинки
                    new WebDriverWait(driver, TimeSpan.FromSeconds(3)).Until(x => x.FindElement(By.CssSelector($"img[title=\"{articul.ToEng()}.jpg\"]")));
                    new WebDriverWait(driver, TimeSpan.FromSeconds(3)).Until(x => x.FindElement(By.CssSelector($"img[title=\"{articul.ToEng()}.jpg\"]")));
                    driver.FindElement(By.CssSelector($"img[title=\"{articul.ToEng()}.jpg\"]")).Click();

                    //Сохранить товар (кнопка)
                    driver.FindElement(By.CssSelector(".btn-primary")).Click();
                }
            }

            /*
             * Двойные wait чтобы не сбивалось. чёто сбивается. Но с двойными вейтами уже вроде нет
            //Закрыть брузер
            driver.Close();
            */
        }

        private void label1_Click(object sender, EventArgs e)
        {
            label1.Text = label1.Text.ToEng();
            Console.WriteLine(label1.Text.ToEng());
        }

        private void button7_Click(object sender, EventArgs e)
        {
            foreach (string pathik in Directory.GetFiles(@"D:\vadymkon\Проги\TRY\addproduct\photoszp"))
                File.Move(pathik, $@"D:\vadymkon\Проги\TRY\addproduct\photoszp1\{ pathik.Replace(@"D:\vadymkon\Проги\TRY\addproduct\photoszp\", "").ToEng()}");
        }

        private void button8_Click(object sender, EventArgs e)
        {
            int col = 1;
            List<string> names = new List<string>();
            List<string> descs = new List<string>();
            List<string> codes = new List<string>();
            List<string> prices = new List<string>();
            List<string> categorys = new List<string>();
            //читаем коды
            using (ExcelPackage package = new ExcelPackage(new FileInfo(@"D:\vadymkon\Проги\TRY\addproduct\test.xlsx")))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                col = 1;
                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    names.Add(worksheet.Cells[row, col].Value?.ToString());
                    if (names[names.Count - 1] == null) { names.RemoveAt(names.Count - 1); break; }
                    Console.WriteLine("NAME:  Row:" + row + " column:" + col + " Value:" + worksheet.Cells[row, col].Value?.ToString().Trim());
                }
                col = 2;
                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    descs.Add(worksheet.Cells[row, col].Value?.ToString());
                    if (descs[descs.Count - 1] == null) { descs.RemoveAt(descs.Count - 1); break; }
                    Console.WriteLine("DESCS:  Row:" + row + " column:" + col + " Value:" + worksheet.Cells[row, col].Value?.ToString().Trim());
                }
                col = 3;
                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    codes.Add(worksheet.Cells[row, col].Value?.ToString());
                    if (codes[codes.Count - 1] == null) { codes.RemoveAt(codes.Count - 1); break; }
                    Console.WriteLine("CODE:  Row:" + row + " column:" + col + " Value:" + worksheet.Cells[row, col].Value?.ToString().Trim());
                }
                col = 4;
                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    prices.Add(worksheet.Cells[row, col].Value?.ToString());
                    if (prices[prices.Count - 1] == null) { prices.RemoveAt(prices.Count - 1); break; }
                    Console.WriteLine("PRICES:  Row:" + row + " column:" + col + " Value:" + worksheet.Cells[row, col].Value?.ToString().Trim());
                }
                col = 5;
                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    categorys.Add(worksheet.Cells[row, col].Value?.ToString());
                    if (categorys[categorys.Count - 1] == null) { categorys.RemoveAt(categorys.Count - 1); break; }
                    Console.WriteLine("CATEGORYS:  Row:" + row + " column:" + col + " Value:" + worksheet.Cells[row, col].Value?.ToString().Trim());
                }
            }
            //Ручное изменение категорий
            for (int i = 0; i < categorys.Count; i++)
                if (categorys[i].Contains("Новорічний мотив"))
                    categorys[i] = "Новорічн";

            Console.WriteLine($"There: {names.Count} {descs.Count} {codes.Count} {prices.Count} {categorys.Count}");
            //categorys.Distinct().ToList().ForEach(Console.WriteLine);
            Dictionary<string, string> categors = new Dictionary<string, string>
            {
                ["Новорічн"] = "59,67",
                ["Лірика"] = "59,65",
                ["Природа"] = "59,68",
                ["Тварини"] = "59,78",
                ["Квіти"] = "59,64",
                ["Релігія"] = "59,72",
                ["Рушники пасхальні"] = "59,74",
                ["Натюрморти"] = "59,66",
                ["Ангелик"] = "59,61",
                ["Лента на спасовский рушник"] = "59,77",
            };

            using (ExcelPackage excel = new ExcelPackage(new FileInfo(@"D:\install\Downloads\template.xlsx")))
            {
                ExcelWorksheet ws1 = excel.Workbook.Worksheets[1];
                ExcelWorksheet ws2 = excel.Workbook.Worksheets[2];

                for (int i = 0; i < codes.Count; i++)
                {
                    ws1.Cells[$"B{i+2}"].Value = names[i];
                    ws1.Cells[$"C{i+2}"].Value = categors[categorys[i]];
                    ws1.Cells[$"L{i+2}"].Value = codes[i];
                    //vyshivka.xn--80aabjf3bxas.xn--j1amh/image/catalog/zpa-004-D.jpg
                    ws1.Cells[$"N{i+2}"].Value = $"catalog/{codes[i].ToEng()}.jpg";
                    if (File.Exists($@"D:\vadymkon\Проги\TRY\addproduct\photos2\{codes[i].ToEng()}-D.jpg")) 
                        ws2.Cells[$"B{i+2}"].Value = $"catalog/{codes[i].ToEng()}-D.jpg";
                    ws1.Cells[$"P{i+2}"].Value = prices[i];
                    ws1.Cells[$"AD{i+2}"].Value = descs[i].Oformlator();
                }
                FileInfo excelFile = new FileInfo($"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}/databasegp.xlsx");
                excel.SaveAs(excelFile);
            }
        }

        async void button9_Click(object sender, EventArgs e)
        {
            //заполнение
            string[] links = File.ReadAllLines(@"D:\vadymkon\Проги\TRY\addproduct\datalinks.txt");

                List<string> linksofjpg = new List<string>();
            IWebDriver driver = new EdgeDriver();
            string path = "";
            for (int i = 0; i < links.Length; i++)
            {
                path = links[i];
                driver.Navigate().GoToUrl(path);
                await Task.Delay(1000);
                string articul = driver.FindElement(By.CssSelector(".sku")).Text;
                if (driver.FindElement(By.CssSelector(".mfn-mim-2")).FindElements(By.TagName("li")).Count == 2)
                {
                    driver.FindElement(By.CssSelector(".mfn-mim-2")).FindElements(By.TagName("li")).ToList()[1].Click();
                    string getlinkofimage = driver.FindElement(By.CssSelector(".flex-active-slide")).FindElement(By.TagName("a")).GetAttribute("href");
                    using (WebClient wc = new WebClient())
                        wc.DownloadFileAsync(new Uri(getlinkofimage), $"{Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)}/photos2/{articul.ToEng()}-D.jpg");
                }
            }
            
        }
    }
    }
    public static class Ext
        {
        public static string OnlyNumbers(this string line)
        {
            List<char> numbers = new List<char>{ '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' };
            string finalstring = "";
            foreach (char a in line) if (numbers.Contains(a)) finalstring += a;
            return finalstring;
        }
        public static string Oformlator(this string line)
        {
            string final = "";
            string[] lines= line.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            for(int i =0; i<lines.Length; i++)
            {
                if (i == 0) final += "<p><span style = \"box-sizing: inherit; -webkit-font-smoothing: antialiased; margin: 0px; padding: 0px; border: 0px; font-variant-numeric: inherit; font-variant-east-asian: inherit; font-weight: 700; font-stretch: inherit; font-size: 17px; line-height: inherit; font-family: &quot;Alegreya Sans&quot;, -apple-system, BlinkMacSystemFont, &quot;Segoe UI&quot;, Roboto, Oxygen-Sans, Ubuntu, Cantarell, &quot;Helvetica Neue&quot;, sans-serif; vertical-align: baseline; color: rgb(98, 98, 98);\">";
                final += lines[i];
                if (i != lines.Length - 1) final += "</span><br style=\"box-sizing: inherit; -webkit-font-smoothing: antialiased; color: rgb(98, 98, 98); font-family: &quot;Alegreya Sans&quot;, -apple-system, BlinkMacSystemFont, &quot;Segoe UI&quot;, Roboto, Oxygen-Sans, Ubuntu, Cantarell, &quot;Helvetica Neue&quot;, sans-serif; font-size: 17px;\"><span style=\"color: rgb(98, 98, 98); font-family: &quot;Alegreya Sans&quot;, -apple-system, BlinkMacSystemFont, &quot;Segoe UI&quot;, Roboto, Oxygen-Sans, Ubuntu, Cantarell, &quot;Helvetica Neue&quot;, sans-serif; font-size: 17px;\">";
                else final += "</span><br></p>";
            }
            return final;
        }
        public static string DeleteFirstLine(this string line)
        {
            string final = "";
            string[] lines= line.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            for(int i =1; i<lines.Length; i++)
            {
                final += lines[i] +"\r\n";
            }
            return final;
        }
        public static Dictionary<char,string> chars = new Dictionary<char, string>
        {
            // АБВГДЕЁЖ ЗИЙКЛ  МНОПР СТУФХЦЧШЩ ЪЫЬЭЮЯ
            // abvgdeezh zijkl mnopr stufhcchs hsch yejuja
            ['А']="a",
            ['Б']="b",
            ['В']="v",
            ['Г']="g",
            ['Д']="d",
            ['Е']="e",
            ['Ё']="e",
            ['Ж']="zh",
            ['З']="z",
            ['И']="i",
            ['Й']="j",
            ['К']="k",
            ['Л']="l",
            ['М']="m",
            ['Н']="n",
            ['О']="o",
            ['П']="p",
            ['Р']="r",
            ['С']="s",
            ['Т']="t",
            ['У']="u",
            ['Ф']="f",
            ['Х']="h",
            ['Ц']="c",
            ['Ч']="ch",
            ['Ш']="sh",
            ['Щ']="sch",
            ['Ъ'] ="",
            ['Ы'] = "y",
            ['Ь'] ="",
            ['Э'] ="e",
            ['Ю'] ="ju",
            ['Я'] = "ja",
        };
        public static string ToEng(this string line)
        {
            string linecopy = "";
            for (int i =0; i < line.Length;i++)
            {
                if (chars.ContainsKey(line[i])) linecopy += (chars[line[i]]);
                else linecopy += line[i];
            }
            return linecopy; 
        }
    }

