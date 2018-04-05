using System;
using System.Threading;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System.Data;
using System.Text;
using GemBox.Spreadsheet;
using System.Linq;
using System.Collections.Generic;

namespace valvoline
{
	[TestFixture]
	public class Loader1 : TestBase
	{
		[SetUp]
		public void Start()
		{
			driver = new ChromeDriver();
			wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
		}

		[Test]
		public void LoadInfo()
		{

			driver.Url = "https://www.valvolineeurope.com/russian/products";
			driver.Manage().Window.Maximize();

			var products = new List <Product>();

			var level1ProdCount = driver.FindElements(By.CssSelector("#products li")).Count;
			for (var i = 0; i < level1ProdCount; i++)
			{
				var level1Cat = driver.FindElements(By.CssSelector("#products li"))[i].Text;
				driver.FindElements(By.CssSelector("#products li"))[i].Click();

				var level2ProdCount = driver.FindElements(By.CssSelector("#products li:not(.header) > a")).Count;
				for (var j = 0; j < level2ProdCount; j++)
				{
					var level2Cat = driver.FindElements(By.CssSelector("#products li:not(.header) > a"))[j].Text;
					driver.FindElements(By.CssSelector("#products li:not(.header) > a"))[j].Click();

					var level3ProdCount = driver.FindElements(By.CssSelector("#products li:not(.header) > a")).Count;
					for (var k = 0; k < level3ProdCount; k++)
					{
						var name = driver.FindElements(By.CssSelector("#products li:not(.header) > a"))[k].Text;
						var link = driver.FindElements(By.CssSelector("#products li:not(.header) > a"))[k].GetAttribute("href");

						var product = new Product
						{
							Name = name,
							Link = link,
							Category1 = level1Cat,
							Category2 = level2Cat
						};
						products.Add(product);
					}
					driver.FindElement(By.CssSelector("#products > div.subNav > a")).Click();
				}
				driver.FindElement(By.CssSelector("#products > div.subNav > a")).Click();
			}

			SpreadsheetInfo.SetLicense("EIKU-U5LX-6MSF-Z84S");
			ExcelFile ef = new ExcelFile();
			ExcelWorksheet ws = ef.Worksheets.Add("все продукты");

			DataTable dt = new DataTable();

			// add columns
			var c1 = dt.Columns.Add("Продукт", typeof(string));
			var c2 = dt.Columns.Add("Категория1", typeof(string));
			var c3 = dt.Columns.Add("Категория2", typeof(string));
			var c4 = dt.Columns.Add("Ссылка", typeof(string));

			foreach (var p in products)
			{
				dt.Rows.Add(
					p.Name,
					p.Category1,
					p.Category2,
					p.Link);
			}

			// add cell
			// ws.Cells[0, 0].Value = "DataTable insert example:";

			// Insert DataTable into an Excel worksheet.
			ws.InsertDataTable(dt,
				new InsertDataTableOptions()
				{
					ColumnHeaders = true,
					StartRow = 0
				});
			// Autofit columns and some print options (for better look when exporting to pdf, xps and printing).
			var columnCount = ws.CalculateMaxUsedColumns();
			for (int i = 0; i < columnCount; i++)
				ws.Columns[i].AutoFit();

			var date = DateTime.Now.ToString("yyyy.MM.dd_HH-mm-ss");
			var fileName = String.Concat(@"D:/valvoline_", date, ".xlsx");
			ef.Save(fileName);
		}
	}
}