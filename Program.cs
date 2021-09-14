using CsvHelper;
using HtmlAgilityPack;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Web;

namespace airtable_templates_scraper
{
    public class CategoryModel
    {
        public string Name { get; set; }
        public string Url { get; set; }
    }

    public class DataModel
    {
        public string Title { get; set; }
        public string Category { get; set; }
        public string Url { get; set; }
        public string CoverImage { get; set; }
        public string Description { get; set; }
        public string PreviewLink { get; set; }
    }
    class Program
    {
        static void Main(string[] args)
        {
            using (var driver = new ChromeDriver())
            {
                driver.Navigate().GoToUrl("https://www.airtable.com/templates");
                IWait<IWebDriver> wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30.00));
                wait.Until(driver1 => ((IJavaScriptExecutor)driver).ExecuteScript("return document.readyState").Equals("complete"));
                
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(driver.PageSource);

                List<CategoryModel> categories = new List<CategoryModel>();
                var categoriesNode = doc.DocumentNode.SelectSingleNode("/html/body/div[1]/div/div/div/div[2]/div/div/div/div/div[1]/div/div[2]/div[2]/div[2]");
                if (categoriesNode != null)
                {
                    foreach (var category in categoriesNode.ChildNodes.Where(x => x.Name == "div"))
                    {
                        CategoryModel cat = new CategoryModel();
                        HtmlDocument sub = new HtmlDocument();
                        sub.LoadHtml(category.InnerHtml);

                        cat.Name = sub.DocumentNode.SelectSingleNode("/a[1]").InnerText;
                        cat.Url = sub.DocumentNode.SelectSingleNode("/a[1]").Attributes.FirstOrDefault(x => x.Name == "href").Value;
                        categories.Add(cat);
                    }
                }

                driver.Manage().Window.Maximize();
                List<DataModel> entries = new List<DataModel>();
                foreach (var category in categories)
                {
                    driver.Navigate().GoToUrl(category.Url);
                    wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30.00));
                    wait.Until(driver1 => ((IJavaScriptExecutor)driver).ExecuteScript("return document.readyState").Equals("complete"));
                    Thread.Sleep(5000);
                    doc = new HtmlDocument();
                    doc.LoadHtml(driver.PageSource);

                    var listNode = doc.DocumentNode.SelectSingleNode("/html/body/div[1]/div/div/div/div[2]/div/div/div/div/div[2]/div[2]/div/div");
                    if (listNode != null)
                    {
                        foreach (var node in listNode.ChildNodes.Where(x=>x.Name=="div"))
                        {
                            DataModel entry = new DataModel() { Category=category.Name};
                            HtmlDocument sub = new HtmlDocument();
                            sub.LoadHtml(node.InnerHtml);

                            var titleNode = sub.DocumentNode.SelectSingleNode("/div[1]/div[1]/a[1]/h4[1]");
                            if (titleNode != null)
                            {
                                entry. Title = HttpUtility.HtmlDecode(titleNode.InnerText);
                            }

                            var UrlNode = sub.DocumentNode.SelectSingleNode("/div[1]/div[1]/a[1]");
                            if (UrlNode != null)
                            {
                            entry.Url = HttpUtility.HtmlDecode(UrlNode.Attributes.FirstOrDefault(x=>x.Name=="href").Value);
                            }

                            var coverNode = sub.DocumentNode.SelectSingleNode("/div[1]/a[1]");
                            if (coverNode != null)
                            {
                                var style = coverNode.Attributes.FirstOrDefault(x => x.Name == "style");
                                if (style != null)
                                {
                                    entry.CoverImage = HttpUtility.HtmlDecode(style.Value.Split(new string[] { "(&quot;" }, StringSplitOptions.None)[1].Replace("&quot;);", ""));
                                }
                                else
                                    Console.WriteLine("Unable to get cover image");
                            }

                            entries.Add(entry);
                        }
                        using (var writer = new StreamWriter("file.csv"))
                        using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
                        {
                            csv.WriteRecords(entries);
                        }
                    }
                    else
                        Console.WriteLine("List node not found");
                }

                
            }
        }
    }
}
