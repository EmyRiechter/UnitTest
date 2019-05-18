using System;
using System.Diagnostics;
using System.Drawing.Imaging;
using System.IO;
using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Winium;

namespace TesteOpet

{
    [TestClass]

    public class UnitTest1
    {
        private WiniumDriver driver;
        private TestContext _testContextInstance;

        [TestInitialize]
        public void Init()
        {
            HtmlReporter = new ExtentHtmlReporter("Report.html");
            Extent = new ExtentReports();
            Extent.AttachReporter(HtmlReporter);

            Test = Extent.CreateTest($"{TestContext.TestName}", "");

        }

        [TestMethod]
        public void TestMethod1()
        {
            //abrir a calculadora
            DesktopOptions option = new DesktopOptions();
            option.ApplicationPath = "C:\\Windows\\system32\\calc.exe";

            driver = new WiniumDriver("Driver", option);

            //clicar nos botões indicados, os encontrando pelo nome
            driver.FindElement(By.Name("Sete")).Click();
            TakeScreenshot(" ");
            driver.FindElement(By.Name("Mais")).Click();
            TakeScreenshot(" ");
            driver.FindElement(By.Name("Dois")).Click();
            TakeScreenshot(" ");
            driver.FindElement(By.Name("Igual a")).Click();
             TakeScreenshot(" ");

        }

        public void TakeScreenshot(string message)
        {
            //tira um print da tela e salva com nome tempScreenshot em formato png
            Screenshot shot = ((ITakesScreenshot)driver).GetScreenshot();
            shot.SaveAsFile("tempScreenshot", ImageFormat.Png);

            //chama o método para converter a imagem
            String TestInfo = "<h5>" + message +
                "</h5><br/><a target='_blank' onClick=OpenImage('data:image/png;base64," +
                ConvertImageToBase64("tempScreenshot") +
                "')><img src='data:image/png;base64," +
                ConvertImageToBase64("tempScreenshot") +
                "' alt=''></a";

            //deleta a imagem em formato png
            File.Delete("tempScreenshot");

            //retorna a imagem no html
            Test.Info(TestInfo);

        }

        public static string ConvertImageToBase64(string fileName)
        {
            //converte a imagem para base64 e retorna em um array
            byte[] imageArray = File.ReadAllBytes(fileName);
            return Convert.ToBase64String(imageArray);
        }

        public TestContext TestContext
        {
            get
            {
                return _testContextInstance;
            }
            set
            {
                _testContextInstance = value;
            }
        }

        //executa sempre no final
        [TestCleanup]
        public void CleanUp()
        {
            foreach (Process p in Process.GetProcessesByName("Winium.Desktop.Driver"))
            {
                //se o driver ficar executando em segundo plano, o projeto não compila, então matamos o driver winium
                p.Kill();
            }

            Extent.Flush();
            HtmlReporter.Stop();

            //salva os resultados no arquivo html
            TestContext.AddResultFile("Report.html");
        }

        public ExtentHtmlReporter HtmlReporter { get; private set; }
        public ExtentReports Extent { get; private set; }
        public ExtentTest Test { get; private set; }

    }

}
