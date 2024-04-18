using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Collections.Generic;
using Newtonsoft.Json;
using System.Net.Http;
using Newtonsoft.Json.Linq;

namespace WindowsFormsApp1
{
    public partial class CouponNotion : Form
    {
        private const string ApiUrl = "https://P11124-game-adapter.qookkagames.com/cms/active_code/change";

        private readonly List<string> playerNames = new List<string>();
        private readonly List<string> playerIds = new List<string>();
        private readonly List<string> ErrorPlayer = new List<string>();
        private readonly List<string> ErrorReason = new List<string>();

        public CouponNotion()
        {
            InitializeComponent();

        }

        private async void button1_Click(object sender, EventArgs e)
        {
            playerNames.Clear();
            playerIds.Clear();
            ErrorReason.Clear();

            // Selenium WebDriver 설정
            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--start-minimized"); // 원하는 옵션 추가 가능
                                                      //options.AddArguments("--headless"); // 브라우저 창을 숨기는 옵션

            string chromeDriverPath = @"C:\\Users\\onlys\\Downloads\\chromedriver-win64\\chromedriver.exe"; // ChromeDriver의 경로 설정
            IWebDriver driver = new ChromeDriver(chromeDriverPath, options);

            driver.Navigate().GoToUrl("https://mountainous-snowman-23c.notion.site/LIST-ce1fc7cfc2b040908f5cd3d9f7015a8b");

            // 웹 페이지가 완전히 로드될 때까지 기다립니다. (예: 10초)
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);

            // 여기서부터 실제 웹 페이지에서 필요한 요소를 찾고 상호작용하는 코드를 작성합니다.
            // 예를 들어, '더 불러오기' 버튼을 찾아 클릭하는 과정을 반복할 수 있습니다.

            // 데이터를 저장할 리스트
            List<Dictionary<string, string>> dataList = new List<Dictionary<string, string>>();

            // '더 불러오기' 텍스트를 가진 모든 <span> 태그를 찾습니다.
            var spans = driver.FindElements(By.XPath("//span[text()='더 불러오기']"));
            // CSS 선택자를 사용하여 요소를 찾습니다.
            var elements = driver.FindElements(By.CssSelector("div.notranslate"));
            foreach (var element in elements)
            {
                txt_coupon.Text = element.Text; // 요소의 텍스트 값을 출력합니다.
            }

            if (spans.Any())
            {
                // 첫 번째 발견된 <span> 태그의 부모 <div>를 찾아 클릭합니다.
                var parentDiv = spans.First().FindElement(By.XPath("./ancestor::div[1]"));
                IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                js.ExecuteScript("arguments[0].scrollIntoView(true);", parentDiv);
                //js.ExecuteScript("arguments[0].click();", element);
                parentDiv.Click();

                Console.WriteLine("부모 <div> 클릭 성공!");

                var dataList2 = new List<KeyValuePair<string, string>>();

                int rowIndex = 0;
                int rowCnt = 1000;

                while (rowIndex < rowCnt)
                {
                    int colIndex = 0;
                    // colIndex는 0과 1을 반복합니다.
                    for (colIndex = 0; colIndex <= 1; colIndex++)
                    {
                        try
                        {
                            var selector = $"div[data-row-index='{rowIndex}'][data-col-index='{colIndex}'] span";
                            var span = driver.FindElement(By.CssSelector(selector));
                            if (span.Text == "") { rowCnt = playerNames.Count; break; }
                            if (span != null)
                            {
                                if (colIndex == 0)
                                {
                                    playerNames.Add(span.Text.Trim());
                                }
                                else
                                {
                                    playerIds.Add(span.Text.Trim());
                                    rowIndex++; // rowIndex를 1씩 증가시킵니다.
                                }
                            }
                            else
                            {
                                // span이 없으면 다음 rowIndex로 이동
                                break;
                            }
                        }
                        catch (NoSuchElementException)
                        {
                            // 더 이상 해당 colIndex의 요소가 없으면 다음 rowIndex로 이동
                            rowCnt = playerNames.Count;
                            break;
                        }
                    }
                }
            }
            driver.Quit();

            using (var client = new HttpClient())
            {
                ErrorReason.Add(playerNames.Count + " 명 완료했습니다." );

                for (int i = 0; i < playerNames.Count; i++)
                {
                    var requestData = new
                    {
                        player_name = playerNames[i],
                        player_id = playerIds[i],
                        code = txt_coupon.ToString()
                    };

                    var json = Newtonsoft.Json.JsonConvert.SerializeObject(requestData);
                    var content = new StringContent(json, Encoding.UTF8, "application/json");

                    var response = await client.PostAsync(ApiUrl, content);
                    var content2 = await response.Content.ReadAsStringAsync();
                    JObject parsedData = JObject.Parse(content2);

                    string response_message = parsedData["message"].ToString();
                    int code = (int)parsedData["code"];

                    switch (code)
                    {
                        case 200:
                            response_message = "교환 성공!";
                            break;
                        case 419:
                            response_message = "해당 쿠폰코드는 최대 교환 인원수를 초과하였거나 존재하지 않는 쿠폰코드입니다.";
                            break;
                        case 10608:
                            response_message = "잘못된 캐릭터 ID 혹은 캐릭터명입니다. 다시 입력해 주세요.";
                            break;
                        case 10610:
                            response_message = "귀하는 이미 해당 쿠폰코드와 중복 사용 불가한 다른 쿠폰코드를 사용했습니다.";
                            break;
                        case 10612:
                            response_message = "귀하는 이미 해당 쿠폰코드를 교환하여 중복 교환이 불가합니다!";
                            break;
                    }
                    if (playerNames[i] != null)
                    {
                        ErrorReason.Add(playerNames[i].ToString() + " : " + response_message);
                    }
                }
                string message = string.Join(Environment.NewLine, ErrorReason);

                MessageBox.Show(message, "Name List", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }
    }
}
