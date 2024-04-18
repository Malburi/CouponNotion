using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Windows.Forms;
using ExcelDataReader;
using Newtonsoft.Json.Linq;
using System.Xml.Serialization;

namespace Coupon
{
    public partial class Coupon : Form
    {
        private const string ApiUrl = "https://P11124-game-adapter.qookkagames.com/cms/active_code/change";

        private readonly List<string> playerNames = new List<string>();
        private readonly List<string> playerIds = new List<string>();
        private readonly List<string> ErrorPlayer = new List<string>();
        private readonly List<string> ErrorReason = new List<string>();

        private string lastFilePath; // 마지막으로 선택한 파일의 경로

        public Coupon()
        {
            InitializeComponent();
            LoadLastFilePath(); // 프로그램 시작 시 마지막으로 선택한 파일의 경로를 불러옴
        }

        private async void button1_ClickAsync(object sender, EventArgs e)
        {
            try
            {
                playerNames.Clear();
                playerIds.Clear();
                ErrorPlayer.Clear();
                ErrorReason.Clear();

                string filePath = lastFilePath; // 마지막으로 선택한 파일의 경로 사용

                if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                {
                    // 저장된 경로가 없거나 파일이 존재하지 않으면 파일 선택 다이얼로그 표시
                    OpenFileDialog dialog = new OpenFileDialog();
                    dialog.Filter = "Excel files|*.xls;*.xlsx";
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        filePath = dialog.FileName;
                        SaveLastFilePath(filePath); // 마지막으로 선택한 파일의 경로 저장
                        lbl_path.Text = "경로 : " + filePath;
                    }
                    else
                    {
                        return; // 파일 선택이 취소된 경우 처리 중단
                    }
                }

                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        while (reader.Read())
                        {
                            if(reader.GetString(0) != null)
                            {
                                string playerName = reader.GetString(0).Trim();
                                string playerId = reader.GetString(1).Trim();

                                playerNames.Add(playerName);
                                playerIds.Add(playerId);
                            }
                        }
                    }
                }

                using (var client = new HttpClient())
                {
                    for (int i = 0; i < playerNames.Count; i++)
                    {
                        var requestData = new
                        {
                            player_name = playerNames[i],
                            player_id = playerIds[i],
                            code = txt_coupon.Text
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
            catch (Exception ex)
            {
                MessageBox.Show($"오류 발생: {ex.Message}");
            }
        }

        private void SaveLastFilePath(string filePath)
        {
            try
            {
                // 설정 파일에 마지막으로 선택한 파일의 경로 저장
                var settings = new ProgramSettings { LastFilePath = filePath };
                XmlSerializer serializer = new XmlSerializer(typeof(ProgramSettings));
                using (FileStream fileStream = new FileStream("settings.xml", FileMode.Create))
                {
                    serializer.Serialize(fileStream, settings);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"경로 저장 중 오류 발생: {ex.Message}");
            }
        }

        private void LoadLastFilePath()
        {
            try
            {
                if (File.Exists("settings.xml"))
                {
                    XmlSerializer serializer = new XmlSerializer(typeof(ProgramSettings));
                    using (FileStream fileStream = new FileStream("settings.xml", FileMode.Open))
                    {
                        ProgramSettings settings = (ProgramSettings)serializer.Deserialize(fileStream);
                        lastFilePath = settings.LastFilePath;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"경로 불러오기 중 오류 발생: {ex.Message}");
            }
        }

        private void Coupon_Load(object sender, EventArgs e)
        {
            if (lastFilePath != null)
            {
                lbl_path.Text = "경로 : " + lastFilePath;
            }
            else
            {
                lbl_path.Text = "경로 : 등록 안됨" ;
            }
            
        }
    }

    public class ProgramSettings
    {
        public string LastFilePath { get; set; }
    }
}
