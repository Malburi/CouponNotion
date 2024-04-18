using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Coupon
{
    static class Program
    {
        private static readonly DateTime startDate = new DateTime(2024, 4, 1); // 프로그램 실행 시작일
        private static readonly DateTime endDate = new DateTime(2024, 12, 30); // 프로그램 실행 종료일

        [STAThread]
        static void Main()
        {
            if (IsWithinDateRange())
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new Coupon());
            }
            else
            {
                MessageBox.Show("프로그램 실행 기간이 아닙니다. 제작자에게 문의해주세요.", "경고", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private static bool IsWithinDateRange()
        {
            // 네트워크 시간 동기화를 통해 실제 시간을 가져옴
            DateTime currentTime = GetNetworkTime();

            // 설정된 시작일과 종료일 사이에 현재 시간이 있는지 확인
            return currentTime >= startDate && currentTime <= endDate;
        }

        private static DateTime GetNetworkTime()
        {
            const string ntpServer = "pool.ntp.org"; // NTP 서버 주소

            try
            {
                // NTP 프로토콜을 사용하여 네트워크 시간을 가져옴
                var ntpData = new byte[48];
                ntpData[0] = 0x1B;
                var addresses = Dns.GetHostEntry(ntpServer).AddressList;
                var ipEndPoint = new IPEndPoint(addresses[0], 123);
                using (var socket = new Socket(AddressFamily.InterNetwork, SocketType.Dgram, ProtocolType.Udp))
                {
                    socket.Connect(ipEndPoint);
                    socket.Send(ntpData);
                    socket.Receive(ntpData);
                }

                const byte serverReplyTime = 40;
                ulong intPart = BitConverter.ToUInt32(ntpData, serverReplyTime);
                ulong fractPart = BitConverter.ToUInt32(ntpData, serverReplyTime + 4);
                intPart = SwapEndianness(intPart);
                fractPart = SwapEndianness(fractPart);
                var milliseconds = (intPart * 1000) + ((fractPart * 1000) / 0x100000000L);
                var networkDateTime = (new DateTime(1900, 1, 1, 0, 0, 0, DateTimeKind.Utc)).AddMilliseconds((long)milliseconds);
                return networkDateTime.ToLocalTime();
            }
            catch
            {
                // 네트워크 시간을 가져오는 데 실패하면 로컬 시간을 반환
                return DateTime.Now;
            }
        }

        private static ulong SwapEndianness(ulong x)
        {
            return (x >> 24) |
                ((x >> 8) & 0x0000FF00) |
                ((x << 8) & 0x00FF0000) |
                (x << 24);
        }
    }
}
