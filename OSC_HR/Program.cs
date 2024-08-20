using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Windows.Devices.Bluetooth;
using Windows.Devices.Bluetooth.GenericAttributeProfile;
using Windows.Devices.Enumeration;
using Windows.Storage.Streams;
using VRChatOSCLib;
using System.Text;

namespace OSC_HR
{
    internal class Program
    {
        private const float timescale = 1000;
        private const bool testing = false;
        private const bool record = true;
        private const float staleDataTime = 20;
        private const float stalenessCatchupFactor = 0.1f;

        private const int hrMin = 40;
        private const int hrMax = 180;
        private const float rriMin = 60f / hrMax;
        private const float rriMax = 60f / hrMin;

        private static VRChatOSC vrcOsc = new VRChatOSC(true);

        private static int last_hr = 0;

        static async Task Main(string[] args)
        {
            AppDomain.CurrentDomain.ProcessExit += new EventHandler(CurrentDomain_ProcessExit);

            while (!await Connect())
            {
                await Task.Delay(100);
            }

            while (true)
            {
                await Task.Delay(1000);
            }
        }


        private static BluetoothLEDevice bleDevice;
        private static GattCharacteristic characteristic;

        private static async Task<bool> Connect()
        {
            int count = 0;

            foreach (var device in await DeviceInformation.FindAllAsync())
            {
                try
                {
                    bleDevice = await BluetoothLEDevice.FromIdAsync(device.Id);

                    if (bleDevice != null && bleDevice.Appearance.Category == BluetoothLEAppearanceCategories.HeartRate)
                    {
                        GattDeviceService service = bleDevice.GetGattService(new Guid("0000180d-0000-1000-8000-00805f9b34fb"));
                        characteristic = service.GetCharacteristics(new Guid("00002a37-0000-1000-8000-00805f9b34fb")).First();

                        if (service != null && characteristic != null)
                        {
                            Console.WriteLine("Found Paired Heart Rate Device");

                            GattCommunicationStatus status = await characteristic.WriteClientCharacteristicConfigurationDescriptorAsync(GattClientCharacteristicConfigurationDescriptorValue.Notify);
                            if (status == GattCommunicationStatus.Success)
                            {
                                bleDevice.ConnectionStatusChanged += ConnectionStatusChanged;
                                characteristic.ValueChanged += ValueChanged;
                                count++;
                                Console.WriteLine("Subscribed to Heart Rate");
                                break;
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    //Console.WriteLine(e.Message);
                }
            }
            return count > 0;
        }

        private static void ValueChanged(GattCharacteristic sender, GattValueChangedEventArgs args)
        {
            try
            {
                var reader = DataReader.FromBuffer(args.CharacteristicValue);

                byte flags = reader.ReadByte();

                int heartRate = -1;

                if ((flags & (1 << 0)) != 0)//16bit HR
                {
                    byte a = reader.ReadByte();
                    byte b = reader.ReadByte();
                    heartRate = b << 8 | a;
                }
                else//8bit HR
                {
                    heartRate = reader.ReadByte();
                }
                Console.WriteLine("Heart Rate: " + heartRate);
                if (last_hr != heartRate)
                {
                    string symbol = "";
                    if (last_hr < heartRate) symbol = "🔺";
                    else if (last_hr > heartRate) symbol = "🔻";

                    vrcOsc.SendChatbox("🤍 " + heartRate, true, false);
                    last_hr = heartRate;
                }
            }
            catch { }
        }

        private static void ConnectionStatusChanged(BluetoothLEDevice sender, object args)
        {
            Console.WriteLine("Connection status: " + sender.ConnectionStatus);
        }

        private static void CurrentDomain_ProcessExit(object sender, EventArgs e)
        {
            if (bleDevice != null)
            {
                bleDevice.ConnectionStatusChanged -= ConnectionStatusChanged;
                characteristic.ValueChanged -= ValueChanged;
                bleDevice.Dispose();
            }

            vrcOsc.Dispose();
        }
    }
}