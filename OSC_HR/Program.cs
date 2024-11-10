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
using System.Diagnostics;

namespace OSC_HR
{
    internal class Program
    {
        private const int hrMin = 40;
        private const int hrMax = 200;
        private const float rriMin = 60f / hrMax;
        private const float rriMax = 60f / hrMin;

        private static VRChatOSC vrcOsc = new VRChatOSC();
        private static DateTime lastSent = DateTime.Now;

        private static int last_hr = 0;

        private static Queue<float> rriQueue = new Queue<float>();
        private static bool isRunning = false;

        static async Task Main(string[] args)
        {
            AppDomain.CurrentDomain.ProcessExit += new EventHandler(CurrentDomain_ProcessExit);

            while (!await Connect())
            {
                await Task.Delay(1000);
            }

            while (Process.GetProcessesByName("vrchat").Length == 0)
            {
                await Task.Delay(1000);
            }

            Console.WriteLine("VRChat running");
            vrcOsc.Connect();


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
                        Console.WriteLine("Found Paired Heart Rate Device");
                        GattDeviceService service = bleDevice.GetGattService(new Guid("0000180d-0000-1000-8000-00805f9b34fb"));
                        characteristic = service.GetCharacteristics(new Guid("00002a37-0000-1000-8000-00805f9b34fb")).First();

                        if (service != null && characteristic != null)
                        {
                            Console.WriteLine("Has Heart Rate Characteristic");

                            GattCommunicationStatus status = await characteristic.WriteClientCharacteristicConfigurationDescriptorAsync(GattClientCharacteristicConfigurationDescriptorValue.Notify);
                            if (status == GattCommunicationStatus.Success)
                            {
                                bleDevice.ConnectionStatusChanged += ConnectionStatusChanged;
                                characteristic.ValueChanged += ValueChanged;
                                count++;
                                Console.WriteLine("Subscribed to Heart Rate");
                                break;
                            }
                            else
                            {
                                Console.WriteLine("Failed to Subscribe");
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                    if (bleDevice != null)
                    {
                        bleDevice.Dispose();
                    }
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
                if (last_hr != heartRate)
                {
                    if ((DateTime.Now - lastSent).TotalSeconds > 1.5)
                    {
                        string symbol = "";
                        if (last_hr < heartRate) symbol = "🔺";
                        else if (last_hr > heartRate) symbol = "🔻";

                        vrcOsc.SendChatbox("🤍 " + heartRate, true, false);
                        last_hr = heartRate;
                        lastSent = DateTime.Now;
                        //Console.WriteLine("Sent Heart Rate: " + heartRate);
                    }
                    else
                    {
                        //Console.WriteLine("Read Heart Rate: " + heartRate);
                    }
                }

                if ((flags & (1 << 4)) != 0)//RRI present (16bit)
                {
                    while (reader.UnconsumedBufferLength > 1)
                    {
                        byte a = reader.ReadByte();
                        byte b = reader.ReadByte();
                        int rri = b << 8 | a;

                        EnqueueRRI(rri / 1024f);
                    }
                }
            }
            catch { }
        }


        private static void EnqueueRRI(float rri)
        {
            if (rri < rriMin || rri > rriMax) //rri invalid
            {
                if (last_hr > hrMin && last_hr < hrMax) //use hr if valid
                {
                    rri = 60f / last_hr;
                }
                else //nothing valid assume 60bpm
                {
                    rri = 1f;
                }
            }

            lock (rriQueue)
            {
                rriQueue.Enqueue(rri);
            }

            if (!isRunning)
            {
                StartProcessing();
            }
        }

        private static async void StartProcessing()
        {
            isRunning = true;
            while (true)
            {
                float rri = 0;

                lock (rriQueue)
                {
                    while (true)
                    {
                        float sum = rriQueue.Sum(x => x);
                        int count = rriQueue.Count;

                        Console.WriteLine(count + " " + sum);

                        //2 slow heart beats of 2 seconds of fast hearbeats
                        if (sum > 2 * rriMax || count > 2 * hrMax / 60f)
                        {
                            rriQueue.Dequeue();
                            Console.WriteLine("Skip");
                        }
                        else
                        {
                            break;
                        }
                    }

                    if (rriQueue.Count > 0)
                    {
                        rri = rriQueue.Dequeue();
                    }
                    else
                    {
                        isRunning = false;
                        Console.WriteLine("Stall");
                        return;
                    }
                }


                //do the work
                _ = vrcOsc.SendParameterAsync("Beat", true);

                await Task.Delay((int)(rri * 1000f));
            }
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