using NAudio.CoreAudioApi;
using NAudio.CoreAudioApi.Interfaces;
using System.Runtime.InteropServices;

namespace SwitchMicMonitorDevice
{
    internal class Program
    {
        private static Guid GUID_MONITORING = new("{24DBB0FC-9311-4B3D-9CF0-18FF155639D4}");

        private static PropertyKey PROP_KEY_MONITORING_DEVICE_ID = new(GUID_MONITORING, 0);
        private static PropertyKey PROP_KEY_MONITORING_ENABLE = new(GUID_MONITORING, 1);

        private static PropVariant PROP_BOOL_ENABLE = new() { vt = (short)VarEnum.VT_BOOL, boolVal = -1 };
        private static PropVariant PROP_BOOL_DISABLE = new() { vt = (short)VarEnum.VT_BOOL, boolVal = 0 };

        static void Main(string[] args)
        {
            MMDeviceEnumerator enumerator = new MMDeviceEnumerator();

            if (args.Length == 0)
            {
                Console.WriteLine("Usage: {0}.exe \"{1}\" \"{2}\"", System.AppDomain.CurrentDomain.FriendlyName, "{ID of output/mic device}", "{ID of input/speaker device}");

                Console.WriteLine("Output devices:");
                foreach (var dev in enumerator.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active | DeviceState.Disabled | DeviceState.Unplugged))
                {
                    Console.WriteLine("{0}\t{1}\t({2})", dev.ID, dev.FriendlyName, dev.State);
                }

                Console.WriteLine("Input devices:");
                foreach (var dev in enumerator.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active | DeviceState.Disabled | DeviceState.Unplugged))
                {
                    Console.WriteLine("{0}\t{1}\t({2})", dev.ID, dev.FriendlyName, dev.State);
                }

                enumerator.Dispose();

                return;
            }

            string micId = args[0];

            MMDevice device = enumerator.GetDevice(micId);
            if (device == null)
            {
                enumerator.Dispose();

                Console.WriteLine("Device {0} not found", micId);

                return;
            }

            device.GetPropertyInformation(StorageAccessMode.ReadWrite);

            PropertyStore store = device.Properties;

            store.SetValue(PROP_KEY_MONITORING_ENABLE, PROP_BOOL_DISABLE);
            store.Commit();

            if (args.Length > 1)
            {
                string outDeviceId = args[1];

                MMDevice dev2 = enumerator.GetDevice(outDeviceId);
                if (dev2 == null) return;

                store.SetValue(PROP_KEY_MONITORING_ENABLE, PROP_BOOL_ENABLE);
                store.Commit();

                IntPtr pointer = Marshal.StringToHGlobalAuto(outDeviceId);

                PropVariant value = new()
                {
                    vt = (short)VarEnum.VT_LPWSTR,
                    pointerValue = pointer
                };

                store.SetValue(PROP_KEY_MONITORING_DEVICE_ID, value);
                store.Commit();

                Marshal.FreeHGlobal(pointer);

                Console.WriteLine("Enabled device monitoring for {0}", device.FriendlyName);
                Console.WriteLine("Target device: {0}", dev2.FriendlyName);
            }
            else
            {
                Console.WriteLine("Disabled device monitoring for {0}", device.FriendlyName);
            }

            enumerator.Dispose();
        }
    }
}
