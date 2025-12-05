using NAudio.CoreAudioApi;
using NAudio.CoreAudioApi.Interfaces;
using System.Diagnostics;
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
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: {0}.exe \"{1}\" \"{2}\"", System.AppDomain.CurrentDomain.FriendlyName, "{ID or name of output/mic device}", "{ID or name of input/speaker device}");
                Console.WriteLine("Running with only output device argument will disable mic monitoring for that device.");

                MMDeviceEnumerator enumerator = new MMDeviceEnumerator();

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

            MMDevice device = FindDevice(micId, DataFlow.Capture);
            if (device == null)
            {
                Console.WriteLine("Output device {0} not found", micId);

                return;
            }

            device.GetPropertyInformation(StorageAccessMode.ReadWrite);

            PropertyStore store = device.Properties;

            store.SetValue(PROP_KEY_MONITORING_ENABLE, PROP_BOOL_DISABLE);
            store.Commit();

            if (args.Length > 1)
            {
                string outDeviceName = args[1];

                MMDevice dev2 = FindDevice(outDeviceName, DataFlow.Render);
                if (dev2 == null)
                {
                    Console.WriteLine("Input device {0} not found", outDeviceName);

                    return;
                }

                string outDeviceId = dev2.ID;

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
        }

        private static MMDevice FindDevice(string nameOrID, DataFlow flow = DataFlow.All)
        {
            MMDeviceEnumerator enumerator = new();

            foreach (var dev in enumerator.EnumerateAudioEndPoints(flow, DeviceState.Active | DeviceState.Disabled | DeviceState.Unplugged))
            {
                if (dev.ID.Equals(nameOrID) || dev.FriendlyName.Contains(nameOrID))
                {
                    enumerator.Dispose();

                    return dev;
                }
            }

            enumerator.Dispose();

            return null;
        }
    }
}
