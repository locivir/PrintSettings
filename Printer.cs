//  Author: Ernst Plaatsman
//  Email:  ernst@locivir.com
//This is free and unencumbered software released into the public domain.
//
//Anyone is free to copy, modify, publish, use, compile, sell, or
//distribute this software, either in source code form or as a compiled
//binary, for any purpose, commercial or non-commercial, and by any
//means.

//In jurisdictions that recognize copyright laws, the author or authors
//of this software dedicate any and all copyright interest in the
//software to the public domain. We make this dedication for the benefit
//of the public at large and to the detriment of our heirs and
//successors. We intend this dedication to be an overt act of
//relinquishment in perpetuity of all present and future rights to this
//software under copyright law.

//THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
//EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
//MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
//IN NO EVENT SHALL THE AUTHORS BE LIABLE FOR ANY CLAIM, DAMAGES OR
//OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE,
//ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
//OTHER DEALINGS IN THE SOFTWARE.
//
//For more information, please refer to <http://unlicense.org/>
using System;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace Locivir.Printing
{
    /// <summary>
    ///  Saves and restores complete printer configurations, including the driver-specific
    ///  DEVMODE, keyed by a caller-supplied application identity so that multiple
    ///  applications can store multiple named configurations without collision.
    ///
    ///  Configurations are stored as registry values under <see cref="ConfigPath"/>
    ///  (default: HKEY_CURRENT_USER\Software\PrintSettings). Each value is named
    ///  "{ApplicationId}:{label}".
    ///
    ///  A single instance is not thread-safe (it holds mutable printer state); use one
    ///  instance per thread, or synchronize access. Separate instances are independent.
    /// </summary>
    public sealed class PrinterConfig
    {
        /// <summary>Default registry path used when none is supplied.</summary>
        public const string DefaultConfigPath = @"HKEY_CURRENT_USER\Software\PrintSettings";

        // Fixed namespace for deriving deterministic GUIDs from application names (RFC 4122 v5).
        private static readonly Guid AppNameNamespace = new Guid("6f9619ff-8b86-d011-b42d-00c04fc964ff");

        private PrinterSettings _prnSettings = new PrinterSettings();

        /// <summary>
        ///  Creates a configuration store for the given application identity.
        /// </summary>
        /// <param name="applicationId">Stable identity of the owning application. Must not be <see cref="Guid.Empty"/>.</param>
        /// <param name="configPath">Registry path under which configurations are stored.</param>
        public PrinterConfig(Guid applicationId, string configPath = DefaultConfigPath)
        {
            if (applicationId == Guid.Empty)
                throw new ArgumentException("Application id must not be Guid.Empty.", nameof(applicationId));
            if (string.IsNullOrWhiteSpace(configPath))
                throw new ArgumentException("Config path must not be null or empty.", nameof(configPath));

            ApplicationId = applicationId;
            ConfigPath = configPath;
        }

        /// <summary>
        ///  Creates a configuration store, deriving a stable identity from the application name.
        ///  The same name always maps to the same GUID (RFC 4122 v5, SHA-1 name hash).
        /// </summary>
        public PrinterConfig(string applicationName, string configPath = DefaultConfigPath)
            : this(DeriveApplicationId(applicationName), configPath)
        {
        }

        /// <summary>Identity this instance stores configurations under.</summary>
        public Guid ApplicationId { get; }

        /// <summary>Registry path under which configurations are stored.</summary>
        public string ConfigPath { get; }

        /// <summary>The current printer settings held by this instance.</summary>
        public PrinterSettings PrinterSettings => _prnSettings;

        /// <summary>Name of the current printer, or an empty string if none is set.</summary>
        public string PrinterName => _prnSettings.PrinterName ?? string.Empty;

        /// <summary>Whether the current default page is landscape.</summary>
        public bool Landscape => _prnSettings.DefaultPageSettings.Landscape;

        /// <summary>Printable area of the current default page, orientation-corrected, in hundredths of an inch.</summary>
        public Size PrintableArea
        {
            get
            {
                PageSettings page = _prnSettings.DefaultPageSettings;
                RectangleF area = page.PrintableArea;
                return page.Landscape
                    ? new Size((int)area.Height, (int)area.Width)
                    : new Size((int)area.Width, (int)area.Height);
            }
        }

        /// <summary>Paper size of the current default page, orientation-corrected, in hundredths of an inch.</summary>
        public Size PaperSize
        {
            get
            {
                PageSettings page = _prnSettings.DefaultPageSettings;
                int width = page.PaperSize.Width;
                int height = page.PaperSize.Height;
                return page.Landscape
                    ? new Size(height, width)
                    : new Size(width, height);
            }
        }

        /// <summary>
        ///  Loads and applies the configuration stored under <paramref name="label"/>.
        ///  If it is missing or cannot be applied, the print dialog is shown so the user
        ///  can choose settings, which are then saved under the same label.
        /// </summary>
        /// <returns>The resulting printer settings.</returns>
        public PrinterSettings SetPrinter(string label, bool showDialog = false)
        {
            PrinterConfigData data = Load(label);
            if (data == null || !Apply(data))
                showDialog = true;

            if (showDialog)
            {
                using (PrintDocument doc = new PrintDocument())
                {
                    doc.PrinterSettings = _prnSettings;
                    doc.DocumentName = label;
                    using (PrintDialog pd = new PrintDialog { Document = doc })
                    {
                        if (pd.ShowDialog() == DialogResult.OK)
                        {
                            _prnSettings = pd.PrinterSettings;
                            SaveCurrent(label);
                        }
                    }
                }
            }
            return _prnSettings;
        }

        /// <summary>
        ///  Captures the current printer's settings (including DEVMODE) and stores them under <paramref name="label"/>.
        /// </summary>
        /// <returns>True if the settings were captured and stored; false if the DEVMODE could not be read.</returns>
        public bool SaveCurrent(string label)
        {
            if (string.IsNullOrEmpty(label))
                throw new ArgumentException("Label must not be null or empty.", nameof(label));

            string devMode = CaptureDevMode();
            if (devMode == null)
                return false;

            PrinterConfigData data = new PrinterConfigData
            {
                Application = ApplicationId,
                Label = label,
                Printer = _prnSettings.PrinterName,
                DevMode = devMode
            };

            using (StringWriter writer = new StringWriter())
            {
                new XmlSerializer(typeof(PrinterConfigData)).Serialize(writer, data);
                Microsoft.Win32.Registry.SetValue(ConfigPath, MakeValueName(label), writer.ToString());
            }
            return true;
        }

        /// <summary>
        ///  Loads the configuration stored under <paramref name="label"/> and applies it to the current printer.
        /// </summary>
        /// <returns>True if a stored configuration was found and applied; otherwise false.</returns>
        public bool ApplySaved(string label)
        {
            PrinterConfigData data = Load(label);
            return data != null && Apply(data);
        }

        private string MakeValueName(string label) => $"{ApplicationId}:{label}";

        private PrinterConfigData Load(string label)
        {
            string xml = Microsoft.Win32.Registry.GetValue(ConfigPath, MakeValueName(label), null) as string;
            if (string.IsNullOrEmpty(xml))
                return null;

            using (StringReader reader = new StringReader(xml.Trim()))
                return new XmlSerializer(typeof(PrinterConfigData)).Deserialize(reader) as PrinterConfigData;
        }

        /// <summary>
        ///  Reads the current printer's DEVMODE into a base64 string.
        /// </summary>
        private string CaptureDevMode()
        {
            if (string.IsNullOrEmpty(_prnSettings.PrinterName))
                return null;

            // Existence guard: fail fast if the printer cannot be opened.
            if (!SafeNativeMethods.OpenPrinter(_prnSettings.PrinterName, out IntPtr hPrinter, IntPtr.Zero))
                return null;

            IntPtr hDevMode = IntPtr.Zero;
            IntPtr pDevMode = IntPtr.Zero;
            try
            {
                hDevMode = _prnSettings.GetHdevmode(_prnSettings.DefaultPageSettings);
                if (hDevMode == IntPtr.Zero)
                    return null;

                pDevMode = SafeNativeMethods.GlobalLock(hDevMode);
                if (pDevMode == IntPtr.Zero)
                    return null;

                DEVMODE devMode = (DEVMODE)Marshal.PtrToStructure(pDevMode, typeof(DEVMODE));
                int size = devMode.dmSize + devMode.dmDriverExtra;
                if (size <= 0)
                    return null;

                byte[] buffer = new byte[size];
                Marshal.Copy(pDevMode, buffer, 0, size);
                return Convert.ToBase64String(buffer);
            }
            finally
            {
                if (pDevMode != IntPtr.Zero)
                    SafeNativeMethods.GlobalUnlock(hDevMode);
                if (hDevMode != IntPtr.Zero)
                    SafeNativeMethods.GlobalFree(hDevMode);
                SafeNativeMethods.ClosePrinter(hPrinter);
            }
        }

        /// <summary>
        ///  Applies a stored DEVMODE to the target printer named in the configuration.
        /// </summary>
        private bool Apply(PrinterConfigData data)
        {
            if (data == null || string.IsNullOrEmpty(data.Printer) || string.IsNullOrEmpty(data.DevMode))
                return false;

            byte[] savedDevMode;
            try
            {
                savedDevMode = Convert.FromBase64String(data.DevMode);
            }
            catch (FormatException)
            {
                return false;
            }

            // Validate the blob is internally consistent before handing it to native code.
            int minSize = Marshal.SizeOf(typeof(DEVMODE));
            if (savedDevMode.Length < minSize)
                return false;

            DEVMODE header;
            GCHandle pin = GCHandle.Alloc(savedDevMode, GCHandleType.Pinned);
            try
            {
                header = (DEVMODE)Marshal.PtrToStructure(pin.AddrOfPinnedObject(), typeof(DEVMODE));
            }
            finally
            {
                pin.Free();
            }

            int declared = header.dmSize + header.dmDriverExtra;
            if (declared <= 0 || savedDevMode.Length < declared)
                return false;

            _prnSettings = new PrinterSettings { PrinterName = data.Printer };

            // Existence guard: fail (and let the caller fall back to the dialog) if the printer is gone.
            if (!SafeNativeMethods.OpenPrinter(data.Printer, out IntPtr hPrinter, IntPtr.Zero))
                return false;

            IntPtr hDevMode = IntPtr.Zero;
            try
            {
                // Allocate our own buffer sized to the saved blob, so we never write past a
                // buffer sized for a different (e.g. current) driver's DEVMODE.
                hDevMode = SafeNativeMethods.GlobalAlloc(GHND, (UIntPtr)(uint)savedDevMode.Length);
                if (hDevMode == IntPtr.Zero)
                    return false;

                IntPtr pDevMode = SafeNativeMethods.GlobalLock(hDevMode);
                if (pDevMode == IntPtr.Zero)
                    return false;

                try
                {
                    Marshal.Copy(savedDevMode, 0, pDevMode, savedDevMode.Length);
                }
                finally
                {
                    SafeNativeMethods.GlobalUnlock(hDevMode);
                }

                _prnSettings.SetHdevmode(hDevMode);
                return true;
            }
            catch (InvalidPrinterException)
            {
                return false;
            }
            finally
            {
                if (hDevMode != IntPtr.Zero)
                    SafeNativeMethods.GlobalFree(hDevMode);
                SafeNativeMethods.ClosePrinter(hPrinter);
            }
        }

        /// <summary>
        ///  Derives a stable GUID from an application name (RFC 4122 v5, SHA-1 name hash).
        /// </summary>
        private static Guid DeriveApplicationId(string applicationName)
        {
            if (string.IsNullOrWhiteSpace(applicationName))
                throw new ArgumentException("Application name must not be null or empty.", nameof(applicationName));

            byte[] namespaceBytes = AppNameNamespace.ToByteArray();
            SwapGuidByteOrder(namespaceBytes); // .NET stores the first three fields little-endian; RFC needs network order.

            byte[] nameBytes = Encoding.UTF8.GetBytes(applicationName);

            byte[] hash;
            using (SHA1 sha1 = SHA1.Create())
            {
                sha1.TransformBlock(namespaceBytes, 0, namespaceBytes.Length, null, 0);
                sha1.TransformFinalBlock(nameBytes, 0, nameBytes.Length);
                hash = sha1.Hash;
            }

            byte[] result = new byte[16];
            Array.Copy(hash, 0, result, 0, 16);
            result[6] = (byte)((result[6] & 0x0F) | 0x50); // version 5
            result[8] = (byte)((result[8] & 0x3F) | 0x80); // RFC 4122 variant
            SwapGuidByteOrder(result);
            return new Guid(result);
        }

        private static void SwapGuidByteOrder(byte[] guid)
        {
            SwapBytes(guid, 0, 3);
            SwapBytes(guid, 1, 2);
            SwapBytes(guid, 4, 5);
            SwapBytes(guid, 6, 7);
        }

        private static void SwapBytes(byte[] guid, int left, int right)
        {
            byte temp = guid[left];
            guid[left] = guid[right];
            guid[right] = temp;
        }

        // GMEM_MOVEABLE | GMEM_ZEROINIT
        private const uint GHND = 0x0042;

        /// <summary>
        ///  Native methods used by this class.
        /// </summary>
        private static class SafeNativeMethods
        {
            [DllImport("winspool.drv", EntryPoint = "OpenPrinterW", CharSet = CharSet.Unicode,
                SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
            public static extern bool OpenPrinter(string pPrinterName, out IntPtr phPrinter, IntPtr pDefault);

            [DllImport("winspool.drv", EntryPoint = "ClosePrinter", SetLastError = true,
                ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
            public static extern bool ClosePrinter(IntPtr hPrinter);

            [DllImport("kernel32.dll", ExactSpelling = true, SetLastError = true)]
            public static extern IntPtr GlobalAlloc(uint uFlags, UIntPtr dwBytes);

            [DllImport("kernel32.dll", ExactSpelling = true)]
            public static extern IntPtr GlobalLock(IntPtr hMem);

            [DllImport("kernel32.dll", ExactSpelling = true)]
            public static extern bool GlobalUnlock(IntPtr hMem);

            [DllImport("kernel32.dll", ExactSpelling = true)]
            public static extern IntPtr GlobalFree(IntPtr hMem);
        }

        // DEVMODE (printer variant). Only the fixed header up to dmDriverExtra is read,
        // but the full documented layout is declared so the marshaled size is correct.
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct DEVMODE
        {
            private const int CCHDEVICENAME = 32;
            private const int CCHFORMNAME = 32;

            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = CCHDEVICENAME)]
            public string dmDeviceName;
            public short dmSpecVersion;
            public short dmDriverVersion;
            public short dmSize;
            public short dmDriverExtra;
            public int dmFields;

            public short dmOrientation;
            public short dmPaperSize;
            public short dmPaperLength;
            public short dmPaperWidth;
            public short dmScale;
            public short dmCopies;
            public short dmDefaultSource;
            public short dmPrintQuality;

            public short dmColor;
            public short dmDuplex;
            public short dmYResolution;
            public short dmTTOption;
            public short dmCollate;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = CCHFORMNAME)]
            public string dmFormName;
            public short dmLogPixels;
            public int dmBitsPerPel;
            public int dmPelsWidth;
            public int dmPelsHeight;
            public int dmDisplayFlags;
            public int dmDisplayFrequency;
            public int dmICMMethod;
            public int dmICMIntent;
            public int dmMediaType;
            public int dmDitherType;
            public int dmReserved1;
            public int dmReserved2;
            public int dmPanningWidth;
            public int dmPanningHeight;
        }
    }

    /// <summary>
    ///  Serializable data-transfer object for a single stored printer configuration.
    /// </summary>
    [Serializable]
    public sealed class PrinterConfigData
    {
        /// <summary>Identity of the application that owns this configuration.</summary>
        public Guid Application { get; set; }

        /// <summary>Human-readable label for this configuration.</summary>
        public string Label { get; set; } = "Default";

        /// <summary>Target printer name.</summary>
        public string Printer { get; set; }

        /// <summary>Base64-encoded DEVMODE for the target printer.</summary>
        public string DevMode { get; set; }
    }
}
