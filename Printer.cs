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
using System.Windows.Forms;
using System.Xml.Serialization;

namespace Locivir.Printing
{
    /// <summary>
    ///  Custom printer configuration that can save all print settings including the device specific settings for multiple configurations per application 
    ///  The configuration is saved in a registry subkey of HKEY_CURRENT_USER. Adjust the constant "ConfigPath" to adjust the subkey.
    ///  
    /// Printer configurations are saved by name with an unique generated GUID for each different application.
    /// </summary>
    public static class Printer
    {
        /// <summary>
        /// Registry path to save the configuration
        /// </summary>
        public const string ConfigPath = @"HKEY_CURRENT_USER\Software\PrintSettings";

        /// <summary>
        ///  Internal class to help save the user settings
        /// </summary>
        [Serializable]
        public class UserSetting //: TypeConverter
        {
            public Guid Application { get; set; } = ApplicationGuid;
            public string Label { get; set; } = "Default";
            public string Printer { get; set; }
            public string DevMode { get; set; }

            public UserSetting() { SaveDevMode(); }

            public UserSetting(string label) : this() { Label = label; } //SaveDevMode(); }

            public static UserSetting Load(string label)
            {
                string configString = (string)Microsoft.Win32.Registry.GetValue(ConfigPath,
                    string.Format("{0}:{1}", ApplicationGuid.ToString(), label), string.Empty);
                if (!string.IsNullOrEmpty(configString))
                    return new XmlSerializer(typeof(UserSetting)).Deserialize(new StringReader(configString.Trim())) as UserSetting;
                else
                    return null;
            }

            public bool Save()
            {
                if (SaveDevMode())
                {
                    //Save the settings
                    using (StringWriter textWriter = new StringWriter())
                    {
                        XmlSerializer xs = new XmlSerializer(typeof(UserSetting));
                        xs.Serialize(textWriter, this);
                        Microsoft.Win32.Registry.SetValue(ConfigPath,
                            string.Format("{0}:{1}", Application.ToString(), Label), textWriter.ToString());
                    }
                }
                else return false;

                return true;
            }

            private bool SaveDevMode()
            {
                IntPtr hDevMode;
                IntPtr pDevMode;
                DEVMODE devMode;
                byte[] aDevMode;

                Printer = prnSettings.PrinterName;

                // Open the printer
                if (SafeNativeMethods.OpenPrinter(prnSettings.PrinterName.Normalize(), out IntPtr hPrinter, IntPtr.Zero))
                {
                    // Get the default printer settings
                    hDevMode = prnSettings.GetHdevmode(prnSettings.DefaultPageSettings);

                    // Obtain a lock on the handle 
                    pDevMode = SafeNativeMethods.GlobalLock(hDevMode);

                    try
                    {
                        // Marshal the memory into our DEVMODE
                        devMode = (DEVMODE)Marshal.PtrToStructure(pDevMode, typeof(DEVMODE));

                        // Copy the DEVMODE to a byte array
                        aDevMode = new byte[devMode.dmSize + devMode.dmDriverExtra];
                        for (int i = 0; i < devMode.dmSize + devMode.dmDriverExtra; ++i)
                            aDevMode[i] = Marshal.ReadByte(pDevMode, i);
                    }
                    catch
                    {
                        return false;
                    }
                    finally
                    {
                        // Unlock the handle                
                        SafeNativeMethods.GlobalUnlock(hDevMode);
                        SafeNativeMethods.GlobalFree(hDevMode);

                        // Close the printer
                        SafeNativeMethods.ClosePrinter(hPrinter);
                    }

                    // Save the byte array as a string
                    DevMode = System.Convert.ToBase64String(aDevMode);
                }
                else return false;

                return true;
            }

            public bool SetActive()
            {
                IntPtr hDevMode;
                IntPtr pDevMode;
                byte[] savedDevMode;

                prnSettings = new PrinterSettings { PrinterName = Printer };

                // Open the printer.
                if (SafeNativeMethods.OpenPrinter(prnSettings.PrinterName.Normalize(), out IntPtr hPrinter, IntPtr.Zero))
                {
                    // get the current DEVMODE position in memory
                    hDevMode = prnSettings.GetHdevmode(prnSettings.DefaultPageSettings);

                    // Obtain a lock 
                    pDevMode = SafeNativeMethods.GlobalLock(hDevMode);

                    try
                    {
                        // Load and convert the saved DEVMODE string into a byte array
                        savedDevMode = System.Convert.FromBase64String(DevMode);

                        // Overwrite the current DEVMODE 
                        for (int i = 0; i < savedDevMode.Length; ++i)
                            Marshal.WriteByte(pDevMode, i, savedDevMode[i]);

                        // It is good to go. All done
                        SafeNativeMethods.GlobalUnlock(hDevMode);

                        // Upload our printer settings to use the one we just overwrote
                        prnSettings.SetHdevmode(hDevMode);
                    }
                    catch
                    {
                        SafeNativeMethods.GlobalUnlock(hDevMode);
                        return false;
                    }
                    finally
                    {
                        SafeNativeMethods.GlobalFree(hDevMode);
                        SafeNativeMethods.ClosePrinter(hPrinter);
                    }
                }
                else
                {
                    return false;
                }

                return true;
            }
        }

        /// <summary>
        ///  Contains native methods used by this class.
        /// </summary>
        [System.Security.SuppressUnmanagedCodeSecurity]
        private static partial class SafeNativeMethods
        {
            /// <summary>
            ///  Opens a printer handle given the printer name
            /// </summary>
            [DllImport("winspool.Drv", EntryPoint = "OpenPrinterA", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true,
                BestFitMapping = false, ThrowOnUnmappableChar = true, CallingConvention = CallingConvention.StdCall)]
            public static extern bool OpenPrinter([MarshalAs(UnmanagedType.LPStr)] string szPrinter, out IntPtr hPrinter, IntPtr pd);

            /// <summary>
            ///  Closes a printer connection given a handle
            /// </summary>
            [DllImport("winspool.Drv", EntryPoint = "ClosePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
            public static extern bool ClosePrinter(IntPtr hPrinter);

            /// <summary>
            ///  Notifies the print spooler that data should be written to the specified printer.
            /// </summary>
            [DllImport("winspool.Drv", EntryPoint = "WritePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
            public static extern bool WritePrinter(IntPtr hPrinter, IntPtr pBytes, Int32 dwCount, out Int32 dwWritten);

            [DllImport("kernel32.dll", ExactSpelling = true)]
            public static extern IntPtr GlobalFree(IntPtr handle);

            [DllImport("kernel32.dll", ExactSpelling = true)]
            public static extern IntPtr GlobalLock(IntPtr handle);

            [DllImport("kernel32.dll", ExactSpelling = true)]
            public static extern bool GlobalUnlock(IntPtr handle);
        }

        // DEVMODE data structure 
        [StructLayout(LayoutKind.Sequential, Pack = 1, CharSet = CharSet.Auto)]
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

            public int dmPositionX;
            public int dmPositionY;
            public int dmDisplayOrientation;
            public int dmDisplayFixedOutput;

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

        private static PrinterSettings prnSettings = new PrinterSettings();
        private static UserSetting settings;
        public static string PrinterName
        {
            get
            {
                return PrinterSettings.PrinterName;
            }
        }
        public static PrinterSettings PrinterSettings
        {
            get
            {
                if (prnSettings == null) prnSettings = new PrinterSettings();
                return prnSettings.DefaultPageSettings.PrinterSettings;
            }
        }
        public static bool Landscape { get { return PrinterSettings.DefaultPageSettings.Landscape; } }
        public static Size PrintableArea
        {
            get
            {
                Size page = new Size();
                if (PrinterSettings.DefaultPageSettings.Landscape)
                {
                    page.Width = (int)PrinterSettings.DefaultPageSettings.PrintableArea.Height;
                    page.Height = (int)PrinterSettings.DefaultPageSettings.PrintableArea.Width;
                }
                else
                {
                    page.Width = (int)PrinterSettings.DefaultPageSettings.PrintableArea.Width;
                    page.Height = (int)PrinterSettings.DefaultPageSettings.PrintableArea.Height;
                }
                return page;

            }
        }
        public static Size PaperSize
        {
            get
            {
                Size page = new Size();
                if (PrinterSettings.DefaultPageSettings.Landscape)
                {
                    page.Width = PrinterSettings.DefaultPageSettings.PaperSize.Height;
                    page.Height = PrinterSettings.DefaultPageSettings.PaperSize.Width;
                }
                else
                {
                    page.Width = PrinterSettings.DefaultPageSettings.PaperSize.Width;
                    page.Height = PrinterSettings.DefaultPageSettings.PaperSize.Height;
                }
                return page;

            }
        }

        public static PrinterSettings SetPrinter(string configLabel) => SetPrinter(configLabel, false);
        public static PrinterSettings SetPrinter(string configLabel, bool showDialog)
        {
            settings = UserSetting.Load(configLabel);
            if (settings == null || !settings.SetActive())
            {
                showDialog = true;
                settings = new UserSetting(configLabel);
            }

            if (showDialog)
                using (PrintDocument doc = new PrintDocument())
                {
                    doc.PrinterSettings = PrinterSettings;
                    doc.DocumentName = configLabel;
                    using (PrintDialog pd = new PrintDialog())
                    {
                        pd.Document = doc;
                        if (pd.ShowDialog() == DialogResult.OK)
                            try
                            {
                                prnSettings = pd.PrinterSettings;
                                settings.Save();
                            }
                            catch { }
                    }
                }
            return prnSettings;
        }

        /// <summary>
        /// Returns ths Guid of the application that is running.
        /// </summary>
        private static Guid ApplicationGuid
        {
            get
            {
                if (System.Reflection.Assembly.GetEntryAssembly() != null)
                    foreach (System.Reflection.CustomAttributeData attr in System.Reflection.Assembly.GetEntryAssembly().CustomAttributes)
                        if (attr.AttributeType == typeof(System.Runtime.InteropServices.GuidAttribute))
                            return Guid.Parse((string)attr.ConstructorArguments[0].Value);
                //throw new ApplicationException("Could not get Application GUID.");
                return Guid.Empty;
            }
        }

    }

}