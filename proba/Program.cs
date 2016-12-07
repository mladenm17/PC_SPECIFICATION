using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management;  
using Microsoft.Win32;
using System.Diagnostics;
using System.IO;
//using PdfSharp;
//using PdfSharp.Drawing;
//using PdfSharp.Pdf;
//using PdfSharp.Pdf.IO;
using System.Security.Principal;
using iTextSharp.text.pdf;
using iTextSharp.text;
using iTextSharp.text.pdf.draw;

namespace proba
{
    class Program : Hardware
    {
        static string count;
        static string HDD2;
        static string OS = "";
        static string ComputerName = "";
        static List<string> SoftwareList = new List<string>();
        static List<string> HDDList = new List<string>();
        static string location = "";
        static DateTime now = DateTime.Now;
        public static string name = "";
        public static string fileName = "";
        public static string fullPath = "";
        public static Int32 unixTimestamp = (Int32)(DateTime.UtcNow.Subtract(new DateTime(1991, 7, 18))).TotalSeconds;

        static void Main(string[] args)
        {
            Console.WriteLine("Unesite ime i prezime: ");
            name = Console.ReadLine();
            getOperatingSystemInfo();
            getProcessorInfo();
            GetCompName();
            GetBoardMaker();
            GetPhysicalMemory();
            Getinstalledsoftware();
            getHDDInfo();
            getGPUInfo();
            GetPdf();
            AddPageNumber();
            
            Console.ReadLine();
        }

        #region Functions

        public static void getGPUInfo()
        {
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_VideoController");

           
            foreach (ManagementObject mo in searcher.Get())
            {
                foreach (PropertyData property in mo.Properties)
                {
                    if (property.Name == "Description")
                    {
                        GPU = property.Value.ToString();
                    }
                }
            }      
        }

        public static void GetCompName()
        {
            ComputerName = WindowsIdentity.GetCurrent().Name.ToString();
        }

        public static void getOperatingSystemInfo()
        {
            ManagementObjectSearcher mos = new ManagementObjectSearcher("select * from Win32_OperatingSystem");
            foreach (ManagementObject managementObject in mos.Get())
            {
                if (managementObject["Caption"] != null)
                {
                    OS = (managementObject["Caption"].ToString());
                }
            }
        }

        public static void getProcessorInfo()
        {
            RegistryKey processor_name = Registry.LocalMachine.OpenSubKey(@"Hardware\Description\System\CentralProcessor\0", RegistryKeyPermissionCheck.ReadSubTree);   //This registry entry contains entry for processor info.

            if (processor_name != null)
            {
                if (processor_name.GetValue("ProcessorNameString") != null)
                {
                    CPU = ((processor_name.GetValue("ProcessorNameString").ToString()));   //Display processor ingo.
                }
            }
        }

        public static void GetBoardMaker()
        {
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_BaseBoard");
            ManagementObjectCollection oCollection = searcher.Get();
            foreach (ManagementObject obj in oCollection)
            {
                Motherboard = (obj.GetPropertyValue("Manufacturer").ToString());
                ManagementObjectSearcher objMOS = new ManagementObjectSearcher("\\root\\cimv2", "SELECT * FROM Win32_ComputerSystem");
                foreach (ManagementObject Mobject in objMOS.Get())
                {
                    Motherboard += "  " + Mobject["Model"].ToString();
                }

            }
        }

        public static string GetPhysicalMemory()
        {
            ManagementScope oMs = new ManagementScope();
            ObjectQuery oQuery = new ObjectQuery("SELECT Capacity FROM Win32_PhysicalMemory");
            ManagementObjectSearcher oSearcher = new ManagementObjectSearcher(oMs, oQuery);
            ManagementObjectCollection oCollection = oSearcher.Get();

            long MemSize = 0;
            long mCap = 0;

            foreach (ManagementObject obj in oCollection)
            {
                mCap = Convert.ToInt64(obj["Capacity"]);
                MemSize += mCap;

            }
            MemSize = (MemSize / 1024) / 1024;
            RAM = (MemSize.ToString() + "MB");
            return "";
        }

        public static void Getinstalledsoftware()
        {
            string uninstallKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
            using (RegistryKey rk = Registry.LocalMachine.OpenSubKey(uninstallKey))
            {
                foreach (string skName in rk.GetSubKeyNames())
                {
                    using (RegistryKey sk = rk.OpenSubKey(skName))
                    {
                        try
                        {
                            SoftwareList.Add(sk.GetValue("DisplayName").ToString());
                            SoftwareList.Sort();
                        }
                        catch (Exception ex)
                        { }
                    }
                }
                count = SoftwareList.Count.ToString();

            }
        }

        public static void getHDDInfo()
        {
            foreach (System.IO.DriveInfo label in System.IO.DriveInfo.GetDrives())
            {
                if (label.IsReady)
                {
                    WqlObjectQuery q = new WqlObjectQuery("SELECT * FROM Win32_DiskDrive");
                    ManagementObjectSearcher res = new ManagementObjectSearcher(q);
                    foreach (ManagementObject o in res.Get())
                    {
                        HDD = label.Name + " " + label.TotalSize / 1000000000 + "GB, " + o["Caption"].ToString();
                        HDDList.Add(HDD);
                    }
                }
            }
        }

        public static void GetPdf()
        {
            string path = @"Zapisi";
            if (!Directory.Exists(path))
            {
                DirectoryInfo di = Directory.CreateDirectory(path);
            }
            
            string nameRacunara = ComputerName.Replace('\\', '_');
            fileName = "Zapis_" + nameRacunara + "_" + unixTimestamp + ".pdf";
            fullPath = path + "/"+fileName;

            Document document = new Document(iTextSharp.text.PageSize.A4, 30, 10, 42, 35);
            PdfWriter wri = PdfWriter.GetInstance(document, new FileStream(fullPath, FileMode.Create));
            document.Open();          
            Paragraph text = new Paragraph();
            text.Add("Lorem ipsum dolor sit amet, consectetur adipiscing elit. Curabitur pulvinar risus at lobortis mattis. Pellentesque euismod blandit arcu quis sagittis. Sed pretium eu nisi vel porta. Donec ac magna nisl. Duis at mi ac est imperdiet placerat tempus sit amet eros. Maecenas nec auctor dolor. Aenean quis dolor non purus volutpat egestas at ac est. Morbi in nulla ut neque egestas maximus. Quisque nibh eros, bibendum nec ante vitae, viverra commodo leo. Praesent vitae scelerisque justo. Mauris diam libero, bibendum ac fermentum a, gravida ");
            text.IndentationLeft = 55;
            text.IndentationRight = 50;
            
            document.Add(text);

            //---------------------------------------TABLE
            //---------------------------------------Specification table
 
            Paragraph line = new Paragraph(" ");
            document.Add(line);
            PdfPTable table = new PdfPTable(2);
            BaseFont bfTimes = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, false);
            Font times = new Font(bfTimes, 14, Font.BOLD, BaseColor.BLACK);
            Font times2 = new Font(bfTimes, 10, Font.NORMAL, BaseColor.BLACK);
            PdfPCell cell = new PdfPCell(new Phrase("Specifikacija:", times));
            cell.Colspan = 2;           
            cell.HorizontalAlignment = Element.ALIGN_LEFT;
            table.AddCell(cell);

            table.AddCell("OS: ");
            table.AddCell(OS);

            table.AddCell("Computer name: ");
            table.AddCell(ComputerName);

            document.Add(table);

            //------------------------------------------HARDWARE table
            
            document.Add(line);
            document.Add(line);

            PdfPTable table2 = new PdfPTable(2);

            PdfPCell cell2 = new PdfPCell(new Phrase("Hardver: ", times));
            cell2.Colspan = 2;
            cell2.HorizontalAlignment = Element.ALIGN_LEFT;
            table2.AddCell(cell2);

            table2.AddCell("CPU: ");
            table2.AddCell(CPU);

            table2.AddCell("RAM: ");
            table2.AddCell(RAM);

            table2.AddCell("Motherboard: ");
            table2.AddCell(Motherboard);

            table2.AddCell("GPU: ");
            table2.AddCell(GPU);

            table2.AddCell("HDD: ");
            var p = new Paragraph();
            foreach (var item in HDDList)
            {
                p.Add(item + "\n");
            }
            table2.AddCell(p);

            document.Add(table2);

            //-------------------------------------SOFTWARE table

            Paragraph line2 = new Paragraph(" ");
            document.Add(line2);
            document.Add(line2);

            PdfPTable table3 = new PdfPTable(1);
          
            PdfPCell cell3 = new PdfPCell(new Phrase("Softver: ", times));
            cell3.Colspan = 2;
            cell3.HorizontalAlignment = Element.ALIGN_LEFT;
            table3.AddCell(cell3);

            foreach (var soft in SoftwareList)
            {
                table3.AddCell(new PdfPCell(new Phrase(soft, times2)));
            }

            document.Add(table3);
            Paragraph instp = new Paragraph("Number of installed programs:   " + count);
            instp.IndentationLeft = 56;
            
            document.Add(instp);

            Paragraph text2 = new Paragraph();
            text2.Add("Lorem ipsum dolor sit amet, consectetur adipiscing elit. Curabitur pulvinar risus at lobortis mattis. Pellentesque euismod blandit arcu quis sagittis. Sed pretium eu nisi vel porta. Donec ac magna nisl. Duis at mi ac est imperdiet placerat tempus sit amet eros. Maecenas nec auctor dolor. Aenean quis dolor non purus volutpat egestas at ac est. Morbi in nulla ut neque egestas maximus. Quisque nibh eros, bibendum nec ante vitae, viverra commodo leo. Praesent vitae scelerisque justo. Mauris diam libero, bibendum ac fermentum a, gravida ");
            text2.IndentationLeft = 55;
            text2.IndentationRight = 50;
            text2.SpacingBefore = 20;
            document.Add(text2);

            //----------------------------------underlines

            Paragraph paraLevo = new Paragraph();
            paraLevo.Add("U Novom Sadu,  ");
            paraLevo.Add("\n"+ now.ToShortDateString());
            paraLevo.Alignment = (Element.ALIGN_LEFT + Convert.ToInt32(200));
            paraLevo.SpacingBefore = 120;
            paraLevo.IndentationLeft = 55;
         
            document.Add(paraLevo);

            Paragraph paraDesno = new Paragraph();
            paraDesno.Add(name);
            paraDesno.Add("\n______________\nPotpis");
            paraDesno.Alignment = Element.ALIGN_RIGHT;
            paraDesno.SpacingBefore = -50;
            paraDesno.IndentationRight = 50;
            document.Add(paraDesno);
            document.Close();
        }

        public static void AddPageNumber()
        {
            byte[] bytes = File.ReadAllBytes(fullPath);
            Font blackFont = FontFactory.GetFont("Arial", 12, Font.NORMAL, BaseColor.BLACK);
            using (MemoryStream stream = new MemoryStream())
            {
                PdfReader reader = new PdfReader(bytes);
                using (PdfStamper stamper = new PdfStamper(reader, stream))
                {
                    int pages = reader.NumberOfPages;
                    for (int i = 1; i <= pages; i++)
                    {
                        ColumnText.ShowTextAligned(stamper.GetUnderContent(i), Element.ALIGN_LEFT, new Phrase(i.ToString()+"/"+pages, blackFont), 568f, 15f, 0);
                    }
                }
                bytes = stream.ToArray();
            }
            File.WriteAllBytes(fullPath, bytes);
            string folder = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName) + @"\Zapisi\";
            string pathFull = folder + fileName;
            Process.Start(pathFull);
        #endregion
    }
        }

}