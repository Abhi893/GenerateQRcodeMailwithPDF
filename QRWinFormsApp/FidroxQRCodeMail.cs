using System;
using System.Windows.Forms;
using OfficeOpenXml;
using System.IO;
//using Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Collections.Generic;
using QRCoder;
using System.Text;
using System.Net;
using System.Net.Mail;
using System.Configuration;
using ClosedXML.Excel;
using System.Data;
using System.Reflection.Metadata;
using PdfSharpCore;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.html.simpleparser;
using System.Drawing;
using Font = iTextSharp.text.Font;
using iTextSharp.text.html;
using System.Data.SqlClient;

namespace QRWinFormsApp
{



    public partial class FidroxQRCodeMail : Form
    {
        public FidroxQRCodeMail()
        {
            InitializeComponent();
        }
       string connection = ConfigurationManager.ConnectionStrings["DBContext"].ToString();

        public List<EmployeeModel> UpdateList=new List<EmployeeModel>();
        static byte[] byteImage;
        string ImageUrl;

       
        

        public TimeZoneInfo INDIAN_ZONE { get; private set; }


        //send mail trigger
        private void button2_Click(object sender, EventArgs e)
        {
            var successCount = 0;
            if (UpdateList.Count == 0)
            {
                MessageBox.Show("Excel Data isEmpty");
            }
            else
            {
                foreach (var emp in UpdateList)
                {
                    var qrCode = LoadQrCode(emp.txtEmpNo);
                    var senEmail = SendEmail(emp.txtEmpEmail, qrCode, emp.txtEmpNo);
                   
                    //var qrCode=LoadQrCode()
                    // sendmail(emp.EmpName,qrCode)
                    if (senEmail == "Mail Sent Successfully")
                    {
                        DeletePDF(emp.txtEmpNo);
                        DeleteQRCodeImage(emp.txtEmpNo);
                        successCount++;
                    }
                }
                if (successCount == UpdateList.Count)
                {
                    MessageBox.Show("Mails Sent Successfully");
                }
            }
        }
        //Upload Button
        private void button1_Click(object sender, EventArgs e)
        {
            //To where your opendialog box get starting location. My initial directory location is desktop.
            openFileDialog1.InitialDirectory = "C://Desktop";
            //Your opendialog box title name.
            openFileDialog1.Title = "Select file to be upload.";
            //which type file format you want to upload in database. just add them.
            openFileDialog1.Filter = "Select Valid Document(*.pdf; *.doc; *.xlsx; *.html)|*.pdf; *.docx; *.xlsx; *.html";
            //FilterIndex property represents the index of the filter currently selected in the file dialog box.
            openFileDialog1.FilterIndex = 1;
            UpdateList = new List<EmployeeModel>();
            try
            {
                if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    if (openFileDialog1.CheckFileExists)
                    {
                        string path = System.IO.Path.GetFullPath(openFileDialog1.FileName);
                        label1.Text = path;
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        using (ExcelPackage package = new ExcelPackage(new FileInfo(path)))
                        {
                            //ExcelWorksheets currentSheet;
                            //package.Workbook.Worksheets worksheets = new package.Workbook.Worksheets();
                            //ExcelWorksheets currentSheet = package.Workbook.Worksheets;
                            ExcelWorksheets currentSheet = package.Workbook.Worksheets;
                            ExcelWorksheet workSheet = currentSheet.First();
                       //     ExcelWorksheet workSheet = currentSheet.First();
                            int noOfCol = workSheet.Dimension.End.Column;
                            int noOfRow = workSheet.Dimension.End.Row;
                            if(noOfRow==1)
                            {
                                MessageBox.Show("Please add the data in excel");
                            }
                            for (int rowIterator = 2; rowIterator <= noOfRow; rowIterator++)
                            {
                                EmployeeModel uploadExcel = new EmployeeModel();
                                if (noOfCol == 2)
                                {
                                    try
                                    {
                                       
                                            if (workSheet.Cells[rowIterator, 1].Value.ToString() != null && workSheet.Cells[rowIterator, 2].Value.ToString() != null)
                                            {

                                                uploadExcel.txtEmpNo = workSheet.Cells[rowIterator, 1].Value.ToString();
                                                uploadExcel.txtEmpEmail = workSheet.Cells[rowIterator, 2].Value.ToString();
                                                UpdateList.Add(uploadExcel);
                                            }
                                        
                                       
                                    }
                                    catch (Exception ex)
                                    {

                                    }
                                }
                            }
                        }




                    }
                }
                else
                {
                    MessageBox.Show("Please Upload document.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //Download Sample Excel Template
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {


            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                XLWorkbook workbook = new XLWorkbook();
                DataTable dt = new DataTable() { TableName = "New Worksheet" };
                DataSet ds = new DataSet();

                //input data
                var columns = new[] { "EmpNumber", "EmpMail", };
               

                //Add columns
                dt.Columns.AddRange(columns.Select(c => new DataColumn(c)).ToArray());

             

                //Convert datatable to dataset and add it to the workbook as worksheet
                ds.Tables.Add(dt);
                workbook.Worksheets.Add(ds);

                //save
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string savePath = Path.Combine(desktopPath, "QRcode.xlsx");
                workbook.SaveAs(savePath, false);
                //ExcelPackage Ep = new ExcelPackage();
                //ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("EmpNoUpload");
                //Sheet.Cells["A1"].Value = "Emp No";
                //Sheet.Cells["B1"].Value = "Emp Mail";

                //Sheet.Protection.IsProtected = true; //--------Protect whole sheet
                //var start = Sheet.Dimension.Start;
                //var end = Sheet.Dimension.End;
                //for (int row = start.Row; row <= end.Row; row++)
                //{ // Row by row..
                //    for (int col = start.Column; col <= end.Column; col++)
                //    { // ... Cell by cell...
                //        Sheet.Column(col).Style.Locked = false;// This got me the actual value I needed.
                //    }
                //}
                //Sheet.Row(1).Style.Locked = true;


                //Sheet.Cells["A:AZ"].AutoFitColumns();
                //Sheet.Row(1).Style.Locked = true;

                ////Response.Clear();
                //HttpResponse response =HttpContext.
                //Sheet.Cells["A:AZ"].AutoFitColumns();
                //response.ContentType = "application/vnd.openxmlformats-     officedocument.spreadsheetml.sheet";
                ////response.BinaryWrite(Ep.GetAsByteArray());
                MessageBox.Show("Downloaded,Saved In Desktop ");
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }



        private string GenerateQRCode(string EmpNo)
        {
            try
            {
                EmpNo += "CL";

                while (EmpNo.Length < 15)
                {
                    EmpNo += '0';
                }

                int empNoLength = EmpNo.Length;

                string encrypvmob = Crypt(EmpNo);

                int encryptLength = encrypvmob.Length;

                QRCodeGenerator qrGenerator = new QRCodeGenerator();
                QRCodeData qrCodeData = qrGenerator.CreateQrCode(encrypvmob, QRCodeGenerator.ECCLevel.H);

                Base64QRCode qrcode = new Base64QRCode(qrCodeData);
                ////System.UI.WebControls.Image imgBarCode = new System.Web.UI.WebControls.Image();
                ////imgBarCode.Height = 150;
                ////imgBarCode.Width = 150;
                //using (Bitmap bitMap = qrcode.GetGraphic(20))
                //{
                //    using (MemoryStream ms = new MemoryStream())
                //    {
                //        bitMap.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                //        byteImage = ms.ToArray();
                //        ImageUrl = "data:image/png;base64," + Convert.ToBase64String(byteImage);
                //    }
                //    plBarCode.Controls.Add(imgBarCode);
                //}
                return (qrcode.GetGraphic(7));
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public string Crypt(string text)
        {
            try
            {
                byte[] ascibytes = Encoding.ASCII.GetBytes(text);
                string encrpt = "";
                foreach (byte bary in ascibytes)
                {
                    encrpt = encrpt + bary;
                }
                return encrpt;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public string LoadQrCode(string txtEmpNo)
        {
            try
            {
                //var empNo = employeeService.GetEmployeeDetailsByEmpNo(EmpNo);
                //if (empNo.Count > 0)
                //{
                    var QRcode = "data:image/png;base64," + GenerateQRCode(txtEmpNo);
                // return Json(QRcode, JsonRequestBehavior.AllowGet);
                //}
                //else
                //{
                //    //return Json("Employee Number " + EmpNo + " Not Registered", JsonRequestBehavior.AllowGet);
                //}
                return QRcode;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        //TO generate QRcode image for PDF
        public string LoadQrCode1(string txtEmpNo)
        {
            try
            {
                //var empNo = employeeService.GetEmployeeDetailsByEmpNo(EmpNo);
                //if (empNo.Count > 0)
                //{
                var QRcode =GenerateQRCode(txtEmpNo);
                // return Json(QRcode, JsonRequestBehavior.AllowGet);
                //}
                //else
                //{
                //    //return Json("Employee Number " + EmpNo + " Not Registered", JsonRequestBehavior.AllowGet);
                //}
                return QRcode;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public string SendEmail( string txtEmpEmail,string qrCode,string txtEmpNo) 
        {

         
            string signature = string.Empty;
            string fromMailAddress = ConfigurationManager.AppSettings["FromAddress"].ToString();
            string toMailAddress = txtEmpEmail;
            string Empnumber = txtEmpNo;
            string fromPwd = ConfigurationManager.AppSettings["FromPwd"].ToString();
            string Signature = ConfigurationManager.AppSettings["Signature"].ToString();
            //string Name = ConfigurationManager.AppSettings["Name"].ToString();
            string query = ConfigurationSettings.AppSettings["Query"].ToString();
            var credentials = new NetworkCredential(fromMailAddress, fromPwd);
            string fromName = ConfigurationSettings.AppSettings["FromName"].ToString();
            var portNo = int.Parse(ConfigurationSettings.AppSettings["sendPort"].ToString().Replace(Environment.NewLine, string.Empty));
            string hostDomain = ConfigurationSettings.AppSettings["Hostdomain"].ToString();
            string toName = ConfigurationSettings.AppSettings["ToName"].ToString();
            string strsubject = ConfigurationSettings.AppSettings["Subject"].ToString();
            string body = ConfigurationSettings.AppSettings["TestMail"];
            string teamAdmin = ConfigurationSettings.AppSettings["TeamAdmin"];
            string pleaseNote = ConfigurationSettings.AppSettings["PleaseNote"];
            string request = ConfigurationSettings.AppSettings["Request"];
            string furtherDetails = ConfigurationSettings.AppSettings["FurtherDetails"];
            string extension = ConfigurationSettings.AppSettings["Extension"];
            string signature1 = ConfigurationManager.AppSettings["Signature1"].ToString();
            string signature2 = ConfigurationManager.AppSettings["Signature2"].ToString();
            string signature3 = ConfigurationManager.AppSettings["Signature3"].ToString();
            string signature4 = ConfigurationManager.AppSettings["Signature4"].ToString();
            string signature5 = ConfigurationManager.AppSettings["Signature5"].ToString();

            string QRCodepdffiles = ConfigurationManager.AppSettings["QRCodePdf"].ToString();
            string QRcodelogfile = ConfigurationManager.AppSettings["QRCodeLogs"].ToString();

            string Message0 = null;
            Message0 = body;
            //Message0 = Message0.Replace("#Signature#", Signature);

            //           string htmlString = "<!DOCTYPE html>" +
            //"<html> " +
            //    "<body> " +

            //     " <div style='border: solid;height:320px; width: 188px; '>" +
            //          " <div id='QRCode' style='text-align:center; '> #QR# </div>" +
            //                    "</div>" +
            //                    "</body> " +
            //"</html>";


            //htmlString = htmlString.Replace("#QR#", "<img  src='"+qrCode+"'"+"/>");
            //htmlString = htmlString.Replace("#QR#", "<img  src='data:image/png;base64,/9j/4AAQSkZJRgABAQAAAQABAAD/2wCEAAkGBwgHBgkIBwgKCgkLDRYPDQwMDRsUFRAWIB0iIiAdHx8kKDQsJCYxJx8fLT0tMTU3Ojo6Iys / RD84QzQ5OjcBCgoKDQwNGg8PGjclHyU3Nzc3Nzc3Nzc3Nzc3Nzc3Nzc3Nzc3Nzc3Nzc3Nzc3Nzc3Nzc3Nzc3Nzc3Nzc3Nzc3N//AABEIAIIAggMBIgACEQEDEQH/xAAcAAEAAgMBAQEAAAAAAAAAAAAAAQcFBggCAwT/xABAEAABAwIDBgMFBgQDCQAAAAABAAIDBAUGESExQVFhcbEHEoETIkKRoSMyUmLB8BQVcpIWJEMlM2OCosLR4fH/xAAUAQEAAAAAAAAAAAAAAAAAAAAA/8QAFBEBAAAAAAAAAAAAAAAAAAAAAP/aAAwDAQACEQMRAD8Au/8AfRRn/wCuaH9jinHXqUDPj6lM/nuC+c80dPC6ad7Y442lznPOQaOJKqHGvilNO+Shww4xQ55PrnN99/8AQDsHM68MkFj4hxbZcPNyudaxs2WYgj96Vw/pGwczkFW168Ya+UuZZKCKmbukqT53f2jID5lVnJI+WR0kr3Pkec3OeSS48STtXlBnq/GeJbgXGovNWAfhif7MfJuSw81VU1BznqZpTxkkLu6+KIPccssRzjlew8WuI7LKUeJ7/RH/AC15r2ZbjO5w+RzCxCILAtHi1fqRwbcYqeviHFvs3/MafRWJhzxFsF8LIXVBoqpxyENUQ3zHg12w9NvJc9og6yzPr2TPTl3VAYP8RLph5zKerc6uto0ML3e/GPyO/Q6dFd9kvNBfaBtdbKgTRO0J2Fh/CRtBQZDP98EB2fQKOH0Cnjr1KD1nzUJrwCIIPH5lfOeaKnhfNO9sccbS5znnINA2kr6Hb2Cp3xgxcaiofhy3yH2MRBrXtP33bQzoNCeem4oMF4hY5nxLVPpKJ74rTE73W7DOR8TuXAep12aYiICIs/hHCdxxTWGKiAjp4yPbVMg92Plzdy7IMBmAM89F94aOrqBnTUlTMP8AhQuf2C6Aw/4fYfsjGO/g21dQ3bUVTQ85/lbsb6BbUxoYAGNDQNjQMgEHLEltuMQzmt1bGOMlM9vcL8uYzI3jaust2p07rC3rCtjvbHC422B8hH+9Y3yPb/zDVBzOi3nHXh3V4djkuFve6rtgObyR9pCPzDe3mPXitGQFl8MYir8M3JtbQPzachNA4+5K3gRx4HaPnniEQdPYdvlFiG1R3CgeTG/R7HfejdvaeYWU37Og4LnPAOKpML3pssjibfUODKpnAbnjm3tmF0VFIyWNskbg+NwDmuBzDgdhCD36lE9/dkiDB40vgw9hyruAy9s1vkgadjpDo30z1PIFc1yyPlkfLK8vke4uc521xO0qy/HG7e2ulDZ2O9ynj9vIPzO0b8gD/cqxQEREH3oaSavraeiphnNUSNiZ1cch6Lpqw2ilsVpgt1E3yxQtyLstXu3uPMlUZ4U07ajHVB5tkTZJB1DT/wCV0Hw06Dggemu4cFPPd3UbuXdTv59kD5Z9lG7Zp3TTLl3U8Tn1PBB5exr2ua9oc1wycDqMuC508QsPsw5iaelgaRSTAT04/C07W+hBHTJdG8PoFU/jvA3yWeq+LzSxk8tD+iCpUREBXZ4M4hdX2eW0VLwZ6DIxEnbCdnyOY6ZKk1sXh/dv5Ni631DjlDJJ7Cbh5X+7r0OR9EHSOQKJkFCDmvHlaa/GN2n82YFQ6NvRnujssCvtWSmorKiY7ZJXvPqSf1XxQEREG4+Ekgix1Ra6vjlYP7Sf0XQGmW3Tuqd8HMLQ1jziGollDqWo8lOxhyBPl94njo7LJXHr69ggb+fZNPTum7l3Tv2QOp17J6dAo9OgT16lA9epVV+O8g/hrPDsPtJX5csgP1Vq99w4LSPFDC0N8s0tx9rIyqt8D3xDP3XAe84EcwNqChEREBCSBm05Eag80Q7EHT9mu8FZaKGqfIPNNTxyHqWgoqPt+KpqWgpqcPdlFE1g14ABEGpysMU0kZ2tcWn0OS8LLYtpDQ4pu9OdjKuQgcifMPoViUBERBb3gZdYnUtwtDnATMk/iYwT95pAa7LoQPmrT0y5d1yxa7jVWm4Q19BKYqiF3ma4b+IPI7F0zY7pBerTS3KlP2VRGH5fgO9vUHMIP38Tv48FHDToFPD6BRxz9SgcdepU8PoOCemu4cFG7XPLugnnu7rWPEi6x2rB9we9wEtRGaaFu8ufp9BmfRbOduZ0/Rc6+IeJH4ixDOYpS630zjHTM+HTQv5kn6ZINXREQERHaDPhqgysFmnmgjlbG4h7Q4HqFKvOwYaijsNtZKPtG0sQdpv8gzRBW3jTa3UeKY69o+yroAdnxs0P08ir9dB+KNhN8wtO6CMvqqPOoiy1Lsh7zQN+bc9OIC58QEREBWN4Q4sba611luEgbSVTwadzjpHKdMujtPXqq5RB1jx16lT33Dgqz8Kca1t2eLJcYZJ5oYi6OrAGQYMhlJz1yB37+Ksznu7oI3a//VOufPsnfstX8QMUSYWszamCkdPNO/2UTj9xjsic378tNg28kGH8WMWNtFqdaaGXKvrGZOLdsUR0LuROoHqdyo0DLZsX6K+tqbjWzVlbM6aomd5nyO2lfnQEREBZTC9sdecRW63tGk07Q/T4Bq7/AKQVi1a3gjYiX1V+mZoM6enJ9C9w+g+aC3QAAAAAAiZD9lEEHbpt7KgPE/Cxw/ejVU0eVtrXF8WWyN/xM/Ucuiv8/TusffrPSX21z2+4M80Uo2jaw7nDmCg5dRZfFGHq7DV0fRVzc2n3oZgPdlZuI58Ru+S+WH7U29XOOiNfS0Rk+7JUkgE8BkNvIkIMavcMUk8rIYGOklkcGsY0ZlxOwAK4rf4O26MD+Y3OqqH72wtETfrmVt2H8HWHD7va22ha2bZ7eRxkk9Cdnpkgx3hthP8AwzaHPqwDcKoh0+WvkG5gPLfzJ5LcN/Psn76J24cUDTLl3WOxBaKa/WiqttYPs5m5ecbY3DVrhzB1WR149TwThp0CDl6/WWtsFzkt9xj8krNWuH3ZG7nNPArHLqG82O2Xym9hdaOKpjacw5w95p/KRqPRaZX+ENimB/gqqtpXbh5xI0fMZ/VBSKLaMaYQ/wAKyxxuu1JVOfshaC2UDiW6jLnmtdpKWetqoqWkifNPK4MjjYMy4oP14es1Vf7xT2yiH2krvfflpGwfeceg/RdLWm309qt1NQUbPLBAwMYN55nrt9Vr+AMHxYWtpM3lkuNQAaiRuo5Mafwj6nVbX+yUE68QiacEQQf3yUenQKT06BPXqUGLxDYbfiK3PornEJIyc2vH3o3cWncVQ+MsE3PC8pdOx1TbnH3Kpjfd6PHwn6FdGa+u4cF4kjZLG6ORjXxuGTmuGYcOBCCg8K+I95sLGU0/+0KFugimdk9g/K/hyOforTsfiJhy8BjRXNo6h2Q9jVkR68Afun0KwuJvCe2173T2SX+XznUwkeaE+m1vppyVaXrBGIrN5nVdtklhH+tTD2rOumo9QEHR7Hte0OY4OadhBzBXrf3PBcr0Vzr7c/8AyNdU0pHwxTOZ8wCszDj3FcIybe6ggfjax3dqDo7oOgU889N5C50k8QcWSjL+dzN/pijH/asTXYhvNeMq67Vkrd7XzuDT6Z5IOhb1i+w2QEV9zp2ygZiCN3nk/tGqrTE/izW1jX0+H4DRxHT+JlydKRyGxv19FpFow7ebw/y2y2VMwP8AqBnlZ/cch9VYWHfCBxcyXEVWANppqU/Rzz2A9UFe2i0XbE1zdFQxS1VQ93mmmeTk3P4nvOzurzwPgihwtB7XNtTcZG5SVJblpva0bm6DmctVsFstlFaaRtJbqWOnp27I425ZnieJ5lfr/ZKBv7ngo4adAnDToFP7J4oJ9SilEA71G9EQNx6p8foiIA2BQERBi73abbWUr31dvpJ3filga4/UKjcYUVJS1Dm01NDCOEcYb2REGHsUMU1Qxs0bJATse0FXthOx2hlM2RlroWyZfeFOwH55IiDaSAAABkOSN2KEQTvPRTuChEDceqlEQeCTmdVCIg//2Q=='>");
            string htmlString = "<!DOCTYPE html>" +
                           "<html> " +
                             "<body> " +
                              " <div>" +

                              " <div> #Name#</div>" +
                                "<br>" +                  
                               " <div> #TestMail#</div>" +
                              "<br>" +
                              "<br>" +
                              " <div> #PleaseNote#</div>" +
                              "<br>" +
                              "<br>" +
                               " <div> #Request#</div>" +
                              "<br>" +
                              "<br>" +
                               " <div> #FurtherDetails#</div>" +
                              "<br>" +
                              "<br>" +
                                   " <div style='text-align:Center'> #QR#" +
                               " <h3 style ='color:black;font-family:Gotham;'>#Emp#</h3></div>" +
                                   "<br>" +
                                   " <div> #Query#</div>" +
                                   "<br>" +
                                   " <div> #Extension#</div>" +
                                   "<br>" +
                                     " <div> #Signature#</div>" +
                                      " <div> #TeamAdmin#</div>" +
                                       "<br>" +
                                       "<br>" +
                                       " <div> #Signature1#</div>" +
                                        " <div> #Signature2#</div>" +
                                      
                                        " <div> #Signature3#</div>" +
                                       
                                        " <div> #Signature4#</div>" +
                                       
                                        " <div> #Signature5#</div>" +
                                       
                                             "</div>" +
                                             "</body> " +
                         "</html>";



                                 htmlString = htmlString.Replace("#Name#", toName);
            htmlString = htmlString.Replace("#TestMail#", body);

            htmlString = htmlString.Replace("#QR#", " <img  src='" + qrCode + "'>");
            htmlString = htmlString.Replace("#Emp#", Empnumber);
            htmlString = htmlString.Replace("#Query#", query);
            htmlString = htmlString.Replace("#Signature#", Signature);
            htmlString = htmlString.Replace("#FurtherDetails#", furtherDetails);
            htmlString = htmlString.Replace("#TeamAdmin#", teamAdmin);
            htmlString = htmlString.Replace("#PleaseNote#", pleaseNote);
            htmlString = htmlString.Replace("#Request#", request);

            htmlString = htmlString.Replace("#Extension#", extension);
            htmlString = htmlString.Replace("#Signature1#", signature1);
            htmlString = htmlString.Replace("#Signature2#", signature2);
            htmlString = htmlString.Replace("#Signature3#", signature3);
            htmlString = htmlString.Replace("#Signature4#", signature4);
            htmlString = htmlString.Replace("#Signature5#", signature5);

            // message.Body = string.Format(htmlString);

            Message0 = Message0.Replace("#qrCode", string.Format(htmlString));

          
          


            

            //Message0 = Message0.Replace("#Signature#", signature);




            // Create and build a new MailMessage object
            //MailMessage message = new MailMessage();
            //message.IsBodyHtml = true;
            //message.From = new MailAddress(fromMailAddress, fromName);
            //message.To.Add(new MailAddress(toMailAddress));
            //message.Subject = strsubject;
            //message.Body = body;
            // Comment or delete the next line if you are not using a configuration set
            // message.Headers.Add("X-SES-CONFIGURATION-SET", CONFIGSET);

            try
            {
                //using (var client = new System.Net.Mail.SmtpClient(HOST, PORT))

                GeneratePDF(txtEmpNo);
             
                var mail = new MailMessage();
                var smtpclient = new SmtpClient();
                mail = new MailMessage()
                {
                    From = new MailAddress(fromMailAddress, fromName),
                    Subject = strsubject,
                    Body = string.Format(htmlString)
                };
                mail.IsBodyHtml = true;
                mail.To.Add(new MailAddress(toMailAddress));

                // Set up the PDF attachment

                string PDFfile = txtEmpNo+ ".pdf";
               // string PDFeFullPath = "D:/QRCodePDFFile/" + PDFfile;
                string PDFeFullPath = QRCodepdffiles+PDFfile;
                Attachment attachment = new Attachment(PDFeFullPath);
                mail.Attachments.Add(attachment);

                //byte[] bytes = memoryStream.ToArray();
                //memoryStream.Close();
                //mail.Attachments.Add(new Attachment(new MemoryStream(bytes), "D:/QRCodePDFFile/ '" + txtEmpNo + "'.pdf"));


                smtpclient = new SmtpClient()
                {
                    Port = portNo,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    UseDefaultCredentials = false,
                    Host = hostDomain,
                    EnableSsl = true,
                    Credentials = credentials
                };

                smtpclient.Send(mail);
                mail.Dispose();
                string logsd = "QRCodeLog.txt";
                string logFilePath = QRcodelogfile+ logsd;
                if (!File.Exists(logFilePath))
                {
                    File.Create(logFilePath).Dispose();
                }

                // Write the log message to the file.
                using (StreamWriter sw = File.AppendText(logFilePath))
                {
                   
                    string logMessage = $"{DateTime.Now}: Mail is Sent to {toMailAddress},{txtEmpNo}";
                    sw.WriteLine(logMessage);
                }

                //DeletePDF(txtEmpNo);
                //DeleteQRCodeImage(txtEmpNo);


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
            return "Mail Sent Successfully";
            //MessageBox.Show("Mails Send Successfully");


        }

        //Generate PDF
        public string GeneratePDF(string txtEmpNo)
        {
            try
            {
                // foreach (var emp in UpdateList)
                //{
                //    //var qrCode = LoadQrCode(emp.txtEmpNo);

                //    ////var qrCode=LoadQrCode()
                //    //// sendmail(emp.EmpName,qrCode)
                //    //var renderer = new HtmlToPdf();

                //    //renderer.RenderHtmlAsPdf(" <img  src='" + qrCode + "'>").SaveAs("D:/QRCodePDFFile/ '"+emp.txtEmpNo +"'.pdf");

                //    // renderer.RenderHtmlAsPdf("<h1>'"+emp.txtEmpNo + "'</h1>").;
                

                   // var webClient = new System.Net.WebClient();
                    var qrCode = LoadQrCode1(txtEmpNo);
                    string base64Image = qrCode;
                 //  string base64Image = "iVBORw0KGgoAAAANSUhEUgAAAUwAAAAvCAYAAACIefMfAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAhdEVYdENyZWF0aW9uIFRpbWUAMjAyMjowNjowMyAyMDoyODoyMRENRd4AAAyeSURBVHhe7Z15jBRFFMafAoIoHqCId9RE45F4Jx4xMTFciVEkiqKIQNTERMAT/xHveAdBY1BQVBBBRDwREFEQD7wwHoiAeMASFEFRQFdByvq2pqaOru6umd3t3YX3S142093VVV0z/fV79ap6d9i8pY1o3WozMQzDMNnsWPrLMAzD5MCCyTAMEwkLJsMwTCQsmAzDMJGwYDIMw0TCgskwDBMJCybDMEwkLJgMwzCRsGAyDMNEwit9Cqamhmjr1tIHjwMOkE8wfoS1GP78k+iPP4iEKG2ogP33J2rVqvSBaTGwYBbI008TXXUVUW1taUOAnj2J+vYlOusson33lV/QDqUdTLNj9GiiYcOINm4sbaiA8eOJ+vQhatu2tIFpEbA/UyAbNuR7IzNmEPXvrzyQDz8k+u+/0g5mm2L5cv5uWyIsmE0IvJMHHiAaPjzsaZx+OtHq1dWFfEzj06FD9RHAYYdxSN4S4ZC8QB55hOjGG4n++Ud9/uYboiOOMOOWX39NdMopRJs2qc/grbeIzjyTb67myqpVrqf4889qWOW330obJJ99RrTXXqUPJXgMs2XCglkgeYIJXnyR6NJLif7+W30eN46oXz+iNm3UZ59Q4qGS5JFfvqnKgqLqBmvXqj7W5XfbjWj33es/ZgwBPe44dX4NRHSffUofcgj1SyXi6pdnYW5gIJiyc9kKsIcfJiFDb/yU60wKppDeiXPMvHkk2rc3x0yeTGLzZvcYfdzgwSR22skca9uTT5L4999kOW3LlpEYOjS9vG1vvum24bvvSFxzjXstaTZrllv23XdJDBmSXnbsWBK1teZ433Td7dqFy9s2Y0ay76QXL669lsTOOyePv/12Ej/9lPxOKrGaGhLSm3TOKwUzeKxt8+er67K/e9tGjyYhBT5YFrZ8OYnrr08vb9trr5GQD+3gediyjQWzQIsRzKlT3Zv5/fdJbNniHnPhhSSkx1k+Js2kNxu8yb76ikSHDuEyIXv5ZSM8KCu9seBxIUNZLdwXXxzXbtz4mza5bdZ1Sy8wWCZk06a5Dw18zmt79+5KlKsVzWoEU0YUUQ8fCOqGDcnyixYl68yyKVOyH0ps6caCWaDlCebHH5PYZRezH4IX+mHbojFoEIn77yfx4IMkhg9P3nizZ7uCK8M1p3yPHiTeeIPEnDkkHnoo7HnBK8Q5UHaPPcz2mLLwhHX9HTua7QMHkrjvPtNuv+zMma53iLrt8hC2119XdY8c6fabtrlzzTlwDXbbL7vMtB3euP0Awb6QMMVYNYIpw/Xysf37u/2y667uuXDN9kMA/WKX79qVxKuvqusaNSr8gMBvIiv6YEs3FswCzRfMLLvrrvSwCR7m+PEkVq0isXWruw+ia99kOI8tuhAwez9CUPscCFd1mP7EE24bUNYWlh9/dMted525PoTWfvvhYT7zjBIV34P75BNXyO+4g8Rff5n9EDz75kcIatd9ww0mTB8zJvmgueQSc10hQYRA2/W/804ynI+xaj3Mp54isXJlsl8+/ZREp07mXLfeSmLjRrMfobz9IFmyxO2Xm24yYTrCertP2Sq3CobJmaKRoWFwVdDkySoxtN9+ySTFySdnD/L/8IM55513qmSEfY5evYiksNQxdSqR9ETKSIEsZ4RDZc8915RF8kontzQTJ5o5pn6C5qSTiFq3Ln0IgHbruqWYJib1n3OOmZrl1y1Fh6ZPN9cyYACRFFeHU091E2sLFxJJz7gQMIkdbQolrk48MT3hB+QDr9wvt9yi+tbul7PPNtc6bZpJJjLVwYLZTLn5ZqKjjyaSHhlJT6fBkN5lme+/Twryr7+mL91EWX0zhsoiMyyfwo2CPecRdWuR0KDutHYvXuyKnwzNE8KELLktNFi901jX0pDY3wkeaL7Ir1uX3i9M5bBgNiGYVoQbHzfmmjVES5cSjRzpej+DBhEtWxb+0cswtm7a0YgRRL17K+8ONw+mlaRx0EFGLGQYWDdHUIuPDCdpyhQj0D16uN5NTFnt2XXvbrxNH3h8drvhGaLduLnTOPBA4zljiSnq1uKAul94Ib1uf4XV8cerc6FO2/CwaEpwTehX9Mv556vfAdqFaUlpwCvV/QJPFX2rvz9McUK/1JaW4nbrlvSsmQrhMcziLCZLDkM2VN7w5eNQzh4PRLb4ggviMs7+GCbMHqeEIXlz5ZXJKUZffplsnz1OaZe1t8G++CJZFu3u0yeu3f4YJgzZc3nDl49B4gd129tgn3/u1v3oo+GkUJZhilE1433VjGFiqlPfvnHj2/4YJswep4Qh8XPFFckpRhgn1gk4tuqMPcxmyHnnueN5mIysgTd12mnKc9CeBCbDv/SSWhU0Zw6RFIdMOnVyQ9KZM4nGjHHHK+fPJzrqqGToirLwejS6rD1mOG+eGk6wy6LdZ5zherBSAOvG1XS7ERZn4dc9a5aqW3tQYO5comOOSbbb5vHHiWbPVnWm2eWXF/NiDHiBWMk1aZLpQ/nArBuH1f3SsaPangb229eLaxs7lkgKfhmc69hjeRJ7vWEPsziL9TCRjba9AztjLsNRZwrOe+8lvQY72+t7mJjMrb0tZJaffVZ5Lb16kbj3XjVvcu3aZPZdl9UZ9lBZKdqpZZEdt708e7qRNjsb7HuYyGLrDD08TV23fLiIe+5RdcuQOli372H6HmhDWqUe5oQJbvZfClsiO9+li9nve5hYVLDnnmofIhOc77bbSPTuTeLuu9Xc0zVrwv3CVrmxYBZosYKJ4+ww0w7J/XB9/fpk+SzBvOgiExK/8kplU2cQNuqy9mT2GIPI2de+bl3ymCzBxJQkfd0Qx0rmEWJKkt0n0ntrtHmIlQrmsGHuA3D16uQxWYKJeZv6tyKjjsRULraGtYzAhWls8CailSuJVqxQhqTPqFHqLUZ2mIl3Y+oQ3Q9LEa7LL7IMQjkd8uaB9zlK0awL+95+WxkyrWhLKMlk89hjRFI0o8v6YaPfbgwp2GF9FqG6Me0ore4jj3RDUYTxaYk0ZNulwEb3YX3J6xdcp/1byALXhX60+wUzCtAv8sHMNATsYRZnvocZY+PGud4Qlk7a4TpW+iBclTdJXSjmnz8rJM+z555z67ZD8jybONEtC6/OLouVPjgf2o0Ei7/SJyskzzOE676nFUr8IIGFEHbSJBIjRqgwVnux1YbtlXqY8JZt7xce4/Tpql/QB35/Z4XkeYZhkaz16Gz5xoJZoFUqmKEXb2ApXN5acoiovvFjsuRZhmWHdnk/S55lWC6py6Ld9nBAyAYMMOFlTJY8yyCAvjhgrNPPHKdZ2nBJnlUqmOiXfv2y+xQiqtvtCybMz5JnGZbRhtbps8UZC2aBhrXLofXW2nr2VOuI4TVgjC9toB6D+BAU/yaDUGIq0IIFxjPxBdP3ULGuWIZudR4NDCLt33xLl6q2wEu0vTS00y/re0RYqqeFB0kZtMfvAwglpiGh3drb8gUTyQv73PBO7bqffz7pgS5enBQ9jGdialPaAwMPBCSkqh3j9AWwWzcSv/8ePlYb+gUJGr/vIJQLF5L46CMzvusLJsaS7TXyWGvu94u9Hxaa8sUWZyyYLdggnFgLri3vJsBaZTsjC7EMJW58UYUI4/x26Oi/8k2bH3qHbk4IxIoVlbXbDjsRnofqhqjaopkVVvttiGlHY5vfprw5k/BmO3c21+u/mEObL6p43wDPx6zOOOnTgtl7b7X6RlvW3EPgL2fs0iVcRnowCSopK39YmciQtW7lTmy77XXkoD51a/w2xLSjsfHblDdnEv0iha8M1vbXt1+YbJr4J8I0JVinjonT+mZChhaTzq++2kx6xvLIQw5J3ohYhpdWVv+LjbSy9WXCBDW7QAs46kZme/BgI/ZYHnnooU0vgkWCl5vYMwXQL1iAMGSIWS7btSv/P6H6wP+iYjsCNxA8F33zxID16lh7DRE8+GCi9etLOyKQoR+dcEL9b060G8Jr/5+cPBYsUG9A2paFAf1y+OFEv/xS2hABBBT/NyrrzVBMOuxhbkdg6SFezjB0aPqLMTQDBxItWWIET5fFsr22OUsGUfbbb9WryRpCsFA3xBd1t8t5eQRek4a6815zty2AfvngA7XEtH370sYU8Fq9RYvUa+xYLKuHPcztFLxrs6ZG/X9seJwIrfGOSYyDYQ05xgntCfI2KIOQOK1s2lhaQ6DrxpgqvF1dd+fOav16Y9bdnEF/6O9E9wu+Q/v73B77paFhwWQYhomEnzkMwzCRsGAyDMNEwoLJMAwTCQsmwzBMJCyYDMMwkbBgMgzDRMKCyTAMEwkLJsMwTCQsmAzDMFEQ/Q+7Q979WiqQHwAAAABJRU5ErkJggg=="; // Replace with your base64 image string
                    byte[] imageBytes = Convert.FromBase64String(base64Image);
                // File.WriteAllBytes("D:/QRCodePDFFile/ '" + emp.txtEmpNo + "'.jpg", imageBytes);

                    string QRCodepdffiles = ConfigurationManager.AppSettings["QRCodePdf"].ToString();

                     string imagefile = txtEmpNo + ".jpg";
                    string imageFullPath = QRCodepdffiles + imagefile;

                    string PDFfile = txtEmpNo + ".pdf";
                    //string PDFeFullPath = "D:/QRCodePDFFile/" + PDFfile;
                    string PDFeFullPath = QRCodepdffiles + PDFfile;

                   
                    File.WriteAllBytes(imageFullPath, imageBytes);
                    using (var document = new iTextSharp.text.Document())
                    {
                        //using (var writer = PdfWriter.GetInstance(document, new FileStream("D:/QRCodePDFFile/ '" +emp.txtEmpNo + "'.pdf", FileMode.Create)))
                        using (var writer = PdfWriter.GetInstance(document, new FileStream(PDFeFullPath, FileMode.Create)))
                        {
                            document.Open();

                            // Add the image to the document
                            //var image = iTextSharp.text.Image.GetInstance("D:/QRCodePDFFile/ '" + emp.txtEmpNo + "'.jpg");
                            var image = iTextSharp.text.Image.GetInstance(imageFullPath);
                            
                            document.Add(image);

                        Paragraph paragraph = new Paragraph(txtEmpNo, FontFactory.GetFont(FontFactory.HELVETICA, 12));
                      //  Paragraph paragraph = new Paragraph("Employee Number:- ", FontFactory.GetFont(FontFactory.HELVETICA, 12));
                        //paragraph.Alignment = Element.ALIGN_CENTER;
                        paragraph.IndentationRight = 100;
                        paragraph.IndentationLeft = 95;
                        //paragraph.Alignment=
                        //PdfPTable pdfPTable = new PdfPTable();
                        //PdfPTable.HorizontalAlignment = Element.ALIGN_CENTER;
                        //paragraph.SpacingAfter = 0;
                        //Chunk chunk1 = new Chunk(txtEmpNo, FontFactory.GetFont(FontFactory.HELVETICA, 12));
                        //paragraph.Add(chunk1);

                        // Add the second chunk of text to the paragraph
                        //Chunk chunk2 = new Chunk(txtEmpNo, FontFactory.GetFont(FontFactory.HELVETICA, 12));
                        //paragraph.Add(chunk2);
                        document.Add(paragraph);
                       
                        
                            document.Close();
                        }
                    } 

                }

          
            
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return null;


        }
        /// <summary>
       
        /// </summary>
        /// <returns></returns>
        /// 
        public string DeletePDF(string txtEmpNo)
        {
           
            try
            {
                string QRCodepdffiles = ConfigurationManager.AppSettings["QRCodePdf"].ToString();
                string PDFfile =txtEmpNo + ".pdf";
                string filePath = QRCodepdffiles + PDFfile;

              

                //if (File.Exists(filePath))
                //{
                //    File.Delete(filePath);
                //    Console.WriteLine("File deleted successfully.");
                //}
                //else
                //{
                //    Console.WriteLine("File does not exist.");
                //}

                using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    //pictureBox1.Image = Image.FromStream(stream);
                    stream.Dispose();
                    if (File.Exists(filePath))
                    {
                        File.Delete(filePath);
                        Console.WriteLine("File deleted successfully.");
                    }
                    else
                    {
                        Console.WriteLine("File does not exist.");
                    }
                }
                


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return null;
        }

        public string DeleteQRCodeImage(string txtEmpNo)
        {
            try
            {
                string QRCodepdffiles = ConfigurationManager.AppSettings["QRCodePdf"].ToString();
                string imagefile =txtEmpNo + ".jpg";
                string imageFullPath = QRCodepdffiles + imagefile;


                //using (FileStream fs = File.Delete(imageFullPath))
                //{
                //    fs.Write(item.File, 0, item.File.Length);
                //}
              //  File.Delete(newPath, item.File);
                if (File.Exists(imageFullPath))
                {
                    File.Delete(imageFullPath);
                    Console.WriteLine("File deleted successfully.");
                }
                else
                {
                    Console.WriteLine("File does not exist.");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return null;
        }


        //Activate Employee
        private void button3_Click(object sender, EventArgs e)
        {
            string QRcodelogfile = ConfigurationManager.AppSettings["QRCodeLogs"].ToString();
           
            var successCount = 0;
            if (UpdateList.Count == 0)
            {
                MessageBox.Show("Excel Data isEmpty");
            }
            foreach (var emp in UpdateList)
            {
                
                try
                {

                    string updateQuery = string.Empty;
                    string SelectQuery = string.Empty;
                    SelectQuery = "select * from tblCMSEmployee where txtEmpNo=" + "'" + emp.txtEmpNo + "'";
                    using (SqlConnection sqlCnn = new SqlConnection(connection))
                    {
                        sqlCnn.Open();
                        using (SqlCommand command = new SqlCommand(SelectQuery, sqlCnn))
                        {
                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                               
                                if (reader.HasRows)
                                {
                                    sqlCnn.Close();
                                    
                                    updateQuery = "UPDATE tblCMSEmployee SET flgActive=1 where txtEmpNo=" + "'" + emp.txtEmpNo + "'";

                                    sqlCnn.Open();
                                    SqlCommand sqlCmd = new SqlCommand(updateQuery, sqlCnn);

                                    int result = sqlCmd.ExecuteNonQuery();

                                    string txtActivatedEmployee = "ActivatedEmployeeLog.txt";
                                    string ActivatedEmployeeLog = QRcodelogfile+ txtActivatedEmployee;
                                    if (!File.Exists(ActivatedEmployeeLog))
                                    {
                                        File.Create(ActivatedEmployeeLog).Dispose();
                                    }

                                    // Write the log message to the file.
                                    using (StreamWriter sw = File.AppendText(ActivatedEmployeeLog))
                                    {

                                        string logMessage = $"{DateTime.Now}: Employee Number :- {emp.txtEmpNo} is Activated Successfully";
                                        sw.WriteLine(logMessage);
                                    }

                                    successCount++;




                                }
                                else
                                {
                                    string txtExceptionfiles = "ExceptionEmployeeLog.txt";
                                    string ExceptionEmployeeLog = QRcodelogfile + txtExceptionfiles;
                                    if (!File.Exists(ExceptionEmployeeLog))
                                    {
                                        File.Create(ExceptionEmployeeLog).Dispose();
                                    }

                                    // Write the log message to the file.
                                    using (StreamWriter sw = File.AppendText(ExceptionEmployeeLog))
                                    {

                                        string logMessage = $"{DateTime.Now}: Employee Number :- {emp.txtEmpNo} not Exit";
                                        sw.WriteLine(logMessage);
                                    }
                                }
                                reader.Dispose();
                            }


                            sqlCnn.Close();
                        }
                    }

                }

                catch (Exception ex)
                {
                    throw new Exception(ex.Message);
                }

            }
            if (successCount == UpdateList.Count)
            {
                MessageBox.Show("Employee Activated Successfully");
            }
        }

        //Deactivate Employee
        private void button4_Click(object sender, EventArgs e)
        {
            var successCount = 0;
            if (UpdateList.Count == 0)
            {
                MessageBox.Show("Excel Data isEmpty");
            }
            foreach (var emp in UpdateList)
            {

                try
                {
                    string QRcodelogfile = ConfigurationManager.AppSettings["QRCodeLogs"].ToString();
                    string updateQuery = string.Empty;
                    string SelectQuery = string.Empty;
                    SelectQuery = "select * from tblCMSEmployee where txtEmpNo=" + "'" + emp.txtEmpNo + "'";
                    using (SqlConnection sqlCnn = new SqlConnection(connection))
                    {
                        sqlCnn.Open();
                        using (SqlCommand command = new SqlCommand(SelectQuery, sqlCnn))
                        {
                            using (SqlDataReader reader = command.ExecuteReader())
                            {

                                if (reader.HasRows)
                                {
                                    sqlCnn.Close();
                                    updateQuery = "UPDATE tblCMSEmployee SET flgActive=0 where txtEmpNo=" + "'" + emp.txtEmpNo + "'";

                                    sqlCnn.Open();
                                    SqlCommand sqlCmd = new SqlCommand(updateQuery, sqlCnn);

                                    int result = sqlCmd.ExecuteNonQuery();
                                    string txtdeactive = "DeactivatedEmployeeLog.txt";
                                    string ActivatedEmployeeLog = QRcodelogfile+txtdeactive;
                                    if (!File.Exists(ActivatedEmployeeLog))
                                    {
                                        File.Create(ActivatedEmployeeLog).Dispose();
                                    }

                                    // Write the log message to the file.
                                    using (StreamWriter sw = File.AppendText(ActivatedEmployeeLog))
                                    {

                                        string logMessage = $"{DateTime.Now}: Employee Number :- {emp.txtEmpNo} is Deactivated Successfully";
                                        sw.WriteLine(logMessage);
                                    }

                                    successCount++;




                                }
                                else
                                {
                                    string txtExceptionfiles = "ExceptionEmployeeLog.txt";

                                    string ExceptionEmployeeLog = QRcodelogfile+txtExceptionfiles;
                                    if (!File.Exists(ExceptionEmployeeLog))
                                    {
                                        File.Create(ExceptionEmployeeLog).Dispose();
                                    }

                                    // Write the log message to the file.
                                    using (StreamWriter sw = File.AppendText(ExceptionEmployeeLog))
                                    {

                                        string logMessage = $"{DateTime.Now}: Employee Number :- {emp.txtEmpNo} not Exit";
                                        sw.WriteLine(logMessage);
                                    }
                                }
                                reader.Dispose();
                            }


                            sqlCnn.Close();
                        }
                    }



                }

                catch (Exception ex)
                {
                    throw new Exception(ex.Message);
                }

            }
            if (successCount == UpdateList.Count)
            {
                MessageBox.Show("Employee Deactivated Successfully");
            }

        }


    }
}

