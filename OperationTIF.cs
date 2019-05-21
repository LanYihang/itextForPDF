using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CPF.Model.Bean;
using CPF.Model.Util;
using System.IO;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System.Drawing.Imaging;

namespace CPF.Control
{
    /// <summary>
    /// 将TIF文件转化成PDF文件的系列操作
    /// </summary>
    public class OperationTIF
    {
        private RightNoticeBean right_notice_bean;//通知书对象
        private string tif_parent_path = "";//TIF图片所在父路径
        private string tif_pdf_path = "";//将TIF转化成PDF后的物理路径
        private string tif_grant_parent_path = "";//TIF图片所在的祖路径
        public OperationTIF() { }

        /// <summary>
        /// 实例化类，期间获得TIF文件路径并将文件转化成PDF
        /// </summary>
        /// <param name="right_notice_bean">通知书对象</param>
        public OperationTIF(RightNoticeBean right_notice_bean) 
        {
            this.right_notice_bean = right_notice_bean; 
            tif_parent_path=Sysinfo.CPC_Software_Root_Directory() + string.Format(@"\cases\notices\{0}\{0}\{0}",right_notice_bean.Right_notice_no);
            tif_grant_parent_path = Sysinfo.CPC_Software_Root_Directory() + string.Format(@"\cases\notices\{0}\{0}", right_notice_bean.Right_notice_no);//需要转换的TIF文件的上上级目录，用以检索附件目录下的文件
            tif_pdf_path = Sysinfo.UPLOAD_FILE_PATH + right_notice_bean.Right_notice_no;
            if (!Directory.Exists(tif_pdf_path)) Directory.CreateDirectory(tif_pdf_path);
        }

        /// <summary>
        /// 加载包含TIF图片的路径集合，即在通知书目录下包含所有的tif图片路径
        /// </summary>
        /// <returns>返回TIF与PDF对应的路径集合</returns>
        public Dictionary<string, string> LoadTifPathInDictionary()
        {
            try
            {
                Dictionary<string, string> tif_pdf_directory = new Dictionary<string, string>();
                DirectoryInfo[] all_tif_directory = new DirectoryInfo(tif_grant_parent_path).GetDirectories();//获得该目录下的所有文件夹，子类文件夹中的文件包含正常通知书与附件
                List<FileInfo> tiff_list = new List<FileInfo>();
                tiff_list.AddRange(new DirectoryInfo(tif_parent_path).GetFiles());//获得正常文件夹如GAxxxxx下的所有文件
                foreach (DirectoryInfo tif_directory in all_tif_directory)
                {
                    if (tif_directory.Name != right_notice_bean.Right_notice_no)
                        tiff_list.AddRange(tif_directory.GetFiles());
                }
                //获得附件文件和详细文件
                for (int i = 0; i < tiff_list.Count; i++)
                {
                    string tif_name = tiff_list[i].Name;
                    string pdf_name = "";
                    if (i > new DirectoryInfo(tif_parent_path).GetFiles().Length - 1)
                    {
                        string c = "00000" + (i + 1);
                        pdf_name = c.Substring(c.Length - 6) + ".pdf";
                    }
                    else
                        pdf_name = tif_name.Replace(".tif", ".pdf");
                    string pdf_path = Sysinfo.UPLOAD_FILE_PATH + right_notice_bean.Right_notice_no + @"\" + pdf_name;//保存PDF文件的路径
                    tif_pdf_directory.Add(tiff_list[i].FullName, pdf_path);
                }
                return tif_pdf_directory;
            }catch(Exception e)
            {
                throw new Exception("LoadTifPathInDictionary异常\r\n" + e.Message);
            }
        }

        /// <summary>
        /// 将TIF图片转化为PDF文件
        /// </summary>
        /// <param name="tif_path">tif路径</param>
        /// <param name="pdf_path">pdf路径</param>
        public void TIF_PDF(string tif_path,string pdf_path) 
        {
            Document pdfdoc =null;
            try
            {
                System.Drawing.Bitmap bm = new System.Drawing.Bitmap(tif_path);
                int total = bm.GetFrameCount(FrameDimension.Page);
                pdfdoc = new Document(PageSize.A4, 50, 50, 50, 50);
                PdfWriter writer = PdfWriter.GetInstance(pdfdoc, new FileStream(pdf_path, FileMode.Create));
                pdfdoc.Open();
                PdfContentByte cb = writer.DirectContent;
                for (int k = 0; k < total; k++)
                {
                    bm.SelectActiveFrame(FrameDimension.Page, k);
                    Image img = Image.GetInstance(bm, null, true);
                    img.ScalePercent(72f / 300f * 100);
                    img.SetAbsolutePosition(0, 0);
                    cb.AddImage(img);
                    pdfdoc.NewPage();
                }
            }
            catch (Exception e)
            {
                throw new Exception("TIF_PDF异常\r\n"+e.Message);
            }
            finally 
            {
                if (pdfdoc != null) pdfdoc.Close();
            }
        }

        /// <summary>
        /// 将生成的所有子PDF文件合并为一个PDF文件
        /// </summary>
        /// <param name="tif_pdf_directory">存储pdf/tif文件路径的字典</param>
        public void All_TIF_PDF(Dictionary<string,string> tif_pdf_directory) 
        {
            Document document=null;
            string new_pdf_path = tif_pdf_path + "\\" + right_notice_bean.Right_notice_file_name + ".pdf";//合并后的PDF路径
            try
            {
                string first_pdf_path = "";
                foreach (string value in tif_pdf_directory.Values)
                {
                    first_pdf_path = value;
                    break;
                }
                document = new Document(new PdfReader(first_pdf_path).GetPageSize(1));
                PdfCopy copy = new PdfCopy(document, new FileStream(new_pdf_path, FileMode.Create));
                document.Open();
                foreach (string value in tif_pdf_directory.Values)
                {
                    PdfReader reader = new PdfReader(value);
                    int n = reader.NumberOfPages;
                    for (int j = 1; j <= n; j++)
                    {
                        document.NewPage();
                        PdfImportedPage page = copy.GetImportedPage(reader, j);
                        copy.AddPage(page);
                    }
                }
            }
            catch (Exception e)
            {
                throw new Exception("All_TIF_PDF异常\r\n" + e.Message);
            }
            finally
            {
                if (document != null) document.Close();
                //将生成的子PDF文件删除
                if (tif_pdf_directory.Count > 0)
                    foreach (string value in tif_pdf_directory.Values)
                        new FileInfo(value).Delete();
            }
        }

    }
}
