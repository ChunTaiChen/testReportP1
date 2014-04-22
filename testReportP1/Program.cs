using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.Presentation;
using ReportTest.framework;
using ReportTest.DAO;
using Aspose.Words;

namespace testReportP1
{
    public class Program
    {
         [FISCA.MainMethod()]
        public static void Main()
        {

            #region Test獎懲
            // 測試按鈕
            K12.Presentation.NLDPanels.Student.ListPaneContexMenu["Test獎懲"].Enable = true;
            K12.Presentation.NLDPanels.Student.ListPaneContexMenu["Test獎懲"].Click += delegate
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0)
                {
                    try
                    {
                        MargeCenter mc = new MargeCenter();
                        // 學生基本資料
                        mc.JoinGroup(new StudentBasicInfo());

                        // 獎懲資料，依學年度學期
                        mc.JoinGroup(new DisciplineDetail(99, 1));

                        // 整理資料，傳入學生系統編號
                        mc.BuildDataTable(K12.Presentation.NLDPanels.Student.SelectedSource);

                        // 讀取樣版
                        Document doc = new Document(new System.IO.MemoryStream(Properties.Resources.自訂報表獎懲測試));
                        DocumentBuilder builder = new DocumentBuilder(doc);

                        // 合併
                        doc.MailMerge.Execute(mc.DataTable);
                        doc.MailMerge.DeleteFields();

                        string filePath = System.Windows.Forms.Application.StartupPath + "\\test獎懲報表.doc";
                        doc.Save(filePath, SaveFormat.Doc);
                        System.Diagnostics.Process.Start(filePath);
                    }
                    catch (Exception ex)
                    {
                        System.Windows.Forms.MessageBox.Show(ex.Message);
                    }
                }
            };
            #endregion

            #region Test獎懲所有學年度學期
             // 測試按鈕
            K12.Presentation.NLDPanels.Student.ListPaneContexMenu["Test獎懲所有學年度學期"].Enable = true;
            K12.Presentation.NLDPanels.Student.ListPaneContexMenu["Test獎懲所有學年度學期"].Click += delegate
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0)
                {
                    try
                    {
                        MargeCenter mc = new MargeCenter();
                        // 學生基本資料
                        mc.JoinGroup(new StudentBasicInfo());

                        // 獎懲資料
                        mc.JoinGroup(new DisciplineDetail());

                        // 整理資料，傳入學生系統編號
                        mc.BuildDataTable(K12.Presentation.NLDPanels.Student.SelectedSource);

                        // 讀取樣版
                        Document doc = new Document(new System.IO.MemoryStream(Properties.Resources.自訂報表獎懲測試_學年度學期));
                        DocumentBuilder builder = new DocumentBuilder(doc);

                        // 合併
                        doc.MailMerge.Execute(mc.DataTable);
                        doc.MailMerge.DeleteFields();

                        string filePath = System.Windows.Forms.Application.StartupPath + "\\test獎懲報表所有學年度學期.doc";
                        doc.Save(filePath, SaveFormat.Doc);
                        System.Diagnostics.Process.Start(filePath);
                    }
                    catch (Exception ex)
                    {
                        System.Windows.Forms.MessageBox.Show(ex.Message);
                    }
                }
            };
            #endregion


            #region Test缺曠
            // 測試按鈕
            K12.Presentation.NLDPanels.Student.ListPaneContexMenu["Test缺曠"].Enable = true;
            K12.Presentation.NLDPanels.Student.ListPaneContexMenu["Test缺曠"].Click += delegate
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0)
                {
                    try
                    {
                        MargeCenter mc = new MargeCenter();
                        // 學生基本資料
                        mc.JoinGroup(new StudentBasicInfo());

                        // 獎懲資料，依學年度學期
                        mc.JoinGroup(new AttendanceDeatil(99, 1));

                        // 整理資料，傳入學生系統編號
                        mc.BuildDataTable(K12.Presentation.NLDPanels.Student.SelectedSource);

                        // 讀取樣版
                        Document doc = new Document(new System.IO.MemoryStream(Properties.Resources.自訂報表缺曠測試));
                        DocumentBuilder builder = new DocumentBuilder(doc);

                        // 合併
                        doc.MailMerge.Execute(mc.DataTable);
                        doc.MailMerge.DeleteFields();

                        string filePath = System.Windows.Forms.Application.StartupPath + "\\testt缺曠報表.doc";
                        doc.Save(filePath, SaveFormat.Doc);
                        System.Diagnostics.Process.Start(filePath);
                    }
                    catch (Exception ex)
                    {
                        System.Windows.Forms.MessageBox.Show(ex.Message);
                    }
                }
            };
            #endregion

            #region Test缺曠所有學年度學期
            // 測試按鈕
            K12.Presentation.NLDPanels.Student.ListPaneContexMenu["Test缺曠所有學年度學期"].Enable = true;
            K12.Presentation.NLDPanels.Student.ListPaneContexMenu["Test缺曠所有學年度學期"].Click += delegate
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0)
                {
                    try
                    {
                        MargeCenter mc = new MargeCenter();
                        // 學生基本資料
                        mc.JoinGroup(new StudentBasicInfo());

                        // 獎懲資料
                        mc.JoinGroup(new AttendanceDeatil());

                        // 整理資料，傳入學生系統編號
                        mc.BuildDataTable(K12.Presentation.NLDPanels.Student.SelectedSource);

                        // 讀取樣版
                        Document doc = new Document(new System.IO.MemoryStream(Properties.Resources.自訂報表缺曠測試_學年度學期));
                        DocumentBuilder builder = new DocumentBuilder(doc);

                        // 合併
                        doc.MailMerge.Execute(mc.DataTable);
                        doc.MailMerge.DeleteFields();

                        string filePath = System.Windows.Forms.Application.StartupPath + "\\testt缺曠報表所有學年度學期.doc";
                        doc.Save(filePath, SaveFormat.Doc);
                        System.Diagnostics.Process.Start(filePath);
                    }
                    catch (Exception ex)
                    {
                        System.Windows.Forms.MessageBox.Show(ex.Message);
                    }
                }
            };
            #endregion
        }
    }
}
