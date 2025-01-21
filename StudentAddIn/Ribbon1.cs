using DB;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using QueryWindow;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace StudentAddIn
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnListStudent_Click(object sender, RibbonControlEventArgs e)
        {
            Form1 f = new Form1();
            if (f.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                int class_id = f.SelectedClassId;
                string[] titles = f.column;



                Excel.Application xlsApp = Globals.ThisAddIn.Application;
                Excel.Worksheet xlsSheet = xlsApp.ActiveSheet;

                // 清空頁面
                xlsSheet.Cells.ClearContents();


                //: Call SP 放到 excel
                //xlsSheet.Range["A1", "A1"].Value2 = class_id.ToString();
                DB.sp_student_list.Param p = new DB.sp_student_list.Param();
                p.class_id = class_id;
                DB.sp_student_list.Row[] rows = DB.sp_student_list.ExecuteArr(p);

                int k = 1;
                foreach (var r in titles)
                {
                    xlsSheet.Cells[1, k] = r;
                    k = k + 1;
                }

                //xlsSheet.Cells[1, 1].Value2 = "id";
                //xlsSheet.Cells[1, 2].Value2 = "name";
                //xlsSheet.Cells[1, 3].Value2 = "age";

                // Method 1 , no good
                //int i = 2;
                //foreach(var r in rows)
                //{
                //    xlsSheet.Cells[i, 1].Value2 = r.id;
                //    xlsSheet.Cells[i, 2].Value2 = r.name;
                //    xlsSheet.Cells[i, 3].Value2 = r.age;
                //    ++i;
                //}

                // Method 2 
                object[,] data = new object[rows.Length, titles.Length];
                for(int j=0; j<rows.Length; j++)
                {
                    for (int h = 0; h < titles.Length; h++)
                    {
                        PropertyInfo prop = typeof(sp_student_list.Row).GetProperty(titles[h]);
                        if (prop != null)
                        {
                            data[j, h] = prop.GetValue(rows[j], null)?.ToString();
                        }
                    }
                    //data[j, 0] = rows[j].id;
                    //data[j, 1] = rows[j].name;
                    //data[j, 2] = rows[j].age;
                }

                Clipboard.SetDataObject(data);

                xlsSheet.Range["A2"].Resize[rows.Length, titles.Length].set_Value(Missing.Value, data);
            }
        }

        static void CopyRowFormat(Worksheet ws_template, Worksheet ws_report, int fromRowIndex, int rowCount)
        {
            //: Copy from
            ws_template.Range["A4", "J4"].Copy();

            //: Paste to            
            Range destRange = ws_report.Range["A" + fromRowIndex, "J" + fromRowIndex + rowCount];
            destRange.PasteSpecial(XlPasteType.xlPasteFormats);
            destRange.PasteSpecial(XlPasteType.xlPasteValidation);
        }
    }
}
