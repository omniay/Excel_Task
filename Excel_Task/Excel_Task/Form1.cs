using Excel_Task;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Excel_Task
{
    public partial class Form1 : Form
    {
        private ExcelPackage excel;
        List<string[]>header;
        public Form1()
        {

            InitializeComponent();
            
        }

        private void openfile_Click(object sender, EventArgs e)
        {
            string filepath="Excel.xlsx";
           
            try
            {
               
                  
                  
                    FileInfo excelFile = new FileInfo(filepath);
                    excel.SaveAs(excelFile);
                    MessageBox.Show("Done");
                
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        
        private void Form1_Load(object sender, EventArgs e)
        {
             excel = new ExcelPackage();

            ExcelData data = new ExcelData();
            ExcelStyles.set_worksheet(excel, "Empolyee");
            ExcelWorksheet sheet =ExcelStyles.get_worksheet(excel,"Empolyee");
            header = new List<string[]>(){
                 new string [] { "EmployeeID", "Empolyee First Name", "Last Name", "Floor", "Bouns" },
             };

            ExcelStyles.set_style(sheet,header);
            ExcelStyles.set_color(sheet.Cells["f7"], Color.Red);
           ExcelStyles.format_money(sheet.Cells["I8:I12"], "$#,##0.00");
            ExcelStyles.set_borders(sheet.Cells["E8:I10"]);
            ExcelStyles.load_data(sheet.Cells[8, 5], data.CellData);
         

        }
    }
}
