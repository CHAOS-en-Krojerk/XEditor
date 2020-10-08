using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Editor
{
    public partial class frmAddTable : Form
    {
        frmEditor fMain;
        public frmAddTable()
        {
            InitializeComponent();
        }

        public frmAddTable(frmEditor frm)
        {
            InitializeComponent();
            fMain = frm;
        }

        private bool IsDigitsOnly(string str)
        {
            foreach (char c in str)
            {
                if (c < '0' || c > '9')
                    return false;
            }

            return true;
        }

        private static String InsertTableInRichTextBox(int rows, int cols, int width)
        {
            //Create StringBuilder Instance
            StringBuilder sringTableRtf = new StringBuilder();

            //beginning of rich text format
            sringTableRtf.Append(@"{\rtf1 ");

            //Variable for cell width
            int cellWidth;

            //Start row
            sringTableRtf.Append(@"\trowd");

            //Loop to create table string
            for (int i = 0; i<rows; i++)
    {
                sringTableRtf.Append(@"\trowd");

                for (int j = 0; j<cols; j++)
       {
                    //Calculate cell end point for each cell
                    cellWidth = (j + 1) * width;

                    //A cell with width 1000 in each iteration.
                    sringTableRtf.Append(@"\cellx" + cellWidth.ToString());
                }

                //Append the row in StringBuilder
                sringTableRtf.Append(@"\intbl \cell \row");
            }
            sringTableRtf.Append(@"\pard");
            sringTableRtf.Append(@"}");

            return sringTableRtf.ToString();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            try
            {
                if(IsDigitsOnly(txtRows.Text.ToString()) && IsDigitsOnly(txtColumns.Text.ToString()) && IsDigitsOnly(txtWidth.Text.ToString()))
                {
                    int rw = int.Parse(txtRows.Text.ToString());
                    int cls = int.Parse(txtColumns.Text.ToString());
                    int width = int.Parse(txtWidth.Text.ToString());
                    this.fMain.rtbDoc.SelectedRtf = InsertTableInRichTextBox(rw, cls, width);
                    this.Close();
                }
                else
                {
                    MessageBox.Show("Numbers Only Allowed!!", "XWriter", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch
            {
                MessageBox.Show("Unable to insert table!", "XWriter", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
