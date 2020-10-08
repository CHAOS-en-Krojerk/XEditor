using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using EXRtbox;

namespace Editor
{
    public partial class frmHyperlink : Form
    {
        private string _selText;
        private string _url;
        frmEditor mMain;
        public frmHyperlink()
        {
            this._selText = "";
            this._url = "";
            InitializeComponent();
        }
        public frmHyperlink(frmEditor frm)
        {
            this._selText = "";
            this._url = "";
            this.mMain = frm;
            InitializeComponent();
        }

        public frmHyperlink(string selText)
        {
            this._selText = selText;
            this._url = "";
            InitializeComponent();
        }

        public frmHyperlink(string selText, frmEditor frm)
        {
            this._selText = selText;
            this._url = "";
            this.mMain = frm;
            InitializeComponent();
        }


        private string GetSelectedText()
        {
            return this._selText;
        }

        private string GetUrl()
        {
            return this._selText;
        }

        private void frmHyperlink_Load(object sender, EventArgs e)
        {
            txtSelText.Text = GetSelectedText();
            if (GetUrl() != "")
                txtLink.Text = GetUrl();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            mMain.rtbDoc.InsertLink(txtSelText.Text,txtLink.Text);
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
