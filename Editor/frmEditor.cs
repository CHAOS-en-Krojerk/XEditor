using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using EXRtbox;

namespace Editor
{
    public partial class frmEditor : Form
    {
        private string currentFile;
        private int checkPrint;
        public frmEditor()
        {
            InitializeComponent();
            currentFile = "";
            this.Text = "XWriter: New Document";
            SetupFontFaceComboBox();
            SetupFontSizeComboBox();
            SetupTimer();
            rtbDoc.DetectUrls = true;
            
        }

        //form loading and closing methods
        //פעולות פתיחת/סגירת טופס
        private void frmEditor_Load(object sender, EventArgs e)
        {
            this.showRegularToolbarToolStripMenuItem.Checked = mainToolbar.Visible;
            this.regularToolStripMenuItem.Checked = mainToolbar.Visible;
            this.showDesignToolbarToolStripMenuItem.Checked = designToolbar.Visible;
            this.designToolStripMenuItem.Checked = designToolbar.Visible;
            this.showStatusBarToolStripMenuItem.Checked = statusStrip1.Visible;
            this.statusBarToolStripMenuItem.Checked = statusStrip1.Visible;
            this.boldButton.Checked = IsBold();
            this.italicButton.Checked = IsItalic();
            this.underlineButton.Checked = IsUnderline();
            this.justLeftButton.Checked = IsJustifyLeft();
            this.justCenterButton.Checked = IsJustifyCenter();
            this.justRightButton.Checked = IsJustifyRight();
            this.rtlToolStripButton.Checked = IsRightToLeft();
            this.ltrToolStripButton.Checked = IsLeftToRight();
        }
        private void frmEditor_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                if (rtbDoc.Modified == true)
                {
                    DialogResult answer;
                    answer = MessageBox.Show("Save current document before exiting?", "Unsaved Document", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (answer == DialogResult.Yes)
                    {
                        saveToolStripMenuItem_Click(this, new EventArgs());
                        rtbDoc.Modified = false;
                        return;
                    }
                    else
                        rtbDoc.Modified = false;
                }
                else
                    rtbDoc.Modified = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        //help methods
        //פעולות עזר
        private void SetupTimer()
        {
            timer2.Interval = 200;
            timer2.Tick += new EventHandler(timer2_Tick);
            timer2.Start();
        }
        private void OpenFile()
        {
            try
            {
                openFileDialog1.Title = "Open File";
                openFileDialog1.DefaultExt = "rtf";
                openFileDialog1.Filter = "Rich Text Files|*.rtf|Text Files|*.txt|HTML Files|*.htm|All Files|*.*";
                openFileDialog1.FilterIndex = 1;
                openFileDialog1.FileName = string.Empty;
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    if (openFileDialog1.FileName == "")
                        return;
                    string strExt;
                    strExt = System.IO.Path.GetExtension(openFileDialog1.FileName);
                    strExt = strExt.ToUpper();
                    if (strExt == ".RTF")
                        rtbDoc.LoadFile(openFileDialog1.FileName, RichTextBoxStreamType.RichText);
                    else
                    {
                        System.IO.StreamReader txtReader;
                        txtReader = new System.IO.StreamReader(openFileDialog1.FileName);
                        rtbDoc.Text = txtReader.ReadToEnd();
                        txtReader.Close();
                        txtReader = null;
                        rtbDoc.SelectionStart = 0;
                        rtbDoc.SelectionLength = 0;
                    }
                    currentFile = openFileDialog1.FileName;
                    rtbDoc.Modified = false;
                    this.Text = "XWriter: " + currentFile.ToString();
                }
                else
                {
                    MessageBox.Show("File opening cancelled by you", "XWriter");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        private void SetupFontFaceComboBox()
        {
            foreach (FontFamily fam in FontFamily.Families)
                fontFaceComboBox.Items.Add(fam.Name);
            fontFaceComboBox.Text = rtbDoc.Font.Name;
        }
        private void SetupFontSizeComboBox()
        {
            for (int x = 1; x <= 75; x++)
                fontSizeComboBox.Items.Add(x);
            fontSizeComboBox.SelectedIndex = (int)rtbDoc.Font.Size-1;
        }
        
        private void printDocument1_BeginPrint(object sender, PrintEventArgs e)
        {
            checkPrint = 0;
        }
        private void printDocument1_PrintPage(object sender, PrintPageEventArgs e)
        {
            checkPrint = rtbDoc.Print(checkPrint, rtbDoc.TextLength, e);
            if (checkPrint < rtbDoc.TextLength)
                e.HasMorePages = true;
            else
                e.HasMorePages = false;
        }
        private bool IsBold()
        {
            return (rtbDoc.SelectionFont.Style == FontStyle.Bold);
        }
        private bool IsItalic()
        {
            return (rtbDoc.SelectionFont.Style == FontStyle.Italic);
        }
        private bool IsUnderline()
        {
            return (rtbDoc.SelectionFont.Style == FontStyle.Underline);
        }
        private bool IsStrikeout()
        {
            return (rtbDoc.SelectionFont.Style == FontStyle.Strikeout);
        }
        private bool IsJustifyLeft()
        {
            return (rtbDoc.SelectionAlignment == HorizontalAlignment.Left);
        }
        private bool IsJustifyCenter()
        {
            return (rtbDoc.SelectionAlignment == HorizontalAlignment.Center);
        }
        private bool IsJustifyRight()
        {
            return (rtbDoc.SelectionAlignment == HorizontalAlignment.Right);
        }
        private bool IsRightToLeft()
        {
            return (rtbDoc.RightToLeft == RightToLeft.Yes);
        }
        private bool IsLeftToRight()
        {
            return (rtbDoc.RightToLeft == RightToLeft.No);
        }
        private void UpdateFontFace()
        {
            fontFaceComboBox.Text = rtbDoc.SelectionFont.Name;
        }
        private void UpdateFontSize()
        {
            fontSizeComboBox.SelectedIndex = (int)rtbDoc.SelectionFont.Size - 1;
        }

        //menu methods
        //פעולות של התפריט
        private void newToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (rtbDoc.Modified == true)
                {
                    DialogResult answer;
                    answer = MessageBox.Show("Save current document before creating new document?", "Unsaved Document", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (answer == DialogResult.No)
                    {
                        currentFile = "";
                        this.Text = "XWriter: New Document";
                        rtbDoc.Modified = false;
                        rtbDoc.Clear();
                        return;
                    }
                    else
                    {
                        saveToolStripMenuItem_Click(this, new EventArgs());
                        rtbDoc.Modified = false;
                        rtbDoc.Clear();
                        currentFile = "";
                        this.Text = "XWriter: New Document";
                        return;
                    }
                }
                else
                {
                    currentFile = "";
                    this.Text = "XWriter: New Document";
                    rtbDoc.Modified = false;
                    rtbDoc.Clear();
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (rtbDoc.Modified == true)
                {
                    DialogResult answer;
                    answer = MessageBox.Show("Save current document before opening document?", "Unsaved Document", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (answer == DialogResult.No)
                    {
                        rtbDoc.Modified = false;
                        OpenFile();
                    }
                    else
                    {
                        saveToolStripMenuItem_Click(this, new EventArgs());
                        OpenFile();
                    }
                }
                else
                    OpenFile();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (currentFile == string.Empty)
                {
                    saveAsToolStripMenuItem_Click(this, e);
                    return;
                }
                else
                {
                    string strExt;
                    strExt = System.IO.Path.GetExtension(currentFile);
                    strExt = strExt.ToUpper();
                    if (strExt == ".RTF")
                        rtbDoc.SaveFile(currentFile);
                    else
                    {
                        System.IO.StreamWriter txtWriter;
                        txtWriter = new System.IO.StreamWriter(currentFile);
                        txtWriter.Write(rtbDoc.Text);
                        txtWriter.Close();
                        txtWriter = null;
                        rtbDoc.SelectionStart = 0;
                        rtbDoc.SelectionLength = 0;
                    }

                    this.Text = "XWriter: " + currentFile.ToString();
                    rtbDoc.Modified = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                saveFileDialog1.Title = "Save File";
                saveFileDialog1.DefaultExt = "rtf";
                saveFileDialog1.Filter = "Rich Text Files|*.rtf|Text Files|*.txt|HTML Files|*.htm|All Files|*.*";
                saveFileDialog1.FilterIndex = 1;
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    if (saveFileDialog1.FileName == "")
                        return;
                    string strExt;
                    strExt = System.IO.Path.GetExtension(saveFileDialog1.FileName);
                    strExt = strExt.ToUpper();
                    if (strExt == ".RTF")
                        rtbDoc.SaveFile(saveFileDialog1.FileName, RichTextBoxStreamType.RichText);
                    else
                    {
                        System.IO.StreamWriter txtWriter;
                        txtWriter = new System.IO.StreamWriter(saveFileDialog1.FileName);
                        txtWriter.Write(rtbDoc.Text);
                        txtWriter.Close();
                        txtWriter = null;
                        rtbDoc.SelectionStart = 0;
                        rtbDoc.SelectionLength = 0;
                    }
                    currentFile = saveFileDialog1.FileName;
                    rtbDoc.Modified = false;
                    this.Text = "XWriter: " + currentFile.ToString();
                    MessageBox.Show(currentFile.ToString() + " saved.", "File Save");
                }
                else
                    MessageBox.Show("File saving cancelled by you", "XWriter");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        private void printPreviewToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                printPreviewDialog1.Document = printDocument1;
                printPreviewDialog1.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }

        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (rtbDoc.Modified == true)
                {
                    DialogResult answer;
                    answer = MessageBox.Show("Save current document before exiting?", "Unsaved Document", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (answer == DialogResult.Yes)
                    {
                        saveToolStripMenuItem_Click(this, new EventArgs());
                        return;
                    }
                    else
                    {
                        rtbDoc.Modified = false;
                        Application.Exit();
                    }
                }
                else
                {
                    rtbDoc.Modified = false;
                    Application.Exit();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        private void undoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (rtbDoc.CanUndo)
                    rtbDoc.Undo();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        private void redoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (rtbDoc.CanRedo)
                {
                    rtbDoc.Redo();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        private void selectAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                rtbDoc.SelectAll();
            }
            catch
            {
                MessageBox.Show("Unable to select all document content.", "SelectAll", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void cutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                rtbDoc.Cut();
            }
            catch
            {
                MessageBox.Show("Unable to cut document content.", "Cut", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                rtbDoc.Copy();
            }
            catch (Exception)
            {
                MessageBox.Show("Unable to copy document content.", "RTE - Copy", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void pasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                rtbDoc.Paste();
            }
            catch
            {
                MessageBox.Show("Unable to copy clipboard content to document.", "RTE - Paste", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void fontStyleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (!(rtbDoc.SelectionFont == null))
                    fontDialog1.Font = rtbDoc.SelectionFont;
                else
                    fontDialog1.Font = null;
                fontDialog1.ShowApply = true;
                if (fontDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    rtbDoc.SelectionFont = fontDialog1.Font;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        private void boldToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (!(rtbDoc.SelectionFont == null))
                {
                    Font currentFont = rtbDoc.SelectionFont;
                    FontStyle newFontStyle;
                    newFontStyle = rtbDoc.SelectionFont.Style ^ FontStyle.Bold;
                    rtbDoc.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, newFontStyle);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        private void italicToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (!(rtbDoc.SelectionFont == null))
                {
                    Font currentFont = rtbDoc.SelectionFont;
                    FontStyle newFontStyle;
                    newFontStyle = rtbDoc.SelectionFont.Style ^ FontStyle.Italic;
                    rtbDoc.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, newFontStyle);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        private void underlineToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (!(rtbDoc.SelectionFont == null))
                {
                    Font currentFont = rtbDoc.SelectionFont;
                    FontStyle newFontStyle;
                    newFontStyle = rtbDoc.SelectionFont.Style ^ FontStyle.Underline;
                    rtbDoc.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, newFontStyle);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        private void normalToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (!(rtbDoc.SelectionFont == null))
                {
                    Font currentFont = rtbDoc.SelectionFont;
                    FontStyle newFontStyle = FontStyle.Regular;
                    rtbDoc.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, newFontStyle);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        private void findToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                frmFind f = new frmFind(this);
                f.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }

        }

        private void findReplaceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                frmReplace f = new frmReplace(this);
                f.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }

        }

        private void printToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                printDialog1.Document = printDocument1;
                if (printDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    printDocument1.Print();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        private void showRegularToolbarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            mainToolbar.Visible = !mainToolbar.Visible;
            this.regularToolStripMenuItem.Checked = mainToolbar.Visible;
        }

        private void showDesignToolbarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            designToolbar.Visible = !designToolbar.Visible;
            this.designToolStripMenuItem.Checked = designToolbar.Visible;
        }

        private void showStatusBarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            statusStrip1.Visible = !statusStrip1.Visible;
            this.statusBarToolStripMenuItem.Checked = statusStrip1.Visible;
        }

        private void contentsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("For any help, contant Nihilanth at FXP.co.il", "Help", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutBox abt = new AboutBox(this);
            abt.ShowDialog();
        }
        //toolbar methods
        //פעולות של הסרגלי כלים
        private void newToolStripButton_Click(object sender, EventArgs e)
        {
            newToolStripMenuItem_Click(this, e);
        }

        private void openToolStripButton_Click(object sender, EventArgs e)
        {
            openToolStripMenuItem_Click(this, e);
        }

        private void saveToolStripButton_Click(object sender, EventArgs e)
        {
            saveToolStripMenuItem_Click(this, e);
        }

        private void printToolStripButton_Click(object sender, EventArgs e)
        {
            printToolStripMenuItem_Click(this, e);
        }

        private void cutToolStripButton_Click(object sender, EventArgs e)
        {
            cutToolStripMenuItem_Click(this, e);
        }

        private void copyToolStripButton_Click(object sender, EventArgs e)
        {
            copyToolStripMenuItem_Click(this, e);
        }

        private void pasteToolStripButton_Click(object sender, EventArgs e)
        {
            pasteToolStripMenuItem_Click(this, e);
        }
        private void findButton_Click(object sender, EventArgs e)
        {
            findToolStripMenuItem_Click(this, e);
        }

        private void linkButton_Click(object sender, EventArgs e)
        {
            try
            {
                if(rtbDoc.SelectedText!=null)
                {
                    frmHyperlink fhp = new frmHyperlink(rtbDoc.SelectedText, this);
                    fhp.ShowDialog();
                }
                 else
                {
                    frmHyperlink fhp = new frmHyperlink(this);
                    fhp.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
            //MessageBox.Show((new NotImplementedException()).Message.ToString(), "Error");
        }

        private void imageButton_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "Insert Image";
            openFileDialog1.DefaultExt = "rtf";
            openFileDialog1.Filter = "Bitmap Files|*.bmp|JPEG Files|*.jpg|GIF Files|*.gif|PNG Files|*.png";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.ShowDialog();
            if (openFileDialog1.FileName == "")
                return;
            try
            {
                string strImagePath = openFileDialog1.FileName;
                Image img = Image.FromFile(strImagePath);
                Clipboard.SetDataObject(img);
                DataFormats.Format df = DataFormats.GetFormat(DataFormats.Bitmap);
                if (this.rtbDoc.CanPaste(df))
                    this.rtbDoc.Paste(df);
            }
            catch
            {
                MessageBox.Show("Unable to insert image format selected.", "XWriter", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void addTabletoolStripButton_Click(object sender, EventArgs e)
        {
            frmAddTable fr = new frmAddTable(this);
            fr.ShowDialog();
        }

        private void helpToolStripButton_Click(object sender, EventArgs e)
        {
            contentsToolStripMenuItem_Click(this, e);
        }

        private void fontFaceComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                Font newFont = new Font(fontFaceComboBox.SelectedItem.ToString(), rtbDoc.SelectionFont.Size, rtbDoc.SelectionFont.Style);
                rtbDoc.SelectionFont = newFont;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        private void fontSizeComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                float newSize;
                float.TryParse(fontSizeComboBox.SelectedItem.ToString(), out newSize);
                Font newFont = new Font(rtbDoc.SelectionFont.Name, newSize, rtbDoc.SelectionFont.Style);
                rtbDoc.SelectionFont = newFont;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        private void boldButton_Click(object sender, EventArgs e)
        {
            boldToolStripMenuItem_Click(this, e);
        }

        private void italicButton_Click(object sender, EventArgs e)
        {
            italicToolStripMenuItem_Click(this, e);
        }

        private void underlineButton_Click(object sender, EventArgs e)
        {
            underlineToolStripMenuItem_Click(this, e);
        }

        private void strikeoutButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (!(rtbDoc.SelectionFont == null))
                {
                    Font currentFont = rtbDoc.SelectionFont;
                    FontStyle newFontStyle;
                    newFontStyle = rtbDoc.SelectionFont.Style ^ FontStyle.Strikeout;
                    rtbDoc.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, newFontStyle);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        private void upperCaseButton_Click(object sender, EventArgs e)
        {
            try
            {
                rtbDoc.SelectedText = rtbDoc.SelectedText.ToUpper();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        private void lowerCaseButton_Click(object sender, EventArgs e)
        {
            try
            {
                rtbDoc.SelectedText = rtbDoc.SelectedText.ToLower();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        private void biggerButton_Click(object sender, EventArgs e)
        {
            try
            {
                float newFontSize = rtbDoc.SelectionFont.SizeInPoints + 2;
                Font newSize = new Font(rtbDoc.SelectionFont.Name, newFontSize, rtbDoc.SelectionFont.Style);
                rtbDoc.SelectionFont = newSize;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        private void smallerButton_Click(object sender, EventArgs e)
        {
            try
            {
                float newFontSize = rtbDoc.SelectionFont.SizeInPoints - 2;
                Font newSize = new Font(rtbDoc.SelectionFont.Name, newFontSize, rtbDoc.SelectionFont.Style);
                rtbDoc.SelectionFont = newSize;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        private void foreColorButon_Click(object sender, EventArgs e)
        {
            try
            {
                colorDialog1.Color = rtbDoc.ForeColor;
                if (colorDialog1.ShowDialog() == DialogResult.OK)
                    rtbDoc.SelectionColor = colorDialog1.Color;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        private void highlightColorButton_Click(object sender, EventArgs e)
        {
            try
            {
                colorDialog1.Color = rtbDoc.SelectionBackColor;
                if (colorDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    rtbDoc.SelectionBackColor = colorDialog1.Color;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        private void btnFontStyle_Click(object sender, EventArgs e)
        {
            fontStyleToolStripMenuItem_Click(this, e);
        }

        private void justLeftButton_Click(object sender, EventArgs e)
        {
            try
            {
                rtbDoc.SelectionAlignment = HorizontalAlignment.Left;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        private void justCenterButton_Click(object sender, EventArgs e)
        {
            try
            {
                rtbDoc.SelectionAlignment = HorizontalAlignment.Center;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        private void justRightButton_Click(object sender, EventArgs e)
        {
            try
            {
                rtbDoc.SelectionAlignment = HorizontalAlignment.Right;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        private void justFullButton_Click(object sender, EventArgs e)
        {
            try
            {
                rtbDoc.SelectionAlignment = HorizontalAlignment.Right;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        private void rtlToolStripButton_Click(object sender, EventArgs e)
        {
            try
            {
                rtbDoc.RightToLeft = RightToLeft.Yes;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");

            }
        }

        private void ltrToolStripButton_Click(object sender, EventArgs e)
        {
            try
            {
                rtbDoc.RightToLeft = RightToLeft.No;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");

            }

        }

        private void btnIndent_Click(object sender, EventArgs e)
        {
            try
            {
                rtbDoc.SelectionIndent += 5;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        private void btnOutdent_Click(object sender, EventArgs e)
        {
            try
            {
                if (rtbDoc.SelectionIndent != 0)
                    rtbDoc.SelectionIndent -= 5;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        private void ulButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (rtbDoc.SelectionBullet == false)
                {
                    rtbDoc.BulletIndent = 10;
                    rtbDoc.SelectionBullet = true;
                }
                else
                    rtbDoc.SelectionBullet = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        private void printButton_Click(object sender, EventArgs e)
        {
            printToolStripButton_Click(this, e);
        }

        private void printPreviewButton_Click(object sender, EventArgs e)
        {
            printPreviewToolStripMenuItem_Click(this, e);
        }

        private void pageSetupToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                pageSetupDialog1.Document = printDocument1;
                pageSetupDialog1.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        //Timer
        //טיימר
        private void timer1_Tick(object sender, EventArgs e)
        //טיימר ראשון מעדכן את הסטטוס
        {
            if (rtbDoc.Text.Length > 0)
            {
                wordsCountLabel.Text = "Characters: " + rtbDoc.Text.Length.ToString();
                linesCountLabel.Text = "Lines: " + rtbDoc.Lines.Length.ToString();
            }
        }
        private void timer2_Tick(object sender, EventArgs e)
        //טיימר שני מעדכן את רשימת הפונטים והגדלים לפי מיקום הסמן
        {
            this.showRegularToolbarToolStripMenuItem.Checked = mainToolbar.Visible;
            this.regularToolStripMenuItem.Checked = mainToolbar.Visible;
            this.showDesignToolbarToolStripMenuItem.Checked = designToolbar.Visible;
            this.designToolStripMenuItem.Checked = designToolbar.Visible;
            this.showStatusBarToolStripMenuItem.Checked = statusStrip1.Visible;
            this.statusBarToolStripMenuItem.Checked = statusStrip1.Visible;
            this.boldButton.Checked = IsBold();
            this.italicButton.Checked = IsItalic();
            this.underlineButton.Checked = IsUnderline();
            this.justLeftButton.Checked = IsJustifyLeft();
            this.justCenterButton.Checked = IsJustifyCenter();
            this.justRightButton.Checked = IsJustifyRight();
            this.rtlToolStripButton.Checked = IsRightToLeft();
            this.ltrToolStripButton.Checked = IsLeftToRight();
            if ((fontFaceComboBox.Focused==false) && (fontSizeComboBox.Focused==false))
            {
                UpdateFontFace();
                UpdateFontSize();
            }

        }

        //context menu methods
        //פעולות של תפריטים קופצים
        private void regularToolStripMenuItem_Click(object sender, EventArgs e)
        {
            showRegularToolbarToolStripMenuItem_Click(this, e);
        }

        private void designToolStripMenuItem_Click(object sender, EventArgs e)
        {
            showDesignToolbarToolStripMenuItem_Click(this, e);
        }

        private void statusBarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            showStatusBarToolStripMenuItem_Click(this, e);
        }

        private void cutToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            cutToolStripMenuItem_Click(this, e);
        }

        private void copyToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            copyToolStripMenuItem_Click(this, e);
        }

        private void pasteToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            pasteToolStripMenuItem_Click(this, e);
        }

        private void selectAllToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            selectAllToolStripMenuItem_Click(this, e);
        }

        private void boldToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            boldToolStripMenuItem_Click(this, e);
        }

        private void italicToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            italicToolStripMenuItem_Click(this, e);
        }

        private void underlineToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            underlineToolStripMenuItem_Click(this, e);
        }

        private void strikeoutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            strikeoutButton_Click(this, e);
        }

        private void fontToolStripMenuItem_Click(object sender, EventArgs e)
        {
            fontStyleToolStripMenuItem_Click(this, e);
        }

        private void colorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreColorButon_Click(this, e);
        }

        private void highlightToolStripMenuItem_Click(object sender, EventArgs e)
        {
            highlightColorButton_Click(this, e);
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            example frm1 = new example();
            frm1.Show();
        }


    }
}
