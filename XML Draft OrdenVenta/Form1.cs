using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace XML_Draft_OrdenVenta
{
    public partial class Form1 : Form
    {
        private SAPbouiCOM.Application SBO_Application;
        private SAPbobsCOM.Company oCompany;
        int errCode;
        string errMsg;

        public Form1()
        {
            InitializeComponent();
            comboBox1.SelectedIndex = 0;
            DIAPIconexion();
        }

        private void SetApplication()
        {
            SAPbouiCOM.SboGuiApi sboGuiApi;
            string sConnectionString;

            sboGuiApi = new SAPbouiCOM.SboGuiApi();

            // by following the steps specified above, the following
            // statment should be suficient for either development or run mode
            sConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056"; 

            // connect to a running SBO Application
            sboGuiApi.Connect(sConnectionString);

            // get an initialized application object
            SBO_Application = sboGuiApi.GetApplication(-1);
        }

        public void DIAPIconexion()
        {
            try
            {
                SetApplication();

                oCompany = new SAPbobsCOM.Company();

                oCompany = (SAPbobsCOM.Company)SBO_Application.Company.GetDICompany();

                label1.Text = oCompany.CompanyName;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        public void saveDocument()
        {
            string path = "";
            string table = "";
            SAPbobsCOM.Documents oDoc = null;

            switch (comboBox1.SelectedIndex)
            {
                case 0:
                    oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                    table = "ODRF";
                    break;
                case 1:
                    oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                    table = "ORDR";
                    break;
                case 2:
                    oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
                    table = "OINV";
                    break;
                case 3:
                    oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes);
                    table = "ORIN";
                    break;
                case 4:
                    oDoc = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oReturns);
                    table = "ORDN";
                    break;
                default:
                    break;
            };

            try
            {
                SAPbobsCOM.Recordset oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                // query to get the docEntry
                oRs.DoQuery("select DocEntry from " + table + " where DocNum = " + textBox1.Text);
                var id = oRs.Fields.Item(0).Value;
                oDoc.GetByKey(Convert.ToInt32(id));
                oCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_AllNodes;
                oCompany.XMLAsString = false;
                FileDialog dlg = new SaveFileDialog();
                dlg.Filter = "xml files (*.xml)|*.xml";
                DialogResult result = dlg.ShowDialog();

                if (result == DialogResult.OK)
                {
                    path = dlg.FileName;
                    //save the file
                    oDoc.SaveXML(path);
                    textBox1.Clear();
                    MessageBox.Show("The document was saved successfully");
                }
            }
            catch (Exception ex)
            { MessageBox.Show(ex.Message); }
        }

        public bool manageError(int i, bool showError)
        {
            // if errCode is different than 0 an error could occur
            if (i != 0)
            {
                oCompany.GetLastError(out errCode, out errMsg);
                if (showError)
                    MessageBox.Show(errCode + " - " + errMsg);

                return false;
            }
            return true;
        }

        public void loadDocument()
        {
            SAPbobsCOM.Documents oDraft = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
            try
            {
                string path = "";
                FileDialog dlg = new OpenFileDialog();
                dlg.Filter = "xml files (*.xml)|*.xml";
                DialogResult result = dlg.ShowDialog();

                if (result == DialogResult.OK)
                {
                    path = dlg.FileName;
                }

                oDraft = (SAPbobsCOM.Documents)oCompany.GetBusinessObjectFromXML(path, 0);
                string i = oDraft.GetType().ToString();
                if (manageError(oDraft.Add(), true))
                {
                    MessageBox.Show("Document inserted correctly");
                }
            }
            catch (Exception ex)
            { MessageBox.Show(ex.Message); }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DIAPIconexion();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            saveDocument();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            loadDocument();
        }
    }
}
