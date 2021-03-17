using System;
using System.Data;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.IO.Ports;

namespace InsertUpdateDeleteDemo
{
    public partial class frmMain : Form
    {
        //OleDbConnection con= new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/FutApsol/DatabaseApplication/FactConnect1.mdb;");
        //OleDbCommand cmd;
        //OleDbDataAdapter adapt;

        static SerialPort myPort;

        private const string conString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/FutApsol/DatabaseApplication/FactConnect1.mdb;";
        readonly OleDbConnection con = new OleDbConnection(conString);
        OleDbCommand cmd;
        OleDbDataAdapter adapt;

        //ID variable used in Updating and Deleting Record
        int ID = 0;
        public frmMain()
        {
            InitializeComponent();
           // con.Open();
            DisplayData();

            try {
                myPort = new SerialPort("COM4", 9600);
                myPort.Open();

                myPort.DataReceived += new SerialDataReceivedEventHandler(myPort_DataReceived);
                myPort.ReceivedBytesThreshold = 1;
            }
            catch(Exception ex) { 
                MessageBox.Show(ex.Message);
                myPort.Close();
            }

            
        }

        public void myPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try {
                string newPortData = myPort.ReadExisting();

                string[] portDataArr = newPortData.Split(',');
                int portDateArrLenght = portDataArr.Length;

                classes.PortDataDetails portDetailsObj = new classes.PortDataDetails();

                if (portDateArrLenght >= 17)
                {
                    portDetailsObj.partNumber = portDataArr[0];
                    portDetailsObj.revisionNumber = portDataArr[1];
                    portDetailsObj.dateTime = portDataArr[2] + ' ' + portDataArr[3];
                    portDetailsObj.testResult = portDataArr[4];
                    portDetailsObj.testingMode = portDataArr[5];
                    portDetailsObj.passCount = portDataArr[6];
                    portDetailsObj.failCount = portDataArr[7];
                    portDetailsObj.failTypePoints = portDataArr[8];
                    portDetailsObj.cutterCount = portDataArr[9];
                    portDetailsObj.labelSerialNumber = portDataArr[10];
                    portDetailsObj.printedBarCode = portDataArr[11];
                    portDetailsObj.operatorCode = portDataArr[12];
                    portDetailsObj.shift = portDataArr[13];
                    portDetailsObj.lotCount = portDataArr[14];
                    portDetailsObj.lotQuantity = portDataArr[15];
                    portDetailsObj.customField = portDataArr[16];

                    //MessageBox.Show(portDetailsObj.ToString());

                    const string insertQuerystr = "insert into Face_Connect(" +
                "part_number, " +
                "date_time, " +
                "test_result, " +
                "testing_mode, " +
                "pass_count, " +
                "fail_count, " +
                "fail_type_points, " +
                "cutter_count, " +
                "label_serial_number, " +
                "printed_barcode, " +
                "operator_code, " +
                "shift, " +
                "lot_count, " +
                "lot_quantity, " +
                "custom_field)" +
                "values(" +
                "@partNumber, " +
                "@dateTime, " +
                "@testResult, " +
                "@testingMode, " +
                "@passCount, " +
                "@failCount, " +
                "@failTypePoints, " +
                "@cutterCount, " +
                "@labelSerialNumber, " +
                "@printedBarCode, " +
                "@operatorCode, " +
                "@shift, " +
                "@lotCount, " +
                "@lotQuantity, " +
                "@customField)";

                    cmd = new OleDbCommand(insertQuerystr, con);

                    cmd.Parameters.AddWithValue("@partNumber", portDetailsObj.partNumber);
                    cmd.Parameters.AddWithValue("@dateTime", portDetailsObj.dateTime);
                    cmd.Parameters.AddWithValue("@testResult", portDetailsObj.testResult);
                    cmd.Parameters.AddWithValue("@testingMode", portDetailsObj.testingMode);
                    cmd.Parameters.AddWithValue("@passCount", portDetailsObj.passCount);
                    cmd.Parameters.AddWithValue("@failCount", portDetailsObj.failCount);
                    cmd.Parameters.AddWithValue("@failTypePoints", portDetailsObj.failTypePoints);
                    cmd.Parameters.AddWithValue("@cutterCount", portDetailsObj.cutterCount);
                    cmd.Parameters.AddWithValue("@labelSerialNumber", portDetailsObj.labelSerialNumber);
                    cmd.Parameters.AddWithValue("@printedBarCode", portDetailsObj.printedBarCode);
                    cmd.Parameters.AddWithValue("@operatorCode", portDetailsObj.operatorCode);
                    cmd.Parameters.AddWithValue("@shift", portDetailsObj.shift);
                    cmd.Parameters.AddWithValue("@lotCount", portDetailsObj.lotCount);
                    cmd.Parameters.AddWithValue("@lotQuantity", portDetailsObj.lotQuantity);
                    cmd.Parameters.AddWithValue("@customField", portDetailsObj.customField);

                    try
                    {

                        con.Open();
                        if (cmd.ExecuteNonQuery() > 0)
                        {

                            MessageBox.Show("Data inserted successfully");
                        }
                        con.Close();

                        //retrieveData();

                        DisplayData();

                    }
                    catch (Exception ex)
                    {
                        con.Close();
                        MessageBox.Show(ex.Message);
                    }


                }

                DisplayData();
            }
            catch { }
        }
        //Insert Data
        private void btn_Insert_Click(object sender, EventArgs e)
        {
            if (txt_Name.Text != "" && txt_State.Text != "")
            {
                cmd = new OleDbCommand("insert into tbl_Record(Name,State) values(@name,@state)", con);
                con.Open();
                cmd.Parameters.AddWithValue("@name", txt_Name.Text);
                cmd.Parameters.AddWithValue("@state", txt_State.Text);
                cmd.ExecuteNonQuery();
                con.Close();
                MessageBox.Show("Record Inserted Successfully");
                DisplayData();
                ClearData();
            }
            else
            {
                MessageBox.Show("Please Provide Details!");
            }
        }
        //Display Data in DataGridView
        private void DisplayData()
        {
            con.Open();
            DataTable dt=new DataTable();
            adapt=new OleDbDataAdapter("select * from tbl_Record",con);
            adapt.Fill(dt);
            dataGridView1.DataSource = dt;
            con.Close();
        }
        //Clear Data 
        private void ClearData()
        {
            txt_Name.Text = "";
            txt_State.Text = "";
            ID = 0;
        }
        //dataGridView1 RowHeaderMouseClick Event
        private void dataGridView1_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            ID = Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString());
            txt_Name.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
            txt_State.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
        }
        //Update Record
        private void btn_Update_Click(object sender, EventArgs e)
        {
            if (txt_Name.Text != "" && txt_State.Text != "")
            {
                cmd = new OleDbCommand("update tbl_Record set Name=@name,State=@state where ID=@id", con);
                con.Open();
                cmd.Parameters.AddWithValue("@id", ID);
                cmd.Parameters.AddWithValue("@name", txt_Name.Text);
                cmd.Parameters.AddWithValue("@state", txt_State.Text);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Record Updated Successfully");
                con.Close();
                DisplayData();
                ClearData();
            }
            else
            {
                MessageBox.Show("Please Select Record to Update");
            }
        }
        //Delete Record
        private void btn_Delete_Click(object sender, EventArgs e)
        {
            if(ID!=0)
            {
                cmd = new OleDbCommand("delete tbl_Record where ID=@id",con);
                con.Open();
                cmd.Parameters.AddWithValue("@id",ID);
                cmd.ExecuteNonQuery();
                con.Close();
                MessageBox.Show("Record Deleted Successfully!");
                DisplayData();
                ClearData();
            }
            else
            {
                MessageBox.Show("Please Select Record to Delete");
            }
        }
    }
}
