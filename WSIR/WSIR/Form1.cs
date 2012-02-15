using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Collections;
using System.IO;
using System.Security.Permissions;
[assembly: CLSCompliant(true)]

namespace WSIR
{
    /*
     * Windows Search Index Reader
     * v1.0
     * This utility program lets you read the index created by
     * windows.
     * This does not help to restore your data.
     * index files only record what-file-is-stored-where kind of information
     * 
     */

    public partial class Form1 : Form
    {

        BackgroundWorker bgw;             

        // Connection string for Windows Search        
        const string connectionString = "Provider=Search.CollatorDSO;Extended Properties=\"Application=Windows\"";
        
        public Form1()
        {
            InitializeComponent();
            bgw = new BackgroundWorker();
            bgw.WorkerReportsProgress = true;
            bgw.WorkerSupportsCancellation = true;

            bgw.DoWork += new DoWorkEventHandler(bgw_DoWork);
            bgw.ProgressChanged += new ProgressChangedEventHandler(bgw_ProgressChanged);
            bgw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bgw_RunWorkerCompleted);

            dataGridView1.Columns.Add("Name","Name");
            dataGridView1.Columns.Add("URL","URL");
            dataGridView1.Columns[0].Width = (dataGridView1.Width / 2) - 50;
            dataGridView1.Columns[1].Width = dataGridView1.Width - dataGridView1.Columns[0].Width - 50;

            lblStatus.Text = "Ready";
            
        }

        /*                      
            SELECT System.ItemName FROM SystemIndex
            SELECT System.ItemName FROM SystemIndex WHERE contains(*, 'keyword*') AND System.Kind = 'email'
                GROUP ON System.Kind AGGREGATE Count()
                OVER (SELECT System.Kind, System.ItemName from SystemIndex)
            
            Recursion depth when expanding chapters for GROUP ON queries.
            0 = stop at first chapter, 1 = stop at second chapter
            By default all chapters are expanded.
        */

        //List All Entries
        private void button1_Click(object sender, EventArgs e)
        {
            executeQuery(@"SELECT System.ItemName,System.ItemUrl From SystemIndex");
        }

        //list filtered entries
        private void btnFilter_Click(object sender, EventArgs e)
        {           
            executeQuery(@"SELECT System.ItemName,System.ItemUrl From SystemIndex WHERE contains(*,'" + txtFilter.Text + "*')");
        }

        //setup the things and start background worker
        private void executeQuery(String query)
        {
            btnFilter.Enabled = false;
            btnExplorerView.Enabled = false;
            btnListEntries.Enabled = false;
            dataGridView1.Rows.Clear();

            try
            {                
                progressBar.Style = ProgressBarStyle.Marquee;
                bgw.RunWorkerAsync(query);
            }
            catch (Exception ae)
            {
                Console.WriteLine(ae);
                Console.WriteLine();
            }
        }        
              
        // create connection-fetch records-send back to UI thread in batch of 100 records
        void bgw_DoWork(object sender, DoWorkEventArgs e)
        {
            int totalCount = 0;
            String query = (String)e.Argument;

            OleDbDataReader myDataReader = null;
            OleDbConnection myOleDbConnection = new OleDbConnection(connectionString);
            OleDbCommand myOleDbCommand = new OleDbCommand(query, myOleDbConnection);

            try
            {
                System.Threading.Thread.Sleep(100);
                
                myOleDbConnection.Open();
                myDataReader = myOleDbCommand.ExecuteReader();
                if (!myDataReader.HasRows)
                {                    
                    return;
                }
                
                //batch counter
                int count = 0;
                
                //list that stores batch of records
                //Record is our own custom class
                List<Record> Records  = new List<Record>();

                //temp string needed to alter Url
                //actual url returned from Search index starts with file:
                String tempUrl;

                while (myDataReader.Read())
                {
                    if (count == 100)
                    {
                        //send batch back to UI thread
                        bgw.ReportProgress(totalCount, Records);
                        Records = new List<Record>();
                        count = 0;
                        
                        //give UI thread little time to process sent batch
                        System.Threading.Thread.Sleep(400);
                    }

                    tempUrl = myDataReader.GetString(1);
                    
                    //url may start as iehistory: or without any suffix
                    //alter the path only if its a file path
                    if(tempUrl.StartsWith("file")) 
                        tempUrl = tempUrl.Substring(5);                    

                    Records.Add(new Record(myDataReader.GetString(0), tempUrl));
                    count++;
                    totalCount++;
                }

                //send batch of last remaining records
                if(Records.Count > 0)
                    bgw.ReportProgress(totalCount, Records);
                                
            }
            catch (Exception ex)
            {
                return;
            }
            finally
            {
                //close everything
                if (myDataReader != null)
                {
                    myDataReader.Close();
                    myDataReader.Dispose();
                }
                // Close the connection
                if (myOleDbConnection.State == System.Data.ConnectionState.Open)
                {
                    myOleDbConnection.Close();
                }
            }
        }

        //process recieved batch of records and add in datagridview
        void bgw_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            List<Record> temp = (List<Record>)e.UserState;
            foreach (Record item in temp)            
                dataGridView1.Rows.Add(new String[]{item.Name,item.URL});

            lblStatus.Text = e.ProgressPercentage + " records found";
        }

        //executes after bgw completed
        void bgw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            progressBar.Style = ProgressBarStyle.Continuous;

            lblStatus.Text = "Done. " + lblStatus.Text;

            btnFilter.Enabled = true;
            btnExplorerView.Enabled = true;
            btnListEntries.Enabled = true;
        }
         
        //generate Explorer view from current DataGridView               
        private void btnExplorerView_Click(object sender, EventArgs e)
        {
            //dataGridView1.Sort(dataGridView1.Columns[1], ListSortDirection.Ascending);
            
            Stack<String> folders = new Stack<string>();
            String path;
            String folderName;            
            String dLetter="";
            TreeNode tempNode;

            treeView1.Nodes.Clear();

            int errorCount = 0;

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                try
                {
                    //get url
                    path = row.Cells[1].Value.ToString();

                    //this is annoying.. had to
                    if (path.Length < 248)
                        //get drive letter
                        dLetter = Path.GetPathRoot(path);
                    else
                        continue;

                    if (!treeView1.Nodes.ContainsKey(dLetter) && !String.IsNullOrEmpty(dLetter))
                    {
                        //treeView does not contain Drive letter
                        //so add it
                        tempNode = treeView1.Nodes.Add(dLetter, dLetter);
                        //assign a list of string to that node
                        //which will be used to store Filelist
                        tempNode.Tag = new List<String>();
                    }
                    else
                    {
                        //treeview contains drive letter

                        //clear stack
                        folders.Clear();

                        //generate stack
                        //if path is like C:\Users\Omie\Desktop
                        //need to first check if 
                        //C:\ node is present, then
                        //Users node, then
                        //Omie node, then
                        //Desktop node
                        //
                        //stack == first in last out
                        //use Path.GetFileName() to get last directory name and push it on stack
                        //set path to its parent directory for next iteration
                        while (!path.Equals(dLetter))
                        {
                            folders.Push(Path.GetFileName(path));
                            path = Path.GetDirectoryName(path);
                        }

                        //set node to start with
                        //access it using key
                        tempNode = treeView1.Nodes[dLetter];
                        
                        //while stack is not empty
                        while (folders.Count > 0)
                        {
                            //get the folder name from top of stack
                            //pop() also removes it from stack                            
                            folderName = folders.Pop();
                                                       
                            //check if popped value is a filename of foldername
                            if(!Path.HasExtension(folderName))
                            {
                                //assuming that a name without extension IS a folder
                                //it may not be true but there isn't any way to figure out if its a folder
                                //if it is a Folder

                                //check if it exists in treeview
                                if(!tempNode.Nodes.ContainsKey(folderName))
                                {
                                    //it doesnt !
                                    //add it
                                    tempNode.Nodes.Add(folderName, folderName);                                    
                                    tempNode.Tag = new List<String>();

                                    //set tempNode for next interation
                                    tempNode = tempNode.Nodes[folderName]; 
                                }
                                else
                                {
                                    //it does exist
                                    //set tempNode for next interation
                                    tempNode = tempNode.Nodes[folderName]; 
                                }
                                    
                            }
                            else
                            {
                                //if its a file

                                //if current node already has a List object assigned to its Tag
                                if (tempNode.Tag != null)
                                {
                                    //add file entry
                                    ((List<String>)tempNode.Tag).Add(folderName);
                                }
                                else
                                {
                                    //create list object
                                    //and add file entry
                                    tempNode.Tag = new List<String>();
                                    ((List<String>)tempNode.Tag).Add(folderName);
                                }
                            }                                
                            
                        }//end while
                        
                    }

                }//try
                catch (Exception ae)
                {
                    errorCount++;
                    continue;                 
                }

            } //for each dataGridView Row
            lblStatus.Text = lblStatus.Text + " " + errorCount + " Non-file-path entries";
            
        }//end generate explorer view

        //show child items in ListView when a node is clicked in TreeView
        private void treeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            try
            {
                listView1.Items.Clear();
                
                //load directories
                //child directories == child nodes
                ListViewItem temp;
                foreach (TreeNode node in e.Node.Nodes)
                {
                    temp=listView1.Items.Add(node.Text);

                    //set to use folder icon
                    temp.ImageIndex = 0;
                }

                //load files
                //file list is saved in Tag property of every node
                List<String> ex = (List<String>)e.Node.Tag;
                foreach (String item in ex)
                {
                    temp=listView1.Items.Add(item);

                    //set to use file icon
                    temp.ImageIndex = 1;
                }
            }
            catch (Exception ae)
            {
                lblStatus.Text = "Some Error Occured";
            }
        }

        //adjust DataGridViewColumn size when form is resized
        private void Form1_Resize(object sender, EventArgs e)
        {
            dataGridView1.Columns[0].Width = (dataGridView1.Width / 2) - 50;
            dataGridView1.Columns[1].Width = dataGridView1.Width - dataGridView1.Columns[0].Width - 50;
        }

        //Set ListView view modes
        private void largeIconsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            listView1.View = View.LargeIcon;
        }
        private void smallIconsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            listView1.View = View.SmallIcon;
        }
        private void listToolStripMenuItem_Click(object sender, EventArgs e)
        {
            listView1.View = View.List;
        }
        private void tileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            listView1.View = View.Tile;
        }
        //End ListView view mode functions

        //show about application form
        private void btnAbout_Click(object sender, EventArgs e)
        {
            using(AboutForm about = new AboutForm())
            {
                about.ShowDialog();
            }
            
        }

        //hope it was good :-)

    }
}
