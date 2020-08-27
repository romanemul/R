using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Windows.Forms;

namespace RRD
{
    public class ComponentTools: IComponentTools
    {
        public BindingSource bindingSource = new BindingSource();
        
        public List<string> DataGridViews;
        public DataGridView _ActiveDataGridView;

        public List<string> ListBoxes;
        public ListBox _ActiveListBox;

        public List<string> ListViews;
        public ListView _ActiveListView;
        private Form activeform;

        public Form Activeform 
        { 
            get => activeform; 
            set => activeform = value; 
        }
        
        public ComponentTools(Form form)
        {
            Activeform = form;            
            InitializeListBoxes();
        }

        private void InitializeListBoxes() 
        {
            if(Activeform.Controls.OfType<ListBox>().Select(element => element.Name).ToList().Count > 0) 
            {
                ListBoxes.AddRange(Activeform.Controls.OfType<ListBox>().Select(element => element.Name).ToList());
            }
            else 
            {
                return;            
            }
        }

        private void InitializeDataGridViews() 
        {
            if (Activeform.Controls.OfType<DataGridView>().Select(element => element.Name).ToList().Count > 0)
            {
                DataGridViews.AddRange(Activeform.Controls.OfType<DataGridView>().Select(element => element.Name).ToList());
            }
            else
            {
                return;
            }
        }

        private void InitializeListViews() 
        {
            if (Activeform.Controls.OfType<ListView>().Select(element => element.Name).ToList().Count > 0)
            {
                ListViews.AddRange(Activeform.Controls.OfType<ListView>().Select(element => element.Name).ToList());
            }
            else
            {
                return;
            }
        }

        private void BindToListbox() 
        {
        }

        private void BindToDataGridView()
        {
        }

        private void BindToListView() 
        {
        }

        public static void DataGridViewToDataTable(DataGridView dgv, DataTable dt) 
        {
            List<string> names = dgv.Columns.Cast<DataGridViewColumn>().Select(a => a.Name).ToList();
        }
        
        public static void DataTableToDataGridView(DataTable dt , DataGridView dgv) 
        {       
            List<string> names = dt.Columns.Cast<DataColumn>().Select(a => a.ColumnName).ToList();


            if(names.Count != dgv.Columns.Count) 
            {
                MessageBox.Show("Vstup a vystup se neshoduje. Nemas stejny pocet sloupcu.");
                return;
            }

            for (int i = 0; i < dgv.Columns.Count; i++)
            {
                dgv.Columns[i].Name = dt.Columns[i].ColumnName;
            }        
        }
    }
}
