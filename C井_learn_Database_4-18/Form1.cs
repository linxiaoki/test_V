using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ADOX;
using System.IO;

namespace C井_learn_Database_4_18
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
           
            //new Access ac = new Access();

            string FilePath = @"C:\自动创建的Access数据表.mdb";

            ADOX.Column column1 = new Column();

            ADOX.Column column2 = new Column();

            ADOX.Column column3 = new Column();

            //设置字段名：名字

            column1.Type = ADOX.DataTypeEnum.adVarWChar;//设置类型为

            column1.DefinedSize = 255;//设置长度

            column1.Name = "名字";//设置字段名

            //设置字段名：性别

            column2.Type = ADOX.DataTypeEnum.adVarWChar;//设置类型为

            column2.DefinedSize = 255;//设置长度

            column2.Name = "性别";//设置字段名

            //设置字段名：年龄

            column3.Type = ADOX.DataTypeEnum.adInteger;//设置类型为

            column3.DefinedSize = 9;//设置长度

            column3.Name = "年龄";//设置字段名

            //column.Properties["AutoIncrement"].Value = true;//设置自动增长

            if (CreateAccessTable(FilePath, "Administrator", column1, column2, column3))
            {

                MessageBox.Show("创建成功", "提示");

            } 
        }
        //创建Access数据库
        public static bool CreateAccess(string strFilePath) {
            ADOX.Catalog clg = new Catalog();
            if (!File.Exists(strFilePath)) {
                try
                {
                    clg.Create("Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " + strFilePath + ";Jet OLEDB:Engine Type=5");   
                }
                catch(System.Exception ex) {
                    MessageBox.Show("数据库创建失败", "提示");
                    return false;
                }
            }
            return true; 
        }
        //创建数据表
        public bool CreateAccessTable(string FilePath, string tableName, params ADOX.Column[] colums)
        {
            bool bolReturn = false;
            ADOX.Catalog clg = new Catalog();
            //数据库文件存在
            try { 
                if(CreateAccess (FilePath)==true){
                    ADODB.Connection cn =new ADODB.Connection();  //连接已创建的数据库文件
                    cn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+FilePath,null,null,-1);
                    clg.ActiveConnection=cn;    //打开已创建的数据库文件
                    
                    ADOX.Table table1 = new ADOX.Table ();
                    table1.Name=tableName;
                    foreach (var column in colums){
                        if(column.Name !=null){
                            table1.Columns.Append (column);
                        }
                    }
                    clg.Tables.Append(table1);
                    cn.Close();
                    bolReturn =true;
                }
            }catch(Exception ex){
                MessageBox.Show("创建失败\r\n" + ex.ToString(), "提示");
            }
            return bolReturn;
        }

    }
    
}
