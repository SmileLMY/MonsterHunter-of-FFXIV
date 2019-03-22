using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Aspose.Cells;

namespace Hunter
{
    public partial class Form1 : Form
    {
        int monsterNumber = 70;//怪物数量
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //单选2默认当前时间
            this.dateTimePicker1.Value = System.DateTime.Now;

            this.FillTable();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //set kill time
            int index = dataGridView1.SelectedRows[0].Index;
            string name = dataGridView1.Rows[index].Cells[2].Value.ToString();
            string nameCD = name + "CD";
            DateTime dateTime;
            if (radioButton1.Checked == true)
            {
                dateTime = System.DateTime.Now;
            }
            else
            {
                dateTime = dateTimePicker1.Value;
            }
            Properties.Settings.Default[name] = dateTime;
            Properties.Settings.Default.Save();
            this.dataGridView1.Rows[index].Cells[4].Value = Properties.Settings.Default[name];
            nameCD = Properties.Settings.Default[nameCD].ToString();
            this.dataGridView1.Rows[index].Cells[5].Value = dateTime.AddMinutes(int.Parse(nameCD));
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            //单选1设置当前时间
            this.radioButton1.Text = System.DateTime.Now.ToString();
            //测试是否进入CD
            for (int i = 0; i < this.dataGridView1.Rows.Count; i++)
            {
                if (this.dataGridView1.Rows[i].Cells[4].Value != null)
                {
                    if (DateTime.Compare(Convert.ToDateTime(this.dataGridView1.Rows[i].Cells[5].Value.ToString()), System.DateTime.Now) < 0)
                    {
                        this.dataGridView1.Rows[i].Cells[1].Style.ForeColor = Color.Green;
                        this.dataGridView1.Rows[i].Cells[2].Style.ForeColor = Color.Green;
                        this.dataGridView1.Rows[i].Cells[3].Style.ForeColor = Color.Green;
                        this.dataGridView1.Rows[i].Cells[4].Style.ForeColor = Color.Green;
                        this.dataGridView1.Rows[i].Cells[5].Style.ForeColor = Color.Green;
                    }
                    else
                    {
                        this.dataGridView1.Rows[i].Cells[1].Style.ForeColor = Color.Black;
                        this.dataGridView1.Rows[i].Cells[2].Style.ForeColor = Color.Black;
                        this.dataGridView1.Rows[i].Cells[3].Style.ForeColor = Color.Black;
                        this.dataGridView1.Rows[i].Cells[4].Style.ForeColor = Color.Black;
                        this.dataGridView1.Rows[i].Cells[5].Style.ForeColor = Color.Black;
                    }
                }
            }


        }

        private void button2_Click(object sender, EventArgs e)
        {
            //clear kill time
            int index = dataGridView1.SelectedRows[0].Index;
            string name = dataGridView1.Rows[index].Cells[2].Value.ToString();
            Properties.Settings.Default[name] = null;
            Properties.Settings.Default.Save();
            this.dataGridView1.Rows[index].Cells[4].Value = null;
            this.dataGridView1.Rows[index].Cells[5].Value = null;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //import
            OpenFileDialog fileDialog = new OpenFileDialog
            {
                Filter = "Excel文件(*.xls,xlsx)|*.xls;*.xlsx"
            };
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                string excelName = fileDialog.FileName;
                Workbook excel = new Workbook(excelName);
                Cells cells = excel.Worksheets[0].Cells;
                //列have标题，name/time
                for (int i = 1; i <= cells.MaxDataRow; i++)//max+1?
                {
                    string name = cells[i, 0].StringValue;
                    DateTime time = Convert.ToDateTime(cells[i, 1].StringValue);
                    Properties.Settings.Default[name] = time;
                    Properties.Settings.Default.Save();
                }
                //重新填充
                this.dataGridView1.Rows.Clear();
                this.FillTable();
                MessageBox.Show("导入完成！");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //export
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel文件(*.xls,xlsx)|*.xls;*.xlsx",
                FileName="Hunter.xlsx"
            };
            string path;
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                if (System.IO.File.Exists(saveFileDialog.FileName))
                {
                    System.IO.File.Delete(saveFileDialog.FileName);
                }
                path = saveFileDialog.FileName.ToString();               
                
                Workbook workbook = new Workbook();                
                Worksheet sheet = workbook.Worksheets[0]; //工作表
                Cells cells = sheet.Cells;//单元格               

                int rowCount = this.monsterNumber+1;
                int columnCount = 2;
                for(int i = 0; i < rowCount; i++)
                {
                    if (i == 0)
                    {
                        cells[i, 0].PutValue("name");
                        cells[i, 1].PutValue("kill time");
                    }
                    else
                    {
                        cells[i, 0].PutValue(this.dataGridView1.Rows[i - 1].Cells[2].Value);
                        cells[i, 1].PutValue(this.dataGridView1.Rows[i - 1].Cells[4].Value.ToString());
                    }
                }
                workbook.Save(path);
            }
        }

        private void FillTable()
        {
            //表格+行
            for (int i = 0; i < monsterNumber; i++)
            {
                this.dataGridView1.Rows.Add();
            }

            //填充数据
            #region 2.0
            #region 中拉01
            this.dataGridView1.Rows[0].Cells[1].Value = "S";
            this.dataGridView1.Rows[0].Cells[2].Value = "护土精灵";
            this.dataGridView1.Rows[0].Cells[3].Value = "中拉诺西亚";
            this.dataGridView1.Rows[0].Cells[4].Value = Properties.Settings.Default.护土精灵;
            this.dataGridView1.Rows[0].Cells[5].Value = Properties.Settings.Default.护土精灵.AddMinutes(Properties.Settings.Default.护土精灵CD);


            this.dataGridView1.Rows[1].Cells[1].Value = "A";
            this.dataGridView1.Rows[1].Cells[2].Value = "丑男子沃迦加";
            this.dataGridView1.Rows[1].Cells[3].Value = "中拉诺西亚";
            this.dataGridView1.Rows[1].Cells[4].Value = Properties.Settings.Default.丑男子沃迦加;
            this.dataGridView1.Rows[1].Cells[5].Value = Properties.Settings.Default.丑男子沃迦加.AddMinutes(Properties.Settings.Default.丑男子沃迦加CD);
            #endregion

            #region 东拉23
            this.dataGridView1.Rows[2].Cells[1].Value = "S";
            this.dataGridView1.Rows[2].Cells[2].Value = "伽洛克";
            this.dataGridView1.Rows[2].Cells[3].Value = "东拉诺西亚";
            this.dataGridView1.Rows[2].Cells[4].Value = Properties.Settings.Default.伽洛克;
            this.dataGridView1.Rows[2].Cells[5].Value = Properties.Settings.Default.伽洛克.AddMinutes(Properties.Settings.Default.伽洛克CD);

            this.dataGridView1.Rows[3].Cells[1].Value = "A";
            this.dataGridView1.Rows[3].Cells[2].Value = "魔导地狱爪";
            this.dataGridView1.Rows[3].Cells[3].Value = "东拉诺西亚";
            this.dataGridView1.Rows[3].Cells[4].Value = Properties.Settings.Default.魔导地狱爪;
            this.dataGridView1.Rows[3].Cells[5].Value = Properties.Settings.Default.魔导地狱爪.AddMinutes(Properties.Settings.Default.魔导地狱爪CD);
            #endregion

            #region 西拉45
            this.dataGridView1.Rows[4].Cells[1].Value = "S";
            this.dataGridView1.Rows[4].Cells[2].Value = "火愤牛";
            this.dataGridView1.Rows[4].Cells[3].Value = "西拉诺西亚";
            this.dataGridView1.Rows[4].Cells[4].Value = Properties.Settings.Default.火愤牛;
            this.dataGridView1.Rows[4].Cells[5].Value = Properties.Settings.Default.火愤牛.AddMinutes(Properties.Settings.Default.火愤牛CD);

            this.dataGridView1.Rows[5].Cells[1].Value = "A";
            this.dataGridView1.Rows[5].Cells[2].Value = "纳恩";
            this.dataGridView1.Rows[5].Cells[3].Value = "西拉诺西亚";
            this.dataGridView1.Rows[5].Cells[4].Value = Properties.Settings.Default.纳恩;
            this.dataGridView1.Rows[5].Cells[5].Value = Properties.Settings.Default.纳恩.AddMinutes(Properties.Settings.Default.纳恩CD);
            #endregion

            #region 拉低67
            this.dataGridView1.Rows[6].Cells[1].Value = "S";
            this.dataGridView1.Rows[6].Cells[2].Value = "咕尔呱洛斯";
            this.dataGridView1.Rows[6].Cells[3].Value = "拉诺西亚低地";
            this.dataGridView1.Rows[6].Cells[4].Value = Properties.Settings.Default.咕尔呱洛斯;
            this.dataGridView1.Rows[6].Cells[5].Value = Properties.Settings.Default.咕尔呱洛斯.AddMinutes(Properties.Settings.Default.咕尔呱洛斯CD);

            this.dataGridView1.Rows[7].Cells[1].Value = "A";
            this.dataGridView1.Rows[7].Cells[2].Value = "乌克提希";
            this.dataGridView1.Rows[7].Cells[3].Value = "拉诺西亚低地";
            this.dataGridView1.Rows[7].Cells[4].Value = Properties.Settings.Default.乌克提希;
            this.dataGridView1.Rows[7].Cells[5].Value = Properties.Settings.Default.乌克提希.AddMinutes(Properties.Settings.Default.乌克提希CD);
            #endregion

            #region 拉外89
            this.dataGridView1.Rows[8].Cells[1].Value = "S";
            this.dataGridView1.Rows[8].Cells[2].Value = "牛头黑神";
            this.dataGridView1.Rows[8].Cells[3].Value = "拉诺西亚外地";
            this.dataGridView1.Rows[8].Cells[4].Value = Properties.Settings.Default.牛头黑神;
            this.dataGridView1.Rows[8].Cells[5].Value = Properties.Settings.Default.牛头黑神.AddMinutes(Properties.Settings.Default.牛头黑神CD);

            this.dataGridView1.Rows[9].Cells[1].Value = "A";
            this.dataGridView1.Rows[9].Cells[2].Value = "角祖";
            this.dataGridView1.Rows[9].Cells[3].Value = "拉诺西亚外地";
            this.dataGridView1.Rows[9].Cells[4].Value = Properties.Settings.Default.角祖;
            this.dataGridView1.Rows[9].Cells[5].Value = Properties.Settings.Default.角祖.AddMinutes(Properties.Settings.Default.角祖CD);
            #endregion

            #region 拉高1011
            this.dataGridView1.Rows[10].Cells[1].Value = "S";
            this.dataGridView1.Rows[10].Cells[2].Value = "南迪";
            this.dataGridView1.Rows[10].Cells[3].Value = "拉诺西亚高地";
            this.dataGridView1.Rows[10].Cells[4].Value = Properties.Settings.Default.南迪;
            this.dataGridView1.Rows[10].Cells[5].Value = Properties.Settings.Default.南迪.AddMinutes(Properties.Settings.Default.南迪CD);

            this.dataGridView1.Rows[11].Cells[1].Value = "A";
            this.dataGridView1.Rows[11].Cells[2].Value = "玛贝利";
            this.dataGridView1.Rows[11].Cells[3].Value = "拉诺西亚高地";
            this.dataGridView1.Rows[11].Cells[4].Value = Properties.Settings.Default.玛贝利;
            this.dataGridView1.Rows[11].Cells[5].Value = Properties.Settings.Default.玛贝利.AddMinutes(Properties.Settings.Default.玛贝利CD);
            #endregion

            #region 东森1213
            this.dataGridView1.Rows[12].Cells[1].Value = "S";
            this.dataGridView1.Rows[12].Cells[2].Value = "乌尔伽鲁";
            this.dataGridView1.Rows[12].Cells[3].Value = "黑衣森林东部林区";
            this.dataGridView1.Rows[12].Cells[4].Value = Properties.Settings.Default.乌尔伽鲁;
            this.dataGridView1.Rows[12].Cells[5].Value = Properties.Settings.Default.乌尔伽鲁.AddMinutes(Properties.Settings.Default.乌尔伽鲁CD);

            this.dataGridView1.Rows[13].Cells[1].Value = "A";
            this.dataGridView1.Rows[13].Cells[2].Value = "千眼凝胶";
            this.dataGridView1.Rows[13].Cells[3].Value = "黑衣森林东部林区";
            this.dataGridView1.Rows[13].Cells[4].Value = Properties.Settings.Default.千眼凝胶;
            this.dataGridView1.Rows[13].Cells[5].Value = Properties.Settings.Default.千眼凝胶.AddMinutes(Properties.Settings.Default.千眼凝胶CD);
            #endregion

            #region 南森1415
            this.dataGridView1.Rows[14].Cells[1].Value = "S";
            this.dataGridView1.Rows[14].Cells[2].Value = "夺心魔";
            this.dataGridView1.Rows[14].Cells[3].Value = "黑衣森林南部林区";
            this.dataGridView1.Rows[14].Cells[4].Value = Properties.Settings.Default.夺心魔;
            this.dataGridView1.Rows[14].Cells[5].Value = Properties.Settings.Default.夺心魔.AddMinutes(Properties.Settings.Default.夺心魔CD);

            this.dataGridView1.Rows[15].Cells[1].Value = "A";
            this.dataGridView1.Rows[15].Cells[2].Value = "盖得";
            this.dataGridView1.Rows[15].Cells[3].Value = "黑衣森林南部林区";
            this.dataGridView1.Rows[15].Cells[4].Value = Properties.Settings.Default.盖得;
            this.dataGridView1.Rows[15].Cells[5].Value = Properties.Settings.Default.盖得.AddMinutes(Properties.Settings.Default.盖得CD);
            #endregion

            #region 南森1617
            this.dataGridView1.Rows[16].Cells[1].Value = "S";
            this.dataGridView1.Rows[16].Cells[2].Value = "千竿口花希达";
            this.dataGridView1.Rows[16].Cells[3].Value = "黑衣森林北部林区";
            this.dataGridView1.Rows[16].Cells[4].Value = Properties.Settings.Default.千竿口花希达;
            this.dataGridView1.Rows[16].Cells[5].Value = Properties.Settings.Default.千竿口花希达.AddMinutes(Properties.Settings.Default.千竿口花希达CD);

            this.dataGridView1.Rows[17].Cells[1].Value = "A";
            this.dataGridView1.Rows[17].Cells[2].Value = "尾宿蛛蝎";
            this.dataGridView1.Rows[17].Cells[3].Value = "黑衣森林北部林区";
            this.dataGridView1.Rows[17].Cells[4].Value = Properties.Settings.Default.尾宿蛛蝎;
            this.dataGridView1.Rows[17].Cells[5].Value = Properties.Settings.Default.尾宿蛛蝎.AddMinutes(Properties.Settings.Default.尾宿蛛蝎CD);
            #endregion

            #region 中森1819
            this.dataGridView1.Rows[18].Cells[1].Value = "S";
            this.dataGridView1.Rows[18].Cells[2].Value = "雷德罗巨蛇";
            this.dataGridView1.Rows[18].Cells[3].Value = "黑衣森林中央林区";
            this.dataGridView1.Rows[18].Cells[4].Value = Properties.Settings.Default.雷德罗巨蛇;
            this.dataGridView1.Rows[18].Cells[5].Value = Properties.Settings.Default.雷德罗巨蛇.AddMinutes(Properties.Settings.Default.雷德罗巨蛇CD);

            this.dataGridView1.Rows[19].Cells[1].Value = "A";
            this.dataGridView1.Rows[19].Cells[2].Value = "弗内乌斯";
            this.dataGridView1.Rows[19].Cells[3].Value = "黑衣森林中央林区";
            this.dataGridView1.Rows[19].Cells[4].Value = Properties.Settings.Default.弗内乌斯;
            this.dataGridView1.Rows[19].Cells[5].Value = Properties.Settings.Default.弗内乌斯.AddMinutes(Properties.Settings.Default.弗内乌斯CD);
            #endregion

            #region 东萨2021
            this.dataGridView1.Rows[20].Cells[1].Value = "S";
            this.dataGridView1.Rows[20].Cells[2].Value = "巴拉乌尔";
            this.dataGridView1.Rows[20].Cells[3].Value = "东萨纳兰";
            this.dataGridView1.Rows[20].Cells[4].Value = Properties.Settings.Default.巴拉乌尔;
            this.dataGridView1.Rows[20].Cells[5].Value = Properties.Settings.Default.巴拉乌尔.AddMinutes(Properties.Settings.Default.巴拉乌尔CD);

            this.dataGridView1.Rows[21].Cells[1].Value = "A";
            this.dataGridView1.Rows[21].Cells[2].Value = "玛赫斯";
            this.dataGridView1.Rows[21].Cells[3].Value = "东萨纳兰";
            this.dataGridView1.Rows[21].Cells[4].Value = Properties.Settings.Default.玛赫斯;
            this.dataGridView1.Rows[21].Cells[5].Value = Properties.Settings.Default.玛赫斯.AddMinutes(Properties.Settings.Default.玛赫斯CD);
            #endregion

            #region 西萨2223
            this.dataGridView1.Rows[22].Cells[1].Value = "S";
            this.dataGridView1.Rows[22].Cells[2].Value = "虚无探索者";
            this.dataGridView1.Rows[22].Cells[3].Value = "西萨纳兰";
            this.dataGridView1.Rows[22].Cells[4].Value = Properties.Settings.Default.虚无探索者;
            this.dataGridView1.Rows[22].Cells[5].Value = Properties.Settings.Default.虚无探索者.AddMinutes(Properties.Settings.Default.虚无探索者CD);

            this.dataGridView1.Rows[23].Cells[1].Value = "A";
            this.dataGridView1.Rows[23].Cells[2].Value = "阿列刻特利昂";
            this.dataGridView1.Rows[23].Cells[3].Value = "西萨纳兰";
            this.dataGridView1.Rows[23].Cells[4].Value = Properties.Settings.Default.阿列刻特利昂;
            this.dataGridView1.Rows[23].Cells[5].Value = Properties.Settings.Default.阿列刻特利昂.AddMinutes(Properties.Settings.Default.阿列刻特利昂CD);
            #endregion

            #region 南萨2425
            this.dataGridView1.Rows[24].Cells[1].Value = "S";
            this.dataGridView1.Rows[24].Cells[2].Value = "努纽努维";
            this.dataGridView1.Rows[24].Cells[3].Value = "南萨纳兰";
            this.dataGridView1.Rows[24].Cells[4].Value = Properties.Settings.Default.努纽努维;
            this.dataGridView1.Rows[24].Cells[5].Value = Properties.Settings.Default.努纽努维.AddMinutes(Properties.Settings.Default.努纽努维CD);

            this.dataGridView1.Rows[25].Cells[1].Value = "A";
            this.dataGridView1.Rows[25].Cells[2].Value = "札尼戈";
            this.dataGridView1.Rows[25].Cells[3].Value = "南萨纳兰";
            this.dataGridView1.Rows[25].Cells[4].Value = Properties.Settings.Default.札尼戈;
            this.dataGridView1.Rows[25].Cells[5].Value = Properties.Settings.Default.札尼戈.AddMinutes(Properties.Settings.Default.札尼戈CD);
            #endregion

            #region 北萨2627
            this.dataGridView1.Rows[26].Cells[1].Value = "S";
            this.dataGridView1.Rows[26].Cells[2].Value = "蚓螈巨虫";
            this.dataGridView1.Rows[26].Cells[3].Value = "北萨纳兰";
            this.dataGridView1.Rows[26].Cells[4].Value = Properties.Settings.Default.蚓螈巨虫;
            this.dataGridView1.Rows[26].Cells[5].Value = Properties.Settings.Default.蚓螈巨虫.AddMinutes(Properties.Settings.Default.蚓螈巨虫CD);

            this.dataGridView1.Rows[27].Cells[1].Value = "A";
            this.dataGridView1.Rows[27].Cells[2].Value = "菲兰德的遗火";
            this.dataGridView1.Rows[27].Cells[3].Value = "北萨纳兰";
            this.dataGridView1.Rows[27].Cells[4].Value = Properties.Settings.Default.菲兰德的遗火;
            this.dataGridView1.Rows[27].Cells[5].Value = Properties.Settings.Default.菲兰德的遗火.AddMinutes(Properties.Settings.Default.菲兰德的遗火CD);
            #endregion

            #region 中萨2829
            this.dataGridView1.Rows[28].Cells[1].Value = "S";
            this.dataGridView1.Rows[28].Cells[2].Value = "布隆特斯";
            this.dataGridView1.Rows[28].Cells[3].Value = "中萨纳兰";
            this.dataGridView1.Rows[28].Cells[4].Value = Properties.Settings.Default.布隆特斯;
            this.dataGridView1.Rows[28].Cells[5].Value = Properties.Settings.Default.布隆特斯.AddMinutes(Properties.Settings.Default.布隆特斯CD);

            this.dataGridView1.Rows[29].Cells[1].Value = "A";
            this.dataGridView1.Rows[29].Cells[2].Value = "花舞仙人刺";
            this.dataGridView1.Rows[29].Cells[3].Value = "中萨纳兰";
            this.dataGridView1.Rows[29].Cells[4].Value = Properties.Settings.Default.花舞仙人刺;
            this.dataGridView1.Rows[29].Cells[5].Value = Properties.Settings.Default.花舞仙人刺.AddMinutes(Properties.Settings.Default.花舞仙人刺CD);
            #endregion

            #region 中高3031
            this.dataGridView1.Rows[30].Cells[1].Value = "S";
            this.dataGridView1.Rows[30].Cells[2].Value = "萨法特";
            this.dataGridView1.Rows[30].Cells[3].Value = "库尔札斯中央高地";
            this.dataGridView1.Rows[30].Cells[4].Value = Properties.Settings.Default.萨法特;
            this.dataGridView1.Rows[30].Cells[5].Value = Properties.Settings.Default.萨法特.AddMinutes(Properties.Settings.Default.萨法特CD);

            this.dataGridView1.Rows[31].Cells[1].Value = "A";
            this.dataGridView1.Rows[31].Cells[2].Value = "马拉克";
            this.dataGridView1.Rows[31].Cells[3].Value = "库尔札斯中央高地";
            this.dataGridView1.Rows[31].Cells[4].Value = Properties.Settings.Default.马拉克;
            this.dataGridView1.Rows[31].Cells[5].Value = Properties.Settings.Default.马拉克.AddMinutes(Properties.Settings.Default.马拉克CD);
            #endregion

            #region 魔杜纳3233
            this.dataGridView1.Rows[32].Cells[1].Value = "S";
            this.dataGridView1.Rows[32].Cells[2].Value = "阿格里帕";
            this.dataGridView1.Rows[32].Cells[3].Value = "魔杜纳";
            this.dataGridView1.Rows[32].Cells[4].Value = Properties.Settings.Default.阿格里帕;
            this.dataGridView1.Rows[32].Cells[5].Value = Properties.Settings.Default.阿格里帕.AddMinutes(Properties.Settings.Default.阿格里帕CD);

            this.dataGridView1.Rows[33].Cells[1].Value = "A";
            this.dataGridView1.Rows[33].Cells[2].Value = "库雷亚";
            this.dataGridView1.Rows[33].Cells[3].Value = "魔杜纳";
            this.dataGridView1.Rows[33].Cells[4].Value = Properties.Settings.Default.库雷亚;
            this.dataGridView1.Rows[33].Cells[5].Value = Properties.Settings.Default.库雷亚.AddMinutes(Properties.Settings.Default.库雷亚CD);
            #endregion
            #endregion

            #region 3.0
            #region 西高343536
            this.dataGridView1.Rows[34].Cells[1].Value = "S";
            this.dataGridView1.Rows[34].Cells[2].Value = "凯撒贝希摩斯";
            this.dataGridView1.Rows[34].Cells[3].Value = "库尔札斯西部高地";
            this.dataGridView1.Rows[34].Cells[4].Value = Properties.Settings.Default.凯撒贝希摩斯;
            this.dataGridView1.Rows[34].Cells[5].Value = Properties.Settings.Default.凯撒贝希摩斯.AddMinutes(Properties.Settings.Default.凯撒贝希摩斯CD);

            this.dataGridView1.Rows[35].Cells[1].Value = "A";
            this.dataGridView1.Rows[35].Cells[2].Value = "米勒卡";
            this.dataGridView1.Rows[35].Cells[3].Value = "库尔札斯西部高地";
            this.dataGridView1.Rows[35].Cells[4].Value = Properties.Settings.Default.米勒卡;
            this.dataGridView1.Rows[35].Cells[5].Value = Properties.Settings.Default.米勒卡.AddMinutes(Properties.Settings.Default.ACD);

            this.dataGridView1.Rows[36].Cells[1].Value = "A";
            this.dataGridView1.Rows[36].Cells[2].Value = "卢芭";
            this.dataGridView1.Rows[36].Cells[3].Value = "库尔札斯西部高地";
            this.dataGridView1.Rows[36].Cells[4].Value = Properties.Settings.Default.卢芭;
            this.dataGridView1.Rows[36].Cells[5].Value = Properties.Settings.Default.卢芭.AddMinutes(Properties.Settings.Default.ACD);
            #endregion

            #region 龙高373839
            this.dataGridView1.Rows[37].Cells[1].Value = "S";
            this.dataGridView1.Rows[37].Cells[2].Value = "神穆尔鸟";
            this.dataGridView1.Rows[37].Cells[3].Value = "龙堡参天高地";
            this.dataGridView1.Rows[37].Cells[4].Value = Properties.Settings.Default.神穆尔鸟;
            this.dataGridView1.Rows[37].Cells[5].Value = Properties.Settings.Default.神穆尔鸟.AddMinutes(Properties.Settings.Default.神穆尔鸟CD);

            this.dataGridView1.Rows[38].Cells[1].Value = "A";
            this.dataGridView1.Rows[38].Cells[2].Value = "双足飞龙之王";
            this.dataGridView1.Rows[38].Cells[3].Value = "龙堡参天高地";
            this.dataGridView1.Rows[38].Cells[4].Value = Properties.Settings.Default.双足飞龙之王;
            this.dataGridView1.Rows[38].Cells[5].Value = Properties.Settings.Default.双足飞龙之王.AddMinutes(Properties.Settings.Default.ACD);

            this.dataGridView1.Rows[39].Cells[1].Value = "A";
            this.dataGridView1.Rows[39].Cells[2].Value = "派拉斯特暴龙";
            this.dataGridView1.Rows[39].Cells[3].Value = "龙堡参天高地";
            this.dataGridView1.Rows[39].Cells[4].Value = Properties.Settings.Default.派拉斯特暴龙;
            this.dataGridView1.Rows[39].Cells[5].Value = Properties.Settings.Default.派拉斯特暴龙.AddMinutes(Properties.Settings.Default.ACD);
            #endregion

            #region 龙低404142
            this.dataGridView1.Rows[40].Cells[1].Value = "S";
            this.dataGridView1.Rows[40].Cells[2].Value = "苍白骑士";
            this.dataGridView1.Rows[40].Cells[3].Value = "龙堡内陆低地";
            this.dataGridView1.Rows[40].Cells[4].Value = Properties.Settings.Default.苍白骑士;
            this.dataGridView1.Rows[40].Cells[5].Value = Properties.Settings.Default.苍白骑士.AddMinutes(Properties.Settings.Default.苍白骑士CD);

            this.dataGridView1.Rows[41].Cells[1].Value = "A";
            this.dataGridView1.Rows[41].Cells[2].Value = "斯特拉斯";
            this.dataGridView1.Rows[41].Cells[3].Value = "龙堡内陆低地";
            this.dataGridView1.Rows[41].Cells[4].Value = Properties.Settings.Default.斯特拉斯;
            this.dataGridView1.Rows[41].Cells[5].Value = Properties.Settings.Default.斯特拉斯.AddMinutes(Properties.Settings.Default.ACD);

            this.dataGridView1.Rows[42].Cells[1].Value = "A";
            this.dataGridView1.Rows[42].Cells[2].Value = "机工兵斯利普金克斯";
            this.dataGridView1.Rows[42].Cells[3].Value = "龙堡内陆低地";
            this.dataGridView1.Rows[42].Cells[4].Value = Properties.Settings.Default.机工兵斯利普金克斯;
            this.dataGridView1.Rows[42].Cells[5].Value = Properties.Settings.Default.机工兵斯利普金克斯.AddMinutes(Properties.Settings.Default.ACD);
            #endregion

            #region 翻云雾海434445
            this.dataGridView1.Rows[43].Cells[1].Value = "S";
            this.dataGridView1.Rows[43].Cells[2].Value = "刚德瑞瓦";
            this.dataGridView1.Rows[43].Cells[3].Value = "翻云雾海";
            this.dataGridView1.Rows[43].Cells[4].Value = Properties.Settings.Default.刚德瑞瓦;
            this.dataGridView1.Rows[43].Cells[5].Value = Properties.Settings.Default.刚德瑞瓦.AddMinutes(Properties.Settings.Default.刚德瑞瓦CD);

            this.dataGridView1.Rows[44].Cells[1].Value = "A";
            this.dataGridView1.Rows[44].Cells[2].Value = "布涅";
            this.dataGridView1.Rows[44].Cells[3].Value = "翻云雾海";
            this.dataGridView1.Rows[44].Cells[4].Value = Properties.Settings.Default.布涅;
            this.dataGridView1.Rows[44].Cells[5].Value = Properties.Settings.Default.布涅.AddMinutes(Properties.Settings.Default.ACD);

            this.dataGridView1.Rows[45].Cells[1].Value = "A";
            this.dataGridView1.Rows[45].Cells[2].Value = "阿伽托斯";
            this.dataGridView1.Rows[45].Cells[3].Value = "翻云雾海";
            this.dataGridView1.Rows[45].Cells[4].Value = Properties.Settings.Default.阿伽托斯;
            this.dataGridView1.Rows[45].Cells[5].Value = Properties.Settings.Default.阿伽托斯.AddMinutes(Properties.Settings.Default.ACD);
            #endregion

            #region 云海464748
            this.dataGridView1.Rows[46].Cells[1].Value = "S";
            this.dataGridView1.Rows[46].Cells[2].Value = "极乐鸟";
            this.dataGridView1.Rows[46].Cells[3].Value = "阿巴拉提亚云海";
            this.dataGridView1.Rows[46].Cells[4].Value = Properties.Settings.Default.极乐鸟;
            this.dataGridView1.Rows[46].Cells[5].Value = Properties.Settings.Default.极乐鸟.AddMinutes(Properties.Settings.Default.极乐鸟CD);

            this.dataGridView1.Rows[47].Cells[1].Value = "A";
            this.dataGridView1.Rows[47].Cells[2].Value = "西斯尤";
            this.dataGridView1.Rows[47].Cells[3].Value = "阿巴拉提亚云海";
            this.dataGridView1.Rows[47].Cells[4].Value = Properties.Settings.Default.西斯尤;
            this.dataGridView1.Rows[47].Cells[5].Value = Properties.Settings.Default.西斯尤.AddMinutes(Properties.Settings.Default.ACD);

            this.dataGridView1.Rows[48].Cells[1].Value = "A";
            this.dataGridView1.Rows[48].Cells[2].Value = "恩克拉多斯";
            this.dataGridView1.Rows[48].Cells[3].Value = "阿巴拉提亚云海";
            this.dataGridView1.Rows[48].Cells[4].Value = Properties.Settings.Default.恩克拉多斯;
            this.dataGridView1.Rows[48].Cells[5].Value = Properties.Settings.Default.恩克拉多斯.AddMinutes(Properties.Settings.Default.ACD);
            #endregion

            #region 魔大陆495051
            this.dataGridView1.Rows[49].Cells[1].Value = "S";
            this.dataGridView1.Rows[49].Cells[2].Value = "卢克洛塔";
            this.dataGridView1.Rows[49].Cells[3].Value = "魔大陆";
            this.dataGridView1.Rows[49].Cells[4].Value = Properties.Settings.Default.卢克洛塔;
            this.dataGridView1.Rows[49].Cells[5].Value = Properties.Settings.Default.卢克洛塔.AddMinutes(Properties.Settings.Default.卢克洛塔CD);

            this.dataGridView1.Rows[50].Cells[1].Value = "A";
            this.dataGridView1.Rows[50].Cells[2].Value = "坎帕提";
            this.dataGridView1.Rows[50].Cells[3].Value = "魔大陆";
            this.dataGridView1.Rows[50].Cells[4].Value = Properties.Settings.Default.坎帕提;
            this.dataGridView1.Rows[50].Cells[5].Value = Properties.Settings.Default.坎帕提.AddMinutes(Properties.Settings.Default.ACD);

            this.dataGridView1.Rows[51].Cells[1].Value = "A";
            this.dataGridView1.Rows[51].Cells[2].Value = "恶臭狂花";
            this.dataGridView1.Rows[51].Cells[3].Value = "魔大陆";
            this.dataGridView1.Rows[51].Cells[4].Value = Properties.Settings.Default.恶臭狂花;
            this.dataGridView1.Rows[51].Cells[5].Value = Properties.Settings.Default.恶臭狂花.AddMinutes(Properties.Settings.Default.ACD);
            #endregion
            #endregion

            #region 4.0
            #region 边区525354
            this.dataGridView1.Rows[52].Cells[1].Value = "S";
            this.dataGridView1.Rows[52].Cells[2].Value = "优昙婆罗花";
            this.dataGridView1.Rows[52].Cells[3].Value = "基拉巴尼亚边区";
            this.dataGridView1.Rows[52].Cells[4].Value = Properties.Settings.Default.优昙婆罗花;
            this.dataGridView1.Rows[52].Cells[5].Value = Properties.Settings.Default.优昙婆罗花.AddMinutes(Properties.Settings.Default.fourSCD);

            this.dataGridView1.Rows[53].Cells[1].Value = "A";
            this.dataGridView1.Rows[53].Cells[2].Value = "女王蜂";
            this.dataGridView1.Rows[53].Cells[3].Value = "基拉巴尼亚边区";
            this.dataGridView1.Rows[53].Cells[4].Value = Properties.Settings.Default.女王蜂;
            this.dataGridView1.Rows[53].Cells[5].Value = Properties.Settings.Default.女王蜂.AddMinutes(Properties.Settings.Default.ACD);

            this.dataGridView1.Rows[54].Cells[1].Value = "A";
            this.dataGridView1.Rows[54].Cells[2].Value = "奥迦斯";
            this.dataGridView1.Rows[54].Cells[3].Value = "基拉巴尼亚边区";
            this.dataGridView1.Rows[54].Cells[4].Value = Properties.Settings.Default.奥迦斯;
            this.dataGridView1.Rows[54].Cells[5].Value = Properties.Settings.Default.奥迦斯.AddMinutes(Properties.Settings.Default.ACD);
            #endregion

            #region 山区555657
            this.dataGridView1.Rows[55].Cells[1].Value = "S";
            this.dataGridView1.Rows[55].Cells[2].Value = "爬骨怪龙";
            this.dataGridView1.Rows[55].Cells[3].Value = "基拉巴尼亚山区";
            this.dataGridView1.Rows[55].Cells[4].Value = Properties.Settings.Default.爬骨怪龙;
            this.dataGridView1.Rows[55].Cells[5].Value = Properties.Settings.Default.爬骨怪龙.AddMinutes(Properties.Settings.Default.fourSCD);

            this.dataGridView1.Rows[56].Cells[1].Value = "A";
            this.dataGridView1.Rows[56].Cells[2].Value = "熔骨炎蝎";
            this.dataGridView1.Rows[56].Cells[3].Value = "基拉巴尼亚山区";
            this.dataGridView1.Rows[56].Cells[4].Value = Properties.Settings.Default.熔骨炎蝎;
            this.dataGridView1.Rows[56].Cells[5].Value = Properties.Settings.Default.熔骨炎蝎.AddMinutes(Properties.Settings.Default.ACD);

            this.dataGridView1.Rows[57].Cells[1].Value = "A";
            this.dataGridView1.Rows[57].Cells[2].Value = "弗克施泰因";
            this.dataGridView1.Rows[57].Cells[3].Value = "基拉巴尼亚山区";
            this.dataGridView1.Rows[57].Cells[4].Value = Properties.Settings.Default.弗克施泰因;
            this.dataGridView1.Rows[57].Cells[5].Value = Properties.Settings.Default.弗克施泰因.AddMinutes(Properties.Settings.Default.ACD);
            #endregion

            #region 湖区585960
            this.dataGridView1.Rows[58].Cells[1].Value = "S";
            this.dataGridView1.Rows[58].Cells[2].Value = "盐和光";
            this.dataGridView1.Rows[58].Cells[3].Value = "基拉巴尼亚湖区";
            this.dataGridView1.Rows[58].Cells[4].Value = Properties.Settings.Default.盐和光;
            this.dataGridView1.Rows[58].Cells[5].Value = Properties.Settings.Default.盐和光.AddMinutes(Properties.Settings.Default.fourSCD);

            this.dataGridView1.Rows[59].Cells[1].Value = "A";
            this.dataGridView1.Rows[59].Cells[2].Value = "马希沙";
            this.dataGridView1.Rows[59].Cells[3].Value = "基拉巴尼亚湖区";
            this.dataGridView1.Rows[59].Cells[4].Value = Properties.Settings.Default.马希沙;
            this.dataGridView1.Rows[59].Cells[5].Value = Properties.Settings.Default.马希沙.AddMinutes(Properties.Settings.Default.ACD);

            this.dataGridView1.Rows[60].Cells[1].Value = "A";
            this.dataGridView1.Rows[60].Cells[2].Value = "泛光晶体";
            this.dataGridView1.Rows[60].Cells[3].Value = "基拉巴尼亚湖区";
            this.dataGridView1.Rows[60].Cells[4].Value = Properties.Settings.Default.泛光晶体;
            this.dataGridView1.Rows[60].Cells[5].Value = Properties.Settings.Default.泛光晶体.AddMinutes(Properties.Settings.Default.ACD);
            #endregion

            #region 红玉海616263
            this.dataGridView1.Rows[61].Cells[1].Value = "S";
            this.dataGridView1.Rows[61].Cells[2].Value = "巨大鳐";
            this.dataGridView1.Rows[61].Cells[3].Value = "红玉海";
            this.dataGridView1.Rows[61].Cells[4].Value = Properties.Settings.Default.巨大鳐;
            this.dataGridView1.Rows[61].Cells[5].Value = Properties.Settings.Default.巨大鳐.AddMinutes(Properties.Settings.Default.fourSCD);

            this.dataGridView1.Rows[62].Cells[1].Value = "A";
            this.dataGridView1.Rows[62].Cells[2].Value = "船幽灵";
            this.dataGridView1.Rows[62].Cells[3].Value = "红玉海";
            this.dataGridView1.Rows[62].Cells[4].Value = Properties.Settings.Default.船幽灵;
            this.dataGridView1.Rows[62].Cells[5].Value = Properties.Settings.Default.船幽灵.AddMinutes(Properties.Settings.Default.ACD);

            this.dataGridView1.Rows[63].Cells[1].Value = "A";
            this.dataGridView1.Rows[63].Cells[2].Value = "鬼观梦";
            this.dataGridView1.Rows[63].Cells[3].Value = "红玉海";
            this.dataGridView1.Rows[63].Cells[4].Value = Properties.Settings.Default.鬼观梦;
            this.dataGridView1.Rows[63].Cells[5].Value = Properties.Settings.Default.鬼观梦.AddMinutes(Properties.Settings.Default.ACD);
            #endregion

            #region 太阳神草原646566
            this.dataGridView1.Rows[64].Cells[1].Value = "S";
            this.dataGridView1.Rows[64].Cells[2].Value = "兀鲁忽乃朝鲁";
            this.dataGridView1.Rows[64].Cells[3].Value = "太阳神草原";
            this.dataGridView1.Rows[64].Cells[4].Value = Properties.Settings.Default.兀鲁忽乃朝鲁;
            this.dataGridView1.Rows[64].Cells[5].Value = Properties.Settings.Default.兀鲁忽乃朝鲁.AddMinutes(Properties.Settings.Default.fourSCD);

            this.dataGridView1.Rows[65].Cells[1].Value = "A";
            this.dataGridView1.Rows[65].Cells[2].Value = "硕姆";
            this.dataGridView1.Rows[65].Cells[3].Value = "太阳神草原";
            this.dataGridView1.Rows[65].Cells[4].Value = Properties.Settings.Default.硕姆;
            this.dataGridView1.Rows[65].Cells[5].Value = Properties.Settings.Default.硕姆.AddMinutes(Properties.Settings.Default.ACD);

            this.dataGridView1.Rows[66].Cells[1].Value = "A";
            this.dataGridView1.Rows[66].Cells[2].Value = "基里麦卡拉";
            this.dataGridView1.Rows[66].Cells[3].Value = "太阳神草原";
            this.dataGridView1.Rows[66].Cells[4].Value = Properties.Settings.Default.基里麦卡拉;
            this.dataGridView1.Rows[66].Cells[5].Value = Properties.Settings.Default.基里麦卡拉.AddMinutes(Properties.Settings.Default.ACD);
            #endregion

            #region 延夏676869
            this.dataGridView1.Rows[67].Cells[1].Value = "S";
            this.dataGridView1.Rows[67].Cells[2].Value = "伽马";
            this.dataGridView1.Rows[67].Cells[3].Value = "延夏";
            this.dataGridView1.Rows[67].Cells[4].Value = Properties.Settings.Default.伽马;
            this.dataGridView1.Rows[67].Cells[5].Value = Properties.Settings.Default.伽马.AddMinutes(Properties.Settings.Default.fourSCD);

            this.dataGridView1.Rows[68].Cells[1].Value = "A";
            this.dataGridView1.Rows[68].Cells[2].Value = "安迦达";
            this.dataGridView1.Rows[68].Cells[3].Value = "延夏";
            this.dataGridView1.Rows[68].Cells[4].Value = Properties.Settings.Default.安迦达;
            this.dataGridView1.Rows[68].Cells[5].Value = Properties.Settings.Default.安迦达.AddMinutes(Properties.Settings.Default.ACD);

            this.dataGridView1.Rows[69].Cells[1].Value = "A";
            this.dataGridView1.Rows[69].Cells[2].Value = "象魔修罗";
            this.dataGridView1.Rows[69].Cells[3].Value = "延夏";
            this.dataGridView1.Rows[69].Cells[4].Value = Properties.Settings.Default.象魔修罗;
            this.dataGridView1.Rows[69].Cells[5].Value = Properties.Settings.Default.象魔修罗.AddMinutes(Properties.Settings.Default.ACD);
            #endregion
            #endregion

            #region 5.0

            #endregion
        }
    }
}
