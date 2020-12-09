using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PlusTest
{
    public partial class fmVolTop : Form
    {
        private DataTable m_volTable = null;
        private string m_sVolUnit = "거래대금(만원)";
        public fmVolTop()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            InitGridTable();

            RqVolTop();
        }

        void InitGridTable()
        {

            if (m_volTable != null)
                m_volTable.Clear();

            m_volTable = new DataTable();

            // 거래대금 단위: 거래소 - 만원, 코스닥 - 천원
            m_sVolUnit = "거래대금(만원)";
            if (m_rdoMarket2.Checked)
                m_sVolUnit = "거래대금(천원)";


            //////////////////////////////////////////////////////
            // 일자별 그리드 
            m_volTable.Columns.Add("순위");
            m_volTable.Columns.Add("종목코드");
            m_volTable.Columns.Add("종목명");
            m_volTable.Columns.Add("현재가");
            m_volTable.Columns.Add("대비");
            m_volTable.Columns.Add("대비율");
            m_volTable.Columns.Add("거래량");
            m_volTable.Columns.Add(m_sVolUnit);


            m_VolTopGrid.DataSource = m_volTable;

            m_VolTopGrid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            m_VolTopGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            m_VolTopGrid.AllowUserToResizeRows = false;


            // 컬럼 정렬
            foreach (DataGridViewColumn cols in m_VolTopGrid.Columns)
            {
                if ((cols.HeaderText == "종목명") || (cols.HeaderText == "종목코드"))
                    cols.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                else
                    cols.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }


            m_VolTopGrid.Columns[0].Width = 50;
            m_VolTopGrid.Font = new Font("Tahoma", 9.75F, FontStyle.Regular);
            m_VolTopGrid.RowHeadersVisible = false;   // 왼쪽 커서 제거 
            m_VolTopGrid.AllowUserToAddRows = false;  // 사용자 행 추가 못하도록 
            m_VolTopGrid.Refresh();
        }

        private bool RqVolTop()
        {
            CPSYSDIBLib.CpSvr7049 objRq = new CPSYSDIBLib.CpSvr7049();


            // 시장구분 "4" 전체, "1" 거래소, "2" 코스닥
            string sMarket = "4";   // 전체 

            if (m_rdoMarket1.Checked)
                sMarket = "1";   // 거래소
            else if (m_rdoMarket2.Checked)
                sMarket = "2";   // 코스닥

            objRq.SetInputValue(0, sMarket);


            // 거래량 상위, 거래대금 상위 선택 - "V" 거래량,  "A" 거래대금 
            string sVolFlag = "V";
            if (m_rdoV2.Checked)
                sVolFlag = "A";

            string sChk1 = "N";
            string sChk2 = "N";
            if (m_chkExc1.Checked == true)
                sChk1 = "Y";
            if (m_chkExc2.Checked == true)
                sChk2 = "Y";
            objRq.SetInputValue(1, sVolFlag);  // 선택 구분: "V" 거래량,  "A" 거래대금 

            objRq.SetInputValue(2, sChk1);  // 관리 제외 Y/N
            objRq.SetInputValue(3, sChk2);  // 우선주 제외 Y/N

            CpCommon.waitRqLimit(COMM_TYPE.SISE);
            int nRet = objRq.BlockRequest();
            if (false == CpCommon.checkRqError(objRq, nRet))
                return false;

            int nCnt = objRq.GetHeaderValue(0);
            Hashtable mapData = new Hashtable();

            for (int i = 0; i < nCnt; i++)
            {
                mapData["순위"] = objRq.GetDataValue(0, i);
                mapData["종목코드"] = objRq.GetDataValue(1, i);
                mapData["종목명"] = objRq.GetDataValue(2, i);
                mapData["현재가"] = objRq.GetDataValue(3, i);
                mapData["대비"] = objRq.GetDataValue(4, i);
                mapData["대비율"] = objRq.GetDataValue(5, i);
                mapData["거래량"] = objRq.GetDataValue(6, i);
                mapData[m_sVolUnit] = objRq.GetDataValue(7, i);

                DataRow row = m_volTable.NewRow();
                foreach (DictionaryEntry item in mapData)
                {
                    if ((string)item.Key == "대비율")
                        row[(string)item.Key] = string.Format("{0:f2}", item.Value);
                    else
                        row[(string)item.Key] = string.Format("{0:N0}", item.Value);
                }

                m_volTable.Rows.Add(row);
            }
            m_VolTopGrid.Refresh();

            return true;
        }



        private void m_rdoV1_Click(object sender, EventArgs e)
        {
            m_btnRetry_Click(null, null);
        }

        private void m_rdoV2_Click(object sender, EventArgs e)
        {
            m_btnRetry_Click(null, null);
        }

        private void m_chkExc1_Click(object sender, EventArgs e)
        {
            m_btnRetry_Click(null, null);
        }

        private void m_chkExc2_Click(object sender, EventArgs e)
        {
            m_btnRetry_Click(null, null);
        }

        private void m_rdoMarketAll_Click(object sender, EventArgs e)
        {
            m_btnRetry_Click(null, null);
        }

        private void m_rdoMarket1_Click(object sender, EventArgs e)
        {
            m_btnRetry_Click(null, null);
        }

        private void m_rdoMarket2_Click(object sender, EventArgs e)
        {
            m_btnRetry_Click(null, null);
        }

        private void m_btnRetry_Click(object sender, EventArgs e)
        {
            InitGridTable();
            RqVolTop();
        }
    } // end of fmVolTop
}
