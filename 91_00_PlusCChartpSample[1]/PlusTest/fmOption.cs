using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PlusTest
{
    public partial class fmOption : Form
    {
        public class CMstData
        {
            //public int nTime;
            public int nMarketStatus;
            public double dOpen;
            public double dHigh;
            public double dLow;
            public double dClose;
            public double dBase;
            public double dDiff;
            public double dDiffR;
            public double dK200;
            public double dK200Prev;
            public double dK200Diff;
            public double dK200DiffR;
            public string sBaseCode;


            public int nOfferA;
            public int nBidA;

            public double[] dOfferP = new double[5];
            public int[] nOfferR = new int[5];

            public double[] dBidP = new double[5];
            public int[] nBidR = new int[5];
        }

        // 컬럼 인덱스 
        const int COL_SELLHOGA = 0;
        const int COL_PRICE = 1;
        const int COL_BUYHOGA = 2;

        const int ROW_SEL_START = 0;
        const int ROW_BUY_START = 6;
        const int ROW_CURPRICE = 5;
        const int ROW_ALL_HOGA = 11;
        const int NUM_HOGA = 5;



        enum PB_TYPE { ALL, CUR, HOGA,K200}

        private DataTable m_DataTable;
        private DataTable m_infoTable;

        private CMstData m_mstData = new CMstData();


        private Dictionary<string, string> m_mapOCodes = new Dictionary<string, string>();



        // 실시간 SB
        private CPSYSDIBLib.OptionCurOnly m_cpOptionCur;
        private CPSYSDIBLib.OptionJpBid m_cpOptionBid;
        private DSCBO1Lib.StockIndexis m_cpUpjongCur;

        private string m_sCode = "";

        public fmOption()
        {
            InitializeComponent();

            LoadOptionCode();

            // 현재가 실시간 객체 색성, 이벤트 함수 연동 
            m_cpOptionCur = new CPSYSDIBLib.OptionCurOnly();
            m_cpOptionCur.Received += new CPSYSDIBLib._ISysDibEvents_ReceivedEventHandler(m_cpOptionCur_Received);

            m_cpOptionBid = new CPSYSDIBLib.OptionJpBid();
            m_cpOptionBid.Received += new CPSYSDIBLib._ISysDibEvents_ReceivedEventHandler(m_cpOptionBid_Received);


            m_cpUpjongCur = new DSCBO1Lib.StockIndexis();
            m_cpUpjongCur.Received += new DSCBO1Lib._IDibEvents_ReceivedEventHandler(m_cpUpjongCur_Received);

        }

        private void LoadOptionCode()
        {
            CPUTILLib.CpOptionCode cpCodes = new CPUTILLib.CpOptionCode();

            int nCnt = cpCodes.GetCount();
            for (int i = 0; i < nCnt; i++)
            {
                string sCode = cpCodes.GetData(0, (short)i);
                string sName = cpCodes.GetData(1, (short)i);

                m_mapOCodes.Add(sCode, sName);
                m_cbCode.Items.Add(sCode + " " + sName);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (false == CpCommon.checkPlusStatus())
            {
                MessageBox.Show("PLUS 실행 상태 점검 필요", "PLUS 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            m_cbCode.SelectedIndex = 0;

        }

        void requestMst(string sCode)
        {
            if (m_sCode != "")
                SbCancel(m_sCode);


            InitGridTable();

            RqOptionMst(sCode);
            setDataToGrid(PB_TYPE.ALL);

            SBInfos(sCode);
        }

        void InitGridTable()
        {
            if (m_DataTable != null)
                m_DataTable.Clear();

            if (m_infoTable != null)
                m_infoTable.Clear();

            m_DataTable = new DataTable();
            m_infoTable = new DataTable();

            //////////////////////////////////////////////////////
            // 기초자산 그리드 
            m_infoTable.Columns.Add("K200");
            m_infoTable.Columns.Add("대비");
            m_infoTable.Columns.Add("대비율");
            m_infoTable.Rows.Add("", "", "");

            m_infoGrid.DataSource = m_infoTable;
            m_infoGrid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            m_infoGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            m_infoGrid.AllowUserToResizeRows = false;
            m_infoGrid.RowsDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            m_infoGrid.Font = new Font("Tahoma", 9.75F, FontStyle.Regular);
            m_infoGrid.RowHeadersVisible = false;   // 왼쪽 커서 제거 
            m_infoGrid.Rows[0].Cells[0].Selected = false;   // 첫 컬럼이 선택되지 않도록 
            m_infoGrid.AllowUserToAddRows = false;  // 사용자 행 추가 못하도록 
            m_infoGrid.Refresh();



            //////////////////////////////////////////////////////
            // 5차 호가 그리드
            m_DataTable.Columns.Add("매도잔량");
            m_DataTable.Columns.Add("호가");
            m_DataTable.Columns.Add("매수잔량");

            for (int i = 0; i < 22; i++)
                m_DataTable.Rows.Add("", "", "");


            m_HogaView.DataSource = m_DataTable;
            m_HogaView.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            m_HogaView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            m_HogaView.AllowUserToResizeRows = false;
            //m_HogaView.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            m_HogaView.RowsDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;


            // 배경색 칠하기 - 매도잔량
            for (int i = ROW_SEL_START; i < ROW_SEL_START + NUM_HOGA; i++)
                m_HogaView[COL_SELLHOGA, i].Style.BackColor = Color.FromArgb(215, 233, 255);


            // 배경색 칠하기 - 매도호가
            for (int i = ROW_SEL_START; i < ROW_SEL_START + NUM_HOGA; i++)
                m_HogaView[COL_PRICE, i].Style.BackColor = Color.FromArgb(215, 233, 255);



            // 배경색 칠하기 - 매수잔량
            for (int i = ROW_BUY_START; i < ROW_BUY_START + NUM_HOGA; i++)
                m_HogaView[COL_BUYHOGA, i].Style.BackColor = Color.FromArgb(255, 227, 227);



            // 배경색 칠하기 - 매수호가
            for (int i = ROW_BUY_START; i < ROW_BUY_START + NUM_HOGA; i++)
                m_HogaView[COL_PRICE, i].Style.BackColor = Color.FromArgb(255, 227, 227);


            // 배경색 칠하기 - 현재가
            m_HogaView[COL_SELLHOGA, ROW_CURPRICE].Style.BackColor = Color.White;
            m_HogaView[COL_PRICE, ROW_CURPRICE].Style.BackColor = Color.White;
            m_HogaView[COL_BUYHOGA, ROW_CURPRICE].Style.BackColor = Color.White;

            m_HogaView.Font = new Font("Tahoma", 9.75F, FontStyle.Regular);
            
            m_HogaView.RowHeadersVisible = false;       //  왼쪽 커서 제거 
            m_HogaView.Rows[0].Cells[0].Selected = false; // 첫 컬럼이 선택되지 않도록 
            m_HogaView.AllowUserToAddRows = false;  // 사용자 행 추가 못하도록 
            m_HogaView.Refresh();



            // 인터넷에서 퍼옴 - 그리드 깜빡임 제거 
            Type dgvType = m_HogaView.GetType();
            PropertyInfo pi = dgvType.GetProperty("DoubleBuffered",
                  BindingFlags.Instance | BindingFlags.NonPublic);
            pi.SetValue(m_HogaView, true, null);
        }

        // 선물 현재가 조회
        private bool RqOptionMst(string sCode)
        {
            DSCBO1Lib.OptionMst objRq = new DSCBO1Lib.OptionMst();
            objRq.SetInputValue(0, sCode);

            CpCommon.waitRqLimit(COMM_TYPE.SISE);
            int nRet = objRq.BlockRequest();
            if (false == CpCommon.checkRqError(objRq, nRet))
                return false;

            
            m_mstData.dOpen = objRq.GetHeaderValue(94);
            m_mstData.dHigh = objRq.GetHeaderValue(95);
            m_mstData.dLow = objRq.GetHeaderValue(96);
            m_mstData.dClose = objRq.GetHeaderValue(93);

            m_mstData.dK200 = objRq.GetHeaderValue(106);
            m_mstData.dK200Diff = objRq.GetHeaderValue(107);
            m_mstData.dK200Prev = m_mstData.dK200 - m_mstData.dK200Diff;
            m_mstData.dK200DiffR = m_mstData.dK200Diff / m_mstData.dK200Prev * 100;


            m_mstData.dDiff = objRq.GetHeaderValue(77);
            m_mstData.dBase = objRq.GetHeaderValue(51);
            m_mstData.dDiff = m_mstData.dClose - m_mstData.dBase;
            if (m_mstData.dBase > 0)
            {
                m_mstData.dDiffR = (double)(m_mstData.dDiff / m_mstData.dBase) * 100;
            }


            int nOP = 58;
            for (int i = 0; i < 5; i++)
            {
                m_mstData.dOfferP[i] = objRq.GetHeaderValue(nOP++);
                m_mstData.dBidP[i] = objRq.GetHeaderValue(nOP++);
                m_mstData.nOfferR[i] = objRq.GetHeaderValue(nOP++);
                m_mstData.nBidR[i] = objRq.GetHeaderValue(nOP++);
            }

            m_mstData.nOfferA = objRq.GetHeaderValue(78);
            m_mstData.nBidA = objRq.GetHeaderValue(79);
            m_mstData.sBaseCode = "U180";
            m_mstData.nMarketStatus = objRq.GetHeaderValue(118);

            return true;
        }

        Color getDiffColor(double dDiff)
        {
            if (dDiff > 0)
                return Color.Red;
            if (dDiff < 0)
                return Color.Blue;
            return Color.Black;
        }
        private void setDataToGrid(PB_TYPE pbtype)
        {
            if (pbtype == PB_TYPE.ALL || pbtype == PB_TYPE.K200)
            {
                // 기초자산 채우기 
                m_infoTable.Rows[0]["K200"] = string.Format("{0:f2}", m_mstData.dK200);
                m_infoTable.Rows[0]["대비"] = string.Format("{0:f2}", m_mstData.dK200Diff);
                m_infoTable.Rows[0]["대비율"] = string.Format("{0:f2}%", m_mstData.dK200DiffR);
                Color clrDiff = getDiffColor(m_mstData.dK200Diff);
                m_infoGrid[0, 0].Style.ForeColor = clrDiff;
                m_infoGrid[1, 0].Style.ForeColor = clrDiff;
                m_infoGrid[2, 0].Style.ForeColor = clrDiff;
            }
            if (pbtype == PB_TYPE.ALL || pbtype == PB_TYPE.HOGA)
            {
                // 매도잔량 | 호가 채우기 
                for (int i = 0; i < NUM_HOGA; i++)
                {
                    int nIndex = NUM_HOGA - i - 1;
                    m_DataTable.Rows[ROW_SEL_START + i]["매도잔량"] = string.Format("{0:N0}", m_mstData.nOfferR[nIndex]);
                    m_DataTable.Rows[ROW_SEL_START + i]["호가"] = string.Format("{0:f2}", m_mstData.dOfferP[nIndex]);
                }
                // 매수잔량 | 호가 채우기 
                for (int i = 0; i < NUM_HOGA; i++)
                {
                    //string sKey = (i + 1).ToString();
                    //m_DataTable.Rows[ROW_BUY_START + i]["호가"] = string.Format("{0:N0}", m_mapMst[sKey + "_매수호가"]);
                    //m_DataTable.Rows[ROW_BUY_START + i]["매수잔량"] = string.Format("{0:N0}", m_mapMst[sKey + "_매수잔량"]);
                    m_DataTable.Rows[ROW_BUY_START + i]["호가"] = string.Format("{0:f2}", m_mstData.dBidP[i]);
                    m_DataTable.Rows[ROW_BUY_START + i]["매수잔량"] = string.Format("{0:N0}", m_mstData.nBidR[i]);
                }


                m_DataTable.Rows[ROW_ALL_HOGA]["매도잔량"] = string.Format("{0:N0}", m_mstData.nOfferA);
                m_DataTable.Rows[ROW_ALL_HOGA]["매수잔량"] = string.Format("{0:N0}", m_mstData.nBidA);

            }

            if (pbtype == PB_TYPE.ALL || pbtype == PB_TYPE.CUR)
            {
                // 현재가 채우기 
                m_DataTable.Rows[ROW_CURPRICE]["매도잔량"] = "현재가";
                m_DataTable.Rows[ROW_CURPRICE]["호가"] = string.Format("{0:f2}", m_mstData.dClose);
                string sDiff = string.Format("{0:f2}", m_mstData.dDiff);
                m_DataTable.Rows[ROW_CURPRICE]["매수잔량"] = sDiff;
                sDiff = string.Format("{0:f2}%", m_mstData.dDiffR);
                m_DataTable.Rows[ROW_CURPRICE - 1]["매수잔량"] = sDiff;

                // 현재가, 대비 색상처리 
                Color clrDiff = getDiffColor(m_mstData.dDiff);

                m_HogaView[COL_PRICE, ROW_CURPRICE].Style.ForeColor = clrDiff;
                m_HogaView[COL_BUYHOGA, ROW_CURPRICE].Style.ForeColor = clrDiff;
                m_HogaView[COL_BUYHOGA, ROW_CURPRICE - 1].Style.ForeColor = clrDiff;
            }

        }


        //////////////////////////////////////////////////////////////////////
        /// 실시간 등록 처리 
        //////////////////////////////////////////////////////////////////////

        // 실시간 시세 및 체결 등록
        private void SBInfos(string sCode)
        {
            // 이전 sb cancel

            // 선물 현재가 실시간 등록
            m_cpOptionCur.SetInputValue(0, sCode);
            m_cpOptionCur.Subscribe();

            // 선물 5차 호가 실시간 등록
            // SubscribeLatest 는 최근 값만 가져온다 --> 성능 개선 효과 
            m_cpOptionBid.SetInputValue(0, sCode);
            m_cpOptionBid.SubscribeLatest();
            m_sCode = sCode;

            m_cpUpjongCur.SetInputValue(0, m_mstData.sBaseCode);
            m_cpUpjongCur.Subscribe();

            Debug.WriteLine("SB START CODE {0}", sCode);

        }

        // 실시간 시세 및 체결 등록 해지 
        private void SbCancel(string sCode)
        {
            if (sCode == "")
                return;

            m_cpOptionCur.SetInputValue(0, sCode);
            m_cpOptionCur.Unsubscribe();

            m_cpOptionBid.SetInputValue(0, sCode);
            m_cpOptionBid.Unsubscribe();

            m_cpUpjongCur.SetInputValue(0, m_mstData.sBaseCode);
            m_cpUpjongCur.Unsubscribe();
        }

        //////////////////////////////////////////////////////////////////////
        /// PLUS 실시간 이벤트 처리 
        //////////////////////////////////////////////////////////////////////

        // 실시간 호가 수신
        private void m_cpOptionBid_Received()
        {
            int nOP = 2;
            int nOPR = 7;
            int nBP = 19;
            int nBPR = 24;
            for (int i = 0; i < 5; i++)
            {
                m_mstData.dOfferP[i] = m_cpOptionBid.GetHeaderValue(nOP++);
                m_mstData.dBidP[i] = m_cpOptionBid.GetHeaderValue(nBP++);
                m_mstData.nOfferR[i] = m_cpOptionBid.GetHeaderValue(nOPR++);
                m_mstData.nBidR[i] = m_cpOptionBid.GetHeaderValue(nBPR++);
            }
                
            m_mstData.nOfferA = m_cpOptionBid.GetHeaderValue(12);
            m_mstData.nBidA = m_cpOptionBid.GetHeaderValue(29);


            setDataToGrid(PB_TYPE.HOGA);

        }
        //실시간 시세체결 수신
        private void m_cpOptionCur_Received()
        {
            //Debug.WriteLine("현재가 체결 PB 수신");

            //m_mstData.nTime = m_cpOptionCur.GetHeaderValue(3);
            m_mstData.dOpen = m_cpOptionCur.GetHeaderValue(4);
            m_mstData.dHigh = m_cpOptionCur.GetHeaderValue(5);
            m_mstData.dLow = m_cpOptionCur.GetHeaderValue(6);

            m_mstData.dClose = m_cpOptionCur.GetHeaderValue(2);
            m_mstData.dDiff = m_cpOptionCur.GetHeaderValue(3);

            m_mstData.dBase = m_mstData.dClose - m_mstData.dDiff;
            m_mstData.dDiffR = 0;
            if (m_mstData.dBase > 0)
            {
                m_mstData.dDiffR = (double)(m_mstData.dDiff / m_mstData.dBase) * 100;
            }

            setDataToGrid(PB_TYPE.CUR);
        }

        private void m_cpUpjongCur_Received()
        {

            m_mstData.dK200 = m_cpUpjongCur.GetHeaderValue(2);
            m_mstData.dK200Diff = m_cpUpjongCur.GetHeaderValue(3);
            m_mstData.dK200DiffR = m_mstData.dK200Diff / m_mstData.dK200Prev * 100;
            setDataToGrid(PB_TYPE.K200);
        }

        private void m_cbCode_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sCodes = m_cbCode.SelectedItem.ToString();
            string[] splits = sCodes.Split(' ');

            requestMst(splits[0]);
        }

        private void fmOption_FormClosed(object sender, FormClosedEventArgs e)
        {
            SbCancel(m_sCode);
        }


    } // end of fmOption
}
