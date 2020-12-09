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
    public partial class fmSFuture : Form
    {
        public class CMstData
        {
            //public int nTime;
            public int nTime;
            public int nOpen;
            public int nHigh;
            public int nLow;
            public int nClose;
            public int nBase;
            public int nDiff;
            public double dDiffR;

            public uint uVol;
            public uint uNoCont;        // 미결제약정
            public int nLogPrice;       // 이론가



            public int nBaseCur;
            public int nBaseDiff;
            public int nBasePrev;
            public int nBaseDiffR;
            public int nBasis;
            public string sBaseCode;


            public int nOfferA;
            public int nBidA;

            public int[] nOfferP = new int[10];
            public int[] nOfferR = new int[10];

            public int[] nBidP = new int[10];
            public int[] nBidR = new int[10];
        }

        // 컬럼 인덱스 
        const int COL_SELLHOGA = 0;
        const int COL_PRICE = 1;
        const int COL_BUYHOGA = 2;

        const int ROW_SEL_START = 0;
        const int ROW_BUY_START = 11;
        const int ROW_CURPRICE = 10;
        const int ROW_ALL_HOGA = 21;
        const int NUM_HOGA = 10;



        enum PB_TYPE { ALL, CUR, BASE }

        private DataTable m_DataTable;
        private DataTable m_infoTable;
        private DataTable m_weekTable;

        private CMstData m_mstData = new CMstData();

        private Dictionary<string, string> m_mapFCodes = new Dictionary<string, string>();

        // 실시간 SB
        private CPSYSDIBLib.FutStockCurS m_cpSFutureCur;
        private DSCBO1Lib.StockCur m_cpStockCur;
        private CPUTILLib.CpCodeMgr m_cpCodeMgr = new CPUTILLib.CpCodeMgr();

        private string m_sCode = "";

        public fmSFuture()
        {
            InitializeComponent();

            LoadFutureCode();

            // 현재가 실시간 객체 색성, 이벤트 함수 연동 
            m_cpSFutureCur = new CPSYSDIBLib.FutStockCurS();
            m_cpSFutureCur.Received += new CPSYSDIBLib._ISysDibEvents_ReceivedEventHandler(m_cpSFutureCur_Received);

            // 기초자산 요청용
            m_cpStockCur = new DSCBO1Lib.StockCur();
            m_cpStockCur.Received += new DSCBO1Lib._IDibEvents_ReceivedEventHandler(m_cpStockCur_Received);

        }

        private void LoadFutureCode()
        {
            object[] codes = m_cpCodeMgr.GetStockFutureList();

            foreach (object code in codes)
            {
                string sCode = (string)code;
                string sName = m_cpCodeMgr.CodeToName(sCode);
                m_mapFCodes.Add(sCode, sName);
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

        void requestSFuture(string sCode)
        {
            if (m_sCode != "")
                SbCancel(m_sCode);


            InitGridTable();

            RqSFutureMst(sCode);
            RqSFutureWeek(sCode);
            setDataToGrid(PB_TYPE.ALL);

            SBSFutureInfos(sCode);
        }

        void InitGridTable()
        {
            if (m_DataTable != null)
                m_DataTable.Clear();

            if (m_infoTable != null)
                m_infoTable.Clear();

            if (m_weekTable != null)
                m_weekTable.Clear();

            m_DataTable = new DataTable();
            m_infoTable = new DataTable();
            m_weekTable = new DataTable();

            //////////////////////////////////////////////////////
            // 기초자산 그리드 
            m_infoTable.Columns.Add("기초자산");
            m_infoTable.Columns.Add("대비");
            m_infoTable.Columns.Add("Basis");
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
            // 일자별 그리드 
            m_weekTable.Columns.Add("날짜");
            m_weekTable.Columns.Add("시가");
            m_weekTable.Columns.Add("고가");
            m_weekTable.Columns.Add("저가");
            m_weekTable.Columns.Add("종가");
            m_weekTable.Columns.Add("대비");
            m_weekTable.Columns.Add("거래량");
            m_weekTable.Columns.Add("이론선물");
            m_weekTable.Columns.Add("미결제약정");
            m_weekTable.Columns.Add("베이시스");

            m_weekGrid.DataSource = m_weekTable;
            m_weekGrid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            m_weekGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            m_weekGrid.AllowUserToResizeRows = false;
            m_weekGrid.RowsDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            m_weekGrid.Columns[0].Width = 50;
            m_weekGrid.Font = new Font("Tahoma", 9.75F, FontStyle.Regular);
            m_weekGrid.RowHeadersVisible = false;   // 왼쪽 커서 제거 
            //m_weekGrid.Rows[0].Cells[0].Selected = false;   // 첫 컬럼이 선택되지 않도록 
            m_weekGrid.AllowUserToAddRows = false;  // 사용자 행 추가 못하도록 
            m_weekGrid.Refresh();


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
        private bool RqSFutureMst(string sCode)
        {
            CPSYSDIBLib.FutStockMst objRq = new CPSYSDIBLib.FutStockMst();
            objRq.SetInputValue(0, sCode);

            CpCommon.waitRqLimit(COMM_TYPE.SISE);
            int nRet = objRq.BlockRequest();
            if (false == CpCommon.checkRqError(objRq, nRet))
                return false;



            m_mstData.nOpen = objRq.GetHeaderValue(7);
            m_mstData.nHigh = objRq.GetHeaderValue(8);
            m_mstData.nLow = objRq.GetHeaderValue(9);
            m_mstData.nClose = objRq.GetHeaderValue(2);
            m_mstData.nBaseCur = objRq.GetHeaderValue(4);
            m_mstData.nBaseDiff = objRq.GetHeaderValue(5);
            m_mstData.nBasePrev = m_mstData.nBaseCur - m_mstData.nBaseDiff;

            m_mstData.nBasis = objRq.GetHeaderValue(6);
            m_mstData.nDiff = objRq.GetHeaderValue(3);
            m_mstData.uVol = objRq.GetHeaderValue(10);
            m_mstData.uNoCont = objRq.GetHeaderValue(12);
            m_mstData.nLogPrice = objRq.GetHeaderValue(14);

            m_mstData.nBase = m_mstData.nClose - m_mstData.nDiff;
            m_mstData.dDiffR = 0;
            if (m_mstData.nBase > 0)
            {
                m_mstData.dDiffR = (double)(m_mstData.nDiff / (double)m_mstData.nBase) * 100;
            }


            int nOP = 25;
            for (int i = 0; i < NUM_HOGA; i++)
            {
                m_mstData.nOfferP[i] = objRq.GetHeaderValue(nOP++);
                m_mstData.nBidP[i] = objRq.GetHeaderValue(nOP++);
                m_mstData.nOfferR[i] = objRq.GetHeaderValue(nOP++);
                m_mstData.nBidR[i] = objRq.GetHeaderValue(nOP++);
            }

            m_mstData.nOfferA = objRq.GetHeaderValue(85);
            m_mstData.nBidA = objRq.GetHeaderValue(86);
            m_mstData.sBaseCode = objRq.GetHeaderValue(91);


            return true;
        }

        // 주식선물 6주간 테스트 
        private bool RqSFutureWeek(string sCode)
        {
            CPSYSDIBLib.FutStockWeek objRq = new CPSYSDIBLib.FutStockWeek();
            objRq.SetInputValue(0, sCode);


            while (true)
            {
                CpCommon.waitRqLimit(COMM_TYPE.SISE);
                int nRet = objRq.BlockRequest();
                if (false == CpCommon.checkRqError(objRq, nRet))
                    return false;

                int nCnt = objRq.GetHeaderValue(0);
                Hashtable mapWeek = new Hashtable();

                for (int i = 0; i < nCnt; i++)
                {
                    int nDate = (int)objRq.GetDataValue(0, i);
                    mapWeek["시가"] = objRq.GetDataValue(1, i);
                    mapWeek["고가"] = objRq.GetDataValue(2, i);
                    mapWeek["저가"] = objRq.GetDataValue(3, i);
                    mapWeek["종가"] = objRq.GetDataValue(4, i);
                    mapWeek["대비"] = objRq.GetDataValue(5, i);
                    mapWeek["거래량"] = objRq.GetDataValue(6, i);
                    mapWeek["이론선물"] = objRq.GetDataValue(7, i);
                    mapWeek["미결제약정"] = objRq.GetDataValue(8, i);
                    mapWeek["베이시스"] = objRq.GetDataValue(9, i);

                    DataRow row = m_weekTable.NewRow();
                    nDate = nDate % 10000;
                    int nMon = nDate / 100;
                    int nDay = nDate % 100;
                    row["날짜"] = string.Format("{0:D2}/{1:D2}", nMon, nDay);
                    foreach (DictionaryEntry item in mapWeek)
                    {
                        row[(string)item.Key] = string.Format("{0:N0}", item.Value);
                    }

                    m_weekTable.Rows.Add(row);
                }

                if (objRq.Continue == 0)
                    break;

            }


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

        private void setDataToWeekGrid()
        {
            m_weekTable.Rows[0]["시가"] = string.Format("{0:N0}", m_mstData.nOpen);
            m_weekTable.Rows[0]["고가"] = string.Format("{0:N0}", m_mstData.nHigh);
            m_weekTable.Rows[0]["저가"] = string.Format("{0:N0}", m_mstData.nLow);
            m_weekTable.Rows[0]["종가"] = string.Format("{0:N0}", m_mstData.nClose);
            m_weekTable.Rows[0]["대비"] = string.Format("{0:N0}", m_mstData.nDiff);
            m_weekTable.Rows[0]["거래량"] = string.Format("{0:N0}", m_mstData.uVol);
            m_weekTable.Rows[0]["이론선물"] = string.Format("{0:N0}", m_mstData.nLogPrice);
            m_weekTable.Rows[0]["미결제약정"] = string.Format("{0:N0}", m_mstData.uNoCont);
            m_weekTable.Rows[0]["베이시스"] = string.Format("{0:N0}", m_mstData.nBasis);

        }
        private void setDataToGrid(PB_TYPE pbtype)
        {
            if (pbtype == PB_TYPE.ALL || pbtype == PB_TYPE.BASE)
            {
                // 기초자산 채우기 
                m_infoTable.Rows[0]["기초자산"] = string.Format("{0:N0}", m_mstData.nBaseCur);
                m_infoTable.Rows[0]["대비"] = string.Format("{0:N0}", m_mstData.nBaseDiff);
                m_infoTable.Rows[0]["Basis"] = string.Format("{0:N0}", m_mstData.nBasis);

                Color clrDiff = getDiffColor(m_mstData.nBaseDiff);
                m_infoGrid[0, 0].Style.ForeColor = clrDiff;
                m_infoGrid[1, 0].Style.ForeColor = clrDiff;
                clrDiff = getDiffColor(m_mstData.nBasis);
                m_infoGrid[2, 0].Style.ForeColor = clrDiff;

            }

            if (pbtype == PB_TYPE.ALL || pbtype == PB_TYPE.CUR)
            {
                // 현재가 채우기 
                m_DataTable.Rows[ROW_CURPRICE]["매도잔량"] = "현재가";
                m_DataTable.Rows[ROW_CURPRICE]["호가"] = string.Format("{0:N0}", m_mstData.nClose);
                string sDiff = string.Format("{0:N0}", m_mstData.nDiff);
                m_DataTable.Rows[ROW_CURPRICE]["매수잔량"] = sDiff;
                sDiff = string.Format("{0:f2}%", m_mstData.dDiffR);
                m_DataTable.Rows[ROW_CURPRICE - 1]["매수잔량"] = sDiff;

                // 현재가, 대비 색상처리 
                Color clrDiff = getDiffColor(m_mstData.nDiff);

                m_HogaView[COL_PRICE, ROW_CURPRICE].Style.ForeColor = clrDiff;
                m_HogaView[COL_BUYHOGA, ROW_CURPRICE].Style.ForeColor = clrDiff;
                m_HogaView[COL_BUYHOGA, ROW_CURPRICE - 1].Style.ForeColor = clrDiff;
            }

            // 매도잔량 | 호가 채우기 
            for (int i = 0; i < NUM_HOGA; i++)
            {
                int nIndex = NUM_HOGA - i - 1;
                m_DataTable.Rows[ROW_SEL_START + i]["매도잔량"] = string.Format("{0:N0}", m_mstData.nOfferR[nIndex]);
                m_DataTable.Rows[ROW_SEL_START + i]["호가"] = string.Format("{0:N0}", m_mstData.nOfferP[nIndex]);
            }
            // 매수잔량 | 호가 채우기 
            for (int i = 0; i < NUM_HOGA; i++)
            {
                m_DataTable.Rows[ROW_BUY_START + i]["호가"] = string.Format("{0:N0}", m_mstData.nBidP[i]);
                m_DataTable.Rows[ROW_BUY_START + i]["매수잔량"] = string.Format("{0:N0}", m_mstData.nBidR[i]);
            }


            m_DataTable.Rows[ROW_ALL_HOGA]["매도잔량"] = string.Format("{0:N0}", m_mstData.nOfferA);
            m_DataTable.Rows[ROW_ALL_HOGA]["매수잔량"] = string.Format("{0:N0}", m_mstData.nBidA);
        }


        //////////////////////////////////////////////////////////////////////
        /// 실시간 등록 처리 
        //////////////////////////////////////////////////////////////////////

        // 실시간 시세 및 체결 등록
        private void SBSFutureInfos(string sCode)
        {
            // 이전 sb cancel

            // 선물 현재가 실시간 등록
            m_cpSFutureCur.SetInputValue(0, sCode);
            m_cpSFutureCur.Subscribe();

            // 선물 5차 호가 실시간 등록
            // SubscribeLatest 는 최근 값만 가져온다 --> 성능 개선 효과 
            m_cpSFutureCur.SetInputValue(0, sCode);
            m_cpSFutureCur.SubscribeLatest();
            m_sCode = sCode;

            // 기초자산 현재가 실시간 등록
            m_cpStockCur.SetInputValue(0, m_mstData.sBaseCode);
            m_cpStockCur.Subscribe();

            Debug.WriteLine("SB START CODE {0}", sCode);

        }

        // 실시간 시세 및 체결 등록 해지 
        private void SbCancel(string sCode)
        {
            if (sCode == "")
                return;

            m_cpSFutureCur.SetInputValue(0, sCode);
            m_cpSFutureCur.Unsubscribe();

            m_cpSFutureCur.SetInputValue(0, sCode);
            m_cpSFutureCur.Unsubscribe();

            m_cpStockCur.SetInputValue(0, m_mstData.sBaseCode);
            m_cpStockCur.Unsubscribe();
        }

        //////////////////////////////////////////////////////////////////////
        /// PLUS 실시간 이벤트 처리 
        //////////////////////////////////////////////////////////////////////

        // 실시간 호가 수신
        //실시간 시세체결 수신
        private void m_cpSFutureCur_Received()
        {
            //Debug.WriteLine("현재가 체결 PB 수신");

            //m_mstData.nTime = m_cpSFutureCur.GetHeaderValue(3);
            m_mstData.nOpen = m_cpSFutureCur.GetHeaderValue(6);
            m_mstData.nHigh = m_cpSFutureCur.GetHeaderValue(7);
            m_mstData.nLow = m_cpSFutureCur.GetHeaderValue(8);

            m_mstData.nClose = m_cpSFutureCur.GetHeaderValue(1);
            m_mstData.nDiff = m_cpSFutureCur.GetHeaderValue(2);

            m_mstData.nBaseCur = m_cpSFutureCur.GetHeaderValue(15);
            m_mstData.nBasis = m_cpSFutureCur.GetHeaderValue(4);
            m_mstData.nBaseDiff = m_mstData.nBaseCur - m_mstData.nBasePrev;


            m_mstData.nBase = m_mstData.nClose - m_mstData.nDiff;
            m_mstData.dDiffR = 0;
            if (m_mstData.nBase > 0)
            {
                m_mstData.dDiffR = (double)(m_mstData.nDiff / (double)m_mstData.nBase) * 100;
            }

            int nOP = 16;
            for (int i = 0; i < NUM_HOGA; i++)
            {
                m_mstData.nOfferP[i] = m_cpSFutureCur.GetHeaderValue(nOP++);
                m_mstData.nBidP[i] = m_cpSFutureCur.GetHeaderValue(nOP++);
                m_mstData.nOfferR[i] = m_cpSFutureCur.GetHeaderValue(nOP++);
                m_mstData.nBidR[i] = m_cpSFutureCur.GetHeaderValue(nOP++);
            }

            m_mstData.nOfferA = m_cpSFutureCur.GetHeaderValue(56);
            m_mstData.nBidA = m_cpSFutureCur.GetHeaderValue(57);

            m_mstData.uVol = m_cpSFutureCur.GetHeaderValue(12);
            m_mstData.uNoCont = m_cpSFutureCur.GetHeaderValue(13);
            m_mstData.nLogPrice = m_cpSFutureCur.GetHeaderValue(3);


            setDataToGrid(PB_TYPE.CUR);
            setDataToGrid(PB_TYPE.BASE);
            setDataToWeekGrid();
        }

        private void m_cpStockCur_Received()
        {

            m_mstData.nBaseCur = m_cpStockCur.GetHeaderValue(13);
            m_mstData.nBaseDiff = m_cpStockCur.GetHeaderValue(2);
            m_mstData.nBasis = m_mstData.nClose - m_mstData.nBaseCur;

            setDataToGrid(PB_TYPE.BASE);
        }




        private void m_cbCode_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sCodes = m_cbCode.SelectedItem.ToString();
            string[] splits = sCodes.Split(' ');

            requestSFuture(splits[0]);
        }

        private void fmSFuture_FormClosed(object sender, FormClosedEventArgs e)
        {
            SbCancel(m_sCode);
        }




    } // end of fmSFuture
}
