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
    public partial class fmStock : Form
    {
        public class CMstData
        {
            public int nTime;
            public int nOpen;
            public int nHigh;
            public int nLow;
            public int nClose;
            public int nBase;
            public int nDiff;
            public double dDiffR;

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


        enum PB_TYPE { ALL, CUR, HOGA}

        private DataTable m_DataTable;
        //private Hashtable m_mapMst = new Hashtable();
        private CMstData m_mstData = new CMstData();



        // 실시간 SB
        private DSCBO1Lib.StockCur m_cpStockCur;
        private DSCBO1Lib.StockJpbid m_cpStockJpBid;
        private CPUTILLib.CpCodeMgr m_cpCodeMgr = new CPUTILLib.CpCodeMgr();

        private string m_sCode = "";

        public fmStock()
        {
            InitializeComponent();

            // 현재가 실시간 객체 색성, 이벤트 함수 연동 
            m_cpStockCur = new DSCBO1Lib.StockCur();
            m_cpStockCur.Received += new DSCBO1Lib._IDibEvents_ReceivedEventHandler(m_cpStockCur_Received);

            m_cpStockJpBid = new DSCBO1Lib.StockJpbid();
            m_cpStockJpBid.Received += new DSCBO1Lib._IDibEvents_ReceivedEventHandler(m_cpStockJpBid_Received);

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (false == CpCommon.checkPlusStatus())
            {
                MessageBox.Show("PLUS 실행 상태 점검 필요", "PLUS 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


            m_etCode.Text = "005930";
            reqeustStock("A005930");
        }

        void reqeustStock(string sCode)
        {
            if (m_sCode != "")
                SbCancel(m_sCode);


            InitGridTable();

            RqStockMst(sCode);
            setDataToGrid(PB_TYPE.ALL);

            SBStockInfos(sCode);
        }

        void InitGridTable()
        {
            if (m_DataTable != null)
                m_DataTable.Clear();


            m_DataTable = new DataTable();



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
            // 왼쪽 커서 제거 
            m_HogaView.RowHeadersVisible = false;

            m_HogaView.Refresh();



            // 인터넷에서 퍼옴 - 그리드 깜빡임 제거 
            Type dgvType = m_HogaView.GetType();
            PropertyInfo pi = dgvType.GetProperty("DoubleBuffered",
                  BindingFlags.Instance | BindingFlags.NonPublic);
            pi.SetValue(m_HogaView, true, null);
        }

        // 주식 현재가 조회
        private bool RqStockMst(string sCode)
        {
            DSCBO1Lib.StockMst objRq = new DSCBO1Lib.StockMst();
            objRq.SetInputValue(0, sCode);

            CpCommon.waitRqLimit(COMM_TYPE.SISE);
            int nRet = objRq.BlockRequest();
            if (false == CpCommon.checkRqError(objRq, nRet))
                return false;



            m_mstData.nTime = objRq.GetHeaderValue(4);
            m_mstData.nOpen = objRq.GetHeaderValue(13);
            m_mstData.nHigh = objRq.GetHeaderValue(14);
            m_mstData.nLow = objRq.GetHeaderValue(15);
            m_mstData.nClose = objRq.GetHeaderValue(11);
            m_mstData.nDiff = objRq.GetHeaderValue(12);


            m_mstData.nBase = m_mstData.nClose - m_mstData.nDiff;
            m_mstData.dDiffR = 0;
            if (m_mstData.nBase > 0)
            {
                m_mstData.dDiffR = (float)(m_mstData.nDiff / (float)m_mstData.nBase) * 100;
            }



            for (int i = 0; i < NUM_HOGA; i++)
            {
                string sKey = (i + 1).ToString();



                m_mstData.nOfferP[i] = objRq.GetDataValue(0, i);
                m_mstData.nBidP[i] = objRq.GetDataValue(1, i);
                m_mstData.nOfferR[i] = objRq.GetDataValue(2, i);
                m_mstData.nBidR[i] = objRq.GetDataValue(3, i);
            }


            //foreach (DictionaryEntry item in m_mapMst)
            //{
            //    Debug.WriteLine("{0}:{1}", item.Key, item.Value);
            //}

            m_mstData.nOfferA = objRq.GetHeaderValue(71);
            m_mstData.nBidA = objRq.GetHeaderValue(73);

            return true;
        }
        private void setDataToGrid(PB_TYPE pbtype)
        {

            if (pbtype == PB_TYPE.ALL || pbtype == PB_TYPE.HOGA)
            {
                // 매도잔량 | 호가 채우기 
                for (int i = 0; i < NUM_HOGA; i++)
                {
                    //string sKey = (10 - i).ToString();
                    //m_DataTable.Rows[ROW_SEL_START + i]["매도잔량"] = string.Format("{0:N0}", m_mapMst[sKey + "_매도잔량"]);
                    //m_DataTable.Rows[ROW_SEL_START + i]["호가"] = string.Format("{0:N0}", m_mapMst[sKey + "_매도호가"]);

                    int nIndex = NUM_HOGA - i - 1;
                    m_DataTable.Rows[ROW_SEL_START + i]["매도잔량"] = string.Format("{0:N0}", m_mstData.nOfferR[nIndex]);
                    m_DataTable.Rows[ROW_SEL_START + i]["호가"] = string.Format("{0:N0}", m_mstData.nOfferP[nIndex]);
                }
                // 매수잔량 | 호가 채우기 
                for (int i = 0; i < NUM_HOGA; i++)
                {
                    //string sKey = (i + 1).ToString();
                    //m_DataTable.Rows[ROW_BUY_START + i]["호가"] = string.Format("{0:N0}", m_mapMst[sKey + "_매수호가"]);
                    //m_DataTable.Rows[ROW_BUY_START + i]["매수잔량"] = string.Format("{0:N0}", m_mapMst[sKey + "_매수잔량"]);
                    m_DataTable.Rows[ROW_BUY_START + i]["호가"] = string.Format("{0:N0}", m_mstData.nBidP[i]);
                    m_DataTable.Rows[ROW_BUY_START + i]["매수잔량"] = string.Format("{0:N0}", m_mstData.nBidR[i]);
                }


                m_DataTable.Rows[ROW_ALL_HOGA]["매도잔량"] = string.Format("{0:N0}", m_mstData.nOfferA);
                m_DataTable.Rows[ROW_ALL_HOGA]["매수잔량"] = string.Format("{0:N0}", m_mstData.nBidA);

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
                int nDiff = (int)m_mstData.nDiff;
                Color clrDiff = Color.FromArgb(0, 0, 0);
                if (nDiff > 0)
                    clrDiff = Color.FromArgb(255, 0, 0);
                else if (nDiff < 0)
                    clrDiff = Color.FromArgb(0, 0, 255);

                m_HogaView[COL_PRICE, ROW_CURPRICE].Style.ForeColor = clrDiff;
                m_HogaView[COL_BUYHOGA, ROW_CURPRICE].Style.ForeColor = clrDiff;
                m_HogaView[COL_BUYHOGA, ROW_CURPRICE - 1].Style.ForeColor = clrDiff;
            }

        }


        //////////////////////////////////////////////////////////////////////
        /// 실시간 등록 처리 
        //////////////////////////////////////////////////////////////////////

        // 실시간 시세 및 체결 등록
        private void SBStockInfos(string sCode)
        {
            // 이전 sb cancel

            // 주식 현재가 실시간 등록
            m_cpStockCur.SetInputValue(0, sCode);
            m_cpStockCur.Subscribe();

            // 주식 10차 호가 실시간 등록
            // SubscribeLatest 는 최근 값만 가져온다 --> 성능 개선 효과 
            m_cpStockJpBid.SetInputValue(0, sCode);
            m_cpStockJpBid.SubscribeLatest();
            m_sCode = sCode;



            Debug.WriteLine("SB START CODE {0}", sCode);

        }

        // 실시간 시세 및 체결 등록 해지 
        private void SbCancel(string sCode)
        {
            if (sCode == "")
                return;

            m_cpStockCur.SetInputValue(0, sCode);
            m_cpStockCur.Unsubscribe();

            m_cpStockJpBid.SetInputValue(0, sCode);
            m_cpStockJpBid.Unsubscribe();

        }

        //////////////////////////////////////////////////////////////////////
        /// PLUS 실시간 이벤트 처리 
        //////////////////////////////////////////////////////////////////////

        // 실시간 호가 수신
        private void m_cpStockJpBid_Received()
        {            
            int nStart = 3;
            for (int i = 0; i < 5; i++)
            {
                //string sKey = (i + 1).ToString();
                //m_mapMst[sKey + "_매도호가"] = m_cpStockJpBid.GetHeaderValue(nStart++);
                //m_mapMst[sKey + "_매수호가"] = m_cpStockJpBid.GetHeaderValue(nStart++);
                //m_mapMst[sKey + "_매도잔량"] = m_cpStockJpBid.GetHeaderValue(nStart++);
                //m_mapMst[sKey + "_매수잔량"] = m_cpStockJpBid.GetHeaderValue(nStart++);

                m_mstData.nOfferP[i] = m_cpStockJpBid.GetHeaderValue(nStart++);
                m_mstData.nBidP[i] = m_cpStockJpBid.GetHeaderValue(nStart++);
                m_mstData.nOfferR[i] = m_cpStockJpBid.GetHeaderValue(nStart++);
                m_mstData.nBidR[i] = m_cpStockJpBid.GetHeaderValue(nStart++);

            }

            nStart = 27;
            for (int i = 5; i < 10; i++)
            {
                //string sKey = (i + 1).ToString();
                //m_mapMst[sKey + "_매도호가"] = m_cpStockJpBid.GetHeaderValue(nStart++);
                //m_mapMst[sKey + "_매수호가"] = m_cpStockJpBid.GetHeaderValue(nStart++);
                //m_mapMst[sKey + "_매도잔량"] = m_cpStockJpBid.GetHeaderValue(nStart++);
                //m_mapMst[sKey + "_매수잔량"] = m_cpStockJpBid.GetHeaderValue(nStart++);


                m_mstData.nOfferP[i] = m_cpStockJpBid.GetHeaderValue(nStart++);
                m_mstData.nBidP[i] = m_cpStockJpBid.GetHeaderValue(nStart++);
                m_mstData.nOfferR[i] = m_cpStockJpBid.GetHeaderValue(nStart++);
                m_mstData.nBidR[i] = m_cpStockJpBid.GetHeaderValue(nStart++);

            }

            m_mstData.nOfferA = m_cpStockJpBid.GetHeaderValue(23);
            m_mstData.nBidA = m_cpStockJpBid.GetHeaderValue(24);


            setDataToGrid(PB_TYPE.HOGA);

        }
        // 실시간 시세체결 수신
        private void m_cpStockCur_Received()
        {
            //Debug.WriteLine("현재가 체결 PB 수신");

            m_mstData.nTime = m_cpStockCur.GetHeaderValue(3);
            m_mstData.nOpen = m_cpStockCur.GetHeaderValue(4);
            m_mstData.nHigh = m_cpStockCur.GetHeaderValue(5);
            m_mstData.nLow = m_cpStockCur.GetHeaderValue(6);

            m_mstData.nClose = m_cpStockCur.GetHeaderValue(13);
            m_mstData.nDiff = m_cpStockCur.GetHeaderValue(2);


            m_mstData.nBase = m_mstData.nClose - m_mstData.nDiff;
            m_mstData.dDiffR = 0;
            if (m_mstData.nBase > 0)
            {
                m_mstData.dDiffR = (float)(m_mstData.nDiff / (float)m_mstData.nBase) * 100;
            }


            setDataToGrid(PB_TYPE.CUR);
        }

        private void m_etCode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                string sCode = m_etCode.Text;

                string sName = m_cpCodeMgr.CodeToName("A" + sCode);
                if (sName == "")
                {
                    sName = m_cpCodeMgr.CodeToName("Q" + sCode);
                    if (sName == "")
                        return;
                    sCode = "Q" + sCode;

                }
                else
                    sCode = "A" + sCode; 

                reqeustStock(sCode);
            }

        }

        private void fmStock_FormClosing(object sender, FormClosingEventArgs e)
        {
            SbCancel(m_sCode);
        }
    } // end of fmStock
}
