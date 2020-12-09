using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

// 관리자 권한 실행 체크
using System.Security.Principal;
using System.Diagnostics;

// PLUS OBJECT
using CPTRADELib;
using System.Windows.Forms;
using System.Threading;


// PLUS 관련 공통 처리 함수 
namespace PlusTest
{

    // enum 정의
    public enum COMM_TYPE { SISE = 1, ACC = 0}

    public class CpCommon
    {
        private static CPTRADELib.CpTdUtil m_cpTdUtil = new CPTRADELib.CpTdUtil();
        private static CPUTILLib.CpCybos m_cpCybos = new CPUTILLib.CpCybos();
        private static string m_sAcc = "";
        private static string m_sAccFlag = "";

        public static string getAccount()
        {
            return m_sAcc;
        }
        public static string getAccFlag()
        {
            return m_sAccFlag;
        }

        // 생성자

        public static bool checkPlusStatus()
        {
            // 어드민 권한 실행 여부 체크 
            string sErrMsg; 
            if (IsUserAdministrator() == false)
            {
                sErrMsg = "관리자 권한으로 실행 안됨";
                Debug.WriteLine(sErrMsg);
                MessageBox.Show(sErrMsg, "실행확인", MessageBoxButtons.OK, MessageBoxIcon.Error);

                return false;
            }
            // 연결 체크 
            CPUTILLib.CpCybos cpCybos = new CPUTILLib.CpCybos();
            if (cpCybos.IsConnect == 0)
            {
                sErrMsg = "PLUS 연결 필요합니다";
                Debug.WriteLine(sErrMsg);
                MessageBox.Show(sErrMsg, "실행확인", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            Debug.WriteLine("연결 정상");



            // 계좌 체크 
            if (m_cpTdUtil.TradeInit() != 0)
            {
                sErrMsg = "TradeInit 실패";
                Debug.WriteLine(sErrMsg);
                MessageBox.Show(sErrMsg, "실행확인", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            Debug.WriteLine("TradeInit 성공");


            //string[] listAcc = (string[])m_cpTdUtil.AccountNumber;
            //if (listAcc.Length == 0)
            //    return false;

            //m_sAcc = listAcc[0];
            //string[] lstAccFlag = (string[])m_cpTdUtil.get_GoodsList(m_sAcc, CPE_ACC_GOODS.CPC_STOCK_ACC);    // 주식
            //if (lstAccFlag.Length == 0)
            //    return false;

            //m_sAccFlag = lstAccFlag[0];

            //Debug.WriteLine("계좌번호 " + m_sAcc + " 계좌구분 " + m_sAccFlag);

            return true;
        }


        public static bool IsUserAdministrator()
        {
            WindowsIdentity user = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new WindowsPrincipal(user);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        public static bool checkRqError(CPTRADELib.ICpTdDib objRq, int nRqRet) 
        {
            if (nRqRet != 0)
            {
                string sErrMsg = "통신 오류: " + nRqRet.ToString();
                MessageBox.Show(sErrMsg, "통신오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            short nStatus = objRq.GetDibStatus();
            string sMsg = objRq.GetDibMsg1();

            if (nStatus != 0) 
            {
                string sErrMsg = "통신상태: " + nStatus.ToString() + "\n오류메시지: " + sMsg;
                MessageBox.Show(sErrMsg, "통신오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                
                return false;
            }

            return true;
        }

        public static bool checkRqError(CPSYSDIBLib.ISysDib objRq, int nRqRet)
        {
            if (nRqRet != 0)
            {
                string sErrMsg = "통신 오류: " + nRqRet.ToString();
                MessageBox.Show(sErrMsg, "통신오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            short nStatus = objRq.GetDibStatus();
            string sMsg = objRq.GetDibMsg1();

            if (nStatus != 0)
            {
                string sErrMsg = "통신상태: " + nStatus.ToString() + "\n오류메시지: " + sMsg;
                MessageBox.Show(sErrMsg, "통신오류", MessageBoxButtons.OK, MessageBoxIcon.Error);

                return false;
            }

            return true;
        }


        public static bool checkRqError(DSCBO1Lib.IDib objRq, int nRqRet)
        {
            if (nRqRet != 0)
            {
                string sErrMsg = "통신 오류: " + nRqRet.ToString();
                MessageBox.Show(sErrMsg, "통신오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            short nStatus = objRq.GetDibStatus();
            string sMsg = objRq.GetDibMsg1();

            if (nStatus != 0)
            {
                string sErrMsg = "통신상태: " + nStatus.ToString() + "\n오류메시지: " + sMsg;
                MessageBox.Show(sErrMsg, "통신오류", MessageBoxButtons.OK, MessageBoxIcon.Error);

                return false;
            }

            return true;
        }

        public static void waitRqLimit(COMM_TYPE commType)
        {
            int nRemainCnt =m_cpCybos.GetLimitRemainCount((CPUTILLib.LIMIT_TYPE)commType);

            if (nRemainCnt > 0) {
                //Debug.WriteLine("남은 회수: " + nRemainCnt.ToString());
                return;
            }


            int nRemainTime = m_cpCybos.LimitRequestRemainTime;
            //Debug.WriteLine("제한 초과로 인해 대기 필요: " + nRemainTime.ToString());
            Thread.Sleep(nRemainTime);

            return;

        }



    }
}


