using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.EMMA;
using DocumentFormat.OpenXml.Office2010.Excel;
using MvcPaging;
using Newtonsoft.Json;
using NPOI.OpenXmlFormats.Dml.Diagram;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.EnterpriseServices.Internal;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using WEB.Export;
using WEB.tpehrmap04;
using WebGrease.Css.Ast;
using WebLib;

namespace WEB.Controllers
{
    [ControllerName(PageName = "獎金系統")]
    public class BonusPersonnelListController : BaseController
    {
        #region 參數
        private string 結轉人員名單Page = "獎金系統-專案清單(結轉人員名單)";
        private string 評核名單查詢與啟動Page = "獎金系統-專案清單(評核名單查詢與啟動)";
        private string 獎金分配Page = "獎金系統-獎金分配";
        private string 核定結果審核Page = "獎金系統-核定結果審核";//(核定結果-固定預算 目前已刪除)
        private string 簽核意見編輯Page = "簽核意見編輯";
        string Bonus_FormKind = "----------";
        ApplicationController APPController = new ApplicationController();
        BPMWebServiceController BPMWebController = new BPMWebServiceController();
        BPMWebService ws = new BPMWebService();
        #endregion

        // GET: Bonus/BonusPersonnelList

        #region List
        [CheckLoginSessionExpired]
        public ActionResult List(int BonusProjectID, string mode)
        {
            ViewBag.Mode = mode;
            PrepareSelectList();
            using (var conn = NewConnection())
            {
                conn.Open();

                var dao = NewDAO(conn);
                var Project = Project_NewDAO(conn);
                var count = dao.SelectCount_結轉名單(true, BonusProjectID: BonusProjectID);
                var data = dao.SelectPage_結轉名單(true, 0, PublicVariable.DefaultPageSize, BonusProjectID: BonusProjectID)
                    .ToPagedList(0, PublicVariable.DefaultPageSize, count);
                ViewBag.BonusProjectID = BonusProjectID;
                ViewBag.BonusYear = Project.獎金年度(BonusProjectID);

                SysLog.Write(LoginUserID, 結轉人員名單Page, SysLog.IntoPageLog(intoPageID: BonusProjectID, pageMode: SysLog.進入頁面.頁面, intoSuccess: true));
                return View(data);
            }
        }

        [CheckLoginSessionExpired]
        [HttpPost]
        public ActionResult List(BonusPersonnelListViewModel model)
        {
            var mode = DAC.GetString(model.Mode).ToLower();
            ViewBag.Mode = mode;
            var BonusProjectID = DAC.GetInt32Nullable(model.BonusProjectID);
            ViewBag.BonusProjectID = BonusProjectID;
            PrepareSelectList();


            if (model.Page > 0)
                model.Page--;
            if (model.Page < 0)
                model.Page = 0;
            if (model.PageSize <= 0)
                model.PageSize = PublicVariable.DefaultPageSize;

            ViewBag.Order = model.Order;
            ViewBag.Decending = model.Decending.ToString().ToLower();
            var order = string.IsNullOrEmpty(model.Order) ? "" : $"{model.Order} {(model.Decending ? "DESC" : "ASC")}";

            using (var conn = NewConnection())
            {
                conn.Open();

                var dao = NewDAO(conn);
                var Project = Project_NewDAO(conn);

                var count = dao.SelectCount_結轉名單(true
                    , BonusProjectID: model.BonusProjectID
                    , EMP_SEQ_NO: model.EMP_SEQ_NO_Search
                    );
                if (count < model.Page * model.PageSize)
                    model.Page = count / model.PageSize;
                var data = dao.SelectPage_結轉名單(true, model.Page * model.PageSize, model.PageSize
                    , BonusProjectID: model.BonusProjectID
                    , EMP_SEQ_NO: model.EMP_SEQ_NO_Search
                    , orderBy: order)
                    .ToPagedList(model.Page, model.PageSize, count);
                ViewBag.BonusProjectID = model.BonusProjectID;
                ViewBag.BonusYear = Project.獎金年度(DAC.GetInt32(BonusProjectID));
                return View("List", data);
            }
        }
        #endregion

        #region Project_Initiation
        [CheckLoginSessionExpired]
        public ActionResult Project_Initiation(int BonusProjectID, string mode)
        {
            ViewBag.Mode = mode;

            PrepareSelectList();
            using (var conn = NewConnection())
            {
                conn.Open();

                var dao = NewDAO(conn);
                var Project = Project_NewDAO(conn);
                var count = dao.SelectCount_不含排除名單(true, BonusProjectID: BonusProjectID);
                var data = dao.SelectPage_不含排除名單(true, 0, PublicVariable.DefaultPageSize, BonusProjectID: BonusProjectID)
                    .ToPagedList(0, PublicVariable.DefaultPageSize, count);

                ViewBag.levelStatus = new DAC_BonusProjectParameterApproval().Select(BonusProjectID: BonusProjectID).FirstOrDefault()?.BPM_Status ?? 0;
                ViewBag.levelBPM_FormNO = new DAC_BonusProjectParameterApproval().Select(BonusProjectID: BonusProjectID).FirstOrDefault()?.BPM_FormNO ?? "";
                ViewBag.BonusProjectID = BonusProjectID;
                ViewBag.BonusYear = Project.獎金年度(BonusProjectID);
                ViewBag.BonusCalculation = Project.BonusCalculation(DAC.GetInt32(BonusProjectID));
                ViewBag.BonusProjecStatus = Project.SelectOne(BonusProjectID).FirstOrDefault()?.BonusProjecStatus ?? 0;
                ViewBag.AllotType = Project.GetAllotType(BonusProjectID);
                SysLog.Write(LoginUserID, 評核名單查詢與啟動Page, SysLog.IntoPageLog(intoPageID: BonusProjectID, pageMode: SysLog.進入頁面.頁面, intoSuccess: true));
                return View(data);
            }
        }
        [CheckLoginSessionExpired]
        [HttpPost]
        public ActionResult Project_Initiation(BonusPersonnelListViewModel model)
        {
            var mode = DAC.GetString(model.Mode).ToLower();
            ViewBag.Mode = mode;
            var BonusProjectID = DAC.GetInt32Nullable(model.BonusProjectID);
            ViewBag.BonusProjectID = BonusProjectID;
            PrepareSelectList();
            ViewBag.levelStatus = new DAC_BonusProjectParameterApproval().Select(BonusProjectID: BonusProjectID).FirstOrDefault()?.BPM_Status ?? 0;
            ViewBag.levelBPM_FormNO = new DAC_BonusProjectParameterApproval().Select(BonusProjectID: BonusProjectID).FirstOrDefault()?.BPM_FormNO ?? "";

            var dac_Project = new DAC_BonusProject();
            var dac_BonusPersonnelList = new DAC_BonusPersonnelList();
            var ProjectItem = dac_Project.SelectOne(DAC.GetInt32(BonusProjectID)).FirstOrDefault();

            switch (model.SubmitButton)
            {
                case "啟動專案":
                    DAC_BonusPersonnelList _BonusPersonnelList = new DAC_BonusPersonnelList();
                    DAC_FL_DEPARTMENT_TW_V _DEPARTMENT_TW_V = new DAC_FL_DEPARTMENT_TW_V();
                    //啟動專案
                    if (ProjectItem != null)
                    {

                        string err = "";

                        if (new DAC_BonusProject().GetAllotType(DAC.GetInt32(BonusProjectID)) == 2)
                        {
                            //檢查扣除金額不得大於加碼預算，若扣除金額大於加碼預算，提示：啟動失敗！下列人員[扣除金額]超過[加碼預算]，請至[獎懲紀錄及預算扣除比例維護]重新調正[扣除金額]。
                            var PersonList = new DAC_BonusPersonnelList().Select_不含排除名單(BonusProjectID: DAC.GetInt32(BonusProjectID));
                            foreach (var p in PersonList)
                            {
                                if (DAC.GetInt32(p.DeductedAmount) > p.加碼預算)
                                    err += "【" + p.EMP_NO + p.EMP_NAME + "】";
                            }
                        }

                        if (err == "")
                        {
                            if ((int)ViewBag.levelStatus != 3)
                            {
                                ViewBag.Message = "獎金預算參數尚未簽核通過!!";
                                SysLog.Write(LoginUserID, 評核名單查詢與啟動Page, SysLog.FunctionLog(model.BonusPersonnelListID, todoSomething: "啟動專案", doSuccess: false, message: ViewBag.Message));
                            }
                            //檢查「獎金計算」是否完成
                            else if (ProjectItem.BonusCalculation != 1)
                            {
                                ViewBag.Message = "尚未進行獎金計算，獎金計算後才能啟動專案";
                                SysLog.Write(LoginUserID, 評核名單查詢與啟動Page, SysLog.FunctionLog(model.BonusPersonnelListID, todoSomething: "啟動專案", doSuccess: false, message: ViewBag.Message));
                            }
                            else
                            {
                                try
                                {
                                    // todo 如果「開放保留個人預算」= 1(BonusProject. ReserveBudget = 1) 時

                                    //1) 先刪除
                                    //BonusApprovalList 
                                    //DepartmentBudget
                                    //BonusFinalDistribution
                                    new DAC_BonusApprovalList().DoBackupandDelete(ProjectItem.BonusProjectID);
                                    new DAC_DepartmentBudget().DoBackupandDelete(ProjectItem.BonusProjectID);
                                    new DAC_BonusFinalDistribution().DoBackupandDelete(ProjectItem.BonusProjectID);
                                    new DAC_BonusFinalDistributionList().DoBackupandDelete(ProjectItem.BonusProjectID);
                                    new DAC_BonusFinalDistributionSUM().DoBackupandDelete(ProjectItem.BonusProjectID);
                                    new DAC_BonusRecordDetail().DoBackupandDelete(ProjectItem.BonusProjectID);
                                    new DAC_BonusRecordMaster().DoBackupandDelete(ProjectItem.BonusProjectID);
                                    //2) 再寫入
                                    //BonusApprovalList
                                    new DAC_BonusPersonnelList().WriteInBonusApprovalList(ProjectItem);
                                    //DepartmentBudget
                                    new DAC_DepartmentBudget().WriteInDepartmentBudget_取得部門(ProjectItem);
                                    //BonusFinalDistribution
                                    new DAC_BonusFinalDistribution().WriteInFinalSign(ProjectItem);
                                    //DAC_BonusApprovalRecord
                                    new DAC_BonusRecordMaster().WriteFirstIn(ProjectItem.BonusProjectID);

                                    //獎金計算完成則將專案設定的「主管作業」改為「是」
                                    ProjectItem.SupervisorJob = 1;
                                    ProjectItem.BonusProjecStatus = 2;//啟動中(畫面顯示：未結案)
                                    dac_Project.UpdateOne(ProjectItem);
                                    SysLog.Write(LoginUserID, 評核名單查詢與啟動Page, SysLog.FunctionLog(ProjectItem.BonusProjectID, todoSomething: "啟動專案", doSuccess: true));
                                    ViewBag.Message = "啟動專案成功";
                                    ViewBag.Mode = "start";
                                }
                                catch (Exception ex)
                                {
                                    ViewBag.Message = "啟動專案失敗，失敗原因:" + ex.Message;
                                    SysLog.Write(LoginUserID, 評核名單查詢與啟動Page, SysLog.FunctionLog(ProjectItem.BonusProjectID, todoSomething: "啟動專案", doSuccess: false, message: ViewBag.Message));
                                }
                            }
                        }
                        else
                        {
                            ViewBag.Message = "啟動失敗！下列人員【扣除金額】超過【加碼預算】，請至【獎懲紀錄及預算扣除比例維護】重新調正【扣除金額】。" + err;
                            SysLog.Write(LoginUserID, 評核名單查詢與啟動Page, SysLog.FunctionLog(model.BonusPersonnelListID, todoSomething: "啟動專案", doSuccess: false, message: ViewBag.Message));
                        }
                    }
                    break;
                case "獎金計算":
                    //獎金計算
                    if ((int)ViewBag.levelStatus != 3)
                    {
                        ViewBag.Message = "獎金預算參數尚未簽核通過!!";
                        SysLog.Write(LoginUserID, 評核名單查詢與啟動Page, SysLog.FunctionLog(model.BonusPersonnelListID, todoSomething: "獎金計算", doSuccess: false, message: ViewBag.Message));
                    }
                    else if (ProjectItem != null)
                    {
                        ProjectItem.BonusCalculation = 0;
                        if (dac_Project.UpdateOne(ProjectItem))
                        {
                            ViewBag.Message = "獎金計算開啟成功";
                            SysLog.Write(LoginUserID, 評核名單查詢與啟動Page, SysLog.FunctionLog(ProjectItem.BonusProjectID, todoSomething: "獎金計算開啟", doSuccess: true));
                        }
                        else
                        {
                            ViewBag.Message = "獎金計算開啟失敗";
                            SysLog.Write(LoginUserID, 評核名單查詢與啟動Page, SysLog.FunctionLog(ProjectItem.BonusProjectID, todoSomething: "獎金計算開啟", doSuccess: false, message: ViewBag.Message));
                        }
                    }
                    else
                    {
                        ViewBag.Message = "找不到專案主檔";
                        SysLog.Write(LoginUserID, 評核名單查詢與啟動Page, SysLog.FunctionLog(ProjectItem.BonusProjectID, todoSomething: "獎金計算開啟", doSuccess: false, message: ViewBag.Message));
                    }
                    break;
                case "手動執行獎金計算(測試用)":
                    //獎金計算
                    if ((int)ViewBag.levelStatus != 3)
                    {
                        ViewBag.Message = "獎金預算參數尚未簽核通過!!";
                        SysLog.Write(LoginUserID, 評核名單查詢與啟動Page, SysLog.FunctionLog(model.BonusPersonnelListID, todoSomething: "獎金計算", doSuccess: false, message: ViewBag.Message));
                    }
                    else if (ProjectItem != null)
                    {
                        ProjectItem.BonusCalculation = 0;
                        if (dac_Project.UpdateOne(ProjectItem))
                        {
                            ViewBag.Message = "獎金計算開啟成功";
                            SysLog.Write(LoginUserID, 評核名單查詢與啟動Page, SysLog.FunctionLog(ProjectItem.BonusProjectID, todoSomething: "獎金計算開啟", doSuccess: true));

                            #region 抓出獎金計算專案
                            foreach (var item in dac_Project.SelectOne(DAC.GetInt32(BonusProjectID)))
                            {
                                dac_BonusPersonnelList.BonusCalculationDLL(item);
                            }
                            #endregion


                        }
                        else
                        {
                            ViewBag.Message = "獎金計算開啟失敗";
                            SysLog.Write(LoginUserID, 評核名單查詢與啟動Page, SysLog.FunctionLog(ProjectItem.BonusProjectID, todoSomething: "獎金計算開啟", doSuccess: false, message: ViewBag.Message));
                        }
                    }
                    else
                    {
                        ViewBag.Message = "找不到專案主檔";
                        SysLog.Write(LoginUserID, 評核名單查詢與啟動Page, SysLog.FunctionLog(ProjectItem.BonusProjectID, todoSomething: "獎金計算開啟", doSuccess: false, message: ViewBag.Message));
                    }
                    break;
            }

            if (model.Page > 0)
                model.Page--;
            if (model.Page < 0)
                model.Page = 0;
            if (model.PageSize <= 0)
                model.PageSize = PublicVariable.DefaultPageSize;

            ViewBag.Order = model.Order;
            ViewBag.Decending = model.Decending.ToString().ToLower();
            var order = string.IsNullOrEmpty(model.Order) ? "" : $"{model.Order} {(model.Decending ? "DESC" : "ASC")}";

            using (var conn = NewConnection())
            {
                conn.Open();

                var dao = NewDAO(conn);
                var Project = Project_NewDAO(conn);

                var count = dao.SelectCount_不含排除名單(true
                    , BonusProjectID: model.BonusProjectID
                    , EMP_SEQ_NO: model.EMP_SEQ_NO_Search
                    );
                if (count < model.Page * model.PageSize)
                    model.Page = count / model.PageSize;
                var data = dao.SelectPage_不含排除名單(true, model.Page * model.PageSize, model.PageSize
                    , BonusProjectID: model.BonusProjectID
                    , EMP_SEQ_NO: model.EMP_SEQ_NO_Search
                    , orderBy: order)
                    .ToPagedList(model.Page, model.PageSize, count);
                ViewBag.BonusProjectID = model.BonusProjectID;
                ViewBag.BonusYear = Project.獎金年度(DAC.GetInt32(BonusProjectID));
                ViewBag.BonusCalculation = Project.BonusCalculation(DAC.GetInt32(BonusProjectID));
                ViewBag.BonusProjecStatus = Project.SelectOne(DAC.GetInt32(BonusProjectID)).FirstOrDefault()?.BonusProjecStatus ?? 0;
                ViewBag.AllotType = Project.GetAllotType(DAC.GetInt32(BonusProjectID));
                return View(data);
            }
        }
        #endregion

        #region BonusDistribution 主管分派作業
        [CheckLoginSessionExpired]
        public ActionResult BonusDistribution(int BonusProjectID)
        {
            DAC_BonusProject Project = new DAC_BonusProject();
            DAC_DepartmentBudget _DepartmentBudget = new DAC_DepartmentBudget();

            string 登入者 = (string)Session[PublicVariable.EMP_SEQ_NO];
            //所有的登入者層級帶入資料，皆需要依照部門來看該登入者所擁有的最高層級，而非直接用登入者來看
            //string 登入者層級 = new DAC_FL_DEPARTMENT_TW_V().GETLEVEL_CODE_主管層級(登入者);
            string 登入者DEPT_SEQ_NO = ((string)Session[PublicVariable.department_SEQ_NO]);

            ViewBag.BonusProjectID = BonusProjectID;
            ViewBag.ProjectName = Project.獎金專案名稱(BonusProjectID);
            ViewBag.HighLevelDepts = _DepartmentBudget.GetDepartmentBudget_兼任部門對應的預算(BonusProjectID, 登入者);
            ViewBag.IS總經理 = false;

            //如果非固定預算(彈性預算=2以及複合預算=3)，則判斷是否第一次操作而尚未儲存過。
            ViewBag.BudgetType = Project.SelectOne(BonusProjectID).FirstOrDefault() != null ? Project.SelectOne(BonusProjectID).FirstOrDefault().BudgetType : 0;
            ViewBag.IsFirstWorkTime = _DepartmentBudget.SelectAll()
                                      .Count(x => x.BonusProjectID == BonusProjectID
                                              && x.EMP_SEQ_NO == 登入者
                                              && x.DB_APPNO == 1) == _DepartmentBudget.SelectAll()
                                      .Count(x => x.BonusProjectID == BonusProjectID
                                              && x.EMP_SEQ_NO == 登入者);

            //如果總經理包含登入者
            if (new DAC_FL_DEPARTMENT_TW_V().Select_總經理().Where(p => p.MAIN_LEADER_EMP_NO == (string)Session[PublicVariable.UserId]).Count() > 0)
                ViewBag.IS總經理 = true;
            var Modify = _DepartmentBudget.Is主管是否可以編輯(BonusProjectID, 登入者);
            string str部門已分派 = "";

            if (Modify.IsModify == false)
            {
                str部門已分派 = "尚有部門 ( ";
                foreach (var d in Modify.deptitem)
                {
                    str部門已分派 += "【" + d.DEPT_NO + " " + d.DEPT_NAME + " 】";
                }
                str部門已分派 += " ) 待上級主管決議是否分派，請待決議後再填寫。";
            }
            var Data = Get單位(DAC.GetInt32(BonusProjectID), false);
            var model = new BonusApprovalListViewModel()
            {
                BonusProjectID = BonusProjectID,
                BonusDistributionData = Data,               
                ReserveBudget = Project.ReserveBudget(BonusProjectID) ?? 0,
                AllotType = Project.GetAllotType(BonusProjectID) ?? 0,
                budget = _DepartmentBudget.保留百分比(BonusProjectID, 登入者),
                btn登入者未進行分派 = _DepartmentBudget.IS登入者未進行分派(BonusProjectID, 登入者),
                btn保留百分比 = _DepartmentBudget.IS編輯保留百分比_new(DAC.GetInt32(BonusProjectID), 登入者, 登入者DEPT_SEQ_NO),
                btn上層主管已分派完畢 = Modify.IsModify,
                str部門已分派 = str部門已分派,
                上階主管分配金額 = Data.Sum(p => p.主管分配金額)
            };

            SysLog.Write(LoginUserID, 獎金分配Page, SysLog.IntoPageLog(BonusProjectID, pageMode: SysLog.進入頁面.頁面, intoSuccess: true));
            return View(model);
        }

        [CheckLoginSessionExpired]
        [HttpPost]
        public ActionResult BonusDistribution(BonusApprovalListViewModel model)
        {
            DAC_BonusProject Project = new DAC_BonusProject();
            DAC_DepartmentBudget _DepartmentBudget = new DAC_DepartmentBudget();
            string 登入者 = (string)Session[PublicVariable.EMP_SEQ_NO];
            //string 登入者層級 = new DAC_FL_DEPARTMENT_TW_V().GETLEVEL_CODE_主管層級(登入者);
            string 登入者DEPT_SEQ_NO = ((string)Session[PublicVariable.department_SEQ_NO]);
            ViewBag.BonusProjectID = model.BonusProjectID;
            ViewBag.ProjectName = Project.獎金專案名稱(DAC.GetInt32(model.BonusProjectID));
            ViewBag.HighLevelDepts = _DepartmentBudget.GetDepartmentBudget_兼任部門對應的預算(model.BonusProjectID, 登入者);
            ViewBag.IS總經理 = false;
            //第一次操作，尚未儲存過。
            ViewBag.IsFirstWorkTime = _DepartmentBudget.SelectAll()
                                      .Count(x => x.BonusProjectID == model.BonusProjectID
                                              && x.EMP_SEQ_NO == 登入者
                                              && x.DB_APPNO == 1) == _DepartmentBudget.SelectAll()
                                      .Count(x => x.BonusProjectID == model.BonusProjectID
                                              && x.EMP_SEQ_NO == 登入者);

            //如果總經理包含登入者
            if (new DAC_FL_DEPARTMENT_TW_V().Select_總經理().Where(p => p.MAIN_LEADER_EMP_NO == (string)Session[PublicVariable.UserId]).Count() > 0)
                ViewBag.IS總經理 = true;
            var Modify = _DepartmentBudget.Is主管是否可以編輯(DAC.GetInt32(model.BonusProjectID), 登入者);
            string str部門已分派 = "";
            if (Modify.IsModify == false)
            {
                str部門已分派 = "尚有部門 ( ";
                foreach (var d in Modify.deptitem)
                {
                    str部門已分派 += "【" + d.DEPT_NO + " " + d.DEPT_NAME + " 】";
                }
                str部門已分派 = " ) 待上級主管決議是否分派，請待決議後再填寫。";
            }

            switch (model.SubmitButton)
            {
                case "存檔": //存保留百分比
                    if (!CheckControl.Is正整數(model.budget))
                    {
                        ViewBag.Message = "存檔失敗，原因 : 請輸入正整數";
                    }
                    else if (DAC.GetInt32(model.budget) < 0 || DAC.GetInt32(model.budget) > 99)
                    {
                        ViewBag.Message = "存檔失敗，原因 : 僅能輸入0 ~ 99";
                    }
                    else
                    {
                        Decimal 保留百分比 = DAC.GetDecimal(DAC.GetDecimal(model.budget) / 100);
                        DAC_FL_DEPARTMENT_TW_V _DEPARTMENT_TW_V = new DAC_FL_DEPARTMENT_TW_V();
                        DAC_BonusApprovalList _BonusApprovalList = new DAC_BonusApprovalList();

                        switch (Project.GetAllotType(DAC.GetInt32(model.BonusProjectID)))
                        {
                            case 1:
                            case 2:
                            case 3:
                                var dList_親核 = _DepartmentBudget.Select(BonusProjectID: model.BonusProjectID, EMP_SEQ_NO: 登入者, ISApproval: true);
                                var dList = _DepartmentBudget.Select(BonusProjectID: model.BonusProjectID, EMP_SEQ_NO: 登入者, ISApproval: false);
                                #region 1、select DepartmentBudget 親核、轄下部門
                                foreach (var D_親核 in dList_親核)
                                {
                                    #region 2、算BonusApprovalList
                                    int 保留款金額 = 0;
                                    var 人員List = _BonusApprovalList.Select_登入者所負責的人員名單(BonusProjectID: model.BonusProjectID, Login_EMP_SEQ_NO: 登入者, DEPT_SEQ_NO: D_親核.DEPT_SEQ_NO, false);
                                    //var 人員List = _BonusApprovalList.Select_轄下人員(BonusProjectID: model.BonusProjectID, EMP_SEQ_NO: 登入者, DEPT_SEQ_NO: D_親核.DEPT_SEQ_NO);
                                    foreach (var P in 人員List)
                                    {
                                        int UnFixedBudget_org = 0;
                                        #region 判斷上一層主管的加碼金額
                                        switch (P.LEVEL_CODE_current)
                                        {
                                            case "1":
                                                UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget));
                                                break;
                                            case "2":
                                                if (P.UnFixedBudget1 != null && P.UnFixedBudget1 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget1));
                                                else
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget));
                                                break;
                                            case "3":
                                                if (P.UnFixedBudget2 != null && P.UnFixedBudget2 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget2));
                                                else if (P.UnFixedBudget1 != null && P.UnFixedBudget1 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget1));
                                                else
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget));
                                                break;
                                            case "4":
                                                if (P.UnFixedBudget3 != null && P.UnFixedBudget3 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget3));
                                                else if (P.UnFixedBudget2 != null && P.UnFixedBudget2 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget2));
                                                else if (P.UnFixedBudget1 != null && P.UnFixedBudget1 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget1));
                                                else
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget));
                                                break;
                                            case "5":
                                                if (P.UnFixedBudget4 != null && P.UnFixedBudget4 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget4));
                                                else if (P.UnFixedBudget3 != null && P.UnFixedBudget3 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget3));
                                                else if (P.UnFixedBudget2 != null && P.UnFixedBudget2 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget2));
                                                else if (P.UnFixedBudget1 != null && P.UnFixedBudget1 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget1));
                                                else
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget));
                                                break;
                                            case "6":
                                                if (P.UnFixedBudget5 != null && P.UnFixedBudget5 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget5));
                                                else if (P.UnFixedBudget4 != null && P.UnFixedBudget4 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget4));
                                                else if (P.UnFixedBudget3 != null && P.UnFixedBudget3 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget3));
                                                else if (P.UnFixedBudget2 != null && P.UnFixedBudget2 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget2));
                                                else if (P.UnFixedBudget1 != null && P.UnFixedBudget1 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget1));
                                                else
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget));
                                                break;
                                            case "7":
                                                if (P.UnFixedBudget6 != null && P.UnFixedBudget6 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget6));
                                                else if (P.UnFixedBudget5 != null && P.UnFixedBudget5 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget5));
                                                else if (P.UnFixedBudget4 != null && P.UnFixedBudget4 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget4));
                                                else if (P.UnFixedBudget3 != null && P.UnFixedBudget3 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget3));
                                                else if (P.UnFixedBudget2 != null && P.UnFixedBudget2 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget2));
                                                else if (P.UnFixedBudget1 != null && P.UnFixedBudget1 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget1));
                                                else
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget));
                                                break;
                                            case "8":
                                                if (P.UnFixedBudget7 != null && P.UnFixedBudget7 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget7));
                                                else if (P.UnFixedBudget6 != null && P.UnFixedBudget6 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget6));
                                                else if (P.UnFixedBudget5 != null && P.UnFixedBudget5 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget5));
                                                else if (P.UnFixedBudget4 != null && P.UnFixedBudget4 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget4));
                                                else if (P.UnFixedBudget3 != null && P.UnFixedBudget3 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget3));
                                                else if (P.UnFixedBudget2 != null && P.UnFixedBudget2 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget2));
                                                else if (P.UnFixedBudget1 != null && P.UnFixedBudget1 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget1));
                                                else
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget));
                                                break;
                                            case "9":
                                                if (P.UnFixedBudget8 != null && P.UnFixedBudget8 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget8));
                                                else if (P.UnFixedBudget7 != null && P.UnFixedBudget7 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget7));
                                                else if (P.UnFixedBudget6 != null && P.UnFixedBudget6 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget6));
                                                else if (P.UnFixedBudget5 != null && P.UnFixedBudget5 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget5));
                                                else if (P.UnFixedBudget4 != null && P.UnFixedBudget4 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget4));
                                                else if (P.UnFixedBudget3 != null && P.UnFixedBudget3 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget3));
                                                else if (P.UnFixedBudget2 != null && P.UnFixedBudget2 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget2));
                                                else if (P.UnFixedBudget1 != null && P.UnFixedBudget1 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget1));
                                                else
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget));
                                                break;
                                            case "10":
                                                if (P.UnFixedBudget9 != null && P.UnFixedBudget9 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget9));
                                                else if (P.UnFixedBudget8 != null && P.UnFixedBudget8 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget8));
                                                else if (P.UnFixedBudget7 != null && P.UnFixedBudget7 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget7));
                                                else if (P.UnFixedBudget6 != null && P.UnFixedBudget6 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget6));
                                                else if (P.UnFixedBudget5 != null && P.UnFixedBudget5 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget5));
                                                else if (P.UnFixedBudget4 != null && P.UnFixedBudget4 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget4));
                                                else if (P.UnFixedBudget3 != null && P.UnFixedBudget3 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget3));
                                                else if (P.UnFixedBudget2 != null && P.UnFixedBudget2 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget2));
                                                else if (P.UnFixedBudget1 != null && P.UnFixedBudget1 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget1));
                                                else
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget));
                                                break;
                                            default:
                                                break;
                                        }
                                        #endregion
                                        int UnFixedBudget_保留後金額 = DAC.GetInt32(Math.Round(UnFixedBudget_org * (1 - 保留百分比), 0, MidpointRounding.AwayFromZero));
                                        保留款金額 += UnFixedBudget_org - UnFixedBudget_保留後金額;
                                        #region 寫入該層級UnFixedBudget
                                        switch (P.LEVEL_CODE_current)
                                        {
                                            case "1":
                                                P.UnFixedBudget1 = StringEncrypt.aesEncryptBase64(DAC.GetString(UnFixedBudget_保留後金額)); //加密存回去
                                                break;
                                            case "2":
                                                P.UnFixedBudget2 = StringEncrypt.aesEncryptBase64(DAC.GetString(UnFixedBudget_保留後金額)); //加密存回去
                                                break;
                                            case "3":
                                                P.UnFixedBudget3 = StringEncrypt.aesEncryptBase64(DAC.GetString(UnFixedBudget_保留後金額)); //加密存回去
                                                break;
                                            case "4":
                                                P.UnFixedBudget4 = StringEncrypt.aesEncryptBase64(DAC.GetString(UnFixedBudget_保留後金額)); //加密存回去
                                                break;
                                            case "5":
                                                P.UnFixedBudget5 = StringEncrypt.aesEncryptBase64(DAC.GetString(UnFixedBudget_保留後金額)); //加密存回去
                                                break;
                                            case "6":
                                                P.UnFixedBudget6 = StringEncrypt.aesEncryptBase64(DAC.GetString(UnFixedBudget_保留後金額)); //加密存回去
                                                break;
                                            case "7":
                                                P.UnFixedBudget7 = StringEncrypt.aesEncryptBase64(DAC.GetString(UnFixedBudget_保留後金額)); //加密存回去
                                                break;
                                            case "8":
                                                P.UnFixedBudget8 = StringEncrypt.aesEncryptBase64(DAC.GetString(UnFixedBudget_保留後金額)); //加密存回去
                                                break;
                                            case "9":
                                                P.UnFixedBudget9 = StringEncrypt.aesEncryptBase64(DAC.GetString(UnFixedBudget_保留後金額)); //加密存回去
                                                break;
                                            case "10":
                                                P.UnFixedBudget10 = StringEncrypt.aesEncryptBase64(DAC.GetString(UnFixedBudget_保留後金額)); //加密存回去
                                                break;
                                            default:
                                                break;
                                        }
                                        #endregion
                                        _BonusApprovalList.UpdateOne(P);
                                    }
                                    #endregion

                                    #region 3、寫入DepartmentBudget.ReserveBudgetRatio && DepartmentBudget.ReserveBudget
                                    int DUnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(D_親核.UnFixedBudget));
                                    int DFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(D_親核.FixedBudget));

                                    D_親核.ReserveBudgetRatio = 保留百分比; //保留百分比
                                    D_親核.ReserveBudget = StringEncrypt.aesEncryptBase64(DAC.GetString(保留款金額)); //保留款金額
                                    D_親核.Amount = StringEncrypt.aesEncryptBase64(DAC.GetString(DFixedBudget_org + DUnFixedBudget_org - 保留款金額));
                                    _DepartmentBudget.UpdateOne(D_親核);

                                    #endregion

                                }
                                foreach (var D in dList)
                                {
                                    #region 2、算BonusApprovalList
                                    int 保留款金額 = 0;
                                    var 人員List = _BonusApprovalList.Select_登入者所負責的人員名單(BonusProjectID: model.BonusProjectID, Login_EMP_SEQ_NO: 登入者, DEPT_SEQ_NO: D.DEPT_SEQ_NO, true);
                                    //var 人員List = _BonusApprovalList.Select_轄下人員(BonusProjectID: model.BonusProjectID, EMP_SEQ_NO: 登入者, DEPT_SEQ_NO: _DEPARTMENT_TW_V.GetChild_DEPT_SEQ_NO_List(D.DEPT_SEQ_NO, DAC_FL_DEPARTMENT_TW_V.資訊類型.代號));
                                    foreach (var P in 人員List)
                                    {
                                        int UnFixedBudget_org = 0;
                                        #region 判斷上一層主管的加碼金額
                                        switch (P.LEVEL_CODE_current)
                                        {
                                            case "1":
                                                UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget));
                                                break;
                                            case "2":
                                                if (P.UnFixedBudget1 != null && P.UnFixedBudget1 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget1));
                                                else
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget));
                                                break;
                                            case "3":
                                                if (P.UnFixedBudget2 != null && P.UnFixedBudget2 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget2));
                                                else if (P.UnFixedBudget1 != null && P.UnFixedBudget1 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget1));
                                                else
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget));
                                                break;
                                            case "4":
                                                if (P.UnFixedBudget3 != null && P.UnFixedBudget3 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget3));
                                                else if (P.UnFixedBudget2 != null && P.UnFixedBudget2 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget2));
                                                else if (P.UnFixedBudget1 != null && P.UnFixedBudget1 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget1));
                                                else
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget));
                                                break;
                                            case "5":
                                                if (P.UnFixedBudget4 != null && P.UnFixedBudget4 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget4));
                                                else if (P.UnFixedBudget3 != null && P.UnFixedBudget3 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget3));
                                                else if (P.UnFixedBudget2 != null && P.UnFixedBudget2 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget2));
                                                else if (P.UnFixedBudget1 != null && P.UnFixedBudget1 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget1));
                                                else
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget));
                                                break;
                                            case "6":
                                                if (P.UnFixedBudget5 != null && P.UnFixedBudget5 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget5));
                                                else if (P.UnFixedBudget4 != null && P.UnFixedBudget4 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget4));
                                                else if (P.UnFixedBudget3 != null && P.UnFixedBudget3 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget3));
                                                else if (P.UnFixedBudget2 != null && P.UnFixedBudget2 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget2));
                                                else if (P.UnFixedBudget1 != null && P.UnFixedBudget1 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget1));
                                                else
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget));
                                                break;
                                            case "7":
                                                if (P.UnFixedBudget6 != null && P.UnFixedBudget6 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget6));
                                                else if (P.UnFixedBudget5 != null && P.UnFixedBudget5 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget5));
                                                else if (P.UnFixedBudget4 != null && P.UnFixedBudget4 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget4));
                                                else if (P.UnFixedBudget3 != null && P.UnFixedBudget3 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget3));
                                                else if (P.UnFixedBudget2 != null && P.UnFixedBudget2 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget2));
                                                else if (P.UnFixedBudget1 != null && P.UnFixedBudget1 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget1));
                                                else
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget));
                                                break;
                                            case "8":
                                                if (P.UnFixedBudget7 != null && P.UnFixedBudget7 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget7));
                                                else if (P.UnFixedBudget6 != null && P.UnFixedBudget6 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget6));
                                                else if (P.UnFixedBudget5 != null && P.UnFixedBudget5 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget5));
                                                else if (P.UnFixedBudget4 != null && P.UnFixedBudget4 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget4));
                                                else if (P.UnFixedBudget3 != null && P.UnFixedBudget3 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget3));
                                                else if (P.UnFixedBudget2 != null && P.UnFixedBudget2 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget2));
                                                else if (P.UnFixedBudget1 != null && P.UnFixedBudget1 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget1));
                                                else
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget));
                                                break;
                                            case "9":
                                                if (P.UnFixedBudget8 != null && P.UnFixedBudget8 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget8));
                                                else if (P.UnFixedBudget7 != null && P.UnFixedBudget7 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget7));
                                                else if (P.UnFixedBudget6 != null && P.UnFixedBudget6 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget6));
                                                else if (P.UnFixedBudget5 != null && P.UnFixedBudget5 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget5));
                                                else if (P.UnFixedBudget4 != null && P.UnFixedBudget4 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget4));
                                                else if (P.UnFixedBudget3 != null && P.UnFixedBudget3 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget3));
                                                else if (P.UnFixedBudget2 != null && P.UnFixedBudget2 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget2));
                                                else if (P.UnFixedBudget1 != null && P.UnFixedBudget1 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget1));
                                                else
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget));
                                                break;
                                            case "10":
                                                if (P.UnFixedBudget9 != null && P.UnFixedBudget9 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget9));
                                                else if (P.UnFixedBudget8 != null && P.UnFixedBudget8 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget8));
                                                else if (P.UnFixedBudget7 != null && P.UnFixedBudget7 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget7));
                                                else if (P.UnFixedBudget6 != null && P.UnFixedBudget6 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget6));
                                                else if (P.UnFixedBudget5 != null && P.UnFixedBudget5 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget5));
                                                else if (P.UnFixedBudget4 != null && P.UnFixedBudget4 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget4));
                                                else if (P.UnFixedBudget3 != null && P.UnFixedBudget3 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget3));
                                                else if (P.UnFixedBudget2 != null && P.UnFixedBudget2 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget2));
                                                else if (P.UnFixedBudget1 != null && P.UnFixedBudget1 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget1));
                                                else
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget));
                                                break;
                                            default:
                                                break;
                                        }
                                        #endregion
                                        int UnFixedBudget_保留後金額 = DAC.GetInt32(Math.Round(UnFixedBudget_org * (1 - 保留百分比), 0, MidpointRounding.AwayFromZero));
                                        保留款金額 += UnFixedBudget_org - UnFixedBudget_保留後金額;
                                        #region 寫入該層級UnFixedBudget
                                        switch (P.LEVEL_CODE_current)
                                        {
                                            case "1":
                                                P.UnFixedBudget1 = StringEncrypt.aesEncryptBase64(DAC.GetString(UnFixedBudget_保留後金額)); //加密存回去
                                                break;
                                            case "2":
                                                P.UnFixedBudget2 = StringEncrypt.aesEncryptBase64(DAC.GetString(UnFixedBudget_保留後金額)); //加密存回去
                                                break;
                                            case "3":
                                                P.UnFixedBudget3 = StringEncrypt.aesEncryptBase64(DAC.GetString(UnFixedBudget_保留後金額)); //加密存回去
                                                break;
                                            case "4":
                                                P.UnFixedBudget4 = StringEncrypt.aesEncryptBase64(DAC.GetString(UnFixedBudget_保留後金額)); //加密存回去
                                                break;
                                            case "5":
                                                P.UnFixedBudget5 = StringEncrypt.aesEncryptBase64(DAC.GetString(UnFixedBudget_保留後金額)); //加密存回去
                                                break;
                                            case "6":
                                                P.UnFixedBudget6 = StringEncrypt.aesEncryptBase64(DAC.GetString(UnFixedBudget_保留後金額)); //加密存回去
                                                break;
                                            case "7":
                                                P.UnFixedBudget7 = StringEncrypt.aesEncryptBase64(DAC.GetString(UnFixedBudget_保留後金額)); //加密存回去
                                                break;
                                            case "8":
                                                P.UnFixedBudget8 = StringEncrypt.aesEncryptBase64(DAC.GetString(UnFixedBudget_保留後金額)); //加密存回去
                                                break;
                                            case "9":
                                                P.UnFixedBudget9 = StringEncrypt.aesEncryptBase64(DAC.GetString(UnFixedBudget_保留後金額)); //加密存回去
                                                break;
                                            case "10":
                                                P.UnFixedBudget10 = StringEncrypt.aesEncryptBase64(DAC.GetString(UnFixedBudget_保留後金額)); //加密存回去
                                                break;
                                            default:
                                                break;
                                        }
                                        #endregion
                                        _BonusApprovalList.UpdateOne(P);
                                    }
                                    #endregion

                                    #region 3、寫入DepartmentBudget.ReserveBudgetRatio && DepartmentBudget.ReserveBudget
                                    int DUnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(D.UnFixedBudget));
                                    int DFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(D.FixedBudget));
                                    D.ReserveBudgetRatio = 保留百分比; //保留百分比
                                    D.ReserveBudget = StringEncrypt.aesEncryptBase64(DAC.GetString(保留款金額)); //保留款金額
                                    D.Amount = StringEncrypt.aesEncryptBase64(DAC.GetString(DFixedBudget_org + DUnFixedBudget_org - 保留款金額));
                                    _DepartmentBudget.UpdateOne(D);
                                    #endregion
                                }
                                #endregion
                                break;
                            case 4:
                                break;
                            default:
                                break;
                        }
                    }
                    break;
                case "儲存分配金額":
                    {
                        long 主管分配金額加總 = model.BonusDistributionData.Sum(p => DAC.GetInt64(p.主管分配金額str.Replace(",", "")));
                        if (主管分配金額加總 > model.上階主管分配金額)
                        {
                            ViewBag.Message1 = "儲存失敗！超過可分配預算";
                        }
                        if (主管分配金額加總 < model.上階主管分配金額)
                        {
                            ViewBag.Message1 = "儲存失敗！尚有可分配預算";
                        }
                        if (主管分配金額加總 == model.上階主管分配金額)
                        {
                            DepartmentBudgetItem item_DepartmentBudget = null;
                            var dList_親核 = _DepartmentBudget.Select(BonusProjectID: model.BonusProjectID, EMP_SEQ_NO: 登入者, ISApproval: true);
                            for (int i = 0; i < dList_親核.Count(); i++)
                            {
                                if (i == 0)
                                {
                                    //畫面上親核的DepartmentBudget 某一列
                                    var d親核 = model.BonusDistributionData.Where(p => p.ISApproval == true).FirstOrDefault();
                                    dList_親核[i].FlexibleBudget = StringEncrypt.aesEncryptBase64(DAC.GetString(d親核.主管分配金額str.Replace(",", "")));
                                    _DepartmentBudget.UpdateOne(dList_親核[i]);
                                }
                                else
                                {
                                    dList_親核[i].FlexibleBudget = StringEncrypt.aesEncryptBase64("0");
                                    _DepartmentBudget.UpdateOne(dList_親核[i]);
                                }
                            }
                            foreach (var BonusDistributionData in model.BonusDistributionData)
                            {
                                if (BonusDistributionData.ISApproval == true)
                                    continue;
                                item_DepartmentBudget = _DepartmentBudget.SelectOne(BonusDistributionData.DepartmentBudgetID).FirstOrDefault();
                                item_DepartmentBudget.FlexibleBudget = StringEncrypt.aesEncryptBase64(DAC.GetString(BonusDistributionData.主管分配金額str.Replace(",", "")));
                                _DepartmentBudget.UpdateOne(item_DepartmentBudget);
                            }

                            ViewBag.Message1 = "儲存成功！";
                        }
                    }
                    break;
            }
            var Data = Get單位(DAC.GetInt32(model.BonusProjectID), false);
            model = new BonusApprovalListViewModel()
            {
                BonusProjectID = DAC.GetInt32(model.BonusProjectID),
                BonusDistributionData = Data,
                ReserveBudget = Project.ReserveBudget(DAC.GetInt32(model.BonusProjectID)) ?? 0,
                AllotType = Project.GetAllotType(DAC.GetInt32(model.BonusProjectID)) ?? 0,
                budget = _DepartmentBudget.保留百分比(DAC.GetInt32(model.BonusProjectID), 登入者),
                btn登入者未進行分派 = _DepartmentBudget.IS登入者未進行分派(model.BonusProjectID, 登入者),
                btn保留百分比 = _DepartmentBudget.IS編輯保留百分比_new(DAC.GetInt32(model.BonusProjectID), 登入者, 登入者DEPT_SEQ_NO),
                btn上層主管已分派完畢 = Modify.IsModify,
                str部門已分派 = str部門已分派,
                上階主管分配金額 = Data.Sum(p => p.主管分配金額)
            };
            return View(model);
        }
        #endregion

        #region BonusDistributionList 主管獎金作業 (新版)
        /// <summary>
        /// 
        /// </summary>
        /// <param name="BonusProjectID"></param>
        /// <returns></returns>
        [CheckLoginSessionExpired]
        public ActionResult BonusDistributionList(int BonusProjectID)
        {
            #region 宣告
            DAC_FL_DEPARTMENT_TW_V _DEPARTMENT_TW_V = new DAC_FL_DEPARTMENT_TW_V();
            DAC_DepartmentBudget _DepartmentBudget = new DAC_DepartmentBudget();
            DAC_BonusProject Project = new DAC_BonusProject();
            DAC_BonusApprovalList _BonusApprovalList = new DAC_BonusApprovalList();
            var 登入者 = (string)Session[PublicVariable.EMP_SEQ_NO];
            var 登入者層級 = new DAC_FL_DEPARTMENT_TW_V().GETLEVEL_CODE_主管層級(登入者);
            var 登入者管理部門Lst = 轄下部門();
            //string 登入者DEPT_SEQ_NO = ((string)Session[PublicVariable.department_SEQ_NO]);
            //string 登入者所屬部門主管DEPT_SEQ_NO = new DAC_FL_DEPARTMENT_TW_V().GETDEPT_主管部門(登入者);
            bool 有被分配部門 = true;
            #endregion


            #region 搜尋功能優化
            //員工 - 僅可選擇管轄部門的員工
            var OwnDeptSql = "DEPT_SEQ_PATH = \\'\\' ";
            foreach (var dept in 登入者管理部門Lst.Split(','))
            {
                OwnDeptSql += " OR DEPT_SEQ_PATH LIKE \\'%|" + dept + "|%\\' ";
            }
            ViewBag.SubDeptLst = OwnDeptSql;
            ViewBag.BonusProjectID = BonusProjectID;
            #endregion


            #region 計算 (抽離)
            ViewBag.AllotType = Project.GetAllotType(BonusProjectID) ?? 0;
            ViewBag.BonusProjectID = BonusProjectID;
            ViewBag.ProjectName = Project.獎金專案名稱(BonusProjectID);
            ViewBag.BonusYear = Project.獎金年度(BonusProjectID);
            ViewBag.ReserveBudget = Project.ReserveBudget(BonusProjectID) ?? 0;
            ////var Data = Get單位(DAC.GetInt32(BonusProjectID), false);   // 效能調整
            ////ViewBag.Salary_D = Data.Sum(p => p.主管分配金額);            // 效能調整
            ////ViewBag.Salary_V = Data.Sum(p => p.保留金額);                // 效能調整
            //var dList = _DepartmentBudget.Select(BonusProjectID: DAC.GetInt32(BonusProjectID), EMP_SEQ_NO: 登入者);

            ////不理解這邊要判斷登入者權限的原因，所以先註解掉，若有問題再進行檢驗
            ////if (DAC.GetInt32(登入者層級) > Project.GetExecutionLevel(BonusProjectID) || dList.Count() == 0)
            //if (dList.Count() == 0)
            //{
            //    有被分配部門 = false;
            //}

            //ViewBag.總預算 = 0; ViewBag.總預算_FixedBudget = 0; ViewBag.總預算_UnFixedBudget = 0; ViewBag.保留金額 = 0;
            //ViewBag.主管微調總金額 = 0; ViewBag.轄下主管調整總額 = 0; ViewBag.已核定總金額 = 0; ViewBag.可用餘額 = 0; ViewBag.btnApproval = false;

            //var 獎金專案明細Item = new DAC_DepartmentBudget().獎金專案明細(DAC.GetInt32(BonusProjectID), 登入者);
            //if (有被分配部門)
            //{
            //    WriteLog($"===計算 總預算_UnFixedBudget[AllotType={DAC.GetString(ViewBag.AllotType)}]=== start");
            //    switch (DAC.GetString(ViewBag.AllotType))
            //    {
            //        case "1":
            //            ViewBag.總預算_UnFixedBudget = 獎金專案明細Item.總預算_UnFixedBudget;
            //            break;
            //        case "2":
            //            ViewBag.總預算_FixedBudget = 獎金專案明細Item.總預算_FixedBudget;
            //            ViewBag.總預算_UnFixedBudget = 獎金專案明細Item.總預算_UnFixedBudget;
            //            break;
            //        case "3":
            //            ViewBag.總預算_UnFixedBudget = 獎金專案明細Item.總預算_UnFixedBudget;
            //            break;
            //        case "4":
            //            break;
            //        default:
            //            break;
            //    }
            //    WriteLog($"===計算 總預算_UnFixedBudget[AllotType={DAC.GetString(ViewBag.AllotType)}]=== end");
            //     // 效能調整//ViewBag.保留金額 = 獎金專案明細Item.保留金額總合;
            //    // 效能調整//ViewBag.總預算 = 獎金專案明細Item.總預算_FixedBudget+ 獎金專案明細Item.總預算_UnFixedBudget + 獎金專案明細Item.保留金額總合 + DAC.GetInt32(ViewBag.Salary_D);
            //    // 效能調整//ViewBag.主管微調總金額 = _BonusApprovalList.主管微調總金額_new(BonusProjectID, 登入者);
            //    WriteLog($"===計算 主管微調總金額=== end");
            //    // 效能調整//ViewBag.轄下主管調整總額 = _BonusApprovalList.轄下主管調整總額_new(DAC.GetInt32(BonusProjectID), 登入者);
            //    WriteLog($"===計算 轄下主管調整總額=== end");
            //    // 效能調整//ViewBag.已核定總金額 = DAC.GetInt32(ViewBag.總預算_FixedBudget) + DAC.GetInt32(ViewBag.總預算_UnFixedBudget) + DAC.GetInt32(ViewBag.主管微調總金額) + DAC.GetInt32(ViewBag.轄下主管調整總額);
            //    // 效能調整//ViewBag.可用餘額 = DAC.GetInt32(ViewBag.總預算) - DAC.GetInt32(ViewBag.已核定總金額);
            //    // 效能調整//ViewBag.btnApproval = _DepartmentBudget.Check是否可送簽_new(BonusProjectID, 登入者);
            //    ViewBag.預定簽核數 = _BonusApprovalList.SelectCount(BonusProjectID: BonusProjectID, ReSignerID: DAC.GetInt32(登入者));
            //    // 效能調整//ViewBag.本次簽核人員已全部送簽 = _BonusApprovalList.本次簽核人員已全部送簽(BonusProjectID, 登入者);
            //    WriteLog($"===Check是否可送簽=== end");
            //}
            #endregion

            #region 判斷
            var Modify = _DepartmentBudget.Is主管是否可以編輯(BonusProjectID, 登入者);
            string str部門已分派 = "";
            if (Modify.IsModify == false)
            {
                str部門已分派 = "尚有部門 ( ";
                foreach (var d in Modify.deptitem)
                {
                    str部門已分派 += "【" + d.DEPT_NO + " " + d.DEPT_NAME + " 】";
                }
                str部門已分派 += " ) 待上級主管決議是否分派，請待決議後再填寫。";
            }
            ViewBag.btn上層主管已分派完畢 = Modify.IsModify;
            ViewBag.str部門已分派 = str部門已分派;
            WriteLog($"===Is主管是否可以編輯=== end");
            #endregion

            #region
            //WriteLog($"===CountBudget=== start");
            //_BonusApprovalList.CountBudget(BonusProjectID, 登入者, 登入者所屬部門主管DEPT_SEQ_NO, 登入者層級, DAC.GetInt32(ViewBag.AllotType));
            //WriteLog($"===CountBudget=== end");
            #endregion

            var order = " LEVEL_CODE, DEPT_NAME, SLY_DEGREE desc ";

            //要加上 可編輯核定金額欄位(上階未送簽過或是上階的主管有駁回)
            if (有被分配部門 == false)
            {
                var data = _BonusApprovalList.SelectPage_人員_new(false, orderBy: order).ToPagedList(0, 100, 0);
                return View(data);
            }
            else
            {
                var count = _BonusApprovalList.SelectCount_人員_new(true
                    , 登入者: 登入者,
                    BonusProjectID: BonusProjectID);
                var data = _BonusApprovalList.SelectPage_人員_new(true, 0, 100
                      , 登入者: 登入者,
                      BonusProjectID: BonusProjectID,
                      orderBy: order)
                      .ToPagedList(0, 100, count);
                WriteLog($"===SelectPage=== end");
                return View(data);
            }
        }

        [CheckLoginSessionExpired]
        [HttpPost]
        public ActionResult BonusDistributionList(BonusApprovalListViewModel model)
        {
            #region 宣告
            DAC_FL_DEPARTMENT_TW_V _DEPARTMENT_TW_V = new DAC_FL_DEPARTMENT_TW_V();
            DAC_DepartmentBudget _DepartmentBudget = new DAC_DepartmentBudget();
            DAC_BonusProject Project = new DAC_BonusProject();
            DAC_BonusApprovalList _BonusApprovalList = new DAC_BonusApprovalList();

            var 登入者 = (string)Session[PublicVariable.EMP_SEQ_NO];
            var 登入者層級 = new DAC_FL_DEPARTMENT_TW_V().GETLEVEL_CODE_主管層級(登入者);
            var 登入者管理部門Lst = 轄下部門();
            //string 登入者DEPT_SEQ_NO = ((string)Session[PublicVariable.department_SEQ_NO]);
            //string 登入者所屬部門主管DEPT_SEQ_NO = new DAC_FL_DEPARTMENT_TW_V().GETDEPT_主管部門(登入者);
            bool 有被分配部門 = true;
            ViewBag.AllotType = Project.GetAllotType(DAC.GetInt32(model.BonusProjectID)) ?? 0;
            #endregion
            switch (model.SubmitButton)
            {
                case "不核定預算為0":
                    var Approval = ApprovalnZero_new(DAC.GetInt32(model.BonusProjectID));
                    if (Approval.Approval)
                    {
                        ViewBag.Message = "[不核定預算為0]寫入成功";
                    }
                    else
                    {
                        ViewBag.Message = "[不核定預算為0]寫入失敗，" + Approval.error;
                    }
                    break;
                case "送出":
                    string 預算送出前檢查(BonusApprovalListItem empApproval, string checkType)
                    {
                        switch (checkType)
                        {
                            case "年終調整金額已逾加碼預算":
                                //只有年終獎金才做處理
                                if ((int)ViewBag.AllotType != 2)
                                    return "";
                                if (DAC.GetInt64(empApproval.PreAdjust_current) + DAC.GetInt64(empApproval.轄下主管調整金額加總) + DAC.GetInt64(empApproval.當前加碼預算) < 0)
                                {
                                    return $"、【{empApproval.EMP_NO} {empApproval.EMP_NAME}】";
                                }
                                break;
                            case "本次發放金額":
                                //[本次發放金額]不能小於0。"
                                if (DAC.GetInt64(empApproval.本次發放金額) + DAC.GetInt64(empApproval.PreAdjust_current) < 0)
                                    return $"、【{empApproval.EMP_NO} {empApproval.EMP_NAME}】";
                                break;
                        }
                        return "";
                    }
                    BonusApprovalListList empsdata = null;
                    if (有被分配部門 == false)
                    {
                        empsdata = _BonusApprovalList.SelectPage_人員_new(false, 0, 100000);
                    }
                    else
                    {
                        empsdata = _BonusApprovalList.SelectPage_人員_new(true, 0, 100000
                              //, 登入者層級: 登入者層級
                              , 登入者: 登入者,
                              BonusProjectID: model.BonusProjectID);
                    }
                    string ms1 = "", ms2 = "";
                    foreach (var empdata in empsdata)
                    {
                        ms1 += 預算送出前檢查(empdata, "年終調整金額已逾加碼預算");
                        ms2 += 預算送出前檢查(empdata, "本次發放金額");
                    }
                    ViewBag.Message = "";
                    if (ms1 != "")
                        ViewBag.Message = ms1.Substring(1) + "的調整金額扣除已逾[加碼預算]+[轄下主管調整金額加總]，請重新填寫[主管調整金額]再送出。";
                    if (ms2 != "")
                        ViewBag.Message = ms2.Substring(1) + "的[本次發放金額]不能小於0。";
                    if (ms1 != "" && ms2 != "")
                        ViewBag.Message = ms1.Substring(1) + "的調整金額扣除已逾[加碼預算]+[轄下主管調整金額加總]，請重新填寫[主管調整金額]再送出。" + ms2.Substring(1) + "的[本次發放金額]不能小於0。";
                    if (_DepartmentBudget.IS登入者分派行為已全數完成_new(model.BonusProjectID, 登入者) == false)
                        ViewBag.Message = "請先將主管分配頁面未決定分派的部門完成。";

                    //沒錯誤訊息才執行送出
                    if (string.IsNullOrEmpty((string)ViewBag.Message))
                    {
                        var ALLApproval = ApprovalAll(DAC.GetInt32(model.BonusProjectID));
                        if (ALLApproval.Approval)
                        {
                            ViewBag.Message = "已送出";
                        }
                        else
                        {
                            ViewBag.Message = "送簽失敗，" + ALLApproval.error;
                        }
                    }
                    break;
            }


            #region 搜尋功能優化
            //員工 - 僅可選擇管轄部門的員工
            var OwnDeptSql = "DEPT_SEQ_PATH = \\'\\' ";
            foreach (var dept in 登入者管理部門Lst.Split(','))
            {
                OwnDeptSql += " OR DEPT_SEQ_PATH LIKE \\'%|" + dept + "|%\\' ";
            }
            ViewBag.SubDeptLst = OwnDeptSql;
            ViewBag.BonusProjectID = model.BonusProjectID;
            #endregion


            #region 計算 (抽離)
            //ViewBag.BonusProjectID = model.BonusProjectID;
            //ViewBag.ProjectName = Project.獎金專案名稱(DAC.GetInt32(model.BonusProjectID));
            //ViewBag.BonusYear = Project.獎金年度(DAC.GetInt32(model.BonusProjectID));
            //ViewBag.ReserveBudget = Project.ReserveBudget(DAC.GetInt32(model.BonusProjectID)) ?? 0;
            ////var Data = Get單位(DAC.GetInt32(model.BonusProjectID), false);  // 效能調整
            ////ViewBag.Salary_D = Data.Sum(p => p.主管分配金額);                 // 效能調整
            ////ViewBag.Salary_V = Data.Sum(p => p.保留金額);                     // 效能調整
            ////model.上階主管分配金額 = Data.Sum(p => p.主管分配金額);               // 效能調整
            //var dList = _DepartmentBudget.Select(BonusProjectID: DAC.GetInt32(model.BonusProjectID), EMP_SEQ_NO: 登入者);
            ////不理解這邊要判斷登入者權限的原因，所以先註解掉，若有問題再進行檢驗
            ////if (DAC.GetInt32(登入者層級) > Project.GetExecutionLevel(BonusProjectID) || dList.Count() == 0)
            //if (dList.Count() == 0)
            //{
            //    有被分配部門 = false;
            //}

            //ViewBag.總預算 = 0; ViewBag.總預算_FixedBudget = 0; ViewBag.總預算_UnFixedBudget = 0; ViewBag.保留金額 = 0;
            //ViewBag.主管微調總金額 = 0; ViewBag.轄下主管調整總額 = 0; ViewBag.已核定總金額 = 0; ViewBag.可用餘額 = 0; ViewBag.btnApproval = false;

            //if (有被分配部門)
            //{
            //    WriteLog($"===計算 總預算_UnFixedBudget[AllotType={DAC.GetString(ViewBag.AllotType)}]=== start");
            //    switch (DAC.GetString(ViewBag.AllotType))
            //    {
            //        case "1":
            //            ViewBag.總預算_UnFixedBudget = new DAC_DepartmentBudget().總預算_UnFixedBudget(DAC.GetInt32(model.BonusProjectID), 登入者);
            //            break;
            //        case "2":
            //            ViewBag.總預算_FixedBudget = new DAC_DepartmentBudget().總預算_FixedBudget(DAC.GetInt32(model.BonusProjectID), 登入者);
            //            ViewBag.總預算_UnFixedBudget = new DAC_DepartmentBudget().總預算_UnFixedBudget(DAC.GetInt32(model.BonusProjectID), 登入者);
            //            break;
            //        case "3":
            //            ViewBag.總預算_UnFixedBudget = new DAC_DepartmentBudget().總預算_UnFixedBudget(DAC.GetInt32(model.BonusProjectID), 登入者);
            //            break;
            //        case "4":
            //            break;
            //        default:
            //            break;
            //    }
            //    WriteLog($"===計算 總預算_UnFixedBudget[AllotType={DAC.GetString(ViewBag.AllotType)}]=== end");
            //    // 效能調整//ViewBag.保留金額 = new DAC_DepartmentBudget().保留金額總合(DAC.GetInt32(model.BonusProjectID), 登入者);
            //    // 效能調整//ViewBag.總預算 = DAC.GetInt32(ViewBag.總預算_FixedBudget) + DAC.GetInt32(ViewBag.總預算_UnFixedBudget) + DAC.GetInt32(ViewBag.保留金額) + DAC.GetInt32(ViewBag.Salary_D);
            //    // 效能調整//ViewBag.主管微調總金額 = _BonusApprovalList.主管微調總金額_new(DAC.GetInt32(model.BonusProjectID), 登入者);
            //    WriteLog($"===計算 主管微調總金額=== end");
            //    //20230617 待確認測試區sp
            //    // 效能調整//ViewBag.轄下主管調整總額 = _BonusApprovalList.轄下主管調整總額_new(DAC.GetInt32(model.BonusProjectID), 登入者);
            //    WriteLog($"===計算 轄下主管調整總額=== end");
            //    // 效能調整//ViewBag.已核定總金額 = DAC.GetInt32(ViewBag.總預算_FixedBudget) + DAC.GetInt32(ViewBag.總預算_UnFixedBudget) + DAC.GetInt32(ViewBag.主管微調總金額) + DAC.GetInt32(ViewBag.轄下主管調整總額);
            //    // 效能調整//ViewBag.可用餘額 = DAC.GetInt32(ViewBag.總預算) - DAC.GetInt32(ViewBag.已核定總金額);
            //    // 效能調整//ViewBag.btnApproval = _DepartmentBudget.Check是否可送簽_new(DAC.GetInt32(model.BonusProjectID), 登入者);
            //    ViewBag.預定簽核數 = _BonusApprovalList.SelectCount(BonusProjectID: DAC.GetInt32(model.BonusProjectID), ReSignerID: DAC.GetInt32(登入者));
            //    // 效能調整//ViewBag.本次簽核人員已全部送簽 = _BonusApprovalList.本次簽核人員已全部送簽(model.BonusProjectID, 登入者);

            //    WriteLog($"===Check是否可送簽=== end");
            //}
            #endregion

            #region 判斷
            var Modify = _DepartmentBudget.Is主管是否可以編輯(DAC.GetInt32(model.BonusProjectID), 登入者);
            string str部門已分派 = "";
            if (Modify.IsModify == false)
            {
                str部門已分派 = "尚有部門 ( ";
                foreach (var d in Modify.deptitem)
                {
                    str部門已分派 += "【" + d.DEPT_NO + " " + d.DEPT_NAME + " 】";
                }
                str部門已分派 += " ) 待上級主管決議是否分派，請待決議後再填寫。";
            }
            ViewBag.btn上層主管已分派完畢 = Modify.IsModify;
            ViewBag.str部門已分派 = str部門已分派;
            WriteLog($"===Is主管是否可以編輯=== end");
            #endregion   

            if (model.Page > 0)
                model.Page--;
            if (model.Page < 0)
                model.Page = 0;
            if (model.PageSize <= 0)
                model.PageSize = 100;

            ViewBag.Order = model.Order;
            ViewBag.Decending = model.Decending.ToString().ToLower();
            var order = string.IsNullOrEmpty(model.Order) ? " LEVEL_CODE, DEPT_NAME, SLY_DEGREE desc " : $"{model.Order} {(model.Decending ? "DESC" : "ASC")}";
            using (var conn = NewConnection())
            {
                conn.Open();
                int? 主管職 = null;
                string 直接間接 = null;
                string deptSeqNoLst = null;

                switch (model.Management_Search)
                {
                    case 1:
                        主管職 = 0;
                        break;
                    case 2:
                        主管職 = 1;
                        break;
                }
                switch (model.DIRECT_Search)
                {
                    case 1:
                        直接間接 = "JO01";
                        break;
                    case 2:
                        直接間接 = "JO02";
                        break;
                }

                if (model.DEPT_SEQ_NO_Search != null)
                {
                    if (model.DEPT_SEQ_NO_IS_UNDER_Search == true)
                    {
                        deptSeqNoLst = string.Join(",", _DEPARTMENT_TW_V
                                                        .GetDEPT_DropDown_CTE(model.DEPT_SEQ_NO_Search)
                                                        .Select(x => "'" + x.Value + "'"));
                    }
                    else
                    {
                        deptSeqNoLst = "'" + model.DEPT_SEQ_NO_Search + "'";
                    }
                    //var tmp = string.Join(",", _DEPARTMENT_TW_V
                    //                           .GetDEPT_DropDown_CTE(model.DEPT_SEQ_NO_Search)
                    //                           .Select(x => "\\'" + x.Value + "\\'"));

                }


                if (有被分配部門 == false)
                {
                    var data = _BonusApprovalList.SelectPage_人員_new(false, orderBy: order).ToPagedList(0, 100, 0);
                    return View(data);
                }
                else
                {                
                    var count = _BonusApprovalList.SelectCount_人員_new(true
                    , 登入者層級: 登入者層級
                    , 登入者: 登入者,
                    BonusProjectID: DAC.GetInt32(model.BonusProjectID),
                    EMP_NO: model.EMP_NO_Search, SLY_DEGREE: DAC.GetString(model.GRADEName_Search),
                    JOBCATEGORY: 直接間接, TITLE_ID: DAC.GetString(model.TITLE_ID_Search),
                    Management: 主管職, DEPT_SEQ_NO: deptSeqNoLst, PERFORMANCE: model.Performance_Search,
                    BonusWorkType: model.BonusWorkType_Search, IsAdjustEmpAmount: model.IsAdjustEmpAmount,
                    cbISApproval: model.cbISApproval);

                    if (count < model.Page * model.PageSize)
                        model.Page = count / model.PageSize;
                    var data = _BonusApprovalList.SelectPage_人員_new(true, model.Page * model.PageSize, model.PageSize
                                , 登入者層級: 登入者層級
                                , 登入者: 登入者,
                                BonusProjectID: DAC.GetInt32(model.BonusProjectID),
                                EMP_NO: model.EMP_NO_Search, SLY_DEGREE: DAC.GetString(model.GRADEName_Search),
                                JOBCATEGORY: 直接間接, TITLE_ID: DAC.GetString(model.TITLE_ID_Search),
                                Management: 主管職, DEPT_SEQ_NO: deptSeqNoLst, PERFORMANCE: model.Performance_Search,
                                BonusWorkType: model.BonusWorkType_Search, IsAdjustEmpAmount: model.IsAdjustEmpAmount,
                                cbISApproval: model.cbISApproval
                                , orderBy: order).ToPagedList(model.Page, model.PageSize, count);

                    switch (order.Split(' ')[0])
                    {
                        case "Performance":
                            data = order.Split(' ')[1] == "ASC"
                                   ? data.OrderBy(x => DAC.GetInt32(x.Performance)).ToPagedList(model.Page, model.PageSize, count)
                                   : data.OrderByDescending(x => DAC.GetInt32(x.Performance)).ToPagedList(model.Page, model.PageSize, count);
                            break;
                        case "BonusBase":
                            data = order.Split(' ')[1] == "ASC"
                                   ? data.OrderBy(x => x.獎金計算基礎).ToPagedList(model.Page, model.PageSize, count)
                                   : data.OrderByDescending(x => x.獎金計算基礎).ToPagedList(model.Page, model.PageSize, count);
                            break;
                        case "PreYearAmount":
                            data = order.Split(' ')[1] == "ASC"
                                   ? data.OrderBy(x => DAC.GetInt32(x.前一年度發放金額)).ToPagedList(model.Page, model.PageSize, count)
                                   : data.OrderByDescending(x => DAC.GetInt32(x.前一年度發放金額)).ToPagedList(model.Page, model.PageSize, count);
                            break;
                        case "PreYearMonth":
                            data = order.Split(' ')[1] == "ASC"
                                   ? data.OrderBy(x => DAC.GetInt32(x.前一年度獎金月數)).ToPagedList(model.Page, model.PageSize, count)
                                   : data.OrderByDescending(x => DAC.GetInt32(x.前一年度獎金月數)).ToPagedList(model.Page, model.PageSize, count);
                            break;
                        case "依公式核定個人預算":
                            data = order.Split(' ')[1] == "ASC"
                                   ? data.OrderBy(x => x.當前加碼預算).ToPagedList(model.Page, model.PageSize, count)
                                   : data.OrderByDescending(x => x.當前加碼預算).ToPagedList(model.Page, model.PageSize, count);
                            break;
                        case "轄下主管調整金額加總":
                            data = order.Split(' ')[1] == "ASC"
                                   ? data.OrderBy(x => DAC.GetInt32(x.轄下主管調整金額加總)).ToPagedList(model.Page, model.PageSize, count)
                                   : data.OrderByDescending(x => DAC.GetInt32(x.轄下主管調整金額加總)).ToPagedList(model.Page, model.PageSize, count);
                            break;
                        case "主管調整金額":
                            data = order.Split(' ')[1] == "ASC"
                                   ? data.OrderBy(x => DAC.GetInt32(x.PreAdjust_current)).ToPagedList(model.Page, model.PageSize, count)
                                   : data.OrderByDescending(x => DAC.GetInt32(x.PreAdjust_current)).ToPagedList(model.Page, model.PageSize, count);
                            break;
                        case "本次發放金額":
                            data = order.Split(' ')[1] == "ASC"
                                   ? data.OrderBy(x => DAC.GetInt32(x.本次發放金額_含當下主管)).ToPagedList(model.Page, model.PageSize, count)
                                   : data.OrderByDescending(x => DAC.GetInt32(x.本次發放金額_含當下主管)).ToPagedList(model.Page, model.PageSize, count);
                            break;
                        case "發放月數":
                            data = order.Split(' ')[1] == "ASC"
                                   ? data.OrderBy(x => DAC.GetInt32(x.本次發放金額_含當下主管) / (DAC.GetDouble(x.獎金計算基礎) * DAC.GetDouble(x.YEAR_SZ))).ToPagedList(model.Page, model.PageSize, count)
                                   : data.OrderByDescending(x => DAC.GetInt32(x.本次發放金額_含當下主管) / (DAC.GetDouble(x.獎金計算基礎) * DAC.GetDouble(x.YEAR_SZ))).ToPagedList(model.Page, model.PageSize, count);
                            break;
                        default:
                            break;                    
                    }
                   
                    //var data = _BonusApprovalList.SelectPage_人員_new(true, model.Page * model.PageSize, model.PageSize
                    //, 登入者層級: 登入者層級
                    //, 登入者: 登入者,
                    //BonusProjectID: DAC.GetInt32(model.BonusProjectID),
                    //EMP_NO: model.EMP_NO_Search, SLY_DEGREE: DAC.GetString(model.GRADEName_Search),
                    //JOBCATEGORY: 直接間接, TITLE_ID: DAC.GetString(model.TITLE_ID_Search),
                    //Management: 主管職, DEPT_SEQ_NO: deptSeqNoLst, PERFORMANCE: model.Performance_Search,
                    //BonusWorkType: model.BonusWorkType_Search,IsAdjustEmpAmount:model.IsAdjustEmpAmount,
                    //cbISApproval: model.cbISApproval
                    //    , orderBy: order)
                    //    .ToPagedList(model.Page, model.PageSize, count);
                    WriteLog($"===SelectPage=== end");
                    return View(data);
                }
            }
        }
        #endregion

        #region BonusDistributionList 主管獎金作業 (舊版)

        [CheckLoginSessionExpired]
        public ActionResult BonusDistributionList_0713(int BonusProjectID)
        {
            #region 宣告
            DAC_FL_DEPARTMENT_TW_V _DEPARTMENT_TW_V = new DAC_FL_DEPARTMENT_TW_V();
            DAC_DepartmentBudget _DepartmentBudget = new DAC_DepartmentBudget();
            DAC_BonusProject Project = new DAC_BonusProject();
            DAC_BonusApprovalList _BonusApprovalList = new DAC_BonusApprovalList();
            string 登入者 = (string)Session[PublicVariable.EMP_SEQ_NO];
            string 登入者層級 = new DAC_FL_DEPARTMENT_TW_V().GETLEVEL_CODE_主管層級(登入者);
            string 登入者DEPT_SEQ_NO = ((string)Session[PublicVariable.department_SEQ_NO]);
            string 登入者所屬部門主管DEPT_SEQ_NO = new DAC_FL_DEPARTMENT_TW_V().GETDEPT_主管部門(登入者);
            bool 有被分配部門 = true;
            #endregion

            #region 計算
            ViewBag.AllotType = Project.GetAllotType(BonusProjectID) ?? 0;
            ViewBag.BonusProjectID = BonusProjectID;
            ViewBag.ProjectName = Project.獎金專案名稱(BonusProjectID);
            ViewBag.BonusYear = Project.獎金年度(BonusProjectID);
            ViewBag.ReserveBudget = Project.ReserveBudget(BonusProjectID) ?? 0;
            var Data = Get單位(DAC.GetInt32(BonusProjectID), false);
            ViewBag.Salary_D = Data.Sum(p => p.主管分配金額);
            ViewBag.Salary_V = Data.Sum(p => p.保留金額);
            var dList = _DepartmentBudget.Select(BonusProjectID: DAC.GetInt32(BonusProjectID), EMP_SEQ_NO: 登入者);
            if (DAC.GetInt32(登入者層級) > Project.GetExecutionLevel(BonusProjectID) || dList.Count() == 0)
            {
                有被分配部門 = false;
            }

            ViewBag.總預算 = 0; ViewBag.總預算_FixedBudget = 0; ViewBag.總預算_UnFixedBudget = 0; ViewBag.保留金額 = 0;
            ViewBag.主管微調總金額 = 0; ViewBag.轄下主管調整總額 = 0; ViewBag.已核定總金額 = 0; ViewBag.可用餘額 = 0; ViewBag.btnApproval = false;
            if (有被分配部門)
            {
                WriteLog($"===計算 總預算_UnFixedBudget[AllotType={DAC.GetString(ViewBag.AllotType)}]=== start");
                switch (DAC.GetString(ViewBag.AllotType))
                {
                    case "1":
                        ViewBag.總預算_UnFixedBudget = new DAC_DepartmentBudget().總預算_UnFixedBudget(DAC.GetInt32(BonusProjectID), 登入者);
                        break;
                    case "2":
                        ViewBag.總預算_FixedBudget = new DAC_DepartmentBudget().總預算_FixedBudget(DAC.GetInt32(BonusProjectID), 登入者);
                        ViewBag.總預算_UnFixedBudget = new DAC_DepartmentBudget().總預算_UnFixedBudget(DAC.GetInt32(BonusProjectID), 登入者);
                        break;
                    case "3":
                        ViewBag.總預算_UnFixedBudget = new DAC_DepartmentBudget().總預算_UnFixedBudget(DAC.GetInt32(BonusProjectID), 登入者);
                        break;
                    case "4":
                        break;
                    default:
                        break;
                }
                WriteLog($"===計算 總預算_UnFixedBudget[AllotType={DAC.GetString(ViewBag.AllotType)}]=== end");
                ViewBag.保留金額 = new DAC_DepartmentBudget().保留金額總合(DAC.GetInt32(BonusProjectID), 登入者);
                ViewBag.總預算 = DAC.GetInt32(ViewBag.總預算_FixedBudget) + DAC.GetInt32(ViewBag.總預算_UnFixedBudget) + DAC.GetInt32(ViewBag.保留金額) + DAC.GetInt32(ViewBag.Salary_D);
                ViewBag.主管微調總金額 = _BonusApprovalList.主管微調總金額(BonusProjectID, 登入者, Level: 登入者層級); //要改
                WriteLog($"===計算 主管微調總金額=== end");
                //20230617 待確認測試區sp
                ViewBag.轄下主管調整總額 = _BonusApprovalList.轄下主管調整總額(DAC.GetInt32(BonusProjectID), 登入者, 登入者所屬部門主管DEPT_SEQ_NO, Level: 登入者層級);
                WriteLog($"===計算 轄下主管調整總額=== end");
                ViewBag.已核定總金額 = DAC.GetInt32(ViewBag.總預算_FixedBudget) + DAC.GetInt32(ViewBag.總預算_UnFixedBudget) + DAC.GetInt32(ViewBag.主管微調總金額) + DAC.GetInt32(ViewBag.轄下主管調整總額);
                ViewBag.可用餘額 = DAC.GetInt32(ViewBag.總預算) - DAC.GetInt32(ViewBag.已核定總金額);
                ViewBag.btnApproval = _DepartmentBudget.Check是否可送簽(BonusProjectID, 登入者, 登入者層級, 登入者所屬部門主管DEPT_SEQ_NO);
                ViewBag.預定簽核數 = _BonusApprovalList.SelectCount(BonusProjectID: BonusProjectID, ReSignerID: DAC.GetInt32(登入者));
                WriteLog($"===Check是否可送簽=== end");
            }
            #endregion

            #region 判斷
            var Modify = _DepartmentBudget.Is主管是否可以編輯(BonusProjectID, 登入者);
            string str部門已分派 = "";
            if (Modify.IsModify == false)
            {
                str部門已分派 = "尚有部門 ( ";
                foreach (var d in Modify.deptitem)
                {
                    str部門已分派 += "【" + d.DEPT_NO + " " + d.DEPT_NAME + " 】";
                }
                str部門已分派 += " ) 待上級主管決議是否分派，請待決議後再填寫。";
            }
            ViewBag.btn上層主管已分派完畢 = Modify.IsModify;
            ViewBag.str部門已分派 = str部門已分派;
            WriteLog($"===Is主管是否可以編輯=== end");
            #endregion

            #region
            //WriteLog($"===CountBudget=== start");
            //_BonusApprovalList.CountBudget(BonusProjectID, 登入者, 登入者所屬部門主管DEPT_SEQ_NO, 登入者層級, DAC.GetInt32(ViewBag.AllotType));
            //WriteLog($"===CountBudget=== end");
            #endregion

            var order = " LEVEL_CODE, DEPT_NAME, SLY_DEGREE desc ";

            //要加上 可編輯核定金額欄位(上階未送簽過或是上階的主管有駁回)
            if (有被分配部門 == false)
            {
                var data = _BonusApprovalList.SelectPage_人員(false, orderBy: order).ToPagedList(0, PublicVariable.DefaultPageSize, 0);
                return View(data);
            }
            else
            {
                var count = _BonusApprovalList.SelectCount_人員(true
                    , 登入者層級: 登入者層級
                    , 登入者: 登入者,
                    登入者DEPT_SEQ_NO: 登入者DEPT_SEQ_NO,
                    BonusProjectID: BonusProjectID);
                var data = _BonusApprovalList.SelectPage_人員(true, 0, PublicVariable.DefaultPageSize
                      , 登入者層級: 登入者層級
                      , 登入者: 登入者,
                      登入者DEPT_SEQ_NO: 登入者DEPT_SEQ_NO,
                      BonusProjectID: BonusProjectID,
                      orderBy: order)
                      .ToPagedList(0, PublicVariable.DefaultPageSize, count);
                WriteLog($"===SelectPage=== end");
                return View(data);
            }
        }
        [CheckLoginSessionExpired]
        [HttpPost]
        public ActionResult BonusDistributionList_0713(BonusApprovalListViewModel model)
        {
            #region 宣告
            DAC_FL_DEPARTMENT_TW_V _DEPARTMENT_TW_V = new DAC_FL_DEPARTMENT_TW_V();
            DAC_DepartmentBudget _DepartmentBudget = new DAC_DepartmentBudget();
            DAC_BonusProject Project = new DAC_BonusProject();
            DAC_BonusApprovalList _BonusApprovalList = new DAC_BonusApprovalList();

            string 登入者 = (string)Session[PublicVariable.EMP_SEQ_NO];
            string 登入者層級 = new DAC_FL_DEPARTMENT_TW_V().GETLEVEL_CODE_主管層級(登入者);
            string 登入者DEPT_SEQ_NO = ((string)Session[PublicVariable.department_SEQ_NO]);
            string 登入者所屬部門主管DEPT_SEQ_NO = new DAC_FL_DEPARTMENT_TW_V().GETDEPT_主管部門(登入者);
            bool 有被分配部門 = true;
            #endregion

            #region 計算
            ViewBag.AllotType = Project.GetAllotType(DAC.GetInt32(model.BonusProjectID)) ?? 0;
            ViewBag.BonusProjectID = model.BonusProjectID;
            ViewBag.ProjectName = Project.獎金專案名稱(DAC.GetInt32(model.BonusProjectID));
            ViewBag.BonusYear = Project.獎金年度(DAC.GetInt32(model.BonusProjectID));
            ViewBag.ReserveBudget = Project.ReserveBudget(DAC.GetInt32(model.BonusProjectID)) ?? 0;
            var Data = Get單位(DAC.GetInt32(model.BonusProjectID), false);
            ViewBag.Salary_D = Data.Sum(p => p.主管分配金額);
            ViewBag.Salary_V = Data.Sum(p => p.保留金額);
            model.上階主管分配金額 = Data.Sum(p => p.主管分配金額);
            var dList = _DepartmentBudget.Select(BonusProjectID: DAC.GetInt32(model.BonusProjectID), EMP_SEQ_NO: 登入者);
            if (DAC.GetInt32(登入者層級) > Project.GetExecutionLevel(model.BonusProjectID) || dList.Count() == 0)
            {
                有被分配部門 = false;
            }

            ViewBag.總預算 = 0; ViewBag.總預算_FixedBudget = 0; ViewBag.總預算_UnFixedBudget = 0; ViewBag.保留金額 = 0;
            ViewBag.主管微調總金額 = 0; ViewBag.轄下主管調整總額 = 0; ViewBag.已核定總金額 = 0; ViewBag.可用餘額 = 0; ViewBag.btnApproval = false;

            if (有被分配部門)
            {
                WriteLog($"===計算 總預算_UnFixedBudget[AllotType={DAC.GetString(ViewBag.AllotType)}]=== start");
                switch (DAC.GetString(ViewBag.AllotType))
                {
                    case "1":
                        ViewBag.總預算_UnFixedBudget = new DAC_DepartmentBudget().總預算_UnFixedBudget(DAC.GetInt32(model.BonusProjectID), 登入者);
                        break;
                    case "2":
                        ViewBag.總預算_FixedBudget = new DAC_DepartmentBudget().總預算_FixedBudget(DAC.GetInt32(model.BonusProjectID), 登入者);
                        ViewBag.總預算_UnFixedBudget = new DAC_DepartmentBudget().總預算_UnFixedBudget(DAC.GetInt32(model.BonusProjectID), 登入者);
                        break;
                    case "3":
                        ViewBag.總預算_UnFixedBudget = new DAC_DepartmentBudget().總預算_UnFixedBudget(DAC.GetInt32(model.BonusProjectID), 登入者);
                        break;
                    case "4":
                        break;
                    default:
                        break;
                }
                WriteLog($"===計算 總預算_UnFixedBudget[AllotType={DAC.GetString(ViewBag.AllotType)}]=== end");
                ViewBag.保留金額 = new DAC_DepartmentBudget().保留金額總合(DAC.GetInt32(model.BonusProjectID), 登入者);
                ViewBag.總預算 = DAC.GetInt32(ViewBag.總預算_FixedBudget) + DAC.GetInt32(ViewBag.總預算_UnFixedBudget) + DAC.GetInt32(ViewBag.保留金額) + DAC.GetInt32(ViewBag.Salary_D);
                ViewBag.主管微調總金額 = _BonusApprovalList.主管微調總金額(DAC.GetInt32(model.BonusProjectID), 登入者, Level: 登入者層級);
                WriteLog($"===計算 主管微調總金額=== end");
                //20230617 待確認測試區sp
                ViewBag.轄下主管調整總額 = _BonusApprovalList.轄下主管調整總額(DAC.GetInt32(model.BonusProjectID), 登入者, 登入者所屬部門主管DEPT_SEQ_NO, Level: 登入者層級);
                WriteLog($"===計算 轄下主管調整總額=== end");
                ViewBag.已核定總金額 = DAC.GetInt32(ViewBag.總預算_FixedBudget) + DAC.GetInt32(ViewBag.總預算_UnFixedBudget) + DAC.GetInt32(ViewBag.主管微調總金額) + DAC.GetInt32(ViewBag.轄下主管調整總額);
                ViewBag.可用餘額 = DAC.GetInt32(ViewBag.總預算) - DAC.GetInt32(ViewBag.已核定總金額);
            }
            #endregion

            #region 判斷
            var Modify = _DepartmentBudget.Is主管是否可以編輯(DAC.GetInt32(model.BonusProjectID), 登入者);
            string str部門已分派 = "";
            if (Modify.IsModify == false)
            {
                str部門已分派 = "尚有部門 ( ";
                foreach (var d in Modify.deptitem)
                {
                    str部門已分派 += "【" + d.DEPT_NO + " " + d.DEPT_NAME + " 】";
                }
                str部門已分派 += " ) 待上級主管決議是否分派，請待決議後再填寫。";
            }
            ViewBag.btn上層主管已分派完畢 = Modify.IsModify;
            ViewBag.str部門已分派 = str部門已分派;
            WriteLog($"===Is主管是否可以編輯=== end");
            #endregion

            switch (model.SubmitButton)
            {
                case "不核定預算為0":
                    var Approval = ApprovalnZero(DAC.GetInt32(model.BonusProjectID));
                    if (Approval.Approval)
                    {
                        ViewBag.Message = "[不核定預算為0]寫入成功";
                    }
                    else
                    {
                        ViewBag.Message = "[不核定預算為0]寫入失敗，" + Approval.error;
                    }
                    break;
                case "送出":
                    string 預算送出前檢查(BonusApprovalListItem empApproval, string checkType)
                    {
                        switch (checkType)
                        {
                            case "年終調整金額已逾加碼預算":
                                //只有年終獎金才做處理
                                if ((int)ViewBag.AllotType != 2)
                                    return "";
                                if (DAC.GetInt64(empApproval.PreAdjust_current) + DAC.GetInt64(empApproval.轄下主管調整金額加總) + DAC.GetInt64(empApproval.當前加碼預算) < 0)
                                {
                                    return $"、【{empApproval.EMP_NO} {empApproval.EMP_NAME}】";
                                }
                                break;
                            case "本次發放金額":       
                                //[本次發放金額]不能小於0。"
                                if (DAC.GetInt64(empApproval.本次發放金額) + DAC.GetInt64(empApproval.PreAdjust_current) < 0)
                                    return $"、【{empApproval.EMP_NO} {empApproval.EMP_NAME}】";
                                break;
                        }
                        return "";
                    }
                    BonusApprovalListList empsdata = null;
                    if (有被分配部門 == false)
                    {
                        empsdata = _BonusApprovalList.SelectPage_人員(false, 0, 100000);
                    }
                    else
                    {
                        empsdata = _BonusApprovalList.SelectPage_人員(true, 0, 100000
                              , 登入者層級: 登入者層級
                              , 登入者: 登入者,
                              登入者DEPT_SEQ_NO: 登入者DEPT_SEQ_NO,
                              BonusProjectID: model.BonusProjectID);
                    }
                    string ms1 = "", ms2 = "";
                    foreach(var empdata in empsdata)
                    {
                        ms1 += 預算送出前檢查(empdata, "年終調整金額已逾加碼預算");
                        ms2 += 預算送出前檢查(empdata, "本次發放金額");
                    }
                    ViewBag.Message = "";
                    if (ms1 != "")
                        ViewBag.Message = ms1.Substring(1) + "的調整金額扣除已逾[加碼預算]+[轄下主管調整金額加總]，請重新填寫[主管調整金額]再送出。";
                    if (ms2 != "")
                        ViewBag.Message = ms2.Substring(1) + "的[本次發放金額]不能小於0。";
                    if (ms1 != "" && ms2 != "")
                        ViewBag.Message = ms1.Substring(1) + "的調整金額扣除已逾[加碼預算]+[轄下主管調整金額加總]，請重新填寫[主管調整金額]再送出。" + ms2.Substring(1) + "的[本次發放金額]不能小於0。";
                    if(_DepartmentBudget.IS登入者分派行為已全數完成(model.BonusProjectID, 登入者, 登入者層級) == false)
                        ViewBag.Message = "請先將主管分配頁面未決定分派的部門完成。";

                    //沒錯誤訊息才執行送出
                    if (string.IsNullOrEmpty((string)ViewBag.Message))
                    {
                        var ALLApproval = ApprovalAll(DAC.GetInt32(model.BonusProjectID));
                        if (ALLApproval.Approval)
                        {
                            ViewBag.Message = "已送出";
                        }
                        else
                        {
                            ViewBag.Message = "送簽失敗，" + ALLApproval.error;
                        }
                    }
                    break;
            }

            ViewBag.預定簽核數 = _BonusApprovalList.SelectCount(BonusProjectID: DAC.GetInt32(model.BonusProjectID), ReSignerID: DAC.GetInt32(登入者));
            if (有被分配部門)
            {
                ViewBag.btnApproval = _DepartmentBudget.Check是否可送簽(DAC.GetInt32(model.BonusProjectID), 登入者, 登入者層級, 登入者所屬部門主管DEPT_SEQ_NO);
                WriteLog($"===Check是否可送簽=== end");
            }
            if (model.Page > 0)
                model.Page--;
            if (model.Page < 0)
                model.Page = 0;
            if (model.PageSize <= 0)
                model.PageSize = PublicVariable.DefaultPageSize;

            ViewBag.Order = model.Order;
            ViewBag.Decending = model.Decending.ToString().ToLower();
            var order = string.IsNullOrEmpty(model.Order) ? " LEVEL_CODE, DEPT_NAME, SLY_DEGREE desc " : $"{model.Order} {(model.Decending ? "DESC" : "ASC")}";
            using (var conn = NewConnection())
            {
                conn.Open();
                int? 主管職 = null;
                string 直接間接 = null;
                switch (model.Management_Search)
                {
                    case 1:
                        主管職 = 0;
                        break;
                    case 2:
                        主管職 = 1;
                        break;
                }
                switch (model.DIRECT_Search)
                {
                    case 1:
                        直接間接 = "JO01";
                        break;
                    case 2:
                        直接間接 = "JO02";
                        break;
                }

                if (有被分配部門 == false)
                {
                    var data = _BonusApprovalList.SelectPage_人員(false, orderBy: order).ToPagedList(0, PublicVariable.DefaultPageSize, 0);
                    return View(data);
                }
                else
                {
                    var count = _BonusApprovalList.SelectCount_人員(true
                    , 登入者層級: 登入者層級
                    , 登入者: 登入者,
                    登入者DEPT_SEQ_NO: 登入者DEPT_SEQ_NO,
                    BonusProjectID: DAC.GetInt32(model.BonusProjectID),
                    EMP_NO: model.EMP_NO_Search, SLY_DEGREE: DAC.GetString(model.GRADEName_Search),
                    JOBCATEGORY: 直接間接, TITLE_ID: DAC.GetString(model.TITLE_ID_Search),
                    Management: 主管職, DEPT_SEQ_NO: model.DEPT_SEQ_NO_Search);

                    if (count < model.Page * model.PageSize)
                        model.Page = count / model.PageSize;
                    var data = _BonusApprovalList.SelectPage_人員(true, model.Page * model.PageSize, model.PageSize
                    , 登入者層級: 登入者層級
                    , 登入者: 登入者,
                    登入者DEPT_SEQ_NO: 登入者DEPT_SEQ_NO,
                    BonusProjectID: DAC.GetInt32(model.BonusProjectID),
                    EMP_NO: model.EMP_NO_Search, SLY_DEGREE: DAC.GetString(model.GRADEName_Search),
                    JOBCATEGORY: 直接間接, TITLE_ID: DAC.GetString(model.TITLE_ID_Search),
                    Management: 主管職, DEPT_SEQ_NO: model.DEPT_SEQ_NO_Search
                        , orderBy: order)
                        .ToPagedList(model.Page, model.PageSize, count);
                    WriteLog($"===SelectPage=== end");
                    return View(data);
                }
            }
        }
        #endregion

        #region  主管核定送簽

        [CheckLoginSessionExpired]
        public ActionResult BonusDistributionReview(int BonusProjectID)
        {
            DAC_BonusProject Project = new DAC_BonusProject();
            DAC_DepartmentBudget _DepartmentBudget = new DAC_DepartmentBudget();

            string 登入者 = (string)Session[PublicVariable.EMP_SEQ_NO];
            //所有的登入者層級帶入資料，皆需要依照部門來看該登入者所擁有的最高層級，而非直接用登入者來看
            //string 登入者層級 = new DAC_FL_DEPARTMENT_TW_V().GETLEVEL_CODE_主管層級(登入者);
            string 登入者DEPT_SEQ_NO = ((string)Session[PublicVariable.department_SEQ_NO]);

            ViewBag.BonusProjectID = BonusProjectID;
            ViewBag.ProjectName = Project.獎金專案名稱(BonusProjectID);
            ViewBag.HighLevelDepts = _DepartmentBudget.GetDepartmentBudget_兼任部門對應的預算(BonusProjectID, 登入者);
            ViewBag.IS總經理 = false;

            //如果非固定預算(彈性預算=2以及複合預算=3)，則判斷是否第一次操作而尚未儲存過。
            ViewBag.BudgetType = Project.SelectOne(BonusProjectID).FirstOrDefault() != null ? Project.SelectOne(BonusProjectID).FirstOrDefault().BudgetType : 0;
            ViewBag.IsFirstWorkTime = _DepartmentBudget.SelectAll()
                                      .Count(x => x.BonusProjectID == BonusProjectID
                                              && x.EMP_SEQ_NO == 登入者
                                              && x.DB_APPNO == 1) == _DepartmentBudget.SelectAll()
                                      .Count(x => x.BonusProjectID == BonusProjectID
                                              && x.EMP_SEQ_NO == 登入者);

            //如果總經理包含登入者
            if (new DAC_FL_DEPARTMENT_TW_V().Select_總經理().Where(p => p.MAIN_LEADER_EMP_NO == (string)Session[PublicVariable.UserId]).Count() > 0)
                ViewBag.IS總經理 = true;
            var Modify = _DepartmentBudget.Is主管是否可以編輯(BonusProjectID, 登入者);
            string str部門已分派 = "";

            if (Modify.IsModify == false)
            {
                str部門已分派 = "尚有部門 ( ";
                foreach (var d in Modify.deptitem)
                {
                    str部門已分派 += "【" + d.DEPT_NO + " " + d.DEPT_NAME + " 】";
                }
                str部門已分派 += " ) 待上級主管決議是否分派，請待決議後再填寫。";
            }
            var Data = Get單位(DAC.GetInt32(BonusProjectID), false);
            var model = new BonusApprovalListViewModel()
            {
                BonusProjectID = BonusProjectID,
                BonusDistributionData = Data,
                ReserveBudget = Project.ReserveBudget(BonusProjectID) ?? 0,
                AllotType = Project.GetAllotType(BonusProjectID) ?? 0,
                budget = _DepartmentBudget.保留百分比(BonusProjectID, 登入者),
                btn登入者未進行分派 = _DepartmentBudget.IS登入者未進行分派(BonusProjectID, 登入者),
                btn保留百分比 = _DepartmentBudget.IS編輯保留百分比_new(DAC.GetInt32(BonusProjectID), 登入者, 登入者DEPT_SEQ_NO),
                btn上層主管已分派完畢 = Modify.IsModify,
                str部門已分派 = str部門已分派,
                上階主管分配金額 = Data.Sum(p => p.主管分配金額)
            };

            SysLog.Write(LoginUserID, 獎金分配Page, SysLog.IntoPageLog(BonusProjectID, pageMode: SysLog.進入頁面.頁面, intoSuccess: true));
            return View(model);
        }

        [CheckLoginSessionExpired]
        [HttpPost]
        public ActionResult BonusDistributionReview(BonusApprovalListViewModel model)
        {
            DAC_BonusProject Project = new DAC_BonusProject();
            DAC_DepartmentBudget _DepartmentBudget = new DAC_DepartmentBudget();
            string 登入者 = (string)Session[PublicVariable.EMP_SEQ_NO];
            //string 登入者層級 = new DAC_FL_DEPARTMENT_TW_V().GETLEVEL_CODE_主管層級(登入者);
            string 登入者DEPT_SEQ_NO = ((string)Session[PublicVariable.department_SEQ_NO]);
            ViewBag.BonusProjectID = model.BonusProjectID;
            ViewBag.ProjectName = Project.獎金專案名稱(DAC.GetInt32(model.BonusProjectID));
            ViewBag.HighLevelDepts = _DepartmentBudget.GetDepartmentBudget_兼任部門對應的預算(model.BonusProjectID, 登入者);
            ViewBag.IS總經理 = false;
            //第一次操作，尚未儲存過。
            ViewBag.IsFirstWorkTime = _DepartmentBudget.SelectAll()
                                      .Count(x => x.BonusProjectID == model.BonusProjectID
                                              && x.EMP_SEQ_NO == 登入者
                                              && x.DB_APPNO == 1) == _DepartmentBudget.SelectAll()
                                      .Count(x => x.BonusProjectID == model.BonusProjectID
                                              && x.EMP_SEQ_NO == 登入者);

            //如果總經理包含登入者
            if (new DAC_FL_DEPARTMENT_TW_V().Select_總經理().Where(p => p.MAIN_LEADER_EMP_NO == (string)Session[PublicVariable.UserId]).Count() > 0)
                ViewBag.IS總經理 = true;
            var Modify = _DepartmentBudget.Is主管是否可以編輯(DAC.GetInt32(model.BonusProjectID), 登入者);
            string str部門已分派 = "";
            if (Modify.IsModify == false)
            {
                str部門已分派 = "尚有部門 ( ";
                foreach (var d in Modify.deptitem)
                {
                    str部門已分派 += "【" + d.DEPT_NO + " " + d.DEPT_NAME + " 】";
                }
                str部門已分派 = " ) 待上級主管決議是否分派，請待決議後再填寫。";
            }

            switch (model.SubmitButton)
            {
                case "存檔": //存保留百分比
                    if (!CheckControl.Is正整數(model.budget))
                    {
                        ViewBag.Message = "存檔失敗，原因 : 請輸入正整數";
                    }
                    else if (DAC.GetInt32(model.budget) < 0 || DAC.GetInt32(model.budget) > 99)
                    {
                        ViewBag.Message = "存檔失敗，原因 : 僅能輸入0 ~ 99";
                    }
                    else
                    {
                        Decimal 保留百分比 = DAC.GetDecimal(DAC.GetDecimal(model.budget) / 100);
                        DAC_FL_DEPARTMENT_TW_V _DEPARTMENT_TW_V = new DAC_FL_DEPARTMENT_TW_V();
                        DAC_BonusApprovalList _BonusApprovalList = new DAC_BonusApprovalList();

                        switch (Project.GetAllotType(DAC.GetInt32(model.BonusProjectID)))
                        {
                            case 1:
                            case 2:
                            case 3:
                                var dList_親核 = _DepartmentBudget.Select(BonusProjectID: model.BonusProjectID, EMP_SEQ_NO: 登入者, ISApproval: true);
                                var dList = _DepartmentBudget.Select(BonusProjectID: model.BonusProjectID, EMP_SEQ_NO: 登入者, ISApproval: false);
                                #region 1、select DepartmentBudget 親核、轄下部門
                                foreach (var D_親核 in dList_親核)
                                {
                                    #region 2、算BonusApprovalList
                                    int 保留款金額 = 0;
                                    var 人員List = _BonusApprovalList.Select_登入者所負責的人員名單(BonusProjectID: model.BonusProjectID, Login_EMP_SEQ_NO: 登入者, DEPT_SEQ_NO: D_親核.DEPT_SEQ_NO, false);
                                    //var 人員List = _BonusApprovalList.Select_轄下人員(BonusProjectID: model.BonusProjectID, EMP_SEQ_NO: 登入者, DEPT_SEQ_NO: D_親核.DEPT_SEQ_NO);
                                    foreach (var P in 人員List)
                                    {
                                        int UnFixedBudget_org = 0;
                                        #region 判斷上一層主管的加碼金額
                                        switch (P.LEVEL_CODE_current)
                                        {
                                            case "1":
                                                UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget));
                                                break;
                                            case "2":
                                                if (P.UnFixedBudget1 != null && P.UnFixedBudget1 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget1));
                                                else
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget));
                                                break;
                                            case "3":
                                                if (P.UnFixedBudget2 != null && P.UnFixedBudget2 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget2));
                                                else if (P.UnFixedBudget1 != null && P.UnFixedBudget1 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget1));
                                                else
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget));
                                                break;
                                            case "4":
                                                if (P.UnFixedBudget3 != null && P.UnFixedBudget3 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget3));
                                                else if (P.UnFixedBudget2 != null && P.UnFixedBudget2 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget2));
                                                else if (P.UnFixedBudget1 != null && P.UnFixedBudget1 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget1));
                                                else
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget));
                                                break;
                                            case "5":
                                                if (P.UnFixedBudget4 != null && P.UnFixedBudget4 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget4));
                                                else if (P.UnFixedBudget3 != null && P.UnFixedBudget3 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget3));
                                                else if (P.UnFixedBudget2 != null && P.UnFixedBudget2 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget2));
                                                else if (P.UnFixedBudget1 != null && P.UnFixedBudget1 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget1));
                                                else
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget));
                                                break;
                                            case "6":
                                                if (P.UnFixedBudget5 != null && P.UnFixedBudget5 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget5));
                                                else if (P.UnFixedBudget4 != null && P.UnFixedBudget4 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget4));
                                                else if (P.UnFixedBudget3 != null && P.UnFixedBudget3 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget3));
                                                else if (P.UnFixedBudget2 != null && P.UnFixedBudget2 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget2));
                                                else if (P.UnFixedBudget1 != null && P.UnFixedBudget1 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget1));
                                                else
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget));
                                                break;
                                            case "7":
                                                if (P.UnFixedBudget6 != null && P.UnFixedBudget6 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget6));
                                                else if (P.UnFixedBudget5 != null && P.UnFixedBudget5 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget5));
                                                else if (P.UnFixedBudget4 != null && P.UnFixedBudget4 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget4));
                                                else if (P.UnFixedBudget3 != null && P.UnFixedBudget3 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget3));
                                                else if (P.UnFixedBudget2 != null && P.UnFixedBudget2 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget2));
                                                else if (P.UnFixedBudget1 != null && P.UnFixedBudget1 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget1));
                                                else
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget));
                                                break;
                                            case "8":
                                                if (P.UnFixedBudget7 != null && P.UnFixedBudget7 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget7));
                                                else if (P.UnFixedBudget6 != null && P.UnFixedBudget6 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget6));
                                                else if (P.UnFixedBudget5 != null && P.UnFixedBudget5 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget5));
                                                else if (P.UnFixedBudget4 != null && P.UnFixedBudget4 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget4));
                                                else if (P.UnFixedBudget3 != null && P.UnFixedBudget3 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget3));
                                                else if (P.UnFixedBudget2 != null && P.UnFixedBudget2 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget2));
                                                else if (P.UnFixedBudget1 != null && P.UnFixedBudget1 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget1));
                                                else
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget));
                                                break;
                                            case "9":
                                                if (P.UnFixedBudget8 != null && P.UnFixedBudget8 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget8));
                                                else if (P.UnFixedBudget7 != null && P.UnFixedBudget7 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget7));
                                                else if (P.UnFixedBudget6 != null && P.UnFixedBudget6 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget6));
                                                else if (P.UnFixedBudget5 != null && P.UnFixedBudget5 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget5));
                                                else if (P.UnFixedBudget4 != null && P.UnFixedBudget4 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget4));
                                                else if (P.UnFixedBudget3 != null && P.UnFixedBudget3 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget3));
                                                else if (P.UnFixedBudget2 != null && P.UnFixedBudget2 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget2));
                                                else if (P.UnFixedBudget1 != null && P.UnFixedBudget1 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget1));
                                                else
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget));
                                                break;
                                            case "10":
                                                if (P.UnFixedBudget9 != null && P.UnFixedBudget9 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget9));
                                                else if (P.UnFixedBudget8 != null && P.UnFixedBudget8 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget8));
                                                else if (P.UnFixedBudget7 != null && P.UnFixedBudget7 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget7));
                                                else if (P.UnFixedBudget6 != null && P.UnFixedBudget6 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget6));
                                                else if (P.UnFixedBudget5 != null && P.UnFixedBudget5 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget5));
                                                else if (P.UnFixedBudget4 != null && P.UnFixedBudget4 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget4));
                                                else if (P.UnFixedBudget3 != null && P.UnFixedBudget3 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget3));
                                                else if (P.UnFixedBudget2 != null && P.UnFixedBudget2 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget2));
                                                else if (P.UnFixedBudget1 != null && P.UnFixedBudget1 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget1));
                                                else
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget));
                                                break;
                                            default:
                                                break;
                                        }
                                        #endregion
                                        int UnFixedBudget_保留後金額 = DAC.GetInt32(Math.Round(UnFixedBudget_org * (1 - 保留百分比), 0, MidpointRounding.AwayFromZero));
                                        保留款金額 += UnFixedBudget_org - UnFixedBudget_保留後金額;
                                        #region 寫入該層級UnFixedBudget
                                        switch (P.LEVEL_CODE_current)
                                        {
                                            case "1":
                                                P.UnFixedBudget1 = StringEncrypt.aesEncryptBase64(DAC.GetString(UnFixedBudget_保留後金額)); //加密存回去
                                                break;
                                            case "2":
                                                P.UnFixedBudget2 = StringEncrypt.aesEncryptBase64(DAC.GetString(UnFixedBudget_保留後金額)); //加密存回去
                                                break;
                                            case "3":
                                                P.UnFixedBudget3 = StringEncrypt.aesEncryptBase64(DAC.GetString(UnFixedBudget_保留後金額)); //加密存回去
                                                break;
                                            case "4":
                                                P.UnFixedBudget4 = StringEncrypt.aesEncryptBase64(DAC.GetString(UnFixedBudget_保留後金額)); //加密存回去
                                                break;
                                            case "5":
                                                P.UnFixedBudget5 = StringEncrypt.aesEncryptBase64(DAC.GetString(UnFixedBudget_保留後金額)); //加密存回去
                                                break;
                                            case "6":
                                                P.UnFixedBudget6 = StringEncrypt.aesEncryptBase64(DAC.GetString(UnFixedBudget_保留後金額)); //加密存回去
                                                break;
                                            case "7":
                                                P.UnFixedBudget7 = StringEncrypt.aesEncryptBase64(DAC.GetString(UnFixedBudget_保留後金額)); //加密存回去
                                                break;
                                            case "8":
                                                P.UnFixedBudget8 = StringEncrypt.aesEncryptBase64(DAC.GetString(UnFixedBudget_保留後金額)); //加密存回去
                                                break;
                                            case "9":
                                                P.UnFixedBudget9 = StringEncrypt.aesEncryptBase64(DAC.GetString(UnFixedBudget_保留後金額)); //加密存回去
                                                break;
                                            case "10":
                                                P.UnFixedBudget10 = StringEncrypt.aesEncryptBase64(DAC.GetString(UnFixedBudget_保留後金額)); //加密存回去
                                                break;
                                            default:
                                                break;
                                        }
                                        #endregion
                                        _BonusApprovalList.UpdateOne(P);
                                    }
                                    #endregion

                                    #region 3、寫入DepartmentBudget.ReserveBudgetRatio && DepartmentBudget.ReserveBudget
                                    int DUnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(D_親核.UnFixedBudget));
                                    int DFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(D_親核.FixedBudget));

                                    D_親核.ReserveBudgetRatio = 保留百分比; //保留百分比
                                    D_親核.ReserveBudget = StringEncrypt.aesEncryptBase64(DAC.GetString(保留款金額)); //保留款金額
                                    D_親核.Amount = StringEncrypt.aesEncryptBase64(DAC.GetString(DFixedBudget_org + DUnFixedBudget_org - 保留款金額));
                                    _DepartmentBudget.UpdateOne(D_親核);

                                    #endregion

                                }
                                foreach (var D in dList)
                                {
                                    #region 2、算BonusApprovalList
                                    int 保留款金額 = 0;
                                    var 人員List = _BonusApprovalList.Select_登入者所負責的人員名單(BonusProjectID: model.BonusProjectID, Login_EMP_SEQ_NO: 登入者, DEPT_SEQ_NO: D.DEPT_SEQ_NO, true);
                                    //var 人員List = _BonusApprovalList.Select_轄下人員(BonusProjectID: model.BonusProjectID, EMP_SEQ_NO: 登入者, DEPT_SEQ_NO: _DEPARTMENT_TW_V.GetChild_DEPT_SEQ_NO_List(D.DEPT_SEQ_NO, DAC_FL_DEPARTMENT_TW_V.資訊類型.代號));
                                    foreach (var P in 人員List)
                                    {
                                        int UnFixedBudget_org = 0;
                                        #region 判斷上一層主管的加碼金額
                                        switch (P.LEVEL_CODE_current)
                                        {
                                            case "1":
                                                UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget));
                                                break;
                                            case "2":
                                                if (P.UnFixedBudget1 != null && P.UnFixedBudget1 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget1));
                                                else
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget));
                                                break;
                                            case "3":
                                                if (P.UnFixedBudget2 != null && P.UnFixedBudget2 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget2));
                                                else if (P.UnFixedBudget1 != null && P.UnFixedBudget1 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget1));
                                                else
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget));
                                                break;
                                            case "4":
                                                if (P.UnFixedBudget3 != null && P.UnFixedBudget3 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget3));
                                                else if (P.UnFixedBudget2 != null && P.UnFixedBudget2 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget2));
                                                else if (P.UnFixedBudget1 != null && P.UnFixedBudget1 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget1));
                                                else
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget));
                                                break;
                                            case "5":
                                                if (P.UnFixedBudget4 != null && P.UnFixedBudget4 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget4));
                                                else if (P.UnFixedBudget3 != null && P.UnFixedBudget3 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget3));
                                                else if (P.UnFixedBudget2 != null && P.UnFixedBudget2 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget2));
                                                else if (P.UnFixedBudget1 != null && P.UnFixedBudget1 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget1));
                                                else
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget));
                                                break;
                                            case "6":
                                                if (P.UnFixedBudget5 != null && P.UnFixedBudget5 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget5));
                                                else if (P.UnFixedBudget4 != null && P.UnFixedBudget4 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget4));
                                                else if (P.UnFixedBudget3 != null && P.UnFixedBudget3 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget3));
                                                else if (P.UnFixedBudget2 != null && P.UnFixedBudget2 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget2));
                                                else if (P.UnFixedBudget1 != null && P.UnFixedBudget1 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget1));
                                                else
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget));
                                                break;
                                            case "7":
                                                if (P.UnFixedBudget6 != null && P.UnFixedBudget6 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget6));
                                                else if (P.UnFixedBudget5 != null && P.UnFixedBudget5 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget5));
                                                else if (P.UnFixedBudget4 != null && P.UnFixedBudget4 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget4));
                                                else if (P.UnFixedBudget3 != null && P.UnFixedBudget3 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget3));
                                                else if (P.UnFixedBudget2 != null && P.UnFixedBudget2 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget2));
                                                else if (P.UnFixedBudget1 != null && P.UnFixedBudget1 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget1));
                                                else
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget));
                                                break;
                                            case "8":
                                                if (P.UnFixedBudget7 != null && P.UnFixedBudget7 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget7));
                                                else if (P.UnFixedBudget6 != null && P.UnFixedBudget6 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget6));
                                                else if (P.UnFixedBudget5 != null && P.UnFixedBudget5 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget5));
                                                else if (P.UnFixedBudget4 != null && P.UnFixedBudget4 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget4));
                                                else if (P.UnFixedBudget3 != null && P.UnFixedBudget3 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget3));
                                                else if (P.UnFixedBudget2 != null && P.UnFixedBudget2 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget2));
                                                else if (P.UnFixedBudget1 != null && P.UnFixedBudget1 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget1));
                                                else
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget));
                                                break;
                                            case "9":
                                                if (P.UnFixedBudget8 != null && P.UnFixedBudget8 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget8));
                                                else if (P.UnFixedBudget7 != null && P.UnFixedBudget7 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget7));
                                                else if (P.UnFixedBudget6 != null && P.UnFixedBudget6 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget6));
                                                else if (P.UnFixedBudget5 != null && P.UnFixedBudget5 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget5));
                                                else if (P.UnFixedBudget4 != null && P.UnFixedBudget4 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget4));
                                                else if (P.UnFixedBudget3 != null && P.UnFixedBudget3 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget3));
                                                else if (P.UnFixedBudget2 != null && P.UnFixedBudget2 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget2));
                                                else if (P.UnFixedBudget1 != null && P.UnFixedBudget1 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget1));
                                                else
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget));
                                                break;
                                            case "10":
                                                if (P.UnFixedBudget9 != null && P.UnFixedBudget9 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget9));
                                                else if (P.UnFixedBudget8 != null && P.UnFixedBudget8 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget8));
                                                else if (P.UnFixedBudget7 != null && P.UnFixedBudget7 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget7));
                                                else if (P.UnFixedBudget6 != null && P.UnFixedBudget6 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget6));
                                                else if (P.UnFixedBudget5 != null && P.UnFixedBudget5 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget5));
                                                else if (P.UnFixedBudget4 != null && P.UnFixedBudget4 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget4));
                                                else if (P.UnFixedBudget3 != null && P.UnFixedBudget3 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget3));
                                                else if (P.UnFixedBudget2 != null && P.UnFixedBudget2 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget2));
                                                else if (P.UnFixedBudget1 != null && P.UnFixedBudget1 != "")
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget1));
                                                else
                                                    UnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(P.UnFixedBudget));
                                                break;
                                            default:
                                                break;
                                        }
                                        #endregion
                                        int UnFixedBudget_保留後金額 = DAC.GetInt32(Math.Round(UnFixedBudget_org * (1 - 保留百分比), 0, MidpointRounding.AwayFromZero));
                                        保留款金額 += UnFixedBudget_org - UnFixedBudget_保留後金額;
                                        #region 寫入該層級UnFixedBudget
                                        switch (P.LEVEL_CODE_current)
                                        {
                                            case "1":
                                                P.UnFixedBudget1 = StringEncrypt.aesEncryptBase64(DAC.GetString(UnFixedBudget_保留後金額)); //加密存回去
                                                break;
                                            case "2":
                                                P.UnFixedBudget2 = StringEncrypt.aesEncryptBase64(DAC.GetString(UnFixedBudget_保留後金額)); //加密存回去
                                                break;
                                            case "3":
                                                P.UnFixedBudget3 = StringEncrypt.aesEncryptBase64(DAC.GetString(UnFixedBudget_保留後金額)); //加密存回去
                                                break;
                                            case "4":
                                                P.UnFixedBudget4 = StringEncrypt.aesEncryptBase64(DAC.GetString(UnFixedBudget_保留後金額)); //加密存回去
                                                break;
                                            case "5":
                                                P.UnFixedBudget5 = StringEncrypt.aesEncryptBase64(DAC.GetString(UnFixedBudget_保留後金額)); //加密存回去
                                                break;
                                            case "6":
                                                P.UnFixedBudget6 = StringEncrypt.aesEncryptBase64(DAC.GetString(UnFixedBudget_保留後金額)); //加密存回去
                                                break;
                                            case "7":
                                                P.UnFixedBudget7 = StringEncrypt.aesEncryptBase64(DAC.GetString(UnFixedBudget_保留後金額)); //加密存回去
                                                break;
                                            case "8":
                                                P.UnFixedBudget8 = StringEncrypt.aesEncryptBase64(DAC.GetString(UnFixedBudget_保留後金額)); //加密存回去
                                                break;
                                            case "9":
                                                P.UnFixedBudget9 = StringEncrypt.aesEncryptBase64(DAC.GetString(UnFixedBudget_保留後金額)); //加密存回去
                                                break;
                                            case "10":
                                                P.UnFixedBudget10 = StringEncrypt.aesEncryptBase64(DAC.GetString(UnFixedBudget_保留後金額)); //加密存回去
                                                break;
                                            default:
                                                break;
                                        }
                                        #endregion
                                        _BonusApprovalList.UpdateOne(P);
                                    }
                                    #endregion

                                    #region 3、寫入DepartmentBudget.ReserveBudgetRatio && DepartmentBudget.ReserveBudget
                                    int DUnFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(D.UnFixedBudget));
                                    int DFixedBudget_org = DAC.GetInt32(StringEncrypt.aesDecryptBase64(D.FixedBudget));
                                    D.ReserveBudgetRatio = 保留百分比; //保留百分比
                                    D.ReserveBudget = StringEncrypt.aesEncryptBase64(DAC.GetString(保留款金額)); //保留款金額
                                    D.Amount = StringEncrypt.aesEncryptBase64(DAC.GetString(DFixedBudget_org + DUnFixedBudget_org - 保留款金額));
                                    _DepartmentBudget.UpdateOne(D);
                                    #endregion
                                }
                                #endregion
                                break;
                            case 4:
                                break;
                            default:
                                break;
                        }
                    }
                    break;
                case "儲存分配金額":
                    {
                        long 主管分配金額加總 = model.BonusDistributionData.Sum(p => DAC.GetInt64(p.主管分配金額str.Replace(",", "")));
                        if (主管分配金額加總 > model.上階主管分配金額)
                        {
                            ViewBag.Message1 = "儲存失敗！超過可分配預算";
                        }
                        if (主管分配金額加總 < model.上階主管分配金額)
                        {
                            ViewBag.Message1 = "儲存失敗！尚有可分配預算";
                        }
                        if (主管分配金額加總 == model.上階主管分配金額)
                        {
                            DepartmentBudgetItem item_DepartmentBudget = null;
                            var dList_親核 = _DepartmentBudget.Select(BonusProjectID: model.BonusProjectID, EMP_SEQ_NO: 登入者, ISApproval: true);
                            for (int i = 0; i < dList_親核.Count(); i++)
                            {
                                if (i == 0)
                                {
                                    //畫面上親核的DepartmentBudget 某一列
                                    var d親核 = model.BonusDistributionData.Where(p => p.ISApproval == true).FirstOrDefault();
                                    dList_親核[i].FlexibleBudget = StringEncrypt.aesEncryptBase64(DAC.GetString(d親核.主管分配金額str.Replace(",", "")));
                                    _DepartmentBudget.UpdateOne(dList_親核[i]);
                                }
                                else
                                {
                                    dList_親核[i].FlexibleBudget = StringEncrypt.aesEncryptBase64("0");
                                    _DepartmentBudget.UpdateOne(dList_親核[i]);
                                }
                            }
                            foreach (var BonusDistributionData in model.BonusDistributionData)
                            {
                                if (BonusDistributionData.ISApproval == true)
                                    continue;
                                item_DepartmentBudget = _DepartmentBudget.SelectOne(BonusDistributionData.DepartmentBudgetID).FirstOrDefault();
                                item_DepartmentBudget.FlexibleBudget = StringEncrypt.aesEncryptBase64(DAC.GetString(BonusDistributionData.主管分配金額str.Replace(",", "")));
                                _DepartmentBudget.UpdateOne(item_DepartmentBudget);
                            }

                            ViewBag.Message1 = "儲存成功！";
                        }
                    }
                    break;
            }
            var Data = Get單位(DAC.GetInt32(model.BonusProjectID), false);
            model = new BonusApprovalListViewModel()
            {
                BonusProjectID = DAC.GetInt32(model.BonusProjectID),
                BonusDistributionData = Data,
                ReserveBudget = Project.ReserveBudget(DAC.GetInt32(model.BonusProjectID)) ?? 0,
                AllotType = Project.GetAllotType(DAC.GetInt32(model.BonusProjectID)) ?? 0,
                budget = _DepartmentBudget.保留百分比(DAC.GetInt32(model.BonusProjectID), 登入者),
                btn登入者未進行分派 = _DepartmentBudget.IS登入者未進行分派(model.BonusProjectID, 登入者),
                btn保留百分比 = _DepartmentBudget.IS編輯保留百分比_new(DAC.GetInt32(model.BonusProjectID), 登入者, 登入者DEPT_SEQ_NO),
                btn上層主管已分派完畢 = Modify.IsModify,
                str部門已分派 = str部門已分派,
                上階主管分配金額 = Data.Sum(p => p.主管分配金額)
            };
            return View(model);
        }

        #endregion



        #region BonusDistributionForChairman 核定結果審核 bak
        [CheckLoginSessionExpired]
        public ActionResult BonusDistributionForChairman_old(int BonusProjectID)
        {
            var Project = new DAC_BonusProject();
            var dac_file = new DAC_FILE_Association();
            ViewBag.BonusProjectID = BonusProjectID;
            ViewBag.ProjectName = Project.獎金專案名稱(BonusProjectID);
            var Data = Get單位(DAC.GetInt32(BonusProjectID), 董事長簽核: true);
            var model = new BonusApprovalListViewModel()
            {
                BonusProjectID = BonusProjectID,
                BonusDistributionData = Data,
                FileData = dac_file.Select(Table: "Bonus", LinkID: BonusProjectID),
                ReserveBudget = Project.ReserveBudget(BonusProjectID),
                budget = "1",
            };

            SysLog.Write(LoginUserID, 核定結果審核Page, SysLog.IntoPageLog(intoPageID: BonusProjectID, pageMode: SysLog.進入頁面.頁面, intoSuccess: true));
            return View(model);
        }
        [CheckLoginSessionExpired]
        [HttpPost]
        public ActionResult BonusDistributionForChairman_old(BonusApprovalListViewModel model)
        {
            var Project = new DAC_BonusProject();
            var dac_file = new DAC_FILE_Association();
            ViewBag.BonusProjectID = model.BonusProjectID;
            ViewBag.ProjectName = Project.獎金專案名稱(DAC.GetInt32(model.BonusProjectID));
            var Data = Get單位(DAC.GetInt32(model.BonusProjectID), 董事長簽核: true);
            model = new BonusApprovalListViewModel()
            {
                BonusDistributionData = Data,
                FileData = dac_file.Select(Table: "Bonus", LinkID: model.BonusProjectID),
                ReserveBudget = Project.ReserveBudget(DAC.GetInt32(model.BonusProjectID)),
                budget = "1",
            };
            return View(model);
        }
        #endregion

        #region Get單位
        /// <summary>
        /// Get單位 (登入者層級這邊是否需要改待確認)
        /// </summary>
        /// <param name="BonusProjectID"></param>
        /// <param name="董事長簽核"></param>
        /// <returns></returns>
        public List<BonusDistributionItem> Get單位(int BonusProjectID, bool 董事長簽核 = false)
        {
            DAC_BonusProject _BonusProject = new DAC_BonusProject();
            DAC_FL_PERSONNEL_TW_V _FL_PERSONNEL_TW_V = new DAC_FL_PERSONNEL_TW_V();
            DAC_FL_DEPARTMENT_TW_V _DEPARTMENT_TW_V = new DAC_FL_DEPARTMENT_TW_V();
            DAC_BonusApprovalList _BonusApprovalList = new DAC_BonusApprovalList();
            var List = new List<BonusDistributionItem>();
            string 登入者 = "";
            string 登入者部門 = "";
            string 登入者管理部門Lst = 轄下部門();
            int 登入者層級 = 0;
            var Project = _BonusProject.SelectOne(BonusProjectID).FirstOrDefault();
            int 專案最低分配層級 = Project?.ExecutionLevel ?? 0;

            if (董事長簽核 == false)
            {
                登入者 = (string)Session[PublicVariable.EMP_SEQ_NO];
                登入者部門 = (string)Session[PublicVariable.department_SEQ_NO]; //DEPT_SEQ_NO
                登入者層級 = DAC.GetInt32(_DEPARTMENT_TW_V.GETLEVEL_CODE_主管層級(登入者));
            }
            else
            {
                var 董事長 = _DEPARTMENT_TW_V.Select(LEVEL_CODE: "1").FirstOrDefault();
                登入者 = 董事長.MAIN_LEADER_NO;
                登入者部門 = 董事長.DEPT_SEQ_NO; //DEPT_SEQ_NO
                登入者層級 = DAC.GetInt32(_DEPARTMENT_TW_V.GetLEVEL_CODE(董事長.DEPT_SEQ_NO));
            }
            ViewBag.IS專案最低層級 = false;
            if (登入者層級 >= 專案最低分配層級)
            {
                ViewBag.IS專案最低層級 = true;
            }
            if (登入者 != "")
            {
                DAC_DepartmentBudget _DepartmentBudget = new DAC_DepartmentBudget();
                var dList_親核 = _DepartmentBudget.Select(BonusProjectID: BonusProjectID, EMP_SEQ_NO: 登入者, ISApproval: true);
                var dList = _DepartmentBudget.Select(BonusProjectID: BonusProjectID, EMP_SEQ_NO: 登入者, ISApproval: false);
                int 總預算 = 0, 轄下保留預算合計 = 0, 轄下固定預算 = 0, 轄下加碼預算 = 0;
                #region 轄下
                foreach (var d in dList)
                {
                    var item = new BonusDistributionItem();
                    item.DepartmentBudgetID = d.DepartmentBudgetID;
                    item.單位 = d.DEPT_NAME;
                    item.主管 = d.MAIN_LEADER_NAME;
                    item.主管_EMP_SEQ_NO = d.MAIN_LEADER_NO;
                    item.人數 = _BonusApprovalList.部門人數(BonusProjectID, d.DEPT_SEQ_NO, true, d.MAIN_LEADER_NO);
                    //人數為0時，幫她自動不分派
                    //if(item.人數 == 0 && d.NoAssign != true)
                    //{
                    //    d.NoAssign = true;
                    //    _DepartmentBudget.UpdateOne(d);
                    //}
                    item.個人預算 = d.固定預算;
                    item.加碼預算 = d.加碼預算;
                    item.DEPT_SEQ_NO = d.DEPT_SEQ_NO;
                    item.保留金額 = DAC.GetDecimal(StringEncrypt.aesDecryptBase64(d.ReserveBudget));
                    item.分給轄下單位總預算 = DAC.GetDecimal(StringEncrypt.aesDecryptBase64(d.Amount))
                        + DAC.GetDecimal(StringEncrypt.aesDecryptBase64(d.FlexibleBudget)); //分紅獎金須加上 主管分配金額(W);
                    item.主管分配金額 = DAC.GetDecimal(StringEncrypt.aesDecryptBase64(d.FlexibleBudget));
                    item.轄下主管核定結果 = DAC.GetDecimal(StringEncrypt.aesDecryptBase64(d.Result));
                    item.IS轄下主管核定結果 = d.ISAssign == true ? true : false;
                    item.差額 = DAC.GetDecimal(StringEncrypt.aesDecryptBase64(d.Difference)); //分紅獎金在送出時寫入已加上 主管分配金額(W);
                    item.IS差額 = d.ISAssign == true ? true : false;
                    item.ISapplicant = false;
                    var a = DAC.GetInt32(_DEPARTMENT_TW_V.GetLEVEL_CODE(d.DEPT_SEQ_NO));
                    item.IS專案最低層級 = false;
                    if (登入者層級 >= 專案最低分配層級)
                    {
                        item.IS專案最低層級 = true;
                    }
                    item.ISApproval = DAC.GetBoolean(d.ISApproval);
                    item.ISAssign = DAC.GetBoolean(d.ISAssign);
                    item.ISDepartmentBudget_ISAssign = DAC.GetBoolean(d.ISAssign);
                    //ISAssign==true==>已分配 || ISAssign==false==>未分配 
                    if (item.ISAssign == false)
                    {
                        var A = _DEPARTMENT_TW_V.GetLEVEL_CODE(d.DEPT_SEQ_NO);
                        if (DAC.GetInt32(_DEPARTMENT_TW_V.GetLEVEL_CODE(d.DEPT_SEQ_NO)) > 專案最低分配層級)
                        {
                            item.ISAssign = true;
                        }
                        //如遇最低層級則不往下分配
                        else if (登入者層級 >= 專案最低分配層級)
                        {
                            item.ISAssign = true;
                        }
                        #region 主管是否已送簽 送簽後則不可再分派
                        if (d.BPM_Status == 1 || d.BPM_Status == 3)//BPM單據狀態(1=簽核中、2=否決、3=簽核完成)
                        {
                            item.ISAssign = true;
                        }
                        #endregion
                    }
                    item.NoAssign = DAC.GetBoolean(d.NoAssign);
                    #region 簽核判斷
                    if (登入者層級 >= 專案最低分配層級)
                    {
                        item.Reject = false;
                        item.ISapplicant = true;
                    }
                    else
                    {
                        if (d.BPM_FormNO == null || d.BPM_FormNO == "")
                        {
                            item.Reject = false;
                        }
                        else
                        {
                            #region 找底下的部門是否都簽核完畢
                            //結案?
                            var 簽核List = BPMWebController.GetSignOff_Bonus(DAC.GetString(Session[PublicVariable.EMP_NO]));
                            bool 簽核 = 簽核List.Contains("'" + d.BPM_FormNO + "'");
                            item.Reject = 簽核;
                            if (ws.GetApproveList(Bonus_FormKind, d.BPM_FormNO)?.Data.Where(x => x.AppStatus == "A").FirstOrDefault()?.AppEmpNo == DAC.GetString(Session[PublicVariable.EMP_NO]))
                            {
                                item.Reject = false;
                                item.ISapplicant = true;
                            }
                            #endregion
                        }
                    }
                    #endregion

                    item.BPM_FormNO = d.BPM_FormNO;
                    item.FormKind = Bonus_FormKind;
                    List.Add(item);

                    總預算 += (DAC.GetInt32(item.個人預算) + DAC.GetInt32(item.加碼預算));
                    轄下保留預算合計 += DAC.GetInt32(item.保留金額);
                    轄下固定預算 += DAC.GetInt32(item.個人預算);
                    轄下加碼預算 += DAC.GetInt32(item.加碼預算);
                }
                #endregion
                #region 親核
                if (dList_親核.Count != 0)
                {
                    var item = new BonusDistributionItem();
                    item.單位 = "親核";
                    item.主管 = (string)Session[PublicVariable.UserName];
                    item.主管_EMP_SEQ_NO = (string)Session[PublicVariable.EMP_SEQ_NO];
                    int 人數 = 0;
                    decimal 個人預算 = 0, 分給轄下單位總預算 = 0, 轄下主管核定結果 = 0, 差額 = 0, 保留金額 = 0, 加碼預算 = 0, 主管分配金額 = 0;
                    bool BPM_FormNO = true;
                    foreach (var d in dList_親核)
                    {
                        //親核找出任一個即可
                        item.DepartmentBudgetID = d.DepartmentBudgetID;
                        人數 += _BonusApprovalList.部門人數(BonusProjectID, d.DEPT_SEQ_NO, false);
                        個人預算 += d.固定預算;
                        加碼預算 += d.加碼預算;
                        保留金額 += DAC.GetDecimal(StringEncrypt.aesDecryptBase64(d.ReserveBudget));
                        分給轄下單位總預算 += DAC.GetDecimal(StringEncrypt.aesDecryptBase64(d.Amount)) 
                            + DAC.GetDecimal(StringEncrypt.aesDecryptBase64(d.FlexibleBudget)); //分紅獎金須加上 主管分配金額(W)
                        轄下主管核定結果 += DAC.GetDecimal(StringEncrypt.aesDecryptBase64(d.Result));
                        差額 += DAC.GetDecimal(StringEncrypt.aesDecryptBase64(d.Difference)); //分紅獎金在送出時已在差額加上 主管分配金額(W);
                        主管分配金額 += DAC.GetDecimal(StringEncrypt.aesDecryptBase64(d.FlexibleBudget)); //主管分配金額(W)
                        if (d.BPM_FormNO == null || d.BPM_FormNO == "")
                            BPM_FormNO = false;
                    }
                    總預算 += (DAC.GetInt32(個人預算) + DAC.GetInt32(加碼預算));
                    item.人數 = 人數;
                    item.個人預算 = 個人預算;
                    item.加碼預算 = 加碼預算;
                    item.DEPT_SEQ_NO = "";
                    item.保留金額 = 保留金額;
                    item.主管分配金額 = 主管分配金額;
                    item.分給轄下單位總預算 = 分給轄下單位總預算;
                    item.轄下主管核定結果 = 轄下主管核定結果;
                    item.差額 = 差額;
                    item.ISApproval = true;
                    item.ISAssign = true;
                    item.Reject = false;
                    item.NoAssign = false;
                    if (登入者層級 >= 專案最低分配層級)
                    {
                        item.IS專案最低層級 = true;
                    }
                    else
                    {
                        item.IS專案最低層級 = false;
                    }
                    //List.Add(item);
                    List.Insert(0, item);
                }
                #endregion
            }
            return List;
        }
        #endregion

        public string 轄下部門()
        {
            string str轄下部門 = "";

            DAC_FL_PERSONNEL_TW_V _FL_PERSONNEL_TW_V = new DAC_FL_PERSONNEL_TW_V();
            DAC_FL_DEPARTMENT_TW_V _DEPARTMENT_TW_V = new DAC_FL_DEPARTMENT_TW_V();

            var 登入者 = (string)Session[PublicVariable.EMP_SEQ_NO];
            var 登入者部門 = (string)Session[PublicVariable.department_SEQ_NO]; //DEPT_SEQ_NO

            if (登入者 != "")
            {
                var Maindept = _DEPARTMENT_TW_V.Get登入者管理部門(登入者);
                List<string> Mlist = new List<string>();
                Mlist = Maindept.Split(',').ToList();
                foreach (var m in Mlist)
                {
                    登入者部門 = m;
                    DAC_BonusPersonnelList _BonusPersonnelList = new DAC_BonusPersonnelList();
                    var deptList = _DEPARTMENT_TW_V.GetChild_DEPT_SEQ_NO_List_不含自己部門(登入者部門, DAC_FL_DEPARTMENT_TW_V.資訊類型.代號, false);
                    if (deptList != "")
                    {
                        List<string> dlist = new List<string>();
                        dlist = deptList.Split(',').ToList();
                        foreach (var d in dlist)
                        {
                            str轄下部門 += d + ",";
                        }
                    }
                }
            }

            if (str轄下部門 != "")
            {
                str轄下部門 = str轄下部門.TrimEnd(',');
            }

            return str轄下部門;
        }

        /// <summary>
        /// 準備 View 要用的下拉清單
        /// </summary>
        protected override void PrepareSelectList()
        {
            base.PrepareSelectList();
        }

        #region New物件
        /// <summary>
        /// 建立新的 DAO 物件
        /// </summary>
        private DAC_BonusPersonnelList NewDAO(DbConnection conn)
        {
            return new DAC_BonusPersonnelList(conn);
        }

        private DAC_BonusProject Project_NewDAO(DbConnection conn)
        {
            return new DAC_BonusProject(conn);
        }
        #endregion

        #region 匯出評核名單
        [CheckLoginSessionExpired]
        public ActionResult Export(int BonusProjectID)
        {
            using (var conn = NewConnection())
            {
                conn.Open();
                var fileName = new DAC_BonusProject().SelectOne(BonusProjectID).FirstOrDefault()?.BonusProjectName ?? "";
                var excelExport = new BonusPersonnelList_Export();
                excelExport.BonusProjectID = BonusProjectID;
                fileName += "評核名單.xls";
                excelExport.Connection = conn;
                var output = excelExport.Export();
                byte[] bytes = output.ToArray();
                SysLog.Write(LoginUserID, 評核名單查詢與啟動Page, SysLog.FunctionLog(todoID: BonusProjectID, todoSomething: SysLog.執行操作.匯出檔案 + "(評核名單.xls)", doSuccess: true));
                return File(bytes, "application/vnd.ms-excel", fileName);
            }
        }
        #endregion

        #region 重新拋轉
        [HttpPost]
        public JsonResult Again(string id)
        {
            bool 結轉人員 = false, 獎金計算 = false;
            DAC_BonusProject project = new DAC_BonusProject();
            DAC_BonusPersonnelList _BonusPersonnelList = new DAC_BonusPersonnelList();
            DAC_PersonnelParameters _PersonnelParameters = new DAC_PersonnelParameters();
            var item = _BonusPersonnelList.SelectOne(DAC.GetInt32(id)).FirstOrDefault();
            if (item != null)
            {
                var Projectitem = project.SelectOne(DAC.GetInt32(item.BonusProjectID)).FirstOrDefault();
                if (Projectitem != null)
                {
                    結轉人員 = _PersonnelParameters.Bonus結轉人員名單(Projectitem, item.EMP_SEQ_NO);
                    獎金計算 = _BonusPersonnelList.BonusCalculationDLL(Projectitem, item.EMP_SEQ_NO);
                }
            }

            if (結轉人員 && 獎金計算)
            {
                SysLog.Write(LoginUserID, 評核名單查詢與啟動Page, SysLog.FunctionLog(todoID: id, todoSomething: "重新拋轉", doSuccess: true));
                return Json(new { success = true, responseText = "重新拋轉成功" }, JsonRequestBehavior.AllowGet);
            }
            else
            {
                SysLog.Write(LoginUserID, 評核名單查詢與啟動Page, SysLog.FunctionLog(todoID: id, todoSomething: "重新拋轉", doSuccess: false));
                return Json(new { success = false, responseText = "重新拋轉失敗" }, JsonRequestBehavior.AllowGet);
            }
        }
        #endregion
        public ActionResult FileDownload(int FileID)
        {
            var dao = new DAC_FILE_UPLOAD();
            var content = dao.GetFileContent(FileID);
            var info = dao.SelectOne(FileID).Any()
                       ?dao.SelectOne(FileID).FirstOrDefault()
                       : new FILE_UPLOADItem();
            string mimeType = MimeMapping.GetMimeMapping(info.ORG_FILE_NAME);

            SysLog.Write(LoginUserID, 評核名單查詢與啟動Page, SysLog.FunctionLog(todoID: FileID, todoSomething: SysLog.執行操作.下載檔案 + $"({info.ORG_FILE_NAME})", doSuccess: true));
            return File(content, mimeType, info.ORG_FILE_NAME);
        }

        #region 分配轄下 & 不分配轄下
        //分配轄下
        public JsonResult AssignAll(List<BudgetCLASS> dataPost)
        {
            //todo 寫入DepartmentBudget
            try
            {
                if (dataPost == null)
                {
                    return Json(new { success = true, responseText = "無轄下分配" }, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    //DAC_DepartmentBudget _DepartmentBudget = new DAC_DepartmentBudget();
                    //foreach (var d in dataPost)
                    //{
                    //    var deptItem = _DepartmentBudget.Select(BonusProjectID: d.BonusProjectID, DEPT_SEQ_NO: d.DEPT_SEQ_NO, EMP_SEQ_NO: (string)Session[PublicVariable.EMP_SEQ_NO]).FirstOrDefault();
                    //    //已執行分派作業的不可再分派
                    //    if (deptItem.ISAssign == true || deptItem.ISApproval == true || deptItem.NoAssign == true)
                    //        dataPost.Remove(d);
                    //}
                    new DAC_DepartmentBudget().WriteInDepartmentBudget_分配轄下(dataPost, (string)Session[PublicVariable.EMP_SEQ_NO]);
                    new DAC_DepartmentBudget().WriteInDepartmentBudget_分配轄下_人員(dataPost, (string)Session[PublicVariable.EMP_SEQ_NO]);

                    return Json(new { success = true, responseText = "分配轄下成功" }, JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception ex)
            {
                return Json(new { success = false, responseText = "分配轄下失敗，失敗原因:" + ex.Message }, JsonRequestBehavior.AllowGet);
            }

        }

        //不分配轄下
        public JsonResult NoAssignAll(List<BudgetCLASS> dataPost)
        {
            try
            {
                if (dataPost == null)
                {
                    return Json(new { success = true, responseText = "無不分派單位" }, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    DAC_DepartmentBudget _DepartmentBudget = new DAC_DepartmentBudget();

                    // -- 若不分派，則回復部門預算到親核部門 START--
                    var budgetTmp = 0;
                    var projInfo = dataPost.FirstOrDefault() != null
                                   ? new DAC_BonusProject().SelectOne(dataPost.FirstOrDefault().BonusProjectID).FirstOrDefault()
                                   : null;
                    var 親核Item = new DepartmentBudgetItem();
                    var myFlexibleBudget = "";

                    if (dataPost.FirstOrDefault() != null)
                    {
                        if (projInfo != null) {
                            //僅調整有用到"主管可分配金額"的類型: 彈性預算=2 & 複合預算=3
                            if (projInfo.BudgetType ==2 || projInfo.BudgetType == 3) 
                            {      
                                親核Item = _DepartmentBudget.Select(BonusProjectID: projInfo.BonusProjectID, EMP_SEQ_NO: (string)Session[PublicVariable.EMP_SEQ_NO], ISApproval: true)
                                                            .Where(x => x.FlexibleBudget != StringEncrypt.aesEncryptBase64("0"))
                                                            .FirstOrDefault();
                                myFlexibleBudget = StringEncrypt.aesDecryptBase64(親核Item.FlexibleBudget);

                                budgetTmp = _DepartmentBudget
                                            .Select(BonusProjectID: projInfo.BonusProjectID, EMP_SEQ_NO: (string)Session[PublicVariable.EMP_SEQ_NO])
                                            .Where(x => dataPost.Select(y => y.DEPT_SEQ_NO).Contains(x.DEPT_SEQ_NO)
                                                     && x.主管可分配金額 != 0)
                                            .GroupBy(x => new { x.EMP_SEQ_NO })
                                            .Select(x => new { x.Key.EMP_SEQ_NO, BUDGET_TMP = x.Sum(y => y.主管可分配金額) })
                                            .FirstOrDefault() != null ? _DepartmentBudget
                                            .Select(BonusProjectID: projInfo.BonusProjectID, EMP_SEQ_NO: (string)Session[PublicVariable.EMP_SEQ_NO])
                                            .Where(x => dataPost.Select(y => y.DEPT_SEQ_NO).Contains(x.DEPT_SEQ_NO)
                                                     && x.主管可分配金額 != 0)
                                            .GroupBy(x => new { x.EMP_SEQ_NO })
                                            .Select(x => new { x.Key.EMP_SEQ_NO, BUDGET_TMP = x.Sum(y => y.主管可分配金額) })
                                            .FirstOrDefault()
                                            .BUDGET_TMP : 0;

                                if (budgetTmp!=0)
                                {
                                    親核Item.FlexibleBudget = StringEncrypt.aesEncryptBase64((int.Parse(myFlexibleBudget) + budgetTmp).ToString());
                                    _DepartmentBudget.UpdateOne(親核Item);                            
                                }
                            }                        
                        }    
                    }
                    else {
                        return Json(new { success = false, responseText = "不分派失敗，失敗原因: 取得獎金專案資訊不足"  }, JsonRequestBehavior.AllowGet);
                    }                  
                    // -- 若不分派，則回復部門預算到親核部門 END--

                    //foreach (var d in dataPost)
                    //{
                    //   var deptItem = _DepartmentBudget.Select(BonusProjectID: d.BonusProjectID, DEPT_SEQ_NO: d.DEPT_SEQ_NO, EMP_SEQ_NO: (string)Session[PublicVariable.EMP_SEQ_NO]).Any()
                    //        ?_DepartmentBudget.Select(BonusProjectID: d.BonusProjectID, DEPT_SEQ_NO: d.DEPT_SEQ_NO, EMP_SEQ_NO: (string)Session[PublicVariable.EMP_SEQ_NO]).FirstOrDefault()
                    //        : new DepartmentBudgetItem();
                    //    if(deptItem.主管可分配金額 != 0)
                    //        return Json(new { success = true, responseText = "若不分派【主管分配金額W】欄位必須為0，請重新確認" }, JsonRequestBehavior.AllowGet);
                    //}

                    foreach (var d in dataPost)
                    {
                        var deptItem = _DepartmentBudget.Select(BonusProjectID: d.BonusProjectID, DEPT_SEQ_NO: d.DEPT_SEQ_NO, EMP_SEQ_NO: (string)Session[PublicVariable.EMP_SEQ_NO]).FirstOrDefault();
                        //已執行分派作業的不可再分派
                        //if (deptItem.ISAssign == true || deptItem.ISApproval == true || deptItem.NoAssign == true)
                        //    continue;
                        if (deptItem != null)
                        {
                            if (projInfo != null)
                            {
                                if ((projInfo.BudgetType == 2 || projInfo.BudgetType == 3) && deptItem.主管可分配金額 != 0)
                                {
                                    deptItem.FlexibleBudget = StringEncrypt.aesEncryptBase64("0");
                                }
                            }

                            deptItem.NoAssign = true;
                            _DepartmentBudget.UpdateOne(deptItem);
                        }
                    }
                    return Json(new { success = true, responseText = "不分派成功" }, JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception ex)
            {
                return Json(new { success = false, responseText = "不分派失敗，失敗原因:" + ex.Message }, JsonRequestBehavior.AllowGet);
            }

        }
        #endregion
        /// <summary>
        /// 暫存該頁面的資料(只限該頁面顯示的比數)
        /// </summary>
        /// <param name="dataPost"></param>
        /// <returns></returns>
        public JsonResult SaveAll(List<PersonnelCLASS> dataPost)
        {
            string 登入者 = (string)Session[PublicVariable.EMP_SEQ_NO];
            string 登入者層級 = new DAC_FL_DEPARTMENT_TW_V().GETLEVEL_CODE_主管層級(登入者);

            Message message = new DAC_BonusApprovalList().WriteInApprovalList_new(dataPost, 登入者);
            return Json(new
            {
                success = message.ISsuccess,
                responseText = message.ErrorMessage == "" ? "更新成功" : "成功筆數:" + message.SuccessCount + "筆\n失敗筆數:" + message.ErrorCount + "筆;" + message.ErrorMessage
            }, JsonRequestBehavior.AllowGet);
        }
        /// <summary>
        /// 暫存該頁面的資料(只限該頁面顯示的比數)
        /// </summary>
        /// <param name="dataPost"></param>
        /// <returns></returns>
        public JsonResult SaveAll_0713(List<PersonnelCLASS> dataPost)
        {
            string 登入者 = (string)Session[PublicVariable.EMP_SEQ_NO];
            string 登入者層級 = new DAC_FL_DEPARTMENT_TW_V().GETLEVEL_CODE_主管層級(登入者);

            Message message = new DAC_BonusApprovalList().WriteInApprovalList(dataPost, 登入者, 登入者層級);
            return Json(new
            {
                success = message.ISsuccess,
                responseText = message.ErrorMessage == "" ? "更新成功" : "成功筆數:" + message.SuccessCount + "筆\n失敗筆數:" + message.ErrorCount + "筆;" + message.ErrorMessage
            }, JsonRequestBehavior.AllowGet);
        }
        public JsonResult deletefile(int Bonusid, int id)
        {
            var dac_file = new DAC_FILE_Association();
            var item = dac_file.Select(Table: "Bonus", LinkID: Bonusid, FILE_CONTENTID: id).Any()
                       ? dac_file.Select(Table: "Bonus", LinkID: Bonusid, FILE_CONTENTID: id).FirstOrDefault()
                       : new FILE_AssociationItem();
            if (dac_file.DeleteOne(item))
            {
                return Json(new { success = true, responseText = item.clientMessage }, JsonRequestBehavior.AllowGet);
            }
            else
            {
                return Json(new { success = false, responseText = item.clientMessage }, JsonRequestBehavior.AllowGet);
            }
        }

        #region SignRemarkPopup
        [CheckLoginSessionExpired]
        public ActionResult SignRemarkPopup(bool agree, int[] AppIDs, string mode, int BonusProjectID, bool CloseOpener = false)
        {
            DoSomething_SyslogEnum = SysLog.執行操作.填寫簽核意見;
            SysLog.Write(LoginUserID, 簽核意見編輯Page, SysLog.FunctionLog(AppIDs, DoSomething_SyslogEnum, null));

            List<int> IDlist = new List<int>();
            foreach (int i in AppIDs)
            {
                var dept = new DAC_DepartmentBudget().Select(BonusProjectID: BonusProjectID
                    , EMP_SEQ_NO: (string)Session[PublicVariable.EMP_SEQ_NO]
                    , DEPT_SEQ_NO: DAC.GetString(i)).FirstOrDefault();
                if (dept != null)
                {
                    IDlist.Add(dept.DepartmentBudgetID);
                }
            }
            int[] arr_id = IDlist.ToArray();
            var model = new SignRemarkViewModel() { Agree = agree, AppIDs = arr_id, Mode = mode, CloseOpener = CloseOpener };
            ViewBag.CloseOpener = false;
            ViewBag.OK = false;
            return View(model);
        }

        [HttpPost]
        [CheckLoginSessionExpired]
        public ActionResult SignRemarkPopup_bak(SignRemarkViewModel model)
        {
            DoSomething_SyslogEnum = SysLog.執行操作.執行簽核;
            string errorMessage = "";
            string UpdataErrorMessage = "";
            string FinalMsg = "";
            var BPMWebController = new BPMWebServiceController();
            var list = new List<BPMWebServiceController.SignOffData>();
            var dao = new DAC_DepartmentBudget();
            var dao_BonusApprovalList = new DAC_BonusApprovalList();
            DAC_FL_DEPARTMENT_TW_V dAC_FL_DEPARTMENT_TW_V = new DAC_FL_DEPARTMENT_TW_V();
            var EMP_NO = DAC.GetString(Session[PublicVariable.EMP_NO]);
            var ws = new BPMWebService();
            var 登入者 = (string)Session[PublicVariable.EMP_SEQ_NO];

            if (model.CloseOpener)
                ViewBag.CloseOpener = true;
            else
                ViewBag.CloseOpener = false;

            string LogPath = PublicVariable.FileUploadPath + "Log/";

            foreach (var AppID in model.AppIDs)
            {
                var Ditem = dao.SelectOne(AppID).FirstOrDefault();
                var BPM_FormNO = Ditem?.BPM_FormNO;
                if (BPM_FormNO == null || BPM_FormNO == "")
                {
                    ViewBag.Message = "查無BPM單號";
                    return View(model);
                }
                var remark = model.Agree ? model.AgreeRemark : model.RejectRemark;
                var item = new BPMWebServiceController.SignOffData() { Form_NO = BPM_FormNO, formKind = Bonus_FormKind, remark = remark };
                list.Add(item);//

                var 轄下List = dao.查詢轄下BPM_NO(登入者, DAC.GetInt32(Ditem.BonusProjectID), new DAC_FL_DEPARTMENT_TW_V().GetChild_DEPT_SEQ_NO_List(Ditem.DEPT_SEQ_NO, DAC_FL_DEPARTMENT_TW_V.資訊類型.代號));

                foreach (var i in 轄下List)
                {
                    item = new BPMWebServiceController.SignOffData() { Form_NO = i.BPM_FormNO, formKind = Bonus_FormKind, remark = remark };
                    list.Add(item);

                    //更新為否決
                    i.BPM_Status = 2;
                    dao.UpdateOne(i);
                }


                if (model.Agree)
                {
                    BPMWebController.AgreeAll(list, EMP_NO, out errorMessage);
                    if (string.IsNullOrEmpty(errorMessage) == false)
                    {
                        FinalMsg += errorMessage;
                    }
                    if (string.IsNullOrEmpty(UpdataErrorMessage) == false)
                    {
                        FinalMsg += UpdataErrorMessage;
                    }
                    ViewBag.Message = "同意成功";
                }
                if (model.Agree == false)
                {
                    list = list.GroupBy(B => B.Form_NO)
                      .Select(l => l.First())
                      .ToList();
                    //否決
                    BPMWebController.ReturnFormAll(list, (string)Session[PublicVariable.EMP_NO], out errorMessage);
                    if (string.IsNullOrEmpty(errorMessage) == false)
                    {
                        FinalMsg += errorMessage;
                    }
                    if (string.IsNullOrEmpty(UpdataErrorMessage) == false)
                    {
                        FinalMsg += UpdataErrorMessage;
                    }

                    #region 駁回人員資料&部門
                    Ditem = dao.SelectOne(AppID).FirstOrDefault();
                    if (Ditem != null)
                    {
                        var levelcode = new DAC_FL_DEPARTMENT_TW_V().GetLEVEL_CODE(Ditem.DEPT_SEQ_NO);
                        var PList = dao_BonusApprovalList.Select_轄下人員(BonusProjectID: Ditem.BonusProjectID, EMP_SEQ_NO: Ditem.EMP_SEQ_NO, DEPT_SEQ_NO: dAC_FL_DEPARTMENT_TW_V.GetChild_DEPT_SEQ_NO_List(Ditem.DEPT_SEQ_NO, DAC_FL_DEPARTMENT_TW_V.資訊類型.代號));
                        foreach (var p in PList)
                        {
                            switch (levelcode)
                            {
                                case "1":
                                    p.ISApproval1 = false;
                                    p.ISApproval2 = false;
                                    p.ISApproval3 = false;
                                    p.ISApproval4 = false;
                                    p.ISApproval5 = false;
                                    p.ISApproval6 = false;
                                    p.ISApproval7 = false;
                                    p.ISApproval8 = false;
                                    p.ISApproval9 = false;
                                    p.ISApproval10 = false;
                                    break;
                                case "2":
                                    p.ISApproval2 = false;
                                    p.ISApproval3 = false;
                                    p.ISApproval4 = false;
                                    p.ISApproval5 = false;
                                    p.ISApproval6 = false;
                                    p.ISApproval7 = false;
                                    p.ISApproval8 = false;
                                    p.ISApproval9 = false;
                                    p.ISApproval10 = false;
                                    break;
                                case "3":
                                    p.ISApproval3 = false;
                                    p.ISApproval4 = false;
                                    p.ISApproval5 = false;
                                    p.ISApproval6 = false;
                                    p.ISApproval7 = false;
                                    p.ISApproval8 = false;
                                    p.ISApproval9 = false;
                                    p.ISApproval10 = false;
                                    break;
                                case "4":
                                    p.ISApproval4 = false;
                                    p.ISApproval5 = false;
                                    p.ISApproval6 = false;
                                    p.ISApproval7 = false;
                                    p.ISApproval8 = false;
                                    p.ISApproval9 = false;
                                    p.ISApproval10 = false;
                                    break;
                                case "5":
                                    p.ISApproval5 = false;
                                    p.ISApproval6 = false;
                                    p.ISApproval7 = false;
                                    p.ISApproval8 = false;
                                    p.ISApproval9 = false;
                                    p.ISApproval10 = false;
                                    break;
                                case "6":
                                    p.ISApproval6 = false;
                                    p.ISApproval7 = false;
                                    p.ISApproval8 = false;
                                    p.ISApproval9 = false;
                                    p.ISApproval10 = false;
                                    break;
                                case "7":
                                    p.ISApproval7 = false;
                                    p.ISApproval8 = false;
                                    p.ISApproval9 = false;
                                    p.ISApproval10 = false;
                                    break;
                                case "8":
                                    p.ISApproval8 = false;
                                    p.ISApproval9 = false;
                                    p.ISApproval10 = false;
                                    break;
                                case "9":
                                    p.ISApproval9 = false;
                                    p.ISApproval10 = false;
                                    break;
                                case "10":
                                    p.ISApproval10 = false;
                                    break;
                                default:
                                    break;
                            }
                            p.ReSignerID = DAC.GetInt32(Ditem.MAIN_LEADER_NO);
                            dao_BonusApprovalList.UpdateOne(p);
                        }
                    }
                    #endregion

                    var BonusProject = new DAC_BonusProject().SelectOne(DAC.GetInt32(Ditem.BonusProjectID)).FirstOrDefault();

                    #region 寄信
                    new DAC_EmailSendLog().BonusFixedBudget_Approval(BonusProject, Ditem, Ditem.DEPT_NAME, "駁回", remark);
                    #endregion

                    ViewBag.Message = "駁回成功";
                }

            }


            if (string.IsNullOrEmpty(FinalMsg) == false)
            {
                ViewBag.Message = FinalMsg;
            }
            ViewBag.OK = true;
            //SysLog.Write(LoginUserID, PageName, SysLog.FunctionLog(BPM_FormNO, DoSomething_SyslogEnum, null, ViewBag.Message));

            return View(model);
        }

        [HttpPost]
        [CheckLoginSessionExpired]
        public ActionResult SignRemarkPopup(SignRemarkViewModel model)
        {
            DoSomething_SyslogEnum = SysLog.執行操作.執行簽核;
            string errorMessage = "";
            string UpdataErrorMessage = "";
            string FinalMsg = "";
            var BPMWebController = new BPMWebServiceController();
            
            var dao = new DAC_DepartmentBudget();
            var dao_BonusApprovalList = new DAC_BonusApprovalList();
            DAC_FL_DEPARTMENT_TW_V dAC_FL_DEPARTMENT_TW_V = new DAC_FL_DEPARTMENT_TW_V();
            var EMP_NO = DAC.GetString(Session[PublicVariable.EMP_NO]);
            var ws = new BPMWebService();
            var 登入者 = (string)Session[PublicVariable.EMP_SEQ_NO];

            if (model.CloseOpener)
                ViewBag.CloseOpener = true;
            else
                ViewBag.CloseOpener = false;

            string LogPath = PublicVariable.FileUploadPath + "Log/";

            foreach (var AppID in model.AppIDs)
            {
                var list = new List<BPMWebServiceController.SignOffData>();
                var Ditem = dao.SelectOne(AppID).FirstOrDefault();
                SysLog.Write(LoginUserID, "獎金分配駁回", "AppID:" + AppID + ", Ditem.DEPT_SEQ_NO:" + Ditem?.DEPT_SEQ_NO);
                var BPM_FormNO = Ditem?.BPM_FormNO;
                if (string.IsNullOrWhiteSpace(BPM_FormNO))
                {
                    continue;
                }
                var remark = model.Agree ? model.AgreeRemark : model.RejectRemark;
                var item = new BPMWebServiceController.SignOffData() { Form_NO = BPM_FormNO, formKind = Bonus_FormKind, remark = remark };
                list.Add(item);

                //先將畫面指定的部門狀態退回
                Ditem.BPM_Status = 2;
                //將差額及主管核定獎金清除
                Ditem.Result = null;
                Ditem.Difference = null;
                dao.UpdateOne(Ditem);

                //再將轄下其他有被分派的部門(代表有啟單)都退回上一關
                var 轄下List = dao.Select_轄下已分派byDEPT_SEQ_NO(Ditem.BonusProjectID, Ditem.DEPT_SEQ_NO);

                foreach (var i in 轄下List)
                {
                    SysLog.Write(LoginUserID, "獎金分配駁回", "i.DEPT_SEQ_NO:" + i.DEPT_SEQ_NO);
                    item = new BPMWebServiceController.SignOffData() { Form_NO = i.BPM_FormNO, formKind = Bonus_FormKind, remark = remark };
                    list.Add(item);

                    //更新為否決
                    i.BPM_Status = 2;
                    dao.UpdateOne(i);
                }


                if (model.Agree == false)
                {
                    list = list.GroupBy(B => B.Form_NO)
                      .Select(l => l.First())
                      .ToList();
                    //否決
                    BPMWebController.ReturnFormAll(list, (string)Session[PublicVariable.EMP_NO], out errorMessage);
                    if (string.IsNullOrEmpty(errorMessage) == false)
                    {
                        FinalMsg += errorMessage;
                    }
                    if (string.IsNullOrEmpty(UpdataErrorMessage) == false)
                    {
                        FinalMsg += UpdataErrorMessage;
                    }

                    #region 駁回人員資料&部門
                    Ditem = dao.SelectOne(AppID).FirstOrDefault();
                    if (Ditem != null)
                    {
                        var levelcode = new DAC_FL_DEPARTMENT_TW_V().GetLEVEL_CODE(Ditem.DEPT_SEQ_NO);
                        var PList = dao_BonusApprovalList.Select_轄下人員(BonusProjectID: Ditem.BonusProjectID, EMP_SEQ_NO: Ditem.EMP_SEQ_NO, DEPT_SEQ_NO: dAC_FL_DEPARTMENT_TW_V.GetChild_DEPT_SEQ_NO_List(Ditem.DEPT_SEQ_NO, DAC_FL_DEPARTMENT_TW_V.資訊類型.代號));
                        SysLog.Write(LoginUserID, "獎金分配駁回-Pre", "Ditem.DEPT_SEQ_NO:" + Ditem.DEPT_SEQ_NO + ", EMP_SEQ_NO = " + Ditem.EMP_SEQ_NO + ", DEPT_SEQ_NO:" + dAC_FL_DEPARTMENT_TW_V.GetChild_DEPT_SEQ_NO_List(Ditem.DEPT_SEQ_NO, DAC_FL_DEPARTMENT_TW_V.資訊類型.代號));
                        foreach (var p in PList)
                        {
                            SysLog.Write(LoginUserID, "獎金分配駁回-In", "EMP_SEQ_NO:" + p.EMP_SEQ_NO);
                            switch (levelcode)
                            {
                                case "1":
                                    p.ISApproval1 = false;
                                    p.ISApproval2 = false;
                                    p.ISApproval3 = false;
                                    p.ISApproval4 = false;
                                    p.ISApproval5 = false;
                                    p.ISApproval6 = false;
                                    p.ISApproval7 = false;
                                    p.ISApproval8 = false;
                                    p.ISApproval9 = false;
                                    p.ISApproval10 = false;
                                    break;
                                case "2":
                                    p.ISApproval2 = false;
                                    p.ISApproval3 = false;
                                    p.ISApproval4 = false;
                                    p.ISApproval5 = false;
                                    p.ISApproval6 = false;
                                    p.ISApproval7 = false;
                                    p.ISApproval8 = false;
                                    p.ISApproval9 = false;
                                    p.ISApproval10 = false;
                                    break;
                                case "3":
                                    p.ISApproval3 = false;
                                    p.ISApproval4 = false;
                                    p.ISApproval5 = false;
                                    p.ISApproval6 = false;
                                    p.ISApproval7 = false;
                                    p.ISApproval8 = false;
                                    p.ISApproval9 = false;
                                    p.ISApproval10 = false;
                                    break;
                                case "4":
                                    p.ISApproval4 = false;
                                    p.ISApproval5 = false;
                                    p.ISApproval6 = false;
                                    p.ISApproval7 = false;
                                    p.ISApproval8 = false;
                                    p.ISApproval9 = false;
                                    p.ISApproval10 = false;
                                    break;
                                case "5":
                                    p.ISApproval5 = false;
                                    p.ISApproval6 = false;
                                    p.ISApproval7 = false;
                                    p.ISApproval8 = false;
                                    p.ISApproval9 = false;
                                    p.ISApproval10 = false;
                                    break;
                                case "6":
                                    p.ISApproval6 = false;
                                    p.ISApproval7 = false;
                                    p.ISApproval8 = false;
                                    p.ISApproval9 = false;
                                    p.ISApproval10 = false;
                                    break;
                                case "7":
                                    p.ISApproval7 = false;
                                    p.ISApproval8 = false;
                                    p.ISApproval9 = false;
                                    p.ISApproval10 = false;
                                    break;
                                case "8":
                                    p.ISApproval8 = false;
                                    p.ISApproval9 = false;
                                    p.ISApproval10 = false;
                                    break;
                                case "9":
                                    p.ISApproval9 = false;
                                    p.ISApproval10 = false;
                                    break;
                                case "10":
                                    p.ISApproval10 = false;
                                    break;
                                default:
                                    break;
                            }
                            p.ReSignerID = DAC.GetInt32(Ditem.MAIN_LEADER_NO);
                            dao_BonusApprovalList.UpdateOne(p);
                        }
                    }
                    #endregion

                    var BonusProject = new DAC_BonusProject().SelectOne(DAC.GetInt32(Ditem.BonusProjectID)).FirstOrDefault();

                    #region 寄信
                    new DAC_EmailSendLog().BonusFixedBudget_Approval(BonusProject, Ditem, Ditem.DEPT_NAME, "駁回", remark);
                    #endregion

                    ViewBag.Message = "駁回成功";
                }

            }


            if (string.IsNullOrEmpty(FinalMsg) == false)
            {
                ViewBag.Message = FinalMsg;
            }
            ViewBag.OK = true;
            //SysLog.Write(LoginUserID, PageName, SysLog.FunctionLog(BPM_FormNO, DoSomething_SyslogEnum, null, ViewBag.Message));

            return View(model);
        }

        /// <summary>
        /// 透過BPM Web Service取得 BPM_FormNO
        /// </summary>
        private string GetBPM_FormNO(DepartmentBudgetItem item)
        {
            CreateResult result啟單 = null;
            string BPM_FormNO = "";
            DAC_FL_PERSONNEL_TW_V daoFL_PERSON = new DAC_FL_PERSONNEL_TW_V();
            DAC_FL_DEPARTMENT_TW_V daoFL_DEPT = new DAC_FL_DEPARTMENT_TW_V();
            DAC_BonusProject _BonusProject = new DAC_BonusProject();
            var item申請人 = daoFL_PERSON.Select(false, EMP_NO: (string)Session[PublicVariable.EMP_NO]).FirstOrDefault();
            if (item申請人 == null)
                throw new Exception("找不到申請人");
            var item異動單_DEPT = daoFL_DEPT.Select(false, DEPT_SEQ_NO: (string)Session[PublicVariable.department_SEQ_NO]).FirstOrDefault();
            if (item異動單_DEPT == null)
                throw new Exception("找不到異動部門");
            var itemBonus = _BonusProject.SelectOne(DAC.GetInt32(item.BonusProjectID)).FirstOrDefault();
            if (itemBonus == null)
                throw new Exception("找不到獎金專案主檔");

            //1.啟單 呼叫BPMWebService(CreateNewForm)

            result啟單 = APPController.CreateBonus(new BonusModel()
            {
                EMP_NO = (string)Session[PublicVariable.EMP_NO],
                EMP_CNAME = item申請人.EMP_NAME,
                DEPT_CNAME = item.DEPT_NAME,
                DEPT_CODE = item.DEPT_NO,
                PROJECT_YEAR = DAC.GetString(itemBonus.BonusYear),
                PROJECT_NO = DAC.GetString(itemBonus.BonusProjectNo),
                PROJECT_CNAME = DAC.GetString(itemBonus.BonusProjectName),
                BUDGET_TYPE = DAC.GetString(SystemVariable.BonusType.GetName(itemBonus.AllotType))
            });

            if (result啟單.Success)
            {
                BPM_FormNO = result啟單.FormNo;
            }
            return BPM_FormNO;
        }

        /// <summary>
        /// 被駁回後重送
        /// </summary>
        private string ResendForm(DepartmentBudgetItem item)
        {
            CreateResult result啟單 = null;
            string BPM_FormNO = "";
            DAC_FL_PERSONNEL_TW_V daoFL_PERSON = new DAC_FL_PERSONNEL_TW_V();
            DAC_FL_DEPARTMENT_TW_V daoFL_DEPT = new DAC_FL_DEPARTMENT_TW_V();
            DAC_BonusProject _BonusProject = new DAC_BonusProject();

            var item申請人 = daoFL_PERSON.Select(false, EMP_NO: (string)Session[PublicVariable.EMP_NO]).FirstOrDefault();
            if (item申請人 == null)
                throw new Exception("找不到申請人");
            var item異動單_DEPT = daoFL_DEPT.Select(false, DEPT_SEQ_NO: (string)Session[PublicVariable.department_SEQ_NO]).FirstOrDefault();
            if (item異動單_DEPT == null)
                throw new Exception("找不到異動部門");
            var itemBonus = _BonusProject.SelectOne(DAC.GetInt32(item.BonusProjectID)).FirstOrDefault();
            if (itemBonus == null)
                throw new Exception("找不到獎金專案主檔");

            tpehrmap04.BPMWebService _ws = new tpehrmap04.BPMWebService();
            var APPLYER_ID = _ws.GetBpmEmpId((string)Session[PublicVariable.EMP_NO]);

            var BonusModel = new BonusModel()
            {
                EMP_NO = (string)Session[PublicVariable.EMP_NO],
                APPLYER_ID = APPLYER_ID,
                EMP_CNAME = item申請人.EMP_NAME,
                DEPT_CNAME = item.DEPT_NAME,
                DEPT_CODE = item.DEPT_NO,
                PROJECT_YEAR = DAC.GetString(itemBonus.BonusYear),
                PROJECT_NO = DAC.GetString(itemBonus.BonusProjectNo),
                PROJECT_CNAME = DAC.GetString(itemBonus.BonusProjectName),
                BUDGET_TYPE = DAC.GetString(SystemVariable.BonusType.GetName(itemBonus.AllotType))
            };


            var result = new tpehrmap04.BPMWebService().SendForm(
                FormKind: Bonus_FormKind,
                FormNo: item.BPM_FormNO,
                EmpNo: (string)Session[PublicVariable.EMP_NO],
                Fields: JsonConvert.SerializeObject(BonusModel, Formatting.Indented)
                );

            if (result.Success)
            {
                return result.Data;
            }

            return null;
        }
        /// <summary>
        /// 送出簽核部門的邏輯為 【轄下】且有【被指派(ISAssign = true)】的部門
        /// </summary>
        /// <param name="BonusProjectID"></param>
        /// <returns></returns>
        private ApprovalMessageCLASS ApprovalAll(int BonusProjectID)
        {
            //簽核邏輯待修正
            var ApprovalMessage = new ApprovalMessageCLASS();
            bool Approval = false;
            string error = "";
            string errorMessage = "";
            DAC_BonusProject _BonusProject = new DAC_BonusProject();
            DAC_FL_PERSONNEL_TW_V _FL_PERSONNEL_TW_V = new DAC_FL_PERSONNEL_TW_V();
            DAC_FL_DEPARTMENT_TW_V _DEPARTMENT_TW_V = new DAC_FL_DEPARTMENT_TW_V();
            DAC_BonusApprovalList _BonusApprovalList = new DAC_BonusApprovalList();
            DAC_DepartmentBudget _DepartmentBudget = new DAC_DepartmentBudget();

            var 登入者seqno = (string)Session[PublicVariable.EMP_SEQ_NO];
            var 登入者工號 = (string)Session[PublicVariable.EMP_NO];
            var 登入者部門 = (string)Session[PublicVariable.department_SEQ_NO]; //DEPT_SEQ_NO
            bool is總經理層級 = "2" == new DAC_FL_DEPARTMENT_TW_V().GETLEVEL_CODE_主管層級(登入者seqno);
            List<string> 簽核者被上階分派的部門s = _DepartmentBudget.GetDepartmentBudget_被上級分派的部門(BonusProjectID, 登入者seqno);

            bool 保留預算 = (_BonusProject.SelectOne(BonusProjectID).FirstOrDefault()?.ReserveBudget ?? 0) == 1 ? true : false;
            //找出登入者的需簽核的部門
            var dList_親核 = _DepartmentBudget.Select(BonusProjectID: BonusProjectID, EMP_SEQ_NO: 登入者seqno, ISApproval: true);
            var dList = _DepartmentBudget.Select(BonusProjectID: BonusProjectID, EMP_SEQ_NO: 登入者seqno, ISApproval: false);

            //1) 先檢查
            List<ApprovalCLASS> approvalCLASSList = new List<ApprovalCLASS>();
            //親核部門則是送簽
            foreach (var A in dList_親核)
            {
                var dacBonusPersonne = new DAC_BonusApprovalList();
                //var List人員 = dacBonusPersonne.Select_人員_byDEPT_SEQ_NO(
                //登入者層級: 登入者層級
                //, 登入者: 登入者seqno,
                //登入者DEPT_SEQ_NO: A.DEPT_SEQ_NO,
                //BonusProjectID: A.BonusProjectID, 含轄下: true);
                int 主管調整金額 = 0;// dacBonusPersonne.主管微調總金額(List人員);
                int 轄下人員金額 = 0;// dacBonusPersonne.轄下人員金額(List人員);
                int 親核固定金額 = 0;// dacBonusPersonne.親核固定金額(List人員);
                int 親核加碼金額 = 0;// dacBonusPersonne.親核加碼金額(List人員);
                int 上階主管分配金額2 = 0;
                int 保留款金額_總計 = 0; int 分配給轄下單位總預算_總計 = 0; int 固定預算_總計 = 0; int 加碼預算_總計 = 0;
                //本次發放金額(g)=固定預算+加碼預算+轄下主管調整金額加總+主管調整金額
                //Daniel: 要將主管分配金額在這邊加入差額，避免後續運算產生問題
                string 上層部門 = "";
                foreach (string 簽核者被上階分派的部門 in 簽核者被上階分派的部門s)
                {
                    List<string> 轄下部門清單 = new DAC_FL_DEPARTMENT_TW_V().GetDeptCTE_V(null, 簽核者被上階分派的部門).Select(p => DAC.GetString(p.Value)).ToList();
                    if (轄下部門清單.Contains(A.DEPT_SEQ_NO))
                    {
                        上層部門 = 簽核者被上階分派的部門;
                        BonusApprovalListList List人員 = dacBonusPersonne.Select_登入者所負責的人員名單(BonusProjectID: A.BonusProjectID, Login_EMP_SEQ_NO: 登入者seqno, 簽核者被上階分派的部門, true);
                        主管調整金額 = dacBonusPersonne.主管微調總金額(List人員);
                        轄下人員金額 = dacBonusPersonne.轄下人員金額(List人員);
                        親核固定金額 = dacBonusPersonne.親核固定金額(List人員);
                        親核加碼金額 = dacBonusPersonne.親核加碼金額(List人員);
                        if (is總經理層級)
                        {
                            上階主管分配金額2 = DAC.GetInt32(DAC.GetDecimal(StringEncrypt.aesDecryptBase64(A.Amount))
                            + DAC.GetDecimal(StringEncrypt.aesDecryptBase64(A.FlexibleBudget)));
                        }
                        else
                        {
                            DepartmentBudgetItem 上階主管分配 = new DAC_DepartmentBudget().Select(BonusProjectID: A.BonusProjectID, DEPT_SEQ_NO: 簽核者被上階分派的部門, ISAssign: true).FirstOrDefault();
                            上階主管分配金額2 = DAC.GetInt32(DAC.GetDecimal(StringEncrypt.aesDecryptBase64(上階主管分配.Amount))
                            + DAC.GetDecimal(StringEncrypt.aesDecryptBase64(上階主管分配.FlexibleBudget)));                
                        }
                        break;
                    }
                }
                if (error == "")
                {
                    var approvalCLASS = new ApprovalCLASS();
                    approvalCLASS.BonusProjectID = BonusProjectID;
                    approvalCLASS.DEPT_SEQ_NO = A.DEPT_SEQ_NO;
                    approvalCLASS.Result = (StringEncrypt.aesEncryptBase64(DAC.GetString(親核固定金額 + 親核加碼金額 + 轄下人員金額 + 主管調整金額)));
                    approvalCLASS.Difference = StringEncrypt.aesEncryptBase64(DAC.GetString(DAC.GetInt64(上階主管分配金額2) - DAC.GetInt64(親核固定金額 + 親核加碼金額 + 轄下人員金額 + 主管調整金額)));

                    approvalCLASS.含轄下 = false;
                    approvalCLASS.IS親核 = true;
                    approvalCLASS.簽核者被上階分派的部門 = 上層部門;
                    approvalCLASSList.Add(approvalCLASS);
                }
            }
            //被上級分派的部門


            //轄下部門則是同意轄下的送簽
            foreach (var D in dList)
            {
                if (error == "")
                {
                    var approvalCLASS = new ApprovalCLASS();
                    approvalCLASS.BonusProjectID = BonusProjectID;
                    approvalCLASS.DEPT_SEQ_NO = D.DEPT_SEQ_NO;

                    //轄下部門不計算Result、Difference
                    if (D.BPM_FormNO != null && D.BPM_FormNO != "")
                        approvalCLASS.GETBPM_NO = false;
                    else
                        approvalCLASS.GETBPM_NO = true;
                    approvalCLASS.含轄下 = true;
                    approvalCLASS.IS親核 = false;

                    foreach (string 簽核者被上階分派的部門 in 簽核者被上階分派的部門s)
                    {
                        List<string> 轄下部門清單 = new DAC_FL_DEPARTMENT_TW_V().GetDeptCTE_V(null, 簽核者被上階分派的部門).Select(p => DAC.GetString(p.Value)).ToList();
                        if (轄下部門清單.Contains(approvalCLASS.DEPT_SEQ_NO))
                        {
                            approvalCLASS.簽核者被上階分派的部門 = 簽核者被上階分派的部門;
                            break;
                        }
                    }

                    approvalCLASSList.Add(approvalCLASS);
                }
            }

            //取得BPM單號、同意
            if (error == "")
            {
                try
                {
                    #region BPM相關事件先處理，避免非預期情況發生時造成資料還原困擾
                    //若BPM發現異常導致後續沒有直營，需要協助處理資料時
                    //則建議針對 DepartmentBudget.Select_轄下已分派 的語法，將此表已簽資料改為非ISAssign，重新操作後執行 xx function再將ISAssign改回來即可

                    //紀錄送簽的部門以及BPM_NO
                    NameValueList BPM送簽部門 = new NameValueList();
                    if (is總經理層級 == false)
                    {
                        foreach (var 簽核者被上階分派的部門 in 簽核者被上階分派的部門s)
                        {
                            //取得此次要啟單的部門主管預算item資料
                            DepartmentBudgetItem item_Superior = _DepartmentBudget.Select(BonusProjectID: BonusProjectID, DEPT_SEQ_NO: 簽核者被上階分派的部門, ISAssign: true).FirstOrDefault();
                            string Superior_BPM_FormNO = item_Superior.BPM_FormNO;
                            //未起過單
                            if (string.IsNullOrWhiteSpace(Superior_BPM_FormNO))
                            {

                                //直接先啟單，取得BPM_FormNO
                                Superior_BPM_FormNO = GetBPM_FormNO(item_Superior);
                                //開單之後自己再送簽一次 (只送上層分派下來的部門，而非所有親核部門)
                                ws.ApproveForm(Bonus_FormKind, Superior_BPM_FormNO, 登入者工號, "Y", null);
                            }
                            //有起過單，為駁回重送
                            else
                            {
                                //只有被否決時才送簽，若狀態為1則已經送出，略過
                                if (item_Superior.BPM_Status == 1)
                                {
                                    //ws.ApproveForm(Bonus_FormKind, Superior_BPM_FormNO, (string)Session[PublicVariable.EMP_NO], "Y", null);
                                    SysLog.Write(LoginUserID, "獎金簽核", "DepartmentBudgetID: " + item_Superior.DepartmentBudgetID + " 已送簽過(略過)");
                                    continue;
                                }
                                //有被否決就要送簽
                                if (item_Superior.BPM_Status == 2)
                                {
                                    ws.ApproveForm(Bonus_FormKind, Superior_BPM_FormNO, 登入者工號, "Y", null);
                                }
                            }
                            BPM送簽部門.Add(new NameValueItem() { Name = 簽核者被上階分派的部門, Value = Superior_BPM_FormNO });
                        }
                    }


                    //宣告待送簽列表，送簽時一併送出
                    List<BPMWebServiceController.SignOffData> list_SignOff = new List<BPMWebServiceController.SignOffData>();
                    //不須送簽但須執行(發生錯誤時的補救)
                    List<BPMWebServiceController.SignOffData> list_SignOff_expect = new List<BPMWebServiceController.SignOffData>();
                    //找出其餘須往下送簽的部門
                    DepartmentBudgetList list_ISAssign = _DepartmentBudget.Select_轄下已分派byEMP_SEQ_NO(BonusProjectID: BonusProjectID, EMP_SEQ_NO: 登入者seqno);
                    foreach (DepartmentBudgetItem item_ISAssign in list_ISAssign)
                    {
                        //if(item_ISAssign.此次強制執行 == true)
                        //{
                        //    //不真正送簽，但程式需往後執行
                        //    //加入要簽核的List
                        //    list_SignOff_expect.Add(new BPMWebServiceController.SignOffData() { Form_NO = item_ISAssign.BPM_FormNO, formKind = Bonus_FormKind });
                        //    SysLog.Write(LoginUserID, "獎金簽核", "DepartmentBudgetID: " + item_ISAssign.DepartmentBudgetID + " 強制執行");
                        //    continue;
                        //}
                        //if(item_ISAssign.BPM_Status == 1)
                        //{
                        //    //已送簽過，略過
                        //    SysLog.Write(LoginUserID, "獎金簽核", "DepartmentBudgetID: " + item_ISAssign.DepartmentBudgetID + " 已送簽過(略過)");
                        //    continue;
                        //}
                        //這張單是否真的輪到他簽核
                        //其他模組皆用待我簽核列表帶出該人員的相關資料，故與BPM同步
                        //獎金模組是簽核當下才去找BPM要簽核的單子，故做了一個檢查的機制 (注意 GetSignOff 取得的單號前後有 ' 符號
                        bool IsSignOff = BPMWebController.GetSignOff_Bonus(登入者工號).Contains("'" + item_ISAssign.BPM_FormNO + "'");
                        //發現異常(不為可簽核部門)
                        if (IsSignOff == false)
                        {
                            //errorMessage = item_ISAssign.DEPT_NAME + "不為登入者可簽核的部門，BPM表單號: " + item_ISAssign.BPM_FormNO;
                            SysLog.Write(LoginUserID, "獎金簽核", item_ISAssign.DEPT_NAME + "不為登入者可簽核的部門，BPM表單號: " + item_ISAssign.BPM_FormNO);
                            continue;
                            //throw new Exception(errorMessage);
                        }

                        //加入要簽核的List
                        list_SignOff.Add(new BPMWebServiceController.SignOffData() { Form_NO = item_ISAssign.BPM_FormNO, formKind = Bonus_FormKind });
                    }
                    //一次送簽
                    BPMWebController.AgreeAll(list_SignOff, 登入者工號, out errorMessage);
                    if (string.IsNullOrEmpty(errorMessage) == false)
                    {
                        errorMessage = "轄下送簽時發生異常：" + errorMessage;
                        SysLog.Write(LoginUserID, "獎金簽核", errorMessage);
                        throw new Exception(errorMessage);
                    }
                    //加入須強制執行的list
                    //list_SignOff.AddRange(list_SignOff_expect);
                    #endregion

                    #region 寫入差額及轄下主管核定結果
                    _DepartmentBudget.AfterBonusSendForm_new(BonusProjectID, 登入者seqno, is總經理層級, 簽核者被上階分派的部門s, BPM送簽部門, approvalCLASSList);
                    #endregion

                    #region 總經理送出後將資料寫進總簽核
                    if (is總經理層級)
                    {
                        DAC_BonusFinalDistribution _BonusFinalDistribution = new DAC_BonusFinalDistribution();
                        DAC_BonusFinalDistributionList _BonusFinalDistributionList = new DAC_BonusFinalDistributionList();
                        
                        if (_BonusFinalDistribution.ResultInFinalDistribution(BonusProjectID, dList_親核, dList) == true)
                        {
                            //寫進總簽核核定
                            if (_BonusFinalDistributionList.ResultInFinalDistributionList(BonusProjectID) == false)
                            {
                                ApprovalMessage.error = "寫入總簽核核定時失敗";
                            }
                        }
                        else
                        {
                            ApprovalMessage.error = "寫入總簽核分配時失敗";
                        }
                    }
                    #endregion

                    //寫進主管各階層核定紀錄 (專案歷程記錄)
                    DAC_BonusRecordDetail _BonusRecordDetail = new DAC_BonusRecordDetail();
                    if (_BonusRecordDetail.ResultInBonusApprovalList(BonusProjectID) == false)
                        ApprovalMessage.error = "寫進主管各階層核定紀錄失敗";

                    if (string.IsNullOrWhiteSpace(ApprovalMessage.error))
                    {
                        ApprovalMessage.Approval = true;
                        ApprovalMessage.error = "";
                    }
                    else
                    {
                        ApprovalMessage.Approval = false;
                    }
                    
                }
                catch (Exception ex)
                {
                    ApprovalMessage.Approval = false;
                    ApprovalMessage.error = ex.Message;
                }
            }
            else
            {
                ApprovalMessage.Approval = false;
                ApprovalMessage.error = error;
            }

            return ApprovalMessage;
        }
        /// <summary>
        /// 送出簽核部門的邏輯為 【轄下】且有【被指派(ISAssign = true)】的部門
        /// </summary>
        /// <param name="BonusProjectID"></param>
        /// <returns></returns>
        private ApprovalMessageCLASS ApprovalAll_0713(int BonusProjectID)
        {
            //簽核邏輯待修正
            var ApprovalMessage = new ApprovalMessageCLASS();
            bool Approval = false;
            string error = "";
            string errorMessage = "";
            DAC_BonusProject _BonusProject = new DAC_BonusProject();
            DAC_FL_PERSONNEL_TW_V _FL_PERSONNEL_TW_V = new DAC_FL_PERSONNEL_TW_V();
            DAC_FL_DEPARTMENT_TW_V _DEPARTMENT_TW_V = new DAC_FL_DEPARTMENT_TW_V();
            DAC_BonusApprovalList _BonusApprovalList = new DAC_BonusApprovalList();
            DAC_DepartmentBudget _DepartmentBudget = new DAC_DepartmentBudget();

            var 登入者seqno = (string)Session[PublicVariable.EMP_SEQ_NO];
            var 登入者工號 = (string)Session[PublicVariable.EMP_NO];
            var 登入者部門 = (string)Session[PublicVariable.department_SEQ_NO]; //DEPT_SEQ_NO
            string 登入者層級 = new DAC_FL_DEPARTMENT_TW_V().GETLEVEL_CODE_主管層級(登入者seqno);
            string 登入者所屬部門主管DEPT_SEQ_NO = new DAC_FL_DEPARTMENT_TW_V().GETDEPT_主管部門(登入者seqno);
            List<string> 簽核者被上階分派的部門s = _DepartmentBudget.GetDepartmentBudget_被上級分派的部門(BonusProjectID, 登入者seqno);

            bool 保留預算 = (_BonusProject.SelectOne(BonusProjectID).FirstOrDefault()?.ReserveBudget ) == 1 ? true : false;
            //找出登入者的需簽核的部門
            var dList_親核 = _DepartmentBudget.Select(BonusProjectID: BonusProjectID, EMP_SEQ_NO: 登入者seqno, ISApproval: true);
            var dList = _DepartmentBudget.Select(BonusProjectID: BonusProjectID, EMP_SEQ_NO: 登入者seqno, ISApproval: false);
            
            //1) 先檢查
            List<ApprovalCLASS> approvalCLASSList = new List<ApprovalCLASS>();
            //親核部門則是送簽
            foreach (var A in dList_親核)
            {
                var dacBonusPersonne = new DAC_BonusApprovalList();
                //var List人員 = dacBonusPersonne.Select_人員_byDEPT_SEQ_NO(
                //登入者層級: 登入者層級
                //, 登入者: 登入者seqno,
                //登入者DEPT_SEQ_NO: A.DEPT_SEQ_NO,
                //BonusProjectID: A.BonusProjectID, 含轄下: true);
                int 主管調整金額 = 0;// dacBonusPersonne.主管微調總金額(List人員);
                int 轄下人員金額 = 0;// dacBonusPersonne.轄下人員金額(List人員);
                int 親核固定金額 = 0;// dacBonusPersonne.親核固定金額(List人員);
                int 親核加碼金額 = 0;// dacBonusPersonne.親核加碼金額(List人員);
                int 上階主管分配金額2 = 0;
                int 保留款金額_總計 = 0; int 分配給轄下單位總預算_總計 = 0; int 固定預算_總計 = 0; int 加碼預算_總計 = 0;
                //本次發放金額(g)=固定預算+加碼預算+轄下主管調整金額加總+主管調整金額
                //Daniel: 要將主管分配金額在這邊加入差額，避免後續運算產生問題
                string 上層部門 = "";
                foreach(string 簽核者被上階分派的部門 in 簽核者被上階分派的部門s)
                {
                    List<string> 轄下部門清單 = new DAC_FL_DEPARTMENT_TW_V().GetDeptCTE_V(null, 簽核者被上階分派的部門).Select(p => DAC.GetString(p.Value)).ToList();
                    if (轄下部門清單.Contains(A.DEPT_SEQ_NO))
                    {
                        上層部門 = 簽核者被上階分派的部門;

                        BonusApprovalListList List人員 = dacBonusPersonne.Select_人員_送出核定(登入者層級, 登入者seqno, 簽核者被上階分派的部門, A.BonusProjectID);
                        主管調整金額 = dacBonusPersonne.主管微調總金額(List人員);
                        轄下人員金額 = dacBonusPersonne.轄下人員金額(List人員);
                        親核固定金額 = dacBonusPersonne.親核固定金額(List人員);
                        親核加碼金額 = dacBonusPersonne.親核加碼金額(List人員);
                        if(登入者層級 != "2")
                        {
                            DepartmentBudgetItem 上階主管分配 = new DAC_DepartmentBudget().Select(BonusProjectID: A.BonusProjectID, DEPT_SEQ_NO: 簽核者被上階分派的部門, ISAssign: true).FirstOrDefault();
                            上階主管分配金額2 = DAC.GetInt32(DAC.GetDecimal(StringEncrypt.aesDecryptBase64(上階主管分配.Amount))
                            + DAC.GetDecimal(StringEncrypt.aesDecryptBase64(上階主管分配.FlexibleBudget)));
                        }
                        else
                        {
                            上階主管分配金額2 = DAC.GetInt32(DAC.GetDecimal(StringEncrypt.aesDecryptBase64(A.Amount))
                            + DAC.GetDecimal(StringEncrypt.aesDecryptBase64(A.FlexibleBudget)));
                        }

                        //var dept = new DAC_DepartmentBudget().Select(EMP_SEQ_NO: 登入者seqno, BonusProjectID: BonusProjectID);
                        //foreach (var T in dept)
                        //{
                        //    if (轄下部門清單.Contains(T.DEPT_SEQ_NO))
                        //    {
                        //        保留款金額_總計 += DAC.GetInt32(StringEncrypt.aesDecryptBase64(T.ReserveBudget));
                        //        分配給轄下單位總預算_總計 += DAC.GetInt32(StringEncrypt.aesDecryptBase64(T.Amount));
                        //        固定預算_總計 += DAC.GetInt32(StringEncrypt.aesDecryptBase64(T.FixedBudget));
                        //        加碼預算_總計 += DAC.GetInt32(StringEncrypt.aesDecryptBase64(T.UnFixedBudget));
                        //    }

                        //}
                        break;
                    }
                }
                if (error == "")
                {
                    var approvalCLASS = new ApprovalCLASS();
                    approvalCLASS.BonusProjectID = BonusProjectID;
                    approvalCLASS.DEPT_SEQ_NO = A.DEPT_SEQ_NO;
                    approvalCLASS.Result = (StringEncrypt.aesEncryptBase64(DAC.GetString(親核固定金額 + 親核加碼金額 + 轄下人員金額 + 主管調整金額)));
                    if (保留預算)
                    {
                        approvalCLASS.Difference = StringEncrypt.aesEncryptBase64(DAC.GetString(DAC.GetInt64(上階主管分配金額2) - DAC.GetInt64(親核固定金額 + 親核加碼金額 + 轄下人員金額 + 主管調整金額)));
                        //approvalCLASS.Difference = (StringEncrypt.aesEncryptBase64(DAC.GetString(DAC.GetInt32(保留款金額_總計 + 分配給轄下單位總預算_總計) - DAC.GetInt32(親核固定金額 + 親核加碼金額 + 轄下人員金額 + 主管調整金額) + DAC.GetInt64(上階主管分配金額2))));
                    }
                    else
                    {
                        ;
                        //approvalCLASS.Difference = StringEncrypt.aesEncryptBase64(DAC.GetString(DAC.GetInt64(上階主管分配金額2) - DAC.GetInt64(親核固定金額 + 親核加碼金額 + 轄下人員金額 + 主管調整金額)));
                        //approvalCLASS.Difference = (StringEncrypt.aesEncryptBase64(DAC.GetString(DAC.GetInt32(固定預算_總計 + 加碼預算_總計) - DAC.GetInt32(親核固定金額 + 親核加碼金額 + 轄下人員金額 + 主管調整金額) + DAC.GetInt64(上階主管分配金額2))));
                    }

                    approvalCLASS.含轄下 = false;
                    approvalCLASS.IS親核 = true;
                    approvalCLASS.簽核者被上階分派的部門 = 上層部門;
                    approvalCLASSList.Add(approvalCLASS);
                }
            }
            //被上級分派的部門


            //轄下部門則是同意轄下的送簽
            foreach (var D in dList)
            {
                if (error == "")
                {
                    var approvalCLASS = new ApprovalCLASS();
                    approvalCLASS.BonusProjectID = BonusProjectID;
                    approvalCLASS.DEPT_SEQ_NO = D.DEPT_SEQ_NO;

                    //轄下部門不計算Result、Difference
                    if (D.BPM_FormNO != null && D.BPM_FormNO != "")
                        approvalCLASS.GETBPM_NO = false;
                    else
                        approvalCLASS.GETBPM_NO = true;
                    approvalCLASS.含轄下 = true;
                    approvalCLASS.IS親核 = false;

                    foreach (string 簽核者被上階分派的部門 in 簽核者被上階分派的部門s)
                    {
                        List<string> 轄下部門清單 = new DAC_FL_DEPARTMENT_TW_V().GetDeptCTE_V(null, 簽核者被上階分派的部門).Select(p => DAC.GetString(p.Value)).ToList();
                        if (轄下部門清單.Contains(approvalCLASS.DEPT_SEQ_NO))
                        {
                            approvalCLASS.簽核者被上階分派的部門 = 簽核者被上階分派的部門;
                            break;
                        }
                    }

                    approvalCLASSList.Add(approvalCLASS);
                }
            }

            //取得BPM單號、同意
            if (error == "")
            {
                try
                {
                    #region BPM相關事件先處理，避免非預期情況發生時造成資料還原困擾
                    //若BPM發現異常導致後續沒有直營，需要協助處理資料時
                    //則建議針對 DepartmentBudget.Select_轄下已分派 的語法，將此表已簽資料改為非ISAssign，重新操作後執行 xx function再將ISAssign改回來即可

                    //紀錄送簽的部門以及BPM_NO
                    NameValueList BPM送簽部門 = new NameValueList();
                    if(登入者層級 != "2")
                    {
                        foreach (var 簽核者被上階分派的部門 in 簽核者被上階分派的部門s)
                        {
                            //取得此次要啟單的部門主管預算item資料
                            DepartmentBudgetItem item_Superior = _DepartmentBudget.Select(BonusProjectID: BonusProjectID, DEPT_SEQ_NO: 簽核者被上階分派的部門, ISAssign: true).FirstOrDefault();
                            string Superior_BPM_FormNO = item_Superior.BPM_FormNO;
                            //未起過單
                            if (string.IsNullOrWhiteSpace(Superior_BPM_FormNO))
                            {

                                //直接先啟單，取得BPM_FormNO
                                Superior_BPM_FormNO = GetBPM_FormNO(item_Superior);
                                //開單之後自己再送簽一次 (只送上層分派下來的部門，而非所有親核部門)
                                ws.ApproveForm(Bonus_FormKind, Superior_BPM_FormNO, 登入者工號, "Y", null);
                            }
                            //有起過單，為駁回重送
                            else
                            {
                                //只有被否決時才送簽，若狀態為1則已經送出，略過
                                if (item_Superior.BPM_Status == 1)
                                {
                                    //ws.ApproveForm(Bonus_FormKind, Superior_BPM_FormNO, (string)Session[PublicVariable.EMP_NO], "Y", null);
                                    SysLog.Write(LoginUserID, "獎金簽核", "DepartmentBudgetID: " + item_Superior.DepartmentBudgetID + " 已送簽過(略過)");
                                    continue;
                                }
                                //有被否決就要送簽
                                if (item_Superior.BPM_Status == 2)
                                {
                                    ws.ApproveForm(Bonus_FormKind, Superior_BPM_FormNO, 登入者工號, "Y", null);
                                }
                            }
                            BPM送簽部門.Add(new NameValueItem() { Name = 簽核者被上階分派的部門, Value = Superior_BPM_FormNO });
                        }
                    }
                    

                    //宣告待送簽列表，送簽時一併送出
                    List<BPMWebServiceController.SignOffData> list_SignOff = new List<BPMWebServiceController.SignOffData>();
                    //不須送簽但須執行(發生錯誤時的補救)
                    List<BPMWebServiceController.SignOffData> list_SignOff_expect = new List<BPMWebServiceController.SignOffData>();
                    //找出其餘須往下送簽的部門
                    DepartmentBudgetList list_ISAssign = _DepartmentBudget.Select_轄下已分派byEMP_SEQ_NO(BonusProjectID: BonusProjectID, EMP_SEQ_NO: 登入者seqno);
                    foreach (DepartmentBudgetItem item_ISAssign in list_ISAssign)
                    {
                        //if(item_ISAssign.此次強制執行 == true)
                        //{
                        //    //不真正送簽，但程式需往後執行
                        //    //加入要簽核的List
                        //    list_SignOff_expect.Add(new BPMWebServiceController.SignOffData() { Form_NO = item_ISAssign.BPM_FormNO, formKind = Bonus_FormKind });
                        //    SysLog.Write(LoginUserID, "獎金簽核", "DepartmentBudgetID: " + item_ISAssign.DepartmentBudgetID + " 強制執行");
                        //    continue;
                        //}
                        //if(item_ISAssign.BPM_Status == 1)
                        //{
                        //    //已送簽過，略過
                        //    SysLog.Write(LoginUserID, "獎金簽核", "DepartmentBudgetID: " + item_ISAssign.DepartmentBudgetID + " 已送簽過(略過)");
                        //    continue;
                        //}
                        //這張單是否真的輪到他簽核
                        //其他模組皆用待我簽核列表帶出該人員的相關資料，故與BPM同步
                        //獎金模組是簽核當下才去找BPM要簽核的單子，故做了一個檢查的機制 (注意 GetSignOff 取得的單號前後有 ' 符號
                        bool IsSignOff = BPMWebController.GetSignOff_Bonus(登入者工號).Contains("'" + item_ISAssign.BPM_FormNO + "'");
                        //發現異常(不為可簽核部門)
                        if (IsSignOff == false)
                        {
                            //errorMessage = item_ISAssign.DEPT_NAME + "不為登入者可簽核的部門，BPM表單號: " + item_ISAssign.BPM_FormNO;
                            SysLog.Write(LoginUserID, "獎金簽核", item_ISAssign.DEPT_NAME + "不為登入者可簽核的部門，BPM表單號: " + item_ISAssign.BPM_FormNO);
                            continue;
                            //throw new Exception(errorMessage);
                        }

                        //加入要簽核的List
                        list_SignOff.Add(new BPMWebServiceController.SignOffData() { Form_NO = item_ISAssign.BPM_FormNO, formKind = Bonus_FormKind });
                    }
                    //一次送簽
                    BPMWebController.AgreeAll(list_SignOff, 登入者工號, out errorMessage);
                    if (string.IsNullOrEmpty(errorMessage) == false)
                    {
                        errorMessage = "轄下送簽時發生異常：" + errorMessage;
                        SysLog.Write(LoginUserID, "獎金簽核", errorMessage);
                        throw new Exception(errorMessage);
                    }
                    //加入須強制執行的list
                    //list_SignOff.AddRange(list_SignOff_expect);
                    #endregion

                    #region 寫入差額及轄下主管核定結果
                    _DepartmentBudget.AfterBonusSendForm(BonusProjectID, 登入者seqno, 登入者所屬部門主管DEPT_SEQ_NO, 登入者層級, 簽核者被上階分派的部門s, BPM送簽部門, approvalCLASSList);
                    #endregion

                    ApprovalMessage.Approval = true;
                    ApprovalMessage.error = "";
                }
                catch (Exception ex)
                {
                    ApprovalMessage.Approval = false;
                    ApprovalMessage.error = ex.Message;
                }
            }
            else
            {
                ApprovalMessage.Approval = false;
                ApprovalMessage.error = error;
            }

            return ApprovalMessage;
        }
        /// <summary>
        /// 送出簽核部門的邏輯為 【轄下】且有【被指派(ISAssign = true)】的部門
        /// </summary>
        /// <param name="BonusProjectID"></param>
        /// <returns></returns>
        private ApprovalMessageCLASS ApprovalAll_bak(int BonusProjectID)
        {
            //簽核邏輯待修正
            var ApprovalMessage = new ApprovalMessageCLASS();
            bool Approval = false;
            string error = "";
            string errorMessage = "";
            DAC_BonusProject _BonusProject = new DAC_BonusProject();
            DAC_FL_PERSONNEL_TW_V _FL_PERSONNEL_TW_V = new DAC_FL_PERSONNEL_TW_V();
            DAC_FL_DEPARTMENT_TW_V _DEPARTMENT_TW_V = new DAC_FL_DEPARTMENT_TW_V();
            DAC_BonusApprovalList _BonusApprovalList = new DAC_BonusApprovalList();
            DAC_DepartmentBudget _DepartmentBudget = new DAC_DepartmentBudget();

            var 登入者seqno = (string)Session[PublicVariable.EMP_SEQ_NO];
            var 登入者工號 = (string)Session[PublicVariable.EMP_NO];
            var 登入者部門 = (string)Session[PublicVariable.department_SEQ_NO]; //DEPT_SEQ_NO
            string 登入者層級 = new DAC_FL_DEPARTMENT_TW_V().GETLEVEL_CODE_主管層級(登入者seqno);
            string 登入者所屬部門主管DEPT_SEQ_NO = new DAC_FL_DEPARTMENT_TW_V().GETDEPT_主管部門(登入者seqno);
            string 簽核者被上階分派的部門 = _DepartmentBudget.GetDepartmentBudget_被上級分派的部門_bak(BonusProjectID, 登入者seqno);

            bool 保留預算 = (_BonusProject.SelectOne(BonusProjectID).FirstOrDefault()?.ReserveBudget ?? 0) == 1 ? true : false;
            //找出登入者的需簽核的部門
            var dList_親核 = _DepartmentBudget.Select(BonusProjectID: BonusProjectID, EMP_SEQ_NO: 登入者seqno, ISApproval: true);
            var dList = _DepartmentBudget.Select(BonusProjectID: BonusProjectID, EMP_SEQ_NO: 登入者seqno, ISApproval: false);

            var dept = new DAC_DepartmentBudget().Select(EMP_SEQ_NO: 登入者seqno, BonusProjectID: BonusProjectID);
            int 保留款金額_總計 = 0; int 分配給轄下單位總預算_總計 = 0; int 固定預算_總計 = 0; int 加碼預算_總計 = 0;
            foreach (var T in dept)
            {
                保留款金額_總計 += DAC.GetInt32(StringEncrypt.aesDecryptBase64(T.ReserveBudget));
                分配給轄下單位總預算_總計 += DAC.GetInt32(StringEncrypt.aesDecryptBase64(T.Amount));
                固定預算_總計 += DAC.GetInt32(StringEncrypt.aesDecryptBase64(T.FixedBudget));
                加碼預算_總計 += DAC.GetInt32(StringEncrypt.aesDecryptBase64(T.UnFixedBudget));
            }

            //1) 先檢查
            List<ApprovalCLASS> approvalCLASSList = new List<ApprovalCLASS>();
            //親核部門則是送簽
            foreach (var A in dList_親核)
            {
                var dacBonusPersonne = new DAC_BonusApprovalList();
                var List人員 = dacBonusPersonne.Select_人員_byDEPT_SEQ_NO(
                登入者層級: 登入者層級
                , 登入者: 登入者seqno,
                登入者DEPT_SEQ_NO: A.DEPT_SEQ_NO,
                BonusProjectID: A.BonusProjectID, 含轄下: true);
                var 主管調整金額 = dacBonusPersonne.主管微調總金額(List人員);
                var 轄下人員金額 = dacBonusPersonne.轄下人員金額(List人員);
                var 親核固定金額 = dacBonusPersonne.親核固定金額(List人員);
                var 親核加碼金額 = dacBonusPersonne.親核加碼金額(List人員);
                //本次發放金額(g)=固定預算+加碼預算+轄下主管調整金額加總+主管調整金額


                if (error == "")
                {
                    var approvalCLASS = new ApprovalCLASS();
                    approvalCLASS.BonusProjectID = BonusProjectID;
                    approvalCLASS.DEPT_SEQ_NO = A.DEPT_SEQ_NO;
                    approvalCLASS.Result = (StringEncrypt.aesEncryptBase64(DAC.GetString(親核固定金額 + 親核加碼金額 + 轄下人員金額 + 主管調整金額)));
                    if (保留預算)
                    {
                        approvalCLASS.Difference = (StringEncrypt.aesEncryptBase64(DAC.GetString(DAC.GetInt32(保留款金額_總計 + 分配給轄下單位總預算_總計) - DAC.GetInt32(親核固定金額 + 親核加碼金額 + 轄下人員金額 + 主管調整金額))));
                    }
                    else
                    {
                        approvalCLASS.Difference = (StringEncrypt.aesEncryptBase64(DAC.GetString(DAC.GetInt32(固定預算_總計 + 加碼預算_總計) - DAC.GetInt32(親核固定金額 + 親核加碼金額 + 轄下人員金額 + 主管調整金額))));
                    }

                    approvalCLASS.含轄下 = false;
                    approvalCLASS.IS親核 = true;
                    approvalCLASSList.Add(approvalCLASS);
                }
            }

            //轄下部門則是同意轄下的送簽
            foreach (var D in dList)
            {
                if (error == "")
                {
                    var approvalCLASS = new ApprovalCLASS();
                    approvalCLASS.BonusProjectID = BonusProjectID;
                    approvalCLASS.DEPT_SEQ_NO = D.DEPT_SEQ_NO;

                    //轄下部門不計算Result、Difference
                    if (D.BPM_FormNO != null && D.BPM_FormNO != "")
                        approvalCLASS.GETBPM_NO = false;
                    else
                        approvalCLASS.GETBPM_NO = true;
                    approvalCLASS.含轄下 = true;
                    approvalCLASS.IS親核 = false;
                    approvalCLASSList.Add(approvalCLASS);
                }
            }

            //取得BPM單號、同意
            if (error == "")
            {
                try
                {
                    #region BPM相關事件先處理，避免非預期情況發生時造成資料還原困擾
                    //若BPM發現異常導致後續沒有直營，需要協助處理資料時
                    //則建議針對 DepartmentBudget.Select_轄下已分派 的語法，將此表已簽資料改為非ISAssign，重新操作後執行 xx function再將ISAssign改回來即可

                    //取得此次要啟單的部門主管預算item資料
                    DepartmentBudgetItem item_Superior = _DepartmentBudget.Select(BonusProjectID: BonusProjectID, DEPT_SEQ_NO: 簽核者被上階分派的部門, ISAssign: true).FirstOrDefault();
                    string Superior_BPM_FormNO = item_Superior.BPM_FormNO;
                    //未起過單
                    if (string.IsNullOrWhiteSpace(Superior_BPM_FormNO))
                    {
                        //直接先啟單，取得BPM_FormNO
                        Superior_BPM_FormNO = GetBPM_FormNO(item_Superior);
                        //開單之後自己再送簽一次 (只送上層分派下來的部門，而非所有親核部門)
                        ws.ApproveForm(Bonus_FormKind, Superior_BPM_FormNO, 登入者工號, "Y", null);
                    }
                    //有起過單，為駁回重送
                    else
                    {
                        //只有被否決時才送簽，若狀態為1則已經送出(非預期錯誤)，略過
                        if (item_Superior.BPM_Status == 1)
                        {
                            //ws.ApproveForm(Bonus_FormKind, Superior_BPM_FormNO, (string)Session[PublicVariable.EMP_NO], "Y", null);
                            SysLog.Write(LoginUserID, "獎金簽核", "DepartmentBudgetID: "+ item_Superior.DepartmentBudgetID + " 已送簽過(略過)");
                        }
                        //有被否決就要送簽
                        if(item_Superior.BPM_Status == 2)
                        {
                            ws.ApproveForm(Bonus_FormKind, Superior_BPM_FormNO, 登入者工號, "Y", null);
                        }
                    }

                    //宣告待送簽列表，送簽時一併送出
                    List<BPMWebServiceController.SignOffData> list_SignOff = new List<BPMWebServiceController.SignOffData>();
                    //找出其餘須往下送簽的部門
                    DepartmentBudgetList list_ISAssign = _DepartmentBudget.Select_轄下已分派byDEPT_SEQ_NO(BonusProjectID: BonusProjectID, Superior_DEPT_SEQ_NO: 簽核者被上階分派的部門);
                    foreach (DepartmentBudgetItem item_ISAssign in list_ISAssign)
                    {
                        //這張單是否真的輪到他簽核
                        //其他模組皆用待我簽核列表帶出該人員的相關資料，故與BPM同步
                        //獎金模組是簽核當下才去找BPM要簽核的單子，故做了一個檢查的機制 (注意 GetSignOff 取得的單號前後有 ' 符號
                        bool IsSignOff = BPMWebController.GetSignOff_Bonus(登入者工號).Contains("'" + item_ISAssign.BPM_FormNO + "'");
                        //發現異常(不為可簽核部門)
                        if (IsSignOff == false)
                        {
                            errorMessage = item_ISAssign.DEPT_NAME + "不為登入者可簽核的部門，BPM表單號: " + item_ISAssign.BPM_FormNO;
                            SysLog.Write(LoginUserID, "獎金簽核", errorMessage);
                            throw new Exception(errorMessage);
                        }

                        //加入要簽核的List
                        list_SignOff.Add(new BPMWebServiceController.SignOffData() { Form_NO = item_ISAssign.BPM_FormNO, formKind = Bonus_FormKind});
                    }
                    //一次送簽
                    BPMWebController.AgreeAll(list_SignOff, 登入者工號, out errorMessage);
                    if (string.IsNullOrEmpty(errorMessage) == false)
                    {
                        errorMessage = "轄下送簽時發生異常：" + errorMessage;
                        SysLog.Write(LoginUserID, "獎金簽核", errorMessage);
                        throw new Exception(errorMessage);
                    }
                    #endregion

                    #region 寫入差額及轄下主管核定結果
                    _DepartmentBudget.AfterBonusSendForm_bak(BonusProjectID, 登入者seqno, 登入者所屬部門主管DEPT_SEQ_NO, 登入者層級, 簽核者被上階分派的部門, Superior_BPM_FormNO, approvalCLASSList);
                    #endregion

                    ApprovalMessage.Approval = true;
                    ApprovalMessage.error = "";
                }
                catch (Exception ex)
                {
                    ApprovalMessage.Approval = false;
                    ApprovalMessage.error = ex.Message;
                }
            }
            else
            {
                ApprovalMessage.Approval = false;
                ApprovalMessage.error = error;
            }

            return ApprovalMessage;
        }
        private ApprovalMessageCLASS ApprovalAll_bak2(int BonusProjectID)
        {
            var ApprovalMessage = new ApprovalMessageCLASS();
            bool Approval = false;
            string error = "";
            string errorMessage = "";
            DAC_BonusProject _BonusProject = new DAC_BonusProject();
            DAC_FL_PERSONNEL_TW_V _FL_PERSONNEL_TW_V = new DAC_FL_PERSONNEL_TW_V();
            DAC_FL_DEPARTMENT_TW_V _DEPARTMENT_TW_V = new DAC_FL_DEPARTMENT_TW_V();
            DAC_BonusApprovalList _BonusApprovalList = new DAC_BonusApprovalList();
            DAC_DepartmentBudget _DepartmentBudget = new DAC_DepartmentBudget();

            var 登入者 = (string)Session[PublicVariable.EMP_SEQ_NO];
            var 登入者部門 = (string)Session[PublicVariable.department_SEQ_NO]; //DEPT_SEQ_NO
            string 登入者層級 = new DAC_FL_DEPARTMENT_TW_V().GETLEVEL_CODE_主管層級(登入者);
            string 登入者所屬部門主管DEPT_SEQ_NO = new DAC_FL_DEPARTMENT_TW_V().GETDEPT_主管部門(登入者);

            bool 保留預算 = (_BonusProject.SelectOne(BonusProjectID).FirstOrDefault()?.ReserveBudget ?? 0) == 1 ? true : false;
            //找出登入者的需簽核的部門
            var dList_親核 = _DepartmentBudget.Select(BonusProjectID: BonusProjectID, EMP_SEQ_NO: 登入者, ISApproval: true);
            var dList = _DepartmentBudget.Select(BonusProjectID: BonusProjectID, EMP_SEQ_NO: 登入者, ISApproval: false);
            //List<BudgetCLASS> dataPost = new List<BudgetCLASS>();

            var dept = new DAC_DepartmentBudget().Select(EMP_SEQ_NO: 登入者, BonusProjectID: BonusProjectID);
            int 保留款金額_總計 = 0; int 分配給轄下單位總預算_總計 = 0; int 固定預算_總計 = 0; int 加碼預算_總計 = 0;
            foreach (var T in dept)
            {
                保留款金額_總計 += DAC.GetInt32(StringEncrypt.aesDecryptBase64(T.ReserveBudget));
                分配給轄下單位總預算_總計 += DAC.GetInt32(StringEncrypt.aesDecryptBase64(T.Amount));
                固定預算_總計 += DAC.GetInt32(StringEncrypt.aesDecryptBase64(T.FixedBudget));
                加碼預算_總計 += DAC.GetInt32(StringEncrypt.aesDecryptBase64(T.UnFixedBudget));
            }

            //1) 先檢查
            List<ApprovalCLASS> approvalCLASSList = new List<ApprovalCLASS>();
            //親核部門則是送簽
            foreach (var A in dList_親核)
            {
                var dacBonusPersonne = new DAC_BonusApprovalList();
                var List人員 = dacBonusPersonne.Select_人員_byDEPT_SEQ_NO(
                登入者層級: 登入者層級
                , 登入者: 登入者,
                登入者DEPT_SEQ_NO: A.DEPT_SEQ_NO,
                BonusProjectID: A.BonusProjectID, 含轄下: true);
                var 主管調整金額 = dacBonusPersonne.主管微調總金額(List人員);
                var 轄下人員金額 = dacBonusPersonne.轄下人員金額(List人員);
                var 親核固定金額 = dacBonusPersonne.親核固定金額(List人員);
                var 親核加碼金額 = dacBonusPersonne.親核加碼金額(List人員);
                //本次發放金額(g)=固定預算+加碼預算+轄下主管調整金額加總+主管調整金額


                if (error == "")
                {
                    var approvalCLASS = new ApprovalCLASS();
                    approvalCLASS.BonusProjectID = BonusProjectID;
                    approvalCLASS.DEPT_SEQ_NO = A.DEPT_SEQ_NO;
                    approvalCLASS.Result = (StringEncrypt.aesEncryptBase64(DAC.GetString(親核固定金額 + 親核加碼金額 + 轄下人員金額 + 主管調整金額)));
                    if (保留預算)
                    {
                        approvalCLASS.Difference = (StringEncrypt.aesEncryptBase64(DAC.GetString(DAC.GetInt32(保留款金額_總計 + 分配給轄下單位總預算_總計) - DAC.GetInt32(親核固定金額 + 親核加碼金額 + 轄下人員金額 + 主管調整金額))));
                    }
                    else
                    {
                        approvalCLASS.Difference = (StringEncrypt.aesEncryptBase64(DAC.GetString(DAC.GetInt32(固定預算_總計 + 加碼預算_總計) - DAC.GetInt32(親核固定金額 + 親核加碼金額 + 轄下人員金額 + 主管調整金額))));
                    }
                    //產生BPM_FormNO
                    if (A.BPM_FormNO == null || A.BPM_FormNO == "")
                        approvalCLASS.GETBPM_NO = true;
                    //退回重送
                    else
                        approvalCLASS.GETBPM_NO = false;
                    approvalCLASS.含轄下 = false;
                    approvalCLASS.IS親核 = true;
                    approvalCLASSList.Add(approvalCLASS);
                }
            }

            //轄下部門則是送簽
            foreach (var D in dList)
            {
                var dacBonusPersonne = new DAC_BonusApprovalList();
                var List人員 = dacBonusPersonne.Select_人員_byDEPT_SEQ_NO(
                登入者層級: 登入者層級
                , 登入者: D.MAIN_LEADER_NO,
                登入者DEPT_SEQ_NO: D.DEPT_SEQ_NO,
                BonusProjectID: D.BonusProjectID, 含轄下: true);
                var 轄下人員金額 = dacBonusPersonne.轄下人員金額(List人員);

                if (error == "")
                {
                    var approvalCLASS = new ApprovalCLASS();
                    approvalCLASS.BonusProjectID = BonusProjectID;
                    approvalCLASS.DEPT_SEQ_NO = D.DEPT_SEQ_NO;

                    //轄下部門不計算Result、Difference

                    if (D.BPM_FormNO != null && D.BPM_FormNO != "")
                        approvalCLASS.GETBPM_NO = false;
                    else
                        approvalCLASS.GETBPM_NO = true;
                    approvalCLASS.含轄下 = true;
                    approvalCLASS.IS親核 = false;
                    approvalCLASSList.Add(approvalCLASS);
                }
            }

            //取得BPM單號、同意
            if (error == "")
            {
                try
                {
                    foreach (var App in approvalCLASSList)
                    {
                        var item = _DepartmentBudget.Select(BonusProjectID: BonusProjectID,
                            EMP_SEQ_NO: 登入者,
                            DEPT_SEQ_NO: DAC.GetString(App.DEPT_SEQ_NO)).FirstOrDefault();
                        if (item != null)
                        {
                            if (App.GETBPM_NO)
                            {
                                var BPM_FormNO = GetBPM_FormNO(item);
                                //開單之後自己再送簽一次
                                ws.ApproveForm(Bonus_FormKind, BPM_FormNO, (string)Session[PublicVariable.EMP_NO], "Y", null);
                                //var BPM_FormNO = "1";
                                //item.BPM_FormNO = BPM_FormNO;
                                item.BPM_FormNO = BPM_FormNO;
                                if (item.EMP_SEQ_NO == 登入者)
                                {
                                    item.BPM_Status = 1;
                                }
                                #region 寫入差額
                                if (App.IS親核)
                                {
                                    item.Result = App.Result;
                                    item.Difference = App.Difference;
                                }
                                #endregion
                                _DepartmentBudget.UpdateOne(item);

                                #region 回寫上一層
                                var item_上一層 = _DepartmentBudget.Select(BonusProjectID: BonusProjectID,
                                    EMP_SEQ_NO: new DAC_FL_DEPARTMENT_TW_V().取得上層部門主管(登入者所屬部門主管DEPT_SEQ_NO),
                                    DEPT_SEQ_NO: DAC.GetString(App.DEPT_SEQ_NO)).FirstOrDefault();
                                if (item_上一層 != null)
                                {
                                    item_上一層.BPM_FormNO = BPM_FormNO;
                                    item_上一層.BPM_Status = 1;
                                    item_上一層.Result = App.Result;
                                    item_上一層.Difference = App.Difference;
                                    _DepartmentBudget.UpdateOne(item_上一層);
                                }
                                #endregion

                                #region 人員名單往上送
                                new DAC_DepartmentBudget().WriteInDepartmentBudget_向上簽核_人員(App, BonusProjectID, 登入者, 登入者所屬部門主管DEPT_SEQ_NO, 登入者層級, App.含轄下);
                                #endregion
                            }
                            else
                            {
                                #region BPM同意_包含轄下

                                var list = new List<BPMWebServiceController.SignOffData>();
                                var 轄下List = _DepartmentBudget.查詢BPM_NO(登入者, DAC.GetInt32(BonusProjectID), new DAC_FL_DEPARTMENT_TW_V().GetChild_DEPT_SEQ_NO_List(登入者所屬部門主管DEPT_SEQ_NO, DAC_FL_DEPARTMENT_TW_V.資訊類型.代號));
                                foreach (var i in 轄下List)
                                {

                                    var BPMitem = new BPMWebServiceController.SignOffData() { Form_NO = i.BPM_FormNO, formKind = Bonus_FormKind };

                                    var 當前簽核者 = ws.GetApproveList(Bonus_FormKind, i.BPM_FormNO).Data.Where(d => d.AppStatus == "U").FirstOrDefault()?.AppEmpNo ?? "0";

                                    if ((string)Session[PublicVariable.EMP_NO] == 當前簽核者)
                                    {
                                        list.Add(BPMitem);
                                    }
                                }

                                BPMWebController.AgreeAll(list, DAC.GetString(Session[PublicVariable.EMP_NO]), out errorMessage);


                                if (item.EMP_SEQ_NO == 登入者)
                                {
                                    item.BPM_Status = 1;
                                }

                                #region 寫入差額
                                if (App.IS親核)
                                {
                                    item.Result = App.Result;
                                    item.Difference = App.Difference;
                                }
                                #endregion

                                _DepartmentBudget.UpdateOne(item);

                                #region 回寫上一層
                                var item_上一層 = _DepartmentBudget.Select(BonusProjectID: BonusProjectID,
                                    EMP_SEQ_NO: new DAC_FL_DEPARTMENT_TW_V().取得上層部門主管(登入者所屬部門主管DEPT_SEQ_NO),
                                    DEPT_SEQ_NO: DAC.GetString(App.DEPT_SEQ_NO)).FirstOrDefault();
                                if (item_上一層 != null)
                                {
                                    item_上一層.Result = App.Result;
                                    item_上一層.Difference = App.Difference;
                                    _DepartmentBudget.UpdateOne(item_上一層);
                                }
                                #endregion

                                #region 人員名單往上送
                                new DAC_DepartmentBudget().WriteInDepartmentBudget_向上簽核_人員(App, BonusProjectID, 登入者, 登入者所屬部門主管DEPT_SEQ_NO, 登入者層級, App.含轄下);
                                #endregion
                                #endregion
                            }

                            var BonusProject = new DAC_BonusProject().SelectOne(DAC.GetInt32(item.BonusProjectID)).FirstOrDefault();
                            #region 寄信
                            new DAC_EmailSendLog().BonusFixedBudget_Approval(BonusProject, item, item.DEPT_NAME, "完成", "");
                            #endregion
                        }
                    }
                    ApprovalMessage.Approval = true;
                    ApprovalMessage.error = "";
                }
                catch (Exception ex)
                {
                    ApprovalMessage.Approval = false;
                    ApprovalMessage.error = ex.Message;
                }
            }
            else
            {
                ApprovalMessage.Approval = false;
                ApprovalMessage.error = error;
            }

            return ApprovalMessage;
        }
        private ApprovalMessageCLASS ApprovalnZero_new(int BonusProjectID)
        {
            var ApprovalMessage = new ApprovalMessageCLASS();
            bool Approval = false;
            string error = "";
            DAC_BonusApprovalList dAC_BonusApprovalList = new DAC_BonusApprovalList();
            var 登入者 = (string)Session[PublicVariable.EMP_SEQ_NO];
            var 登入者DEPT_SEQ_NO = ((string)Session[PublicVariable.department_SEQ_NO]);
            //string 登入者層級 = new DAC_FL_DEPARTMENT_TW_V().GETLEVEL_CODE_主管層級(登入者);
            var 人員 = dAC_BonusApprovalList.SelectPage_人員_new(false, 0, 100000, BonusProjectID: BonusProjectID, 登入者: 登入者, ReSignerID: 登入者);
            if (人員 != null)
            {
                try
                {
                    foreach (var item in 人員)
                    {
                        if (item.PreAdjust_current == null || item.PreAdjust_current == "")
                        {
                            switch (item.LEVEL_CODE_current)
                            {
                                case "1":
                                    item.SignerDep1 = 登入者DEPT_SEQ_NO;
                                    item.PreLeaderID1 = 登入者;
                                    item.PreAdjust1 = StringEncrypt.aesEncryptBase64("0");
                                    break;
                                case "2":
                                    item.SignerDep2 = 登入者DEPT_SEQ_NO;
                                    item.PreLeaderID2 = 登入者;
                                    item.PreAdjust2 = StringEncrypt.aesEncryptBase64("0");
                                    break;
                                case "3":
                                    item.SignerDep3 = 登入者DEPT_SEQ_NO;
                                    item.PreLeaderID3 = 登入者;
                                    item.PreAdjust3 = StringEncrypt.aesEncryptBase64("0");
                                    break;
                                case "4":
                                    item.SignerDep4 = 登入者DEPT_SEQ_NO;
                                    item.PreLeaderID4 = 登入者;
                                    item.PreAdjust4 = StringEncrypt.aesEncryptBase64("0");
                                    break;
                                case "5":
                                    item.SignerDep5 = 登入者DEPT_SEQ_NO;
                                    item.PreLeaderID5 = 登入者;
                                    item.PreAdjust5 = StringEncrypt.aesEncryptBase64("0");
                                    break;
                                case "6":
                                    item.SignerDep6 = 登入者DEPT_SEQ_NO;
                                    item.PreLeaderID6 = 登入者;
                                    item.PreAdjust6 = StringEncrypt.aesEncryptBase64("0");
                                    break;
                                case "7":
                                    item.SignerDep7 = 登入者DEPT_SEQ_NO;
                                    item.PreLeaderID7 = 登入者;
                                    item.PreAdjust7 = StringEncrypt.aesEncryptBase64("0");
                                    break;
                                case "8":
                                    item.SignerDep8 = 登入者DEPT_SEQ_NO;
                                    item.PreLeaderID8 = 登入者;
                                    item.PreAdjust8 = StringEncrypt.aesEncryptBase64("0");
                                    break;
                                case "9":
                                    item.SignerDep9 = 登入者DEPT_SEQ_NO;
                                    item.PreLeaderID9 = 登入者;
                                    item.PreAdjust9 = StringEncrypt.aesEncryptBase64("0");
                                    break;
                                case "10":
                                    item.SignerDep10 = 登入者DEPT_SEQ_NO;
                                    item.PreLeaderID10 = 登入者;
                                    item.PreAdjust10 = StringEncrypt.aesEncryptBase64("0");
                                    break;
                                default:
                                    break;
                            }
                            dAC_BonusApprovalList.UpdateOne(item);
                        }
                    }
                    Approval = true;
                }
                catch (Exception ex)
                {
                    Approval = false;
                    error += ex.Message;
                }
            }
            else
            {
                Approval = false;
                error += "尚無核定人員";
            }
            ApprovalMessage.Approval = Approval;
            ApprovalMessage.error = error;
            return ApprovalMessage;
        }
        private ApprovalMessageCLASS ApprovalnZero(int BonusProjectID)
        {
            var ApprovalMessage = new ApprovalMessageCLASS();
            bool Approval = false;
            string error = "";
            DAC_BonusApprovalList dAC_BonusApprovalList = new DAC_BonusApprovalList();
            var 登入者 = (string)Session[PublicVariable.EMP_SEQ_NO];
            var 登入者DEPT_SEQ_NO = ((string)Session[PublicVariable.department_SEQ_NO]);
            string 登入者層級 = new DAC_FL_DEPARTMENT_TW_V().GETLEVEL_CODE_主管層級(登入者);
            var 人員 = dAC_BonusApprovalList.Select當下需簽核人員(登入者層級: 登入者層級, 登入者: 登入者,
                 BonusProjectID: BonusProjectID);
            if (人員 != null)
            {
                try
                {
                    foreach (var item in 人員)
                    {
                        if (item.PreAdjust_current == null || item.PreAdjust_current == "")
                        {
                            switch (登入者層級)
                            {
                                case "1":
                                    item.SignerDep1 = 登入者DEPT_SEQ_NO;
                                    item.PreLeaderID1 = 登入者;
                                    item.PreAdjust1 = StringEncrypt.aesEncryptBase64("0");
                                    break;
                                case "2":
                                    item.SignerDep2 = 登入者DEPT_SEQ_NO;
                                    item.PreLeaderID2 = 登入者;
                                    item.PreAdjust2 = StringEncrypt.aesEncryptBase64("0");
                                    break;
                                case "3":
                                    item.SignerDep3 = 登入者DEPT_SEQ_NO;
                                    item.PreLeaderID3 = 登入者;
                                    item.PreAdjust3 = StringEncrypt.aesEncryptBase64("0");
                                    break;
                                case "4":
                                    item.SignerDep4 = 登入者DEPT_SEQ_NO;
                                    item.PreLeaderID4 = 登入者;
                                    item.PreAdjust4 = StringEncrypt.aesEncryptBase64("0");
                                    break;
                                case "5":
                                    item.SignerDep5 = 登入者DEPT_SEQ_NO;
                                    item.PreLeaderID5 = 登入者;
                                    item.PreAdjust5 = StringEncrypt.aesEncryptBase64("0");
                                    break;
                                case "6":
                                    item.SignerDep6 = 登入者DEPT_SEQ_NO;
                                    item.PreLeaderID6 = 登入者;
                                    item.PreAdjust6 = StringEncrypt.aesEncryptBase64("0");
                                    break;
                                case "7":
                                    item.SignerDep7 = 登入者DEPT_SEQ_NO;
                                    item.PreLeaderID7 = 登入者;
                                    item.PreAdjust7 = StringEncrypt.aesEncryptBase64("0");
                                    break;
                                case "8":
                                    item.SignerDep8 = 登入者DEPT_SEQ_NO;
                                    item.PreLeaderID8 = 登入者;
                                    item.PreAdjust8 = StringEncrypt.aesEncryptBase64("0");
                                    break;
                                case "9":
                                    item.SignerDep9 = 登入者DEPT_SEQ_NO;
                                    item.PreLeaderID9 = 登入者;
                                    item.PreAdjust9 = StringEncrypt.aesEncryptBase64("0");
                                    break;
                                case "10":
                                    item.SignerDep10 = 登入者DEPT_SEQ_NO;
                                    item.PreLeaderID10 = 登入者;
                                    item.PreAdjust10 = StringEncrypt.aesEncryptBase64("0");
                                    break;
                                default:
                                    break;
                            }
                            dAC_BonusApprovalList.UpdateOne(item);
                        }
                    }
                    Approval = true;
                }
                catch (Exception ex)
                {
                    Approval = false;
                    error += ex.Message;
                }
            }
            else
            {
                Approval = false;
                error += "尚無核定人員";
            }
            ApprovalMessage.Approval = Approval;
            ApprovalMessage.error = error;
            return ApprovalMessage;
        }

        #endregion


        #region BonusProjectInformation
        [CheckLoginSessionExpired]
        public ActionResult BonusProjectInfoPopup(int BonusProjectID)
        {
            #region 宣告
            var _DEPARTMENT_TW_V = new DAC_FL_DEPARTMENT_TW_V();
            var _DepartmentBudget = new DAC_DepartmentBudget();
            var Project = new DAC_BonusProject();
            var _BonusApprovalList = new DAC_BonusApprovalList();
            var 登入者 = (string)Session[PublicVariable.EMP_SEQ_NO];
            var 登入者層級 = new DAC_FL_DEPARTMENT_TW_V().GETLEVEL_CODE_主管層級(登入者);
            var 有被分配部門 = true;
            var S_FixedBudget = new decimal(0);    // 總預算_固定預算
            var S_UnFixedBudget = new decimal(0);  // 總預算_非固定預算
            var S_Amt = new decimal(0); // 總預算(S)
            var A_Amt = new decimal(0); // 個人預算總合(A)
            var B_Amt = new decimal(0); // 固定預算總合(B)
            var C_Amt = new decimal(0); // 加碼預算總合(C)
            var D_Amt = new decimal(0); // 主管彈性可分配預算總額(D)
            var E_Amt = new decimal(0); // 轄下主管調整金額(E)
            var F_Amt = new decimal(0); // 主管調整總額(F)
            var G_Amt = new decimal(0); // 已核定總額(G) 
            var H_Amt = new decimal(0); // 剩餘預算(H)            
            var V_Amt = new decimal(0); // 保留金額總合(V)


            #endregion

            #region 計算
            var allotType = Project.GetAllotType(BonusProjectID) ?? 0;
            var reserveBudget = Project.ReserveBudget(BonusProjectID) ?? 0;
            var Data = Get單位(DAC.GetInt32(BonusProjectID), false);   //效能調整            
            D_Amt = Data.Sum(p => p.主管分配金額);                      //效能調整 //主管彈性可分配預算總額(D)
            V_Amt = Data.Sum(p => p.保留金額);                         //效能調整 //保留金額總合(V)

            var dList = _DepartmentBudget.Select(BonusProjectID: DAC.GetInt32(BonusProjectID), EMP_SEQ_NO: 登入者);


            //ViewBag.總預算 = 0; 
            //ViewBag.總預算_FixedBudget = 0; 
            //ViewBag.總預算_UnFixedBudget = 0; 
            //ViewBag.保留金額 = 0;
            //ViewBag.主管微調總金額 = 0; 
            //ViewBag.轄下主管調整總額 = 0; 
            //ViewBag.已核定總金額 = 0; 
            //ViewBag.可用餘額 = 0; 
            //ViewBag.btnApproval = false;

            var 獎金專案明細Item = new DAC_DepartmentBudget().獎金專案明細(DAC.GetInt32(BonusProjectID), 登入者);

            if (有被分配部門)
            {
                //S_FixedBudget = 獎金專案明細Item.總預算_FixedBudget;    // 總預算_固定預算
                //S_UnFixedBudget = 獎金專案明細Item.總預算_UnFixedBudget;  // 總預算_非固定預算

                WriteLog($"===計算 總預算_UnFixedBudget[AllotType={DAC.GetString(allotType)}]=== start");
                switch (DAC.GetString(allotType))
                {
                    case "1":
                        S_UnFixedBudget = 獎金專案明細Item.總預算_UnFixedBudget;
                        break;
                    case "2":
                        S_FixedBudget = 獎金專案明細Item.總預算_FixedBudget;
                        S_UnFixedBudget = 獎金專案明細Item.總預算_UnFixedBudget;
                        break;
                    case "3":
                        S_UnFixedBudget = 獎金專案明細Item.總預算_UnFixedBudget;
                        break;
                    case "4":
                        break;
                    default:
                        break;
                }
                WriteLog($"===計算 總預算_UnFixedBudget[AllotType={DAC.GetString(allotType)}]=== end");
   
                V_Amt = 獎金專案明細Item.保留金額總合; //效能調整 //保留金額總合(V)
                S_Amt = 獎金專案明細Item.總預算_FixedBudget+ 獎金專案明細Item.總預算_UnFixedBudget + 獎金專案明細Item.保留金額總合 + DAC.GetInt32(D_Amt); //效能調整 //總預算
                F_Amt = _BonusApprovalList.主管微調總金額_new(BonusProjectID, 登入者); // 效能調整 //主管調整總額  //WriteLog($"===計算 主管微調總金額=== end");
                E_Amt = _BonusApprovalList.轄下主管調整總額_new(DAC.GetInt32(BonusProjectID), 登入者); // 效能調整 //轄下主管調整總額(E)  //WriteLog($"===計算 轄下主管調整總額=== end");
                G_Amt = DAC.GetInt32(S_FixedBudget) + DAC.GetInt32(S_UnFixedBudget) + DAC.GetInt32(F_Amt) + DAC.GetInt32(E_Amt); // 效能調整 //已核定總金額(G)
                H_Amt = DAC.GetInt32(S_Amt) - DAC.GetInt32(G_Amt); // 效能調整 //剩餘預算(H)
                //ViewBag.btnApproval = _DepartmentBudget.Check是否可送簽_new(BonusProjectID, 登入者); // 效能調整//
                //ViewBag.預定簽核數 = _BonusApprovalList.SelectCount(BonusProjectID: BonusProjectID, ReSignerID: DAC.GetInt32(登入者));
                //ViewBag.本次簽核人員已全部送簽 = _BonusApprovalList.本次簽核人員已全部送簽(BonusProjectID, 登入者);// 效能調整//
                //WriteLog($"===Check是否可送簽=== end");
            }


            #endregion

            var data = new BonusProjInfoViewModel
            {
                ProjectId = BonusProjectID,
                ProjectName = Project.獎金專案名稱(BonusProjectID),
                BonusYear = Project.獎金年度(BonusProjectID),
                AllotType = allotType,
                IsReserveBudget = reserveBudget,
                S_Amount = S_Amt, // 總預算(S)
                A_Amount = A_Amt, // 個人預算總合(A)
                D_Amount = D_Amt, // 主管彈性可分配預算總額(D)
                V_Amount = V_Amt, // 保留金額總合(V)
                E_Amount = E_Amt, // 轄下主管調整金額(E) 
                F_Amount = F_Amt, // 主管調整總額(F)
                G_Amount = G_Amt, // 已核定總額(G) 
                H_Amount = H_Amt  // 剩餘預算(H)
            };



            return View(data);

        }

        [CheckLoginSessionExpired]
        public JsonResult BonusProjectInfo(int BonusProjectID)
        {
            #region 宣告
            var _DEPARTMENT_TW_V = new DAC_FL_DEPARTMENT_TW_V();
            var _DepartmentBudget = new DAC_DepartmentBudget();
            var Project = new DAC_BonusProject();
            var _BonusApprovalList = new DAC_BonusApprovalList();
            var 登入者 = (string)Session[PublicVariable.EMP_SEQ_NO];
            var 登入者層級 = new DAC_FL_DEPARTMENT_TW_V().GETLEVEL_CODE_主管層級(登入者);
            var 有被分配部門 = true;
            var S_FixedBudget = new decimal(0);    // 總預算_固定預算
            var S_UnFixedBudget = new decimal(0);  // 總預算_非固定預算
            var S_Amt = new decimal(0); // 總預算(S)
            var A_Amt = new decimal(0); // 個人預算總合(A)
            var B_Amt = new decimal(0); // 固定預算總合(B)
            var C_Amt = new decimal(0); // 加碼預算總合(C)
            var D_Amt = new decimal(0); // 主管彈性可分配預算總額(D)
            var E_Amt = new decimal(0); // 轄下主管調整金額(E)
            var F_Amt = new decimal(0); // 主管調整總額(F)
            var G_Amt = new decimal(0); // 已核定總額(G) 
            var H_Amt = new decimal(0); // 剩餘預算(H)            
            var V_Amt = new decimal(0); // 保留金額總合(V)


            #endregion

            #region 計算
            var allotType = Project.GetAllotType(BonusProjectID) ?? 0;
            var reserveBudget = Project.ReserveBudget(BonusProjectID) ?? 0;
            var Data = Get單位(DAC.GetInt32(BonusProjectID), false);   //效能調整            
            D_Amt = Data.Sum(p => p.主管分配金額);                      //效能調整 //主管彈性可分配預算總額(D)
            V_Amt = Data.Sum(p => p.保留金額);                         //效能調整 //保留金額總合(V)

            var dList = _DepartmentBudget.Select(BonusProjectID: DAC.GetInt32(BonusProjectID), EMP_SEQ_NO: 登入者);
            var 獎金專案明細Item = new DAC_DepartmentBudget().獎金專案明細(DAC.GetInt32(BonusProjectID), 登入者);

            if (有被分配部門)
            {
                //S_FixedBudget = 獎金專案明細Item.總預算_FixedBudget;    // 總預算_固定預算
                //S_UnFixedBudget = 獎金專案明細Item.總預算_UnFixedBudget;  // 總預算_非固定預算

                WriteLog($"===計算 總預算_UnFixedBudget[AllotType={DAC.GetString(allotType)}]=== start");
                switch (DAC.GetString(allotType))
                {
                    case "1":
                        S_UnFixedBudget = 獎金專案明細Item.總預算_UnFixedBudget;
                        break;
                    case "2":
                        S_FixedBudget = 獎金專案明細Item.總預算_FixedBudget;
                        S_UnFixedBudget = 獎金專案明細Item.總預算_UnFixedBudget;
                        break;
                    case "3":
                        S_UnFixedBudget = 獎金專案明細Item.總預算_UnFixedBudget;
                        break;
                    case "4":
                        break;
                    default:
                        break;
                }
                WriteLog($"===計算 總預算_UnFixedBudget[AllotType={DAC.GetString(allotType)}]=== end");

                V_Amt = 獎金專案明細Item.保留金額總合; //效能調整 //保留金額總合(V)
                S_Amt = 獎金專案明細Item.總預算_FixedBudget + 獎金專案明細Item.總預算_UnFixedBudget + 獎金專案明細Item.保留金額總合 + DAC.GetInt32(D_Amt); //效能調整 //總預算
                F_Amt = _BonusApprovalList.主管微調總金額_new(BonusProjectID, 登入者); // 效能調整 //主管調整總額  //WriteLog($"===計算 主管微調總金額=== end");
                E_Amt = _BonusApprovalList.轄下主管調整總額_new(DAC.GetInt32(BonusProjectID), 登入者); // 效能調整 //轄下主管調整總額(E)  //WriteLog($"===計算 轄下主管調整總額=== end");
                G_Amt = DAC.GetInt32(S_FixedBudget) + DAC.GetInt32(S_UnFixedBudget) + DAC.GetInt32(F_Amt) + DAC.GetInt32(E_Amt); // 效能調整 //已核定總金額(G)
                H_Amt = DAC.GetInt32(S_Amt) - DAC.GetInt32(G_Amt); // 效能調整 //剩餘預算(H)
            }

            #endregion

            var data = new BonusProjInfoViewModel
            {
                ProjectId = BonusProjectID,
                ProjectName = Project.獎金專案名稱(BonusProjectID),
                BonusYear = Project.獎金年度(BonusProjectID),
                AllotType = allotType,
                IsReserveBudget = reserveBudget,
                S_Amount = S_Amt, // 總預算(S)
                A_Amount = A_Amt, // 個人預算總合(A)
                D_Amount = D_Amt, // 主管彈性可分配預算總額(D)
                V_Amount = V_Amt, // 保留金額總合(V)
                E_Amount = E_Amt, // 轄下主管調整金額(E) 
                F_Amount = F_Amt, // 主管調整總額(F)
                G_Amount = G_Amt, // 已核定總額(G) 
                H_Amount = H_Amt  // 剩餘預算(H)
            };

            string val = string.Empty;

            val = JsonConvert.SerializeObject(data, Formatting.Indented, new JsonSerializerSettings
            {
                ReferenceLoopHandling = ReferenceLoopHandling.Ignore
            });

            return Json(val, JsonRequestBehavior.AllowGet);


        }




        #endregion

        #region InstructionPopup
        [CheckLoginSessionExpired]
        public ActionResult InstructionPopup(int BonusProjectId, int BonusApprovalLstID, string EmpSeqNo)
        {
            var dac_BonusAppLst = new DAC_BonusApprovalList();
            var data = new List<BonusInstructionListViewModel>();
            var 登入者 = (string)Session[PublicVariable.EMP_SEQ_NO];
            var 登入者層級 = new DAC_FL_DEPARTMENT_TW_V().GETLEVEL_CODE_主管層級(登入者);
            var 部門層級 = typeof(SystemVariable.部門層級_enum);

            var item = dac_BonusAppLst.Select_人員_byEMP_SEQ_NO(
                BonusProjectID: BonusProjectId,
                BonusApprovalListID: BonusApprovalLstID,
                EMP_SEQ_NO: EmpSeqNo
                ) != null
                ? dac_BonusAppLst.Select_人員_byEMP_SEQ_NO(
                BonusProjectID: BonusProjectId,
                BonusApprovalListID: BonusApprovalLstID,
                EMP_SEQ_NO: EmpSeqNo
                ).FirstOrDefault()
                : new BonusApprovalListItem();


            // 單筆取出 再根據登入者層級顯示資料 ( 項次 層級  主管工號 主管姓名  調整金額 主管備註 )
            var empLst = new DAC_FL_PERSONNEL_TW_V();
            var deptLst = new DAC_FL_DEPARTMENT_TW_V();

            for (var deptLv = 10; deptLv >= int.Parse(登入者層級); deptLv--)
            {
                var leader = new FL_PERSONNEL_TW_VItem();

                switch (deptLv)
                {
                    case 1:
                        // var DeptLevel = deptLst.GETLEVEL_CODE_主管層級(item.PreLeaderID1);
                        leader = empLst.Select(EMP_SEQ_NO: item.PreLeaderID1) != null
                                 ? empLst.Select(EMP_SEQ_NO: item.PreLeaderID1).FirstOrDefault()
                                 : null;

                        if (leader != null)
                            data.Add(new BonusInstructionListViewModel
                            {
                                Id = 0,
                                ProjectId = BonusProjectId,
                                AppLstId = BonusApprovalLstID,
                                EmpSeqNo = EmpSeqNo,
                                DeptLvNo = deptLv,
                                DeptLevel = deptLv.EnumGetName(部門層級, ""),// item.PreLeaderID1 == null ? "董事長" : deptLst.GETLEVEL_CODE_主管層級(item.PreLeaderID1),
                                DeptName = leader == null ? "" : leader.DEPT_NAME,
                                LeaderNo = leader == null ? "" : leader.EMP_NO,
                                LeaderName = leader == null ? "" : leader.EMP_NAME,
                                AdjustmentAmt = item.PreAdjust1_解密.ToString("N0"),
                                Instruction = item.Instruction1 == null ? "" : item.Instruction1
                            });

                        break;
                    case 2:

                        leader = empLst.Select(EMP_SEQ_NO: item.PreLeaderID2) != null
                                 ? empLst.Select(EMP_SEQ_NO: item.PreLeaderID2).FirstOrDefault()
                                 : new FL_PERSONNEL_TW_VItem();

                        if (leader != null)
                            data.Add(new BonusInstructionListViewModel
                            {
                                Id = 0,
                                ProjectId = BonusProjectId,
                                AppLstId = BonusApprovalLstID,
                                EmpSeqNo = EmpSeqNo,
                                DeptLvNo = deptLv,
                                DeptLevel = deptLv.EnumGetName(部門層級, ""),//item.PreLeaderID2 == null ? "總處" : deptLst.GETLEVEL_CODE_主管層級(item.PreLeaderID2),
                                DeptName = leader == null ? "" : leader.DEPT_NAME,
                                LeaderNo = leader == null ? "" : leader.EMP_NO,
                                LeaderName = leader == null ? "" : leader.EMP_NAME,
                                AdjustmentAmt = item.PreAdjust2_解密.ToString("N0"),
                                Instruction = item.Instruction2 == null ? "" : item.Instruction2
                            });

                        //data.Add("總處," + item.PreLeaderID2 + "," + "," + "," + item.PreAdjust2_解密 + "," + item.Instruction2); 
                        
                        break;
                    case 3:

                        leader = empLst.Select(EMP_SEQ_NO: item.PreLeaderID3) != null
                                 ? empLst.Select(EMP_SEQ_NO: item.PreLeaderID3).FirstOrDefault()
                                 : null;

                        if (leader != null)
                            data.Add(new BonusInstructionListViewModel
                            {
                                Id = 0,
                                ProjectId = BonusProjectId,
                                AppLstId = BonusApprovalLstID,
                                EmpSeqNo = EmpSeqNo,
                                DeptLvNo = deptLv,
                                DeptLevel = deptLv.EnumGetName(部門層級, ""),//item.PreLeaderID3 == null ? "群2" : deptLst.GETLEVEL_CODE_主管層級(item.PreLeaderID3),
                                DeptName = leader == null ? "" : leader.DEPT_NAME,
                                LeaderNo = leader == null ? "" : leader.EMP_NO,
                                LeaderName = leader == null ? "" : leader.EMP_NAME,
                                AdjustmentAmt = item.PreAdjust3_解密.ToString("N0"),
                                Instruction = item.Instruction3 == null ? "" : item.Instruction3
                            });

                        //data.Add("總處," + item.PreLeaderID2 + "," + "," + "," + item.PreAdjust2_解密 + "," + item.Instruction2); 

                        break;
                    case 4:

                        leader = empLst.Select(EMP_SEQ_NO: item.PreLeaderID4) != null
                                 ? empLst.Select(EMP_SEQ_NO: item.PreLeaderID4).FirstOrDefault()
                                 : null;

                        if (leader != null)
                            data.Add(new BonusInstructionListViewModel
                            {
                                Id = 0,
                                ProjectId = BonusProjectId,
                                AppLstId = BonusApprovalLstID,
                                EmpSeqNo = EmpSeqNo,
                                DeptLvNo = deptLv,
                                DeptLevel = deptLv.EnumGetName(部門層級, ""),//item.PreLeaderID4 == null ? "群1" : deptLst.GETLEVEL_CODE_主管層級(item.PreLeaderID4),
                                DeptName = leader == null ? "" : leader.DEPT_NAME,
                                LeaderNo = leader == null ? "" : leader.EMP_NO,
                                LeaderName = leader == null ? "" : leader.EMP_NAME,
                                AdjustmentAmt = item.PreAdjust4_解密.ToString("N0"),
                                Instruction = item.Instruction4 == null ? "" : item.Instruction4
                            });

                        //data.Add("總處," + item.PreLeaderID2 + "," + "," + "," + item.PreAdjust2_解密 + "," + item.Instruction2); 

                        break;
                    case 5:

                        leader = empLst.Select(EMP_SEQ_NO: item.PreLeaderID5) != null
                                 ? empLst.Select(EMP_SEQ_NO: item.PreLeaderID5).FirstOrDefault()
                                 : null;

                        if (leader != null)
                            data.Add(new BonusInstructionListViewModel
                            {
                                Id = 0,
                                ProjectId = BonusProjectId,
                                AppLstId = BonusApprovalLstID,
                                EmpSeqNo = EmpSeqNo,
                                DeptLvNo = deptLv,
                                DeptLevel = deptLv.EnumGetName(部門層級, ""),//item.PreLeaderID5 == null ? "中心" : deptLst.GETLEVEL_CODE_主管層級(item.PreLeaderID5),
                                DeptName = leader == null ? "" : leader.DEPT_NAME,
                                LeaderNo = leader == null ? "" : leader.EMP_NO,
                                LeaderName = leader == null ? "" : leader.EMP_NAME,
                                AdjustmentAmt = item.PreAdjust5_解密.ToString("N0"),
                                Instruction = item.Instruction5 == null ? "" : item.Instruction5
                            });

                        //data.Add("總處," + item.PreLeaderID2 + "," + "," + "," + item.PreAdjust2_解密 + "," + item.Instruction2); 

                        break;
                    case 6:

                        leader = empLst.Select(EMP_SEQ_NO: item.PreLeaderID6) != null
                                 ? empLst.Select(EMP_SEQ_NO: item.PreLeaderID6).FirstOrDefault()
                                 : null;

                        if (leader != null)
                            data.Add(new BonusInstructionListViewModel
                            {
                                Id = 0,
                                ProjectId = BonusProjectId,
                                AppLstId = BonusApprovalLstID,
                                EmpSeqNo = EmpSeqNo,
                                DeptLvNo = deptLv,
                                DeptLevel = deptLv.EnumGetName(部門層級, ""),//item.PreLeaderID6 == null ? "處" : deptLst.GETLEVEL_CODE_主管層級(item.PreLeaderID6),
                                DeptName = leader == null ? "" : leader.DEPT_NAME,
                                LeaderNo = leader == null ? "" : leader.EMP_NO,
                                LeaderName = leader == null ? "" : leader.EMP_NAME,
                                AdjustmentAmt = item.PreAdjust6_解密.ToString("N0"),
                                Instruction = item.Instruction6 == null ? "" : item.Instruction6
                            });

                        //data.Add("總處," + item.PreLeaderID2 + "," + "," + "," + item.PreAdjust2_解密 + "," + item.Instruction2); 

                        break;
                    case 7:

                        leader = empLst.Select(EMP_SEQ_NO: item.PreLeaderID7) != null
                                 ? empLst.Select(EMP_SEQ_NO: item.PreLeaderID7).FirstOrDefault()
                                 : null;

                        if (leader != null)
                            data.Add(new BonusInstructionListViewModel
                            {
                                Id = 0,
                                ProjectId = BonusProjectId,
                                AppLstId = BonusApprovalLstID,
                                EmpSeqNo = EmpSeqNo,
                                DeptLvNo = deptLv,
                                DeptLevel = deptLv.EnumGetName(部門層級, ""),//item.PreLeaderID7 == null ? "部" : deptLst.GETLEVEL_CODE_主管層級(item.PreLeaderID7),
                                DeptName = leader == null ? "" : leader.DEPT_NAME,
                                LeaderNo = leader == null ? "" : leader.EMP_NO,
                                LeaderName = leader == null ? "" : leader.EMP_NAME,
                                AdjustmentAmt = item.PreAdjust7_解密.ToString("N0"),
                                Instruction = item.Instruction7 == null ? "" : item.Instruction7
                            });

                        //data.Add("總處," + item.PreLeaderID2 + "," + "," + "," + item.PreAdjust2_解密 + "," + item.Instruction2); 

                        break;
                    case 8:

                        leader = empLst.Select(EMP_SEQ_NO: item.PreLeaderID8) != null
                                 ? empLst.Select(EMP_SEQ_NO: item.PreLeaderID8).FirstOrDefault()
                                 : null;

                        if (leader != null)
                            data.Add(new BonusInstructionListViewModel
                            {
                                Id = 0,
                                ProjectId = BonusProjectId,
                                AppLstId = BonusApprovalLstID,
                                EmpSeqNo = EmpSeqNo,
                                DeptLvNo = deptLv,
                                DeptLevel = deptLv.EnumGetName(部門層級, ""),//item.PreLeaderID8 == null ? "課" : deptLst.GETLEVEL_CODE_主管層級(item.PreLeaderID8),
                                DeptName = leader == null ? "" : leader.DEPT_NAME,
                                LeaderNo = leader == null ? "" : leader.EMP_NO,
                                LeaderName = leader == null ? "" : leader.EMP_NAME,
                                AdjustmentAmt = item.PreAdjust8_解密.ToString("N0"),
                                Instruction = item.Instruction8 == null ? "" : item.Instruction8
                            });

                        //data.Add("總處," + item.PreLeaderID2 + "," + "," + "," + item.PreAdjust2_解密 + "," + item.Instruction2); 

                        break;
                    case 9:

                        leader = empLst.Select(EMP_SEQ_NO: item.PreLeaderID9) != null
                                 ? empLst.Select(EMP_SEQ_NO: item.PreLeaderID9).FirstOrDefault()
                                 : null;

                        if (leader != null)
                            data.Add(new BonusInstructionListViewModel
                            {
                                Id = 0,
                                ProjectId = BonusProjectId,
                                AppLstId = BonusApprovalLstID,
                                EmpSeqNo = EmpSeqNo,
                                DeptLvNo = deptLv,
                                DeptLevel = deptLv.EnumGetName(部門層級, ""),//item.PreLeaderID9 == null ? "組" : deptLst.GETLEVEL_CODE_主管層級(item.PreLeaderID9),
                                DeptName = leader == null ? "" : leader.DEPT_NAME,
                                LeaderNo = leader == null ? "" : leader.EMP_NO,
                                LeaderName = leader == null ? "" : leader.EMP_NAME,
                                AdjustmentAmt = item.PreAdjust9_解密.ToString("N0"),
                                Instruction = item.Instruction9 == null ? "" : item.Instruction9
                            });

                        //data.Add("總處," + item.PreLeaderID2 + "," + "," + "," + item.PreAdjust2_解密 + "," + item.Instruction2); 
                        break;
                    case 10:

                        leader = empLst.Select(EMP_SEQ_NO: item.PreLeaderID10) != null
                                 ? empLst.Select(EMP_SEQ_NO: item.PreLeaderID10).FirstOrDefault()
                                 : null;

                        if (leader != null)
                            data.Add(new BonusInstructionListViewModel
                            {
                                Id = 0,
                                ProjectId = BonusProjectId,
                                AppLstId = BonusApprovalLstID,
                                EmpSeqNo = EmpSeqNo,
                                DeptLvNo = deptLv,
                                DeptLevel = deptLv.EnumGetName(部門層級, ""),//item.PreLeaderID10 == null ? "班" : deptLst.GETLEVEL_CODE_主管層級(item.PreLeaderID10),
                                DeptName = leader == null ? "" : leader.DEPT_NAME,
                                LeaderNo = leader == null ? "" : leader.EMP_NO,
                                LeaderName = leader == null ? "" : leader.EMP_NAME,
                                AdjustmentAmt = item.PreAdjust10_解密.ToString("N0"),
                                Instruction = item.Instruction10 == null ? "" : item.Instruction10
                            });

                        break;


                }
            }

            return View(data);

        }

        [HttpPost]
        [CheckLoginSessionExpired]
        public JsonResult InstructionPopup(BonusInstructionListViewModel model)
        {
            var dac_BonusAppLst = new DAC_BonusApprovalList();

            if (model.SubmitButton == "儲存")
            {
                var tmpModel = dac_BonusAppLst.Select_人員_byEMP_SEQ_NO(
                             BonusProjectID: model.ProjectId,
                             BonusApprovalListID: model.AppLstId,
                             EMP_SEQ_NO: model.EmpSeqNo
                           );
                if (tmpModel != null)
                {
                    var item = tmpModel.FirstOrDefault();

                    switch (model.DeptLvNo)
                    {
                        case 1:  item.Instruction1  = model.Instruction;  break;
                        case 2:  item.Instruction2  = model.Instruction;  break;
                        case 3:  item.Instruction3  = model.Instruction;  break;
                        case 4:  item.Instruction4  = model.Instruction;  break;
                        case 5:  item.Instruction5  = model.Instruction;  break;
                        case 6:  item.Instruction6  = model.Instruction;  break;
                        case 7:  item.Instruction7  = model.Instruction;  break;
                        case 8:  item.Instruction8  = model.Instruction;  break;
                        case 9:  item.Instruction9  = model.Instruction;  break;
                        case 10: item.Instruction10 = model.Instruction;  break;
                        default:  break;
                    }
                    dac_BonusAppLst.UpdateOne(item);
                }
                else {
                    return null;
                }
            }
            else {
                return null;
            }



            string val = string.Empty;

            val = JsonConvert.SerializeObject(model, Formatting.Indented, new JsonSerializerSettings
            {
                ReferenceLoopHandling = ReferenceLoopHandling.Ignore
            });

            return Json(val, JsonRequestBehavior.AllowGet);

        }
        #endregion


        //private bool WriteToFinalSign(int BonusProjectID, DepartmentBudgetList dList_親核, DepartmentBudgetList dList, bool is總經理層級)
        //{
        //    bool result = false;
        //    int 總預算 = 1;
        //    int 非內部人 = 2;
        //    DAC_BonusProject _BonusProject = new DAC_BonusProject();
        //    DAC_FL_PERSONNEL_TW_V _FL_PERSONNEL_TW_V = new DAC_FL_PERSONNEL_TW_V();
        //    DAC_FL_DEPARTMENT_TW_V _DEPARTMENT_TW_V = new DAC_FL_DEPARTMENT_TW_V();
        //    DAC_BonusApprovalList _BonusApprovalList = new DAC_BonusApprovalList();
        //    DAC_DepartmentBudget _DepartmentBudget = new DAC_DepartmentBudget();


        //    var 登入者seqno = (string)Session[PublicVariable.EMP_SEQ_NO];
        //    var 登入者工號 = (string)Session[PublicVariable.EMP_NO];
        //    var 登入者部門 = (string)Session[PublicVariable.department_SEQ_NO]; //DEPT_SEQ_NO
        //    #region 總經理送出後將資料寫進總簽核
        //    DAC_BonusFinalDistribution _BonusFinalDistribution = new DAC_BonusFinalDistribution();
        //    //先寫入單一總經理資料
        //    if (is總經理層級)
        //    {
        //        int? AllotType = _BonusProject.SelectOne(BonusProjectID).FirstOrDefault()?.AllotType;
        //        int 本次可分配人數_A = 0, 無發放金額人數 = 0;
        //        string 實際分配人數比例_C = "";
        //        decimal 分給轄下單位總預算 = 0, 轄下主管核定結果 = 0, 差額 = 0, 保留金額 = 0, 主管分配金額 = 0, 公式計算預算 = 0, 部門調整金額 = 0, 董事長加碼 = 0, 固定預算 = 0, 加碼預算 = 0;
        //        #region 親核
        //        foreach (var d in dList_親核)
        //        {
        //            本次可分配人數_A += _BonusApprovalList.部門人數(BonusProjectID, d.DEPT_SEQ_NO, false);
        //            無發放金額人數 += _BonusApprovalList.部門無發放金額人數(登入者seqno, BonusProjectID, d.DEPT_SEQ_NO, false);
        //            部門調整金額 += _BonusApprovalList.部門調整後金額加總(登入者seqno, BonusProjectID, d.DEPT_SEQ_NO, false);
        //            主管分配金額 += DAC.GetDecimal(StringEncrypt.aesDecryptBase64(d.FlexibleBudget)); //主管分配金額(W)
        //        }
        //        BonusFinalDistributionItem item_final = _BonusFinalDistribution.Select(BonusProjectID: BonusProjectID, DistributionType: 非內部人, GM_SEQ_NO: 登入者seqno, DEPT_SEQ_NO: "親核").FirstOrDefault();
        //        item_final.ActualEMP_B = 本次可分配人數_A - 無發放金額人數;
        //        item_final.ActualRadio_C = 本次可分配人數_A == 0 ? "0%" : DAC.GetString(Math.Round(DAC.GetDecimal(本次可分配人數_A - 無發放金額人數) / DAC.GetDecimal(本次可分配人數_A), 1, MidpointRounding.AwayFromZero) *100) +"%";
        //        item_final.ManagersAdjust_E = StringEncrypt.aesEncryptBase64(DAC.GetString(部門調整金額));
        //        item_final.ManagersExtra_F = StringEncrypt.aesEncryptBase64(DAC.GetString(主管分配金額));
        //        item_final.SumBudget2_H = StringEncrypt.aesEncryptBase64(DAC.GetString(部門調整金額 + 主管分配金額 + 董事長加碼));
        //        if (_BonusFinalDistribution.UpdateOne(item_final) == false)
        //        {
        //            SysLog.Write(LoginUserID, "獎金簽核", "送出至獎金總簽核失敗，原因：" + item_final.clientMessage);
        //            return false;
        //        }
        //        #endregion


        //        #region 轄下
        //        foreach (var d in dList)
        //        {
        //            BonusFinalDistributionItem item_subfinal = _BonusFinalDistribution.Select(BonusProjectID: BonusProjectID, DistributionType: 非內部人, GM_SEQ_NO: 登入者seqno, DEPT_SEQ_NO: d.DEPT_SEQ_NO).FirstOrDefault();
        //            item_final.ActualEMP_B = 本次可分配人數_A - 無發放金額人數;
        //            item_final.ActualRadio_C = 本次可分配人數_A == 0 ? "0%" : DAC.GetString(Math.Round(DAC.GetDecimal(本次可分配人數_A - 無發放金額人數) / DAC.GetDecimal(本次可分配人數_A), 1, MidpointRounding.AwayFromZero) * 100) + "%";
        //            item_final.ManagersAdjust_E = StringEncrypt.aesEncryptBase64(DAC.GetString(部門調整金額));
        //            item_final.ManagersExtra_F = StringEncrypt.aesEncryptBase64(DAC.GetString(主管分配金額));
        //            item_final.SumBudget2_H = StringEncrypt.aesEncryptBase64(DAC.GetString(部門調整金額 + 主管分配金額 + 董事長加碼));

        //            BonusFinalSign_NonInsider subItem = new BonusFinalSign_NonInsider();
        //            subItem.GM_NAME = 登入者seqno;
        //            subItem.總經理已送簽 = 總經理已送簽;
        //            subItem.DEPT_NAME = d.DEPT_NAME;
        //            subItem.MANAGER_NAME = d.MAIN_LEADER_NAME;
        //            subItem.本次可分配人數_A = _BonusApprovalList.部門人數(BonusProjectID, d.DEPT_SEQ_NO, true, d.MAIN_LEADER_NO);
        //            subItem.實際分配人數_B = subItem.本次可分配人數_A - _BonusApprovalList.部門無發放金額人數(emp_da, BonusProjectID, d.DEPT_SEQ_NO, true);
        //            subItem.實際分配人數比例_C = subItem.本次可分配人數_A == 0 ? "0%" : DAC.GetString(Math.Round(DAC.GetDecimal(subItem.實際分配人數_B) / DAC.GetDecimal(subItem.本次可分配人數_A) * 100, 1, MidpointRounding.AwayFromZero)) + "%";
        //            subItem.公式計算預算_D = AllotType == 1 || AllotType == 3 ? d.加碼預算 : (d.加碼預算 + d.固定預算); //對於年終跟分紅，個人預算(U)即加碼預算
        //            subItem.主管調整後金額_E = _BonusApprovalList.部門調整後金額加總(GM_EMP_SEQ_NO, BonusProjectID, d.DEPT_SEQ_NO, true);
        //            subItem.主管額外可分配預算_F = DAC.GetInt64(StringEncrypt.aesDecryptBase64(d.FlexibleBudget));
        //            subItem.董事長加碼_G = _BonusApprovalList.部門董事長加碼金額(BonusProjectID, d.DEPT_SEQ_NO, true);
        //            subItem.合計發送總額_H = subItem.主管調整後金額_E + subItem.主管額外可分配預算_F + subItem.董事長加碼_G;
        //            subList.Add(subItem);
        //        }
        //        #endregion



        //    }



        //    _BonusProject.GetProjectGMList(BonusProjectID);
        //    #endregion
        //    return result;
        //}
    }

    #region 自訂Class
    public class BudgetCLASS
    {
        public int BonusProjectID { get; set; }
        public string DEPT_SEQ_NO { get; set; }
        //單位
        public string DEPT_NO { get; set; }
        //個人預算
        public string FixedBudget { get; set; }
        //保留預算百分比
        public string ReserveBudgetRatio { get; set; }
        //分給轄下單位總預算
        public string Amount { get; set; }
        //轄下主管核定結果
        public string Result { get; set; }
        //是否分配轄下
        public bool Mainclick { get; set; }

    }
    public class ApprovalCLASS
    {
        public int BonusProjectID { get; set; }
        public string DEPT_SEQ_NO { get; set; }
        public string Result { get; set; }
        public string Difference { get; set; }
        public bool GETBPM_NO { get; set; }
        public bool 含轄下 { get; set; }
        public bool IS親核 { get; set; }
        public string 簽核者被上階分派的部門 { get; set; }
    }
    public class ApprovalMessageCLASS
    {
        public bool Approval { get; set; }
        public string error { get; set; }
    }
    #endregion


}