using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using AVEVA.CDMS.Server;
using AVEVA.CDMS.Common;
using AVEVA.CDMS.WebApi;

namespace AVEVA.CDMS.HXEPC_Plugins
{
    public class SelectSystem
    {
        //获取文件类型
        public static JObject GetSelectSystemList(string sid, string ProjectKeyword, string page, string Filter)
        {
            ExReJObject reJo = new ExReJObject();

            try
            {
                page = page ?? "1";
                string limit = "50";
                page = (Convert.ToInt32(page) - 1).ToString();
                int CurPage = Convert.ToInt32(page);

                User curUser = DBSourceController.GetCurrentUser(sid);
                if (curUser == null)
                {
                    reJo.msg = "登录验证失败！请尝试重新登录！";
                    return reJo.Value;
                }

                DBSource dbsource = curUser.dBSource;
                if (dbsource == null)
                {
                    reJo.msg = "登录验证失败！请尝试重新登录！";
                    return reJo.Value;
                }

                Project m_prj = dbsource.GetProjectByKeyWord(ProjectKeyword);
                if (m_prj == null)
                {
                    reJo.msg = "参数错误，目录不存在！";
                    return reJo.Value;
                }

                Project prjProject = CommonFunction.getParentProjectByTempDefn(m_prj, "HXNY_DOCUMENTSYSTEM");

                if (prjProject == null)
                {
                    reJo.msg = "获取项目目录失败！";
                    return reJo.Value;
                }

                Filter = Filter.Trim().ToLower();

                string curSystemCode = "";

                curSystemCode = prjProject.Code;//.GetValueByKeyWord("PRO_COMPANY");
                if (string.IsNullOrEmpty(curSystemCode))
                {
                    reJo.msg = "获取项目来源失败！";
                    return reJo.Value;
                }

                JArray jaData = new JArray();

                #region 添加厂房
                //获取所有参建单位
                List<DictData> dictDataList = dbsource.GetDictDataList("System");
                List<DictData> resultDDList = new List<DictData>();
                ////按代码排序
                //dictDataList.Sort(delegate (DictData x, DictData y)
                //{
                //    return x.O_Code.CompareTo(y.O_Code);
                //});


                foreach (DictData data6 in dictDataList)
                {
                    //判断是否符合过滤条件
                    if (!string.IsNullOrEmpty(Filter) &&
                        data6.O_Code.ToLower().IndexOf(Filter) < 0 &&
                        data6.O_Desc.ToLower().IndexOf(Filter) < 0 &&
                        data6.O_sValue1.ToLower().IndexOf(Filter) < 0
                        )
                    {
                        continue;
                    }

                    if (data6.O_sValue2 != curSystemCode) {
                        continue;
                    }
                        resultDDList.Add(data6);
                }

                #endregion


                reJo.total = resultDDList.Count();
                int ShowNum = 50;

                List<DictData> resDDList = resultDDList.Skip(CurPage * ShowNum).Take(ShowNum).ToList();

                foreach (DictData data6 in resDDList)
                {
                    //if (data6.O_sValue1 == curSystemCode)
                    {
                        JObject joData = new JObject(
                            new JProperty("systemType", "系统"),
                            new JProperty("systemId", data6.O_ID.ToString()),
                            new JProperty("systemCode", data6.O_Code),
                            new JProperty("systemDesc", data6.O_Desc)
                            );
                        jaData.Add(joData);
                    }
                }


                //获取所有项目部门（区分项目）
                //List<DictData> departDdList = dbsource.GetDictDataList("DeparDate");

                //foreach (DictData data6 in departDdList)
                //{
                //    //if (data6.O_sValue1 == curSystemCode)
                //    {
                //        JObject joData = new JObject(
                //            new JProperty("systemType", "项目部门"),
                //            new JProperty("systemId", data6.O_ID.ToString()),
                //            new JProperty("systemCode", data6.O_Code),
                //            new JProperty("systemDesc", data6.O_Desc)
                //            );
                //          jaData.Add(joData);
                //    }
                //}

                ////获取所有项目部门（不区分项目）
                //List<DictData> departDdList = dbsource.GetDictDataList("Communication");

                ////按代码排序
                //departDdList.Sort(delegate (DictData x, DictData y)
                //{
                //    return x.O_sValue1.CompareTo(y.O_sValue1);
                //});

                //foreach (DictData data6 in departDdList)
                //{
                //    //判断是否符合过滤条件
                //    if (!string.IsNullOrEmpty(Filter) &&
                //        data6.O_sValue1.ToLower().IndexOf(Filter) < 0 && data6.O_Desc.ToLower().IndexOf(Filter) < 0)
                //    {
                //        continue;
                //    }

                //    //if (data6.O_sValue1 == curSystemCode)
                //    if (!string.IsNullOrEmpty(data6.O_sValue1.Trim()))
                //    {
                //        JObject joData = new JObject(
                //            new JProperty("systemType", "项目部门"),
                //            new JProperty("systemId", data6.O_ID.ToString()),
                //            new JProperty("systemCode", data6.O_sValue1),
                //            new JProperty("systemDesc", data6.O_Desc)
                //            );
                //        jaData.Add(joData);
                //    }
                //}

                reJo.data = jaData;
                reJo.success = true;
                return reJo.Value;

            }
            catch (Exception e)
            {
                reJo.msg = e.Message;
                CommonController.WebWriteLog(reJo.msg);
            }

            return reJo.Value;
        }
    }

    struct systemObj {
        string Type{get;set;}
        string O_ID { get; set; }
        string Code { get; set; }
        string Desc { get; set; }
    }
}
