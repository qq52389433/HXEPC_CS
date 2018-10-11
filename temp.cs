        private void btnOK_Click(object sender, EventArgs e)
        {
              try
            {
                if (this.IsUpEdition && String.IsNullOrEmpty(this.txtChangeReason.Text))
                    {
                        MessageBox.Show("请选择设计变更原因！");
                        return;
                    }
                if (this.txtDocNum2.Text.Trim() == "00")
                {
                    MessageBox.Show("请正确填写资料单编号！不能以第00号命名");
                }
                else if (((this.m_ucFileContainer.Items.Count > 0) || (DialogResult.No != MessageBox.Show("当前尚未选择附件，是否确定不选择附件提资？", "操作确认", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk))) && (this.m_Project != null))
                {
                    this.m_Project.dBSource.ProgramRun = true;
                    if (this.CheckInputOK())
                    {
                        this.Cursor = Cursors.WaitCursor;
                        List<TempDefn> tempDefnByCode = this.m_Project.dBSource.GetTempDefnByCode("GEDI_INTERFACEPROJECT");
                        TempDefn defn = (tempDefnByCode != null) ? tempDefnByCode[0] : null;
                        TempDefn mTempDefn = null;
                        if (defn != null)
                        {
                            mTempDefn = defn.GetTempDefnByCode("GEDI_EXCHANGEDOC");
                        }
                        if (mTempDefn == null)
                        {
                            AssistFun.PopUpPrompt("没有与其相关的模板管理，创建无法正常完成");
                            this.Cursor = Cursors.Default;
                        }
                        else
                        {
                            string mProjectName = this.txtDocNum1.Text + "字第" + this.txtDocNum2.Text + "号" + this.label12.Text;
                            if (!this.IsUpEdition && (this.MenuName == ""))
                            {
                                Project projectByName = this.m_Project.GetProjectByName(mProjectName);
                                if (projectByName == null)
                                {
                                    projectByName = this.m_Project.NewProject(mProjectName, this.txtTitle.Text);
                                }
                                if (projectByName == null)
                                {
                                    CDMS.Write("当前m_Project " + this.m_Project.ToString);
                                    MessageBox.Show("创建资料目录失败！");
                                    return;
                                }
                                this.m_Project = projectByName;
                            }
                            AttrData attrDataByKeyWord = this.m_Project.GetAttrDataByKeyWord("EXCHANGETITLE");
                            if (attrDataByKeyWord != null)
                            {
                                attrDataByKeyWord.SetCodeDesc(this.txtTitle.Text);
                            }
                            if (this.IsUpEdition)
                            {
                                AttrData data2 = this.m_Project.GetAttrDataByKeyWord("UPEDITIONNUM");
                                if (data2 != null)
                                {
                                    data2.SetCodeDesc((Convert.ToInt32(data2.ToString) + 1).ToString());
                                }
                            }
                            this.m_Project.AttrDataList.SaveData();
                            string codeDescStr = this.ProcessEnclosure(mProjectName + " 互提资料单");
                            string str3 = "";
                            IEnumerable<string> source = from docx in this.m_Project.DocList select docx.Code;
                            str3 = mProjectName + " 互提资料单";
                            if (source.Contains<string>(str3))
                            {
                                for (int i = 1; i < 0x3e8; i++)
                                {
                                    str3 = mProjectName + " 互提资料单" + i.ToString();
                                    if (!source.Contains<string>(str3))
                                    {
                                        break;
                                    }
                                }
                            }
                            Doc item = this.m_Project.NewDoc(str3 + ".docx", str3, "", mTempDefn);
                            if (item == null)
                            {
                                AssistFun.PopUpPrompt("新建互提资料单出错！");
                                this.Cursor = Cursors.Default;
                            }
                            else
                            {
                                string str4 = "";
                                string str5 = "";
                                for (int j = 0; j < this.cblInProfession.CheckedItems.Count; j++)
                                {
                                    string str6 = this.cblInProfession.CheckedItems[j].ToString();
                                    str4 = str4 + str6 + ",";
                                    str5 = str5 + str6.Substring(str6.IndexOf("__") + 2) + "\n";
                                }
                                if (str4.EndsWith(","))
                                {
                                    str4 = str4.Remove(str4.Length - 1);
                                }
                                if (str5.EndsWith("\n"))
                                {
                                    str5 = str5.Remove(str5.Length - 1);
                                }
                                item.GetAttrDataByKeyWord("ED_PROFESSION").SetCodeDesc(str4);
                                item.GetAttrDataByKeyWord("ED_TITLE").SetCodeDesc(this.txtTitle.Text);
                                item.GetAttrDataByKeyWord("ED_CONTENT").SetCodeDesc(this.txtContent.Text);
                                item.GetAttrDataByKeyWord("ED_FILENO").SetCodeDesc(this.txtDocNum1.Text + "字第" + this.txtDocNum2.Text + "号");
                                item.GetAttrDataByKeyWord("ED_CHECKER").SetCodeDesc(this.txtChecker.Text);
                                item.GetAttrDataByKeyWord("ED_AUDITOR").SetCodeDesc(this.txtAduit.Text);
                                item.GetAttrDataByKeyWord("ED_SECTION").SetCodeDesc(this.txtSection.Text);
                                item.GetAttrDataByKeyWord("ED_RECEIVER").SetCodeDesc(this.txtReceiver.Text);
                                item.GetAttrDataByKeyWord("ED_MAKEDATE").SetCodeDesc(DateTime.Now.ToString("yyyy-MM-dd"));
                                item.GetAttrDataByKeyWord("ED_ENCLOSURE").SetCodeDesc(codeDescStr);
                                //设计变更原因
                                if (this.txtChangeReason.Text != "")
                                {
                                    item.GetAttrDataByKeyWord("ED_CHANGEREASON").SetCodeDesc(this.txtChangeReason.Text);
                                }
                                //item.AttrDataList.SaveData();

                                AttrData data3 = item.GetAttrDataByKeyWord("ED_IMPORTANCEDOC");
                                if (this.rbPriImport.Checked)
                                {
                                    data3.SetCodeDesc("综合性重要资料");
                                }
                                else if (this.rbPriNormal.Checked)
                                {
                                    data3.SetCodeDesc("一般资料");
                                }
                                else if (this.rbPriProfess.Checked)
                                {
                                    data3.SetCodeDesc("专业性重要资料");
                                }
                                data3 = item.GetAttrDataByKeyWord("ED_EXCHANGESTATUS");
                                if (this.rbStatePre.Checked)
                                {
                                    data3.SetCodeDesc(this.rbStatePre.Text);
                                }
                                else
                                {
                                    data3.SetCodeDesc(this.rbStateFin.Text);
                                }
                                item.AttrDataList.SaveData();
                                this.m_DocList.Add(item);
                                data3 = item.Project.GetAttrDataByKeyWord("ED_IMPORTANCE");
                                if (this.rbPriImport.Checked)
                                {
                                    data3.SetCodeDesc("综合性重要资料");
                                }
                                else if (this.rbPriNormal.Checked)
                                {
                                    data3.SetCodeDesc("一般资料");
                                }
                                else if (this.rbPriProfess.Checked)
                                {
                                    data3.SetCodeDesc("专业性重要资料");
                                }
                                item.Project.AttrDataList.SaveData();
                                Hashtable htUserKeyWord = new Hashtable();
                                if (this.rbType1.Checked)
                                {
                                    htUserKeyWord.Add("CHECKRESULT_AGREE", "▽");
                                }
                                else if (this.rbType2.Checked)
                                {
                                    htUserKeyWord.Add("CHECKRESULT_NOTAGREE", "▽");
                                }
                                else if (this.rbType3.Checked)
                                {
                                    htUserKeyWord.Add("CHECKRESULT_AGREEBUTVIEW", "▽");
                                }
                                if (this.rbStateFin.Checked)
                                {
                                    htUserKeyWord.Add("STATE_FIN", "√");
                                }
                                else
                                {
                                    htUserKeyWord.Add("STATE_PRE", "√");
                                }
                                htUserKeyWord.Add("DESIGNPHASE", this.txtPhase.Text);
                                htUserKeyWord.Add("DOCNUMBER1", mProjectName);
                                htUserKeyWord.Add("SENDDATE", DateTime.Now.ToString("yyyy年MM月dd日"));
                                htUserKeyWord.Add("OUTPROFESSION", this.txtOutProfession.Text);
                                htUserKeyWord.Add("INPROFESSION", str5);
                                htUserKeyWord.Add("RECEIVER", this.txtReceiver.Text);
                                htUserKeyWord.Add("MORESIGNIDEA", this.txtTitle.Text + "\r\n       " + this.txtContent.Text);
                                htUserKeyWord.Add("ENCLOSURE", codeDescStr);
                                htUserKeyWord.Add("PREPAREDSIGN1", this.m_Project.dBSource.LoginUser.Code);
                                this.Cursor = Cursors.WaitCursor;
                                string workingPath = this.m_Project.dBSource.LoginUser.WorkingPath;
                                AttrData data4 = this.m_Project.GetAttrDataByKeyWord("ISSAVE");
                                if (data4 != null)
                                {
                                    data4.SetCodeDesc("否");
                                }
                                string str7 = "专业间互提资料单";
                                FTPFactory factory = this.m_Project.Storage.FTP ?? new FTPFactory(this.m_Project.Storage);
                                string locFileName = this.m_Project.dBSource.LoginUser.WorkingPath + item.Code + ".docx";
                                factory.download(@"\ISO\" + str7 + ".docx", locFileName, false);
                                CDMSOffice office = new CDMSOffice {
                                    CloseApp = true,
                                    VisibleApp = false
                                };
                                office.Release(true);
                                office.WriteDataToDocument(item, locFileName, htUserKeyWord, htUserKeyWord);
                                factory.upload(locFileName, item.FullPathFile);
                                factory.close();
                                FileInfo info = new FileInfo(locFileName);
                                int length = (int) info.Length;
                                item.O_size = new int?(length);
                                item.Modify();
                                base.DialogResult = DialogResult.OK;
                                this.Cursor = Cursors.Default;
                                base.Close();
                                Yandingsoft.CDMS.PlugIns.CommonFunction.InsertDocListAndOpenDoc(this.m_DocList, item);
                                this.StartWorkFlow(item, "HTWORKFLOW");
                                this.m_Project.dBSource.ProgramRun = false;
                                CallBackParam param = new CallBackParam();
                                if (ExMenu.callTheApp != null)
                                {
                                    CallBackResult result;
                                    param.mask = 2;
                                    param.dList = new List<Doc>(new Doc[] { item });
                                    param.callType = enCallBackType.DocSelectD;
                                    param.dbs = item.dBSource;
                                    ExMenu.callTheApp(param, out result);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                AssistFun.PopUpPrompt(exception.Message);
                this.Cursor = Cursors.Default;
                this.m_Project.dBSource.ProgramRun = false;
            }
        }
