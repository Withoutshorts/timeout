<%@ Page Language="C#" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<script runat="server">

    const string PATH2UPLOAD = @"D:\webserver\wwwroot\timeout_xp\wwwroot\ver2_10\inc\upload\";

    ozJobData jobData;
    List<ozActivity> lstActivities;

    protected void Page_Load(object sender, EventArgs e)
    {
        if (this.ViewState["jobData"] != null)
        {
            jobData = (ozJobData)this.ViewState["jobData"];
            InitJobList();
        }
        if (this.ViewState["activities"] != null)
        {
            lstActivities = (List<ozActivity>)this.ViewState["activities"];
            InitActivity();
        }
    }

    protected void btnUpload_Click(object sender, EventArgs e)
    {        
        try
        {
            string fileName = FileUpload1.FileName;
            if (fileName == string.Empty)
            {
                lblFileName.Text = string.Empty;
                lblUploadStatus.Text = "Filen blev ikke uploadet.";
            }
            else
            {
                Boolean fileOK = false;
                string lto = string.Empty;
                if (!String.IsNullOrEmpty(Request["lto"]))
                    lto = Request["lto"];
                String path = PATH2UPLOAD + lto + "\\";
                if (FileUpload1.HasFile)
                {
                    String fileExtension = System.IO.Path.GetExtension(FileUpload1.FileName).ToLower();
                    String[] allowedExtensions = { ".xml", ".csv" };
                    for (int i = 0; i < allowedExtensions.Length; i++)
                    {
                        if (fileExtension == allowedExtensions[i])
                        {
                            fileOK = true;
                            break;
                        }
                    }

                    if (fileOK)
                    {
                        FileUpload1.PostedFile.SaveAs(path + FileUpload1.FileName);
                        lblFileName.Text = FileUpload1.FileName;
                        lblUploadStatus.Text = " uploadet.";
                    }
                    else
                    {
                        lblFileName.Text = string.Empty;
                        lblUploadStatus.Text = "Filen blev ikke uploadet.";
                    }
                }
            }
        }
        catch (Exception ex)
        {
            lblFileName.Text = string.Empty;
            lblStatus.Text = ex.Message;
        }
    }

    protected void btnReadJob_Click(object sender, EventArgs e)
    {
        if (lblFileName.Text != string.Empty && !String.IsNullOrEmpty("editor") && !String.IsNullOrEmpty("lto"))
        {
            //Init jobData with Xml data
            jobData = ozJob.ReadXml(PATH2UPLOAD + Request["lto"] + "\\" + lblFileName.Text);
            //Init viewstate jobData to keep the job list data while post back
            this.ViewState["jobData"] = jobData;
            //Init user interface with jobData
            txtJobName.Text = jobData.JobName;
            txtStartDato.Text = jobData.StartDate.ToShortDateString();
            txtSlutDato.Text = jobData.EndDate.ToShortDateString();
            txtInternNote.Text = jobData.InternNote;
            lblRekvisitionsNr.Text = jobData.OrderNo;
            lblBruttoOmsaetning.Text = jobData.Brutto.ToString();

            //Init Customer list for user to select
            ozCustomer.GetCustomerList(ref ddlKundenavn, jobData.CustomerName, Request["lto"]);
            if (ddlKundenavn.Items.Count > 0)
                jobData.CustomerID = SelectHelper.GetSelected(ref ddlKundenavn, jobData.CustomerName);

            //Init Contact person list for user to select
            ozCustomer.GetCustomerContactPersons(ref ddlCustomerContactPerson, jobData.CustomerContactPerson.Substring(0, 1), jobData.CustomerID, Request["lto"]);
            SelectHelper.GetSelected(ref ddlCustomerContactPerson, jobData.CustomerContactPerson);

            //Init job ansvarlig list for user to select
            ozCustomer.GetEmployee(ref ddlJobAnsvarlig, Request["lto"]);

            //Init view state activities
            lstActivities = ozActivity.GetActivities(Request["lto"]);
            this.ViewState["activities"] = lstActivities;

            //Init job list table on user interface
            TableRow headerRow = tblJobList.Rows[0];
            tblJobList.Rows.Clear();
            tblJobList.Rows.Add(headerRow);
            InitJobList();
            //Init activity listbox on user infterface
            ddlActivity.Items.Clear();
            InitActivity();
        }
    }

    protected void btnCreateJob_Click(object sender, EventArgs e)
    {
        //Update jobData object after user input
        try
        {
            jobData.CustomerID = int.Parse(ddlKundenavn.SelectedValue);
            jobData.DS_JobName = txtJobName.Text;
            jobData.LicenceID = Request["lto"];
            jobData.Editor = Request["editor"];
        }
        catch (Exception ex)
        {
            lblStatus.Text = "Create job before sending dataset error: we are experiencing a problem in creating CustomerID, DS_JobName, LicenceID, Editor: " + ex.Message;
        }
        try
        {
            if (ddlCustomerContactPerson.SelectedIndex != -1)
                jobData.CustomerContactPerson = ddlCustomerContactPerson.SelectedItem.Text;
            if (ddlCustomerContactPerson.SelectedIndex != -1)
                jobData.CustomerContactPersonId = int.Parse(ddlCustomerContactPerson.SelectedValue);
            if (ddlJobAnsvarlig.SelectedIndex != -1)
            {
                jobData.JobInChargePerson = ddlJobAnsvarlig.SelectedItem.Text;
                jobData.JobInChargePersonId = int.Parse(ddlJobAnsvarlig.SelectedValue);
            }
        }
        catch (Exception ex)
        {
            lblStatus.Text = "Create job before sending dataset error: we are experiencing a problem in creating CustomerContactPerson, CustomerContactPersonId, JobInChargePerson, JobInChargePersonId: " + ex.Message;
        }
        try
        {
            DateTime start = DateTime.MinValue;
            DateTime.TryParse(txtStartDato.Text, new System.Globalization.CultureInfo("en-US"), System.Globalization.DateTimeStyles.AllowWhiteSpaces, out start);
            if (start == DateTime.MinValue)
                DateTime.TryParse(txtStartDato.Text, new System.Globalization.CultureInfo("da-DK"), System.Globalization.DateTimeStyles.AllowWhiteSpaces, out start);
            jobData.StartDate = start;
            DateTime end = DateTime.MinValue;
            DateTime.TryParse(txtSlutDato.Text, new System.Globalization.CultureInfo("en-US"), System.Globalization.DateTimeStyles.AllowWhiteSpaces, out end);
            if (end == DateTime.MinValue)
                DateTime.TryParse(txtSlutDato.Text, new System.Globalization.CultureInfo("da-DK"), System.Globalization.DateTimeStyles.AllowWhiteSpaces, out end);
            jobData.EndDate = end;
        }
        catch (Exception ex)
        {
            lblStatus.Text = "Create job before sending dataset error: we are experiencing a problem in creating StartDate, EndDate: " + ex.Message;
        }
        try
        {
            jobData.JobDescription = txtJobDescription.Text;
            jobData.InternNote = txtInternNote.Text;
            jobData.Type = int.Parse(RadioType.SelectedValue);
            if (Request.Form["phaseType"] != null)
            {
                string phaseType = Request.Form["phaseType"];
                jobData.PhaseType = int.Parse(phaseType);
            }
        }
        catch (Exception ex)
        {
            lblStatus.Text = "Create job before sending dataset error: we are experiencing a problem in creating JobDescription, InternNote, Type, PhaseType: " + ex.Message;
        }
        try
        {
            jobData.ListActivityId.Clear();
            if (jobData.PhaseType == 0)
            {
                foreach (ListItem item in ddlActivity.Items)
                {
                    if (item.Selected)
                        jobData.ListActivityId.Add(int.Parse(item.Value));
                }
            }
        }
        catch (Exception ex)
        {
            lblStatus.Text = "Create job before sending dataset error: we are experiencing a problem in creating ListActivityId: " + ex.Message;
        }
        try
        {
            UpdateJobList();
        }
        catch (Exception ex)
        {
            lblStatus.Text = "Create job before sending dataset error: we are experiencing a problem in UpdateJobList: " + ex.Message;
        }

        System.Data.DataSet ds = new System.Data.DataSet();
        try
        {
            ds = ConvertHelper.Convert(jobData);
        }
        catch (Exception ex)
        {
            lblStatus.Text = "Convert to dataset error: " + ex.Message;
        }

        try
        {
            dk.outzource.TO_CREATEJOB createJobService = new dk.outzource.TO_CREATEJOB();
            createJobService.createjob(ds);
        }
        catch (Exception ex)
        {
            lblStatus.Text = "Create job service error: " + ex.Message;
        }

        try
        {
            GridView1.DataSource = ds;
            GridView1.DataBind();
        }
        catch
        {
        }
    }

    /// <summary>
    /// Update object jobData's CustomerID with dropdownlist ddlKundenavn selected value
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    protected void ddlKundenavn_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (ddlKundenavn.SelectedIndex != -1)
            jobData.CustomerID = int.Parse(ddlKundenavn.SelectedValue);
    }

    /// <summary>
    /// Initialize multi selectable list box ddlActivity with object lstActivities data
    /// </summary>
    private void InitActivity()
    {
        foreach (ozActivity item in lstActivities)
        {
            ListItem li = new ListItem(item.Name, item.id.ToString());
            ddlActivity.Items.Add(li);
        }
    }

    /// <summary>
    /// Initialize table tblJobList with jobData.JobList data
    /// </summary>
    private void InitJobList()
    {
        if (jobData.JobList.Count > 0)
            tblJobList.Visible = true;

        foreach (ozJob job in jobData.JobList)
        {
            TableRow tr = new TableRow();

            TableCell tcAmount = new TableCell();
            TextBox txtAmount = new TextBox();
            txtAmount.Text = job.Amount.ToString();
            tcAmount.Controls.Add(txtAmount);

            TableCell tcName = new TableCell();
            TextBox txtName = new TextBox();
            txtName.Text = "pos " + job.Pos + " " + job.MaterialNo + "(" + job.Description[0] + ")";
            tcName.Controls.Add(txtName);

            TableCell tcDate = new TableCell();
            TextBox txtDate = new TextBox();
            txtDate.Text = job.DeliverDate.ToShortDateString();
            tcDate.Controls.Add(txtDate);

            TableCell tcStkPrice = new TableCell();
            TextBox txtStkPrice = new TextBox();
            txtStkPrice.Text = job.Price.ToString();
            tcStkPrice.Controls.Add(txtStkPrice);

            TableCell tcPrice = new TableCell();
            TextBox txtPrice = new TextBox();
            txtPrice.Text = job.TotalPrice.ToString();
            tcPrice.Controls.Add(txtPrice);

            tr.Cells.Add(tcAmount);
            tr.Cells.Add(tcName);
            tr.Cells.Add(tcDate);
            tr.Cells.Add(tcStkPrice);
            tr.Cells.Add(tcPrice);
            tblJobList.Rows.Add(tr);
        }

        for (int i = 0; i < 20 - jobData.JobList.Count; i++)
        {
            TableRow tr = new TableRow();

            TableCell tcAmount = new TableCell();
            TextBox txtAmount = new TextBox();
            tcAmount.Controls.Add(txtAmount);

            TableCell tcName = new TableCell();
            TextBox txtName = new TextBox();
            tcName.Controls.Add(txtName);

            TableCell tcDate = new TableCell();
            TextBox txtDate = new TextBox();
            tcDate.Controls.Add(txtDate);

            TableCell tcStkPrice = new TableCell();
            TextBox txtStkPrice = new TextBox();
            tcStkPrice.Controls.Add(txtStkPrice);

            TableCell tcPrice = new TableCell();
            TextBox txtPrice = new TextBox();
            tcPrice.Controls.Add(txtPrice);

            tr.Cells.Add(tcAmount);
            tr.Cells.Add(tcName);
            tr.Cells.Add(tcDate);
            tr.Cells.Add(tcStkPrice);
            tr.Cells.Add(tcPrice);
            tblJobList.Rows.Add(tr);
        }
    }

    /// <summary>
    /// Update job list in jobData object
    /// </summary>
    private void UpdateJobList()
    {
        jobData.JobList.Clear();

        foreach (TableRow row in tblJobList.Rows)
        {
            ozJob job = new ozJob();

            if (row.Cells[0].GetType().Name == "TableHeaderCell")
                continue;

            int amount = 1;
            int.TryParse(((TextBox)row.Cells[0].Controls[0]).Text, out amount);
            job.DS_Amount = amount;
            job.DS_Name = ((TextBox)row.Cells[1].Controls[0]).Text.Trim();
            DateTime deliverDate = DateTime.Today;
            DateTime.TryParse(((TextBox)row.Cells[2].Controls[0]).Text, out deliverDate);
            job.DS_Date = deliverDate;
            if (jobData.EndDate == DateTime.MaxValue)
                jobData.EndDate = job.DS_Date;
            decimal price = 0;
            decimal.TryParse(((TextBox)row.Cells[3].Controls[0]).Text, out price);
            job.DS_Price = price;
            decimal totalPrice = 0;
            decimal.TryParse(((TextBox)row.Cells[4].Controls[0]).Text, out totalPrice);
            job.DS_TotalPrice = totalPrice;

            if (job.DS_Name != string.Empty)
                jobData.JobList.Add(job);
        }
    }
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Create job from xml file</title>
    <script type="text/javascript">
        function phaseTypeOnCheck(sender) {
            if (sender == "RadioPhaseType_Activity") {
                document.getElementById("RadioPhaseType_Phase").checked = false;
                document.getElementById("ddlActivity").style.display = "none";
                document.getElementById("phaseType").value = "1";
            }
            else {
                document.getElementById("RadioPhaseType_Activity").checked = false;
                document.getElementById("ddlActivity").style.display = "block";
                document.getElementById("phaseType").value = "0";
            }

        }
    </script>
    <style>
        .activityList
        {
            display: none;
            width: 300px;
            height: 300px;
            margin-left: 20px;
        }
    </style>
</head>
<body style="font-family: Arial" leftmargin="20" topmargin="20">
    <table cellpadding="0" cellspacing="5" border="0" width="100%">
        <tr>
            <td>
                <img alt="logo" src="../ill/outzource_logo_200.gif" />
            </td>
        </tr>
    </table>
    <form id="form1" runat="server">
    <div>
        <h4>
            TimeOut fil upload (opret job)</h4>
        <p>
            <asp:FileUpload ID="FileUpload1" runat="server" />
            <asp:Button ID="btnUpload" runat="server" OnClick="btnUpload_Click" Text="Upload" />
        </p>
        <p>
            <asp:Label ID="lblFileName" runat="server"></asp:Label>
            <asp:Label ID="lblUploadStatus" runat="server"></asp:Label>
        </p>
        <p>
            <asp:Label ID="lblStatus" runat="server" Font-Size="Small" ForeColor="Red"></asp:Label>
            <br />
            <asp:Button ID="btnReadJob" runat="server" OnClick="btnReadJob_Click" Text="Read job" />
        </p>
        <table>
            <tr>
                <td>
                    <asp:Label ID="Label1" runat="server" Text="Kundenavn:"></asp:Label>
                </td>
                <td>
                    <asp:DropDownList ID="ddlKundenavn" runat="server" OnSelectedIndexChanged="ddlKundenavn_SelectedIndexChanged">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="Label2" runat="server" Text="Kunde kontaktperson:"></asp:Label>
                </td>
                <td>
                    <asp:DropDownList ID="ddlCustomerContactPerson" runat="server">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="Label3" runat="server" Text="Jobnavn:"></asp:Label>
                </td>
                <td>
                    <asp:TextBox ID="txtJobName" runat="server"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="Label7" runat="server" Text="Jobansvarlig:"></asp:Label>
                </td>
                <td>
                    <asp:DropDownList ID="ddlJobAnsvarlig" runat="server">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="Label4" runat="server" Text="Start dato:"></asp:Label>
                </td>
                <td>
                    <asp:TextBox ID="txtStartDato" runat="server"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="Label5" runat="server" Text="Slut dato:"></asp:Label>
                </td>
                <td>
                    <asp:TextBox ID="txtSlutDato" runat="server"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="Label6" runat="server" Text="Jobbeskrivelse: "></asp:Label>
                </td>
                <td>
                    <asp:TextBox ID="txtJobDescription" runat="server" Columns="50" MaxLength="250"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="Label8" runat="server" Text="Intern note: "></asp:Label>
                </td>
                <td>
                    <asp:TextBox ID="txtInternNote" runat="server" Columns="50" MaxLength="250"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="Label10" runat="server" Text="Rekvisitions nr. / Inkøbsordre:"></asp:Label>
                </td>
                <td>
                    <asp:Label ID="lblRekvisitionsNr" runat="server"></asp:Label>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="Label12" runat="server" Text="Brutto omsætning: "></asp:Label>
                </td>
                <td>
                    <asp:Label ID="lblBruttoOmsaetning" runat="server"></asp:Label>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="Label9" runat="server" Text="Type:"></asp:Label>
                </td>
                <td>
                    <asp:RadioButtonList ID="RadioType" runat="server" RepeatDirection="Horizontal">
                        <asp:ListItem Value="0">Lbn. timer</asp:ListItem>
                        <asp:ListItem Selected="True" Value="1">Fast pris</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td style="vertical-align: top">
                    <asp:Label ID="Label11" runat="server" Text="Faser/Aktiveter:"></asp:Label>
                </td>
                <td>
                    <table>
                        <tr>
                            <td style="vertical-align: top">
                                <input type="radio" id="RadioPhaseType_Activity" name="RadioPhaseType_Activity" value="1"
                                    checked="checked" onclick="phaseTypeOnCheck('RadioPhaseType_Activity')" />
                                <asp:Label ID="Label13" runat="server" Text="Aktiveter"></asp:Label>
                                <input type="radio" id="RadioPhaseType_Phase" name="RadioPhaseType_Phase" value="0"
                                    onclick="phaseTypeOnCheck('RadioPhaseType_Phase')" />
                                <asp:Label ID="Label14" runat="server" Text="Faser"></asp:Label>
                                <input runat="server" type="hidden" id="phaseType" value="1" />
                            </td>
                            <td style="vertical-align: top">
                                <asp:ListBox ID="ddlActivity" ClientIDMode="Inherit" runat="server" SelectionMode="Multiple"
                                    CssClass="activityList"></asp:ListBox>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <asp:Table ID="tblJobList" runat="server" Visible="false">
            <asp:TableHeaderRow runat="server">
                <asp:TableHeaderCell runat="server">Antal</asp:TableHeaderCell>
                <asp:TableHeaderCell runat="server">Navn</asp:TableHeaderCell>
                <asp:TableHeaderCell ID="TableHeaderCell1" runat="server">Lev. dato</asp:TableHeaderCell>
                <asp:TableHeaderCell runat="server">Stk. pris</asp:TableHeaderCell>
                <asp:TableHeaderCell runat="server">Pris</asp:TableHeaderCell>
            </asp:TableHeaderRow>
        </asp:Table>
        <br />
        <asp:Button ID="btnCreateJob" runat="server" Text="Create job" OnClick="btnCreateJob_Click" />
        <br />
        <br />
        <hr />
        <b>TEST AREA</b>
        <hr />
        <br />
        <asp:GridView ID="GridView1" runat="server">
        </asp:GridView>
    </div>
    </form>
</body>
</html>
