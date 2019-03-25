Imports Word = Microsoft.Office.Interop.Word
'Using Microsoft.Office.Interop.Word;
'Imports System

'Imports System.Collections.Generic
'Imports System.IO
'Imports System.Linq
'Imports System.Text



'Option Strict Off
Public Class Form1
    Private SQL As New SQLControl
    Public AuthUser As String

    Public itemp As Integer
    Public box(15) As Integer
    Public boxpara(13, 3) As Integer
    Public Lbox(3) As Integer
    Public Lboxpara(15, 3) As Integer
    Public cmnH As Integer = 26
    Public coid As Integer = 100
    Public ICompany As Integer

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TEST
        Dim Mhigh As Integer = Screen.PrimaryScreen.Bounds.Height()
        Dim Mwidth As Integer = Screen.PrimaryScreen.Bounds.Width()
        Dim FVirt As Integer = (Mhigh / 100) * 90
        Dim FHoz As Integer = (Mwidth / 100) * 90
        Dim XF As Integer = (Mwidth - FHoz) / 2
        Dim YF As Integer = (Mhigh - FVirt) / 2

        SetBounds(XF, YF, FHoz, FVirt)
        'TEST

        GBLogged.Hide()
        BtnLogedN.Hide()
        'CallGbloged()
        XF = (FHoz / 2) - (340 / 2)
        YF = (FVirt / 2) - (124 / 2)

        TabConMaster.Hide()
        GBLogin.SetBounds(XF, YF, 340, 125)
        GBLogin.Show()



    End Sub

    Private Sub Getscrinfo(sender As Object, e As EventArgs) Handles Me.SizeChanged
        TabConSizeSet()
    End Sub

    Private Sub TabConSizeSet()
        Dim TH As Integer
        Dim Mhigh As Integer = Screen.PrimaryScreen.Bounds.Height()
        Dim Mwidth As Integer = Screen.PrimaryScreen.Bounds.Width()

        'If System.Windows.Forms.FormWindowState.Maximized Then
        If Me.WindowState = FormWindowState.Maximized Then
            TH = (Mhigh / 100) * 85
            ' TH = ((Mhigh / 4) * 3) - 90
            TabConMaster.SetBounds(10, 70, 1230, TH)
            TabConMaster.Show()
            TH = Mwidth - 100
            BtnLogedN.SetBounds(TH, 2, 60, 60)
            TH = Mwidth - 330 - 100
            GBLogged.SetBounds(TH, 2, 330, 107)
        Else
            'TH = (Mhigh / 20) * 17
            TH = (Mhigh / 100) * 70
            TabConMaster.SetBounds(10, 70, 1230, TH)
            TabConMaster.Show()
            TH = (Mwidth / 100) * 90 - 100
            BtnLogedN.SetBounds(TH, 2, 60, 60)
            TH = (Mwidth / 100) * 90 - 330 - 100
            GBLogged.SetBounds(TH, 2, 330, 107)
        End If
    End Sub

    Private Sub BtnLogin_Click(sender As Object, e As EventArgs) Handles BtnLogin.Click
        If SQL.HasConnection = True Then
            If IsAuthenticated() = True Then
                AuthUser = TxtUserName.Text
                'GetUserInfo()
                'MsgBox("Login Successfull")

                GBLogin.Hide()
                'GBLogged.SetBounds(900, 2, 60, 60)
                GBLogged.Hide()
                GetUserDetails(AuthUser)

                'REFERSH ALL GB
                GBFindEmpR.Hide()
                GBSelectCompany.Hide()


                GBAddEmp.Hide()
                GBAddApp.Hide()

                ''GBCoHierarchy.Hide()
                ''GBAddCoLevel.Hide()

                GBListApp.Hide()
                GBViewApp.Hide()
                GBEditGenInfo.Hide()
                GBEditDrvInfo.Hide()
                GBEditEduInfo.Hide()
                GBAddPreLi.Hide()
                GBAddPreEmp.Hide()
                GBAddPreExp.Hide()
                GBAddPreVio.Hide()
                GBAddPreAcc.Hide()
                GBAddPAddr.Hide()
                GBAddAppAtt.Hide()
                GBReportVApp.Hide()
                GBOtherVApp.Hide()


                ''btnAddEmp.Enabled = True
                ''btnAddApplicant.Enabled = True
                ''btnAppProcess.Enabled = True
                BtnFindEmpN.Enabled = False
                BtnFindEmpC.Enabled = False
                'SetBounds(50, 50, 1250, 700)
                AutoScroll = False

                'TAB MASTER SIZE
                'Dim Mhigh As Integer = Screen.PrimaryScreen.Bounds.Height()
                'Dim Mwidth As Integer = Screen.PrimaryScreen.Bounds.Width()
                'Dim TH As Integer = (((Mhigh / 4) * 3) / 20) * 17
                'Dim TH As Integer = ((Mhigh / 4) * 3) - 90
                'TabConMaster.SetBounds(10, 70, 1230, TH)
                TabConSizeSet()
                ''TabConMaster.Show()
                'Initial Layout(LEFT COL)
                'PicBoxIMFILogo.SetBounds(4, 40, 150, 50)
                PicBoxIMFILogo.Show()
                LBoxesView()
            End If

        End If
    End Sub




    Private Function IsAuthenticated() As Boolean
        'CLEAR EXISTING REC
        If SQL.DBDS IsNot Nothing Then
            SQL.DBDS.Clear()
        End If

        SQL.RunQuery("SELECT Count(UserName) As UserCount " &
                     "FROM MasterUser " &
                     "WHERE UserName='" & TxtUserName.Text & "' " &
                     "AND Passwd='" & TxtPassword.Text & "'COLLATE SQL_Latin1_General_CP1_CS_AS")

        If SQL.DBDS.Tables(0).Rows(0).Item("UserCount") = 1 Then
            Return True
        End If

        MsgBox("Invalid user credintial", MsgBoxStyle.Critical, "LOGIN FAILED!")
        Return False
    End Function

    Private Sub GetUserDetails(Username As String)
        Dim str1 As String
        Dim str2 As String
        SQL.AddParam("@user", Username)
        SQL.ExecQuery("SELECT TOP 1 * FROM MasterUser " &
                      "WHERE UserName = @user;")

        If SQL.RecordCount < 1 Then Exit Sub

        For Each r As DataRow In SQL.DBDT.Rows
            str1 = r("FUserName")
            str1 = Mid(str1, 1, 1)
            str2 = r("LUserName")
            str2 = Mid(str2, 1, 1)
            BtnLogedN.Text = str1.ToUpper + str2.ToUpper
            LblUName.Text = r("FUserName") & " " & r("LUserName")
            LblCompany.Text = r("Company")
            LblAccNo.Text = "Account Number  " & r("AccountID")
        Next
        '
        BtnLogedN.Show()

        'GBLogged.Show()
    End Sub




    Private Sub BtnLogedN_Click(sender As Object, e As EventArgs) Handles BtnLogedN.MouseHover
        itemp = 1
        CallGbloged()
    End Sub
    Private Sub TabHome_Click(sender As Object, e As EventArgs) Handles TabHome.MouseHover, TabEmp.MouseHover, TabQualification.MouseHover, TabDrug.MouseHover
        itemp = 0
        CallGbloged()
    End Sub
    Private Sub CallGbloged()
        If itemp > 0 Then
            '
            GBLogged.Show()
        Else
            GBLogged.Hide()
        End If
    End Sub

    Private Sub BtnLogout_Click(sender As Object, e As EventArgs) Handles BtnLogout.Click
        'AuthUser = txtUserName.Text
        'GetUserInfo()
        TabConMaster.Hide()
        GBLogged.Hide()
        BtnLogedN.Hide()
        'MsgBox("You are now completly Logout.")
        BtnLogedN.Text = ""
        LblUName.Text = ""
        LblCompany.Text = ""
        LblAccNo.Text = " "

        TxtUserName.Clear()
        TxtPassword.Clear()

        'GBLogin.SetBounds(455, 280, 340, 125)
        GBLogin.Show()
    End Sub

    '#########################################
    '  LEFT BOXES LAYOUT SHRINK AND FULL OPEN
    '#########################################
    Private Sub LBoxesView()

        'BOX INITIAL PARAM
        Lbox(0) = 5
        Lbox(1) = 5
        Lbox(2) = 5
        Lbox(3) = 1

        BtnOpenFE.Hide()
        BtnShrinkFE.Show()
        BtnOpenCT.Hide()
        BtnShrinkCT.Show()
        BtnOpenRA.Hide()
        BtnShrinkRA.Show()
        BtnOpenML.Show()
        BtnShrinkML.Hide()

        Lboxpara(0, 0) = 6
        Lboxpara(0, 1) = 20
        Lboxpara(0, 2) = 240
        Lboxpara(0, 3) = 114

        Lboxpara(1, 0) = 6
        Lboxpara(1, 1) = 52
        Lboxpara(1, 2) = 240
        Lboxpara(1, 3) = 110

        Lboxpara(2, 0) = 6
        Lboxpara(2, 1) = 82
        Lboxpara(2, 2) = 240
        Lboxpara(2, 3) = 240

        Lboxpara(3, 0) = 6
        Lboxpara(3, 1) = 112
        Lboxpara(3, 2) = 240
        Lboxpara(3, 3) = 162

        CloseOpen()

    End Sub
    Private Sub CloseOpen()
        Dim i As Integer = 0
        Dim j As Integer = 0

        Dim x As Integer
        Dim y As Integer
        Dim w As Integer
        Dim h As Integer

        Dim t As Integer = 0
        Dim ty As Integer = 0
        Dim th As Integer = 0

        For i = 0 To 3
            x = 0
            y = 0
            w = 0
            h = 0

            For j = 0 To 3

                t = Lboxpara(i, j)

                If j = 0 Then
                    x = t
                End If
                If j = 1 Then
                    y = t
                End If
                If j = 2 Then
                    w = t
                End If
                If j = 3 Then
                    h = t
                End If
            Next

            'BOX 1

            If i = 0 Then
                If Lbox(i) > 1 Then
                    ty = y
                    th = h
                    GBFindEmp.SetBounds(x, ty, w, th)

                Else
                    ty = y
                    th = cmnH
                    GBFindEmp.SetBounds(x, ty, w, th)
                End If
            End If

            'BOX 2
            If i = 1 Then
                y = ty + th + 2
                If Lbox(i) > 1 Then
                    ty = y
                    th = h
                    GBCommon.SetBounds(x, ty, w, th)
                Else
                    ty = y
                    th = cmnH
                    GBCommon.SetBounds(x, ty, w, th)
                End If
            End If

            'BOX 3
            If i = 2 Then
                y = ty + th + 2
                If Lbox(i) > 1 Then
                    ty = y
                    th = h
                    GBRelated.SetBounds(x, ty, w, th)
                Else
                    ty = y
                    th = cmnH
                    GBRelated.SetBounds(x, ty, w, th)
                End If
            End If

            'BOX 4
            If i = 3 Then
                y = ty + th + 2
                If Lbox(i) > 1 Then
                    ty = y
                    th = h
                    GBMaintLook.SetBounds(x, ty, w, th)
                Else
                    ty = y
                    th = cmnH
                    GBMaintLook.SetBounds(x, ty, w, th)
                End If
            End If

        Next
        GBFindEmp.Show()
        GBCommon.Show()
        GBRelated.Show()
        GBMaintLook.Show()

    End Sub
    Private Sub BtnOpenFE_Click(sender As Object, e As EventArgs) Handles BtnOpenFE.Click
        Lbox(0) = 5
        BtnOpenFE.Hide()
        BtnShrinkFE.Show()
        CloseOpen()
    End Sub
    Private Sub BtnShrinkFE_Click(sender As Object, e As EventArgs) Handles BtnShrinkFE.Click
        Lbox(0) = 1
        BtnOpenFE.Show()
        BtnShrinkFE.Hide()
        CloseOpen()
    End Sub

    Private Sub BtnOpenCT_Click(sender As Object, e As EventArgs) Handles BtnOpenCT.Click
        Lbox(1) = 5
        BtnOpenCT.Hide()
        BtnShrinkCT.Show()
        CloseOpen()
    End Sub
    Private Sub BtnShrinkCT_Click(sender As Object, e As EventArgs) Handles BtnShrinkCT.Click
        Lbox(1) = 1
        BtnOpenCT.Show()
        BtnShrinkCT.Hide()
        CloseOpen()
    End Sub

    Private Sub BtnOpenRA_Click(sender As Object, e As EventArgs) Handles BtnOpenRA.Click
        Lbox(2) = 5
        BtnOpenRA.Hide()
        BtnShrinkRA.Show()
        CloseOpen()
    End Sub
    Private Sub BtnShrinkRA_Click(sender As Object, e As EventArgs) Handles BtnShrinkRA.Click
        Lbox(2) = 1
        BtnOpenRA.Show()
        BtnShrinkRA.Hide()
        CloseOpen()
    End Sub

    Private Sub BtnOpenML_Click(sender As Object, e As EventArgs) Handles BtnOpenML.Click
        Lbox(3) = 5
        BtnOpenML.Hide()
        BtnShrinkML.Show()
        CloseOpen()
    End Sub
    Private Sub BtnShrinkML_Click(sender As Object, e As EventArgs) Handles BtnShrinkML.Click
        Lbox(3) = 1
        BtnOpenML.Show()
        BtnShrinkML.Hide()
        CloseOpen()
    End Sub

    '#############################
    '   FIND EMPLOYEE RESULT
    '##############################
    Private Sub TxtFindEmp_TextChanged(sender As Object, e As EventArgs) Handles TxtFindEmp.TextChanged
        'BASIC VALIDATION
        If Not String.IsNullOrWhiteSpace(TxtFindEmp.Text) Then
            BtnFindEmpN.Enabled = True
            BtnFindEmpC.Enabled = True
        Else
            BtnFindEmpN.Enabled = False
            BtnFindEmpC.Enabled = False
        End If
    End Sub

    Public ichk As Integer

    Private Sub BtnFindEmpN_Click(sender As Object, e As EventArgs) Handles BtnFindEmpN.Click
        ichk = 0
        GBFindEmpR.SetBounds(260, 20, 700, 650)
        GBFindEmpR.Show()
        FetchFindEResult()
    End Sub

    Private Sub BtnFindEmpC_Click(sender As Object, e As EventArgs) Handles BtnFindEmpC.Click
        ichk = 1
        GBFindEmpR.SetBounds(260, 20, 700, 650)
        GBFindEmpR.Show()
        FetchFindEResult()
    End Sub
    Private Sub FetchFindEResult()
        Dim tempstr1 As String = ""
        Dim tempstr2 As String = ""
        Dim tempstr3 As String = ""
        Dim sumstr As String = ""

        CLBFindEmpResult.Items.Clear()

        If SQL.DBDT IsNot Nothing Then
            SQL.DBDT.Clear()
        End If

        If ichk > 0 Then
            'tempstr1 = CInt(TxtFindEmp.Text)
            SQL.AddParam("@TFindEmp", TxtFindEmp.Text)
            SQL.ExecQuery("SELECT Employees.EmpFirstName,Employees.EmpLastName,Employees.EmpCode,Company.CompanyName " &
                        "FROM Employees " &
                        "INNER JOIN Company ON Employees.CompanyID=Company.CompanyID " &
                        "WHERE EmpCode=@TFindEmp  AND ApplicationID=0 " &
                        "ORDER BY EmpFirstName ASC;")
        Else
            tempstr1 = TxtFindEmp.Text
            'SQL.AddParam("@TFindEmp", tempstr1)
            SQL.ExecQuery("SELECT Employees.EmpFirstName,Employees.EmpLastName,Employees.EmpCode,Company.CompanyName " &
                        "FROM Employees " &
                        "INNER JOIN Company ON Employees.CompanyID=Company.CompanyID " &
                        "WHERE (EmpFirstName LIKE '%" & tempstr1 & "%' AND ApplicationID=0) OR (EmpLastName LIKE '%" & tempstr1 & "%' AND ApplicationID=0) " &
                        "ORDER BY EmpFirstName ASC;")
        End If
        '
        If SQL.RecordCount < 1 Then Exit Sub

        For Each r As DataRow In SQL.DBDT.Rows

            tempstr1 = Trim(r("EmpFirstName"))
            tempstr2 = Trim(r("EmpLastName"))
            tempstr3 = Trim(r("CompanyName"))
            sumstr = tempstr1 & "," & tempstr2 & vbTab & vbTab & vbTab & r("EmpCode") & vbTab & vbTab & vbTab & tempstr3

            CLBFindEmpResult.Items.Add(sumstr)
        Next
    End Sub

    Private Sub BtnCancelFER_Click(sender As Object, e As EventArgs) Handles BtnCancelFER.Click
        GBFindEmpR.Hide()
        TxtFindEmp.Clear()
        BtnFindEmpC.Enabled = False
        BtnFindEmpN.Enabled = False
    End Sub

    '#############################
    '   SELECT COMPANY
    '##############################
    Public CoIDName As String
    Private Sub SelCompanyView()
        GBSelectCompany.SetBounds(260, 20, 550, 200)
        BtnContSCU.Enabled = False
        BtnContSCL.Enabled = False
        GBSelectCompany.Show()
        TxtCompanyID.Hide()

        Dim tempstr1 As String = ""

        If SQL.DBDT IsNot Nothing Then
            SQL.DBDT.Clear()
        End If

        'SQL.AddParam("@TFindEmp", TxtFindEmp.Text)
        SQL.ExecQuery("SELECT CompanyName " &
                      "FROM Company " &
                      "ORDER BY CompanyName ASC;")
        '
        If SQL.RecordCount < 1 Then Exit Sub
        Dim i As Integer = 0
        For Each r As DataRow In SQL.DBDT.Rows
            i += 1
            tempstr1 = Trim(r("CompanyName"))
            Select Case i > 0
                Case i = 1
                    RBCompany1.Text = tempstr1
                Case i = 2
                    RBCompany2.Text = tempstr1
                Case i = 3
                    RBCompany3.Text = tempstr1
                Case i = 4
                    RBCompany4.Text = tempstr1
            End Select
        Next
    End Sub

    Private Sub BtnContSCU_Click(sender As Object, e As EventArgs) Handles BtnContSCU.Click, BtnContSCL.Click
        'coid = TxtCompanyID.Text
        'ClearGBSelCo()
        GBSelectCompany.Hide()
        Select Case ICompany > 0
            Case ICompany = 1
                ' CALL ADD NEW EMPLOYEE
                TxtEmpCode.Enabled = False
                BtnSaveAE.Enabled = False
                BtnSaveAddAE.Enabled = False
                GBAddEmp.SetBounds(260, 20, 700, 740)
                GBAddEmp.Show()
                TxtCoidAE.Text = coid
                TxtSCoidAE.Text = coid


                FetchCombBoxAE()
                FetchJobClass()
                FetchStatus()
                ' MsgBox("Add An EMPLOYEE")
                TxtCoidAE.Hide()
                TxtSCoidAE.Hide()
                TxtSSCoidAE.Hide()
            Case ICompany = 2
                ' CALL ADD NEW APPLICANT
                LblTitleApp.Text = "Add New Applicant"
                BtnUpdANA.Hide()
                BtnUpdCancelANA.Hide()
                BtnUpdEdGenInfo.Hide()

                TxtCoidApp.Text = coid
                TxtWHereFm.Enabled = False
                TxtWHereTo.Enabled = False
                DTPWHFm.Enabled = False
                DTPWHTo.Enabled = False
                BtnSaveANA.Enabled = False
                BtnSaveAddANA.Enabled = False
                BtnSaveANA.Show()
                BtnSaveAddANA.Show()
                ClearAddApp()
                GBAddApp.SetBounds(260, 20, 700, 950)
                GBAddApp.Show()
                'MsgBox("Add An APPLICANT")
                TxtEmpIDApp.Hide()
                TxtCoidApp.Hide()
            Case ICompany = 3
                ' CALL COMPANY HIERARCHY VIEW
                ''GBCoHierarchy.SetBounds(260, 20, 670, 350)
                ''GBCoHierarchy.Show()
                ''TxtCoidCHI.Text = coid
                ''TreeCompanyView()
                ''TxtNodeCount.Clear()
                '' TxtNodeCount.Text = TreeVComp1.TopNode.Nodes.Count()
        End Select
    End Sub

    Private Sub BtnCancelSCU_Click(sender As Object, e As EventArgs) Handles BtnCancelSCU.Click, BtnCancelSCL.Click
        coid = 0
        CoIDName = ""
        ClearGBSelCo()
        GBSelectCompany.Hide()
    End Sub

    Private Sub ClearGBSelCo()
        RBCompany1.Checked = False
        RBCompany2.Checked = False
        RBCompany3.Checked = False
        RBCompany4.Checked = False
    End Sub

    Private Sub GBSelectCompany_Enter(sender As Object, e As EventArgs) Handles RBCompany1.CheckedChanged, RBCompany2.CheckedChanged, RBCompany3.CheckedChanged, RBCompany4.CheckedChanged
        CheckedCompany()
        'TxtCompanyID.Text = coid
    End Sub

    Private Sub CheckedCompany()
        Dim tempstr1 As String = ""

        If SQL.DBDT IsNot Nothing Then
            SQL.DBDT.Clear()
        End If
        If RBCompany1.Checked Then
            tempstr1 = RBCompany1.Text
        ElseIf RBCompany2.Checked Then
            tempstr1 = RBCompany2.Text
        ElseIf RBCompany3.Checked Then
            tempstr1 = RBCompany3.Text
        ElseIf RBCompany4.Checked Then
            tempstr1 = RBCompany4.Text
        End If

        SQL.AddParam("@TCompanyN", tempstr1)
        SQL.ExecQuery("SELECT CompanyID " &
                      "FROM Company " &
                      "WHERE CompanyName=@TCompanyN;")
        '
        If SQL.RecordCount < 1 Then Exit Sub

        For Each r As DataRow In SQL.DBDT.Rows
            coid = r("CompanyID")
        Next
        ' CHECKED CO ID AND COMPANY
        TxtCompanyID.Text = coid
        CoIDName = tempstr1
        BtnContSCU.Enabled = True
        BtnContSCL.Enabled = True
    End Sub

    '#############################
    '   ADD AN EMPLOYEE
    '##############################
    Private Sub BtnAddEmp_Click(sender As Object, e As EventArgs) Handles BtnAddEmp.Click
        ''GBListApp.Hide()
        SelCompanyView()
        ICompany = 1
        ' CHECK SELECT COPMANY
    End Sub

    Private Sub BtnCancelAE_Click(sender As Object, e As EventArgs) Handles BtnCancelAE.Click
        GBClearAddEmp()
        CombBoxRepL.ResetText()
        GBAddEmp.Hide()
        ClearGBSelCo()
    End Sub

    Private Sub BtnSaveAE_Click(sender As Object, e As EventArgs) Handles BtnSaveAE.Click
        InsertAddEmp()
        GBClearAddEmp()
        CombBoxRepL.ResetText()
        GBAddEmp.Hide()
        ClearGBSelCo()
    End Sub
    Private Sub BtnSaveAddAE_Click(sender As Object, e As EventArgs) Handles BtnSaveAddAE.Click
        InsertAddEmp()
        GBClearAddEmp()
    End Sub
    Private Sub InsertAddEmp()
        Dim str1 As String

        SQL.AddParam("@AppID", False)
        SQL.AddParam("@FName", TxtFname.Text)
        SQL.AddParam("@MName", TxtMI.Text)
        SQL.AddParam("@LName", TxtLname.Text)
        SQL.AddParam("@Coid", TxtCoidAE.Text)
        If TxtSCoidAE.Text > 49 Then
            SQL.AddParam("@CoReport", TxtSCoidAE.Text)
        Else
            SQL.AddParam("@CoReport", TxtCoidAE.Text)
        End If
        SQL.AddParam("@JClass", CombBoxJClass.Text)
        SQL.AddParam("@JStatus", CombBoxStatus.Text)
        SQL.AddParam("@DtHire", TxtDtHire.Text)
        SQL.AddParam("@DtInact", TxtInADt.Text)
        SQL.AddParam("@DtSenior", TxtSenDt.Text)
        str1 = TxtSSN.Text
        SQL.AddParam("@FSsn", Mid(str1, 1, 3))
        SQL.AddParam("@MSsn", Mid(str1, 5, 2))
        SQL.AddParam("@LSsn", Mid(str1, 8, 4))
        SQL.AddParam("@EmpAddr", TxtEmpAddr.Text)
        SQL.AddParam("@EmpApt", TxtEmpApt.Text)
        SQL.AddParam("@EmpCity", TxtEmpCity.Text)
        SQL.AddParam("@EmpSt", CombBoxEmpState.Text)
        SQL.AddParam("@EmpFZip", TxtZip.Text)
        SQL.AddParam("@EmpDob", TxtEmpDOB.Text)
        SQL.AddParam("@EmpPhone", TxtEPhone.Text)
        SQL.AddParam("@EmpEmail", TxtEEmail.Text)

        SQL.ExecQuery("INSERT INTO Employees " &
                        "(ApplicationID,EmpFirstName,EmpMI,EmpLastName,CompanyID,ReportLevel,JobClass, " &
                        "EmpStatus,DateHire,DateInActive,DateSeniority,Fssn,Mssn,Lssn,EmpAddress,EmpApt, " &
                        "EmpCity,EmpState,EmpFirstZip,EmpDOB,EmpPhone,EmpEmail) " &
                        "VALUES " &
                        "(@AppID,@FName,@MName,@LName,@Coid,@CoReport,@JClass,@JStatus,@DtHire,@DtInact, " &
                        "@DtSenior,@FSsn,@MSsn,@LSsn,@EmpAddr,@EmpApt,@EmpCity,@EmpSt,@EmpFZip,@EmpDob, " &
                        "@EmpPhone,@EmpEmail); " &
                        "BEGIN TRANSACTION; " &
                        "COMMIT;", True)

        'ERROR
        If SQL.HasException(True) Then Exit Sub

        MsgBox("New Employee added Successfully")
    End Sub
    Private Sub FetchCombBoxAE()
        'Dim Tempstr1 As String

        CombBoxRepL.Items.Clear()
        If SQL.DBDT IsNot Nothing Then
            SQL.DBDT.Clear()
        End If

        SQL.AddParam("@TCompanyID", coid)
        SQL.ExecQuery("SELECT CompanyName " &
                      "FROM Company " &
                      "WHERE CompanyID=@TCompanyID;")
        '
        If SQL.RecordCount < 1 Then Exit Sub

        For Each r As DataRow In SQL.DBDT.Rows
            CombBoxRepL.Items.Add(r("CompanyName"))
            CombBoxRepL.Text = r("CompanyName")
        Next
        'CombBoxRepL.Items.Add(CoIDName)
        If SQL.DBDT IsNot Nothing Then
            SQL.DBDT.Clear()
        End If

        SQL.AddParam("@TCompanyID", coid)
        SQL.ExecQuery("SELECT SubCoName " &
                      "FROM SubCompany1 " &
                      "WHERE CompanyID=@TCompanyID;")
        '
        If SQL.RecordCount < 1 Then Exit Sub

        For Each r As DataRow In SQL.DBDT.Rows
            CombBoxRepL.Items.Add("--" & r("SubCoName"))
            'CombBoxRepL.Text = r("CompanyName")
        Next

    End Sub
    Private Sub FetchJobClass()
        ' REFRESS JOBCLASS ITEMS
        CombBoxJClass.Items.Clear()

        If SQL.DBDT IsNot Nothing Then
            SQL.DBDT.Clear()
        End If
        SQL.ExecQuery("SELECT DISTINCT JobClass FROM Employees; ")

        If SQL.RecordCount < 1 Then Exit Sub
        If SQL.HasException(True) Then Exit Sub
        'LOOP
        For Each r As DataRow In SQL.DBDT.Rows
            CombBoxJClass.Items.Add(r("JobClass"))
        Next
    End Sub
    Private Sub FetchStatus()
        ' REFRESS STATUS ITEMS

        CombBoxStatus.Items.Clear()

        If SQL.DBDT IsNot Nothing Then
            SQL.DBDT.Clear()
        End If
        SQL.ExecQuery("SELECT DISTINCT EmpStatus FROM Employees; ")

        If SQL.RecordCount < 1 Then Exit Sub
        If SQL.HasException(True) Then Exit Sub
        'LOOP
        For Each r As DataRow In SQL.DBDT.Rows

            CombBoxStatus.Items.Add(r("EmpStatus"))
        Next
    End Sub
    'SELECT SUB COMPANY
    Private Sub CombBoxRepL_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CombBoxRepL.SelectedIndexChanged, CombBoxRepL.Click

        SelectedCOID()
    End Sub
    Private Sub SelectedCOID()
        Dim str1 As String
        Dim loc1 As Integer
        str1 = CombBoxRepL.Text
        loc1 = InStr(str1, "--")

        str1 = Mid(str1, 3)
        'TxtSCoidAE.Text = str1
        str1 = Trim(str1)
        If loc1 > 0 Then
            TxtSCoidAE.Clear()
            If SQL.DBDT IsNot Nothing Then
                SQL.DBDT.Clear()
            End If

            SQL.AddParam("@TCompanyN", str1)
            SQL.ExecQuery("SELECT SubCoID1 " &
                          "FROM SubCompany1 " &
                          "WHERE SubCoName=@TCompanyN;")
            '
            If SQL.RecordCount < 1 Then Exit Sub

            For Each r As DataRow In SQL.DBDT.Rows
                TxtSCoidAE.Text = r("SubCoID1")
            Next
        End If
    End Sub

    ' RIGHT TRIM DATE
    Private Sub DTPDtHire_ValueChanged(sender As Object, e As EventArgs) Handles DTPDtHire.ValueChanged
        TxtDtHire.Text = DateTrimR(DTPDtHire.Value)

    End Sub
    Private Sub DTPInADt_ValueChanged(sender As Object, e As EventArgs) Handles DTPInADt.ValueChanged
        TxtInADt.Text = DateTrimR(DTPInADt.Value)
    End Sub
    Private Sub DTPSenDt_ValueChanged(sender As Object, e As EventArgs) Handles DTPSenDt.ValueChanged
        TxtSenDt.Text = DateTrimR(DTPSenDt.Value)
    End Sub
    Private Sub DTPEmpDOB_ValueChanged(sender As Object, e As EventArgs) Handles DTPEmpDOB.ValueChanged
        TxtEmpDOB.Text = DateTrimR(DTPEmpDOB.Value)
    End Sub

    Private Function DateTrimR(trimdt As String)
        Dim loc As Integer
        Dim tmpstr1 As String
        tmpstr1 = trimdt
        loc = Len(tmpstr1)
        If loc < 11 Then

        Else
            loc = InStr(tmpstr1, " ")
            tmpstr1 = Mid(tmpstr1, 1, loc)
        End If

        Return tmpstr1
    End Function
    ' VALIDATION OF SSN
    Private Sub TxtSSN_TextChanged(sender As Object, e As EventArgs) Handles TxtSSN.TextChanged

        TxtSSN.Text = SSNFormat(TxtSSN.Text)

    End Sub
    Private Function SSNFormat(txtssn As String)
        Dim tempstr1 As String
        Dim tempstr2 As String
        Dim tempstr3 As String
        Dim l1 As Integer
        tempstr3 = txtssn
        l1 = Len(tempstr3)
        If l1 = 9 Then

            tempstr1 = Mid(tempstr3, 1, 3)
            tempstr2 = Mid(tempstr3, 4, 2)
            tempstr3 = Mid(tempstr3, 6, l1)
            txtssn = tempstr1 + "-" + tempstr2 + "-" + tempstr3
        Else
            txtssn = tempstr3
        End If
        Return txtssn
    End Function
    ' CLEAR ALL FIELD AE GB
    Private Sub GBClearAddEmp()
        TxtFname.Clear()
        TxtMI.Clear()
        TxtLname.Clear()

        TxtEmpCode.Clear()
        TxtDtHire.Clear()
        TxtInADt.Clear()
        TxtSenDt.Clear()
        TxtEmpAddr.Clear()
        TxtEmpCity.Clear()
        TxtEmpDOB.Clear()

        TxtEEmail.Clear()
        TxtEmpApt.Clear()
        TxtZip.Clear()
        TxtEPhone.Clear()
        TxtSSN.Clear()

        'CombBoxRepL.ResetText()
        CombBoxJClass.ResetText()
        CombBoxStatus.ResetText()
        CombBoxEmpState.ResetText()

    End Sub

    Private Sub TxtFname_TextChanged(sender As Object, e As EventArgs) Handles TxtFname.TextChanged, TxtLname.TextChanged
        'BASIC VALIDATION
        If Not String.IsNullOrWhiteSpace(TxtFname.Text) AndAlso Not String.IsNullOrWhiteSpace(TxtLname.Text) Then
            BtnSaveAE.Enabled = True
            BtnSaveAddAE.Enabled = True
        Else
            BtnSaveAE.Enabled = False
            BtnSaveAddAE.Enabled = False
        End If
    End Sub

    '#############################
    '   ADD AN APPLICANT
    '##############################
    Private Sub BtnAddApplicant_Click(sender As Object, e As EventArgs) Handles BtnAddApplicant.Click
        'GBListApp.Hide()
        SelCompanyView()
        ICompany = 2
        ' CHECK SELECT COPMANY
    End Sub

    Private Sub BtnCancelANA_Click(sender As Object, e As EventArgs) Handles BtnCancelANA.Click
        ClearAddApp()
        GBAddApp.Hide()
        ClearGBSelCo()
        BtnAppProcess.Enabled = True
    End Sub

    Private Sub BtnSaveANA_Click(sender As Object, e As EventArgs) Handles BtnSaveANA.Click
        InsertAddApp()
        ClearAddApp()
        GBAddApp.Hide()
        ClearGBSelCo()
    End Sub
    Private Sub BtnSaveAddANA_Click(sender As Object, e As EventArgs) Handles BtnSaveAddANA.Click
        InsertAddApp()
        ClearAddApp()
    End Sub
    Private Sub DTPAppDt_ValueChanged(sender As Object, e As EventArgs) Handles DTPAppDt.ValueChanged
        TxtAppDt.Text = DateTrimR(DTPAppDt.Value)
    End Sub
    Private Sub DTPDOBAA_ValueChanged(sender As Object, e As EventArgs) Handles DTPDOBAA.ValueChanged
        TxtDOBAA.Text = DateTrimR(DTPDOBAA.Value)
    End Sub
    Private Sub DTPWHFm_ValueChanged(sender As Object, e As EventArgs) Handles DTPWHFm.ValueChanged
        TxtWHereFm.Text = DateTrimR(DTPWHFm.Value)
    End Sub
    Private Sub DTPWHTo_ValueChanged(sender As Object, e As EventArgs) Handles DTPWHTo.ValueChanged
        TxtWHereTo.Text = DateTrimR(DTPWHTo.Value)
    End Sub
    Private Sub DTPLastEmp_ValueChanged(sender As Object, e As EventArgs) Handles DTPLastEmp.ValueChanged
        TxtLastEmp.Text = DateTrimR(DTPLastEmp.Value)
    End Sub
    Private Sub TxtSSNAA_TextChanged(sender As Object, e As EventArgs) Handles TxtSSNAA.TextChanged
        TxtSSNAA.Text = SSNFormat(TxtSSNAA.Text)
    End Sub

    Private Sub InsertAddApp()
        Dim str1 As String
        SQL.AddParam("@AppID", True)
        SQL.AddParam("@FName", TxtFNameAA.Text)
        SQL.AddParam("@MName", TxtMIAA.Text)
        SQL.AddParam("@LName", TxtLnameAA.Text)
        SQL.AddParam("@Coid", TxtCoidApp.Text)
        SQL.AddParam("@AppDt", TxtAppDt.Text)
        str1 = TxtSSNAA.Text
        SQL.AddParam("@FSsn", Mid(str1, 1, 3))
        SQL.AddParam("@MSsn", Mid(str1, 5, 2))
        SQL.AddParam("@LSsn", Mid(str1, 8, 4))
        SQL.AddParam("@EmpAddr", TxtAddrAA.Text)
        SQL.AddParam("@EmpApt", TxtAptAA.Text)
        SQL.AddParam("@EmpCity", TxtCityAA.Text)
        SQL.AddParam("@EmpSt", CombBoxStateAA.Text)
        SQL.AddParam("@EmpFZip", TxtZipAA.Text)
        SQL.AddParam("@LivYY", TxtLivAddYY.Text)
        SQL.AddParam("@LivMM", TxtLivAddMM.Text)
        SQL.AddParam("@DOB", TxtDOBAA.Text)
        SQL.AddParam("@HPhone", TxtHPhoneAA.Text)
        SQL.AddParam("@WPhone", TxtWPhoneAA.Text)
        SQL.AddParam("@EmpEmail", TxtEmailAA.Text)
        SQL.AddParam("@AppPos", TxtAppPosition.Text)
        SQL.AddParam("@LegalR", TxtLegalRAA.Text)
        SQL.AddParam("@ProofA", CBProofAge.Checked)
        SQL.AddParam("@WorkB", CBWorkBefore.Checked)
        SQL.AddParam("@WorkBWH", TxtWBeforeWH.Text)
        SQL.AddParam("@WorkHFrom", TxtWHereFm.Text)
        SQL.AddParam("@WorkHTo", TxtWHereTo.Text)
        SQL.AddParam("@Pay", TxtPay.Text)
        SQL.AddParam("@Position", TxtPosition.Text)
        SQL.AddParam("@ReasonL", TxtReason.Text)
        SQL.AddParam("@Employed", CBEmployed.Checked)
        SQL.AddParam("@LastEmployed", TxtLastEmp.Text)
        SQL.AddParam("@Refered", TxtRefBy.Text)
        SQL.AddParam("@ExpPay", TxtExptedPay.Text)

        SQL.ExecQuery("INSERT INTO Employees " &
                        "(ApplicationID,EmpFirstName,EmpMI,EmpLastName,CompanyID,AppDate, " &
                        "Fssn,Mssn,Lssn,EmpAddress,EmpApt,EmpCity,EmpState,EmpFirstZip,LivedAddYY,LivedAddMM, " &
                        "EmpDOB,HomePhone,WorkPhone,EmpEmail,AppPosition,LegalRTW,ProofAge,WorkBefore,WorkBeforeWh, " &
                        "WorkHereFM,WorkHereTo,Pay,Position,ReasonLeave,Employed,LastEmp,ReferedBy,ExpectedPay) " &
                        "VALUES " &
                        "(@AppID,@FName,@MName,@LName,@Coid,@AppDt,@FSsn,@MSsn,@LSsn, " &
                        "@EmpAddr,@EmpApt,@EmpCity,@EmpSt,@EmpFZip,@LivYY,@LivMM,@DOB,@HPhone, " &
                        "@WPhone,@EmpEmail,@AppPos,@LegalR,@ProofA,@WorkB,@WorkBWH,@WorkHFrom,@WorkHTo, " &
                        "@Pay,@Position,@ReasonL,@Employed,@LastEmployed,@Refered,@ExpPay); " &
                        "BEGIN TRANSACTION; " &
                        "COMMIT;", True)

        'ERROR
        If SQL.HasException(True) Then Exit Sub

        MsgBox("New Applicant added Successfully")
    End Sub

    Private Sub CBWorkBefore_CheckedChanged(sender As Object, e As EventArgs) Handles CBWorkBefore.CheckedChanged
        If CBWorkBefore.Checked Then
            TxtWHereFm.Enabled = True
            TxtWHereTo.Enabled = True
            DTPWHFm.Enabled = True
            DTPWHTo.Enabled = True
            TxtPay.Enabled = True
            TxtPosition.Enabled = True
            TxtReason.Enabled = True
        Else
            TxtWHereFm.Enabled = False
            TxtWHereTo.Enabled = False
            DTPWHFm.Enabled = False
            DTPWHTo.Enabled = False
            TxtPay.Enabled = False
            TxtPosition.Enabled = False
            TxtReason.Enabled = False
        End If
    End Sub
    Private Sub CBEmployed_CheckedChanged(sender As Object, e As EventArgs) Handles CBEmployed.CheckedChanged
        If CBEmployed.Checked Then

            TxtLastEmp.Enabled = False

            DTPLastEmp.Enabled = False
            DTPLastEmp.ResetText()
            TxtLastEmp.Clear()
        Else
            TxtLastEmp.Enabled = True
            DTPLastEmp.Enabled = True
        End If
    End Sub
    Private Sub TxtFNameAA_TextChanged(sender As Object, e As EventArgs) Handles TxtAppDt.TextChanged, TxtFNameAA.TextChanged, TxtLnameAA.TextChanged
        If Not String.IsNullOrWhiteSpace(TxtFNameAA.Text) AndAlso Not String.IsNullOrWhiteSpace(TxtLnameAA.Text) AndAlso Not String.IsNullOrWhiteSpace(TxtAppDt.Text) Then

            BtnSaveANA.Enabled = True
            BtnSaveAddANA.Enabled = True
        Else
            BtnSaveANA.Enabled = False
            BtnSaveAddANA.Enabled = False
        End If
    End Sub
    Private Sub ClearAddApp()

        TxtFNameAA.Clear()
        TxtMIAA.Clear()
        TxtLnameAA.Clear()
        'TxtCoidApp.Clear()
        TxtAppDt.Clear()
        TxtSSNAA.Clear()
        TxtAddrAA.Clear()
        TxtAptAA.Clear()
        TxtCityAA.Clear()
        CombBoxStateAA.ResetText()
        TxtZipAA.Clear()
        TxtLivAddYY.Clear()
        TxtLivAddMM.Clear()
        TxtDOBAA.Clear()
        TxtHPhoneAA.Clear()
        TxtWPhoneAA.Clear()
        TxtEmailAA.Clear()
        TxtAppPosition.Clear()
        TxtLegalRAA.Clear()
        CBProofAge.Checked = False
        CBWorkBefore.Checked = False
        TxtWBeforeWH.Clear()
        TxtWHereFm.Clear()
        TxtWHereTo.Clear()
        TxtPay.Clear()
        TxtPosition.Clear()

        TxtReason.Clear()
        CBEmployed.Checked = False
        TxtLastEmp.Clear()
        TxtRefBy.Clear()
        TxtExptedPay.Clear()

    End Sub

    ''#############################################################
    '   APPLICATION PROCESSING AND LIST OF CURRENT APPLICATION
    '##############################################################
    Private Sub BtnAppProcess_Click(sender As Object, e As EventArgs) Handles BtnAppProcess.Click
        'OPEN THE SCOPE OF ADD NEW Applicant

        GBAddEmp.Hide()
        GBAddApp.Hide()

        'HIDE LEFT
        LboxHide()

        'RELOCATE GB Add Applicant Info & SHOW
        GBListApp.SetBounds(260, 20, 700, 650)
        GBListApp.Show()

        FetchAppProcess()
        'AllAppFieldClear()
        'BASIC VALIDATION
        BtnAddEmp.Enabled = True
        BtnAddApplicant.Enabled = True
        BtnAppProcess.Enabled = False
        BtnDeleteApp.Enabled = False
        BtnEditApp.Enabled = False
        BtnViewApp.Enabled = False
    End Sub
    Private Sub LboxHide()
        GBFindEmp.Hide()
        GBCommon.Hide()
        GBRelated.Hide()
        GBMaintLook.Hide()
    End Sub
    Private Sub FetchAppProcess()
        ' REFRESS JOBCLASS ITEMS
        CLBApplicant.Items.Clear()

        If SQL.DBDT IsNot Nothing Then
            SQL.DBDT.Clear()
        End If
        'SQL.AddParam("@Jclass", txtEmpReportL.Text)
        SQL.ExecQuery("SELECT EmpFirstName,EmpLastName,AppDate,Lssn FROM Employees " &
                     "WHERE ApplicationID=1 " &
                     "ORDER BY EmpFirstName ASC;")

        ' If SQL.RecordCount < 1 Then Exit Sub
        If SQL.HasException(True) Then Exit Sub
        'LOOP
        Dim countltr As Integer = 0
        'Dim blank1st As Integer = 0
        Dim str1 As String = ""
        Dim str2 As String = ""
        Dim str3 As String = ""
        Dim str4 As String = ""
        Dim sumstr As String = ""
        For Each r As DataRow In SQL.DBDT.Rows
            str1 = Trim(r("EmpFirstName"))
            str2 = Trim(r("EmpLastName"))
            'blank1st = InStr(r("AppDate"), " ")
            str3 = DateTrimR(r("AppDate"))
            'str3 = Mid(r("AppDate"), 1, blank1st)
            str3 = Replace$(str3, "/", "-")
            str4 = Trim(r("Lssn"))
            countltr = str1.Length + str2.Length
            If countltr <= 13 Then
                sumstr = str1 & "," & str2 & vbTab & vbTab & vbTab & str3 & vbTab & vbTab & vbTab & "XXX-XX-" & str4
            Else
                sumstr = str1 & "," & str2 & vbTab & vbTab & str3 & vbTab & vbTab & vbTab & "XXX-XX-" & str4
            End If

            CLBApplicant.Items.Add(sumstr)

        Next
    End Sub

    Private Sub BtnCancelApp_Click(sender As Object, e As EventArgs) Handles BtnCancelApp.Click
        GBListApp.Hide()
        'SHOW LEFT
        'GBFindEmp.Show()
        'GBCommon.Show()
        'GBRelated.Show()
        'GBMaintLook.Show()
        LBoxesView()
        BtnAppProcess.Enabled = True
    End Sub

    Private Sub CLBApplicant_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CLBApplicant.SelectedIndexChanged, CLBApplicant.Click, CLBApplicant.MouseDoubleClick
        If CLBApplicant.CheckedItems.Count() = 0 Then
            BtnDeleteApp.Enabled = False
            BtnEditApp.Enabled = False
            BtnViewApp.Enabled = False
        Else
            BtnDeleteApp.Enabled = True
            BtnEditApp.Enabled = True
            BtnViewApp.Enabled = True
        End If
    End Sub

    Private Sub BtnDeleteApp_Click(sender As Object, e As EventArgs) Handles BtnDeleteApp.Click
        Dim tempstr As String = ""
        Dim DelString As String = ""
        Dim commaLoc As Integer
        Dim c As Integer

        For Each i As String In CLBApplicant.CheckedItems
            tempstr = ""
            commaLoc = InStr(i, ",")
            tempstr = Mid(i, 1, commaLoc - 1)
            SQL.AddParam("@Fname" & c, tempstr)
            DelString += "DELETE FROM Employees " &
                    "WHERE ApplicationID=1 AND " &
                    "EmpFirstName = @Fname" & c & ";"
            c += 1
        Next
        'EXECUTE BULK DELETE CMD
        SQL.ExecQuery(DelString)

        If SQL.HasException(True) Then Exit Sub

        MsgBox("Selected User has been deleted")
        FetchAppProcess()

    End Sub

    '###########*************############

    Private Sub BtnEditApp_Click(sender As Object, e As EventArgs) Handles BtnEditApp.Click
        'OPEN THE SCOPE OF ADD NEW Applicant

        ClearAddApp()
        GBAddEmp.Hide()
        GBListApp.Hide()
        'RELOCATE GB Add Applicant Info & SHOW
        LblTitleApp.Text = "Edit Applicant"
        GBAddApp.SetBounds(260, 20, 700, 950)

        GBAddApp.Show()
        BtnAddEmp.Enabled = True
        BtnAppProcess.Enabled = False
        BtnAddApplicant.Enabled = True
        'BASIC VALIDATION

        BtnSaveANA.Hide()
        BtnSaveAddANA.Hide()
        BtnCancelANA.Hide()
        BtnUpdANA.SetBounds(150, 910, 81, 24)
        BtnUpdCancelANA.SetBounds(235, 910, 81, 24)
        BtnUpdEdGenInfo.SetBounds(325, 910, 165, 24)
        BtnUpdANA.Show()
        BtnUpdCancelANA.Show()
        BtnUpdEdGenInfo.Show()

        'TxtEmpIDApp.Hide()
        'TxtCoidApp.Hide()
        FetchEditApp()
    End Sub
    'Public EmpFID As Object
    Public tfull As String

    Private Sub FetchEditApp()
        Dim FullName As String
        Dim EmpCode As Integer
        Dim tempstr1 As String = ""
        Dim tempstr2 As String = ""
        Dim commaLoc As Integer
        'Dim c As Integer
        commaLoc = InStr(CLBApplicant.CheckedItems(0), ",")
        tempstr1 = Mid(CLBApplicant.CheckedItems(0), 1, commaLoc - 1)
        commaLoc = Len(CLBApplicant.CheckedItems(0))
        tempstr2 = Mid(CLBApplicant.CheckedItems(0), commaLoc - 3, 4)

        If SQL.DBDT IsNot Nothing Then
            SQL.DBDT.Clear()
        End If

        SQL.AddParam("@user", tempstr1)
        SQL.AddParam("@TempLssn", tempstr2)
        SQL.ExecQuery("SELECT TOP 1 * FROM Employees " &
                      "WHERE ApplicationID=1 AND EmpFirstName=@user AND Lssn=@TempLssn;")

        If SQL.RecordCount < 1 Then Exit Sub

        For Each r As DataRow In SQL.DBDT.Rows
            tempstr1 = If(IsDBNull(r("EmpMI")), String.Empty, r("EmpMI").ToString)
            FullName = r("EmpFirstName") & " " & tempstr1 & " " & r("EmpLastName")
            tfull = FullName
            LblTitleApp.Text = "Edit Applicant " & r("EmpFirstName") & " " & tempstr1 & " " & r("EmpLastName")

            EmpCode = r("EmpCode")

            TxtEmpIDApp.Text = r("EmpCode")
            TxtCoidApp.Text = r("CompanyID")
            'TxtCoidApp.Text = DirectCast(EmpFID, Integer)
            TxtFNameAA.Text = r("EmpFirstName")
            TxtMIAA.Text = If(IsDBNull(r("EmpMI")), String.Empty, r("EmpMI").ToString)
            TxtLnameAA.Text = r("EmpLastName")

            TxtAppDt.Text = DateTrimR(r("AppDate"))

            'txtAppMssn.Text = r("Mssn")
            'txtAppLssn.Text = r("Lssn")
            tempstr1 = r("Fssn")
            tempstr1 = tempstr1 & "-" & r("Mssn")
            TxtSSNAA.Text = tempstr1 & "-" & r("Lssn")
            'TxtSSNAA.Text = tempstr2
            TxtAddrAA.Text = If(IsDBNull(r("EmpAddress")), String.Empty, r("EmpAddress").ToString)
            TxtAptAA.Text = If(IsDBNull(r("EmpApt")), String.Empty, r("EmpApt").ToString)
            TxtCityAA.Text = If(IsDBNull(r("EmpCity")), String.Empty, r("EmpCity").ToString)
            CombBoxStateAA.Text = If(IsDBNull(r("EmpState")), String.Empty, r("EmpState").ToString)
            TxtZipAA.Text = If(IsDBNull(r("EmpFirstZip")), String.Empty, r("EmpFirstZip").ToString)
            'txtAppLastZip.Text = If(IsDBNull(r("EmpLastZip")), String.Empty, r("EmpLastZip").ToString)
            TxtLivAddYY.Text = If(IsDBNull(r("LivedAddYY")), String.Empty, r("LivedAddYY").ToString)
            TxtLivAddMM.Text = If(IsDBNull(r("LivedAddMM")), String.Empty, r("LivedAddMM").ToString)
            TxtHPhoneAA.Text = If(IsDBNull(r("HomePhone")), String.Empty, Trim(r("HomePhone").ToString))
            TxtWPhoneAA.Text = If(IsDBNull(r("WorkPhone")), String.Empty, Trim(r("WorkPhone").ToString))
            TxtEmailAA.Text = If(IsDBNull(r("EmpEmail")), String.Empty, r("EmpEmail").ToString)

            'TxtDOBAA.Text = If(Mid$(r("EmpDOB"), 1, 8) = "1/1/1900", String.Empty, r("EmpDOB").ToString)
            TxtDOBAA.Text = If(DateTrimR(r("EmpDOB")) = "1/1/1900", String.Empty, DateTrimR(r("EmpDOB")).ToString)
            TxtAppPosition.Text = If(IsDBNull(r("AppPosition")), String.Empty, r("AppPosition").ToString)
            TxtLegalRAA.Text = If(IsDBNull(r("LegalRTW")), String.Empty, r("LegalRTW").ToString)
            TxtWBeforeWH.Text = If(IsDBNull(r("WorkBeforeWh")), String.Empty, r("WorkBeforeWh").ToString)
            TxtWHereFm.Text = DateTrimR(If(Mid$(r("WorkHereFM"), 1, 8) = "1/1/1900", String.Empty, r("WorkHereFM").ToString))
            TxtWHereTo.Text = DateTrimR(If(Mid$(r("WorkHereTo"), 1, 8) = "1/1/1900", String.Empty, r("WorkHereTo").ToString))
            TxtPay.Text = If(IsDBNull(r("Pay")), String.Empty, CDec(r("Pay")).ToString("F2"))
            TxtPosition.Text = If(IsDBNull(r("Position")), String.Empty, r("Position").ToString)
            TxtReason.Text = If(IsDBNull(r("ReasonLeave")), String.Empty, r("ReasonLeave").ToString)

            TxtLastEmp.Text = DateTrimR(If(Mid$(r("LastEmp"), 1, 8) = "1/1/1900", String.Empty, r("LastEmp").ToString))
            TxtRefBy.Text = r("ReferedBy") & ""
            TxtExptedPay.Text = CDec(r("ExpectedPay")).ToString("F2") & ""

            CBProofAge.Checked = If(IsDBNull(r("ProofAge")), False, r("ProofAge"))
            CBWorkBefore.Checked = If(IsDBNull(r("WorkBefore")), False, r("WorkBefore"))
            CBEmployed.Checked = If(IsDBNull(r("Employed")), False, r("Employed"))
            '-------------FETCH EDIT GEN INFO-----------
            LblEGI.Text = "Edit General Info - " & FullName
            TxtEmpCodeEGI.Text = EmpCode
            CBUnPerFunc.Checked = If(IsDBNull(r("GenPerformFunc")), False, r("GenPerformFunc"))
            TxtGenResonUnFunc.Text = If(IsDBNull(r("GenReasonUnFunc")), String.Empty, r("GenReasonUnFunc").ToString)
            TxtGenNote.Text = If(IsDBNull(r("GenNoteInfo")), String.Empty, r("GenNoteInfo").ToString)
            '-------------FETCH EDIT DRIVING INFO-----------
            LblDLI.Text = "Edit Driving Info - " & FullName
            TxtEmpCodeDL.Text = EmpCode
            COMBStateDL.Text = If(IsDBNull(r("LiState")), String.Empty, r("LiState").ToString)
            'tempstr1 = If(Mid$(r("LiExpDate"), 1, 8) = "1/1/1900" Or IsDBNull(r("LiExpDate")), String.Empty, r("LiExpDate").ToString)
            tempstr1 = If(IsDBNull(r("LiExpDate")), String.Empty, r("LiExpDate").ToString)
            TxtExpDL.Text = DateTrimR(tempstr1)
            COMBClassDL.Text = If(IsDBNull(r("LiClass")), String.Empty, r("LiClass").ToString)
            TxtDL.Text = If(IsDBNull(r("LiNumber")), String.Empty, r("LiNumber").ToString)
            CBHazTank.Checked = If(IsDBNull(r("LiEndorHazTank")), False, r("LiEndorHazTank"))
            CBDTrailer.Checked = If(IsDBNull(r("LiEndorDubTrail")), False, r("LiEndorDubTrail"))
            CBHaz.Checked = If(IsDBNull(r("LiEndorHaz")), False, r("LiEndorHaz"))
            CBPax.Checked = If(IsDBNull(r("LiEndorPassenger")), False, r("LiEndorPassenger"))
            CBTank.Checked = If(IsDBNull(r("LiEndorTank")), False, r("LiEndorTank"))
            CBAirBreak.Checked = If(IsDBNull(r("LiRestricAB")), False, r("LiRestricAB"))
            CBLiDen.Checked = If(IsDBNull(r("LiDenied")), False, r("LiDenied"))
            TxtLiDenReason.Text = If(IsDBNull(r("LiDenReason")), String.Empty, r("LiDenReason").ToString)
            CBLiRevok.Checked = If(IsDBNull(r("LiRevok")), False, r("LiRevok"))
            TxtLiStateOp.Text = If(IsDBNull(r("LiStateOp")), String.Empty, r("LiStateOp").ToString)
            TxtLiSpecialE.Text = If(IsDBNull(r("LiSpecialE")), String.Empty, r("LiSpecialE").ToString)
            '-------------FETCH EDIT EDU AND TRG INFO-----------
            LblEdu.Text = "Edit Education And Training Info - " & FullName
            TxtEmpCodeEdu.Text = EmpCode
            COMBElementary.Text = If(IsDBNull(r("EdnElmGrade")), String.Empty, r("EdnElmGrade").ToString)
            COMBHighS.Text = If(IsDBNull(r("EdnHighGrade")), String.Empty, r("EdnHighGrade").ToString)
            COMBCollege.Text = If(IsDBNull(r("EdnCollegeGrade")), String.Empty, r("EdnCollegeGrade").ToString)
            TxtSchoolCity.Text = If(IsDBNull(r("EdnLastSCity")), String.Empty, r("EdnLastSCity").ToString)

            TxtTrg.Text = If(IsDBNull(r("EdnTraining")), String.Empty, r("EdnTraining").ToString)
            TxtDAwd.Text = If(IsDBNull(r("EdnDrvAwards")), String.Empty, r("EdnDrvAwards").ToString)
            TxtOExp.Text = If(IsDBNull(r("EdnOthExp")), String.Empty, r("EdnOthExp").ToString)
            TxtCourse.Text = If(IsDBNull(r("EdnCoursesT")), String.Empty, r("EdnCoursesT").ToString)
        Next
    End Sub


    Private Sub BtnUpdANA_Click(sender As Object, e As EventArgs) Handles BtnUpdANA.Click
        '################UPDATE EDIT APPLICANT

        UpdAddApp()
        'ClearAddApp()
        GBAddApp.Hide()
        'BtnAppProcess.Enabled = True
        tempcode = TxtEmpIDApp.Text()
        'tFullName = TxtFname.Text() + " " + TxtLname.Text()
        BoxViewApp()
        LblViewApp.Text = "View Applcant " & tfull
        FetchViewApp()
    End Sub
    Private Sub BtnUpdCancelANA_Click(sender As Object, e As EventArgs) Handles BtnUpdCancelANA.Click

        GBAddApp.Hide()
        tempcode = TxtEmpIDApp.Text()
        BoxViewApp()
        'tfull = TxtFname.Text()
        LblViewApp.Text = "View Applcant " & tfull
        FetchViewApp()
    End Sub
    Private Sub BtnUpdEdGenInfo_Click(sender As Object, e As EventArgs) Handles BtnUpdEdGenInfo.Click
        UpdAddApp()
        'ClearAddApp()
        GBAddApp.Hide()
        '  CALL General Info
        GBViewApp.Hide()
        'TxtEmpCodeEGI.Hide()
        GBEditGenInfo.SetBounds(260, 20, 700, 260)
        GBEditGenInfo.Show()
        '*** FETCH GENERAL INFO
        FetchEditApp()
    End Sub
    Private Sub UpdAddApp()
        Dim str1 As String

        SQL.AddParam("@FName", TxtFNameAA.Text)
        SQL.AddParam("@MName", TxtMIAA.Text)
        SQL.AddParam("@LName", TxtLnameAA.Text)
        SQL.AddParam("@Coid", TxtCoidApp.Text)
        SQL.AddParam("@AppDt", TxtAppDt.Text)
        str1 = TxtSSNAA.Text
        SQL.AddParam("@FSsn", Mid(str1, 1, 3))
        SQL.AddParam("@MSsn", Mid(str1, 5, 2))
        SQL.AddParam("@LSsn", Mid(str1, 8, 4))
        SQL.AddParam("@EmpAddr", TxtAddrAA.Text)
        SQL.AddParam("@EmpApt", TxtAptAA.Text)
        SQL.AddParam("@EmpCity", TxtCityAA.Text)
        SQL.AddParam("@EmpSt", CombBoxStateAA.Text)
        SQL.AddParam("@EmpFZip", TxtZipAA.Text)
        SQL.AddParam("@LivYY", TxtLivAddYY.Text)
        SQL.AddParam("@LivMM", TxtLivAddMM.Text)
        SQL.AddParam("@DOB", TxtDOBAA.Text)
        SQL.AddParam("@HPhone", TxtHPhoneAA.Text)
        SQL.AddParam("@WPhone", TxtWPhoneAA.Text)
        SQL.AddParam("@EmpEmail", TxtEmailAA.Text)
        SQL.AddParam("@AppPos", TxtAppPosition.Text)
        SQL.AddParam("@LegalR", TxtLegalRAA.Text)
        SQL.AddParam("@ProofA", CBProofAge.Checked)
        SQL.AddParam("@WorkB", CBWorkBefore.Checked)
        SQL.AddParam("@WorkBWH", TxtWBeforeWH.Text)
        SQL.AddParam("@WorkHFrom", TxtWHereFm.Text)
        SQL.AddParam("@WorkHTo", TxtWHereTo.Text)
        SQL.AddParam("@Pay", TxtPay.Text)
        SQL.AddParam("@Position", TxtPosition.Text)
        SQL.AddParam("@ReasonL", TxtReason.Text)
        SQL.AddParam("@Employed", CBEmployed.Checked)
        SQL.AddParam("@LEmployed", TxtLastEmp.Text)
        SQL.AddParam("@Refered", TxtRefBy.Text)
        SQL.AddParam("@ExpPay", TxtExptedPay.Text)
        '
        SQL.AddParam("@EmpcodeA", TxtEmpIDApp.Text)
        SQL.ExecQuery("UPDATE Employees " &
                      "SET EmpFirstName=@Fname, EmpMI=@Mname, EmpLastName=@Lname, AppDate=@AppDt, Fssn=@Fssn, Mssn=@Mssn, CompanyID=@Coid, " &
                      "Lssn=@Lssn, EmpAddress=@EmpAddr, EmpApt=@EmpApt, EmpCity=@EmpCity, EmpState=@EmpSt, EmpFirstZip=@EmpFZip, " &
                      "LivedAddYY=@LivYY, LivedAddMM=@LivMM, HomePhone=@HPhone, WorkPhone=@WPhone, EmpDOB=@DOB, EmpEmail=@EmpEmail,  " &
                      "AppPosition=@AppPos, LegalRTW=@LegalR, ProofAge=@ProofA, WorkBefore=@WorkB, WorkBeforeWh=@WorkBWH, WorkHereFM=@WorkHFrom, " &
                      "WorkHereTo=@WorkHTo, Pay=@Pay, Position=@Position, ReasonLeave=@ReasonL, Employed=@Employed, LastEmp=@LEmployed, " &
                      "Referedby=@Refered, ExpectedPay=@ExpPay " &
                      "WHERE EmpCode=@EmpcodeA; " &
                      "BEGIN TRANSACTION; " &
                      "COMMIT;")
        'ERROR
        If SQL.HasException(True) Then Exit Sub

        MsgBox("Application Updated Successfully")
    End Sub

    '####################
    '  VIEW APPLICANT
    '####################
    Public tempcode As Integer
    Public tempFulName As String

    Private Sub BtnViewApp_Click(sender As Object, e As EventArgs) Handles BtnViewApp.Click
        GBListApp.Hide()
        'HIDE LEFT
        GBFindEmp.Hide()
        GBCommon.Hide()
        GBRelated.Hide()
        GBMaintLook.Hide()
        'SHOW LEFT
        GBReportVApp.SetBounds(6, 20, 240, 84)
        GBReportVApp.Show()
        GBOtherVApp.SetBounds(6, 114, 240, 120)
        GBOtherVApp.Show()


        BoxViewApp()
        'FIND FULL NAME AND EMPCODE
        Dim tempstr1 As String = ""
        Dim tempstr2 As String = ""
        Dim commaLoc As Integer
        'Dim c As Integer
        commaLoc = InStr(CLBApplicant.CheckedItems(0), ",")
        tempstr1 = Mid(CLBApplicant.CheckedItems(0), 1, commaLoc - 1)
        commaLoc = Len(CLBApplicant.CheckedItems(0))
        tempstr2 = Mid(CLBApplicant.CheckedItems(0), commaLoc - 3, 4)

        If SQL.DBDT IsNot Nothing Then
            SQL.DBDT.Clear()
        End If

        SQL.AddParam("@Fname", tempstr1)
        SQL.AddParam("@TempLssn", tempstr2)
        SQL.ExecQuery("SELECT TOP 1 * FROM Employees " &
                      "WHERE ApplicationID=1 AND EmpFirstName=@Fname AND Lssn=@TempLssn;")

        If SQL.RecordCount < 1 Then Exit Sub

        For Each r As DataRow In SQL.DBDT.Rows
            'txtEmpID.Text = r("EmpCode")
            tempstr1 = If(IsDBNull(r("EmpMI")), String.Empty, r("EmpMI").ToString)

            tempFulName = " - " & r("EmpFirstName") & " " & tempstr1 & " " & r("EmpLastName")
            tempcode = r("EmpCode")
        Next

        LblViewApp.Text = "View Applcant" & tempFulName
        FetchViewApp()
    End Sub
    Private Sub BtnCancelView_Click(sender As Object, e As EventArgs) Handles BtnCancelView.Click
        'AutoScroll = False
        GBViewApp.Hide()
        GBOtherVApp.Hide()
        GBReportVApp.Hide()

        BtnAppProcess.Enabled = True

        'SHOW LEFT
        GBFindEmp.Show()
        GBCommon.Show()
        GBRelated.Show()
        GBMaintLook.Show()

    End Sub
    Private Sub BoxViewApp()
        GBViewApp.SetBounds(260, 20, 760, 2500)
        GBViewApp.Show()
        'BOX INITIAL PARAM
        box(0) = 5
        box(1) = 5
        box(2) = 5
        box(3) = 1
        box(4) = 1
        box(5) = 1
        box(6) = 1
        box(7) = 1
        box(8) = 1
        box(9) = 1
        box(10) = 1

        BtnOpenPI.Hide()
        BtnShrinkPI.Show()
        BtnOpenGI.Hide()
        BtnShrinkGI.Show()
        BtnOpenDI.Hide()
        BtnShrinkDI.Show()
        BtnOpenEdu.Show()
        BtnShrinkEdu.Hide()
        BtnOpenPLI.Show()
        BtnShrinkPLI.Hide()
        BtnOpenPEI.Show()
        BtnShrinkPEI.Hide()
        BtnOpenPDE.Show()
        BtnShrinkPDE.Hide()
        BtnOpenVH.Show()
        BtnShrinkVH.Hide()
        BtnOpenAH.Show()
        BtnShrinkAH.Hide()
        BtnOpenPA.Show()
        BtnShrinkPA.Hide()
        BtnOpenAA.Show()
        BtnShrinkAA.Hide()

        boxpara(0, 0) = 5
        boxpara(0, 1) = 60
        boxpara(0, 2) = 750
        boxpara(0, 3) = 380

        boxpara(1, 0) = 5
        boxpara(1, 1) = 92
        boxpara(1, 2) = 750
        boxpara(1, 3) = 140

        boxpara(2, 0) = 5
        boxpara(2, 1) = 122
        boxpara(2, 2) = 750
        boxpara(2, 3) = 240

        boxpara(3, 0) = 5
        boxpara(3, 1) = 152
        boxpara(3, 2) = 750
        boxpara(3, 3) = 240

        boxpara(4, 0) = 5
        boxpara(4, 1) = 182
        boxpara(4, 2) = 750
        boxpara(4, 3) = 175

        boxpara(5, 0) = 5
        boxpara(5, 1) = 212
        boxpara(5, 2) = 750
        boxpara(5, 3) = 175

        boxpara(6, 0) = 5
        boxpara(6, 1) = 242
        boxpara(6, 2) = 750
        boxpara(6, 3) = 155

        boxpara(7, 0) = 5
        boxpara(7, 1) = 272
        boxpara(7, 2) = 750
        boxpara(7, 3) = 155

        boxpara(8, 0) = 5
        boxpara(8, 1) = 302
        boxpara(8, 2) = 750
        boxpara(8, 3) = 155

        boxpara(9, 0) = 5
        boxpara(9, 1) = 332
        boxpara(9, 2) = 750
        boxpara(9, 3) = 155

        boxpara(10, 0) = 5
        boxpara(10, 1) = 332
        boxpara(10, 2) = 750
        boxpara(10, 3) = 155

        BtnEditPLIView.Enabled = False
        BtnDelPLIView.Enabled = False

        BtnDelPEIView.Enabled = False
        BtnEditPEIView.Enabled = False

        BtnEditPDExView.Enabled = False
        BtnDelPDExView.Enabled = False

        BtnEditVHView.Enabled = False
        BtnDelVHView.Enabled = False

        BtnEditAHView.Enabled = False
        BtnDelAHView.Enabled = False

        BtnEditPAView.Enabled = False
        BtnDelPAView.Enabled = False

        BtnEditAAView.Enabled = False
        BtnDelAAView.Enabled = False
        BtnOpenFileAAView.Enabled = False



        CloseOpenViewApp()

    End Sub

    Private Sub CloseOpenViewApp()
        Dim i As Integer = 0
        Dim j As Integer = 0

        Dim x As Integer
        Dim y As Integer
        Dim w As Integer
        Dim h As Integer

        Dim t As Integer = 0
        Dim ty As Integer = 0
        Dim th As Integer = 0

        For i = 0 To 10
            x = 0
            y = 0
            w = 0
            h = 0

            For j = 0 To 3

                t = boxpara(i, j)

                If j = 0 Then
                    x = t
                End If
                If j = 1 Then
                    y = t
                End If
                If j = 2 Then
                    w = t
                End If
                If j = 3 Then
                    h = t
                End If
            Next

            'BOX 1

            If i = 0 Then
                If box(i) > 1 Then
                    ty = y
                    th = h
                    GBPinfoView.SetBounds(x, ty, w, th)
                Else
                    ty = y
                    th = cmnH
                    GBPinfoView.SetBounds(x, ty, w, th)
                End If
            End If

            'BOX 2
            If i = 1 Then
                y = ty + th + 2
                If box(i) > 1 Then
                    ty = y
                    th = h
                    GBGInfoView.SetBounds(x, ty, w, th)
                Else
                    ty = y
                    th = cmnH
                    GBGInfoView.SetBounds(x, ty, w, th)
                End If
            End If
            'BOX 3
            If i = 2 Then
                y = ty + th + 2
                If box(i) > 1 Then
                    ty = y
                    th = h
                    GBDInfoView.SetBounds(x, ty, w, th)
                Else
                    ty = y
                    th = cmnH
                    GBDInfoView.SetBounds(x, ty, w, th)
                End If
            End If

            'BOX 4
            If i = 3 Then
                y = ty + th + 2
                If box(i) > 1 Then
                    ty = y
                    th = h
                    GBEduView.SetBounds(x, ty, w, th)
                Else
                    ty = y
                    th = cmnH
                    GBEduView.SetBounds(x, ty, w, th)
                End If
            End If
            'BOX 5
            If i = 4 Then
                y = ty + th + 2
                If box(i) > 1 Then
                    ty = y
                    th = h
                    GBPreLiView.SetBounds(x, ty, w, th)
                Else
                    ty = y
                    th = cmnH
                    GBPreLiView.SetBounds(x, ty, w, th)
                End If
            End If

            'BOX 6
            If i = 5 Then
                y = ty + th + 2
                If box(i) > 1 Then
                    ty = y
                    th = h
                    GBPreEiView.SetBounds(x, ty, w, th)
                Else
                    ty = y
                    th = cmnH
                    GBPreEiView.SetBounds(x, ty, w, th)
                End If
            End If
            'BOX 7
            If i = 6 Then
                y = ty + th + 2
                If box(i) > 1 Then
                    ty = y
                    th = h
                    GBPDrvExpView.SetBounds(x, ty, w, th)
                Else
                    ty = y
                    th = cmnH
                    GBPDrvExpView.SetBounds(x, ty, w, th)
                End If
            End If

            'BOX 8
            If i = 7 Then
                y = ty + th + 2
                If box(i) > 1 Then
                    ty = y
                    th = h
                    GBVioHisView.SetBounds(x, ty, w, th)
                Else
                    ty = y
                    th = cmnH
                    GBVioHisView.SetBounds(x, ty, w, th)
                End If
            End If
            'BOX 9
            If i = 8 Then
                y = ty + th + 2
                If box(i) > 1 Then
                    ty = y
                    th = h
                    GBAccHisView.SetBounds(x, ty, w, th)
                Else
                    ty = y
                    th = cmnH
                    GBAccHisView.SetBounds(x, ty, w, th)
                End If
            End If

            'BOX 10
            If i = 9 Then
                y = ty + th + 2
                If box(i) > 1 Then
                    ty = y
                    th = h
                    GBPreAddrView.SetBounds(x, ty, w, th)
                Else
                    ty = y
                    th = cmnH
                    GBPreAddrView.SetBounds(x, ty, w, th)
                End If
            End If
            'BOX 11
            If i = 10 Then
                y = ty + th + 2
                If box(i) > 1 Then
                    ty = y
                    th = h
                    GBAppAttachView.SetBounds(x, ty, w, th)
                Else
                    ty = y
                    th = cmnH
                    GBAppAttachView.SetBounds(x, ty, w, th)
                End If
            End If
        Next
    End Sub

    Private Sub CLBApplicant_SelectedIndexChanged(sender As Object, e As MouseEventArgs)

    End Sub
    Private Sub BtnOpenPI_Click(sender As Object, e As EventArgs) Handles BtnOpenPI.Click
        box(0) = 5
        BtnOpenPI.Hide()
        BtnShrinkPI.Show()
        CloseOpenViewApp()
    End Sub
    Private Sub BtnShrinkPI_Click(sender As Object, e As EventArgs) Handles BtnShrinkPI.Click
        box(0) = 1
        BtnOpenPI.Show()
        BtnShrinkPI.Hide()
        CloseOpenViewApp()
    End Sub

    Private Sub BtnOpenGI_Click(sender As Object, e As EventArgs) Handles BtnOpenGI.Click
        box(1) = 5
        BtnOpenGI.Hide()
        BtnShrinkGI.Show()
        CloseOpenViewApp()
    End Sub
    Private Sub BtnShrinkGI_Click(sender As Object, e As EventArgs) Handles BtnShrinkGI.Click
        box(1) = 1
        BtnOpenGI.Show()
        BtnShrinkGI.Hide()
        CloseOpenViewApp()
    End Sub
    Private Sub BtnOpenDI_Click(sender As Object, e As EventArgs) Handles BtnOpenDI.Click
        box(2) = 5
        BtnOpenDI.Hide()
        BtnShrinkDI.Show()
        CloseOpenViewApp()
    End Sub
    Private Sub BtnShrinkDI_Click(sender As Object, e As EventArgs) Handles BtnShrinkDI.Click
        box(2) = 1
        BtnOpenDI.Show()
        BtnShrinkDI.Hide()
        CloseOpenViewApp()
    End Sub

    Private Sub BtnOpenEdu_Click(sender As Object, e As EventArgs) Handles BtnOpenEdu.Click
        box(3) = 5
        BtnOpenEdu.Hide()
        BtnShrinkEdu.Show()
        CloseOpenViewApp()
    End Sub
    Private Sub BtnShrinkEdu_Click(sender As Object, e As EventArgs) Handles BtnShrinkEdu.Click
        box(3) = 1
        BtnOpenEdu.Show()
        BtnShrinkEdu.Hide()
        CloseOpenViewApp()
    End Sub
    Private Sub BtnOpenPLI_Click(sender As Object, e As EventArgs) Handles BtnOpenPLI.Click
        box(4) = 5
        BtnOpenPLI.Hide()
        BtnShrinkPLI.Show()
        CloseOpenViewApp()
    End Sub
    Private Sub BtnShrinkPLI_Click(sender As Object, e As EventArgs) Handles BtnShrinkPLI.Click
        box(4) = 1
        BtnOpenPLI.Show()
        BtnShrinkPLI.Hide()
        CloseOpenViewApp()
    End Sub

    Private Sub BtnOpenPEI_Click(sender As Object, e As EventArgs) Handles BtnOpenPEI.Click
        box(5) = 5
        BtnOpenPEI.Hide()
        BtnShrinkPEI.Show()
        CloseOpenViewApp()
    End Sub
    Private Sub BtnShrinkPEI_Click(sender As Object, e As EventArgs) Handles BtnShrinkPEI.Click
        box(5) = 1
        BtnOpenPEI.Show()
        BtnShrinkPEI.Hide()
        CloseOpenViewApp()
    End Sub
    Private Sub BtnOpenPDE_Click(sender As Object, e As EventArgs) Handles BtnOpenPDE.Click
        box(6) = 5
        BtnOpenPDE.Hide()
        BtnShrinkPDE.Show()
        CloseOpenViewApp()
    End Sub
    Private Sub BtnShrinkPDE_Click(sender As Object, e As EventArgs) Handles BtnShrinkPDE.Click
        box(6) = 1
        BtnOpenPDE.Show()
        BtnShrinkPDE.Hide()
        CloseOpenViewApp()
    End Sub

    Private Sub BtnOpenVH_Click(sender As Object, e As EventArgs) Handles BtnOpenVH.Click
        box(7) = 5
        BtnOpenVH.Hide()
        BtnShrinkVH.Show()
        CloseOpenViewApp()
    End Sub
    Private Sub BtnShrinkVH_Click(sender As Object, e As EventArgs) Handles BtnShrinkVH.Click
        box(7) = 1
        BtnOpenVH.Show()
        BtnShrinkVH.Hide()
        CloseOpenViewApp()
    End Sub
    Private Sub BtnOpenAH_Click(sender As Object, e As EventArgs) Handles BtnOpenAH.Click
        box(8) = 5
        BtnOpenAH.Hide()
        BtnShrinkAH.Show()
        CloseOpenViewApp()
    End Sub
    Private Sub BtnShrinkAH_Click(sender As Object, e As EventArgs) Handles BtnShrinkAH.Click
        box(8) = 1
        BtnOpenAH.Show()
        BtnShrinkAH.Hide()
        CloseOpenViewApp()
    End Sub

    Private Sub BtnOpenPA_Click(sender As Object, e As EventArgs) Handles BtnOpenPA.Click
        box(9) = 5
        BtnOpenPA.Hide()
        BtnShrinkPA.Show()
        CloseOpenViewApp()
    End Sub
    Private Sub BtnShrinkPA_Click(sender As Object, e As EventArgs) Handles BtnShrinkPA.Click
        box(9) = 1
        BtnOpenPA.Show()
        BtnShrinkPA.Hide()
        CloseOpenViewApp()
    End Sub
    Private Sub BtnOpenAA_Click(sender As Object, e As EventArgs) Handles BtnOpenAA.Click
        box(10) = 5
        BtnOpenAA.Hide()
        BtnShrinkAA.Show()
        CloseOpenViewApp()
    End Sub
    Private Sub BtnShrinkAA_Click(sender As Object, e As EventArgs) Handles BtnShrinkAA.Click
        box(10) = 1
        BtnOpenAA.Show()
        BtnShrinkAA.Hide()
        CloseOpenViewApp()
    End Sub

    Private Sub FetchViewApp()
        Dim tempstr1 As String = ""
        Dim tempstr2 As String = ""
        Dim commaLoc As Integer

        If SQL.DBDT IsNot Nothing Then
            SQL.DBDT.Clear()
        End If

        LBPinfoView.Items.Clear()
        LBGinfoView.Items.Clear()
        LBDinfoView.Items.Clear()
        LBEduTrgInfo.Items.Clear()

        SQL.AddParam("@TEmpCode", tempcode)
        'SQL.AddParam("@TempLssn", tempstr2)
        SQL.ExecQuery("SELECT TOP 1 * FROM Employees " &
                      "WHERE ApplicationID=1  AND EmpCode=@TEmpCode;")

        If SQL.RecordCount < 1 Then Exit Sub
        For Each r As DataRow In SQL.DBDT.Rows
            '#txtEmpID.Text = r("EmpCode")
            tempstr1 = If(IsDBNull(r("EmpMI")), String.Empty, r("EmpMI").ToString)
            'gbViewApp.Text = "View Applcant - " & r("EmpFirstName") & " " & tempstr1 & " " & r("EmpLastName")
            'gbEditGenInfo.Text = "Edit General Info - " & r("EmpFirstName") & " " & tempstr1 & " " & r("EmpLastName")
            'gbEditDrvInfo.Text = "Edit Driver Info - " & r("EmpFirstName") & " " & tempstr1 & " " & r("EmpLastName")
            'gbEditEduInfo.Text = "Edit Education And Training Info - " & r("EmpFirstName") & " " & tempstr1 & " " & r("EmpLastName")


            '#gbAddPreLi.Text = "Add Previous License Info - " & r("EmpFirstName") & " " & tempstr1 & " " & r("EmpLastName")
            'tempFulName = " - " & r("EmpFirstName") & " " & tempstr1 & " " & r("EmpLastName")
            '*******************
            'LIST BOX ENTRY PERSONAL INFO                              
            LBPinfoView.Items.Add("Date of Application" & vbTab & DateTrimR(r("AppDate")))
            tempstr1 = If(IsDBNull(r("AppPosition")), String.Empty, r("AppPosition").ToString)
            LBPinfoView.Items.Add("Position(s) Applied for" & vbTab & tempstr1)
            LBPinfoView.Items.Add("Social Security" & vbTab & "XXX-XXX-" & r("Lssn"))
            LBPinfoView.Items.Add("Home Address" & vbTab & If(IsDBNull(r("EmpAddress")), String.Empty, r("EmpAddress").ToString))
            LBPinfoView.Items.Add(vbTab & vbTab & If(IsDBNull(r("EmpApt")), String.Empty, r("EmpApt").ToString))

            tempstr1 = If(IsDBNull(r("EmpCity")), String.Empty, r("EmpCity").ToString) & "," &
                                " " & If(IsDBNull(r("EmpState")), String.Empty, r("EmpState").ToString) &
                                " " & If(IsDBNull(r("EmpFirstZip")), String.Empty, r("EmpFirstZip").ToString) &
                                "-" & If(IsDBNull(r("EmpLastZip")), String.Empty, r("EmpLastZip").ToString)
            LBPinfoView.Items.Add(vbTab & vbTab & tempstr1)

            tempstr1 = "Home Phone" & vbTab & If(IsDBNull(r("HomePhone")), String.Empty, r("HomePhone").ToString) & vbTab &
                        "Work Phone" & vbTab & If(IsDBNull(r("WorkPhone")), String.Empty, r("WorkPhone").ToString)
            LBPinfoView.Items.Add(tempstr1)
            tempstr1 = If(IsDBNull(r("LivedAddYY")), String.Empty, r("LivedAddYY").ToString)
            tempstr2 = If(IsDBNull(r("LivedAddMM")), String.Empty, r("LivedAddMM").ToString)
            LBPinfoView.Items.Add("How long at this address" & vbTab & tempstr1 & " Year and " & tempstr2 & " Months")
            LBPinfoView.Items.Add("Email Address " & vbTab & If(IsDBNull(r("EmpEmail")), String.Empty, r("EmpEmail").ToString))

            'txtAppFssn.Text = r("Fssn")
            'txtAppMssn.Text = r("Mssn")
            tempstr1 = If(IsDBNull(r("LegalRTW")), String.Empty, r("LegalRTW").ToString)
            LBPinfoView.Items.Add("Do you have the legal right to work in the USA?" & vbTab & tempstr1)
            tempstr1 = If(Mid$(r("EmpDOB"), 1, 8) = "1/1/1900", String.Empty, r("EmpDOB").ToString)

            If r("ProofAge") = 0 Then
                tempstr2 = "No"
            Else
                tempstr2 = "Yes"
            End If
            LBPinfoView.Items.Add("Date of Birth" & vbTab & Mid(tempstr1, 1, 10) & vbTab & "Can provide proof of Age?" & vbTab & tempstr2)

            If r("WorkBefore") = 0 Then
                tempstr1 = "No"
            Else
                tempstr1 = "Yes"
            End If
            tempstr2 = If(IsDBNull(r("WorkBeforeWh")), String.Empty, r("WorkBeforeWh").ToString)
            LBPinfoView.Items.Add("Have you worked for this Company before?   " & tempstr1 & vbTab & "Where?  " & tempstr2)

            tempstr1 = If(Mid$(r("WorkHereFM"), 1, 8) = "1/1/1900", String.Empty, r("WorkHereFM").ToString)
            tempstr2 = If(Mid$(r("WorkHereTo"), 1, 8) = "1/1/1900", String.Empty, r("WorkHereTo").ToString)
            commaLoc = If(IsDBNull(r("Pay")), String.Empty, r("Pay").ToString)

            LBPinfoView.Items.Add("Date From   " & Mid(tempstr1, 1, 10) & "To  " & Mid(tempstr2, 1, 10) & vbTab & " Rate of Pay Position  " & commaLoc)

            tempstr1 = If(IsDBNull(r("ReasonLeave")), String.Empty, r("ReasonLeave").ToString)
            LBPinfoView.Items.Add("Reason for leaving" & vbTab & tempstr1)

            If r("Employed") = 0 Then
                tempstr1 = "No"
                If Mid$(r("LastEmp"), 1, 8) = "1/1/1900" Then
                    commaLoc = 0
                Else
                    ' tempstr2 = If(Mid$(r("LastEmp"), 1, 8) = "1/1/1900", String.Empty, r("LastEmp").ToString)
                    tempstr2 = r("LastEmp").ToString
                    tempstr2 = DateTrimR(tempstr2)

                    Dim TempDate As Date = CDate(tempstr2)
                    Dim date2 As New System.DateTime(TempDate.Year, TempDate.Month, TempDate.Day)
                    Dim date1 = Now()


                    Dim TempDiff As System.TimeSpan
                    TempDiff = date1.Subtract(date2)

                    Dim RemDays = (Int(TempDiff.TotalDays))

                    commaLoc = RemDays / 30
                End If

            Else
                tempstr1 = "Yes"
                commaLoc = 0
            End If

            LBPinfoView.Items.Add("Are you now Employed?   " & tempstr1 & vbTab & "     If not, how long since leaving last employment?  " & commaLoc & " Months")

            tempstr1 = If(IsDBNull(r("ReferedBy")), String.Empty, r("ReferedBy").ToString)
            tempstr2 = If(IsDBNull(r("ExpectedPay")), String.Empty, r("ExpectedPay").ToString)

            commaLoc = Len(tempstr2)
            tempstr2 = Mid(tempstr2, 1, commaLoc - 2)
            LBPinfoView.Items.Add("Who referred you? " & vbTab & tempstr1 & vbTab & "Rate of pay expected?  " & tempstr2)
            '####################
            ' VIEW  GENERAL INFO
            '####################
            'txtEmpCodeEGI.Clear()
            'commaLoc = If(IsDBNull(r("EmpCode")), 0, r("EmpCode"))
            'txtEmpCodeEGI.Text = commaLoc

            commaLoc = If(IsDBNull(r("GenPerformFunc")), 0, r("GenPerformFunc"))
            If commaLoc = 0 Then
                tempstr1 = "No"
                'CBUnPerFunc.Checked = False

            Else
                tempstr1 = "Yes"
                ' CBUnPerFunc.Checked = True
            End If
            LBGinfoView.Items.Add("Unable to Perform Function? " & vbTab & tempstr1)
            tempstr1 = If(IsDBNull(r("GenReasonUnFunc")), String.Empty, r("GenReasonUnFunc").ToString)
            LBGinfoView.Items.Add("Reason " & vbTab & tempstr1)
            'txtGenResonUnFunc.Text = tempstr1
            tempstr1 = If(IsDBNull(r("GenNoteInfo")), String.Empty, r("GenNoteInfo").ToString)
            LBGinfoView.Items.Add("Notes " & vbTab & tempstr1)
            'txtGenNote.Text = tempstr1
            '##################
            'VIEW DRIVING INFO
            '##################
            'txtEmpCodeDL.Clear()
            commaLoc = If(IsDBNull(r("EmpCode")), 0, r("EmpCode"))
            'txtEmpCodeDL.Text = commaLoc
            tempstr1 = If(IsDBNull(r("LiState")), String.Empty, r("LiState").ToString)
            'COMBState.Text = tempstr1
            tempstr1 = If(IsDBNull(r("LiNumber")), String.Empty, r("LiNumber").ToString)
            'txtDL.Text = tempstr1

            tempstr2 = If(IsDBNull(r("LiExpDate")), String.Empty, DateTrimR(r("LiExpDate")).ToString)
            'tempstr2 = If(Mid$(r("LiExpDate"), 1, 8) = "1/1/1900", String.Empty, r("LiExpDate").ToString)
            commaLoc = InStr(tempstr2, " ")
            LBDinfoView.Items.Add("License " & vbTab & tempstr1 & vbTab & vbTab & vbTab & "Expires  " & tempstr2)
            'txtExpDL.Text = Mid(tempstr2, 1, commaLoc)
            tempstr1 = If(IsDBNull(r("LiClass")), String.Empty, r("LiClass").ToString)
            LBDinfoView.Items.Add("Class  " & vbTab & tempstr1)
            'COMBClassDL.Text = tempstr1

            commaLoc = If(IsDBNull(r("LiEndorHazTank")), 0, r("LiEndorHazTank"))
            If commaLoc = 0 Then
                tempstr1 = ""
                ' CBHazTank.Checked = False
            Else
                tempstr1 = "HAZ, Tank"
                'CBHazTank.Checked = True
            End If
            tempstr2 = tempstr1

            commaLoc = If(IsDBNull(r("LiEndorDubTrail")), 0, r("LiEndorDubTrail"))
            If commaLoc = 0 Then
                tempstr1 = ""
                'CBDoubleTrailer.Checked = False
            Else
                tempstr1 = "Double/Trip Trailer"
                tempstr2 = tempstr2 & ", " & tempstr1
                'CBDoubleTrailer.Checked = True
            End If

            commaLoc = If(IsDBNull(r("LiEndorHaz")), 0, r("LiEndorHaz"))
            If commaLoc = 0 Then
                tempstr1 = ""
                'CBHaz.Checked = False
            Else
                tempstr1 = "Haz Material"
                tempstr2 = tempstr2 & ", " & tempstr1
                'CBHaz.Checked = False
            End If

            commaLoc = If(IsDBNull(r("LiEndorPassenger")), 0, r("LiEndorPassenger"))
            If commaLoc = 0 Then
                tempstr1 = ""
                'CBPax.Checked = False
            Else
                tempstr1 = "Passenger"
                tempstr2 = tempstr2 & ", " & tempstr1
                ' CBPax.Checked = True
            End If

            commaLoc = If(IsDBNull(r("LiEndorTank")), 0, r("LiEndorTank"))
            If commaLoc = 0 Then
                tempstr1 = ""
                'CBTank.Checked = False
            Else
                tempstr1 = "Tank"
                tempstr2 = tempstr2 & ", " & tempstr1
                'CBTank.Checked = True
            End If

            LBDinfoView.Items.Add("Endorsement " & vbTab & tempstr2)

            commaLoc = If(IsDBNull(r("LiDenied")), 0, r("LiDenied"))
            If commaLoc = 0 Then
                tempstr1 = "No"
                'CBLiDen.Checked = False
            Else
                tempstr1 = "Yes"
                'CBLiDen.Checked = True
            End If
            LBDinfoView.Items.Add("License Denied? " & vbTab & tempstr1)

            tempstr1 = If(IsDBNull(r("LiDenReason")), String.Empty, r("LiDenReason").ToString)
            LBDinfoView.Items.Add("Reason " & vbTab & tempstr1)
            'txtLiDenReason.Text = tempstr1

            commaLoc = If(IsDBNull(r("LiRevok")), 0, r("LiRevok"))
            If commaLoc = 0 Then
                tempstr1 = "No"
                'CBLiRevok.Checked = False
            Else
                tempstr1 = "Yes"
                'CBLiRevok.Checked = True
            End If
            LBDinfoView.Items.Add("License Revoked? " & vbTab & tempstr1)

            tempstr1 = If(IsDBNull(r("LiRevReason")), String.Empty, r("LiRevReason").ToString)
            LBDinfoView.Items.Add("Reason " & vbTab & tempstr1)
            'txtLiRevReason.Text = tempstr1
            tempstr1 = If(IsDBNull(r("LiStateOp")), String.Empty, r("LiStateOp").ToString)
            LBDinfoView.Items.Add("States Operated In " & vbTab & tempstr1)
            'txtLiStateOp.Text = tempstr1
            tempstr1 = If(IsDBNull(r("LiSpecialE")), String.Empty, r("LiSpecialE").ToString)
            LBDinfoView.Items.Add("Special Equipment " & vbTab & tempstr1)
            'txtLiSpecialE.Text = tempstr1
            '##################
            'VI EDU AND TRG INFO
            '##################
            'txtEmpCodeEdu.Clear()
            commaLoc = If(IsDBNull(r("EmpCode")), 0, r("EmpCode"))
            'txtEmpCodeEdu.Text = commaLoc

            commaLoc = If(IsDBNull(r("EdnElmGrade")), 0, r("EdnElmGrade"))
            LBEduTrgInfo.Items.Add("Elementary Gread Completed?" & vbTab & commaLoc)
            'COMBElementary.Text = commaLoc
            commaLoc = If(IsDBNull(r("EdnHighGrade")), 0, r("EdnHighGrade"))
            LBEduTrgInfo.Items.Add("High School Grade Completed?" & vbTab & commaLoc)
            'COMBHighS.Text = commaLoc
            commaLoc = If(IsDBNull(r("EdnCollegeGrade")), 0, r("EdnCollegeGrade"))
            LBEduTrgInfo.Items.Add("College Grade Completed?" & vbTab & commaLoc)
            'COMBCollege.Text = commaLoc
            tempstr1 = If(IsDBNull(r("EdnLastSchool")), String.Empty, r("EdnLastSchool").ToString)
            LBEduTrgInfo.Items.Add("Last School Name? " & vbTab & tempstr1)
            'txtSchoolName.Text = tempstr1
            tempstr1 = If(IsDBNull(r("EdnLastSCity")), String.Empty, r("EdnLastSCity").ToString)
            LBEduTrgInfo.Items.Add("Last School City? " & vbTab & tempstr1)
            'txtSchoolCity.Text = tempstr1
            tempstr1 = If(IsDBNull(r("EdnTraining")), String.Empty, r("EdnTraining").ToString)
            LBEduTrgInfo.Items.Add("Training " & vbTab & tempstr1)
            'txtTrg.Text = tempstr1
            tempstr1 = If(IsDBNull(r("EdnDrvAwards")), String.Empty, r("EdnDrvAwards").ToString)
            LBEduTrgInfo.Items.Add("Driving Awards " & vbTab & tempstr1)
            'txtDAwd.Text = tempstr1
            tempstr1 = If(IsDBNull(r("EdnOthExp")), String.Empty, r("EdnOthExp").ToString)
            LBEduTrgInfo.Items.Add("Other Experience " & vbTab & tempstr1)
            'txtOExp.Text = tempstr1
            tempstr1 = If(IsDBNull(r("EdnCoursesT")), String.Empty, r("EdnCoursesT").ToString)
            LBEduTrgInfo.Items.Add("Courses Taken " & vbTab & tempstr1)
            'txtCourse.Text = tempstr1
            'txtAppCompany.Text = If(IsDBNull(r("CompanyID")), String.Empty, r("CompanyID").ToString)
        Next
        'FETCH PREVIOUS LICENSE INFO
        FetchViewPreLi()
        'FETCH PREVIOUS EMPLOYER INFO
        FetchViewPreEmp()
        'FETCH PREVIOUS EXPERIENCE
        FetchViewPreExp()
        'FETCH PREVIOUS VIOLATIONS
        FetchViewPreVio()
        'FETCH PREVIOUS ACCIDENT
        FetchViewPreAcc()
        'FETCH PREVIOUS ADDRESS
        FetchViewPreAddr()
        'FETCH APPLICANT ATTACHMENT
        FetchViewAppAtc()
    End Sub
    'FETCH PREVIOUS LICENSE INFO
    Private Sub FetchViewPreLi()
        CLBPreLicenInfo.Items.Clear()

        If SQL.DBDT IsNot Nothing Then
            SQL.DBDT.Clear()
        End If

        SQL.AddParam("@Code", tempcode)
        SQL.ExecQuery("SELECT PreLiState,PreLiNumber,PreLiClass,PreLiExpDate FROM PreLicenInfo " &
                         "WHERE EmpCode=@Code " &
                         "ORDER BY PreLiNumber ASC;")

        ' If SQL.RecordCount < 1 Then Exit Sub
        If SQL.HasException(True) Then Exit Sub
        'LOOP
        Dim countltr As Integer = 0
        Dim blank1st As Integer = 0
        Dim str1 As String = ""
        Dim str2 As String = ""
        Dim str3 As String = ""
        Dim str4 As String = ""
        Dim sumstr As String = ""
        For Each r As DataRow In SQL.DBDT.Rows
            str1 = Trim(r("PreLiNumber"))

            str2 = Trim(If(IsDBNull(r("PreLiClass")), String.Empty, r("PreLiClass").ToString))

            str3 = If(IsDBNull(r("PreLiExpDate")), String.Empty, r("PreLiExpDate").ToString)
            str3 = DateTrimR(str3)
            'blank1st = InStr(str3, " ")
            'str3 = Mid(str3, 1, blank1st)
            'str3 = Replace$(str3, "/", "-")
            str4 = Trim(If(IsDBNull(r("PreLiState")), String.Empty, r("PreLiState").ToString))
            'countltr = str1.Length + str2.Length
            'If countltr < 13 Or countltr = 13 Then
            sumstr = str1 & vbTab & vbTab & str2 & vbTab & vbTab & str4 & vbTab & vbTab & vbTab & str3
            '
            CLBPreLicenInfo.Items.Add(sumstr)
        Next
    End Sub

    'FETCH PREVIOUS EMPLOYER INFO
    Private Sub FetchViewPreEmp()
        CLBPreEmpInfo.Items.Clear()

        If SQL.DBDT IsNot Nothing Then
            SQL.DBDT.Clear()
        End If

        SQL.AddParam("@PECode", tempcode)
        SQL.ExecQuery("SELECT PreEmpName,PreEmpWorkFm,PreEmpWorkTo FROM PreEmpInfo " &
                         "WHERE EmpCode=@PECode " &
                         "ORDER BY PreEmpName ASC;")

        ' If SQL.RecordCount < 1 Then Exit Sub
        If SQL.HasException(True) Then Exit Sub
        'LOOP
        Dim countltr As Integer = 0
        Dim blank1st As Integer = 0
        Dim str1 As String = ""
        Dim str2 As String = ""
        Dim str3 As String = ""
        ' Dim str4 As String = ""
        Dim sumstr As String = ""
        For Each r As DataRow In SQL.DBDT.Rows
            str1 = Trim(r("PreEmpName"))

            str2 = Trim(If(IsDBNull(r("PreEmpWorkFm")), String.Empty, r("PreEmpWorkFm").ToString))
            str2 = DateTrimR(str2)
            If str2 = "1/1/1900" Then
                'str2 = String.Empty
                str2 = ""
            Else
                str2 = DateTrimR(str2)

            End If


            str3 = If(IsDBNull(r("PreEmpWorkTo")), String.Empty, r("PreEmpWorkTo").ToString)
            If DateTrimR(str3) = "1/1/1900" Then
                'str2 = String.Empty
                str3 = ""
            Else
                str3 = DateTrimR(str3)

            End If

            countltr = Len(str1)

            Select Case countltr > 0
                Case countltr > 0 And countltr < 7
                    sumstr = str1 & vbTab & vbTab & vbTab & vbTab & vbTab & str2 & vbTab & vbTab & str3
                Case countltr > 6 And countltr < 12
                    sumstr = str1 & vbTab & vbTab & vbTab & vbTab & str2 & vbTab & vbTab & str3
                Case countltr > 11 And countltr < 24
                    sumstr = str1 & vbTab & vbTab & vbTab & str2 & vbTab & vbTab & str3
            End Select

            CLBPreEmpInfo.Items.Add(sumstr)
        Next
    End Sub

    'FETCH PREVIOUS EXPERIENCE
    Private Sub FetchViewPreExp()
        CLBPreDrvEx.Items.Clear()

        If SQL.DBDT IsNot Nothing Then
            SQL.DBDT.Clear()
        End If

        SQL.AddParam("@PECode", tempcode)
        SQL.ExecQuery("SELECT EqptClass,DriveFm,DriveTo FROM PreDrvExp " &
                         "WHERE EmpCode=@PECode " &
                         "ORDER BY EqptClass ASC;")

        ' If SQL.RecordCount < 1 Then Exit Sub
        If SQL.HasException(True) Then Exit Sub
        'LOOP
        Dim countltr As Integer = 0
        Dim str1 As String = ""
        Dim str2 As String = ""
        Dim str3 As String = ""

        Dim sumstr As String = ""
        For Each r As DataRow In SQL.DBDT.Rows
            str1 = Trim(r("EqptClass"))

            str2 = Trim(If(IsDBNull(r("DriveFm")), String.Empty, r("DriveFm").ToString))
            If Mid(str2, 1, 8) = "1/1/1900" Then
                'str2 = String.Empty
                str2 = ""
            Else
                str2 = DateTrimR(str2)
            End If

            str3 = If(IsDBNull(r("DriveTo")), String.Empty, r("DriveTo").ToString)
            If Mid(str3, 1, 8) = "1/1/1900" Then
                'str2 = String.Empty
                str3 = ""
            Else
                str3 = DateTrimR(str3)
            End If

            countltr = Len(str1)

            Select Case countltr > 0
                Case countltr > 0 And countltr < 20
                    sumstr = str1 & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & str2 & vbTab & str3
                Case countltr > 19 And countltr < 30
                    sumstr = str1 & vbTab & vbTab & vbTab & str2 & vbTab & str3
                Case countltr > 29 And countltr < 50
                    sumstr = str1 & vbTab & vbTab & str2 & vbTab & str3
            End Select

            CLBPreDrvEx.Items.Add(sumstr)
        Next
    End Sub

    'FETCH PREVIOUS VIOLATIONS
    Private Sub FetchViewPreVio()
        CLBViolation.Items.Clear()

        If SQL.DBDT IsNot Nothing Then
            SQL.DBDT.Clear()
        End If

        SQL.AddParam("@PECode", tempcode)
        SQL.ExecQuery("SELECT Charge,Loc,VioDate FROM Violations " &
                         "WHERE EmpCode=@PECode " &
                         "ORDER BY Charge ASC;")

        ' If SQL.RecordCount < 1 Then Exit Sub
        If SQL.HasException(True) Then Exit Sub
        'LOOP
        Dim countltr As Integer = 0
        'Dim blank1st As Integer = 0
        Dim str1 As String = ""
        Dim str2 As String = ""
        Dim str3 As String = ""
        ' Dim str4 As String = ""
        Dim sumstr As String = ""
        For Each r As DataRow In SQL.DBDT.Rows
            str1 = Trim(r("Charge"))

            str2 = Trim(If(IsDBNull(r("Loc")), String.Empty, r("Loc").ToString))

            str3 = If(IsDBNull(r("VioDate")), String.Empty, r("VioDate").ToString)
            If Mid(str3, 1, 8) = "1/1/1900" Then
                'str2 = String.Empty
                str3 = ""
            Else
                str3 = DateTrimR(str3)
                'blank1st = InStr(str3, " ")
                'str3 = Mid(str3, 1, blank1st)
            End If

            countltr = Len(str1)

            Select Case countltr > 0
                Case countltr > 1 And countltr < 15
                    sumstr = str1 & vbTab & vbTab & vbTab & vbTab & vbTab & str2 & vbTab & vbTab & str3
                Case countltr > 14 And countltr < 25
                    sumstr = str1 & vbTab & vbTab & vbTab & str2 & vbTab & vbTab & str3
                Case countltr > 24 And countltr < 36
                    sumstr = str1 & vbTab & str2 & vbTab & vbTab & str3
            End Select

            CLBViolation.Items.Add(sumstr)
        Next
    End Sub

    'FETCH PREVIOUS ACCIDENT
    Private Sub FetchViewPreAcc()

        CLBAccident.Items.Clear()

        If SQL.DBDT IsNot Nothing Then
            SQL.DBDT.Clear()
        End If

        SQL.AddParam("@PECode", tempcode)
        SQL.ExecQuery("SELECT AccNature,AccDate,NoFatality,NoInjury FROM PreAccident " &
                         "WHERE EmpCode=@PECode " &
                         "ORDER BY AccDate ASC;")

        ' If SQL.RecordCount < 1 Then Exit Sub
        If SQL.HasException(True) Then Exit Sub
        'LOOP
        Dim countltr As Integer = 0
        Dim blank1st As Integer = 0
        Dim str1 As String = ""
        Dim str2 As String = ""
        Dim str3 As String = ""
        Dim str4 As String = ""
        Dim sumstr As String = ""
        For Each r As DataRow In SQL.DBDT.Rows
            str1 = Trim(r("AccNature"))

            str2 = If(IsDBNull(r("AccDate")), String.Empty, r("AccDate").ToString)
            If Mid(str2, 1, 8) = "1/1/1900" Then
                'str2 = String.Empty
                str2 = ""
            Else
                str2 = Trim(str2)
                blank1st = InStr(str2, " ")
                str2 = Mid(str2, 1, blank1st)
            End If

            str3 = Trim(If(IsDBNull(r("NoFatality")), String.Empty, r("NoFatality").ToString))
            str4 = Trim(If(IsDBNull(r("NoInjury")), String.Empty, r("NoInjury").ToString))

            countltr = Len(str1)

            Select Case countltr > 0
                Case countltr > 1 And countltr < 15
                    sumstr = str1 & vbTab & vbTab & vbTab & vbTab & str2 & vbTab & vbTab & str3 & vbTab & vbTab & str4
                Case countltr > 14 And countltr < 25
                    sumstr = str1 & vbTab & vbTab & vbTab & str2 & vbTab & vbTab & str3 & vbTab & vbTab & str4
                Case countltr > 24 And countltr < 36
                    sumstr = str1 & vbTab & str2 & vbTab & vbTab & str3 & vbTab & vbTab & str4
            End Select

            CLBAccident.Items.Add(sumstr)
        Next
    End Sub
    'FETCH PREVIOUS ADDRESS
    Private Sub FetchViewPreAddr()

        CLBPreAddr.Items.Clear()

        If SQL.DBDT IsNot Nothing Then
            SQL.DBDT.Clear()
        End If

        SQL.AddParam("@PECode", tempcode)
        SQL.ExecQuery("SELECT PAddr1,PAddr2,PCity,PState,PZip FROM PreAddr " &
                         "WHERE EmpCode=@PECode " &
                         "ORDER BY PAddr1 ASC;")

        ' If SQL.RecordCount < 1 Then Exit Sub
        If SQL.HasException(True) Then Exit Sub
        'LOOP
        Dim countltr As Integer = 0
        Dim blank1st As Integer = 0
        Dim str1 As String = ""
        Dim str2 As String = ""
        Dim str3 As String = ""
        Dim str4 As String = ""
        Dim sumstr As String = ""
        For Each r As DataRow In SQL.DBDT.Rows
            str1 = Trim(r("PAddr1"))

            ' str2 = If(IsDBNull(r("PAddr2")), String.Empty, r("PAddr2").ToString)
            'str1 = str1 & vbTab & str2

            str2 = Trim(If(IsDBNull(r("PCity")), String.Empty, r("PCity").ToString))
            str3 = Trim(If(IsDBNull(r("PState")), String.Empty, r("PState").ToString))
            str4 = Trim(If(IsDBNull(r("PZip")), String.Empty, r("PZip").ToString))

            countltr = Len(str1)

            Select Case countltr > 0
                Case countltr > 1 And countltr < 20
                    sumstr = str1 & vbTab & vbTab & vbTab & str2 & vbTab & str3 & vbTab & str4
                Case countltr > 19 And countltr < 25
                    sumstr = str1 & vbTab & vbTab & str2 & vbTab & str3 & vbTab & str4
                Case countltr > 24 And countltr < 36
                    sumstr = str1 & vbTab & str2 & vbTab & str3 & vbTab & str4
            End Select

            CLBPreAddr.Items.Add(sumstr)
        Next
    End Sub
    'FETCH APPLICANT ATTACHMENT
    Private Sub FetchViewAppAtc()
        CLBAppAttachments.Items.Clear()

        If SQL.DBDT IsNot Nothing Then
            SQL.DBDT.Clear()
        End If

        SQL.AddParam("@PECode", tempcode)
        SQL.ExecQuery("SELECT FileName,FileDate FROM FileLoc " &
                         "WHERE EmpCode=@PECode " &
                         "ORDER BY FileName ASC;")

        ' If SQL.RecordCount < 1 Then Exit Sub
        If SQL.HasException(True) Then Exit Sub
        'LOOP
        Dim countltr As Integer = 0
        Dim blank1st As Integer = 0
        Dim str1 As String = ""
        Dim str2 As String = ""
        Dim str3 As String = ""
        Dim str4 As String = ""
        Dim sumstr As String = ""
        For Each r As DataRow In SQL.DBDT.Rows
            str1 = Trim(r("FileName"))

            str2 = If(IsDBNull(r("FileDate")), String.Empty, r("FileDate").ToString)

            If Mid(str2, 1, 8) = "1/1/1900" Then
                'str2 = String.Empty
                str2 = ""
            Else
                str2 = Trim(str2)
                blank1st = InStr(str2, " ")
                str2 = Mid(str2, 1, blank1st)
            End If

            countltr = Len(str1)

            Select Case countltr > 0
                Case countltr > 1 And countltr < 15
                    sumstr = str1 & vbTab & vbTab & vbTab & vbTab & str2
                Case countltr > 14 And countltr < 25
                    sumstr = str1 & vbTab & vbTab & vbTab & str2
                Case countltr > 24 And countltr < 36
                    sumstr = str1 & vbTab & str2
            End Select

            CLBAppAttachments.Items.Add(sumstr)
        Next
    End Sub
    ' ------------------FETCH EDIT APPLICATION END-------------
    ' ------------------EDIT PERSONAL INFO APP -------------
    Private Sub BtnEditPView_Click(sender As Object, e As EventArgs) Handles BtnEditPView.Click
        GBViewApp.Hide()

        GBAddApp.SetBounds(260, 20, 700, 950)
        'gbAddApp.Text = "Edit Applicant"
        GBAddApp.Show()
        BtnAddEmp.Enabled = True
        BtnAppProcess.Enabled = False
        BtnAddApplicant.Enabled = True
        'BASIC VALIDATION
        ''gbSelectAppCompany.Hide()
        ''gbGeneral2.Show()
        BtnSaveANA.Hide()
        BtnCancelANA.Hide()
        BtnSaveAddANA.Hide()
        BtnUpdANA.SetBounds(150, 910, 81, 24)
        BtnUpdCancelANA.SetBounds(235, 910, 81, 24)
        BtnUpdEdGenInfo.SetBounds(325, 910, 165, 24)
        BtnUpdANA.Show()
        BtnUpdCancelANA.Show()
        BtnUpdEdGenInfo.Show()
        TxtCoidApp.Hide()
        TxtEmpIDApp.Hide()
        FetchEditApp()
    End Sub

    ' ------------------EDIT GENERAL INFO APP -------------
    Private Sub BtnEditGView_Click(sender As Object, e As EventArgs) Handles BtnEditGView.Click
        GBViewApp.Hide()
        TxtEmpCodeEGI.Hide()
        GBEditGenInfo.SetBounds(260, 20, 700, 260)
        GBEditGenInfo.Show()
        '*** FETCH GENERAL INFO
        FetchEditApp()

    End Sub

    Private Sub BtnCancelGI_Click(sender As Object, e As EventArgs) Handles BtnCancelGI.Click
        GBEditGenInfo.Hide()
        'btnAppProcess.Enabled = True
        BoxViewApp()
        FetchViewApp()

    End Sub

    Private Sub BtnUpdGI_Click(sender As Object, e As EventArgs) Handles BtnUpdGI.Click
        UpdateGenInfo()
    End Sub

    Private Sub UpdateGenInfo()
        SQL.AddParam("@EmpCodeEGI", TxtEmpCodeEGI.Text)

        SQL.AddParam("@UnPerFunc", CBUnPerFunc.Checked)
        SQL.AddParam("@ReasonUnFunc", TxtGenResonUnFunc.Text)
        SQL.AddParam("@GenNote", TxtGenNote.Text)

        SQL.ExecQuery("UPDATE Employees " &
                          "SET GenPerformFunc=@UnPerFunc, GenReasonUnFunc=@ReasonUnFunc, GenNoteInfo=@GenNote " &
                          "WHERE EmpCode=@EmpCodeEGI; " &
                          "BEGIN TRANSACTION; " &
                          "COMMIT;")
        'ERROR
        If SQL.HasException(True) Then Exit Sub

        MsgBox("General Info Updated Successfully")
    End Sub

    Private Sub BtnUpdEDrvInfo_Click(sender As Object, e As EventArgs) Handles BtnUpdEDrvInfo.Click
        UpdateGenInfo()
        GBEditGenInfo.Hide()
        'CALL EDIT DRIVING INFO
        GBEditDrvInfo.SetBounds(260, 20, 700, 640)
        GBEditDrvInfo.Show()
        'TxtEmpCodeEGI.Hide()
        '*** FETCH DRIVING LICENSE INFO
        FetchEditApp()
    End Sub


    ' ------------------EDIT DRIVING INFO APP -------------
    Private Sub BtnEditDLView_Click(sender As Object, e As EventArgs) Handles BtnEditDLView.Click
        GBViewApp.Hide()
        TxtEmpCodeDL.Hide()
        GBEditDrvInfo.SetBounds(260, 20, 700, 640)
        GBEditDrvInfo.Show()
        '*** FETCH DRIVING LICENSE INFO
        FetchEditApp()
    End Sub

    Private Sub BtnCancelEDrvL_Click(sender As Object, e As EventArgs) Handles BtnCancelEDrvL.Click
        GBEditDrvInfo.Hide()
        'btnAppProcess.Enabled = True
        BoxViewApp()
        FetchViewApp()
    End Sub

    Private Sub BtnUpdEDrvL_Click(sender As Object, e As EventArgs) Handles BtnUpdEDrvL.Click
        UpdateDrvLiInfo()
    End Sub

    Private Sub BtnUpdEEduTrgInfo_Click(sender As Object, e As EventArgs) Handles BtnUpdEEduTrgInfo.Click
        UpdateDrvLiInfo()
        GBEditDrvInfo.Hide()
        'CALL EDIT EDU AND TRG INFO
        GBEditEduInfo.SetBounds(260, 20, 700, 480)
        GBEditEduInfo.Show()
        '*** FETCH DRIVING LICENSE INFO
        FetchEditApp()
    End Sub

    Private Sub DTPDrvL_ValueChanged(sender As Object, e As EventArgs) Handles DTPDrvL.ValueChanged, DTPDrvL.Click, DTPDatePL.ValueChanged, DTPDatePL.Click
        TxtExpDL.Text = DateTrimR(DTPDrvL.Value)
    End Sub

    Private Sub UpdateDrvLiInfo()
        SQL.AddParam("@EmpCodeDL", TxtEmpCodeDL.Text)

        'SQL.AddParam("@LiState", CBUnPerFunc.Checked)
        SQL.AddParam("@LiState", COMBStateDL.Text)
        SQL.AddParam("@LiExpDate", TxtExpDL.Text)
        SQL.AddParam("@LiClass", COMBClassDL.Text)
        SQL.AddParam("@LiNumber", TxtDL.Text)

        SQL.AddParam("@LiEndorHTank", CBHazTank.Checked)
        SQL.AddParam("@LiEndorDTrail", CBDTrailer.Checked)
        SQL.AddParam("@LiEndorHz", CBHaz.Checked)
        SQL.AddParam("@LiEndorPx", CBPax.Checked)
        SQL.AddParam("@LiEndorTnk", CBTank.Checked)
        SQL.AddParam("@LiRestAB", CBAirBreak.Checked)

        SQL.AddParam("@LiDen", CBLiDen.Checked)
        SQL.AddParam("@LiDenReason", TxtLiDenReason.Text)

        SQL.AddParam("@LiRevk", CBLiRevok.Checked)
        SQL.AddParam("@LiRevReason", TxtLiRevReason.Text)

        SQL.AddParam("@LiStateOp", TxtLiStateOp.Text)
        SQL.AddParam("@LiSpecialE", TxtLiSpecialE.Text)

        SQL.ExecQuery("UPDATE Employees " &
                          "SET LiState=@LiState, LiExpDate=@LiExpDate, LiClass=@LiClass, LiNumber=@LiNumber, LiEndorHazTank=@LiEndorHTank, " &
                          "LiEndorDubTrail=@LiEndorDTrail, LiEndorHaz=@LiEndorHz, LiEndorPassenger=@LiEndorPx, LiEndorTank=@LiEndorTnk, " &
                          "LiRestricAB=@LiRestAB, LiDenied=@LiDen, LiDenReason=@LiDenReason, LiRevok=@LiRevk, LiRevReason=@LiRevReason, " &
                          "LiStateOp=@LiStateOp,LiSpecialE=@LiSpecialE " &
                          "WHERE EmpCode=@EmpCodeDL; " &
                          "BEGIN TRANSACTION; " &
                          "COMMIT;")
        'ERROR
        If SQL.HasException(True) Then Exit Sub

        MsgBox("Driving Info Updated Successfully")
    End Sub

    ' ------------------EDIT EDU AND TRG INFO APP -------------
    Private Sub BtnEditEduView_Click(sender As Object, e As EventArgs) Handles BtnEditEduView.Click
        GBViewApp.Hide()
        TxtEmpCodeEdu.Hide()
        GBEditEduInfo.SetBounds(260, 20, 700, 480)
        GBEditEduInfo.Show()
        '*** FETCH DRIVING LICENSE INFO
        FetchEditApp()
    End Sub

    Private Sub BtnCancelEdu_Click(sender As Object, e As EventArgs) Handles BtnCancelEdu.Click
        GBEditEduInfo.Hide()
        'btnAppProcess.Enabled = True
        BoxViewApp()
        FetchViewApp()
    End Sub

    Private Sub BtnUpdEdu_Click(sender As Object, e As EventArgs) Handles BtnUpdEdu.Click
        UpdateEduInfo()

    End Sub
    Private Sub UpdateEduInfo()
        SQL.AddParam("@EmpCodeEdu", TxtEmpCodeEdu.Text)

        SQL.AddParam("@EdnElmGrd", COMBElementary.Text)
        SQL.AddParam("@EdnHighGrd", COMBHighS.Text)
        SQL.AddParam("@EdnCollegeGrd", COMBCollege.Text)
        SQL.AddParam("@SchoolName", TxtSchoolName.Text)
        SQL.AddParam("@LCity", TxtSchoolCity.Text)
        SQL.AddParam("@EdnTrg", TxtTrg.Text)
        SQL.AddParam("@EdnDrvAwd", TxtDAwd.Text)

        SQL.AddParam("@EdnOExp", TxtOExp.Text)
        SQL.AddParam("@EdnCourses", TxtCourse.Text)

        SQL.ExecQuery("UPDATE Employees " &
                          "SET EdnElmGrade=@EdnElmGrd, EdnHighGrade=@EdnHighGrd, EdnCollegeGrade=@EdnCollegeGrd, EdnLastSchool=@SchoolName, " &
                          "EdnLastSCity=@LCity, EdnTraining=@EdnTrg, EdnDrvAwards=@EdnDrvAwd, EdnOthExp=@EdnOExp, EdnCoursesT=@EdnCourses " &
                          "WHERE EmpCode=@EmpCodeEdu; " &
                          "BEGIN TRANSACTION; " &
                          "COMMIT;")
        'ERROR
        If SQL.HasException(True) Then Exit Sub

        MsgBox("Educational Info Updated Successfully")
    End Sub
    '##############################################
    '     PREVIOUS LICENSE INFO DT: 1/29/2018
    '##############################################
    ' ------------------ADD PREVIOUS LICENSE INFO APP -------------
    Private Sub BtnAddPLIView_Click(sender As Object, e As EventArgs) Handles BtnAddPLIView.Click
        '  ADD PREVIOUS LICENSE INFO

        GBViewApp.Hide()

        GBAddPreLi.SetBounds(260, 20, 700, 250)
        GBAddPreLi.Show()

        BtnUpdPLI.Hide()
        BtnSavePLI.Show()
        BtnSaveAddPLI.Show()
        LblPreLI.Text = "Add Previous License Info" & tempFulName
        TxtEmpCodePL.Clear()
        TxtEmpCodePL.Text = tempcode
        TxtEmpCodePL.Hide()
        TxtPLID.Hide()
        CleargbAddPLI()

        BtnSavePLI.Enabled = False
        BtnSaveAddPLI.Enabled = False
    End Sub

    Private Sub BtnCancelPLI_Click(sender As Object, e As EventArgs) Handles BtnCancelPLI.Click
        'txtEmpCodePL.Clear()
        GBAddPreLi.Hide()
        'btnAppProcess.Enabled = True
        BoxViewApp()
        FetchViewApp()
    End Sub
    Private Sub BtnSavePLI_Click(sender As Object, e As EventArgs) Handles BtnSavePLI.Click
        InsertAddPreLi()
        'txtEmpCodePL.Clear()
        GBAddPreLi.Hide()
        'btnAppProcess.Enabled = True
        BoxViewApp()
        FetchViewApp()

    End Sub
    Private Sub BtnSaveAddPLI_Click(sender As Object, e As EventArgs) Handles BtnSaveAddPLI.Click
        InsertAddPreLi()
        CleargbAddPLI()
    End Sub


    Private Sub txtPLNumber_TextChanged(sender As Object, e As EventArgs) Handles TxtPLNumber.TextChanged
        BtnSavePLI.Enabled = True
        BtnSaveAddPLI.Enabled = True
    End Sub
    Private Sub DTPDatePL_ValueChanged(sender As Object, e As EventArgs) Handles DTPDatePL.ValueChanged, DTPDatePL.Click

        TxtPLExpDate.Text = DateTrimR(DTPDatePL.Value)
    End Sub

    Private Sub InsertAddPreLi()

        SQL.AddParam("@PLiEmpCode", TxtEmpCodePL.Text)

        SQL.AddParam("@PLiState", CombPLState.Text)
        SQL.AddParam("@PLiNumber", TxtPLNumber.Text)
        SQL.AddParam("@PLiClass", CombPLClass.Text)
        SQL.AddParam("@PLiExpDate", TxtPLExpDate.Text)

        SQL.ExecQuery("INSERT INTO PreLicenInfo " &
                       "(PreLiState,EmpCode,PreLiNumber,PreLiClass,PreLiExpDate) " &
                       "VALUES (@PLiState,@PLiEmpCode,@PLiNumber,@PLiClass,@PLiExpDate); " &
                       "BEGIN TRANSACTION; " &
                       "COMMIT;", True)

        'ERROR
        If SQL.HasException(True) Then Exit Sub

        MsgBox("Previous Driving License Info Added Successfully")

    End Sub
    Private Sub CleargbAddPLI()
        '   CLEAR PREVIOUS LI INFO
        CombPLState.ResetText()
        TxtPLNumber.Clear()
        CombPLClass.ResetText()
        DTPDatePL.ResetText()
        TxtPLExpDate.Text = ""

    End Sub
    ' ------------------EDIT PREVIOUS LICENSE INFO APP -------------
    Private Sub CLBPreLicenInfo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CLBPreLicenInfo.SelectedIndexChanged, CLBPreLicenInfo.Click, CLBPreLicenInfo.DoubleClick
        If CLBPreLicenInfo.CheckedItems.Count() = 0 Then
            BtnDelPLIView.Enabled = False
            BtnEditPLIView.Enabled = False
        Else
            BtnDelPLIView.Enabled = True
            BtnEditPLIView.Enabled = True
        End If
    End Sub

    Private Sub BtnEditPLIView_Click(sender As Object, e As EventArgs) Handles BtnEditPLIView.Click
        '   EDIT PREVIOUS LICENSE INFO
        GBViewApp.Hide()
        'HEAD CHK
        TxtEmpCodePL.Clear()
        TxtPLID.Clear()
        TxtEmpCodePL.Hide()
        TxtPLID.Hide()

        GBAddPreLi.SetBounds(260, 20, 700, 250)
        GBAddPreLi.Show()
        BtnSavePLI.Hide()
        BtnSaveAddPLI.Hide()
        BtnUpdPLI.SetBounds(90, 200, 80, 24)
        BtnUpdPLI.Show()
        LblPreLI.Text = "Edit Previous License Info" & tempFulName
        BtnSavePLI.Enabled = False
        BtnSaveAddPLI.Enabled = False
        '*** FETCH PREVIOUS LICENSE INFO
        CleargbAddPLI()

        FetchPLI()

    End Sub

    Private Sub FetchPLI()
        Dim tempstr1 As String = ""
        'Dim tempstr2 As String = ""
        Dim aLoc As Integer

        aLoc = InStr(CLBPreLicenInfo.CheckedItems(0), vbTab)
        tempstr1 = Mid(CLBPreLicenInfo.CheckedItems(0), 1, aLoc - 1)
        '
        If SQL.DBDT IsNot Nothing Then
            SQL.DBDT.Clear()
        End If

        SQL.AddParam("@licen1", tempstr1)
        SQL.AddParam("@tmpEmpCode1", tempcode)
        SQL.ExecQuery("SELECT TOP 1 * FROM PreLicenInfo " &
                      "WHERE PreLiNumber=@licen1 AND EmpCode=@tmpEmpCode1;")
        '
        If SQL.RecordCount < 1 Then Exit Sub
        'TxtEmpCodePL.Text = tempcode
        For Each r As DataRow In SQL.DBDT.Rows
            TxtEmpCodePL.Text = r("Empcode")
            TxtPLID.Text = r("PreLicenID")
            CombPLState.Text = If(IsDBNull(r("PreLiState")), String.Empty, r("PreLiState").ToString)
            TxtPLNumber.Text = If(IsDBNull(r("PreLiNumber")), String.Empty, r("PreLiNumber").ToString)
            CombPLClass.Text = If(IsDBNull(r("PreLiClass")), String.Empty, r("PreLiClass").ToString)
            TxtPLExpDate.Text = DateTrimR(If(IsDBNull(r("PreLiExpDate")), String.Empty, r("PreLiExpDate").ToString))
        Next
    End Sub

    Private Sub BtnUpdPLI_Click(sender As Object, e As EventArgs) Handles BtnUpdPLI.Click
        UpdatePLI()
    End Sub
    Private Sub UpdatePLI()
        SQL.AddParam("@PLicenID", TxtPLID.Text)

        SQL.AddParam("@PLiState", CombPLState.Text)
        SQL.AddParam("@PEmpCode", TxtEmpCodePL.Text)
        SQL.AddParam("@PLiNumber", TxtPLNumber.Text)
        SQL.AddParam("@PLiClass", CombPLClass.Text)
        SQL.AddParam("@PLiExpDate", TxtPLExpDate.Text)

        SQL.ExecQuery("UPDATE PreLicenInfo " &
                       "SET PreLiState=@PLiState, EmpCode=@PEmpCode,PreLiNumber=@PLiNumber, PreLiClass=@PLiClass, PreLiExpDate=@PLiExpDate " &
                       "WHERE PreLicenID=@PLicenID; " &
                       "BEGIN TRANSACTION; " &
                       "COMMIT;", True)
        'ERROR
        If SQL.HasException(True) Then Exit Sub

        MsgBox("Previous Driving License Info Updated Successfully")
    End Sub

    ' ------------------DELETE PREVIOUS LICENSE INFO APP -------------
    Private Sub BtnDelPLIView_Click(sender As Object, e As EventArgs) Handles BtnDelPLIView.Click
        '  DELETE PREVIOUS LICENSE INFO
        Dim tempstr1 As String = ""
        Dim aLoc As Integer

        aLoc = InStr(CLBPreLicenInfo.CheckedItems(0), vbTab)
        tempstr1 = Mid(CLBPreLicenInfo.CheckedItems(0), 1, aLoc - 1)

        If SQL.DBDT IsNot Nothing Then
            SQL.DBDT.Clear()
        End If

        SQL.AddParam("@licen1", tempstr1)
        SQL.AddParam("@tmpEmpCode1", tempcode)
        SQL.ExecQuery("DELETE FROM PreLicenInfo " &
                      "WHERE PreLiNumber=@licen1 AND EmpCode=@tmpEmpCode1;")
        MsgBox("Selecte Previous License deleted")

        FetchViewPreLi()
        BtnDelPLIView.Enabled = False
        BtnEditPLIView.Enabled = False

    End Sub

    '##############################################
    '     PREVIOUS EMPLOYER INFO DT: 1/31/2019
    '##############################################
    ' ------------------ADD PREVIOUS EMPLOYER INFO APP -------------
    Private Sub BtnAddPEIView_Click(sender As Object, e As EventArgs) Handles BtnAddPEIView.Click
        GBViewApp.Hide()

        GBAddPreEmp.SetBounds(260, 20, 700, 770)
        GBAddPreEmp.Show()
        BtnUpdPEmp.Hide()

        LblPEI.Text = "Add Previous Employer Info" & tempFulName
        ClearPreEmp()
        TxtEmpCodePEI.Clear()
        TxtEmpCodePEI.Text = tempcode
        TxtEmpCodePEI.Hide()
        TxtPEmpID.Hide()
        BtnSavePEmp.Enabled = False
        BtnSaveAddPEmp.Enabled = False

    End Sub

    Private Sub BtnCancelPEmp_Click(sender As Object, e As EventArgs) Handles BtnCancelPEmp.Click
        GBAddPreEmp.Hide()
        BtnAppProcess.Enabled = True
        BoxViewApp()
        FetchViewApp()
    End Sub
    Private Sub BtnSavePEmp_Click(sender As Object, e As EventArgs) Handles BtnSavePEmp.Click
        InsertPreEmp()
        GBAddPreEmp.Hide()
        BtnAppProcess.Enabled = True
        BoxViewApp()
        FetchViewApp()
    End Sub

    Private Sub BtnSaveAddPEmp_Click(sender As Object, e As EventArgs) Handles BtnSaveAddPEmp.Click
        InsertPreEmp()
        ClearPreEmp()
    End Sub
    Private Sub InsertPreEmp()
        SQL.AddParam("@PEmpName", TxtPEmpName.Text)

        SQL.AddParam("@PEmpAdd1", TxtPEmpAdd1.Text)
        SQL.AddParam("@PEmpAdd2", TxtPEmpAdd2.Text)
        SQL.AddParam("@PEmpCity", TxtPEmpCity.Text)
        SQL.AddParam("@PEmpState", CombPEmpState.Text)
        SQL.AddParam("@PEmpZip1", TxtPEmpZip1.Text)
        SQL.AddParam("@PEmpZip2", TxtPEmpZip2.Text)

        SQL.AddParam("@PEmpCode", TxtEmpCodePEI.Text)
        SQL.AddParam("@PEmpEmail", TxtPEmpEmail.Text)
        SQL.AddParam("@PEmpPhone", TxtPEmpPhone.Text)
        SQL.AddParam("@PEmpFax", TxtPEmpFax.Text)
        SQL.AddParam("@PEmpCName", TxtPEmpContName.Text)
        SQL.AddParam("@PEmpCNo", TxtPEmpContNo.Text)
        SQL.AddParam("@PEmpWFm", TxtPEmpWFm.Text)
        SQL.AddParam("@PEmpWTo", TxtPEmpWTo.Text)
        SQL.AddParam("@PEmpPost", TxtPEmpPosition.Text)
        SQL.AddParam("@PEmpWage", TxtPEmpWage.Text)

        SQL.AddParam("@PEmpFMCSR", CBPEmpFMCSR.Checked)
        SQL.AddParam("@PEmpCFR", CBPEmp49CFR.Checked)
        SQL.AddParam("@PEmpRL", TxtPEmpReasonL.Text)

        SQL.ExecQuery("INSERT INTO PreEmpInfo " &
                       "(PreEmpName,PreEmpAdd1,PreEmpAdd2,PreEmpCity,PreEmpState,PreEmpZip1,PreEmpZip2,EmpCode,PreEmpEmail,PreEmpPhone,PreEmpFax, " &
                       "PreEmpContact,PreEmpContNo,PreEmpWorkFm,PreEmpWorkTo,PreEmpPosition,PreEmpWage,PreEmpFMCSR,PreEmp49CFR,PreEmpReasonL) " &
                       "VALUES (@PEmpName,@PEmpAdd1,@PEmpAdd2,@PEmpCity,@PEmpState,@PEmpZip1,@PEmpZip2,@PEmpCode,@PEmpEmail,@PEmpPhone,@PEmpFax, " &
                       "@PEmpCName,@PEmpCNo,@PEmpWFm,@PEmpWTo,@PEmpPost,@PEmpWage,@PEmpFMCSR,@PEmpCFR,@PEmpRL); " &
                       "BEGIN TRANSACTION; " &
                       "COMMIT;", True)

        'ERROR
        If SQL.HasException(True) Then Exit Sub



        MsgBox("Previous Employer Info Added Successfully")
    End Sub
    'Private Sub TxtPEmpName_TextChanged(sender As Object, e As EventArgs) Handles TxtPEmpName.TextChanged, TxtPEmpAdd1.TextChanged, TxtPEmpCity.TextChanged
    Private Sub txtPEmpName_TextChanged(sender As Object, e As EventArgs) Handles TxtPEmpName.TextChanged, TxtPEmpAdd1.TextChanged, TxtPEmpCity.TextChanged, TxtPEmpWFm.TextChanged, TxtPEmpWTo.TextChanged, TxtEqptType.TextChanged, TxtDrvTo.TextChanged, TxtDrvFm.TextChanged, TxtVioLoc.TextChanged, TxtVioDt.TextChanged, TxtPenalty.TextChanged, TxtNoInjury.TextChanged, TxtNoFatality.TextChanged, TxtAccDt.TextChanged, TxtPCity.TextChanged, TxtPAddr2.TextChanged, TxtPAddr1.TextChanged, TxtAttFileName.TextChanged, TxtAttDate.TextChanged
        'BASIC VALIDATION
        If Not String.IsNullOrWhiteSpace(TxtPEmpName.Text) AndAlso Not String.IsNullOrWhiteSpace(TxtPEmpAdd1.Text) AndAlso Not String.IsNullOrWhiteSpace(TxtPEmpCity.Text) AndAlso Not String.IsNullOrWhiteSpace(TxtPEmpWFm.Text) AndAlso Not String.IsNullOrWhiteSpace(TxtPEmpWTo.Text) Then
            BtnSavePEmp.Enabled = True
            BtnSaveAddPEmp.Enabled = True
        Else
            BtnSavePEmp.Enabled = False
            BtnSaveAddPEmp.Enabled = False
        End If
    End Sub

    Private Sub ClearPreEmp()

        TxtPEmpName.Clear()
        TxtPEmpAdd1.Clear()
        TxtPEmpAdd2.Clear()
        TxtPEmpCity.Clear()

        CombPEmpState.ResetText()
        TxtPEmpZip1.Clear()
        TxtPEmpZip2.Clear()
        'txtEmpCodePEI.Clear()
        TxtPEmpEmail.Clear()
        TxtPEmpPhone.Clear()
        TxtPEmpFax.Clear()
        TxtPEmpContName.Clear()
        TxtPEmpContNo.Clear()
        DTPTo.ResetText()
        DTPFm.ResetText()
        TxtPEmpWFm.Clear()
        TxtPEmpWTo.Clear()
        TxtPEmpPosition.Clear()
        TxtPEmpWage.Clear()

        CBPEmpFMCSR.Checked = False
        CBPEmp49CFR.Checked = False
        TxtPEmpReasonL.Clear()

    End Sub

    Private Sub DTPFm_ValueChanged(sender As Object, e As EventArgs) Handles DTPFm.ValueChanged
        TxtPEmpWFm.Text = DateTrimR(DTPFm.Value)
    End Sub
    Private Sub DTPTo_ValueChanged(sender As Object, e As EventArgs) Handles DTPTo.ValueChanged
        TxtPEmpWTo.Text = DateTrimR(DTPTo.Value)
    End Sub
    ' ------------------EDIT PREVIOUS EMPLOYER INFO APP -------------
    Private Sub CLBPreEmpInfo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CLBPreEmpInfo.SelectedIndexChanged, CLBPreEmpInfo.Click, CLBPreEmpInfo.DoubleClick
        If CLBPreEmpInfo.CheckedItems.Count() = 0 Then
            BtnDelPEIView.Enabled = False
            BtnEditPEIView.Enabled = False
        Else
            BtnDelPEIView.Enabled = True
            BtnEditPEIView.Enabled = True
        End If
    End Sub
    Private Sub BtnEditPEIView_Click(sender As Object, e As EventArgs) Handles BtnEditPEIView.Click
        '   EDIT PREVIOUS EMP INFO

        GBViewApp.Hide()

        TxtEmpCodePEI.Clear()
        'TxtEmpCodePEI.Hide()
        GBAddPreEmp.SetBounds(260, 20, 700, 770)
        GBAddPreEmp.Show()
        BtnSavePEmp.Hide()
        BtnSaveAddPEmp.Hide()
        BtnUpdPEmp.SetBounds(131, 730, 81, 24)
        BtnUpdPEmp.Show()
        GBAddPreEmp.Text = "Edit Previous Employer Info" & tempFulName
        BtnSavePEmp.Enabled = False
        BtnSaveAddPEmp.Enabled = False
        '*** FETCH PREVIOUS LICENSE INFO
        CleargbAddPLI()
        TxtEmpCodePEI.Hide()
        TxtPEmpID.Hide()
        FetchPEI()
    End Sub

    Private Sub FetchPEI()
        Dim tempstr1 As String = ""
        Dim tempstr2 As String = ""
        Dim tempdt1 As DateTime
        Dim tempdt2 As DateTime
        Dim aLoc As Integer

        aLoc = InStr(CLBPreEmpInfo.CheckedItems(0), vbTab)
        tempstr1 = Mid(CLBPreEmpInfo.CheckedItems(0), 1, aLoc - 1)

        aLoc = InStrRev(CLBPreEmpInfo.CheckedItems(0), "/")
        aLoc = aLoc + 5
        tempstr2 = Mid(CLBPreEmpInfo.CheckedItems(0), aLoc - 10, 10)

        tempdt1 = Convert.ToDateTime(tempstr2)
        tempdt2 = DateAdd("n", 1, tempdt1)

        If SQL.DBDT IsNot Nothing Then
            SQL.DBDT.Clear()
        End If

        SQL.AddParam("@PEName", tempstr1)
        SQL.AddParam("@tmEmpCode", tempcode)
        SQL.AddParam("@Adate", tempdt1)
        SQL.AddParam("@Bdate", tempdt2)
        SQL.ExecQuery("SELECT TOP 1 * FROM PreEmpInfo " &
                      "WHERE PreEmpName=@PEName AND EmpCode=@tmEmpCode AND PreEmpWorkTo BETWEEN @Adate AND @Bdate;")

        If SQL.RecordCount < 1 Then Exit Sub
        'TxtEmpCodePEI.Text = tempcode
        For Each r As DataRow In SQL.DBDT.Rows
            'TxtEmpCodePEI.Text = r("PreEmpID")
            TxtPEmpID.Text = r("PreEmpID")
            TxtPEmpName.Text = r("PreEmpName")
            TxtPEmpAdd1.Text = If(IsDBNull(r("PreEmpAdd1")), String.Empty, r("PreEmpAdd1").ToString)

            TxtPEmpAdd2.Text = If(IsDBNull(r("PreEmpAdd2")), String.Empty, r("PreEmpAdd2").ToString)

            TxtPEmpCity.Text = If(IsDBNull(r("PreEmpCity")), String.Empty, r("PreEmpCity").ToString)
            CombPEmpState.Text = If(IsDBNull(r("PreEmpState")), String.Empty, r("PreEmpState").ToString)
            TxtPEmpZip1.Text = If(IsDBNull(r("PreEmpZip1")), String.Empty, r("PreEmpZip1").ToString)
            TxtPEmpZip2.Text = If(IsDBNull(r("PreEmpZip2")), String.Empty, r("PreEmpZip2").ToString)

            TxtEmpCodePEI.Text = r("EmpCode")
            TxtPEmpEmail.Text = If(IsDBNull(r("PreEmpEmail")), String.Empty, r("PreEmpEmail").ToString)
            TxtPEmpPhone.Text = If(IsDBNull(r("PreEmpPhone")), String.Empty, r("PreEmpPhone").ToString)
            TxtPEmpFax.Text = If(IsDBNull(r("PreEmpFax")), String.Empty, r("PreEmpFax").ToString)
            TxtPEmpContName.Text = If(IsDBNull(r("PreEmpContact")), String.Empty, r("PreEmpContact").ToString)
            TxtPEmpContNo.Text = If(IsDBNull(r("PreEmpContNo")), String.Empty, r("PreEmpContNo").ToString)
            'NEED ATTN
            TxtPEmpWFm.Text = DateTrimR(If(IsDBNull(r("PreEmpWorkFm")), String.Empty, r("PreEmpWorkFm").ToString))

            TxtPEmpWTo.Text = DateTrimR(If(IsDBNull(r("PreEmpWorkTo")), String.Empty, r("PreEmpWorkTo").ToString))

            TxtPEmpPosition.Text = If(IsDBNull(r("PreEmpPosition")), String.Empty, r("PreEmpPosition").ToString)
            TxtPEmpWage.Text = If(IsDBNull(r("PreEmpWage")), String.Empty, r("PreEmpWage").ToString)

            CBPEmpFMCSR.Checked = If(IsDBNull(r("PreEmpFMCSR")), False, r("PreEmpFMCSR"))
            CBPEmp49CFR.Checked = If(IsDBNull(r("PreEmp49CFR")), False, r("PreEmp49CFR"))

            TxtPEmpReasonL.Text = If(IsDBNull(r("PreEmpReasonL")), String.Empty, r("PreEmpReasonL").ToString)

        Next
    End Sub
    Private Sub BtnUpdPEmp_Click(sender As Object, e As EventArgs) Handles BtnUpdPEmp.Click
        ' Update all edited Pre Emp Info
        UpdatePEI()
    End Sub
    Private Sub UpdatePEI()
        SQL.AddParam("@PEmpID", TxtPEmpID.Text)

        SQL.AddParam("@PEName", TxtPEmpName.Text)
        SQL.AddParam("@PEadd1", TxtPEmpAdd1.Text)
        SQL.AddParam("@PEadd2", TxtPEmpAdd2.Text)
        SQL.AddParam("@PEcity", TxtPEmpCity.Text)
        SQL.AddParam("@PEState", CombPEmpState.Text)
        SQL.AddParam("@PEzip1", TxtPEmpZip1.Text)
        SQL.AddParam("@PEzip2", TxtPEmpZip2.Text)
        SQL.AddParam("@PEmail", TxtPEmpEmail.Text)
        SQL.AddParam("@PEPhone", TxtPEmpPhone.Text)
        SQL.AddParam("@PEFax", TxtPEmpFax.Text)
        SQL.AddParam("@PEcont", TxtPEmpContName.Text)
        SQL.AddParam("@PEcontNo", TxtPEmpContNo.Text)
        SQL.AddParam("@PEWFm", TxtPEmpWFm.Text)
        SQL.AddParam("@PEWTo", TxtPEmpWTo.Text)
        SQL.AddParam("@PEmpPo", TxtPEmpPosition.Text)
        SQL.AddParam("@PEmpWage", TxtPEmpWage.Text)
        SQL.AddParam("@PEmpFMCSR", CBPEmpFMCSR.Checked)
        SQL.AddParam("@PE49", CBPEmp49CFR.Checked)

        SQL.AddParam("@PEmpRsn", TxtPEmpReasonL.Text)

        SQL.ExecQuery("UPDATE PreEmpInfo " &
                      "SET PreEmpName=@PEName,PreEmpAdd1=@PEadd1,PreEmpAdd2=@PEadd2,PreEmpCity=@PEcity,PreEmpState=@PEState,PreEmpZip1=@PEzip1, " &
                      "PreEmpZip2=@PEzip2,PreEmpEmail=@PEmail,PreEmpPhone=@PEPhone,PreEmpFax=@PEFax,PreEmpContact=@PEcont,PreEmpContNo=@PEcontNo, " &
                      "PreEmpWorkFm=@PEWFm,PreEmpWorkTo=@PEWTo,PreEmpPosition=@PEmpPo,PreEmpWage=@PEmpWage,PreEmpFMCSR=@PEmpFMCSR,PreEmp49CFR=@PE49, " &
                      "PreEmpReasonL=@PEmpRsn " &
                      "WHERE PreEmpID=@PEmpID; " &
                      "BEGIN TRANSACTION; " &
                      "COMMIT;", True)

        'ERROR
        If SQL.HasException(True) Then Exit Sub

        MsgBox("Previous Employer Info Updated Successfully")
    End Sub

    ' ------------------DEL PREVIOUS EMPLOYER INFO APP -------------
    Private Sub BtnDelPEIView_Click(sender As Object, e As EventArgs) Handles BtnDelPEIView.Click
        '  DELETE PREVIOUS EMPLOYER INFO
        Dim tempstr1 As String = ""
        Dim tempstr2 As String = ""
        Dim tempdt1 As DateTime
        Dim tempdt2 As DateTime
        Dim aLoc As Integer

        aLoc = InStr(CLBPreEmpInfo.CheckedItems(0), vbTab)
        tempstr1 = Mid(CLBPreEmpInfo.CheckedItems(0), 1, aLoc - 1)

        aLoc = InStrRev(CLBPreEmpInfo.CheckedItems(0), "/")
        aLoc = aLoc + 5
        tempstr2 = Mid(CLBPreEmpInfo.CheckedItems(0), aLoc - 10, 10)

        tempdt1 = Convert.ToDateTime(tempstr2)
        tempdt2 = DateAdd("n", 1, tempdt1)

        If SQL.DBDT IsNot Nothing Then
            SQL.DBDT.Clear()
        End If

        SQL.AddParam("@PEmpName", tempstr1)
        SQL.AddParam("@tmpEmpCode", tempcode)
        SQL.AddParam("@Adate", tempdt1)
        SQL.AddParam("@Bdate", tempdt2)
        SQL.ExecQuery("DELETE FROM PreEmpInfo " &
                      "WHERE PreEmpName=@PEmpName AND EmpCode=@tmpEmpCode AND PreEmpWorkTo BETWEEN @Adate AND @Bdate;")

        MsgBox("Selected Previous Employer Info deleted")

        FetchViewPreEmp()
        'FetchViewPreLi()
        BtnDelPEIView.Enabled = False
        BtnEditPEIView.Enabled = False
    End Sub

    '##############################################
    '    PREVIOUS DRIVING EXPERIENCE DT: 1/31/2019
    '##############################################
    ' ------------------ADD PREVIOUS DRIVING EXPERIENCE APP -------------
    Private Sub BtnAddPDExView_Click(sender As Object, e As EventArgs) Handles BtnAddPDExView.Click
        GBViewApp.Hide()

        GBAddPreExp.SetBounds(260, 20, 700, 270)
        GBAddPreExp.Show()
        BtnUpdPEx.Hide()

        LblPExp.Text = "Add Previous Driving Experience" & tempFulName
        ClearPDrvExp()
        TxtTempIDPEx.Clear()
        TxtTempIDPEx.Text = tempcode
        'Temp Code Hide
        TxtTempIDPEx.Hide()
        TxtPDExpID.Hide()
        BtnSavePEx.Enabled = False
        BtnSaveAddPEx.Enabled = False
        BtnSavePEx.Show()
        BtnSaveAddPEx.Show()


    End Sub

    Private Sub BtnCancelPEx_Click(sender As Object, e As EventArgs) Handles BtnCancelPEx.Click
        GBAddPreExp.Hide()
        'btnAppProcess.Enabled = True
        BoxViewApp()
        FetchViewApp()
    End Sub

    Private Sub BtnSavePEx_Click(sender As Object, e As EventArgs) Handles BtnSavePEx.Click
        InsertPreExp()
        GBAddPreExp.Hide()
        'btnAppProcess.Enabled = True
        BoxViewApp()
        FetchViewApp()
    End Sub

    Private Sub BtnSaveAddPEx_Click(sender As Object, e As EventArgs) Handles BtnSaveAddPEx.Click
        InsertPreExp()
        ClearPDrvExp()
    End Sub

    Private Sub InsertPreExp()

        SQL.AddParam("@PEqptClass", CombEqptClass.Text)
        SQL.AddParam("@PEmpCode", TxtTempIDPEx.Text)
        SQL.AddParam("@PEqptType", TxtEqptType.Text)
        SQL.AddParam("@DrFm", TxtDrvFm.Text)
        SQL.AddParam("@DrTo", TxtDrvTo.Text)
        SQL.AddParam("@DrMiles", TxtDrvMiles.Text)

        SQL.ExecQuery("INSERT INTO PreDrvExp " &
                       "(EqptClass,EmpCode,EqptType,DriveFm,DriveTo,DriveMiles) " &
                       "VALUES (@PEqptClass,@PEmpCode,@PEqptType,@DrFm,@DrTo,@DrMiles); " &
                       "BEGIN TRANSACTION; " &
                       "COMMIT;", True)

        'ERROR
        If SQL.HasException(True) Then Exit Sub

        MsgBox("Previous Driving Experience Info Added Successfully")

    End Sub

    Private Sub CombEqptClass_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CombEqptClass.SelectedIndexChanged, TxtEqptType.TextChanged
        'BASIC VALIDATION
        If Not String.IsNullOrWhiteSpace(CombEqptClass.Text) AndAlso Not String.IsNullOrWhiteSpace(TxtEqptType.Text) Then

            BtnSavePEx.Enabled = True
            BtnSaveAddPEx.Enabled = True
        Else
            BtnSavePEx.Enabled = False
            BtnSaveAddPEx.Enabled = False
        End If
    End Sub

    Private Sub DTPDrvFm_ValueChanged(sender As Object, e As EventArgs) Handles DTPDrvFm.ValueChanged

        TxtDrvFm.Text = DateTrimR(DTPDrvFm.Value)
    End Sub
    Private Sub DTPDrvTo_ValueChanged(sender As Object, e As EventArgs) Handles DTPDrvTo.ValueChanged

        TxtDrvTo.Text = DateTrimR(DTPDrvTo.Value)
    End Sub

    Private Sub ClearPDrvExp()

        CombEqptClass.ResetText()
        'txtTempIDPEx.Clear()
        TxtEqptType.Clear()
        DTPDrvFm.ResetText()
        DTPDrvTo.ResetText()
        TxtDrvFm.Clear()
        TxtDrvTo.Clear()
        TxtDrvMiles.Clear()
    End Sub

    ' ------------------EDIT PREVIOUS DRIVING EXPERIENCE APP -------------
    Private Sub CLBPreDrvEx_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CLBPreDrvEx.SelectedIndexChanged, CLBPreDrvEx.Click, CLBPreDrvEx.DoubleClick
        If CLBPreDrvEx.CheckedItems.Count() = 0 Then
            BtnDelPDExView.Enabled = False
            BtnEditPDExView.Enabled = False
        Else
            BtnDelPDExView.Enabled = True
            BtnEditPDExView.Enabled = True
        End If
    End Sub

    Private Sub BtnEditPDExView_Click(sender As Object, e As EventArgs) Handles BtnEditPDExView.Click
        '   EDIT PREVIOUS EXPERIENCE 

        GBViewApp.Hide()

        TxtTempIDPEx.Clear()
        'Temp Code Hide
        TxtTempIDPEx.Hide()
        TxtPDExpID.Hide()

        GBAddPreExp.SetBounds(260, 20, 700, 270)
        GBAddPreExp.Show()
        BtnSavePEx.Hide()
        BtnSaveAddPEx.Hide()
        BtnUpdPEx.SetBounds(145, 225, 81, 24)
        BtnUpdPEx.Show()
        LblPExp.Text = "Edit Previous Driving Experience" & tempFulName
        BtnSavePEx.Enabled = False
        BtnSaveAddPEx.Enabled = False
        '*** FETCH PREVIOUS EXPERIENCE INFO

        ClearPDrvExp()

        FetchPEx()
    End Sub

    Private Sub FetchPEx()
        Dim tempstr1 As String = ""
        Dim tempstr2 As String = ""
        Dim aLoc As Integer
        Dim tempdt1 As DateTime
        Dim tempdt2 As DateTime

        aLoc = InStr(CLBPreDrvEx.CheckedItems(0), vbTab)
        tempstr1 = Mid(CLBPreDrvEx.CheckedItems(0), 1, aLoc - 1)
        '
        aLoc = InStrRev(CLBPreDrvEx.CheckedItems(0), "/")
        aLoc = aLoc + 5
        tempstr2 = Mid(CLBPreDrvEx.CheckedItems(0), aLoc - 10, 10)

        tempdt1 = Convert.ToDateTime(tempstr2)
        tempdt2 = DateAdd("n", 1, tempdt1)

        If SQL.DBDT IsNot Nothing Then
            SQL.DBDT.Clear()
        End If

        SQL.AddParam("@EClass", tempstr1)
        SQL.AddParam("@tmpEmpCode", tempcode)
        SQL.AddParam("@Adate", tempdt1)
        SQL.AddParam("@Bdate", tempdt2)
        SQL.ExecQuery("SELECT TOP 1 * FROM PreDrvExp " &
                      "WHERE EqptClass=@EClass AND EmpCode=@tmpEmpCode AND DriveTo BETWEEN @Adate AND @Bdate;")

        If SQL.RecordCount < 1 Then Exit Sub
        TxtTempIDPEx.Text = tempcode
        For Each r As DataRow In SQL.DBDT.Rows
            TxtPDExpID.Text = r("PreDrvExpID")
            'TxtTempIDPEx.Text =
            CombEqptClass.Text = r("EqptClass")
            TxtEqptType.Text = If(IsDBNull(r("EqptType")), String.Empty, r("EqptType").ToString)
            TxtDrvFm.Text = DateTrimR(If(IsDBNull(r("DriveFm")), String.Empty, r("DriveFm").ToString))
            TxtDrvTo.Text = DateTrimR(If(IsDBNull(r("DriveTo")), String.Empty, r("DriveTo").ToString))
            TxtDrvMiles.Text = If(IsDBNull(r("DriveMiles")), String.Empty, r("DriveMiles").ToString)

        Next
    End Sub

    Private Sub BtnUpdPEx_Click(sender As Object, e As EventArgs) Handles BtnUpdPEx.Click

        SQL.AddParam("@PDrvExID", TxtPDExpID.Text)

        SQL.AddParam("@PEqptClass", CombEqptClass.Text)

        SQL.AddParam("@PEqptType", TxtEqptType.Text)
        SQL.AddParam("@DrFm", TxtDrvFm.Text)
        SQL.AddParam("@DrTo", TxtDrvTo.Text)
        SQL.AddParam("@DrMiles", TxtDrvMiles.Text)

        SQL.ExecQuery("UPDATE PreDrvExp " &
                     "SET EqptClass=@PEqptClass,EqptType=@PEqptType,DriveFm=@DrFm,DriveTo=@DrTo,DriveMiles=@DrMiles " &
                     "WHERE PreDrvExpID=@PDrvExID; " &
                     "BEGIN TRANSACTION; " &
                     "COMMIT;", True)

        'ERROR
        If SQL.HasException(True) Then Exit Sub

        MsgBox("Previous Driving Experience Updated Successfully")

    End Sub

    ' ------------------DEL PREVIOUS DRIVING EXPERIENCE APP -------------
    Private Sub BtnDelPDExView_Click(sender As Object, e As EventArgs) Handles BtnDelPDExView.Click
        '  DELETE PREVIOUS EXPERIENCE INFO
        Dim tempstr1 As String = ""
        Dim tempstr2 As String = ""
        Dim tempdt1 As DateTime
        Dim tempdt2 As DateTime
        Dim aLoc As Integer

        aLoc = InStr(CLBPreDrvEx.CheckedItems(0), vbTab)
        tempstr1 = Mid(CLBPreDrvEx.CheckedItems(0), 1, aLoc - 1)

        aLoc = InStrRev(CLBPreDrvEx.CheckedItems(0), "/")
        aLoc = aLoc + 5
        tempstr2 = Mid(CLBPreDrvEx.CheckedItems(0), aLoc - 10, 10)

        tempdt1 = Convert.ToDateTime(tempstr2)
        tempdt2 = DateAdd("n", 1, tempdt1)

        If SQL.DBDT IsNot Nothing Then
            SQL.DBDT.Clear()
        End If

        SQL.AddParam("@EClass", tempstr1)
        SQL.AddParam("@tmpEmpCode", tempcode)
        SQL.AddParam("@Adate", tempdt1)
        SQL.AddParam("@Bdate", tempdt2)
        SQL.ExecQuery("DELETE FROM PreDrvExp " &
                      "WHERE EqptClass=@EClass AND EmpCode=@tmpEmpCode AND DriveTo BETWEEN @Adate AND @Bdate;")

        MsgBox("Selected Previous Experience deleted")

        FetchViewPreExp()

        BtnDelPDExView.Enabled = False
        BtnEditPDExView.Enabled = False
    End Sub



    '##############################################
    '    PREVIOUS VIOLATION INFO DT: 1/31/2019
    '############################################
    ' ------------------ADD PREVIOUS VIOLATION INFO APP -------------
    Private Sub BtnAddVHView_Click(sender As Object, e As EventArgs) Handles BtnAddVHView.Click
        GBViewApp.Hide()

        GBAddPreVio.SetBounds(260, 20, 700, 240)
        GBAddPreVio.Show()
        BtnUpdVI.Hide()

        LblPVI.Text = "Add Previous Violation Info" & tempFulName
        ClearPreVio()
        TxtTmpViID.Clear()
        TxtTmpViID.Text = tempcode
        TxtTmpViID.Hide()
        TxtPVID.Hide()
        BtnSaveVI.Enabled = False
        BtnSaveAddVI.Enabled = False
        BtnSaveVI.Show()
        BtnSaveAddVI.Show()
    End Sub


    Private Sub BtnCancelVI_Click(sender As Object, e As EventArgs) Handles BtnCancelVI.Click
        GBAddPreVio.Hide()
        'btnAppProcess.Enabled = True
        BoxViewApp()
        FetchViewApp()
    End Sub

    Private Sub BtnSaveVI_Click(sender As Object, e As EventArgs) Handles BtnSaveVI.Click
        InsertPreVio()
        GBAddPreVio.Hide()
        'btnAppProcess.Enabled = True
        BoxViewApp()
        FetchViewApp()
    End Sub

    Private Sub InsertPreVio()

        SQL.AddParam("@Vcharge", TxtCharge.Text)
        SQL.AddParam("@PEmpCode", TxtTmpViID.Text)
        SQL.AddParam("@Vloc", TxtVioLoc.Text)
        SQL.AddParam("@Viodt", TxtVioDt.Text)
        SQL.AddParam("@VPenalty", TxtPenalty.Text)

        SQL.ExecQuery("INSERT INTO Violations " &
                       "(Charge,EmpCode,Loc,VioDate,Penalty) " &
                       "VALUES (@Vcharge,@PEmpCode,@Vloc,@Viodt,@Vpenalty); " &
                       "BEGIN TRANSACTION; " &
                       "COMMIT;", True)

        'ERROR
        If SQL.HasException(True) Then Exit Sub

        MsgBox("Previous Violation Info Added Successfully")
    End Sub

    Private Sub TxtCharge_TextChanged(sender As Object, e As EventArgs) Handles TxtCharge.TextChanged, TxtVioLoc.TextChanged, TxtAccNature.TextChanged
        'BASIC VALIDATION
        If Not String.IsNullOrWhiteSpace(TxtCharge.Text) AndAlso Not String.IsNullOrWhiteSpace(TxtVioLoc.Text) Then

            BtnSaveVI.Enabled = True
            BtnSaveAddVI.Enabled = True
        Else
            BtnSaveVI.Enabled = False
            BtnSaveAddVI.Enabled = False
        End If
    End Sub

    Private Sub DTPVioDt_ValueChanged(sender As Object, e As EventArgs) Handles DTPVioDt.ValueChanged, DTPAccDt.ValueChanged

        TxtVioDt.Text = DateTrimR(DTPVioDt.Value)
    End Sub

    Private Sub BtnSaveAddVI_Click(sender As Object, e As EventArgs) Handles BtnSaveAddVI.Click
        InsertPreVio()
        ClearPreVio()
    End Sub

    Private Sub ClearPreVio()

        TxtCharge.Clear()
        TxtVioLoc.Clear()
        DTPVioDt.ResetText()
        TxtVioDt.Clear()
        TxtPenalty.Clear()

    End Sub
    ' ------------------EDIT PREVIOUS VIOLATION INFO APP -------------
    Private Sub CLBViolation_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CLBViolation.SelectedIndexChanged, CLBViolation.Click, CLBViolation.DoubleClick
        If CLBViolation.CheckedItems.Count() = 0 Then
            BtnDelVHView.Enabled = False
            BtnEditVHView.Enabled = False
        Else
            BtnDelVHView.Enabled = True
            BtnEditVHView.Enabled = True
        End If
    End Sub

    Private Sub BtnEditVHView_Click(sender As Object, e As EventArgs) Handles BtnEditVHView.Click
        '   EDIT PREVIOUS VIOLATION                                                      

        GBViewApp.Hide()

        TxtTmpViID.Clear()
        'Temp Code Hide
        TxtTmpViID.Hide()
        TxtPVID.Hide()
        GBAddPreVio.SetBounds(260, 20, 700, 240)
        GBAddPreVio.Show()
        BtnSaveVI.Hide()
        BtnSaveAddVI.Hide()
        BtnUpdVI.SetBounds(136, 195, 81, 24)
        BtnUpdVI.Show()
        LblPVI.Text = "Edit Previous Violation" & tempFulName
        BtnSaveVI.Enabled = False
        BtnSaveAddVI.Enabled = False
        '*** FETCH PREVIOUS VIOLATION INFO

        ClearPreVio()

        FetchPreVio()

    End Sub

    Private Sub FetchPreVio()
        Dim tempstr1 As String = ""
        Dim tempstr2 As String = ""
        Dim aLoc As Integer
        Dim tempdt1 As DateTime
        Dim tempdt2 As DateTime
        '######
        '#######
        aLoc = InStr(CLBViolation.CheckedItems(0), vbTab)
        tempstr1 = Mid(CLBViolation.CheckedItems(0), 1, aLoc - 1)
        '
        aLoc = InStrRev(CLBViolation.CheckedItems(0), "/")
        aLoc = aLoc + 5
        tempstr2 = Mid(CLBViolation.CheckedItems(0), aLoc - 10, 10)

        tempdt1 = Convert.ToDateTime(tempstr2)

        tempdt2 = DateAdd("n", 1, tempdt1)

        If SQL.DBDT IsNot Nothing Then
            SQL.DBDT.Clear()
        End If

        SQL.AddParam("@TCharge", tempstr1)
        SQL.AddParam("@tmpEmpCode", tempcode)
        SQL.AddParam("@Adate", tempdt1)
        SQL.AddParam("@Bdate", tempdt2)
        SQL.ExecQuery("SELECT TOP 1 * FROM Violations " &
                      "WHERE Charge=@TCharge AND EmpCode=@tmpEmpCode AND VioDate BETWEEN @Adate AND @Bdate;")
        '
        If SQL.RecordCount < 1 Then Exit Sub
        TxtTmpViID.Text = tempcode
        For Each r As DataRow In SQL.DBDT.Rows
            TxtPVID.Text = r("ViolationID")
            'TxtTmpViID.Text = r("ViolationID")
            TxtCharge.Text = r("Charge")
            TxtVioLoc.Text = If(IsDBNull(r("Loc")), String.Empty, r("Loc").ToString)
            TxtVioDt.Text = DateTrimR(If(IsDBNull(r("VioDate")), String.Empty, r("VioDate").ToString))
            TxtPenalty.Text = If(IsDBNull(r("Penalty")), String.Empty, r("Penalty").ToString)

        Next
    End Sub

    Private Sub BtnUpdVI_Click(sender As Object, e As EventArgs) Handles BtnUpdVI.Click

        SQL.AddParam("@VioID", TxtPVID.Text)

        SQL.AddParam("@TCharge", TxtCharge.Text)

        SQL.AddParam("@VLoc", TxtVioLoc.Text)
        SQL.AddParam("@Vdt", TxtVioDt.Text)
        SQL.AddParam("@Tpenalty", TxtPenalty.Text)

        SQL.ExecQuery("UPDATE Violations " &
                     "SET Charge=@TCharge,Loc=@VLoc,VioDate=@Vdt,Penalty=@Tpenalty " &
                     "WHERE ViolationID=@VioID; " &
                     "BEGIN TRANSACTION; " &
                     "COMMIT;", True)
        'ERROR
        If SQL.HasException(True) Then Exit Sub

        MsgBox("Previous Violation Updated Successfully")
    End Sub
    ' ------------------DEL PREVIOUS VIOLATION INFO APP -------------
    Private Sub BtnDelVHView_Click(sender As Object, e As EventArgs) Handles BtnDelVHView.Click
        '  DELETE PREVIOUS VIOLATION INFO
        Dim tempstr1 As String = ""
        Dim tempstr2 As String = ""
        Dim tempdt1 As DateTime
        Dim tempdt2 As DateTime
        Dim aLoc As Integer

        aLoc = InStr(CLBViolation.CheckedItems(0), vbTab)
        tempstr1 = Mid(CLBViolation.CheckedItems(0), 1, aLoc - 1)

        aLoc = InStrRev(CLBViolation.CheckedItems(0), "/")
        aLoc = aLoc + 5
        tempstr2 = Mid(CLBViolation.CheckedItems(0), aLoc - 10, 10)

        tempdt1 = Convert.ToDateTime(tempstr2)

        tempdt2 = DateAdd("n", 1, tempdt1)

        If SQL.DBDT IsNot Nothing Then
            SQL.DBDT.Clear()
        End If

        SQL.AddParam("@TCharge", tempstr1)
        SQL.AddParam("@tmpEmpCode", tempcode)
        SQL.AddParam("@Adate", tempdt1)
        SQL.AddParam("@Bdate", tempdt2)
        SQL.ExecQuery("DELETE FROM Violations " &
                      "WHERE Charge=@TCharge AND EmpCode=@tmpEmpCode AND VioDate BETWEEN @Adate AND @Bdate;")

        MsgBox("Selected Previous Violation deleted")

        FetchViewPreVio()

        BtnDelVHView.Enabled = False
        BtnEditVHView.Enabled = False
    End Sub



    '#####################################
    '   PREVIOUS ACCIDENT HISTORY 2/1/19
    '#####################################
    ' ------------------ADD PREVIOUS ACCIDENT HISTORY APP -------------
    Private Sub BtnAddAHView_Click(sender As Object, e As EventArgs) Handles BtnAddAHView.Click
        GBViewApp.Hide()

        GBAddPreAcc.SetBounds(260, 20, 700, 230)
        GBAddPreAcc.Show()
        BtnUpdPAH.Hide()

        LblPAI.Text = "Add Previous Accident Info" & tempFulName
        ClearPreAcc()
        TxtTempAccID.Clear()
        TxtTempAccID.Text = tempcode
        TxtTempAccID.Hide()
        TxtPAccID.Hide()
        BtnSavePAH.Enabled = False
        BtnSaveAddPAH.Enabled = False
        BtnSavePAH.Show()
        BtnSaveAddPAH.Show()

    End Sub

    Private Sub BtnCancelPAH_Click(sender As Object, e As EventArgs) Handles BtnCancelPAH.Click
        GBAddPreAcc.Hide()
        'btnAppProcess.Enabled = True
        BoxViewApp()
        FetchViewApp()
    End Sub

    Private Sub DTPAccDt_ValueChanged(sender As Object, e As EventArgs) Handles DTPAccDt.ValueChanged

        TxtAccDt.Text = DateTrimR(DTPAccDt.Value)
    End Sub

    Private Sub BtnSavePAH_Click(sender As Object, e As EventArgs) Handles BtnSavePAH.Click
        InsertPreAcc()
        GBAddPreAcc.Hide()
        'btnAppProcess.Enabled = True
        BoxViewApp()
        FetchViewApp()
    End Sub

    Private Sub BtnSaveAddPAH_Click(sender As Object, e As EventArgs) Handles BtnSaveAddPAH.Click
        InsertPreAcc()
        ClearPreAcc()
    End Sub

    Private Sub ClearPreAcc()

        TxtAccNature.Clear()
        DTPAccDt.ResetText()
        TxtAccDt.Clear()
        TxtNoFatality.Clear()
        TxtNoInjury.Clear()

    End Sub

    Private Sub txtAccNature_TextChanged(sender As Object, e As EventArgs) Handles TxtAccNature.TextChanged, TxtAccDt.TextChanged
        'BASIC VALIDATION
        If Not String.IsNullOrWhiteSpace(TxtAccNature.Text) AndAlso Not String.IsNullOrWhiteSpace(TxtAccDt.Text) Then

            BtnSavePAH.Enabled = True
            BtnSaveAddPAH.Enabled = True
        Else
            BtnSavePAH.Enabled = False
            BtnSaveAddPAH.Enabled = False
        End If
    End Sub

    Private Sub InsertPreAcc()
        SQL.AddParam("@PAccNature", TxtAccNature.Text)
        SQL.AddParam("@PEmpCode", TxtTempAccID.Text)
        SQL.AddParam("@PAccDt", TxtAccDt.Text)
        SQL.AddParam("@NumFatality", TxtNoFatality.Text)
        SQL.AddParam("@NumInjury", TxtNoInjury.Text)

        SQL.ExecQuery("INSERT INTO PreAccident " &
                       "(AccNature,EmpCode,AccDate,NoFatality,NoInjury) " &
                       "VALUES (@PAccNature,@PEmpCode,@PAccDt,@NumFatality,@NumInjury); " &
                       "BEGIN TRANSACTION; " &
                       "COMMIT;", True)

        'ERROR
        If SQL.HasException(True) Then Exit Sub

        MsgBox("Previous Accident Info Added Successfully")
    End Sub

    ' ------------------EDIT PREVIOUS ACCIDENT HISTORY APP -------------
    Private Sub CLBAccident_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CLBAccident.SelectedIndexChanged, CLBAccident.Click, CLBAccident.DoubleClick
        If CLBAccident.CheckedItems.Count() = 0 Then
            BtnDelAHView.Enabled = False
            BtnEditAHView.Enabled = False
        Else
            BtnDelAHView.Enabled = True
            BtnEditAHView.Enabled = True
        End If
    End Sub

    Private Sub BtnEditAHView_Click(sender As Object, e As EventArgs) Handles BtnEditAHView.Click
        GBViewApp.Hide()

        TxtTempAccID.Clear()
        '
        TxtTempAccID.Hide()
        TxtPAccID.Hide()

        GBAddPreAcc.SetBounds(260, 20, 700, 230)
        GBAddPreAcc.Show()
        BtnSavePAH.Hide()
        BtnSaveAddPAH.Hide()
        BtnUpdPAH.SetBounds(136, 195, 81, 24)
        BtnUpdPAH.Show()
        LblPAI.Text = "Edit Previous Accident Info" & tempFulName
        BtnSavePAH.Enabled = False
        BtnSaveAddPAH.Enabled = False
        '*** FETCH PREVIOUS ACCIDENT INFO

        ClearPreAcc()

        FetchPreAcc()
    End Sub

    Private Sub FetchPreAcc()
        Dim tempstr1 As String = ""
        Dim tempstr2 As String = ""
        Dim aLoc As Integer
        Dim tempdt1 As DateTime
        Dim tempdt2 As DateTime

        aLoc = InStr(CLBAccident.CheckedItems(0), vbTab)
        tempstr1 = Mid(CLBAccident.CheckedItems(0), 1, aLoc - 1)

        aLoc = InStrRev(CLBAccident.CheckedItems(0), "/")
        aLoc = aLoc + 5
        tempstr2 = Mid(CLBAccident.CheckedItems(0), aLoc - 10, 10)

        tempdt1 = Convert.ToDateTime(tempstr2)

        tempdt2 = DateAdd("n", 1, tempdt1)

        If SQL.DBDT IsNot Nothing Then
            SQL.DBDT.Clear()
        End If

        SQL.AddParam("@PANature", tempstr1)
        SQL.AddParam("@tmpEmpCode", tempcode)
        SQL.AddParam("@Adate", tempdt1)
        SQL.AddParam("@Bdate", tempdt2)
        SQL.ExecQuery("SELECT TOP 1 * FROM PreAccident " &
                      "WHERE AccNature=@PANature AND EmpCode=@tmpEmpCode AND AccDate BETWEEN @Adate AND @Bdate;")
        '
        If SQL.RecordCount < 1 Then Exit Sub
        TxtTempAccID.Text = tempcode
        For Each r As DataRow In SQL.DBDT.Rows

            TxtPAccID.Text = r("PAccidentID")
            TxtAccNature.Text = r("AccNature")
            TxtAccDt.Text = DateTrimR(If(IsDBNull(r("AccDate")), String.Empty, r("AccDate").ToString))
            TxtNoFatality.Text = If(IsDBNull(r("NoFatality")), String.Empty, r("NoFatality").ToString)
            TxtNoInjury.Text = If(IsDBNull(r("NoInjury")), String.Empty, r("NoInjury").ToString)

        Next
    End Sub

    Private Sub BtnUpdPAH_Click(sender As Object, e As EventArgs) Handles BtnUpdPAH.Click
        SQL.AddParam("@PAccID", TxtPAccID.Text)

        SQL.AddParam("@ANature", TxtAccNature.Text)

        SQL.AddParam("@ADate", TxtAccDt.Text)
        SQL.AddParam("@NFatality", TxtNoFatality.Text)
        SQL.AddParam("@NInjury", TxtNoInjury.Text)

        SQL.ExecQuery("UPDATE PreAccident " &
                     "SET AccNature=@ANature,AccDate=@ADate,NoFatality=@NFatality,NoInjury=@NInjury " &
                     "WHERE PAccidentID=@PAccID; " &
                     "BEGIN TRANSACTION; " &
                     "COMMIT;", True)
        'ERROR
        If SQL.HasException(True) Then Exit Sub

        MsgBox("Previous Accident Info Updated Successfully")
    End Sub
    ' ------------------DEL PREVIOUS ACCIDENT HISTORY APP -------------
    Private Sub BtnDelAHView_Click(sender As Object, e As EventArgs) Handles BtnDelAHView.Click
        '  DELETE PREVIOUS ACCIDENT INFO
        Dim tempstr1 As String = ""
        Dim tempstr2 As String = ""


        Dim tempdt As DateTime
        Dim tempdt2 As DateTime
        Dim aLoc As Integer

        aLoc = InStr(CLBAccident.CheckedItems(0), vbTab)
        tempstr1 = Mid(CLBAccident.CheckedItems(0), 1, aLoc - 1)

        aLoc = 0
        aLoc = InStrRev(CLBAccident.CheckedItems(0), "/")
        aLoc = aLoc + 5
        tempstr2 = Mid(CLBAccident.CheckedItems(0), aLoc - 10, 10)

        tempdt = Convert.ToDateTime(tempstr2)

        tempdt2 = DateAdd("n", 1, tempdt)

        If SQL.DBDT IsNot Nothing Then
            SQL.DBDT.Clear()
        End If

        SQL.AddParam("@ANature", tempstr1)
        SQL.AddParam("@tmpEmpCode", tempcode)
        SQL.AddParam("@Adate", tempdt)
        SQL.AddParam("@Bdate", tempdt2)

        SQL.ExecQuery("DELETE FROM PreAccident " &
                      "WHERE AccNature=@ANature AND EmpCode=@tmpEmpCode AND AccDate BETWEEN @Adate AND @Bdate;")

        MsgBox("Selected Previous Accident info deleted")

        FetchViewPreAcc()

        BtnDelAHView.Enabled = False
        BtnEditAHView.Enabled = False
    End Sub


    '##############################################
    '     PREVIOUS ADDRESS INFO DT: 2/1/2019
    '##############################################
    ' ------------------ADD PREVIOUS ACCIDENT HISTORY APP -------------
    Private Sub BtnAddPAView_Click(sender As Object, e As EventArgs) Handles BtnAddPAView.Click
        GBViewApp.Hide()

        GBAddPAddr.SetBounds(260, 20, 700, 285)
        GBAddPAddr.Show()
        BtnUpdPAI.Hide()

        LblPAddr.Text = "Add Previous Address Info" & tempFulName
        ClearPreAddr()
        TxtPTempID.Clear()
        TxtPTempID.Text = tempcode
        TxtPTempID.Hide()
        TxtPAddrID.Hide()
        BtnSavePAI.Enabled = False
        BtnSaveAddPAI.Enabled = False
        BtnSavePAI.Show()
        BtnSaveAddPAI.Show()
    End Sub

    Private Sub BtnCancelPAI_Click(sender As Object, e As EventArgs) Handles BtnCancelPAI.Click
        GBAddPAddr.Hide()
        'btnAppProcess.Enabled = True
        BoxViewApp()
        FetchViewApp()
    End Sub

    Private Sub BtnSavePAI_Click(sender As Object, e As EventArgs) Handles BtnSavePAI.Click
        InsertPreAddr()
        GBAddPAddr.Hide()
        'btnAppProcess.Enabled = True
        BoxViewApp()
        FetchViewApp()
    End Sub

    Private Sub BtnSaveAddPAI_Click(sender As Object, e As EventArgs) Handles BtnSaveAddPAI.Click
        InsertPreAddr()
        ClearPreAddr()
    End Sub

    Private Sub InsertPreAddr()
        SQL.AddParam("@PAAddr1", TxtPAddr1.Text)
        SQL.AddParam("@PAAddr2", TxtPAddr2.Text)
        SQL.AddParam("@PAEmpCode", TxtPTempID.Text)
        SQL.AddParam("@PACity", TxtPCity.Text)
        SQL.AddParam("@PAState", CombPState.Text)
        SQL.AddParam("@PAZip", TxtPZip.Text)
        SQL.AddParam("@PAHlong", TxtPHlong.Text)

        SQL.ExecQuery("INSERT INTO PreAddr " &
                       "(PAddr1,PAddr2,EmpCode,PCity,PState,PZip,Hlong) " &
                       "VALUES (@PAAddr1,@PAAddr2,@PAEmpCode,@PACity,@PAState,@PAZip,@PAHlong); " &
                       "BEGIN TRANSACTION; " &
                       "COMMIT;", True)

        'ERROR
        If SQL.HasException(True) Then Exit Sub

        MsgBox("Previous Address Info Added Successfully")
    End Sub

    Private Sub txtPAddr1_TextChanged(sender As Object, e As EventArgs) Handles TxtPAddr1.TextChanged, TxtPCity.TextChanged, CombPState.TextChanged
        'BASIC VALIDATION
        If Not String.IsNullOrWhiteSpace(TxtPAddr1.Text) AndAlso Not String.IsNullOrWhiteSpace(TxtPCity.Text) AndAlso Not String.IsNullOrWhiteSpace(CombPState.Text) Then

            BtnSavePAI.Enabled = True
            BtnSaveAddPAI.Enabled = True
        Else
            BtnSavePAI.Enabled = False
            BtnSaveAddPAI.Enabled = False
        End If
    End Sub

    Private Sub ClearPreAddr()

        TxtPAddr1.Clear()
        TxtPAddr2.Clear()
        ' txtPTempID.Text)
        TxtPCity.Clear()
        CombPState.ResetText()
        TxtPZip.Clear()
        TxtPHlong.Clear()

    End Sub


    ' ------------------EDIT PREVIOUS ACCIDENT HISTORY APP -------------
    Private Sub CLBPreAddr_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CLBPreAddr.SelectedIndexChanged, CLBPreAddr.Click, CLBPreAddr.DoubleClick
        If CLBPreAddr.CheckedItems.Count() = 0 Then
            BtnDelPAView.Enabled = False
            BtnEditPAView.Enabled = False
        Else
            BtnDelPAView.Enabled = True
            BtnEditPAView.Enabled = True
        End If
    End Sub
    Private Sub BtnEditPAView_Click(sender As Object, e As EventArgs) Handles BtnEditPAView.Click
        GBViewApp.Hide()

        TxtPTempID.Clear()
        '
        TxtPTempID.Hide()
        TxtPAddrID.Hide()
        GBAddPAddr.SetBounds(260, 20, 700, 285)
        GBAddPAddr.Show()
        BtnSavePAI.Hide()
        BtnSaveAddPAI.Hide()

        BtnUpdPAI.SetBounds(133, 243, 81, 24)
        BtnUpdPAI.Show()
        LblPAddr.Text = "Edit Previous Address Info" & tempFulName

        '*** FETCH PREVIOUS ADDRESS INFO

        ClearPreAddr()

        FetchPreAddr()
    End Sub

    Private Sub FetchPreAddr()
        Dim tempstr1 As String = ""
        ' Dim tempstr2 As String = ""
        Dim aLoc As Integer

        aLoc = InStr(CLBPreAddr.CheckedItems(0), vbTab)
        tempstr1 = Mid(CLBPreAddr.CheckedItems(0), 1, aLoc - 1)

        If SQL.DBDT IsNot Nothing Then
            SQL.DBDT.Clear()
        End If

        SQL.AddParam("@PAAddr1", tempstr1)
        SQL.AddParam("@tmpEmpCode", tempcode)

        SQL.ExecQuery("SELECT TOP 1 * FROM PreAddr " &
                      "WHERE PAddr1=@PAAddr1 AND EmpCode=@tmpEmpCode;")
        '
        If SQL.RecordCount < 1 Then Exit Sub
        TxtPTempID.Text = tempcode
        For Each r As DataRow In SQL.DBDT.Rows

            TxtPAddrID.Text = r("PAddrID")

            TxtPAddr1.Text = r("PAddr1")
            TxtPAddr2.Text = If(IsDBNull(r("PAddr2")), String.Empty, r("PAddr2").ToString)
            TxtPCity.Text = If(IsDBNull(r("PCity")), String.Empty, r("PCity").ToString)
            CombPState.Text = If(IsDBNull(r("PState")), String.Empty, r("PState").ToString)
            TxtPZip.Text = If(IsDBNull(r("PZip")), String.Empty, r("PZip").ToString)
            TxtPHlong.Text = If(IsDBNull(r("Hlong")), String.Empty, r("Hlong").ToString)

        Next
    End Sub
    Private Sub BtnUpdPAI_Click(sender As Object, e As EventArgs) Handles BtnUpdPAI.Click
        ' Update all edited Pre Address Info
        UpdatePreAddr()
    End Sub

    Private Sub UpdatePreAddr()
        SQL.AddParam("@PTempID", TxtPAddrID.Text)
        SQL.AddParam("@Addr1", TxtPAddr1.Text)
        SQL.AddParam("@Addr2", TxtPAddr2.Text)
        SQL.AddParam("@PCity", TxtPCity.Text)
        SQL.AddParam("@PState", CombPState.Text)
        SQL.AddParam("@PZip", TxtPZip.Text)
        SQL.AddParam("@PHlong", TxtPHlong.Text)

        SQL.ExecQuery("UPDATE PreAddr " &
                     "SET PAddr1=@Addr1,PAddr2=@Addr2,PCity=@Pcity,PState=@PState,PZip=@PZip,Hlong=@PHlong " &
                     "WHERE PAddrID=@PTempID; " &
                     "BEGIN TRANSACTION; " &
                     "COMMIT;", True)
        'ERROR
        If SQL.HasException(True) Then Exit Sub

        MsgBox("Previous Address Info Updated Successfully")
    End Sub
    ' ------------------DEL PREVIOUS ACCIDENT HISTORY APP -------------
    Private Sub BtnDelPAView_Click(sender As Object, e As EventArgs) Handles BtnDelPAView.Click

        Dim tempstr1 As String = ""
        ' Dim tempstr2 As String = ""
        Dim aLoc As Integer

        aLoc = InStr(CLBPreAddr.CheckedItems(0), vbTab)
        tempstr1 = Mid(CLBPreAddr.CheckedItems(0), 1, aLoc - 1)

        If SQL.DBDT IsNot Nothing Then
            SQL.DBDT.Clear()
        End If

        SQL.AddParam("@PAAddr1", tempstr1)
        SQL.AddParam("@tmpEmpCode", tempcode)

        SQL.ExecQuery("DELETE FROM PreAddr " &
                      "WHERE PAddr1=@PAAddr1 AND EmpCode=@tmpEmpCode ;")

        MsgBox("Selected Previous Address info deleted")

        FetchViewPreAddr()

        BtnDelPAView.Enabled = False
        BtnEditPAView.Enabled = False
    End Sub

    '#######################################################
    '       APPLICANT ATTACHMENTS
    '#########################################################
    ' ------------------ADD APPLICANT ATTACHMENT -------------
    Private Sub BtnAddAAView_Click(sender As Object, e As EventArgs) Handles BtnAddAAView.Click
        GBViewApp.Hide()

        GBAddAppAtt.SetBounds(260, 20, 700, 210)
        GBAddAppAtt.Show()
        BtnUpdAA.Hide()

        LblAAA.Text = "Add Applicant Attachment" & tempFulName
        ClearFileLoc()
        TxtTempID.Clear()
        TxtTempID.Text = tempcode
        TxtTempID.Hide()
        TxtFileID.Hide()

        LblFileName.Show()

        BtnBrowseAA.Show()
        BtnSaveAA.Show()
        BtnSaveAddAA.Show()
        BtnSaveAA.Enabled = False
        BtnSaveAddAA.Enabled = False
    End Sub

    Public Flocstr As String

    Private Sub BtnBrowseAA_Click(sender As Object, e As EventArgs) Handles BtnBrowseAA.Click
        Dim opf As OpenFileDialog = New OpenFileDialog

        opf.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        opf.Filter = "PDF Files (*.pdf) |*.pdf|All Files (*.*)|*.*"

        If (opf.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK) Then
            Dim FileName As String = opf.FileName
            Dim tempstr As String = ""

            'TxtFilePathFm.Text = FileName

            'GET FILE NAME
            Dim sLoc As Integer
            Flocstr = FileName
            sLoc = InStrRev(FileName, "\")
            tempstr = Mid(FileName, sLoc + 1)
            LblFileName.Text = tempstr

        Else
            LblFileName.Text = "No file chosen"
            'TxtFilePathFm.Clear()

        End If

    End Sub

    Private Sub BtnCancelAA_Click(sender As Object, e As EventArgs) Handles BtnCancelAA.Click
        GBAddAppAtt.Hide()
        'btnAppProcess.Enabled = True
        BoxViewApp()
        FetchViewApp()
    End Sub

    Private Sub DTPAttDate_ValueChanged(sender As Object, e As EventArgs) Handles DTPAttDate.ValueChanged

        TxtAttDate.Text = DateTrimR(DTPAttDate.Value)
    End Sub
    'Public str11 As String
    Public str22 As String

    Private Sub BtnSaveAA_Click(sender As Object, e As EventArgs) Handles BtnSaveAA.Click
        '##### FILE COPY TO NETWORK
        CopyFile()

        InsertFileLoc()
        GBAddAppAtt.Hide()
        'btnAppProcess.Enabled = True
        BoxViewApp()
        FetchViewApp()

    End Sub

    Private Sub BtnSaveAddAA_Click(sender As Object, e As EventArgs) Handles BtnSaveAddAA.Click
        CopyFile()

        InsertFileLoc()
        ClearFileLoc()
    End Sub

    Private Sub CopyFile()
        Dim str11 As String
        Dim strday As String
        Dim dloc As Integer

        Dim Ran1 As Integer
        Dim Ran2 As Integer
        Dim max1 As Integer = 999
        Dim max2 As Integer = 999
        Dim min1 As Integer = 100
        Dim min2 As Integer = 100

        Ran1 = Int((max1 - min1 + 1) * Rnd() + min1)
        Ran2 = Int((max2 - min2 + 1) * Rnd() + min2)

        str11 = Flocstr

        ' strday = DateValue(Now())
        strday = CStr(Now())
        'strday = Mid(strday, 1, 2)
        dloc = InStr(strday, "/")
        strday = Mid(strday, dloc + 1)
        dloc = InStr(strday, "/")
        strday = Mid(strday, 1, dloc - 1)

        str22 = "\\IMFI-LENOVO-111\ShareFile\" + "Temp" + "-" + strday + "-" + CStr(Ran2) + "-" + CStr(Hour(Now())) + CStr(Minute(Now())) + CStr(Second(Now())) + ".pdf"

        FileCopy(str11, str22)
    End Sub

    Private Sub InsertFileLoc()

        SQL.AddParam("@FileN", TxtAttFileName.Text)

        SQL.AddParam("@PAEmpCode", TxtTempID.Text)
        SQL.AddParam("@FileDt", TxtAttDate.Text)
        SQL.AddParam("@FileLoc", str22)

        SQL.ExecQuery("INSERT INTO FileLoc " &
                       "(FileName,EmpCode,FileDate,FileLoc) " &
                       "VALUES (@FileN,@PAEmpCode,@FileDt,@FileLoc); " &
                       "BEGIN TRANSACTION; " &
                       "COMMIT;", True)

        'ERROR
        If SQL.HasException(True) Then Exit Sub

        MsgBox("Selected file Added Successfully")

    End Sub

    Private Sub lblFileName_Click(sender As Object, e As EventArgs) Handles LblFileName.Click, LblFileName.TextChanged, TxtAttFileName.TextChanged
        If Not String.IsNullOrWhiteSpace(TxtAttFileName.Text) Then
            If LblFileName.Text = "No file chosen" Then
                BtnSaveAA.Enabled = False
                BtnSaveAddAA.Enabled = False
            Else
                BtnSaveAA.Enabled = True
                BtnSaveAddAA.Enabled = True
            End If
        Else
            BtnSaveAA.Enabled = False
            BtnSaveAddAA.Enabled = False
        End If
    End Sub

    Private Sub ClearFileLoc()
        TxtAttFileName.Clear()
        DTPAttDate.ResetText()
        TxtAttDate.Clear()
        LblFileName.Text = "No file chosen"
    End Sub
    ' ------------------EDIT APPLICANT ATTACHMENT -------------
    Private Sub CLBAppAttachments_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CLBAppAttachments.SelectedIndexChanged, CLBAppAttachments.Click, CLBAppAttachments.DoubleClick
        If CLBAppAttachments.CheckedItems.Count() = 0 Then
            BtnDelAAView.Enabled = False
            BtnEditAAView.Enabled = False
            BtnOpenFileAAView.Enabled = False
        Else
            BtnDelAAView.Enabled = True
            BtnEditAAView.Enabled = True
            BtnOpenFileAAView.Enabled = True
        End If
    End Sub

    Private Sub BtnEditAAView_Click(sender As Object, e As EventArgs) Handles BtnEditAAView.Click
        GBViewApp.Hide()

        TxtTempID.Clear()
        TxtFileID.Clear()
        '
        TxtTempID.Hide()
        TxtFileID.Hide()
        GBAddAppAtt.SetBounds(260, 20, 700, 210)
        GBAddAppAtt.Show()

        LblTitleFile.Text = "Uploaded File Name"

        BtnBrowseAA.Hide()
        LblFileName.SetBounds(152, 140, 150, 18)

        BtnSaveAA.Hide()
        BtnSaveAddAA.Hide()

        BtnUpdAA.SetBounds(152, 170, 81, 24)
        BtnUpdAA.Show()
        LblAAA.Text = "Edit Applicant Attachment Info" & tempFulName
        'BtnSaveAA.Enabled = False
        'BtnSaveAddAA.Enabled = False
        '*** FETCH PREVIOUS ADDRESS INFO

        ClearFileLoc()

        FetchFileLoc()
    End Sub

    Private Sub FetchFileLoc()
        Dim tempstr1 As String = ""
        Dim tempstr2 As String = ""
        Dim aLoc As Integer
        Dim tempdt1 As DateTime
        Dim tempdt2 As DateTime

        aLoc = InStr(CLBAppAttachments.CheckedItems(0), vbTab)
        tempstr1 = Mid(CLBAppAttachments.CheckedItems(0), 1, aLoc - 1)

        aLoc = InStrRev(CLBAppAttachments.CheckedItems(0), "/")
        aLoc = aLoc + 5
        tempstr2 = Mid(CLBAppAttachments.CheckedItems(0), aLoc - 10, 10)

        tempdt1 = Convert.ToDateTime(tempstr2)

        tempdt2 = DateAdd("n", 1, tempdt1)

        If SQL.DBDT IsNot Nothing Then
            SQL.DBDT.Clear()
        End If

        SQL.AddParam("@TFileName", tempstr1)
        SQL.AddParam("@tmpEmpCode", tempcode)
        SQL.AddParam("@Adate", tempdt1)
        SQL.AddParam("@Bdate", tempdt2)
        SQL.ExecQuery("SELECT TOP 1 * FROM FileLoc " &
                      "WHERE FileName=@TFileName AND EmpCode=@tmpEmpCode AND FileDate BETWEEN @Adate AND @Bdate;")
        '
        If SQL.RecordCount < 1 Then Exit Sub
        TxtTempID.Text = tempcode
        For Each r As DataRow In SQL.DBDT.Rows

            TxtFileID.Text = r("FileID")
            LblFileName.Text = r("FileName")
            TxtAttDate.Text = DateTrimR(If(IsDBNull(r("FileDate")), String.Empty, r("FileDate").ToString))

        Next
    End Sub

    Private Sub BtnUpdAA_Click(sender As Object, e As EventArgs) Handles BtnUpdAA.Click

        SQL.AddParam("@PTempID", TxtFileID.Text)
        SQL.AddParam("@FileN", TxtAttFileName.Text)
        SQL.AddParam("@FileDt", TxtAttDate.Text)

        SQL.ExecQuery("UPDATE FileLoc " &
                     "SET FileName=@FileN,FileDate=@FileDt " &
                     "WHERE FileID=@PTempID; " &
                     "BEGIN TRANSACTION; " &
                     "COMMIT;", True)
        'ERROR
        If SQL.HasException(True) Then Exit Sub

        MsgBox("Applicant Attachment Info Updated Successfully")
        GBAddAppAtt.Hide()
        BoxViewApp()
        FetchViewApp()
    End Sub
    ' ------------------DELETE APPLICANT ATTACHMENT -------------
    Private Sub BtnDelAAView_Click(sender As Object, e As EventArgs) Handles BtnDelAAView.Click
        '  DELETE APPLICANT ATTACHMENTS
        Dim tempstr1 As String = ""
        Dim tempstr2 As String = ""

        Dim tempdt As DateTime
        Dim tempdt2 As DateTime
        Dim aLoc As Integer

        aLoc = InStr(CLBAppAttachments.CheckedItems(0), vbTab)
        tempstr1 = Mid(CLBAppAttachments.CheckedItems(0), 1, aLoc - 1)

        aLoc = 0
        aLoc = InStrRev(CLBAppAttachments.CheckedItems(0), "/")
        aLoc = aLoc + 5
        tempstr2 = Mid(CLBAppAttachments.CheckedItems(0), aLoc - 10, 10)

        tempdt = Convert.ToDateTime(tempstr2)

        tempdt2 = DateAdd("n", 1, tempdt)

        If SQL.DBDT IsNot Nothing Then
            SQL.DBDT.Clear()
        End If

        SQL.AddParam("@FileN", tempstr1)
        SQL.AddParam("@tmpEmpCode", tempcode)
        SQL.AddParam("@Adate", tempdt)
        SQL.AddParam("@Bdate", tempdt2)

        SQL.ExecQuery("DELETE FROM FileLoc " &
                      "WHERE FileName=@FileN AND EmpCode=@tmpEmpCode AND FileDate BETWEEN @Adate AND @Bdate;")

        MsgBox("Selected Applicant Attacment  deleted")

        'FetchFileLoc()
        'FetchViewPreAddr()
        FetchViewApp()
        'FetchViewAppAtc()
        BtnDelAAView.Enabled = False
        BtnEditAAView.Enabled = False
        BtnOpenFileAAView.Enabled = False
    End Sub

    ' ------------------OPEN APPLICANT ATTACHMENT -------------
    Private Sub BtnOpenFileAAView_Click(sender As Object, e As EventArgs) Handles BtnOpenFileAAView.Click

        Dim tempstr1 As String = ""
        Dim tempstr2 As String = ""
        Dim aLoc As Integer
        Dim tempdt1 As DateTime
        Dim tempdt2 As DateTime

        aLoc = InStr(CLBAppAttachments.CheckedItems(0), vbTab)
        tempstr1 = Mid(CLBAppAttachments.CheckedItems(0), 1, aLoc - 1)

        aLoc = InStrRev(CLBAppAttachments.CheckedItems(0), "/")
        aLoc = aLoc + 5
        tempstr2 = Mid(CLBAppAttachments.CheckedItems(0), aLoc - 10, 10)

        tempdt1 = Convert.ToDateTime(tempstr2)

        tempdt2 = DateAdd("n", 1, tempdt1)

        If SQL.DBDT IsNot Nothing Then
            SQL.DBDT.Clear()
        End If

        SQL.AddParam("@TFileName", tempstr1)
        SQL.AddParam("@tmpEmpCode", tempcode)
        SQL.AddParam("@Adate", tempdt1)
        SQL.AddParam("@Bdate", tempdt2)
        SQL.ExecQuery("SELECT TOP 1 * FROM FileLoc " &
                      "WHERE FileName=@TFileName AND EmpCode=@tmpEmpCode AND FileDate BETWEEN @Adate AND @Bdate;")
        '
        If SQL.RecordCount < 1 Then Exit Sub

        For Each r As DataRow In SQL.DBDT.Rows

            tempstr1 = If(IsDBNull(r("FileLoc")), String.Empty, r("FileLoc").ToString)

        Next

        Process.Start(tempstr1)

    End Sub
    'PRINT APPLICATION


    Private Sub BtnPrintApp_Click(sender As Object, e As EventArgs) Handles BtnPrintApp.Click
        TempletApp()
    End Sub
    Private Sub TempletApp()
        Dim tempstr1 As String = ""
        Dim tempstr2 As String = ""
        Dim coidapp As Integer
        Dim oWord As Word.Application
        Dim oDoc As Word.Document

        'START WORD AND OPEN THE DOCUMENT TEMPLATE
        oWord = CreateObject("Word.Application")
        oWord.Visible = False
        'oWord.Application.Visible = False
        oWord.Application.ShowWindowsInTaskbar = False


        'Get the Word Tempelate
        'oDoc = oWord.Documents.Add("\\NJEDIOB01\MyShare\templet\PrintApp001.dotx")
        oDoc = oWord.Documents.Add("\\IMFI-LENOVO-111\ShareFile\MyShare\templet\Emp_App.dotx")
        'oDoc.Application.Visible = False
        'oDoc.Application.ShowWindowsInTaskbar = False
        'GET APPLICANT DATA

        SQL.AddParam("@tmpEmpCode", tempcode)


        SQL.ExecQuery("SELECT TOP 1 * FROM Employees " &
                      "WHERE EmpCode=@tmpEmpCode AND ApplicationID=1;")
        '
        If SQL.RecordCount < 1 Then Exit Sub

        For Each r As DataRow In SQL.DBDT.Rows

            oDoc.Bookmarks.Item("APPDT").Range.Text = DateTrimR(r("AppDate"))
            oDoc.Bookmarks.Item("POSAPP1").Range.Text = If(IsDBNull(r("AppPosition")), String.Empty, r("AppPosition").ToString)
            ' tempstr1 = If(IsDBNull(r("EmpFirstName")), String.Empty, r("EmpFirstName").ToString)
            'tempstr2 = If(IsDBNull(r("EmpLastName")), String.Empty, r("EmpLastName").ToString)
            oDoc.Bookmarks.Item("LNAME").Range.Text = If(IsDBNull(r("EmpLastName")), String.Empty, r("EmpLastName").ToString) + ", "
            oDoc.Bookmarks.Item("FNAME").Range.Text = If(IsDBNull(r("EmpFirstName")), String.Empty, r("EmpFirstName").ToString)
            oDoc.Bookmarks.Item("MNAME").Range.Text = If(IsDBNull(r("EmpMI")), String.Empty, r("EmpMI").ToString)
            oDoc.Bookmarks.Item("SNF").Range.Text = If(IsDBNull(r("Fssn")), String.Empty, r("Fssn").ToString)
            oDoc.Bookmarks.Item("SM").Range.Text = If(IsDBNull(r("Mssn")), String.Empty, r("Mssn").ToString)
            oDoc.Bookmarks.Item("SNL").Range.Text = If(IsDBNull(r("Lssn")), String.Empty, r("Lssn").ToString)
            oDoc.Bookmarks.Item("EADD").Range.Text = If(IsDBNull(r("EmpAddress")), String.Empty, r("EmpAddress").ToString)
            oDoc.Bookmarks.Item("EAPT").Range.Text = If(IsDBNull(r("EmpApt")), String.Empty, r("EmpApt").ToString)
            oDoc.Bookmarks.Item("ECITY").Range.Text = If(IsDBNull(r("EmpCity")), String.Empty, r("EmpCity").ToString)

            oDoc.Bookmarks.Item("ESTATE").Range.Text = If(IsDBNull(r("EmpState")), String.Empty, r("EmpState").ToString)
            oDoc.Bookmarks.Item("EZIP").Range.Text = If(IsDBNull(r("EmpFirstZip")), String.Empty, r("EmpFirstZip").ToString)
            oDoc.Bookmarks.Item("HPHONE").Range.Text = If(IsDBNull(r("HomePhone")), String.Empty, r("HomePhone").ToString)
            'tempstr1 = If(IsDBNull(r("ProofAge")), String.Empty, r("ProofAge").ToString)
            If r("ProofAge") = 1 Then
                tempstr2 = "Yes"
            Else
                tempstr2 = "No"
            End If
            oDoc.Bookmarks.Item("PROOFAGE").Range.Text = tempstr2
            coidapp = r("CompanyID")
        Next

        'GET INFO COMPANY
        If SQL.DBDT IsNot Nothing Then
            SQL.DBDT.Clear()
        End If
        'oDoc.Bookmarks.Item("CONAME").Range.Text = coidapp
        SQL.AddParam("@tmpCoCode", coidapp)


        SQL.ExecQuery("SELECT TOP 1 * FROM Company " &
                      "WHERE CompanyID=@tmpCoCode;")
        '
        If SQL.RecordCount < 1 Then Exit Sub

        For Each r As DataRow In SQL.DBDT.Rows

            oDoc.Bookmarks.Item("CONAME").Range.Text = If(IsDBNull(r("CompanyName")), String.Empty, r("CompanyName").ToString)
            oDoc.Bookmarks.Item("COADDR1").Range.Text = If(IsDBNull(r("CoAddress")), String.Empty, r("CoAddress").ToString)
            oDoc.Bookmarks.Item("COCITY").Range.Text = If(IsDBNull(r("CoCity")), String.Empty, r("CoCity").ToString)
            oDoc.Bookmarks.Item("COSTATE").Range.Text = If(IsDBNull(r("CoState")), String.Empty, r("CoState").ToString)
            oDoc.Bookmarks.Item("COZIP").Range.Text = If(IsDBNull(r("CoZipCode")), String.Empty, r("CoZipCode").ToString)
        Next

        ' Save my document as defoult word format and Close document and quit word app
        oDoc.SaveAs2("\\IMFI-LENOVO-111\ShareFile\MyShare\bbb12")
        oDoc.Close()
        oWord.Quit()
        'Convert saved word document to pdf
        Dim Source As String = "\\IMFI-LENOVO-111\ShareFile\MyShare\bbb12.docx"
        Dim Target As String = "\\IMFI-LENOVO-111\ShareFile\MyShare\Aaaa12"
        Word2PDF(Source, Target)
        '###########
        'Delete Word file
        Dim FileToDelete As String

        FileToDelete = "\\IMFI-LENOVO-111\ShareFile\MyShare\bbb12.docx"
        'FileToDelete = "C:\Users\Owner\Documents\testDelete.txt"
        If System.IO.File.Exists(FileToDelete) = True Then
            'System.IO.File.Move(FileToDelete, ">>bbb12.docx")
            'System.IO.File.Delete(FileToDelete)
            'MsgBox("File Deleted")

        End If
        '###########
        'open pdf file to browser
        Process.Start("\\IMFI-LENOVO-111\ShareFile\MyShare\Aaaa12.pdf")
    End Sub

    'WORD TO PDF CONVERSION

    Private Sub Word2PDF(Source As Object, Target As Object)
        Dim MSdoc As Word.Application
        Dim Unknown As Object = Type.Missing
        MSdoc = CreateObject("Word.Application")
        'Creating the instance of Word Application          
        'If MSdoc Is Nothing Then
        'MSdoc = CreateObject("Word.Application")
        'E nd If

        Try
            MSdoc.Visible = False
            ' MSdoc.Application.Visible = False
            MSdoc.ShowWindowsInTaskbar = False
            'MSdoc.Application.ShowWindowsInTaskbar = False


            MSdoc.Documents.Open(Source, Unknown, Unknown, Unknown, Unknown, Unknown,
            Unknown, Unknown, Unknown, Unknown, Unknown, Unknown,
            Unknown, Unknown, Unknown, Unknown)

            'MSdoc.Application.Visible = False
            'MSdoc.ShowWindowsInTaskbar = False
            'MSdoc.ShowWindowsInTaskbar = False
            'MSdoc.Application.ShowWindowsInTaskbar = False

            MSdoc.WindowState = Microsoft.Office.Interop.Word.WdWindowState.wdWindowStateMinimize

            Dim format As Object = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF

            MSdoc.ActiveDocument.SaveAs(Target, format, Unknown, Unknown, Unknown, Unknown,
            Unknown, Unknown, Unknown, Unknown, Unknown, Unknown,
            Unknown, Unknown, Unknown, Unknown)


        Catch e As Exception
            MessageBox.Show(e.Message)
        Finally

            If MSdoc IsNot Nothing Then
                MSdoc.Visible = False
                ' MSdoc.Application.Visible = False
                MSdoc.ShowWindowsInTaskbar = False
                'MsgBox("Check pt 1")
                MSdoc.Documents.Close(Unknown, Unknown, Unknown)
                'MsgBox("Check pt 2")
            End If

            ' for closing the application
            'WordDoc.Quit(Unknown, Unknown, Unknown)
            MSdoc.Quit(Unknown, Unknown, Unknown)

        End Try
    End Sub

    Private Sub BtnAllFormApp_Click(sender As Object, e As EventArgs) Handles BtnAllFormApp.Click
        Dim Source As String = "\\IMFI-LENOVO-111\ShareFile\MyShare\abcd.docx"
        Dim Target As String = "\\IMFI-LENOVO-111\ShareFile\MyShare\abcd1"
        Word2PDF(Source, Target)
    End Sub


End Class
