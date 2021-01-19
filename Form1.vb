Public Class Form1
    Inherits System.Windows.Forms.Form

    Private objMutex As System.Threading.Mutex
    Dim waitDlg As WaitDialog   ''進行状況フォームクラス  

    Dim SqlCmd1 As New SqlClient.SqlCommand
    Dim DaList1 As New SqlClient.SqlDataAdapter
    Dim DsList1 As New DataSet
    Dim DtView1 As DataView

    Dim strSQL, str_ANS, Err_F As String
    Dim WK_close_date As Date
    Dim i, r, r1, r2 As Integer

#Region " Windows フォーム デザイナで生成されたコード "

    Public Sub New()
        MyBase.New()

        ' この呼び出しは Windows フォーム デザイナで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後に初期化を追加します。

    End Sub

    ' Form は、コンポーネント一覧に後処理を実行するために dispose をオーバーライドします。
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    ' Windows フォーム デザイナで必要です。
    Private components As System.ComponentModel.IContainer

    ' メモ : 以下のプロシージャは、Windows フォーム デザイナで必要です。
    'Windows フォーム デザイナを使って変更してください。  
    ' コード エディタを使って変更しないでください。
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Date1 As GrapeCity.Win.Input.Interop.Date
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form1))
        Me.Button1 = New System.Windows.Forms.Button
        Me.Date1 = New GrapeCity.Win.Input.Interop.Date
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Button2 = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        CType(Me.Date1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button1.Location = New System.Drawing.Point(32, 136)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(80, 32)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "締め処理"
        '
        'Date1
        '
        Me.Date1.DisabledForeColor = System.Drawing.SystemColors.WindowText
        Me.Date1.DisplayFormat = New GrapeCity.Win.Input.Interop.DateDisplayFormat("yyy/MM", "", "")
        Me.Date1.DropDown = New GrapeCity.Win.Input.Interop.DropDown(GrapeCity.Win.Input.Interop.ButtonPosition.Inside, True, GrapeCity.Win.Input.Interop.Visibility.NotShown, System.Windows.Forms.FlatStyle.System)
        Me.Date1.DropDownCalendar.Size = New System.Drawing.Size(158, 165)
        Me.Date1.Enabled = False
        Me.Date1.Format = New GrapeCity.Win.Input.Interop.DateFormat("yyyy/MM", "", "")
        Me.Date1.HighlightText = GrapeCity.Win.Input.Interop.HighlightText.All
        Me.Date1.Location = New System.Drawing.Point(96, 24)
        Me.Date1.MinDate = New GrapeCity.Win.Input.Interop.DateTimeEx(New Date(2000, 1, 1, 0, 0, 0, 0))
        Me.Date1.Name = "Date1"
        Me.Date1.Shortcuts = New GrapeCity.Win.Input.Interop.ShortcutCollection(New String() {"F2", "F5"}, New GrapeCity.Win.Input.Interop.KeyActions() {GrapeCity.Win.Input.Interop.KeyActions.Clear, GrapeCity.Win.Input.Interop.KeyActions.Now})
        Me.Date1.Size = New System.Drawing.Size(72, 32)
        Me.Date1.TabIndex = 0
        Me.Date1.TextHAlign = GrapeCity.Win.Input.Interop.AlignHorizontal.Center
        Me.Date1.TextVAlign = GrapeCity.Win.Input.Interop.AlignVertical.Middle
        Me.Date1.Value = New GrapeCity.Win.Input.Interop.DateTimeEx(New Date(2007, 5, 24, 12, 19, 58, 0))
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Navy
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label1.ForeColor = System.Drawing.SystemColors.Window
        Me.Label1.Location = New System.Drawing.Point(32, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 32)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "締月"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(255, Byte), CType(255, Byte))
        Me.Label2.Location = New System.Drawing.Point(192, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(136, 16)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Label2"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Label2.Visible = False
        '
        'Button2
        '
        Me.Button2.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button2.Location = New System.Drawing.Point(232, 136)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(80, 32)
        Me.Button2.TabIndex = 2
        Me.Button2.Text = "終了"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Location = New System.Drawing.Point(152, 64)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(88, 32)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Label3"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Location = New System.Drawing.Point(152, 96)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(88, 32)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "Label4"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(255, Byte), CType(255, Byte))
        Me.Label5.Location = New System.Drawing.Point(192, 16)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(136, 16)
        Me.Label5.TabIndex = 6
        Me.Label5.Text = "Label5"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Label5.Visible = False
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Location = New System.Drawing.Point(32, 96)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(112, 32)
        Me.Label6.TabIndex = 8
        Me.Label6.Text = "ベスト分： "
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Location = New System.Drawing.Point(32, 64)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(112, 32)
        Me.Label7.TabIndex = 7
        Me.Label7.Text = "オール電化分： "
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(7, 16)
        Me.ClientSize = New System.Drawing.Size(338, 183)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Date1)
        Me.Controls.Add(Me.Button1)
        Me.Font = New System.Drawing.Font("MS UI Gothic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "加入締め処理 Var 2.0"
        CType(Me.Date1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objMutex = New System.Threading.Mutex(False, "best_close")
        If objMutex.WaitOne(0, False) = False Then
            MessageBox.Show("すでに起動しています", "実行結果")
            Application.Exit()
        End If
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        DB_INIT()
        DB_OPEN()
        DsList1.Clear()

        '締月セット
        strSQL = "SELECT * FROM CLS_CODE WHERE (CLS_NO = '999')"
        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        SqlCmd1.CommandTimeout = 3000
        DaList1.Fill(DsList1, "CLS_CODE")
        DtView1 = New DataView(DsList1.Tables("CLS_CODE"), "CLS_CODE='2'", "", DataViewRowState.CurrentRows)
        If DtView1.Count <> 0 Then
            WK_close_date = DateAdd(DateInterval.Month, 1, DtView1(0)("CLS_CODE_NAME"))
            Date1.Value = Format(WK_close_date, "yyyy/MM")
        End If
        Call Date1_Leave(sender, e)

        '対象件数
        'ALL8
        strSQL = "SELECT ordr_no, line_no, seq, cont_flg, prch_date, fin_date, cxl_date, close_date"
        strSQL += " FROM All8_Wrn_sub"
        strSQL += " WHERE (cont_flg = 'A')"
        strSQL += " AND (fin_date < CONVERT(DATETIME, '" & Label5.Text & "', 102))"
        strSQL += " AND (close_date IS NULL)"
        strSQL += " OR (cont_flg = 'C')"
        strSQL += " AND (cxl_date < CONVERT(DATETIME, '" & Label5.Text & "', 102))"
        strSQL += " AND (cxl_close_date IS NULL)"
        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        SqlCmd1.CommandTimeout = 600
        r1 = DaList1.Fill(DsList1, "All8_Wrn_sub")
        'BEST
        strSQL = "SELECT ordr_no, line_no, seq, cont_flg, prch_date, fin_date, cxl_date, close_date"
        strSQL += " FROM Wrn_sub"
        strSQL += " WHERE (cont_flg = 'A')"
        strSQL += " AND (fin_date < CONVERT(DATETIME, '" & Label5.Text & "', 102))"
        strSQL += " AND (close_date IS NULL)"
        strSQL += " OR (cont_flg = 'C')"
        strSQL += " AND (cxl_date < CONVERT(DATETIME, '" & Label5.Text & "', 102))"
        strSQL += " AND (cxl_close_date IS NULL)"
        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        SqlCmd1.CommandTimeout = 600
        r2 = DaList1.Fill(DsList1, "Wrn_sub")

        If r1 <> 0 Then
            Label3.Text = Format(r1, "##,##0") & " 件"
        Else
            Label3.Text = "0 件"
        End If

        If r2 <> 0 Then
            Label4.Text = Format(r2, "##,##0") & " 件"
        Else
            Label4.Text = "0 件"
        End If

        If r1 + r2 <> 0 Then
            Button1.Enabled = True
        Else
            Button1.Enabled = False
        End If

        DB_CLOSE()
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub Date1_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Date1.Leave
        Label2.Text = DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 1, CDate(Date1.Text & "/01")))
        Label5.Text = DateAdd(DateInterval.Month, 1, CDate(Date1.Text & "/01"))
    End Sub

    'Sub F_CHK()
    '    Err_F = "0"

    '    If Format(WK_close_date, "yyyy/MM") > Date1.Text Then
    '        If MsgBox("既に締め処理済、追加で処理をしますか？", MsgBoxStyle.YesNo, "Error") = MsgBoxResult.Yes Then
    '            Err_F = "0"
    '        Else
    '            Err_F = "1" : Exit Sub
    '        End If

    '    End If
    'End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        str_ANS = MsgBox("加入データの締め処理を行います。よろしいですか？", MsgBoxStyle.OKCancel, "確認")
        If str_ANS = "1" Then   'OK
            Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
            'Call F_CHK()
            'If Err_F = "0" Then

            ' 進行状況ダイアログの初期化処理
            waitDlg = New WaitDialog        ' 進行状況ダイアログ
            waitDlg.Owner = Me              ' ダイアログのオーナーを設定する
            waitDlg.MainMsg = Nothing       ' 処理の概要を表示
            waitDlg.ProgressMax = 0         ' 全体の処理件数を設定
            waitDlg.ProgressMin = 0         ' 処理件数の最小値を設定（0件から開始）
            waitDlg.ProgressStep = 1        ' 何件ごとにメータを進めるかを設定
            waitDlg.ProgressValue = 0       ' 最初の件数を設定
            Me.Enabled = False              ' オーナーのフォームを無効にする
            waitDlg.Show()                  ' 進行状況ダイアログを表示する

            waitDlg.MainMsg = "オール電化　締め処理　実行中"      ' 進行状況ダイアログのメーターを設定
            waitDlg.ProgressMsg = ""        ' 進行状況ダイアログのメーターを設定
            waitDlg.ProgressMax = r1        ' 全体の処理件数を設定
            waitDlg.ProgressValue = 0       ' 最初の件数を設定
            Application.DoEvents()          ' メッセージ処理を促して表示を更新する

            DB_OPEN()

            DtView1 = New DataView(DsList1.Tables("All8_Wrn_sub"), "", "", DataViewRowState.CurrentRows)
            For i = 0 To DtView1.Count - 1

                waitDlg.ProgressMsg = Fix((i + 1) * 100 / DtView1.Count) & "%　（" & Format(i + 1, "##,##0") & " / " & Format(DtView1.Count, "##,##0") & " 件）"
                waitDlg.Text = "実行中・・・" & Fix((i + 1) * 100 / DtView1.Count) & "%"
                Application.DoEvents()  ' メッセージ処理を促して表示を更新する
                waitDlg.PerformStep()   ' 処理カウントを1ステップ進める

                strSQL = "UPDATE All8_Wrn_sub"
                If DtView1(i)("cont_flg") = "A" Then
                    strSQL = strSQL & " SET close_date = CONVERT(DATETIME, '" & Label2.Text & "', 102)"
                    strSQL = strSQL & ", close_cont_flg = cont_flg"
                Else
                    If IsDBNull(DtView1(i)("close_date")) Then  '同月ｷｬﾝｾﾙ
                        strSQL = strSQL & " SET close_date = CONVERT(DATETIME, '" & Label2.Text & "', 102)"
                        strSQL = strSQL & ", cxl_close_date = CONVERT(DATETIME, '" & Label2.Text & "', 102)"
                        strSQL = strSQL & ", close_cont_flg = cont_flg"
                    Else                                        '別月ｷｬﾝｾﾙ
                        strSQL = strSQL & " SET cxl_close_date = CONVERT(DATETIME, '" & Label2.Text & "', 102)"
                    End If
                End If
                strSQL = strSQL & " WHERE (ordr_no = '" & DtView1(i)("ordr_no") & "')"
                strSQL = strSQL & " AND (line_no = '" & DtView1(i)("line_no") & "')"
                strSQL = strSQL & " AND (seq = " & DtView1(i)("seq") & ")"
                SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
                SqlCmd1.CommandTimeout = 3600
                SqlCmd1.ExecuteNonQuery()

            Next

            'cls
            strSQL = "UPDATE CLS_CODE"
            strSQL = strSQL & " SET CLS_CODE_NAME = '" & Label2.Text & "'"
            strSQL = strSQL & " WHERE (CLS_NO = '999') AND (CLS_CODE = '2')"
            SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
            SqlCmd1.CommandTimeout = 3600
            SqlCmd1.ExecuteNonQuery()

            waitDlg.MainMsg = "ベスト電器　締め処理　実行中"      ' 進行状況ダイアログのメーターを設定
            waitDlg.ProgressMsg = ""        ' 進行状況ダイアログのメーターを設定
            waitDlg.ProgressMax = r2        ' 全体の処理件数を設定
            waitDlg.ProgressValue = 0       ' 最初の件数を設定
            Application.DoEvents()          ' メッセージ処理を促して表示を更新する

            DtView1 = New DataView(DsList1.Tables("Wrn_sub"), "", "", DataViewRowState.CurrentRows)
            For i = 0 To DtView1.Count - 1

                waitDlg.ProgressMsg = Fix((i + 1) * 100 / DtView1.Count) & "%　（" & Format(i + 1, "##,##0") & " / " & Format(DtView1.Count, "##,##0") & " 件）"
                waitDlg.Text = "実行中・・・" & Fix((i + 1) * 100 / DtView1.Count) & "%"
                Application.DoEvents()  ' メッセージ処理を促して表示を更新する
                waitDlg.PerformStep()   ' 処理カウントを1ステップ進める

                strSQL = "UPDATE Wrn_sub"
                If DtView1(i)("cont_flg") = "A" Then
                    strSQL = strSQL & " SET close_date = CONVERT(DATETIME, '" & Label2.Text & "', 102)"
                    strSQL = strSQL & ", close_cont_flg = cont_flg"
                Else
                    If IsDBNull(DtView1(i)("close_date")) Then  '同月ｷｬﾝｾﾙ
                        strSQL = strSQL & " SET close_date = CONVERT(DATETIME, '" & Label2.Text & "', 102)"
                        strSQL = strSQL & ", cxl_close_date = CONVERT(DATETIME, '" & Label2.Text & "', 102)"
                        strSQL = strSQL & ", close_cont_flg = cont_flg"
                    Else                                        '別月ｷｬﾝｾﾙ
                        strSQL = strSQL & " SET cxl_close_date = CONVERT(DATETIME, '" & Label2.Text & "', 102)"
                    End If
                End If
                strSQL = strSQL & " WHERE (ordr_no = '" & DtView1(i)("ordr_no") & "')"
                strSQL = strSQL & " AND (line_no = '" & DtView1(i)("line_no") & "')"
                strSQL = strSQL & " AND (seq = " & DtView1(i)("seq") & ")"
                SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
                SqlCmd1.CommandTimeout = 3600
                SqlCmd1.ExecuteNonQuery()

            Next

            'Input_Seq
            strSQL = "DELETE FROM Input_Seq"
            SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
            SqlCmd1.CommandTimeout = 600
            SqlCmd1.ExecuteNonQuery()

            'Count_tbl
            strSQL = "UPDATE Count_tbl SET seq = 1001 WHERE (cls = '002')"
            SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
            SqlCmd1.CommandTimeout = 600
            SqlCmd1.ExecuteNonQuery()

            'cls
            strSQL = "UPDATE CLS_CODE"
            strSQL = strSQL & " SET CLS_CODE_NAME = '" & Label2.Text & "'"
            strSQL = strSQL & " WHERE (CLS_NO = '999') AND (CLS_CODE = '3')"
            SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
            SqlCmd1.CommandTimeout = 3600
            SqlCmd1.ExecuteNonQuery()

            DB_CLOSE()
            Button1.Enabled = False
            MsgBox("締め処理しました。", MsgBoxStyle.Information, "終了")
            waitDlg.Close()         '進行状況ダイアログを閉じる
            Me.Enabled = True       'オーナーのフォームを有効にする
            'End If
            Me.Cursor = System.Windows.Forms.Cursors.Default
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Application.Exit()  '終了
    End Sub
End Class
