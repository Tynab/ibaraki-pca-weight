Imports System.ComponentModel
Imports System.Math
Imports System.Net
Imports System.Windows.Forms
Imports System.Windows.Forms.Keys

''' <summary>
''' Form hiển thị tiến độ tải bộ cài mới và khởi chạy installer sau khi tải thành công.
''' </summary>
Public Class FrmUpdate
#Region "Fields"
    Private ReadOnly _wc As New WebClient()
    Private _downloadSucceeded As Boolean
#End Region

#Region "Overridden"
    ''' <summary>
    ''' Ẩn form khỏi danh sách Alt+Tab để hộp cập nhật hoạt động như sub window.
    ''' </summary>
    Protected Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or &H80
            Return cp
        End Get
    End Property

    ''' <summary>
    ''' Chặn Alt+F4 trong lúc cập nhật để tránh hủy tải không chủ ý.
    ''' </summary>
    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
        Return keyData = (Alt Or F4) OrElse MyBase.ProcessCmdKey(msg, keyData)
    End Function
#End Region

#Region "Events"
    ''' <summary>
    ''' Khởi tạo trạng thái hiển thị khi form cập nhật được load.
    ''' </summary>
    Private Sub FrmUpdate_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lblCapacity.Text = ""
        lblPercent.Text = ""
        pnlProgressBar.Width = 1
    End Sub

    ''' <summary>
    ''' Chuẩn bị thư mục tạm và bắt đầu tải bộ cài mới.
    ''' </summary>
    Private Sub FrmUpdate_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        FIFrm()
        CrtDirAdv(FRNT_PATH)
        DelFileAdv(FILE_SETUP_ADR)
        AddHandler _wc.DownloadProgressChanged, AddressOf Upd_DownloadProgressChanged
        AddHandler _wc.DownloadFileCompleted, AddressOf Upd_DownloadFileCompleted

        Try
            Dim setupUrl As String = _wc.DownloadString(My.Resources.link_app)
            _wc.DownloadFileAsync(New Uri(setupUrl), FILE_SETUP_ADR)
        Catch ex As Exception
            ErrSty($"{vbCrLf}Không thể tải bản cập nhật: {ex.Message}{vbCrLf}")
            Close()
        End Try
    End Sub

    ''' <summary>
    ''' Cập nhật dung lượng, phần trăm và thanh tiến độ khi WebClient báo tiến trình tải.
    ''' </summary>
    Private Sub Upd_DownloadProgressChanged(sender As Object, e As DownloadProgressChangedEventArgs)
        lblCapacity.Text = String.Format("{0} MB / {1} MB", (e.BytesReceived / 1024D / 1024D).ToString("0.00"), (e.TotalBytesToReceive / 1024D / 1024D).ToString("0.00"))
        lblPercent.Text = $"{e.ProgressPercentage}%"

        Dim progressWidth As Integer = CInt(Ceiling(e.ProgressPercentage * pnlMain.ClientSize.Width / 100D))
        pnlProgressBar.Width = Min(pnlMain.ClientSize.Width, Max(1, progressWidth))
    End Sub

    ''' <summary>
    ''' Đóng form sau khi tải xong, hoặc ghi lỗi nếu quá trình tải thất bại.
    ''' </summary>
    Private Sub Upd_DownloadFileCompleted(sender As Object, e As AsyncCompletedEventArgs)
        If e.Cancelled Then
            ErrSty($"{vbCrLf}Quá trình tải bản cập nhật đã bị hủy.{vbCrLf}")
        ElseIf e.Error IsNot Nothing Then
            ErrSty($"{vbCrLf}Tải bản cập nhật thất bại: {e.Error.Message}{vbCrLf}")
        Else
            _downloadSucceeded = True
            lblPercent.Text = "100%"
            pnlProgressBar.Width = pnlMain.ClientSize.Width
        End If

        tmrMain.StopAdv()
        Close()
    End Sub

    ''' <summary>
    ''' Dự phòng cho timer cũ: chỉ đóng form khi download đã hoàn tất thành công.
    ''' </summary>
    Private Sub TmrMain_Tick(sender As Object, e As EventArgs) Handles tmrMain.Tick
        If _downloadSucceeded Then
            tmrMain.StopAdv()
            Close()
        End If
    End Sub

    ''' <summary>
    ''' Chạy hiệu ứng fade out khi form cập nhật đóng.
    ''' </summary>
    Private Sub FrmUpdate_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        FOFrm()
    End Sub

    ''' <summary>
    ''' Dọn WebClient và chỉ chạy installer nếu file cập nhật đã tải thành công.
    ''' </summary>
    Private Sub FrmUpdate_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        RemoveHandler _wc.DownloadProgressChanged, AddressOf Upd_DownloadProgressChanged
        RemoveHandler _wc.DownloadFileCompleted, AddressOf Upd_DownloadFileCompleted
        _wc.Dispose()

        If _downloadSucceeded Then
            System.Diagnostics.Process.Start(FILE_SETUP_ADR)
            KillPrcs(My.Resources.app_name)
        End If
    End Sub
#End Region
End Class
