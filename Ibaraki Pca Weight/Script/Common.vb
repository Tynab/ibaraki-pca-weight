Imports System.Console
Imports System.ConsoleColor
Imports System.Diagnostics.Process
Imports System.IO
Imports System.IO.Directory
Imports System.Net
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Threading.Thread
Imports System.Windows.Forms
Imports Excel = Microsoft.Office.Interop.Excel

''' <summary>
''' Chứa các helper dùng chung: kiểm tra cập nhật, điều khiển Excel, nhập liệu console và hiệu ứng giao diện.
''' </summary>
Friend Module Common
#Region "Helper"
    ''' <summary>
    ''' Kiểm tra máy hiện tại có truy cập được mạng hay không.
    ''' </summary>
    ''' <returns>True nếu endpoint cơ bản phản hồi, ngược lại là False.</returns>
    Private Function IsNetAvail() As Boolean
        Try
            Using objResp As WebResponse = WebRequest.Create(New Uri(My.Resources.link_base)).GetResponse()
                Return True
            End Using
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Kiểm tra phiên bản trên server và mở form cập nhật khi có bản mới.
    ''' </summary>
    Private Sub ChkUpd()
        HdrSty("アップデートの確認...")
        If Not IsNetAvail() Then
            Return
        End If

        Try
            Using client As New WebClient()
                Dim versionText As String = client.DownloadString(My.Resources.link_ver)
                If Not versionText.Contains(My.Resources.app_ver) Then
                    MsgBox($"「{My.Resources.app_true_name}」新しいバージョンが利用可能！", 262144, Title:="更新")
                    Using frmUpd As New FrmUpdate()
                        frmUpd.ShowDialog()
                    End Using
                End If
            End Using
        Catch ex As Exception
            ErrSty($"{vbCrLf}Không thể kiểm tra bản cập nhật: {ex.Message}{vbCrLf}")
        End Try
    End Sub

    ''' <summary>
    ''' Lưu trạng thái license hợp lệ vào user settings.
    ''' </summary>
    Friend Sub UpdVldLic()
        My.Settings.Chk_Key = True
        My.Settings.Save()
    End Sub

    ''' <summary>
    ''' Tăng opacity để form xuất hiện mềm hơn.
    ''' </summary>
    <Extension()>
    Friend Sub FIFrm(frm As Form)
        While frm.Opacity < 1
            frm.Opacity = System.Math.Min(1, frm.Opacity + 0.05)
            frm.Update()
            Sleep(10)
        End While
    End Sub

    ''' <summary>
    ''' Giảm opacity để form đóng mềm hơn.
    ''' </summary>
    <Extension()>
    Friend Sub FOFrm(frm As Form)
        While frm.Opacity > 0
            frm.Opacity = System.Math.Max(0, frm.Opacity - 0.05)
            frm.Update()
            Sleep(10)
        End While
    End Sub
#End Region

#Region "Master"
    ''' <summary>
    ''' Đóng tất cả tiến trình có tên được truyền vào.
    ''' </summary>
    ''' <param name="name">Tên tiến trình không bao gồm phần mở rộng.</param>
    Friend Sub KillPrcs(name As String)
        For Each item In GetProcessesByName(name)
            Try
                item.Kill()
                item.WaitForExit(5000)
            Catch ex As Exception
                ErrSty($"{vbCrLf}Không thể đóng tiến trình {name}: {ex.Message}{vbCrLf}")
            Finally
                item.Dispose()
            End Try
        Next
    End Sub

    ''' <summary>
    ''' Yêu cầu người dùng đóng Excel rồi dọn các tiến trình Excel còn sót.
    ''' </summary>
    Private Sub KillXl()
        Clear()
        HdrSty("警告：このアプリケーションを使用する前に、すべての「エクセル」を閉じてください。「エンター」キーを押して続行します...")
        ReadLine()
        KillPrcs(XL_NAME)
    End Sub

    ''' <summary>
    ''' Chạy luồng chính: kiểm tra cập nhật, chọn file Excel, ghi dữ liệu và mở lại file sau khi lưu.
    ''' </summary>
    Friend Sub RunApp()
        ChkUpd()
        KillXl()

        Dim filePath As String = Nothing
        Using ofd As New OpenFileDialog With {
            .Multiselect = False,
            .Title = "「エクセル」ドキュメントを開く",
            .Filter = "「エクセル」ドキュメント|*.xlsx;*.xls"
        }
            If ofd.ShowDialog() <> DialogResult.OK Then
                Return
            End If
            filePath = ofd.FileName
        End Using

        Dim xlApp As Excel.Application = Nothing
        Dim workbook As Excel.Workbook = Nothing
        Dim workbookClosed As Boolean = False

        Try
            xlApp = New Excel.Application()
            workbook = xlApp.Workbooks.Open(filePath)
            WtIbarakiPca(xlApp)
            workbook.Close(SaveChanges:=True)
            workbookClosed = True
            System.Diagnostics.Process.Start(filePath)
        Catch ex As Exception
            ErrSty($"{vbCrLf}Không thể xử lý file Excel: {ex.Message}{vbCrLf}")
        Finally
            If workbook IsNot Nothing Then
                If Not workbookClosed Then
                    Try
                        workbook.Close(SaveChanges:=False)
                    Catch
                    End Try
                End If
                ReleaseComObject(workbook)
            End If

            If xlApp IsNot Nothing Then
                Try
                    xlApp.Quit()
                Catch
                End Try
                ReleaseComObject(xlApp)
            End If
        End Try
    End Sub
#End Region

#Region "Main"
    ''' <summary>
    ''' Tạo thư mục nếu đường dẫn chưa tồn tại.
    ''' </summary>
    ''' <param name="path">Đường dẫn thư mục cần tạo.</param>
    Friend Sub CrtDirAdv(path As String)
        If Not Exists(path) Then
            CreateDirectory(path)
        End If
    End Sub

    ''' <summary>
    ''' Xóa file nếu file đang tồn tại.
    ''' </summary>
    ''' <param name="path">Đường dẫn file cần xóa.</param>
    Friend Sub DelFileAdv(path As String)
        If File.Exists(path) Then
            File.Delete(path)
        End If
    End Sub

    ''' <summary>
    ''' Hỏi lựa chọn dạng 1/0 và lặp lại cho đến khi người dùng nhập hợp lệ.
    ''' </summary>
    ''' <param name="caption">Nội dung câu hỏi.</param>
    ''' <returns>0 hoặc 1 theo lựa chọn của người dùng.</returns>
    Friend Function HdrYNQ(caption As String) As Double
        Dim value As Double = HdrDWrng(caption)
        Do Until value = 0 OrElse value = 1
            value = HdrDErr(caption)
        Loop
        Return value
    End Function

    ''' <summary>
    ''' Ghi trực tiếp một giá trị vào ô Excel.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang thao tác.</param>
    ''' <param name="cell">Địa chỉ ô Excel.</param>
    ''' <param name="value">Giá trị cần ghi.</param>
    Friend Sub DctVal(xlApp As Excel.Application, cell As String, value As Object)
        SetRangeValue(xlApp, cell, value)
    End Sub

    ''' <summary>
    ''' Ghi giá trị vào ô Excel và tô màu ô để đánh dấu dữ liệu được chỉnh tay.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang thao tác.</param>
    ''' <param name="cell">Địa chỉ ô Excel.</param>
    ''' <param name="value">Giá trị cần ghi.</param>
    Private Sub ModVal(xlApp As Excel.Application, cell As String, value As Object)
        SetRangeValue(xlApp, cell, value, highlight:=True)
    End Sub

    ''' <summary>
    ''' Xóa nội dung của ô hoặc toàn bộ vùng merge chứa ô đó.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang thao tác.</param>
    ''' <param name="cell">Địa chỉ ô Excel.</param>
    Friend Sub ClrVal(xlApp As Excel.Application, cell As String)
        Dim target As Excel.Range = Nothing
        Dim mergeArea As Excel.Range = Nothing

        Try
            target = CType(xlApp.Range(cell), Excel.Range)
            mergeArea = CType(target.MergeArea, Excel.Range)
            mergeArea.ClearContents()
        Finally
            If mergeArea IsNot Nothing AndAlso Not Object.ReferenceEquals(mergeArea, target) Then
                ReleaseComObject(mergeArea)
            End If
            ReleaseComObject(target)
        End Try
    End Sub

    ''' <summary>
    ''' Nhận chuỗi từ console rồi ghi vào một ô Excel.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang thao tác.</param>
    ''' <param name="caption">Nhãn nhập liệu hiển thị trên console.</param>
    ''' <param name="cell">Địa chỉ ô Excel.</param>
    Friend Sub PubSVal(xlApp As Excel.Application, caption As String, cell As String)
        DctVal(xlApp, cell, DtlSInp(caption))
    End Sub

    ''' <summary>
    ''' Chỉ ghi số lượng vào Excel khi giá trị lớn hơn 0.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang thao tác.</param>
    ''' <param name="cell">Địa chỉ ô Excel.</param>
    ''' <param name="value">Số lượng người dùng nhập.</param>
    Friend Sub PubDVal(xlApp As Excel.Application, cell As String, value As Double)
        If value > 0 Then
            DctVal(xlApp, cell, value)
        End If
    End Sub

    ''' <summary>
    ''' Ghi tên, trọng lượng và số lượng cốt thép vào dòng Excel tương ứng.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang thao tác.</param>
    ''' <param name="row">Dòng Excel cần ghi.</param>
    ''' <param name="name">Tên hoặc quy cách cốt thép.</param>
    ''' <param name="weight">Trọng lượng đơn vị.</param>
    ''' <param name="number">Số lượng người dùng nhập.</param>
    Friend Sub PubDModVal(xlApp As Excel.Application, row As String, name As String, weight As Double, number As Double)
        PubDModCells(xlApp, row, number, name:=name, weight:=weight)
    End Sub

    ''' <summary>
    ''' Ghi tiêu đề, tên, trọng lượng và số lượng cốt thép vào dòng Excel tương ứng.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang thao tác.</param>
    ''' <param name="row">Dòng Excel cần ghi.</param>
    ''' <param name="title">Tiêu đề nhóm cốt thép.</param>
    ''' <param name="name">Tên hoặc quy cách cốt thép.</param>
    ''' <param name="weight">Trọng lượng đơn vị.</param>
    ''' <param name="number">Số lượng người dùng nhập.</param>
    Friend Sub PubDModVal(xlApp As Excel.Application, row As String, title As String, name As String, weight As Double, number As Double)
        PubDModCells(xlApp, row, number, title:=title, name:=name, weight:=weight)
    End Sub

    ''' <summary>
    ''' Ghi đường kính, tiêu đề, tên, trọng lượng và số lượng cốt thép vào dòng Excel tương ứng.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang thao tác.</param>
    ''' <param name="row">Dòng Excel cần ghi.</param>
    ''' <param name="d">Đường kính cốt thép.</param>
    ''' <param name="title">Tiêu đề nhóm cốt thép.</param>
    ''' <param name="name">Tên hoặc quy cách cốt thép.</param>
    ''' <param name="weight">Trọng lượng đơn vị.</param>
    ''' <param name="number">Số lượng người dùng nhập.</param>
    Friend Sub PubDModVal(xlApp As Excel.Application, row As String, d As String, title As String, name As String, weight As Double, number As Double)
        PubDModCells(xlApp, row, number, d:=d, title:=title, name:=name, weight:=weight)
    End Sub

    ''' <summary>
    ''' Ghi đường kính, tiêu đề, tên, trọng lượng, đơn giá và số lượng cốt thép vào dòng Excel tương ứng.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang thao tác.</param>
    ''' <param name="row">Dòng Excel cần ghi.</param>
    ''' <param name="d">Đường kính cốt thép.</param>
    ''' <param name="title">Tiêu đề nhóm cốt thép.</param>
    ''' <param name="name">Tên hoặc quy cách cốt thép.</param>
    ''' <param name="weight">Trọng lượng đơn vị.</param>
    ''' <param name="price">Đơn giá áp dụng.</param>
    ''' <param name="number">Số lượng người dùng nhập.</param>
    Friend Sub PubDModVal(xlApp As Excel.Application, row As String, d As String, title As String, name As String, weight As Double, price As Double, number As Double)
        PubDModCells(xlApp, row, number, d:=d, title:=title, name:=name, weight:=weight, price:=price)
    End Sub

    ''' <summary>
    ''' Ghi dữ liệu cốt thép theo từng cột chuẩn của template Excel khi số lượng lớn hơn 0.
    ''' </summary>
    Private Sub PubDModCells(xlApp As Excel.Application, row As String, number As Double, Optional d As String = Nothing, Optional title As String = Nothing, Optional name As String = Nothing, Optional weight As Nullable(Of Double) = Nothing, Optional price As Nullable(Of Double) = Nothing)
        If number <= 0 Then
            Return
        End If

        If Not String.IsNullOrEmpty(d) Then
            DctVal(xlApp, $"S{row}", d)
        End If
        If Not String.IsNullOrEmpty(title) Then
            DctVal(xlApp, $"X{row}", title)
        End If
        If Not String.IsNullOrEmpty(name) Then
            DctVal(xlApp, $"AH{row}", name)
        End If
        If weight.HasValue Then
            ModVal(xlApp, $"CM{row}", weight.Value)
        End If
        If price.HasValue Then
            ModVal(xlApp, $"CQ{row}", price.Value)
        End If
        DctVal(xlApp, $"BA{row}", number)
    End Sub

    ''' <summary>
    ''' Ghi giá trị vào ô Excel, tùy chọn tô màu cho ô được chỉnh.
    ''' </summary>
    Private Sub SetRangeValue(xlApp As Excel.Application, cell As String, value As Object, Optional highlight As Boolean = False)
        Dim target As Excel.Range = Nothing
        Dim interior As Excel.Interior = Nothing

        Try
            target = CType(xlApp.Range(cell), Excel.Range)
            target.FormulaR1C1 = value
            If highlight Then
                interior = target.Interior
                interior.Color = RGB(0, 176, 240)
            End If
        Finally
            ReleaseComObject(interior)
            ReleaseComObject(target)
        End Try
    End Sub

    ''' <summary>
    ''' Giải phóng COM object để hạn chế Excel bị treo lại trong nền.
    ''' </summary>
    Private Sub ReleaseComObject(comObject As Object)
        If comObject Is Nothing OrElse Not Marshal.IsComObject(comObject) Then
            Return
        End If

        Try
            Marshal.FinalReleaseComObject(comObject)
        Catch ex As ArgumentException
        End Try
    End Sub
#End Region

#Region "Timer"
    ''' <summary>
    ''' Bắt đầu timer nếu timer chưa chạy.
    ''' </summary>
    <Extension()>
    Friend Sub StrtAdv(tmr As Timer)
        If Not tmr.Enabled Then
            tmr.Start()
        End If
    End Sub

    ''' <summary>
    ''' Dừng timer nếu timer đang chạy.
    ''' </summary>
    <Extension()>
    Friend Sub StopAdv(tmr As Timer)
        If tmr.Enabled Then
            tmr.Stop()
        End If
    End Sub
#End Region

#Region "Actor"
    ''' <summary>
    ''' In tiêu đề cảnh báo bằng màu vàng đậm.
    ''' </summary>
    ''' <param name="caption">Nội dung cần in.</param>
    Private Sub HdrSty(caption As String)
        ForegroundColor = DarkYellow
        Write(caption)
    End Sub

    ''' <summary>
    ''' In phần giới thiệu bằng màu xanh dương.
    ''' </summary>
    ''' <param name="caption">Nội dung cần in.</param>
    Private Sub IntroSty(caption As String)
        ForegroundColor = Blue
        Write(caption)
    End Sub

    ''' <summary>
    ''' In tên chương trình bằng màu xanh lá.
    ''' </summary>
    ''' <param name="caption">Nội dung cần in.</param>
    Private Sub TitSty(caption As String)
        ForegroundColor = Green
        Write(caption)
    End Sub

    ''' <summary>
    ''' In nhãn nhập liệu bằng màu xanh cyan.
    ''' </summary>
    ''' <param name="caption">Nội dung cần in.</param>
    Private Sub InpSty(caption As String)
        ForegroundColor = Cyan
        Write(caption)
    End Sub

    ''' <summary>
    ''' In mô tả bổ sung bằng màu tím.
    ''' </summary>
    ''' <param name="caption">Nội dung cần in.</param>
    Private Sub DescSty(caption As String)
        ForegroundColor = Magenta
        Write(caption)
    End Sub

    ''' <summary>
    ''' In lựa chọn hoặc cảnh báo bằng màu vàng.
    ''' </summary>
    ''' <param name="caption">Nội dung cần in.</param>
    Private Sub WrngSty(caption As String)
        ForegroundColor = Yellow
        Write(caption)
    End Sub

    ''' <summary>
    ''' In thông báo lỗi bằng màu đỏ.
    ''' </summary>
    ''' <param name="caption">Nội dung cần in.</param>
    Friend Sub ErrSty(caption As String)
        ForegroundColor = Red
        Write(caption)
    End Sub

    ''' <summary>
    ''' In tiền tố nhập liệu rồi trả màu chữ về trắng.
    ''' </summary>
    ''' <param name="caption">Nội dung tiền tố.</param>
    Private Sub PrefInp(caption As String)
        InpSty(caption)
        ForegroundColor = White
    End Sub

    ''' <summary>
    ''' In tiền tố lựa chọn rồi trả màu chữ về trắng.
    ''' </summary>
    ''' <param name="caption">Nội dung tiền tố.</param>
    Private Sub PrefSel(caption As String)
        WrngSty(caption)
        ForegroundColor = White
    End Sub

    ''' <summary>
    ''' In tiền tố lỗi nhập liệu với phần nhập màu đỏ.
    ''' </summary>
    ''' <param name="caption">Nội dung tiền tố.</param>
    Private Sub PrefWrng(caption As String)
        WrngSty(caption)
        ForegroundColor = Red
    End Sub

    ''' <summary>
    ''' In mô tả phụ sau nhãn nhập liệu.
    ''' </summary>
    ''' <param name="description">Mô tả phụ cần in.</param>
    Private Sub SfxDesc(description As String)
        DescSty(description)
        PrefInp(": ")
    End Sub

    ''' <summary>
    ''' Vẽ lại màn hình console giới thiệu trước mỗi nhóm câu hỏi.
    ''' </summary>
    Private Sub Intro()
        Clear()
        IntroSty(My.Resources.gr_name & vbCrLf)
        IntroSty(My.Resources.cc_text & vbCrLf)
        TitSty(vbCrLf & My.Resources.app_true_name & vbCrLf & vbCrLf)
    End Sub

    ''' <summary>
    ''' Hiển thị màn hình đầu nhóm rồi nhận số từ người dùng.
    ''' </summary>
    ''' <param name="caption">Nhãn nhập liệu.</param>
    ''' <returns>Giá trị số người dùng nhập.</returns>
    Friend Function HdrDInp(caption As String) As Double
        Intro()
        Return DtlDInp(caption)
    End Function

    ''' <summary>
    ''' Hiển thị màn hình đầu nhóm kèm thông báo dạng cảnh báo.
    ''' </summary>
    ''' <param name="caption">Nội dung cảnh báo.</param>
    Friend Sub HdrWrng(caption As String)
        Intro()
        WrngSty(caption)
    End Sub

    ''' <summary>
    ''' Hiển thị màn hình đầu nhóm rồi nhận lựa chọn dạng số.
    ''' </summary>
    ''' <param name="caption">Nhãn lựa chọn.</param>
    ''' <returns>Giá trị số người dùng nhập.</returns>
    Friend Function HdrDWrng(caption As String) As Double
        Intro()
        PrefSel(caption)
        Return Val(ReadLine())
    End Function

    ''' <summary>
    ''' Hiển thị lại câu hỏi lỗi rồi nhận lựa chọn dạng số.
    ''' </summary>
    ''' <param name="caption">Nhãn lựa chọn.</param>
    ''' <returns>Giá trị số người dùng nhập.</returns>
    Friend Function HdrDErr(caption As String) As Double
        Intro()
        PrefWrng(caption)
        Return Val(ReadLine())
    End Function

    ''' <summary>
    ''' Nhận số từ console ở cùng màn hình hiện tại.
    ''' </summary>
    ''' <param name="caption">Nhãn nhập liệu.</param>
    ''' <returns>Giá trị số người dùng nhập.</returns>
    Friend Function DtlDInp(caption As String) As Double
        PrefInp(caption)
        Return Val(ReadLine())
    End Function

    ''' <summary>
    ''' Nhận chuỗi từ console ở cùng màn hình hiện tại.
    ''' </summary>
    ''' <param name="caption">Nhãn nhập liệu.</param>
    ''' <returns>Chuỗi người dùng nhập.</returns>
    Friend Function DtlSInp(caption As String) As String
        PrefInp(caption)
        Return ReadLine()
    End Function

    ''' <summary>
    ''' Nhận số từ console với phần mô tả phụ ở cuối nhãn.
    ''' </summary>
    ''' <param name="caption">Nhãn nhập liệu.</param>
    ''' <param name="description">Mô tả phụ.</param>
    ''' <returns>Giá trị số người dùng nhập.</returns>
    Friend Function DtlDInpDesc(caption As String, description As String) As Double
        InpSty(caption)
        SfxDesc(description)
        Return Val(ReadLine())
    End Function
#End Region
End Module
