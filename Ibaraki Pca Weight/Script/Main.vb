Imports System.Console
Imports System.Text.Encoding
Imports System.Windows.Forms
Imports System.Windows.Forms.DialogResult
Imports System.Windows.Forms.MessageBox
Imports System.Windows.Forms.MessageBoxButtons

''' <summary>
''' Entry point của ứng dụng console/WinForms dùng để kiểm tra license rồi chạy luồng xử lý Excel.
''' </summary>
Public Module Main
    ''' <summary>
    ''' Thiết lập encoding UTF-8, xác thực license nếu cần và khởi động ứng dụng chính.
    ''' </summary>
    Public Sub Main()
        OutputEncoding = UTF8
        If My.Settings.Chk_Key OrElse ValidateLicense() Then
            RunApp()
        End If
    End Sub

    ''' <summary>
    ''' Nhận serial từ người dùng và cho phép thử lại đến khi đúng hoặc người dùng hủy.
    ''' </summary>
    ''' <returns>True nếu license hợp lệ, ngược lại là False.</returns>
    Private Function ValidateLicense() As Boolean
        Do
            If InputBox("シリアルを入力", "ライセンスキー") = My.Resources.key_ser Then
                UpdVldLic()
                Return True
            End If

            If Show("ライセンスが間違っています！", "エラー", RetryCancel, MessageBoxIcon.Error) <> Retry Then
                ErrSty("終了するには、任意のキーを押してください...")
                ReadKey()
                Return False
            End If
        Loop
    End Function
End Module
