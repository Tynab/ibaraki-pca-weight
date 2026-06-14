Imports System.Environment
Imports System.Environment.SpecialFolder

''' <summary>
''' Khai báo các hằng số và đường dẫn dùng chung cho Excel và bộ cài cập nhật.
''' </summary>
Friend Module Constant
    Friend Const XL_NAME As String = "excel"
    Friend ReadOnly FILE_SETUP_NAME As String = $"{My.Resources.app_name} Setup.msi"
    Friend ReadOnly BACK_PATH As String = GetFolderPath(ApplicationData)
    Friend ReadOnly FRNT_PATH As String = System.IO.Path.Combine(BACK_PATH, My.Resources.co_name)
    Friend ReadOnly FILE_SETUP_ADR As String = System.IO.Path.Combine(FRNT_PATH, FILE_SETUP_NAME)
End Module
