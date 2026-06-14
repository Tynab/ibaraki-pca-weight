# Ibaraki PCa Weight

Ứng dụng VB.NET hỗ trợ team 西山 của エマール group nhập nhanh dữ liệu vật tư/trọng lượng cho mẫu Excel 茨城 (プレキャス) 重量 từ đối tác 文化シャッター. Chương trình chạy dạng console có dùng WinForms cho hộp license, chọn file Excel và màn hình cập nhật.

## Chức năng chính

- Kiểm tra license cục bộ trước khi chạy luồng xử lý.
- Kiểm tra phiên bản mới từ server cấu hình trong `My.Resources`.
- Yêu cầu đóng Excel, mở workbook người dùng chọn, nhập số lượng theo từng nhóm cấu kiện/vật tư và lưu lại workbook.
- Tự tải bộ cài `.msi` mới khi có phiên bản cập nhật.
- Ghi dữ liệu vào đúng ô của template Excel, đồng thời tô màu các ô trọng lượng/đơn giá được chỉnh tay.

## Luồng xử lý

1. `Script/Main.vb` thiết lập UTF-8 và xác thực license nếu user setting `Chk_Key` chưa hợp lệ.
2. `Script/Common.vb` kiểm tra cập nhật, yêu cầu đóng Excel, mở file qua `OpenFileDialog`, gọi service nghiệp vụ và dọn Excel COM object sau khi lưu.
3. `Script/Service.vb` điều phối thứ tự các nhóm nhập liệu theo template 茨城 PCa.
4. `Script/Util.vb` chứa mapping nghiệp vụ từ từng câu hỏi sang ô Excel cụ thể.
5. `Control/FrmUpdate.vb` tải installer mới, hiển thị tiến độ và chỉ chạy installer khi download hoàn tất thành công.

## Summary source tự viết

| File | Vai trò | Ghi chú cleanup |
| --- | --- | --- |
| `Script/Main.vb` | Entry point, kiểm tra license và gọi `RunApp()`. | Bỏ `GoTo`, chuyển sang vòng lặp `ValidateLicense()` rõ ràng hơn. |
| `Script/Constant.vb` | Khai báo tên process Excel, tên file setup và đường dẫn lưu installer. | Dùng `Path.Combine` thay cho nối chuỗi đường dẫn thủ công. |
| `Script/Common.vb` | Helper dùng chung cho update check, Excel Interop, nhập liệu console, timer và hiệu ứng form. | Thêm kiểu trả về rõ ràng, sửa điều kiện nhập 1/0, dispose `WebClient`, dọn workbook/Excel COM object, ghi ô Excel trực tiếp qua `Range` thay vì `Activate()/ActiveCell`. |
| `Script/Service.vb` | Điều phối các nhóm câu hỏi theo đúng thứ tự template. | Việt hóa comment và giữ nguyên thứ tự nghiệp vụ để tránh lệch form Excel. |
| `Script/Util.vb` | Mapping số lượng vật tư/thép sang các ô như `BA35`, `CM120`, `CQ97`. | Việt hóa XML summary/param, sửa typo `choosen` thành `chosen`, giữ nguyên label Nhật và địa chỉ ô. |
| `Control/FrmUpdate.vb` | Form tải bản cập nhật và chạy installer. | Chuyển logic đóng form sang `DownloadFileCompleted`, không chạy installer khi tải lỗi, dispose `WebClient`. |
| `My Project/AssemblyInfo.vb` | Metadata assembly và version. | Việt hóa comment metadata thủ công. |

Các file `*.Designer.vb`, `*.resx`, `Application.myapp`, `Settings.settings`, icon/gif/png và nội dung trong `packages/` được xem là designer/autogen hoặc asset nên không sửa comment thủ công.

## Code demo

```vb
''' <summary>
''' Ghi lựa chọn vận chuyển bằng xe 2 tấn và số dòng thép mặc định.
''' </summary>
''' <param name="xlApp">Ứng dụng Excel đang thao tác.</param>
''' <param name="chosen">Giá trị lựa chọn 1/0.</param>
Friend Sub Fare(xlApp As Application, chosen As Double)
    If chosen = 1 Then
        DctVal(xlApp, "BA158", chosen)
    End If
    DctVal(xlApp, "BA108", 5) ' D13
    DctVal(xlApp, "BA109", 3) ' D10
End Sub
```

## Build

Project target `.NET Framework 4.8.1`, dùng Visual Studio/MSBuild cho project VB.NET cũ:

```powershell
& 'C:\Program Files\Microsoft Visual Studio\18\Professional\MSBuild\Current\Bin\MSBuild.exe' '.\Ibaraki Pca Weight\Ibaraki Pca Weight.vbproj' /p:Configuration=Release /p:Platform=AnyCPU /m
```

`dotnet build` với SDK mới có thể lỗi resource non-string của WinForms/.NET Framework cũ; nên ưu tiên MSBuild đi kèm Visual Studio.

## Package

- `Microsoft.Office.Interop.Excel` 15.0.4795.1001

## Hình minh họa

<p align="center">
  <img src="pic/0.png" alt="Màn hình mask của Ibaraki PCa Weight">
</p>

<p align="center">
  <img src="pic/1.png" width="48" alt="Biểu tượng package">
</p>
