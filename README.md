# TOUHOKU (PCA-BF) WEIGHT

Ứng dụng hỗ trợ đội 西山 thuộc エマールグループ nhập và chuyển dữ liệu nhanh hơn cho mẫu trọng lượng 茨城 (プレキャス - BF) từ đối tác 文化シャッター.

## Icon

<p align='center'>
<img src='pic/0.png'></img>
</p>

## Thành Phần Chính

| Thành phần | Tệp | Mô tả |
|---|---|---|
| `Main` | `Ibaraki Pca Weight/Script/Main.vb` | Điểm vào ứng dụng: thiết lập console UTF-8, kiểm tra serial bản quyền và gọi `RunApp`. |
| `Common` | `Ibaraki Pca Weight/Script/Common.vb` | Hàm dùng chung: kiểm tra cập nhật, thao tác Excel COM, định dạng console, hiệu ứng biểu mẫu, quản lý tiến trình và file. |
| `Constant` | `Ibaraki Pca Weight/Script/Constant.vb` | Hằng số toàn ứng dụng: tên tiến trình Excel, tên file cài đặt và đường dẫn trong AppData. |
| `Service` | `Ibaraki Pca Weight/Script/Service.vb` | Điều phối luồng `WtIbarakiPca`, chạy lần lượt toàn bộ nhóm câu hỏi nhập liệu. |
| `Util` | `Ibaraki Pca Weight/Script/Util.vb` | Logic nghiệp vụ: mỗi nhóm nhập liệu được ánh xạ tới các ô Excel tương ứng. |
| `FrmUpdate` | `Ibaraki Pca Weight/Control/FrmUpdate.vb` | Biểu mẫu cập nhật: tải file MSI, hiển thị tiến độ, chặn Alt+F4 và chạy trình cài đặt khi tải xong. |

## Luồng Chạy

```text
Main()
 ├─ ValidateLicense()        ← kiểm tra serial trong tài nguyên
 └─ RunApp()
     ├─ ChkUpd()            ← kiểm tra phiên bản qua mạng; mở FrmUpdate nếu có bản mới
     ├─ KillXl()            ← kết thúc các tiến trình excel.exe đang mở
     ├─ OpenFileDialog      ← người dùng chọn file *.xlsx / *.xls
     └─ ProcessWorkbook()
         ├─ Excel.Application (COM)
         ├─ WtIbarakiPca()  ← ghi toàn bộ nhóm dữ liệu vào sổ tính
         ├─ Workbook.Close(Save)
         └─ Process.Start() ← mở lại file đã lưu bằng ứng dụng mặc định
```

## Nhóm Nhập Liệu

| # | Câu hỏi | Ô Excel |
|---|---|---|
| 1 | 運賃 (2トン車) | BA108, BA109, BA158 |
| 2 | スラブフック型 D13 | BA35 - BA44 |
| 3 | スラブＬ型 D13 | BA45 - BA54 |
| 4 | スラブ直 D13 | BA55 - BA66 |
| 5 | スラブ補強フック型 D10 | BA67 - BA76 |
| 6 | スラブ補強直 D10 | BA77 - BA86 |
| 7 | 下端 D13, bắt buộc | BA110 - BA120 |
| 8 | 下端 D16 | BA97 - BA106 |
| 9 | 端部 D10 | BA87 - BA96 |
| 10 | スリーブ | BA124 |
| 11 | コーナー | BA121 - BA123 |
| 12 | 土間用さし | BA137 |
| 13 | Ｕ型 D16 | BA126 |
| 14 | ハンチ H250 | BA125 - BA127 |
| 15 | 深基礎用端部スラブ D10 | BA128 |
| 16 | 電気温水器 | BA30 |
| 17 | 副資材リスト, gồm vật tư phụ và thông tin công trình | BA140 - BA157, BF155, CB155, AD6, BO3, BJ13, BJ14 |

## Minh Họa Mã Nguồn

Trích nguyên văn từ `Ibaraki Pca Weight/Script/Util.vb`:

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

## Gói Phụ Thuộc

<img src='pic/1.png' align='left' width='3%' height='3%'></img>
<div style='display:flex;'>

- Microsoft.Office.Interop.Excel » 15.0.4795.1001

</div>
