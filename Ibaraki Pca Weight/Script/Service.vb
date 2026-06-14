Imports Microsoft.Office.Interop.Excel

''' <summary>
''' Điều phối toàn bộ quy trình nhập số lượng và ghi dữ liệu vào template Excel Ibaraki PCa.
''' </summary>
Friend Module Service
    ''' <summary>
    ''' Chạy lần lượt các nhóm câu hỏi nghiệp vụ và cập nhật workbook Excel đang mở.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang thao tác.</param>
    Friend Sub WtIbarakiPca(xlApp As Application)
        ' Nhóm vận chuyển.
        Fare(xlApp, HdrYNQ(vbTab & vbTab & "運賃 (2トン車): "))
        ' Nhóm slab hook D13.
        SlabHookType(xlApp, HdrYNQ(vbTab & vbTab & "スラブフック型 (D13): "))
        ' Nhóm slab L D13.
        SlabLType(xlApp, HdrYNQ(vbTab & vbTab & "スラブＬ型 (D13): "))
        ' Nhóm slab thẳng D13.
        SlabStr(xlApp, HdrYNQ(vbTab & vbTab & "スラブ直 (D13): "))
        ' Nhóm slab gia cường hook D10.
        SlabReinfHookType(xlApp, HdrYNQ(vbTab & vbTab & "スラブ補強フック型 (D10): "))
        ' Nhóm slab gia cường thẳng D10.
        SlabReinfStr(xlApp, HdrYNQ(vbTab & vbTab & "スラブ補強直 (D10): "))
        ' Nhóm thép dưới D13 luôn được nhập.
        HdrWrng(vbTab & vbTab & "下端 (D13)" & vbCrLf)
        LwrEndD13(xlApp)
        ' Nhóm thép dưới D16 tùy chọn.
        LwrEndD16(xlApp, HdrYNQ(vbTab & vbTab & "下端 (D16): "))
        ' Nhóm thép đầu biên D10.
        Edge(xlApp, HdrYNQ(vbTab & vbTab & "端部 (D10): "))
        ' Số lượng sleeve.
        PubDVal(xlApp, "BA124", HdrDInp(vbTab & vbTab & "スリーブ: "))
        ' Nhóm thép góc.
        HdrWrng(vbTab & vbTab & "コーナー" & vbCrLf)
        JtCor(xlApp)
        ' Vật tư sàn đất.
        PubDVal(xlApp, "BA137", HdrDInp(vbTab & vbTab & "土間用さし: "))
        ' Nhóm U type D16.
        PubDModVal(xlApp, "126", "（Ｕノ字型）", "900×80×900", 3.1, HdrDInp(vbTab & vbTab & "Ｕ型 (D16): "))
        ' Nhóm haunch H250.
        Haunch(xlApp, HdrYNQ(vbTab & vbTab & "ハンチ (H250): "))
        ' End slab cho móng sâu.
        PubDModVal(xlApp, "128", "650×250　　フック付", 0.6, HdrDInp(vbTab & vbTab & "深基礎用端部スラブ (D10): "))
        ' Máy nước nóng điện.
        ElecWtrHtr(xlApp, HdrDInp(vbTab & vbTab & "電気温水器: "))
        ' Danh sách phụ kiện và vật tư phụ.
        HdrWrng(vbTab & vbTab & "副資材リスト" & vbCrLf)
        Parts(xlApp)
    End Sub
End Module
