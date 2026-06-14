Imports Microsoft.Office.Interop.Excel

''' <summary>
''' Tập hợp các hàm nghiệp vụ ghi số lượng vật tư/thép vào đúng ô của template Ibaraki PCa.
''' </summary>
Friend Module Util
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

    ''' <summary>
    ''' Nhập số lượng cho nhóm slab hook type D13.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang thao tác.</param>
    ''' <param name="chosen">Giá trị lựa chọn 1/0.</param>
    Friend Sub SlabHookType(xlApp As Application, chosen As Double)
        If chosen = 1 Then
            PubDVal(xlApp, "BA35", DtlDInp(vbTab & "95×5405: "))
            PubDVal(xlApp, "BA36", DtlDInp(vbTab & "95×4905: "))
            PubDVal(xlApp, "BA37", DtlDInp(vbTab & "95×4405: "))
            PubDVal(xlApp, "BA38", DtlDInp(vbTab & "95×3905: "))
            PubDVal(xlApp, "BA39", DtlDInp(vbTab & "95×3405: "))
            PubDVal(xlApp, "BA40", DtlDInp(vbTab & "95×2905: "))
            PubDVal(xlApp, "BA41", DtlDInp(vbTab & "95×2405: "))
            PubDVal(xlApp, "BA42", DtlDInp(vbTab & "95×1905: "))
            PubDVal(xlApp, "BA43", DtlDInp(vbTab & "95×1405: "))
            PubDVal(xlApp, "BA44", DtlDInp(vbTab & "95× 905: "))
        End If
    End Sub

    ''' <summary>
    ''' Nhập số lượng cho nhóm slab L type D13.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang thao tác.</param>
    ''' <param name="chosen">Giá trị lựa chọn 1/0.</param>
    Friend Sub SlabLType(xlApp As Application, chosen As Double)
        If chosen = 1 Then
            PubDVal(xlApp, "BA45", DtlDInp(vbTab & "180×5320: "))
            PubDVal(xlApp, "BA46", DtlDInp(vbTab & "180×4820: "))
            PubDVal(xlApp, "BA47", DtlDInp(vbTab & "180×4320: "))
            PubDVal(xlApp, "BA48", DtlDInp(vbTab & "180×3820: "))
            PubDVal(xlApp, "BA49", DtlDInp(vbTab & "180×3320: "))
            PubDVal(xlApp, "BA50", DtlDInp(vbTab & "180×2820: "))
            PubDVal(xlApp, "BA51", DtlDInp(vbTab & "180×2320: "))
            PubDVal(xlApp, "BA52", DtlDInp(vbTab & "180×1820: "))
            PubDVal(xlApp, "BA53", DtlDInp(vbTab & "180×1320: "))
            PubDVal(xlApp, "BA54", DtlDInp(vbTab & "180× 820: "))
        End If
    End Sub

    ''' <summary>
    ''' Nhập số lượng cho nhóm slab thẳng D13.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang thao tác.</param>
    ''' <param name="chosen">Giá trị lựa chọn 1/0.</param>
    Friend Sub SlabStr(xlApp As Application, chosen As Double)
        If chosen = 1 Then
            PubDVal(xlApp, "BA55", DtlDInp(vbTab & "5500: "))
            PubDVal(xlApp, "BA56", DtlDInp(vbTab & "5000: "))
            PubDVal(xlApp, "BA57", DtlDInp(vbTab & "4500: "))
            PubDVal(xlApp, "BA58", DtlDInp(vbTab & "4000: "))
            PubDVal(xlApp, "BA59", DtlDInp(vbTab & "3500: "))
            PubDVal(xlApp, "BA60", DtlDInp(vbTab & "3000: "))
            PubDVal(xlApp, "BA61", DtlDInp(vbTab & "2500: "))
            PubDVal(xlApp, "BA62", DtlDInp(vbTab & "2000: "))
            PubDVal(xlApp, "BA63", DtlDInp(vbTab & "1500: "))
            PubDModVal(xlApp, "64", "1300", 1.4, DtlDInp(vbTab & "1300: "))
            PubDVal(xlApp, "BA65", DtlDInp(vbTab & "1000: "))
            PubDVal(xlApp, "BA66", DtlDInp(vbTab & " 900: "))
        End If
    End Sub

    ''' <summary>
    ''' Nhập số lượng cho nhóm slab gia cường hook type D10.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang thao tác.</param>
    ''' <param name="chosen">Giá trị lựa chọn 1/0.</param>
    Friend Sub SlabReinfHookType(xlApp As Application, chosen As Double)
        If chosen = 1 Then
            PubDModVal(xlApp, "67", "95×5405", 3.2, DtlDInp(vbTab & "95×5405: "))
            PubDModVal(xlApp, "68", "95×4905", 2.9, DtlDInp(vbTab & "95×4905: "))
            PubDModVal(xlApp, "69", "95×4405", 2.6, DtlDInp(vbTab & "95×4405: "))
            PubDModVal(xlApp, "70", "95×3905", 2.4, DtlDInp(vbTab & "95×3905: "))
            PubDModVal(xlApp, "71", "95×3405", 2.1, DtlDInp(vbTab & "95×3405: "))
            PubDModVal(xlApp, "72", "95×2905", 1.8, DtlDInp(vbTab & "95×2905: "))
            PubDModVal(xlApp, "73", "95×2405", 1.5, DtlDInp(vbTab & "95×2405: "))
            PubDModVal(xlApp, "74", "95×1905", 1.2, DtlDInp(vbTab & "95×1905: "))
            PubDModVal(xlApp, "75", "95×1405", 0.9, DtlDInp(vbTab & "95×1405: "))
            PubDModVal(xlApp, "76", "95× 905", 0.6, DtlDInp(vbTab & "95× 905: "))
        End If
    End Sub

    ''' <summary>
    ''' Nhập số lượng cho nhóm slab gia cường thẳng D10.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang thao tác.</param>
    ''' <param name="chosen">Giá trị lựa chọn 1/0.</param>
    Friend Sub SlabReinfStr(xlApp As Application, chosen As Double)
        If chosen = 1 Then
            PubDVal(xlApp, "BA77", DtlDInp(vbTab & "5500: "))
            PubDVal(xlApp, "BA78", DtlDInp(vbTab & "5000: "))
            PubDVal(xlApp, "BA79", DtlDInp(vbTab & "4500: "))
            PubDVal(xlApp, "BA80", DtlDInp(vbTab & "4000: "))
            PubDVal(xlApp, "BA81", DtlDInp(vbTab & "3500: "))
            PubDVal(xlApp, "BA82", DtlDInp(vbTab & "3000: "))
            PubDVal(xlApp, "BA83", DtlDInp(vbTab & "2500: "))
            PubDVal(xlApp, "BA84", DtlDInp(vbTab & "2000: "))
            PubDVal(xlApp, "BA85", DtlDInp(vbTab & "1500: "))
            PubDVal(xlApp, "BA86", DtlDInp(vbTab & "1000: "))
        End If
    End Sub

    ''' <summary>
    ''' Nhập nhóm thép dưới D13 và dòng bổ sung 750x1250.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang thao tác.</param>
    Friend Sub LwrEndD13(xlApp As Application)
        PubDVal(xlApp, "BA110", DtlDInp(vbTab & "    5500: "))
        PubDVal(xlApp, "BA111", DtlDInp(vbTab & "    5000: "))
        PubDVal(xlApp, "BA112", DtlDInp(vbTab & "    4500: "))
        PubDVal(xlApp, "BA113", DtlDInp(vbTab & "    4000: "))
        PubDVal(xlApp, "BA114", DtlDInp(vbTab & "    3500: "))
        PubDVal(xlApp, "BA115", DtlDInp(vbTab & "    3000: "))
        PubDVal(xlApp, "BA116", DtlDInp(vbTab & "    2500: "))
        PubDVal(xlApp, "BA117", DtlDInp(vbTab & "    2000: "))
        PubDVal(xlApp, "BA118", DtlDInp(vbTab & "    1500: "))
        PubDVal(xlApp, "BA119", DtlDInp(vbTab & "    1000: "))
        PubDModVal(xlApp, "120", "D13", "（下部主筋）", "750×1250", 2.1, DtlDInp(vbTab & "750×1250: "))
    End Sub

    ''' <summary>
    ''' Nhập nhóm thép dưới D16, kèm đơn giá cho các dòng có cấu hình giá.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang thao tác.</param>
    ''' <param name="chosen">Giá trị lựa chọn 1/0.</param>
    Friend Sub LwrEndD16(xlApp As Application, chosen As Double)
        If chosen = 1 Then
            PubDModVal(xlApp, "97", "D16", "（下部主筋　直筋）", "5500", 8.6, My.Settings.Pr_D16, DtlDInp(vbTab & "5500: "))
            PubDModVal(xlApp, "98", "D16", "（下部主筋　直筋）", "5000", 7.8, My.Settings.Pr_D16, DtlDInp(vbTab & "5000: "))
            PubDModVal(xlApp, "99", "D16", "（下部主筋　直筋）", "4500", 7.1, My.Settings.Pr_D16, DtlDInp(vbTab & "4500: "))
            PubDModVal(xlApp, "100", "D16", "（下部主筋　直筋）", "4000", 6.3, My.Settings.Pr_D16, DtlDInp(vbTab & "4000: "))
            PubDModVal(xlApp, "101", "D16", "（下部主筋　直筋）", "3500", 5.5, My.Settings.Pr_D16, DtlDInp(vbTab & "3500: "))
            PubDModVal(xlApp, "102", "D16", "（下部主筋　直筋）", "3000", 4.9, DtlDInp(vbTab & "3000: "))
            PubDModVal(xlApp, "103", "D16", "（下部主筋　直筋）", "2500", 4.1, DtlDInp(vbTab & "2500: "))
            PubDModVal(xlApp, "104", "D16", "（下部主筋　直筋）", "2000", 3.3, DtlDInp(vbTab & "2000: "))
            PubDModVal(xlApp, "105", "D16", "（下部主筋　直筋）", "1500", 2.5, DtlDInp(vbTab & "1500: "))
            PubDModVal(xlApp, "106", "D16", "（下部主筋　直筋）", "1000", 1.7, DtlDInp(vbTab & "1000: "))
        End If
    End Sub

    ''' <summary>
    ''' Nhập số lượng cho nhóm thép đầu biên D10.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang thao tác.</param>
    ''' <param name="chosen">Giá trị lựa chọn 1/0.</param>
    Friend Sub Edge(xlApp As Application, chosen As Double)
        If chosen = 1 Then
            PubDVal(xlApp, "BA87", DtlDInp(vbTab & "5500: "))
            PubDVal(xlApp, "BA88", DtlDInp(vbTab & "5000: "))
            PubDVal(xlApp, "BA89", DtlDInp(vbTab & "4500: "))
            PubDVal(xlApp, "BA90", DtlDInp(vbTab & "4000: "))
            PubDVal(xlApp, "BA91", DtlDInp(vbTab & "3500: "))
            PubDVal(xlApp, "BA92", DtlDInp(vbTab & "3000: "))
            PubDVal(xlApp, "BA93", DtlDInp(vbTab & "2500: "))
            PubDVal(xlApp, "BA94", DtlDInp(vbTab & "2000: "))
            PubDVal(xlApp, "BA95", DtlDInp(vbTab & "1500: "))
            PubDVal(xlApp, "BA96", DtlDInp(vbTab & "1000: "))
        End If
    End Sub

    ''' <summary>
    ''' Nhập số lượng cho nhóm thép góc.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang thao tác.</param>
    Friend Sub JtCor(xlApp As Application)
        PubDModVal(xlApp, "121", "D16", "（コーナー）", "900×900", 2.9, DtlDInp(vbTab & "D16: "))
        PubDVal(xlApp, "BA123", DtlDInp(vbTab & "D13: "))
        PubDVal(xlApp, "BA122", DtlDInp(vbTab & "D10: "))
    End Sub

    ''' <summary>
    ''' Nhập số lượng cho nhóm haunch H250.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang thao tác.</param>
    ''' <param name="chosen">Giá trị lựa chọn 1/0.</param>
    Friend Sub Haunch(xlApp As Application, chosen As Double)
        If chosen = 1 Then
            PubDModVal(xlApp, "126", "900×曲（H250）×900", 3.5, DtlDInp(vbTab & "D16: "))
            PubDModVal(xlApp, "127", "D13", "750×曲（H250）×750", 2, DtlDInp(vbTab & "D13: "))
            PubDModVal(xlApp, "125", "D10", "（斜筋）", "600×曲（H250）×600", 0.9, DtlDInp(vbTab & "D10: "))
        End If
    End Sub

    ''' <summary>
    ''' Ghi hoặc xóa số lượng máy nước nóng điện.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang thao tác.</param>
    ''' <param name="value">Số lượng máy nước nóng điện.</param>
    Friend Sub ElecWtrHtr(xlApp As Application, value As Double)
        If value > 0 Then
            DctVal(xlApp, "BA30", value)
        Else
            ClrVal(xlApp, "BA30")
        End If
    End Sub

    ''' <summary>
    ''' Nhập thông tin công trình và danh sách vật tư phụ.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang thao tác.</param>
    Friend Sub Parts(xlApp As Application)
        Dim name = $"{DtlSInp(vbTab & "邸名" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & ": ")}様邸"
        DctVal(xlApp, "BJ13", name)
        CType(xlApp.ActiveSheet, Worksheet).Name = name
        PubSVal(xlApp, vbTab & "邸名コード" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & ": ", "AD6")
        PubSVal(xlApp, vbTab & "納品日" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & ": ", "BO3")
        PubSVal(xlApp, vbTab & "住所" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & ": ", "BJ14")
        PubDVal(xlApp, "BA149", DtlDInpDesc(vbTab & "カットスクリュー・Ⅱ (袋)", vbTab & vbTab & vbTab & "[M12ﾖｳ 50ｺｲﾘ]" & vbTab))
        PubDVal(xlApp, "BA140", DtlDInpDesc(vbTab & "PCa基礎梁 敷き鉄板 (枚)", vbTab & vbTab & vbTab & vbTab & "[1.6t×200×300]" & vbTab))
        PubDVal(xlApp, "BA141", DtlDInpDesc(vbTab & "PCa基礎梁 敷き鉄板 (枚)", vbTab & vbTab & vbTab & vbTab & "[3.2t×200×300]" & vbTab))
        PubDVal(xlApp, "BA142", DtlDInpDesc(vbTab & "PCa基礎梁 土台用アンカー (本)", vbTab & vbTab & vbTab & "[M12×147]" & vbTab))
        PubDVal(xlApp, "BA143", DtlDInpDesc(vbTab & "PCa基礎梁 ホールダウンカンカーボルト (本)", vbTab & "[M12×387]" & vbTab))
        PubDVal(xlApp, "BA145", DtlDInpDesc(vbTab & "PCa基礎梁 ホールダウンカンカーボルト (本)", vbTab & "[M12×170]" & vbTab))
        PubDVal(xlApp, "BA153", DtlDInpDesc(vbTab & "マグネット差筋アンカーD13 (ｾｯﾄ)", vbTab & vbTab & vbTab & "[直]" & vbTab & vbTab))
        PubDVal(xlApp, "BA154", DtlDInpDesc(vbTab & "マグネット差筋アンカーD13 (ｾｯﾄ)", vbTab & vbTab & vbTab & "[曲]" & vbTab & vbTab))
        PubDVal(xlApp, "BA151", DtlDInpDesc(vbTab & "排水用スリーブホルダー・D10用 (箱)", vbTab & vbTab & "[50ﾊﾟｲ 75ﾊﾟｲ]" & vbTab))
        PubDVal(xlApp, "BA150", DtlDInpDesc(vbTab & "給水用スリーブホルダー・D10用 (箱)", vbTab & vbTab & "[50ﾊﾟｲ]" & vbTab & vbTab))
        PubDVal(xlApp, "BA144", DtlDInpDesc(vbTab & "PCa基礎梁 BF用柱脚両ネジボルト (本)", vbTab & vbTab & "[M16×70]" & vbTab))
        Dim curingShRingTree = DtlDInpDesc(vbTab & "養生シート輪木 (ｾｯﾄ)", vbTab & vbTab & vbTab & vbTab & "[3.6×5.4]" & vbTab)
        If curingShRingTree > 0 Then
            DctVal(xlApp, "BA155", curingShRingTree)
        Else
            DctVal(xlApp, "BA155", 1)
            ClrVal(xlApp, "BF155")
            ClrVal(xlApp, "CB155")
        End If
        PubDVal(xlApp, "BA156", DtlDInp(vbTab & "カットスクリューⅡ・専用ピット (個)" & vbTab & vbTab & vbTab & vbTab & ": "))
        PubDVal(xlApp, "BA146", DtlDInpDesc(vbTab & "スペーサーブロック (個)", vbTab & vbTab & vbTab & vbTab & "[H60]" & vbTab & vbTab))
        PubDVal(xlApp, "BA147", DtlDInpDesc(vbTab & "スペーサーブロック (個)", vbTab & vbTab & vbTab & vbTab & "[H80]" & vbTab & vbTab))
        ' Nhóm vật tư mở rộng so với form cơ bản.
        PubDVal(xlApp, "BA152", DtlDInpDesc(vbTab & "アンカーボルトセット (ｾｯﾄ)", vbTab & vbTab & vbTab & "[M18×380]" & vbTab))
        PubDVal(xlApp, "BA157", DtlDInpDesc(vbTab & "Ｕボルト (ｾｯﾄ)", vbTab & vbTab & vbTab & vbTab & vbTab & "[M8]" & vbTab & vbTab))
    End Sub
End Module
