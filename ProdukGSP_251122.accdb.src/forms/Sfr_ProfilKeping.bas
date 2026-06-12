Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11520
    DatasheetFontHeight =11
    ItemSuffix =48
    RecSrcDt = Begin
        0xa45357f6aa43e640
    End
    RecordSource ="T_ProfilKeping"
    Caption ="Sfr_ProfilKeping"
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =60.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ComboBox
            AddColon = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =0
            Name ="FormHeader"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
        End
        Begin Section
            Height =8710
            Name ="Detail"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1890
                    Top =360
                    Height =300
                    ColumnWidth =1440
                    Name ="KepingAutoID"
                    ControlSource ="KepingAutoID"

                    LayoutCachedLeft =1890
                    LayoutCachedTop =360
                    LayoutCachedWidth =3330
                    LayoutCachedHeight =660
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =360
                            Width =1440
                            Height =310
                            Name ="KepingAutoID_Label"
                            Caption ="KepingAutoID"
                            LayoutCachedLeft =360
                            LayoutCachedTop =360
                            LayoutCachedWidth =1800
                            LayoutCachedHeight =670
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1890
                    Top =780
                    Width =4170
                    Height =580
                    ColumnWidth =3000
                    TabIndex =1
                    Name ="KepingID"
                    ControlSource ="KepingID"

                    LayoutCachedLeft =1890
                    LayoutCachedTop =780
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =1360
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =780
                            Width =1440
                            Height =310
                            Name ="KepingID_Label"
                            Caption ="KepingID"
                            LayoutCachedLeft =360
                            LayoutCachedTop =780
                            LayoutCachedWidth =1800
                            LayoutCachedHeight =1090
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1890
                    Top =1440
                    Width =1410
                    Height =310
                    ColumnWidth =1410
                    TabIndex =2
                    Name ="BlankoKepingAutoID"
                    ControlSource ="BlankoKepingAutoID"

                    LayoutCachedLeft =1890
                    LayoutCachedTop =1440
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =1750
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =1440
                            Width =1440
                            Height =310
                            Name ="BlankoKepingAutoID_Label"
                            Caption ="BlankoKepingAutoID"
                            LayoutCachedLeft =360
                            LayoutCachedTop =1440
                            LayoutCachedWidth =1800
                            LayoutCachedHeight =1750
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1890
                    Top =1860
                    Width =4170
                    Height =580
                    ColumnWidth =3000
                    TabIndex =3
                    Name ="DeskripsiKeping"
                    ControlSource ="DeskripsiProfilKeping"

                    LayoutCachedLeft =1890
                    LayoutCachedTop =1860
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =2440
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =1860
                            Width =1440
                            Height =310
                            Name ="DeskripsiKeping_Label"
                            Caption ="DeskripsiProfilKeping"
                            LayoutCachedLeft =360
                            LayoutCachedTop =1860
                            LayoutCachedWidth =1800
                            LayoutCachedHeight =2170
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1890
                    Top =2520
                    Width =4170
                    Height =580
                    ColumnWidth =3000
                    TabIndex =4
                    Name ="JenisRivet_deprecated"
                    ControlSource ="JenisRivet_deprecated"

                    LayoutCachedLeft =1890
                    LayoutCachedTop =2520
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =3100
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =2520
                            Width =1440
                            Height =310
                            Name ="JenisRivet_deprecated_Label"
                            Caption ="JenisRivet_deprecated"
                            LayoutCachedLeft =360
                            LayoutCachedTop =2520
                            LayoutCachedWidth =1800
                            LayoutCachedHeight =2830
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1890
                    Top =3180
                    Width =4170
                    Height =580
                    ColumnWidth =3000
                    TabIndex =5
                    Name ="UkuranRivet_deprecated"
                    ControlSource ="UkuranRivet_deprecated"

                    LayoutCachedLeft =1890
                    LayoutCachedTop =3180
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =3760
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =3180
                            Width =1440
                            Height =310
                            Name ="UkuranRivet_deprecated_Label"
                            Caption ="UkuranRivet_deprecated"
                            LayoutCachedLeft =360
                            LayoutCachedTop =3180
                            LayoutCachedWidth =1800
                            LayoutCachedHeight =3490
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1890
                    Top =3840
                    TabIndex =6
                    Name ="LubangFlag_deprecated"
                    ControlSource ="LubangFlag_deprecated"
                    StatusBarText ="'gunakan rivetID apabila tanpa lubang"

                    LayoutCachedLeft =1890
                    LayoutCachedTop =3840
                    LayoutCachedWidth =2150
                    LayoutCachedHeight =4080
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =3840
                            Width =1440
                            Height =310
                            Name ="LubangFlag_deprecated_Label"
                            Caption ="LubangFlag_deprecated"
                            LayoutCachedLeft =360
                            LayoutCachedTop =3840
                            LayoutCachedWidth =1800
                            LayoutCachedHeight =4150
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1890
                    Top =4260
                    Width =1410
                    Height =310
                    ColumnWidth =1410
                    TabIndex =7
                    Name ="UkuranLubang_deprecated"
                    ControlSource ="UkuranLubang_deprecated"

                    LayoutCachedLeft =1890
                    LayoutCachedTop =4260
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =4570
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =4260
                            Width =1440
                            Height =310
                            Name ="UkuranLubang_deprecated_Label"
                            Caption ="UkuranLubang_deprecated"
                            LayoutCachedLeft =360
                            LayoutCachedTop =4260
                            LayoutCachedWidth =1800
                            LayoutCachedHeight =4570
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1890
                    Top =4680
                    Width =1410
                    Height =310
                    ColumnWidth =1410
                    TabIndex =8
                    Name ="PolaLubangID"
                    ControlSource ="PolaLubangID"

                    LayoutCachedLeft =1890
                    LayoutCachedTop =4680
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =4990
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =4680
                            Width =1440
                            Height =310
                            Name ="PolaLubangID_Label"
                            Caption ="PolaLubangID"
                            LayoutCachedLeft =360
                            LayoutCachedTop =4680
                            LayoutCachedWidth =1800
                            LayoutCachedHeight =4990
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1890
                    Top =5100
                    Width =1410
                    Height =310
                    ColumnWidth =1410
                    TabIndex =9
                    Name ="RivetID"
                    ControlSource ="RivetID"

                    LayoutCachedLeft =1890
                    LayoutCachedTop =5100
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =5410
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =5100
                            Width =1440
                            Height =310
                            Name ="RivetID_Label"
                            Caption ="RivetID"
                            LayoutCachedLeft =360
                            LayoutCachedTop =5100
                            LayoutCachedWidth =1800
                            LayoutCachedHeight =5410
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1890
                    Top =5520
                    TabIndex =10
                    Name ="Indikator/Lis"
                    ControlSource ="Indikator-Lis"
                    StatusBarText ="wear indikator untuk lis"
                    EventProcPrefix ="Indikator_Lis"

                    LayoutCachedLeft =1890
                    LayoutCachedTop =5520
                    LayoutCachedWidth =2150
                    LayoutCachedHeight =5760
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =5520
                            Width =1440
                            Height =310
                            Name ="Indikator/Lis_Label"
                            Caption ="Indikator-Lis"
                            EventProcPrefix ="Indikator_Lis_Label"
                            LayoutCachedLeft =360
                            LayoutCachedTop =5520
                            LayoutCachedWidth =1800
                            LayoutCachedHeight =5830
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1890
                    Top =5940
                    TabIndex =11
                    Name ="Coak"
                    ControlSource ="Coak"
                    StatusBarText ="Pencos (setengah lingkaran untuk wear indikator)"

                    LayoutCachedLeft =1890
                    LayoutCachedTop =5940
                    LayoutCachedWidth =2150
                    LayoutCachedHeight =6180
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =5940
                            Width =1440
                            Height =310
                            Name ="Coak_Label"
                            Caption ="Coak"
                            LayoutCachedLeft =360
                            LayoutCachedTop =5940
                            LayoutCachedWidth =1800
                            LayoutCachedHeight =6250
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1890
                    Top =6360
                    TabIndex =12
                    Name ="Koak"
                    ControlSource ="Koak"
                    StatusBarText ="koak seperti k-492/k-492 koak"

                    LayoutCachedLeft =1890
                    LayoutCachedTop =6360
                    LayoutCachedWidth =2150
                    LayoutCachedHeight =6600
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =6360
                            Width =1440
                            Height =310
                            Name ="Koak_Label"
                            Caption ="Koak"
                            LayoutCachedLeft =360
                            LayoutCachedTop =6360
                            LayoutCachedWidth =1800
                            LayoutCachedHeight =6670
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1890
                    Top =6780
                    Width =4170
                    Height =1120
                    ColumnWidth =3000
                    TabIndex =13
                    Name ="Catatan"
                    ControlSource ="Catatan"

                    LayoutCachedLeft =1890
                    LayoutCachedTop =6780
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =7900
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =6780
                            Width =1440
                            Height =310
                            Name ="Catatan_Label"
                            Caption ="Catatan"
                            LayoutCachedLeft =360
                            LayoutCachedTop =6780
                            LayoutCachedWidth =1800
                            LayoutCachedHeight =7090
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1890
                    Top =7980
                    Width =3340
                    Height =310
                    ColumnWidth =3000
                    TabIndex =14
                    Name ="Lebar_deprecated"
                    ControlSource ="Lebar_deprecated"
                    StatusBarText ="(mm)"

                    LayoutCachedLeft =1890
                    LayoutCachedTop =7980
                    LayoutCachedWidth =5230
                    LayoutCachedHeight =8290
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =7980
                            Width =1440
                            Height =310
                            Name ="Lebar_deprecated_Label"
                            Caption ="Lebar_deprecated"
                            LayoutCachedLeft =360
                            LayoutCachedTop =7980
                            LayoutCachedWidth =1800
                            LayoutCachedHeight =8290
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1890
                    Top =8400
                    Width =3340
                    Height =310
                    ColumnWidth =3000
                    TabIndex =15
                    Name ="Panjang_deprecated"
                    ControlSource ="Panjang_deprecated"
                    StatusBarText ="(mm)"

                    LayoutCachedLeft =1890
                    LayoutCachedTop =8400
                    LayoutCachedWidth =5230
                    LayoutCachedHeight =8710
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =8400
                            Width =1440
                            Height =310
                            Name ="Panjang_deprecated_Label"
                            Caption ="Panjang_deprecated"
                            LayoutCachedLeft =360
                            LayoutCachedTop =8400
                            LayoutCachedWidth =1800
                            LayoutCachedHeight =8710
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7770
                    Top =360
                    Width =3690
                    Height =580
                    ColumnWidth =3000
                    TabIndex =16
                    Name ="Tebal_deprecated"
                    ControlSource ="Tebal_deprecated"
                    StatusBarText ="(mm)"

                    LayoutCachedLeft =7770
                    LayoutCachedTop =360
                    LayoutCachedWidth =11460
                    LayoutCachedHeight =940
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6240
                            Top =360
                            Width =1440
                            Height =310
                            Name ="Tebal_deprecated_Label"
                            Caption ="Tebal_deprecated"
                            LayoutCachedLeft =6240
                            LayoutCachedTop =360
                            LayoutCachedWidth =7680
                            LayoutCachedHeight =670
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7770
                    Top =1020
                    Width =1410
                    Height =310
                    ColumnWidth =1410
                    TabIndex =17
                    Name ="DiameterDrum_deprecated"
                    ControlSource ="DiameterDrum_deprecated"
                    StatusBarText ="(mm)"

                    LayoutCachedLeft =7770
                    LayoutCachedTop =1020
                    LayoutCachedWidth =9180
                    LayoutCachedHeight =1330
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6240
                            Top =1020
                            Width =1440
                            Height =310
                            Name ="DiameterDrum_deprecated_Label"
                            Caption ="DiameterDrum_deprecated"
                            LayoutCachedLeft =6240
                            LayoutCachedTop =1020
                            LayoutCachedWidth =7680
                            LayoutCachedHeight =1330
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =2880
                    Left =7770
                    Top =1440
                    Width =3340
                    Height =310
                    ColumnWidth =3000
                    TabIndex =18
                    Name ="PolaLubangID_deprecated2"
                    ControlSource ="PolaLubangID_deprecated2"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [T_UkuranPolaLubang].[PolaLubangID], [T_UkuranPolaLubang].[PolaLubang], ["
                        "T_UkuranPolaLubang].[Catatan] FROM T_UkuranPolaLubang ORDER BY [PolaLubang], [Po"
                        "laLubangID], [Catatan]; "
                    ColumnWidths ="0;1440;1440"

                    LayoutCachedLeft =7770
                    LayoutCachedTop =1440
                    LayoutCachedWidth =11110
                    LayoutCachedHeight =1750
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6240
                            Top =1440
                            Width =1440
                            Height =310
                            Name ="PolaLubangID_deprecated2_Label"
                            Caption ="PolaLubangID_deprecated2"
                            LayoutCachedLeft =6240
                            LayoutCachedTop =1440
                            LayoutCachedWidth =7680
                            LayoutCachedHeight =1750
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =2550
                    Left =7770
                    Top =1860
                    Width =3340
                    Height =310
                    ColumnWidth =3000
                    TabIndex =19
                    Name ="PolaLubangID_deprecated"
                    ControlSource ="PolaLubangID_deprecated"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [T_UkuranPolaLubang].PolaLubangID, [T_UkuranPolaLubang].PolaLubang, [T_Uk"
                        "uranPolaLubang].Catatan FROM T_UkuranPolaLubang ORDER BY [T_UkuranPolaLubang].Po"
                        "laLubang; "
                    ColumnWidths ="0;650;1900"
                    AllowValueListEdits =0

                    LayoutCachedLeft =7770
                    LayoutCachedTop =1860
                    LayoutCachedWidth =11110
                    LayoutCachedHeight =2170
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6240
                            Top =1860
                            Width =1440
                            Height =310
                            Name ="PolaLubangID_deprecated_Label"
                            Caption ="PolaLubangID_deprecated"
                            LayoutCachedLeft =6240
                            LayoutCachedTop =1860
                            LayoutCachedWidth =7680
                            LayoutCachedHeight =2170
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7770
                    Top =2280
                    Width =1410
                    Height =310
                    ColumnWidth =1410
                    TabIndex =20
                    Name ="JumlahLubang_deprecated"
                    ControlSource ="JumlahLubang_deprecated"
                    StatusBarText ="Pindah ke dalam tabel PolaLubang"

                    LayoutCachedLeft =7770
                    LayoutCachedTop =2280
                    LayoutCachedWidth =9180
                    LayoutCachedHeight =2590
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6240
                            Top =2280
                            Width =1440
                            Height =310
                            Name ="JumlahLubang_deprecated_Label"
                            Caption ="JumlahLubang_deprecated"
                            LayoutCachedLeft =6240
                            LayoutCachedTop =2280
                            LayoutCachedWidth =7680
                            LayoutCachedHeight =2590
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7770
                    Top =2700
                    Width =3690
                    Height =580
                    ColumnWidth =3000
                    TabIndex =21
                    Name ="UkuranRivetID_deprecated"
                    ControlSource ="UkuranRivetID_deprecated"

                    LayoutCachedLeft =7770
                    LayoutCachedTop =2700
                    LayoutCachedWidth =11460
                    LayoutCachedHeight =3280
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6240
                            Top =2700
                            Width =1440
                            Height =310
                            Name ="UkuranRivetID_deprecated_Label"
                            Caption ="UkuranRivetID_deprecated"
                            LayoutCachedLeft =6240
                            LayoutCachedTop =2700
                            LayoutCachedWidth =7680
                            LayoutCachedHeight =3010
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =7770
                    Top =3360
                    TabIndex =22
                    Name ="TebalTipisFlag_deprecated"
                    ControlSource ="TebalTipisFlag_deprecated"
                    StatusBarText ="Apabila Kepingan ini ada Tebal Tipis"

                    LayoutCachedLeft =7770
                    LayoutCachedTop =3360
                    LayoutCachedWidth =8030
                    LayoutCachedHeight =3600
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6240
                            Top =3360
                            Width =1440
                            Height =310
                            Name ="TebalTipisFlag_deprecated_Label"
                            Caption ="TebalTipisFlag_deprecated"
                            LayoutCachedLeft =6240
                            LayoutCachedTop =3360
                            LayoutCachedWidth =7680
                            LayoutCachedHeight =3670
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7770
                    Top =3780
                    Width =3690
                    Height =580
                    ColumnWidth =3000
                    TabIndex =23
                    Name ="Tebal/Tipis_deprecated"
                    ControlSource ="Tebal/Tipis_deprecated"
                    StatusBarText ="(mm)"
                    EventProcPrefix ="Tebal_Tipis_deprecated"

                    LayoutCachedLeft =7770
                    LayoutCachedTop =3780
                    LayoutCachedWidth =11460
                    LayoutCachedHeight =4360
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6240
                            Top =3780
                            Width =1440
                            Height =310
                            Name ="Tebal/Tipis_deprecated_Label"
                            Caption ="Tebal/Tipis_deprecated"
                            EventProcPrefix ="Tebal_Tipis_deprecated_Label"
                            LayoutCachedLeft =6240
                            LayoutCachedTop =3780
                            LayoutCachedWidth =7680
                            LayoutCachedHeight =4090
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
