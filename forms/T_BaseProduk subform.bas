Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9390
    DatasheetFontHeight =11
    ItemSuffix =17
    Left =90
    Top =130
    Right =18910
    Bottom =11130
    OrderBy ="[Lookup_StatusBaseProduk].[StatusProduk]"
    RecSrcDt = Begin
        0xdb465134941fe640
    End
    RecordSource ="T_BaseProduk"
    Caption ="T_BaseProduk subform"
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
            Height =4540
            Name ="Detail"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2690
                    Top =360
                    Width =3340
                    Height =310
                    ColumnWidth =3000
                    Name ="SetAutoID"
                    ControlSource ="SetAutoID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [T_PcsSetBOM].[SetAutoID], [T_PcsSetBOM].[SetID] FROM T_PcsSetBOM; "
                    ColumnWidths ="0;1441"
                    StatusBarText ="Kode Base Produk (lookup PcsSetBom Table)"
                    AllowValueListEdits =0

                    LayoutCachedLeft =2690
                    LayoutCachedTop =360
                    LayoutCachedWidth =6030
                    LayoutCachedHeight =670
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =360
                            Width =2240
                            Height =310
                            Name ="SetAutoID_Label"
                            Caption ="SetAutoID"
                            LayoutCachedLeft =360
                            LayoutCachedTop =360
                            LayoutCachedWidth =2600
                            LayoutCachedHeight =670
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2690
                    Top =780
                    Width =3340
                    Height =310
                    ColumnWidth =3000
                    TabIndex =1
                    Name ="Deskripsi"
                    ControlSource ="Deskripsi"

                    LayoutCachedLeft =2690
                    LayoutCachedTop =780
                    LayoutCachedWidth =6030
                    LayoutCachedHeight =1090
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =780
                            Width =2240
                            Height =310
                            Name ="Deskripsi_Label"
                            Caption ="Deskripsi"
                            LayoutCachedLeft =360
                            LayoutCachedTop =780
                            LayoutCachedWidth =2600
                            LayoutCachedHeight =1090
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =2690
                    Top =1200
                    ColumnWidth =2060
                    TabIndex =2
                    Name ="CampurKepingFlag_deprecated"
                    ControlSource ="CampurKepingFlag_deprecated"
                    StatusBarText ="info ini bisa diambil dari qcalPcsPerSetCampurKeping"

                    LayoutCachedLeft =2690
                    LayoutCachedTop =1200
                    LayoutCachedWidth =2950
                    LayoutCachedHeight =1440
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =1200
                            Width =2240
                            Height =310
                            Name ="CampurKepingFlag_deprecated_Label"
                            Caption ="CampurKepingFlag_deprecated"
                            LayoutCachedLeft =360
                            LayoutCachedTop =1200
                            LayoutCachedWidth =2600
                            LayoutCachedHeight =1510
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2690
                    Top =1620
                    Width =970
                    Height =310
                    ColumnWidth =970
                    TabIndex =3
                    Name ="pcs/set_deprecated"
                    ControlSource ="pcs/set_deprecated"
                    StatusBarText ="info ini bisa diambil dari qcalPcsPerSetCampurKeping"
                    EventProcPrefix ="pcs_set_deprecated"

                    LayoutCachedLeft =2690
                    LayoutCachedTop =1620
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =1930
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =1620
                            Width =2240
                            Height =310
                            Name ="pcs/set_deprecated_Label"
                            Caption ="pcs/set_deprecated"
                            EventProcPrefix ="pcs_set_deprecated_Label"
                            LayoutCachedLeft =360
                            LayoutCachedTop =1620
                            LayoutCachedWidth =2600
                            LayoutCachedHeight =1930
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2690
                    Top =2040
                    Width =3340
                    Height =310
                    ColumnWidth =3000
                    TabIndex =4
                    Name ="AxlePosID"
                    ControlSource ="AxlePosID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [T_ListAxlePos].[AxlePosID] AS Expr1, [T_ListAxlePos].[AxlePos] AS Expr2 "
                        "FROM T_ListAxlePos; "
                    ColumnWidths ="0;1441"
                    StatusBarText ="(Lookup) Jenis Produk: F/R,Handbrake, Trailer, Lempengan dari T_AxlePos"
                    AllowValueListEdits =0

                    LayoutCachedLeft =2690
                    LayoutCachedTop =2040
                    LayoutCachedWidth =6030
                    LayoutCachedHeight =2350
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =2040
                            Width =2240
                            Height =310
                            Name ="AxlePosID_Label"
                            Caption ="AxlePosID"
                            LayoutCachedLeft =360
                            LayoutCachedTop =2040
                            LayoutCachedWidth =2600
                            LayoutCachedHeight =2350
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2690
                    Top =2460
                    Width =3340
                    Height =310
                    ColumnWidth =3000
                    TabIndex =5
                    Name ="KategoriProdukID"
                    ControlSource ="KategoriProdukID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [T_KategoriProduk].[KategoriProdukID], [T_KategoriProduk].[KategoriProduk"
                        "] FROM T_KategoriProduk; "
                    ColumnWidths ="0;1441"
                    StatusBarText ="(Lookup) dari T_KategoriProduk"
                    AllowValueListEdits =0

                    LayoutCachedLeft =2690
                    LayoutCachedTop =2460
                    LayoutCachedWidth =6030
                    LayoutCachedHeight =2770
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =2460
                            Width =2240
                            Height =310
                            Name ="KategoriProdukID_Label"
                            Caption ="KategoriProdukID"
                            LayoutCachedLeft =360
                            LayoutCachedTop =2460
                            LayoutCachedWidth =2600
                            LayoutCachedHeight =2770
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =2690
                    Top =2880
                    Width =3340
                    Height =310
                    ColumnWidth =3000
                    TabIndex =6
                    Name ="StatusBaseProduk"
                    ControlSource ="StatusBaseProduk"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT T_ListStatusBaseProduk.[StatusProdukID], [T_ListStatusBaseProduk].[Status"
                        "Produk] FROM T_ListStatusBaseProduk ORDER BY [StatusProdukID]; "
                    ColumnWidths ="0;1440"
                    StatusBarText ="(Lookup) List Status dari Sample, Prototype, Produksi"
                    AllowValueListEdits =0

                    LayoutCachedLeft =2690
                    LayoutCachedTop =2880
                    LayoutCachedWidth =6030
                    LayoutCachedHeight =3190
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =2880
                            Width =2240
                            Height =310
                            Name ="StatusBaseProduk_Label"
                            Caption ="StatusBaseProduk"
                            LayoutCachedLeft =360
                            LayoutCachedTop =2880
                            LayoutCachedWidth =2600
                            LayoutCachedHeight =3190
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2690
                    Top =3300
                    Width =6640
                    Height =1120
                    ColumnWidth =3000
                    TabIndex =7
                    Name ="Catatan"
                    ControlSource ="Catatan"
                    StatusBarText ="Ini untuk memuat Catatan Panjang tentang Produk terkait"

                    LayoutCachedLeft =2690
                    LayoutCachedTop =3300
                    LayoutCachedWidth =9330
                    LayoutCachedHeight =4420
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =3300
                            Width =2240
                            Height =310
                            Name ="Catatan_Label"
                            Caption ="Catatan"
                            LayoutCachedLeft =360
                            LayoutCachedTop =3300
                            LayoutCachedWidth =2600
                            LayoutCachedHeight =3610
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    ColumnCount =2
                    Left =2700
                    Height =300
                    TabIndex =8
                    Name ="SetID"
                    ControlSource ="BaseProdukName"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [T_PcsSetBOM].[SetAutoID], [T_PcsSetBOM].[SetID] FROM T_PcsSetBOM; "
                    ColumnWidths ="0;1441"
                    StatusBarText ="Copy dari SetAutoID"
                    AllowValueListEdits =0

                    LayoutCachedLeft =2700
                    LayoutCachedWidth =4140
                    LayoutCachedHeight =300
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1200
                            Width =530
                            Height =300
                            Name ="Label16"
                            Caption ="BaseProdukName"
                            LayoutCachedLeft =1200
                            LayoutCachedWidth =1730
                            LayoutCachedHeight =300
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
