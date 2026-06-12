Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11470
    DatasheetFontHeight =11
    ItemSuffix =14
    RecSrcDt = Begin
        0xe26f53d4ddeae540
    End
    RecordSource ="Q_ProdukZebrake"
    Caption ="Q_ProdukZebrake subreport"
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    FitToPage =1
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
            ShowDatePicker =0
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
            KeepTogether = NotDefault
            Height =435
            Name ="ReportHeader"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    TextAlign =1
                    Left =360
                    Top =60
                    Width =2040
                    Height =315
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="KodeBarang_Label"
                    Caption ="KodeBarang"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =360
                    LayoutCachedTop =60
                    LayoutCachedWidth =2400
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =1
                    Left =2460
                    Top =60
                    Width =2040
                    Height =315
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Kendaraan_Label"
                    Caption ="Kendaraan"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =2460
                    LayoutCachedTop =60
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =1
                    Left =4560
                    Top =60
                    Width =1020
                    Height =315
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="AxlePosID_Label"
                    Caption ="AxlePosID"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =4560
                    LayoutCachedTop =60
                    LayoutCachedWidth =5580
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =1
                    Left =5640
                    Top =60
                    Width =2040
                    Height =315
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="pcs/set_Label"
                    Caption ="pcs/set"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="pcs_set_Label"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =5640
                    LayoutCachedTop =60
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =1
                    Left =7740
                    Top =60
                    Width =1020
                    Height =315
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="KategoriProdukID_Label"
                    Caption ="KategoriProdukID"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =7740
                    LayoutCachedTop =60
                    LayoutCachedWidth =8760
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =1
                    Left =8820
                    Top =60
                    Width =2040
                    Height =315
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="KodeBarangCust_Label"
                    Caption ="KodeBarangCust"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =8820
                    LayoutCachedTop =60
                    LayoutCachedWidth =10860
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =1
                    Left =10920
                    Top =60
                    Width =490
                    Height =315
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="CampurKepingFlag_Label"
                    Caption ="CampurKepingFlag"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =10920
                    LayoutCachedTop =60
                    LayoutCachedWidth =11410
                    LayoutCachedHeight =375
                End
            End
        End
        Begin PageHeader
            Height =0
            Name ="PageHeaderSection"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin Section
            KeepTogether = NotDefault
            Height =450
            Name ="Detail"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =360
                    Top =60
                    Width =2040
                    Height =330
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="KodeBarang"
                    ControlSource ="KodeBarang"
                    GridlineColor =10921638

                    LayoutCachedLeft =360
                    LayoutCachedTop =60
                    LayoutCachedWidth =2400
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =2460
                    Top =60
                    Width =2040
                    Height =330
                    ColumnWidth =5595
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Kendaraan"
                    ControlSource ="Deskripsi_deprecated"
                    GridlineColor =10921638

                    LayoutCachedLeft =2460
                    LayoutCachedTop =60
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =390
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =4560
                    Top =60
                    Width =1020
                    Height =330
                    ColumnWidth =600
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =3484194
                    Name ="AxlePosID"
                    ControlSource ="AxlePosID_deprecated"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [T_AxlePos].[AxlePosID], [T_AxlePos].[AxlePos] FROM T_AxlePos; "
                    ColumnWidths ="0;1440"
                    StatusBarText ="Jenis Produk: F/R,Handbrake, Trailer, Lempengan"
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =4560
                    LayoutCachedTop =60
                    LayoutCachedWidth =5580
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =5640
                    Top =60
                    Width =2040
                    Height =330
                    ColumnWidth =1140
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="pcs/set"
                    ControlSource ="pcs/set"
                    EventProcPrefix ="pcs_set"
                    GridlineColor =10921638

                    LayoutCachedLeft =5640
                    LayoutCachedTop =60
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =390
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =7740
                    Top =60
                    Width =1020
                    Height =330
                    ColumnWidth =2400
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =3484194
                    Name ="KategoriProdukID"
                    ControlSource ="KategoriProdukID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [T_KategoriProduk].[KategoriProdukID], [T_KategoriProduk].[KategoriProduk"
                        "] FROM T_KategoriProduk; "
                    ColumnWidths ="0;1440"
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =7740
                    LayoutCachedTop =60
                    LayoutCachedWidth =8760
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =8820
                    Top =60
                    Width =2040
                    Height =330
                    ColumnWidth =2070
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="KodeBarangCust"
                    ControlSource ="KodeBarangCust"
                    GridlineColor =10921638

                    LayoutCachedLeft =8820
                    LayoutCachedTop =60
                    LayoutCachedWidth =10860
                    LayoutCachedHeight =390
                End
                Begin CheckBox
                    Left =10920
                    Top =60
                    TabIndex =6
                    BorderColor =10921638
                    Name ="CampurKepingFlag"
                    ControlSource ="CampurKepingFlag_deprecated"
                    StatusBarText ="Apakah ini menandakan kalo satu set terisi dari dua atau lebih jenis keping yang"
                        " berbeda"
                    GridlineColor =10921638

                    LayoutCachedLeft =10920
                    LayoutCachedTop =60
                    LayoutCachedWidth =11180
                    LayoutCachedHeight =300
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
