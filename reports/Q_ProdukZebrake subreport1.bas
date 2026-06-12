Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11715
    DatasheetFontHeight =11
    ItemSuffix =26
    Left =360
    Top =2460
    RecSrcDt = Begin
        0x8e1c4bf4ddeae540
    End
    RecordSource ="Q_ProdukZebrake"
    Caption ="Q_ProdukZebrake subreport1"
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =555
            Name ="ReportHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin EmptyCell
                    Left =300
                    Top =60
                    Width =690
                    Height =450
                    Name ="Auto_LogoEmptyCell"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =300
                    LayoutCachedTop =60
                    LayoutCachedWidth =990
                    LayoutCachedHeight =510
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    Left =1050
                    Top =60
                    Width =9121
                    Height =450
                    FontSize =18
                    Name ="Auto_Header0"
                    Caption ="Catalog Produk Zebrake"
                    FontName ="Calibri Light"
                    GroupTable =1
                    GridlineColor =10921638
                    HorizontalAnchor =2
                    LayoutCachedLeft =1050
                    LayoutCachedTop =60
                    LayoutCachedWidth =10171
                    LayoutCachedHeight =510
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ThemeFontIndex =0
                    BorderThemeColorIndex =2
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =1
                End
            End
        End
        Begin PageHeader
            Height =360
            Name ="PageHeaderSection"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin Section
            KeepTogether = NotDefault
            Height =495
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =2040
                    Height =330
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="KodeBarang"
                    ControlSource ="KodeBarang"
                    GridlineColor =10921638

                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =2100
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =2160
                    Top =60
                    Width =2370
                    Height =330
                    ColumnWidth =5595
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Kendaraan"
                    ControlSource ="Kendaraan"
                    GridlineColor =10921638

                    LayoutCachedLeft =2160
                    LayoutCachedTop =60
                    LayoutCachedWidth =4530
                    LayoutCachedHeight =390
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =4590
                    Top =60
                    Width =690
                    Height =330
                    ColumnWidth =600
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =3484194
                    Name ="AxlePosID"
                    ControlSource ="AxlePosID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [T_AxlePos].[AxlePosID], [T_AxlePos].[AxlePos] FROM T_AxlePos; "
                    ColumnWidths ="0;1440"
                    StatusBarText ="Jenis Produk: F/R,Handbrake, Trailer, Lempengan"
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =4590
                    LayoutCachedTop =60
                    LayoutCachedWidth =5280
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =5340
                    Top =60
                    Width =885
                    Height =330
                    ColumnWidth =1140
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="pcs/set"
                    ControlSource ="pcs/set"
                    EventProcPrefix ="pcs_set"
                    GridlineColor =10921638

                    LayoutCachedLeft =5340
                    LayoutCachedTop =60
                    LayoutCachedWidth =6225
                    LayoutCachedHeight =390
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =6345
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

                    LayoutCachedLeft =6345
                    LayoutCachedTop =60
                    LayoutCachedWidth =7365
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =7425
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

                    LayoutCachedLeft =7425
                    LayoutCachedTop =60
                    LayoutCachedWidth =9465
                    LayoutCachedHeight =390
                End
                Begin CheckBox
                    Left =9525
                    Top =60
                    Width =485
                    Height =300
                    TabIndex =6
                    BorderColor =10921638
                    Name ="CampurKepingFlag"
                    ControlSource ="CampurKepingFlag"
                    StatusBarText ="Apakah ini menandakan kalo satu set terisi dari dua atau lebih jenis keping yang"
                        " berbeda"
                    GridlineColor =10921638

                    LayoutCachedLeft =9525
                    LayoutCachedTop =60
                    LayoutCachedWidth =10010
                    LayoutCachedHeight =360
                End
            End
        End
        Begin PageFooter
            Height =405
            Name ="PageFooterSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =15
                    Top =15
                    Width =3600
                    Height =315
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text22"
                    ControlSource ="=\"Halaman: \" & [Page] & \" of \" & [Pages]"
                    GridlineColor =10921638

                    LayoutCachedLeft =15
                    LayoutCachedTop =15
                    LayoutCachedWidth =3615
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9240
                    Top =30
                    Width =885
                    Height =315
                    TabIndex =1
                    BorderColor =10921638
                    Name ="Auto_Date"
                    ControlSource ="=Date()"
                    Format ="Medium Date"
                    GroupTable =2
                    GridlineColor =10921638
                    HorizontalAnchor =2

                    LayoutCachedLeft =9240
                    LayoutCachedTop =30
                    LayoutCachedWidth =10125
                    LayoutCachedHeight =345
                    ColumnEnd =1
                    LayoutGroup =2
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =2
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =15
            Name ="ReportFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
