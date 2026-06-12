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
    Width =10785
    DatasheetFontHeight =11
    ItemSuffix =32
    Left =360
    Top =2460
    RecSrcDt = Begin
        0x862c5b92fdeae540
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
            Height =907
            Name ="ReportHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin EmptyCell
                    Left =135
                    Top =225
                    Width =0
                    Height =450
                    Name ="Auto_LogoEmptyCell"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =135
                    LayoutCachedTop =225
                    LayoutCachedWidth =135
                    LayoutCachedHeight =675
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    TextAlign =2
                    Left =195
                    Top =225
                    Width =9915
                    Height =450
                    FontSize =18
                    FontWeight =700
                    Name ="Auto_Header0"
                    Caption ="Catalog Produk Zebrake"
                    FontName ="Calibri Light"
                    GroupTable =1
                    GridlineColor =10921638
                    HorizontalAnchor =2
                    LayoutCachedLeft =195
                    LayoutCachedTop =225
                    LayoutCachedWidth =10110
                    LayoutCachedHeight =675
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
            Height =405
            Name ="PageHeaderSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    Left =60
                    Top =30
                    Width =1230
                    Height =315
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Label27"
                    Caption ="Kode Barang"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =30
                    LayoutCachedWidth =1290
                    LayoutCachedHeight =345
                End
                Begin Label
                    Left =2190
                    Top =30
                    Width =1575
                    Height =315
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Label28"
                    Caption ="Jenis Kendaraan"
                    GridlineColor =10921638
                    LayoutCachedLeft =2190
                    LayoutCachedTop =30
                    LayoutCachedWidth =3765
                    LayoutCachedHeight =345
                End
                Begin Label
                    Left =4590
                    Top =30
                    Width =585
                    Height =315
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Label29"
                    Caption ="F/R/T"
                    GridlineColor =10921638
                    LayoutCachedLeft =4590
                    LayoutCachedTop =30
                    LayoutCachedWidth =5175
                    LayoutCachedHeight =345
                End
                Begin Label
                    Left =5355
                    Top =30
                    Width =765
                    Height =315
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Label30"
                    Caption ="Pcs/Set"
                    GridlineColor =10921638
                    LayoutCachedLeft =5355
                    LayoutCachedTop =30
                    LayoutCachedWidth =6120
                    LayoutCachedHeight =345
                End
                Begin Label
                    Left =8475
                    Top =30
                    Width =1350
                    Height =315
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Label31"
                    Caption ="Mixed Keping"
                    GridlineColor =10921638
                    LayoutCachedLeft =8475
                    LayoutCachedTop =30
                    LayoutCachedWidth =9825
                    LayoutCachedHeight =345
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
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
                    CanGrow = NotDefault
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
                    ControlSource ="Deskripsi_deprecated"
                    GridlineColor =10921638

                    LayoutCachedLeft =2160
                    LayoutCachedTop =60
                    LayoutCachedWidth =4530
                    LayoutCachedHeight =390
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
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
                    TextAlign =2
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
                Begin TextBox
                    CanGrow = NotDefault
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =6480
                    Top =60
                    Width =2985
                    Height =330
                    ColumnWidth =2070
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="KodeBarangCust"
                    ControlSource ="KodeBarangCust"
                    GridlineColor =10921638

                    LayoutCachedLeft =6480
                    LayoutCachedTop =60
                    LayoutCachedWidth =9465
                    LayoutCachedHeight =390
                End
                Begin CheckBox
                    Left =9525
                    Top =60
                    Width =485
                    Height =300
                    TabIndex =5
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
            Height =435
            Name ="PageFooterSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =45
                    Top =75
                    Width =3600
                    Height =315
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text22"
                    ControlSource ="=\"Halaman: \" & [Page] & \" of \" & [Pages]"
                    GridlineColor =10921638

                    LayoutCachedLeft =45
                    LayoutCachedTop =75
                    LayoutCachedWidth =3645
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6420
                    Top =90
                    Width =3675
                    Height =315
                    TabIndex =1
                    BorderColor =10921638
                    Name ="Auto_Date"
                    ControlSource ="=Date()"
                    Format ="Medium Date"
                    GroupTable =2
                    GridlineColor =10921638
                    HorizontalAnchor =2

                    LayoutCachedLeft =6420
                    LayoutCachedTop =90
                    LayoutCachedWidth =10095
                    LayoutCachedHeight =405
                    ColumnEnd =1
                    LayoutGroup =2
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin Label
                    TextAlign =2
                    Left =2895
                    Top =75
                    Width =4290
                    Height =345
                    FontWeight =700
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Label26"
                    Caption ="PT Galangsa Spartindo Perkasa"
                    GridlineColor =10921638
                    LayoutCachedLeft =2895
                    LayoutCachedTop =75
                    LayoutCachedWidth =7185
                    LayoutCachedHeight =420
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
