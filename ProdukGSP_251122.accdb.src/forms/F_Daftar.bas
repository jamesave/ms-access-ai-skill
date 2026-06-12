Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11520
    DatasheetFontHeight =11
    ItemSuffix =17
    Left =1280
    Top =390
    Right =25010
    Bottom =11660
    RecSrcDt = Begin
        0x624399f2fd42e640
    End
    RecordSource ="T_Produk"
    Caption ="F_Daftar"
    DatasheetFontName ="Calibri"
    AllowDatasheetView =0
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
            Height =1380
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =3
                    Left =360
                    Top =720
                    Width =780
                    Height =300
                    Name ="ProdukID_Label"
                    Caption ="ProdukID"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =360
                    LayoutCachedTop =720
                    LayoutCachedWidth =1140
                    LayoutCachedHeight =1020
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =1200
                    Top =720
                    Width =1200
                    Height =300
                    Name ="KodeBarang_Label"
                    Caption ="KodeBarang"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =1200
                    LayoutCachedTop =720
                    LayoutCachedWidth =2400
                    LayoutCachedHeight =1020
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =2460
                    Top =720
                    Width =3420
                    Height =300
                    Name ="KodeBarangCust_Label"
                    Caption ="KodeBarangCust"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =2460
                    LayoutCachedTop =720
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =1020
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =5940
                    Top =720
                    Width =3420
                    Height =300
                    Name ="DeskripsiCust_Label"
                    Caption ="DeskripsiCust"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =5940
                    LayoutCachedTop =720
                    LayoutCachedWidth =9360
                    LayoutCachedHeight =1020
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =9420
                    Top =720
                    Width =780
                    Height =300
                    Name ="CustomerID_Label"
                    Caption ="CustomerID"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =9420
                    LayoutCachedTop =720
                    LayoutCachedWidth =10200
                    LayoutCachedHeight =1020
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =10260
                    Top =720
                    Width =780
                    Height =300
                    Name ="SetAutoID_Label"
                    Caption ="SetAutoID"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =10260
                    LayoutCachedTop =720
                    LayoutCachedWidth =11040
                    LayoutCachedHeight =1020
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =11100
                    Top =720
                    Width =360
                    Height =300
                    Name ="BaseProdukID_Label"
                    Caption ="BaseProdukID"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =11100
                    LayoutCachedTop =720
                    LayoutCachedWidth =11460
                    LayoutCachedHeight =1020
                End
                Begin Label
                    OverlapFlags =215
                    Left =60
                    Top =60
                    Width =3602
                    Height =1020
                    FontSize =20
                    Name ="Label14"
                    Caption ="Daftar Produk"
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =3662
                    LayoutCachedHeight =1080
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =7920
                    Top =180
                    Width =2460
                    Height =300
                    Name ="Combo15"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [T_ListCustomer].[CustomerID], [T_ListCustomer].[NamaCustomer] FROM T_Lis"
                        "tCustomer ORDER BY [CustomerID]; "
                    ColumnWidths ="0;1440"

                    LayoutCachedLeft =7920
                    LayoutCachedTop =180
                    LayoutCachedWidth =10380
                    LayoutCachedHeight =480
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6120
                            Top =180
                            Width =1570
                            Height =320
                            Name ="NamaCustomer_Label"
                            Caption ="NamaCustomer"
                            LayoutCachedLeft =6120
                            LayoutCachedTop =180
                            LayoutCachedWidth =7690
                            LayoutCachedHeight =500
                        End
                    End
                End
            End
        End
        Begin Section
            Height =480
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =360
                    Top =60
                    Width =780
                    Height =300
                    ColumnWidth =1440
                    Name ="ProdukID"
                    ControlSource ="ProdukID"

                    LayoutCachedLeft =360
                    LayoutCachedTop =60
                    LayoutCachedWidth =1140
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1200
                    Top =60
                    Width =1200
                    Height =300
                    ColumnWidth =2240
                    TabIndex =1
                    Name ="KodeBarang"
                    ControlSource ="KodeBarang"

                    LayoutCachedLeft =1200
                    LayoutCachedTop =60
                    LayoutCachedWidth =2400
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2460
                    Top =60
                    Width =3420
                    Height =300
                    ColumnWidth =3000
                    TabIndex =2
                    Name ="KodeBarangCust"
                    ControlSource ="KodeBarangCust"
                    StatusBarText ="Deskripsi Customer"

                    LayoutCachedLeft =2460
                    LayoutCachedTop =60
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5940
                    Top =60
                    Width =3420
                    Height =300
                    ColumnWidth =3000
                    TabIndex =3
                    Name ="DeskripsiCust"
                    ControlSource ="DeskripsiCust"
                    StatusBarText ="Untuk Barang MI, SCB, GSP dsb"

                    LayoutCachedLeft =5940
                    LayoutCachedTop =60
                    LayoutCachedWidth =9360
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9420
                    Top =60
                    Width =780
                    Height =300
                    ColumnWidth =1410
                    TabIndex =4
                    Name ="CustomerID"
                    ControlSource ="CustomerID"

                    LayoutCachedLeft =9420
                    LayoutCachedTop =60
                    LayoutCachedWidth =10200
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10260
                    Top =60
                    Width =780
                    Height =300
                    ColumnWidth =1410
                    TabIndex =5
                    Name ="SetAutoID"
                    ControlSource ="SetAutoID"

                    LayoutCachedLeft =10260
                    LayoutCachedTop =60
                    LayoutCachedWidth =11040
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =11100
                    Top =60
                    Width =360
                    Height =300
                    ColumnWidth =1410
                    TabIndex =6
                    Name ="BaseProdukID"
                    ControlSource ="BaseProdukID"

                    LayoutCachedLeft =11100
                    LayoutCachedTop =60
                    LayoutCachedWidth =11460
                    LayoutCachedHeight =360
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
