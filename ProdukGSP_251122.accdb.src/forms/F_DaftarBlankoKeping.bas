Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    KeyPreview = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =12960
    DatasheetFontHeight =11
    ItemSuffix =34
    Left =4840
    Top =780
    Right =21320
    Bottom =11130
    OrderBy ="Tebal DESC"
    RecSrcDt = Begin
        0x6fb00f23db3de640
    End
    RecordSource ="T_BlankoKeping"
    Caption ="T_BlankoKeping"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    OnKeyDown ="[Event Procedure]"
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
        Begin CommandButton
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
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
        Begin FormHeader
            Height =1080
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =60
                    Top =720
                    Width =1200
                    Height =300
                    FontWeight =700
                    Name ="BlankoKepingID_Label"
                    Caption ="ID"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =60
                    LayoutCachedTop =720
                    LayoutCachedWidth =1260
                    LayoutCachedHeight =1020
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =1320
                    Top =720
                    Width =3720
                    Height =300
                    FontWeight =700
                    Name ="DeskripsiBlankoKeping_Label"
                    Caption ="Deskripsi Keping"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =1320
                    LayoutCachedTop =720
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =1020
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =5100
                    Top =720
                    Width =660
                    Height =300
                    FontWeight =700
                    Name ="lblPanjang"
                    Caption ="Pjg"
                    OnClick ="[Event Procedure]"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =5100
                    LayoutCachedTop =720
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =1020
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =5820
                    Top =720
                    Width =660
                    Height =300
                    FontWeight =700
                    Name ="lblLebar"
                    Caption ="Lebar"
                    OnClick ="[Event Procedure]"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =5820
                    LayoutCachedTop =720
                    LayoutCachedWidth =6480
                    LayoutCachedHeight =1020
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =7260
                    Top =720
                    Width =360
                    Height =300
                    FontWeight =700
                    Name ="lblTebalTipisFlag"
                    Caption ="R/R"
                    OnClick ="[Event Procedure]"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =7260
                    LayoutCachedTop =720
                    LayoutCachedWidth =7620
                    LayoutCachedHeight =1020
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =6540
                    Top =720
                    Width =660
                    Height =300
                    FontWeight =700
                    Name ="lblTebal"
                    Caption ="Tebal"
                    OnClick ="[Event Procedure]"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =6540
                    LayoutCachedTop =720
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =1020
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =7680
                    Top =720
                    Width =660
                    Height =300
                    FontWeight =700
                    Name ="lblTebal2"
                    Caption ="Tbl 2"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =7680
                    LayoutCachedTop =720
                    LayoutCachedWidth =8340
                    LayoutCachedHeight =1020
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =8400
                    Top =720
                    Width =660
                    Height =300
                    FontWeight =700
                    Name ="lblDiameterDrum"
                    Caption ="Ø Drm"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =8400
                    LayoutCachedTop =720
                    LayoutCachedWidth =9060
                    LayoutCachedHeight =1020
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =9120
                    Top =720
                    Width =660
                    Height =300
                    FontWeight =700
                    Name ="lblDiaBrakeShoe"
                    Caption ="Ø Shoe"
                    OnClick ="[Event Procedure]"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =9120
                    LayoutCachedTop =720
                    LayoutCachedWidth =9780
                    LayoutCachedHeight =1020
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =10560
                    Top =720
                    Width =2280
                    Height =300
                    FontWeight =700
                    Name ="lblCatatanBlankoKeping"
                    Caption ="Catatan"
                    OnClick ="[Event Procedure]"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =10560
                    LayoutCachedTop =720
                    LayoutCachedWidth =12840
                    LayoutCachedHeight =1020
                End
                Begin Label
                    OverlapFlags =85
                    Left =60
                    Top =60
                    Width =3722
                    Height =540
                    FontSize =20
                    Name ="Label20"
                    Caption ="Daftar Blanko Keping"
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =3782
                    LayoutCachedHeight =600
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7200
                    Top =180
                    Width =2040
                    Height =420
                    Name ="cmdEditBlankoBtn"
                    Caption ="Detail Keping Blanko"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =7200
                    LayoutCachedTop =180
                    LayoutCachedWidth =9240
                    LayoutCachedHeight =600
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4680
                    Top =180
                    Width =2340
                    Height =420
                    TabIndex =1
                    Name ="cmdTambahKepingBlankoBtn"
                    Caption ="Tambah Keping Blanko"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =4680
                    LayoutCachedTop =180
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =600
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9420
                    Top =180
                    Width =2340
                    Height =420
                    TabIndex =2
                    Name ="cmdExportExcelBtn"
                    Caption ="Export ke Excel"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =9420
                    LayoutCachedTop =180
                    LayoutCachedWidth =11760
                    LayoutCachedHeight =600
                End
                Begin Label
                    OverlapFlags =85
                    Left =9840
                    Top =720
                    Width =660
                    Height =300
                    FontWeight =700
                    Name ="lblPjgB"
                    Caption ="PjgB"
                    LayoutCachedLeft =9840
                    LayoutCachedTop =720
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =1020
                End
            End
        End
        Begin Section
            Height =420
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =1200
                    Height =300
                    ColumnWidth =1440
                    Name ="BlankoKepingID"
                    ControlSource ="BlankoKepingID"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =1260
                    LayoutCachedHeight =360
                    DisplayAsHyperlink =2
                    ForeThemeColorIndex =10
                    ForeTint =100.0
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1320
                    Top =60
                    Width =3720
                    Height =300
                    ColumnWidth =3000
                    TabIndex =1
                    Name ="DeskripsiBlankoKeping"
                    ControlSource ="DeskripsiBlankoKeping"
                    StatusBarText ="Deskripsi Dasar Blanko (Sebelum lobang/finishing).  Lebih untuk mengetahui cetak"
                        " mould"

                    LayoutCachedLeft =1320
                    LayoutCachedTop =60
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5100
                    Top =60
                    Width =660
                    Height =310
                    ColumnWidth =1470
                    TabIndex =2
                    Name ="Panjang"
                    ControlSource ="Panjang"
                    Format ="General Number"
                    StatusBarText ="(mm)"

                    LayoutCachedLeft =5100
                    LayoutCachedTop =60
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =370
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5820
                    Top =60
                    Width =660
                    Height =310
                    ColumnWidth =1410
                    TabIndex =3
                    Name ="Lebar"
                    ControlSource ="Lebar"
                    StatusBarText ="(mm)"

                    LayoutCachedLeft =5820
                    LayoutCachedTop =60
                    LayoutCachedWidth =6480
                    LayoutCachedHeight =370
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =7320
                    Top =120
                    TabIndex =4
                    Name ="Tebal/TipisFlag"
                    ControlSource ="Tebal_TipisFlag"
                    StatusBarText ="Untuk ketebalan yang tidak sama antara satu ujung dengan yang lain"
                    EventProcPrefix ="Tebal_TipisFlag"

                    LayoutCachedLeft =7320
                    LayoutCachedTop =120
                    LayoutCachedWidth =7580
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6540
                    Top =60
                    Width =660
                    Height =310
                    ColumnWidth =3000
                    TabIndex =5
                    Name ="Tebal"
                    ControlSource ="Tebal"
                    StatusBarText ="(mm). Tebal (untuk keping yang tebal-tipis, ketebalan yang tertinggi- yang lebih"
                        " tepis di Tebal2)"

                    LayoutCachedLeft =6540
                    LayoutCachedTop =60
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =370
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7680
                    Top =60
                    Width =660
                    Height =310
                    ColumnWidth =3000
                    TabIndex =6
                    Name ="Tebal2"
                    ControlSource ="Tebal2"
                    StatusBarText ="(mm) Apabila ada yang bagian tebal tipis.  Ketebalan tipis untuk yang tipis"

                    LayoutCachedLeft =7680
                    LayoutCachedTop =60
                    LayoutCachedWidth =8340
                    LayoutCachedHeight =370
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8400
                    Top =60
                    Width =660
                    Height =310
                    ColumnWidth =1410
                    TabIndex =7
                    Name ="DiameterDrum"
                    ControlSource ="DiameterDrum"
                    Format ="General Number"

                    LayoutCachedLeft =8400
                    LayoutCachedTop =60
                    LayoutCachedWidth =9060
                    LayoutCachedHeight =370
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9120
                    Top =60
                    Width =660
                    Height =310
                    ColumnWidth =1410
                    TabIndex =8
                    Name ="DiameterBrakeShoe"
                    ControlSource ="DiameterBrakeShoe"

                    LayoutCachedLeft =9120
                    LayoutCachedTop =60
                    LayoutCachedWidth =9780
                    LayoutCachedHeight =370
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9840
                    Top =60
                    Width =660
                    Height =300
                    TabIndex =9
                    Name ="Panjang2"
                    ControlSource ="Panjang2"
                    StatusBarText ="(mm) Untuk Panjang Bagian Bawah"

                    LayoutCachedLeft =9840
                    LayoutCachedTop =60
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10560
                    Top =60
                    Width =2280
                    Height =300
                    ColumnWidth =3000
                    TabIndex =10
                    Name ="CatatanBlankoKeping"
                    ControlSource ="CatatanBlankoKeping"

                    LayoutCachedLeft =10560
                    LayoutCachedTop =60
                    LayoutCachedWidth =12840
                    LayoutCachedHeight =360
                End
            End
        End
        Begin FormFooter
            Height =720
            Name ="FormFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =11400
                    Top =360
                    Width =1500
                    Height =300
                    Name ="Label25"
                    Caption ="T_BlankoKeping"
                    LayoutCachedLeft =11400
                    LayoutCachedTop =360
                    LayoutCachedWidth =12900
                    LayoutCachedHeight =660
                End
                Begin Label
                    OverlapFlags =85
                    Left =60
                    Top =60
                    Width =11250
                    Height =600
                    Name ="Label27"
                    Caption ="P = Panjang Atas. D = Diameter Drum. T2= Tebal kedua (untuk tebal tipis), D2 = D"
                        "iameter Shoe. R/R untuk Keping tebal/tipis\015\012Klik di label kolom untuk sort"
                        "ir.  Cari dengan menggunakan tombol huruf dan angka."
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =11310
                    LayoutCachedHeight =660
                End
            End
        End
    End
End
CodeBehindForm
' See "F_DaftarBlankoKeping.cls"
