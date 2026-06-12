Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9390
    DatasheetFontHeight =11
    ItemSuffix =26
    Left =410
    Top =4300
    Right =9790
    Bottom =8950
    RecSrcDt = Begin
        0x28bc6746ee48e640
    End
    RecordSource ="T_BlankoKeping"
    Caption ="sfrm_BlankoKeping"
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
            Height =0
            Name ="FormHeader"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
        End
        Begin Section
            Height =4200
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2450
                    Top =60
                    Width =6640
                    Height =310
                    ColumnWidth =3000
                    Name ="BlankoKepingID"
                    ControlSource ="BlankoKepingID"

                    LayoutCachedLeft =2450
                    LayoutCachedTop =60
                    LayoutCachedWidth =9090
                    LayoutCachedHeight =370
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =60
                            Width =2240
                            Height =310
                            Name ="BlankoKepingID_Label"
                            Caption ="BlankoKepingID"
                            LayoutCachedLeft =120
                            LayoutCachedTop =60
                            LayoutCachedWidth =2360
                            LayoutCachedHeight =370
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2450
                    Top =420
                    Width =6640
                    Height =300
                    ColumnWidth =3000
                    TabIndex =1
                    Name ="DeskripsiBlankoKeping"
                    ControlSource ="DeskripsiBlankoKeping"
                    StatusBarText ="Deskripsi Dasar Blanko (Sebelum lobang/finishing).  Lebih untuk mengetahui cetak"
                        " mould"

                    LayoutCachedLeft =2450
                    LayoutCachedTop =420
                    LayoutCachedWidth =9090
                    LayoutCachedHeight =720
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =420
                            Width =2240
                            Height =310
                            Name ="DeskripsiBlankoKeping_Label"
                            Caption ="DeskripsiBlankoKeping"
                            LayoutCachedLeft =120
                            LayoutCachedTop =420
                            LayoutCachedWidth =2360
                            LayoutCachedHeight =730
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2450
                    Top =780
                    Width =3340
                    Height =310
                    ColumnWidth =3000
                    TabIndex =2
                    Name ="JenisBlankoKepingID"
                    ControlSource ="JenisBlankoKepingID"
                    StatusBarText ="1 = Brake Lining; 2= Lempengan/SegiEmpat; 3= Kopling/Donut; 9=Lainnya"

                    LayoutCachedLeft =2450
                    LayoutCachedTop =780
                    LayoutCachedWidth =5790
                    LayoutCachedHeight =1090
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =780
                            Width =2240
                            Height =310
                            Name ="JenisBlankoKepingID_Label"
                            Caption ="JenisBlankoKepingID"
                            LayoutCachedLeft =120
                            LayoutCachedTop =780
                            LayoutCachedWidth =2360
                            LayoutCachedHeight =1090
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2450
                    Top =1200
                    Width =1410
                    Height =310
                    ColumnWidth =1470
                    TabIndex =3
                    Name ="Panjang"
                    ControlSource ="Panjang"
                    Format ="General Number"
                    StatusBarText ="(mm)"

                    LayoutCachedLeft =2450
                    LayoutCachedTop =1200
                    LayoutCachedWidth =3860
                    LayoutCachedHeight =1510
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1200
                            Width =2240
                            Height =310
                            Name ="Panjang_Label"
                            Caption ="Panjang"
                            LayoutCachedLeft =120
                            LayoutCachedTop =1200
                            LayoutCachedWidth =2360
                            LayoutCachedHeight =1510
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6530
                    Top =1200
                    Width =1410
                    Height =310
                    ColumnWidth =1410
                    TabIndex =4
                    Name ="Panjang2"
                    ControlSource ="Panjang2"
                    StatusBarText ="(mm) Untuk Panjang Bagian Bawah"

                    LayoutCachedLeft =6530
                    LayoutCachedTop =1200
                    LayoutCachedWidth =7940
                    LayoutCachedHeight =1510
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4740
                            Top =1200
                            Width =1700
                            Height =310
                            Name ="Panjang2_Label"
                            Caption ="Panjang2"
                            LayoutCachedLeft =4740
                            LayoutCachedTop =1200
                            LayoutCachedWidth =6440
                            LayoutCachedHeight =1510
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2450
                    Top =1620
                    Width =1410
                    Height =310
                    ColumnWidth =1410
                    TabIndex =5
                    Name ="Lebar"
                    ControlSource ="Lebar"
                    StatusBarText ="(mm)"

                    LayoutCachedLeft =2450
                    LayoutCachedTop =1620
                    LayoutCachedWidth =3860
                    LayoutCachedHeight =1930
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1620
                            Width =2240
                            Height =310
                            Name ="Lebar_Label"
                            Caption ="Lebar"
                            LayoutCachedLeft =120
                            LayoutCachedTop =1620
                            LayoutCachedWidth =2360
                            LayoutCachedHeight =1930
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =2450
                    Top =2040
                    TabIndex =6
                    Name ="Tebal_TipisFlag"
                    ControlSource ="Tebal_TipisFlag"
                    StatusBarText ="Untuk ketebalan yang tidak sama antara satu ujung dengan yang lain"

                    LayoutCachedLeft =2450
                    LayoutCachedTop =2040
                    LayoutCachedWidth =2710
                    LayoutCachedHeight =2280
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =2040
                            Width =2240
                            Height =310
                            Name ="Tebal_TipisFlag_Label"
                            Caption ="Tebal_TipisFlag"
                            LayoutCachedLeft =120
                            LayoutCachedTop =2040
                            LayoutCachedWidth =2360
                            LayoutCachedHeight =2350
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2450
                    Top =2460
                    Width =1410
                    Height =310
                    ColumnWidth =3000
                    TabIndex =7
                    Name ="Tebal"
                    ControlSource ="Tebal"
                    StatusBarText ="(mm). Tebal (untuk keping yang tebal-tipis, ketebalan yang tertinggi- yang lebih"
                        " tepis di Tebal2)"

                    LayoutCachedLeft =2450
                    LayoutCachedTop =2460
                    LayoutCachedWidth =3860
                    LayoutCachedHeight =2770
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =2460
                            Width =2240
                            Height =310
                            Name ="Tebal_Label"
                            Caption ="Tebal"
                            LayoutCachedLeft =120
                            LayoutCachedTop =2460
                            LayoutCachedWidth =2360
                            LayoutCachedHeight =2770
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6530
                    Top =2460
                    Width =1410
                    Height =310
                    ColumnWidth =3000
                    TabIndex =8
                    Name ="Tebal2"
                    ControlSource ="Tebal2"
                    StatusBarText ="(mm) Apabila ada yang bagian tebal tipis.  Ketebalan tipis untuk yang tipis"

                    LayoutCachedLeft =6530
                    LayoutCachedTop =2460
                    LayoutCachedWidth =7940
                    LayoutCachedHeight =2770
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4740
                            Top =2460
                            Width =1700
                            Height =310
                            Name ="Tebal2_Label"
                            Caption ="Tebal2"
                            LayoutCachedLeft =4740
                            LayoutCachedTop =2460
                            LayoutCachedWidth =6440
                            LayoutCachedHeight =2770
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2450
                    Top =2880
                    Width =1410
                    Height =310
                    ColumnWidth =1410
                    TabIndex =9
                    Name ="DiameterDrum"
                    ControlSource ="DiameterDrum"
                    Format ="General Number"

                    LayoutCachedLeft =2450
                    LayoutCachedTop =2880
                    LayoutCachedWidth =3860
                    LayoutCachedHeight =3190
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =2880
                            Width =2240
                            Height =310
                            Name ="DiameterDrum_Label"
                            Caption ="DiameterDrum"
                            LayoutCachedLeft =120
                            LayoutCachedTop =2880
                            LayoutCachedWidth =2360
                            LayoutCachedHeight =3190
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6530
                    Top =2880
                    Width =1410
                    Height =310
                    ColumnWidth =1410
                    TabIndex =10
                    Name ="DiameterBrakeShoe"
                    ControlSource ="DiameterBrakeShoe"

                    LayoutCachedLeft =6530
                    LayoutCachedTop =2880
                    LayoutCachedWidth =7940
                    LayoutCachedHeight =3190
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4740
                            Top =2880
                            Width =1700
                            Height =310
                            Name ="DiameterBrakeShoe_Label"
                            Caption ="DiameterBrakeShoe"
                            LayoutCachedLeft =4740
                            LayoutCachedTop =2880
                            LayoutCachedWidth =6440
                            LayoutCachedHeight =3190
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2450
                    Top =3300
                    Width =6640
                    Height =580
                    ColumnWidth =3000
                    TabIndex =11
                    Name ="CatatanBlankoKeping"
                    ControlSource ="CatatanBlankoKeping"

                    LayoutCachedLeft =2450
                    LayoutCachedTop =3300
                    LayoutCachedWidth =9090
                    LayoutCachedHeight =3880
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =3300
                            Width =2240
                            Height =310
                            Name ="CatatanBlankoKeping_Label"
                            Caption ="CatatanBlankoKeping"
                            LayoutCachedLeft =120
                            LayoutCachedTop =3300
                            LayoutCachedWidth =2360
                            LayoutCachedHeight =3610
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
