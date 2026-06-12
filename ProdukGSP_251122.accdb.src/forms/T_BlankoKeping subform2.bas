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
    ItemSuffix =24
    OrderBy ="[T_BlankoKeping].[DeskripsiBlankoKeping]"
    RecSrcDt = Begin
        0xed59d650aa43e640
    End
    RecordSource ="T_BlankoKeping"
    Caption ="T_BlankoKeping subform2"
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
            Height =6160
            Name ="Detail"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2690
                    Top =360
                    Height =300
                    ColumnWidth =1440
                    Name ="BlankoKepingAutoID"
                    ControlSource ="BlankoKepingAutoID"

                    LayoutCachedLeft =2690
                    LayoutCachedTop =360
                    LayoutCachedWidth =4130
                    LayoutCachedHeight =660
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =360
                            Width =2240
                            Height =310
                            Name ="BlankoKepingAutoID_Label"
                            Caption ="BlankoKepingAutoID"
                            LayoutCachedLeft =360
                            LayoutCachedTop =360
                            LayoutCachedWidth =2600
                            LayoutCachedHeight =670
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2690
                    Top =780
                    Width =6640
                    Height =580
                    ColumnWidth =3000
                    TabIndex =1
                    Name ="BlankoKepingID"
                    ControlSource ="BlankoKepingID"

                    LayoutCachedLeft =2690
                    LayoutCachedTop =780
                    LayoutCachedWidth =9330
                    LayoutCachedHeight =1360
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =780
                            Width =2240
                            Height =310
                            Name ="BlankoKepingID_Label"
                            Caption ="BlankoKepingID"
                            LayoutCachedLeft =360
                            LayoutCachedTop =780
                            LayoutCachedWidth =2600
                            LayoutCachedHeight =1090
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2690
                    Top =1440
                    Width =6640
                    Height =580
                    ColumnWidth =3000
                    TabIndex =2
                    Name ="DeskripsiBlankoKeping"
                    ControlSource ="DeskripsiBlankoKeping"
                    StatusBarText ="Deskripsi Dasar Blanko (Sebelum lobang/finishing).  Lebih untuk mengetahui cetak"
                        " mould"

                    LayoutCachedLeft =2690
                    LayoutCachedTop =1440
                    LayoutCachedWidth =9330
                    LayoutCachedHeight =2020
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =1440
                            Width =2240
                            Height =310
                            Name ="DeskripsiBlankoKeping_Label"
                            Caption ="DeskripsiBlankoKeping"
                            LayoutCachedLeft =360
                            LayoutCachedTop =1440
                            LayoutCachedWidth =2600
                            LayoutCachedHeight =1750
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2690
                    Top =2100
                    Width =1470
                    Height =310
                    ColumnWidth =1470
                    TabIndex =3
                    Name ="Panjang"
                    ControlSource ="Panjang"
                    Format ="General Number"
                    StatusBarText ="(mm)"

                    LayoutCachedLeft =2690
                    LayoutCachedTop =2100
                    LayoutCachedWidth =4160
                    LayoutCachedHeight =2410
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =2100
                            Width =2240
                            Height =310
                            Name ="Panjang_Label"
                            Caption ="Panjang"
                            LayoutCachedLeft =360
                            LayoutCachedTop =2100
                            LayoutCachedWidth =2600
                            LayoutCachedHeight =2410
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2690
                    Top =2520
                    Width =1410
                    Height =310
                    ColumnWidth =1410
                    TabIndex =4
                    Name ="Panjang2"
                    ControlSource ="Panjang2"
                    StatusBarText ="(mm) Untuk Panjang Bagian Bawah"

                    LayoutCachedLeft =2690
                    LayoutCachedTop =2520
                    LayoutCachedWidth =4100
                    LayoutCachedHeight =2830
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =2520
                            Width =2240
                            Height =310
                            Name ="Panjang2_Label"
                            Caption ="Panjang2"
                            LayoutCachedLeft =360
                            LayoutCachedTop =2520
                            LayoutCachedWidth =2600
                            LayoutCachedHeight =2830
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2690
                    Top =2940
                    Width =1410
                    Height =310
                    ColumnWidth =1410
                    TabIndex =5
                    Name ="Lebar"
                    ControlSource ="Lebar"
                    StatusBarText ="(mm)"

                    LayoutCachedLeft =2690
                    LayoutCachedTop =2940
                    LayoutCachedWidth =4100
                    LayoutCachedHeight =3250
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =2940
                            Width =2240
                            Height =310
                            Name ="Lebar_Label"
                            Caption ="Lebar"
                            LayoutCachedLeft =360
                            LayoutCachedTop =2940
                            LayoutCachedWidth =2600
                            LayoutCachedHeight =3250
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =2690
                    Top =3360
                    TabIndex =6
                    Name ="Tebal_TipisFlag"
                    ControlSource ="Tebal_TipisFlag"
                    StatusBarText ="Untuk ketebalan yang tidak sama antara satu ujung dengan yang lain"

                    LayoutCachedLeft =2690
                    LayoutCachedTop =3360
                    LayoutCachedWidth =2950
                    LayoutCachedHeight =3600
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =3360
                            Width =2240
                            Height =310
                            Name ="Tebal_TipisFlag_Label"
                            Caption ="Tebal_TipisFlag"
                            LayoutCachedLeft =360
                            LayoutCachedTop =3360
                            LayoutCachedWidth =2600
                            LayoutCachedHeight =3670
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2690
                    Top =3780
                    Width =3340
                    Height =310
                    ColumnWidth =3000
                    TabIndex =7
                    Name ="Tebal"
                    ControlSource ="Tebal"
                    StatusBarText ="(mm). Tebal (untuk keping yang tebal-tipis, ketebalan yang tertinggi- yang lebih"
                        " tepis di Tebal2)"

                    LayoutCachedLeft =2690
                    LayoutCachedTop =3780
                    LayoutCachedWidth =6030
                    LayoutCachedHeight =4090
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =3780
                            Width =2240
                            Height =310
                            Name ="Tebal_Label"
                            Caption ="Tebal"
                            LayoutCachedLeft =360
                            LayoutCachedTop =3780
                            LayoutCachedWidth =2600
                            LayoutCachedHeight =4090
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2690
                    Top =4200
                    Width =3340
                    Height =310
                    ColumnWidth =3000
                    TabIndex =8
                    Name ="Tebal2"
                    ControlSource ="Tebal2"
                    StatusBarText ="(mm) Apabila ada yang bagian tebal tipis.  Ketebalan tipis untuk yang tipis"

                    LayoutCachedLeft =2690
                    LayoutCachedTop =4200
                    LayoutCachedWidth =6030
                    LayoutCachedHeight =4510
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =4200
                            Width =2240
                            Height =310
                            Name ="Tebal2_Label"
                            Caption ="Tebal2"
                            LayoutCachedLeft =360
                            LayoutCachedTop =4200
                            LayoutCachedWidth =2600
                            LayoutCachedHeight =4510
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2690
                    Top =4620
                    Width =1410
                    Height =310
                    ColumnWidth =1410
                    TabIndex =9
                    Name ="DiameterDrum"
                    ControlSource ="DiameterDrum"
                    Format ="General Number"

                    LayoutCachedLeft =2690
                    LayoutCachedTop =4620
                    LayoutCachedWidth =4100
                    LayoutCachedHeight =4930
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =4620
                            Width =2240
                            Height =310
                            Name ="DiameterDrum_Label"
                            Caption ="DiameterDrum"
                            LayoutCachedLeft =360
                            LayoutCachedTop =4620
                            LayoutCachedWidth =2600
                            LayoutCachedHeight =4930
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2690
                    Top =5040
                    Width =1410
                    Height =310
                    ColumnWidth =1410
                    TabIndex =10
                    Name ="DiameterBrakeShoe"
                    ControlSource ="DiameterBrakeShoe"

                    LayoutCachedLeft =2690
                    LayoutCachedTop =5040
                    LayoutCachedWidth =4100
                    LayoutCachedHeight =5350
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =5040
                            Width =2240
                            Height =310
                            Name ="DiameterBrakeShoe_Label"
                            Caption ="DiameterBrakeShoe"
                            LayoutCachedLeft =360
                            LayoutCachedTop =5040
                            LayoutCachedWidth =2600
                            LayoutCachedHeight =5350
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2690
                    Top =5460
                    Width =6640
                    Height =580
                    ColumnWidth =3000
                    TabIndex =11
                    Name ="CatatanBlankoKeping"
                    ControlSource ="CatatanBlankoKeping"

                    LayoutCachedLeft =2690
                    LayoutCachedTop =5460
                    LayoutCachedWidth =9330
                    LayoutCachedHeight =6040
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =5460
                            Width =2240
                            Height =310
                            Name ="CatatanBlankoKeping_Label"
                            Caption ="CatatanBlankoKeping"
                            LayoutCachedLeft =360
                            LayoutCachedTop =5460
                            LayoutCachedWidth =2600
                            LayoutCachedHeight =5770
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
