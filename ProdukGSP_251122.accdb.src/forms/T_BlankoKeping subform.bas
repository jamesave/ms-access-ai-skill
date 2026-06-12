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
    Width =9390
    DatasheetFontHeight =11
    ItemSuffix =18
    Filter ="([T_BlankoKeping].[Lebar] In (163,165))"
    RecSrcDt = Begin
        0x3f09cff57520e640
    End
    RecordSource ="T_BlankoKeping"
    Caption ="T_BlankoKeping subform"
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
            Height =4660
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
                    Name ="BlankoKepingID"
                    ControlSource ="BlankoKepingID"

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
                            Name ="BlankoKepingID_Label"
                            Caption ="BlankoKepingID"
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
                    Name ="DeskripsiBlankoKeping"
                    ControlSource ="DeskripsiBlankoKeping"
                    StatusBarText ="Deskripsi Dasar Blanko (Sebelum lobang/finishing).  Lebih untuk mengetahui cetak"
                        " mould"

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
                            Name ="DeskripsiBlankoKeping_Label"
                            Caption ="DeskripsiBlankoKeping"
                            LayoutCachedLeft =360
                            LayoutCachedTop =780
                            LayoutCachedWidth =2600
                            LayoutCachedHeight =1090
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2690
                    Top =1440
                    Width =1470
                    Height =310
                    ColumnWidth =1470
                    TabIndex =2
                    Name ="Panjang"
                    ControlSource ="Panjang"
                    Format ="General Number"
                    StatusBarText ="(mm)"

                    LayoutCachedLeft =2690
                    LayoutCachedTop =1440
                    LayoutCachedWidth =4160
                    LayoutCachedHeight =1750
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =1440
                            Width =2240
                            Height =310
                            Name ="Panjang_Label"
                            Caption ="Panjang"
                            LayoutCachedLeft =360
                            LayoutCachedTop =1440
                            LayoutCachedWidth =2600
                            LayoutCachedHeight =1750
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2690
                    Top =1860
                    Width =1410
                    Height =310
                    ColumnWidth =1410
                    TabIndex =3
                    Name ="Lebar"
                    ControlSource ="Lebar"
                    StatusBarText ="(mm)"

                    LayoutCachedLeft =2690
                    LayoutCachedTop =1860
                    LayoutCachedWidth =4100
                    LayoutCachedHeight =2170
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =1860
                            Width =2240
                            Height =310
                            Name ="Lebar_Label"
                            Caption ="Lebar"
                            LayoutCachedLeft =360
                            LayoutCachedTop =1860
                            LayoutCachedWidth =2600
                            LayoutCachedHeight =2170
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =2690
                    Top =2280
                    TabIndex =4
                    Name ="Tebal/TipisFlag"
                    ControlSource ="Tebal_TipisFlag"
                    StatusBarText ="Untuk ketebalan yang tidak sama antara satu ujung dengan yang lain"
                    EventProcPrefix ="Tebal_TipisFlag"

                    LayoutCachedLeft =2690
                    LayoutCachedTop =2280
                    LayoutCachedWidth =2950
                    LayoutCachedHeight =2520
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =2280
                            Width =2240
                            Height =310
                            Name ="Tebal/TipisFlag_Label"
                            Caption ="Tebal_TipisFlag"
                            EventProcPrefix ="Tebal_TipisFlag_Label"
                            LayoutCachedLeft =360
                            LayoutCachedTop =2280
                            LayoutCachedWidth =2600
                            LayoutCachedHeight =2590
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2690
                    Top =2700
                    Width =3340
                    Height =310
                    ColumnWidth =3000
                    TabIndex =5
                    Name ="Tebal"
                    ControlSource ="Tebal"
                    StatusBarText ="(mm). Tebal (untuk keping yang tebal-tipis, ketebalan yang tertinggi- yang lebih"
                        " tepis di Tebal2)"

                    LayoutCachedLeft =2690
                    LayoutCachedTop =2700
                    LayoutCachedWidth =6030
                    LayoutCachedHeight =3010
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =2700
                            Width =2240
                            Height =310
                            Name ="Tebal_Label"
                            Caption ="Tebal"
                            LayoutCachedLeft =360
                            LayoutCachedTop =2700
                            LayoutCachedWidth =2600
                            LayoutCachedHeight =3010
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2690
                    Top =3120
                    Width =3340
                    Height =310
                    ColumnWidth =3000
                    TabIndex =6
                    Name ="Tebal2"
                    ControlSource ="Tebal2"
                    StatusBarText ="(mm) Apabila ada yang bagian tebal tipis.  Ketebalan tipis untuk yang tipis"

                    LayoutCachedLeft =2690
                    LayoutCachedTop =3120
                    LayoutCachedWidth =6030
                    LayoutCachedHeight =3430
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =3120
                            Width =2240
                            Height =310
                            Name ="Tebal2_Label"
                            Caption ="Tebal2"
                            LayoutCachedLeft =360
                            LayoutCachedTop =3120
                            LayoutCachedWidth =2600
                            LayoutCachedHeight =3430
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2690
                    Top =3540
                    Width =1410
                    Height =310
                    ColumnWidth =1410
                    TabIndex =7
                    Name ="DiameterDrum"
                    ControlSource ="DiameterDrum"
                    Format ="General Number"

                    LayoutCachedLeft =2690
                    LayoutCachedTop =3540
                    LayoutCachedWidth =4100
                    LayoutCachedHeight =3850
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =3540
                            Width =2240
                            Height =310
                            Name ="DiameterDrum_Label"
                            Caption ="DiameterDrum"
                            LayoutCachedLeft =360
                            LayoutCachedTop =3540
                            LayoutCachedWidth =2600
                            LayoutCachedHeight =3850
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2690
                    Top =3960
                    Width =6640
                    Height =580
                    ColumnWidth =3000
                    TabIndex =8
                    Name ="CatatanBlankoKeping"
                    ControlSource ="CatatanBlankoKeping"

                    LayoutCachedLeft =2690
                    LayoutCachedTop =3960
                    LayoutCachedWidth =9330
                    LayoutCachedHeight =4540
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =3960
                            Width =2240
                            Height =310
                            Name ="CatatanBlankoKeping_Label"
                            Caption ="CatatanBlankoKeping"
                            LayoutCachedLeft =360
                            LayoutCachedTop =3960
                            LayoutCachedWidth =2600
                            LayoutCachedHeight =4270
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
