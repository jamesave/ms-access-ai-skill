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
    ItemSuffix =6
    RecSrcDt = Begin
        0x20a5d8b4b02ce640
    End
    RecordSource ="T_ListPolaLubang"
    Caption ="T_ListPolaLubang subform"
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
            Height =2110
            Name ="Detail"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2690
                    Top =360
                    Width =6640
                    Height =580
                    ColumnWidth =3000
                    Name ="PolaLubang"
                    ControlSource ="PolaLubang"

                    LayoutCachedLeft =2690
                    LayoutCachedTop =360
                    LayoutCachedWidth =9330
                    LayoutCachedHeight =940
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =360
                            Width =2240
                            Height =310
                            Name ="PolaLubang_Label"
                            Caption ="PolaLubang"
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
                    Top =1020
                    Width =6640
                    Height =580
                    ColumnWidth =3000
                    TabIndex =1
                    Name ="LokasiImage"
                    ControlSource ="LokasiImage"

                    LayoutCachedLeft =2690
                    LayoutCachedTop =1020
                    LayoutCachedWidth =9330
                    LayoutCachedHeight =1600
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =1020
                            Width =2240
                            Height =310
                            Name ="LokasiImage_Label"
                            Caption ="LokasiImage"
                            LayoutCachedLeft =360
                            LayoutCachedTop =1020
                            LayoutCachedWidth =2600
                            LayoutCachedHeight =1330
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2690
                    Top =1680
                    Width =3340
                    Height =310
                    ColumnWidth =3000
                    TabIndex =2
                    Name ="JumlahLubang"
                    ControlSource ="JumlahLubang"

                    LayoutCachedLeft =2690
                    LayoutCachedTop =1680
                    LayoutCachedWidth =6030
                    LayoutCachedHeight =1990
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =1680
                            Width =2240
                            Height =310
                            Name ="JumlahLubang_Label"
                            Caption ="JumlahLubang"
                            LayoutCachedLeft =360
                            LayoutCachedTop =1680
                            LayoutCachedWidth =2600
                            LayoutCachedHeight =1990
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
