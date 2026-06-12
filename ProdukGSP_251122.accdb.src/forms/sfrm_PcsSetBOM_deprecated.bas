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
    ItemSuffix =10
    RecSrcDt = Begin
        0x417586f43642e640
    End
    RecordSource ="T_PcsSetBOM"
    Caption ="T_PcsSetBOM subform"
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
            Height =2710
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
                    Name ="SetBOMID"
                    ControlSource ="SetBOMID"

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
                            Name ="SetBOMID_Label"
                            Caption ="SetBOMID"
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
                    Width =1410
                    Height =310
                    ColumnWidth =1410
                    TabIndex =1
                    Name ="BaseProdukID"
                    ControlSource ="BaseProdukID"

                    LayoutCachedLeft =2690
                    LayoutCachedTop =780
                    LayoutCachedWidth =4100
                    LayoutCachedHeight =1090
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =780
                            Width =2240
                            Height =310
                            Name ="BaseProdukID_Label"
                            Caption ="BaseProdukID"
                            LayoutCachedLeft =360
                            LayoutCachedTop =780
                            LayoutCachedWidth =2600
                            LayoutCachedHeight =1090
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2690
                    Top =1200
                    Width =2240
                    Height =310
                    ColumnWidth =2240
                    TabIndex =2
                    Name ="SetID_deprecated"
                    ControlSource ="SetID_deprecated"
                    StatusBarText ="Name is in BaseProduct table"

                    LayoutCachedLeft =2690
                    LayoutCachedTop =1200
                    LayoutCachedWidth =4930
                    LayoutCachedHeight =1510
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =1200
                            Width =2240
                            Height =310
                            Name ="SetID_deprecated_Label"
                            Caption ="SetID_deprecated"
                            LayoutCachedLeft =360
                            LayoutCachedTop =1200
                            LayoutCachedWidth =2600
                            LayoutCachedHeight =1510
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2690
                    Top =1620
                    Width =6640
                    Height =580
                    ColumnWidth =3000
                    TabIndex =3
                    Name ="KepingID"
                    ControlSource ="KepingID"

                    LayoutCachedLeft =2690
                    LayoutCachedTop =1620
                    LayoutCachedWidth =9330
                    LayoutCachedHeight =2200
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =1620
                            Width =2240
                            Height =310
                            Name ="KepingID_Label"
                            Caption ="KepingID"
                            LayoutCachedLeft =360
                            LayoutCachedTop =1620
                            LayoutCachedWidth =2600
                            LayoutCachedHeight =1930
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2690
                    Top =2280
                    Width =1410
                    Height =310
                    ColumnWidth =1410
                    TabIndex =4
                    Name ="Qty"
                    ControlSource ="Qty"

                    LayoutCachedLeft =2690
                    LayoutCachedTop =2280
                    LayoutCachedWidth =4100
                    LayoutCachedHeight =2590
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =2280
                            Width =2240
                            Height =310
                            Name ="Qty_Label"
                            Caption ="Qty"
                            LayoutCachedLeft =360
                            LayoutCachedTop =2280
                            LayoutCachedWidth =2600
                            LayoutCachedHeight =2590
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
