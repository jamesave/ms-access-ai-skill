Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =1860
    DatasheetFontHeight =11
    ItemSuffix =12
    Left =2860
    Top =2730
    Right =9380
    Bottom =4170
    RecSrcDt = Begin
        0x914e0f651342e640
    End
    RecordSource ="qcal_ProdukNmPcsPerSetCampurKeping"
    Caption ="sfrm_qcal_ProdukNmPcsPerSetCampurKeping"
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    SplitFormDatasheet =1
    SplitFormDatasheet =1
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
            Height =720
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =60
                    Width =360
                    Height =310
                    ColumnWidth =3000
                    BackColor =62207
                    Name ="pcs/set"
                    ControlSource ="pcs/set"
                    EventProcPrefix ="pcs_set"

                    LayoutCachedLeft =60
                    LayoutCachedWidth =420
                    LayoutCachedHeight =310
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =93
                            Left =420
                            Width =720
                            Height =310
                            BackColor =62207
                            Name ="pcs/set_Label"
                            Caption ="pcs/set"
                            EventProcPrefix ="pcs_set_Label"
                            LayoutCachedLeft =420
                            LayoutCachedWidth =1140
                            LayoutCachedHeight =310
                            BackThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =60
                    Top =390
                    TabIndex =1
                    Name ="Check10"
                    ControlSource ="CampurKepingFlagCalc"

                    LayoutCachedLeft =60
                    LayoutCachedTop =390
                    LayoutCachedWidth =320
                    LayoutCachedHeight =630
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =290
                            Top =360
                            Width =1440
                            Height =300
                            Name ="Label11"
                            Caption ="Campur Keping"
                            LayoutCachedLeft =290
                            LayoutCachedTop =360
                            LayoutCachedWidth =1730
                            LayoutCachedHeight =660
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
