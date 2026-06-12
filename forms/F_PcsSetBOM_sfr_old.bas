Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =4320
    DatasheetFontHeight =11
    ItemSuffix =23
    Left =1420
    Top =2480
    Right =5830
    Bottom =4430
    RecSrcDt = Begin
        0x03aa9263931fe640
    End
    RecordSource ="T_PcsSetBOM"
    Caption ="F_PcsSetBOM_sfr"
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
            Height =360
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    OverlapFlags =85
                    Width =2280
                    Height =300
                    ForeColor =0
                    Name ="Label16"
                    Caption ="Profil Keping"
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =300
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =3000
                    Width =1200
                    Height =300
                    Name ="IDLbl"
                    Caption ="ID"
                    LayoutCachedLeft =3000
                    LayoutCachedWidth =4200
                    LayoutCachedHeight =300
                End
                Begin Label
                    OverlapFlags =85
                    Left =2340
                    Width =570
                    Height =300
                    ForeColor =0
                    Name ="qtylbl"
                    Caption ="Qty"
                    LayoutCachedLeft =2340
                    LayoutCachedWidth =2910
                    LayoutCachedHeight =300
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            Height =360
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2340
                    Width =570
                    Height =300
                    TabIndex =1
                    Name ="Qty"
                    ControlSource ="Qty"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =2340
                    LayoutCachedWidth =2910
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    IsHyperlink = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3000
                    Width =1200
                    Height =300
                    TabIndex =2
                    Name ="IDLink"
                    ControlSource ="KepingID"

                    LayoutCachedLeft =3000
                    LayoutCachedWidth =4200
                    LayoutCachedHeight =300
                    ForeThemeColorIndex =10
                    ForeTint =100.0
                    HyperlinkTarget ="form.F_ProfilKepingDetail"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Width =2280
                    Height =300
                    Name ="KepingID"
                    ControlSource ="KepingID"
                    OnClick ="[Event Procedure]"

                    LayoutCachedWidth =2280
                    LayoutCachedHeight =300
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
CodeBehindForm
' See "F_PcsSetBOM_sfr_old.cls"
