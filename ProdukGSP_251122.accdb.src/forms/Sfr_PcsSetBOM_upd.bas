Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6300
    DatasheetFontHeight =11
    ItemSuffix =34
    Left =620
    Top =380
    Right =6920
    Bottom =3970
    AfterInsert ="[Event Procedure]"
    RecSrcDt = Begin
        0x52ef6eb31573e640
    End
    RecordSource ="SELECT T_PcsSetBOM.*, T_ProfilKeping.KepingAutoID, T_ProfilKeping.DeskripsiProfi"
        "lKeping FROM T_ProfilKeping INNER JOIN T_PcsSetBOM ON T_ProfilKeping.KepingID = "
        "T_PcsSetBOM.KepingID; "
    Caption ="Sfr_PcsSetBOM"
    OnCurrent ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
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
            Height =720
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =780
                    Top =360
                    Width =1320
                    Height =300
                    ForeColor =0
                    Name ="Label16"
                    Caption ="Profil Keping"
                    LayoutCachedLeft =780
                    LayoutCachedTop =360
                    LayoutCachedWidth =2100
                    LayoutCachedHeight =660
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =120
                    Top =360
                    Width =600
                    Height =300
                    ForeColor =0
                    Name ="IDLbl"
                    Caption ="KAUID"
                    LayoutCachedLeft =120
                    LayoutCachedTop =360
                    LayoutCachedWidth =720
                    LayoutCachedHeight =660
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =2160
                    Top =360
                    Width =570
                    Height =300
                    ForeColor =0
                    Name ="qtylbl"
                    Caption ="Qty"
                    LayoutCachedLeft =2160
                    LayoutCachedTop =360
                    LayoutCachedWidth =2730
                    LayoutCachedHeight =660
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2340
                    Width =540
                    Height =300
                    FontSize =12
                    FontWeight =700
                    BackColor =62207
                    Name ="txtSumPcsSet"
                    ControlSource ="=Sum([Qty])"

                    LayoutCachedLeft =2340
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =300
                    BackThemeColorIndex =-1
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =93
                            Left =2880
                            Width =1440
                            Height =300
                            FontSize =12
                            FontWeight =700
                            BackColor =62207
                            Name ="lblSumPcsSet"
                            Caption ="  Pcs/Set"
                            LayoutCachedLeft =2880
                            LayoutCachedWidth =4320
                            LayoutCachedHeight =300
                            BackThemeColorIndex =-1
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =120
                    Width =2220
                    Height =300
                    TabIndex =1
                    BackColor =62207
                    ForeColor =255
                    Name ="txtCampurKepingIndicator"

                    LayoutCachedLeft =120
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =300
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
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
                    Left =2160
                    Width =570
                    Height =300
                    Name ="Qty"
                    ControlSource ="Qty"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =2160
                    LayoutCachedWidth =2730
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    IsHyperlink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =120
                    Width =600
                    Height =300
                    TabIndex =1
                    Name ="IDLink"
                    ControlSource ="KepingAutoID"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =120
                    LayoutCachedWidth =720
                    LayoutCachedHeight =300
                    ForeThemeColorIndex =10
                    ForeTint =100.0
                    HyperlinkTarget ="form.F_ProfilKepingDetail"
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =780
                    Width =1320
                    Height =300
                    TabIndex =2
                    Name ="KepingID"
                    ControlSource ="KepingID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT T_ProfilKeping.KepingID FROM T_ProfilKeping; "
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =780
                    LayoutCachedWidth =2100
                    LayoutCachedHeight =300
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    ForeShade =100.0
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    IsHyperlink = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2820
                    Width =2520
                    Height =300
                    TabIndex =3
                    Name ="txtDeskripsiKeping"
                    ControlSource ="DeskripsiProfilKeping"

                    LayoutCachedLeft =2820
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =300
                    ForeThemeColorIndex =10
                    ForeTint =100.0
                    HyperlinkTarget ="form.F_ProfilKepingDetail"
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
' See "Sfr_PcsSetBOM_upd.cls"
