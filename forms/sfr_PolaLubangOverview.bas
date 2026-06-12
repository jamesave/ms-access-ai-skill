Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    ViewsAllowed =1
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =13889
    DatasheetFontHeight =11
    ItemSuffix =28
    Left =120
    Top =4320
    Right =14390
    Bottom =8150
    OrderBy ="PolaLubangID"
    RecSrcDt = Begin
        0x2dd3b9667640e640
    End
    RecordSource ="T_UkuranPolaLubang"
    Caption ="F_PolaLubangView"
    DatasheetFontName ="Aptos"
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
            TextFontFamily =0
            FontSize =11
            FontName ="Aptos"
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
            TextFontFamily =0
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Aptos"
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
            Height =360
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =660
                    Width =2850
                    Height =315
                    Name ="Catatan_Label"
                    Caption ="Catatan"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =660
                    LayoutCachedWidth =3510
                    LayoutCachedHeight =315
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =1
                    Left =3510
                    Width =780
                    Height =315
                    Name ="L_Label"
                    Caption ="L"
                    OnClick ="[Event Procedure]"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =3510
                    LayoutCachedWidth =4290
                    LayoutCachedHeight =315
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =1
                    Left =4290
                    Width =1020
                    Height =315
                    Name ="X_Label"
                    Caption ="X"
                    OnClick ="[Event Procedure]"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =4290
                    LayoutCachedWidth =5310
                    LayoutCachedHeight =315
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =1
                    Left =5310
                    Width =720
                    Height =315
                    Name ="A_Label"
                    Caption ="A"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =5310
                    LayoutCachedWidth =6030
                    LayoutCachedHeight =315
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =1
                    Left =6030
                    Width =780
                    Height =315
                    Name ="B_Label"
                    Caption ="B"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =6030
                    LayoutCachedWidth =6810
                    LayoutCachedHeight =315
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =1
                    Left =6810
                    Width =780
                    Height =315
                    Name ="C_Label"
                    Caption ="C"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =6810
                    LayoutCachedWidth =7590
                    LayoutCachedHeight =315
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =1
                    Left =7590
                    Width =780
                    Height =315
                    Name ="D_Label"
                    Caption ="D"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =7590
                    LayoutCachedWidth =8370
                    LayoutCachedHeight =315
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =1
                    Left =8370
                    Width =780
                    Height =315
                    Name ="E_Label"
                    Caption ="E"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =8370
                    LayoutCachedWidth =9150
                    LayoutCachedHeight =315
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =1
                    Left =9150
                    Width =780
                    Height =315
                    Name ="F_Label"
                    Caption ="F"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =9150
                    LayoutCachedWidth =9930
                    LayoutCachedHeight =315
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =1
                    Left =9930
                    Width =780
                    Height =315
                    Name ="G_Label"
                    Caption ="G"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =9930
                    LayoutCachedWidth =10710
                    LayoutCachedHeight =315
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =1
                    Left =10710
                    Width =780
                    Height =315
                    Name ="H_Label"
                    Caption ="H"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =10710
                    LayoutCachedWidth =11490
                    LayoutCachedHeight =315
                End
                Begin Label
                    OverlapFlags =85
                    Left =60
                    Width =540
                    Height =300
                    Name ="PolaLubang_Lbl"
                    Caption ="PL_ID"
                    LayoutCachedLeft =60
                    LayoutCachedWidth =600
                    LayoutCachedHeight =300
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
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =660
                    Width =2850
                    Height =300
                    ColumnWidth =3000
                    Name ="Catatan"
                    ControlSource ="DetailPolaLubang"
                    StatusBarText ="Informasi Tentang Jarak Lobang yang mungkin berbeda dengan gambar"

                    LayoutCachedLeft =660
                    LayoutCachedWidth =3510
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =3510
                    Width =780
                    Height =300
                    ColumnWidth =3000
                    TabIndex =1
                    Name ="L"
                    ControlSource ="L"

                    LayoutCachedLeft =3510
                    LayoutCachedWidth =4290
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =4290
                    Width =1020
                    Height =300
                    ColumnWidth =3000
                    TabIndex =2
                    Name ="X"
                    ControlSource ="X"

                    LayoutCachedLeft =4290
                    LayoutCachedWidth =5310
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =5310
                    Width =720
                    Height =300
                    ColumnWidth =3000
                    TabIndex =3
                    Name ="A"
                    ControlSource ="A"

                    LayoutCachedLeft =5310
                    LayoutCachedWidth =6030
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =6030
                    Width =780
                    Height =300
                    ColumnWidth =3000
                    TabIndex =4
                    Name ="B"
                    ControlSource ="B"

                    LayoutCachedLeft =6030
                    LayoutCachedWidth =6810
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =6810
                    Width =780
                    Height =300
                    ColumnWidth =3000
                    TabIndex =5
                    Name ="C"
                    ControlSource ="C"

                    LayoutCachedLeft =6810
                    LayoutCachedWidth =7590
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =7590
                    Width =780
                    Height =300
                    ColumnWidth =3000
                    TabIndex =6
                    Name ="D"
                    ControlSource ="D"

                    LayoutCachedLeft =7590
                    LayoutCachedWidth =8370
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =8370
                    Width =780
                    Height =300
                    ColumnWidth =3000
                    TabIndex =7
                    Name ="E"
                    ControlSource ="E"

                    LayoutCachedLeft =8370
                    LayoutCachedWidth =9150
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =9150
                    Width =780
                    Height =300
                    ColumnWidth =3000
                    TabIndex =8
                    Name ="F"
                    ControlSource ="F"

                    LayoutCachedLeft =9150
                    LayoutCachedWidth =9930
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =9930
                    Width =780
                    Height =300
                    ColumnWidth =3000
                    TabIndex =9
                    Name ="G"
                    ControlSource ="G"

                    LayoutCachedLeft =9930
                    LayoutCachedWidth =10710
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =10710
                    Width =780
                    Height =300
                    ColumnWidth =3000
                    TabIndex =10
                    Name ="H"
                    ControlSource ="H"

                    LayoutCachedLeft =10710
                    LayoutCachedWidth =11490
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    IsHyperlink = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =60
                    Width =540
                    Height =300
                    ColumnWidth =2010
                    TabIndex =11
                    Name ="PolaLubangID"
                    ControlSource ="PolaLubangID"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =60
                    LayoutCachedWidth =600
                    LayoutCachedHeight =300
                    ForeThemeColorIndex =10
                    ForeTint =100.0
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
CodeBehindForm
' See "sfr_PolaLubangOverview.cls"
