Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6000
    DatasheetFontHeight =11
    ItemSuffix =14
    Right =6000
    Bottom =6530
    RecSrcDt = Begin
        0x184312b3d076e640
    End
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
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            TextFontFamily =0
            FontSize =11
            FontWeight =400
            FontName ="Aptos"
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
        Begin ListBox
            TextFontFamily =0
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Aptos"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            Height =6540
            BackColor =14151142
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =180
                    Top =1080
                    Width =5400
                    Height =540
                    FontSize =20
                    BackColor =13952764
                    Name ="Label0"
                    Caption ="Menu Utama"
                    LayoutCachedLeft =180
                    LayoutCachedTop =1080
                    LayoutCachedWidth =5580
                    LayoutCachedHeight =1620
                    BackThemeColorIndex =-1
                End
                Begin ListBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =180
                    Top =1680
                    Width =4020
                    Height =3600
                    Name ="lst_Menu"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT T_MenuUtama.TargetObject, T_MenuUtama.ItemDescription, T_MenuUtama.Detail"
                        "Description FROM T_MenuUtama ORDER BY T_MenuUtama.SortOrder; "
                    ColumnWidths ="0;1440;0"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =180
                    LayoutCachedTop =1680
                    LayoutCachedWidth =4200
                    LayoutCachedHeight =5280
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4320
                    Top =1680
                    Width =1256
                    Height =576
                    TabIndex =1
                    Name ="cmdPilihMenu"
                    Caption ="Pilih"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Open Form"

                    LayoutCachedLeft =4320
                    LayoutCachedTop =1680
                    LayoutCachedWidth =5576
                    LayoutCachedHeight =2256
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =180
                    Top =5400
                    Width =5400
                    Height =1020
                    TabIndex =2
                    Name ="txt_ItemDesc"

                    LayoutCachedLeft =180
                    LayoutCachedTop =5400
                    LayoutCachedWidth =5580
                    LayoutCachedHeight =6420
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =180
                    Top =120
                    Width =5400
                    Height =910
                    FontSize =18
                    BackColor =3751056
                    ForeColor =5676533
                    Name ="Label6"
                    Caption ="Sireprog \015\012(Sistem Referensi Produk PT GSP)"
                    LayoutCachedLeft =180
                    LayoutCachedTop =120
                    LayoutCachedWidth =5580
                    LayoutCachedHeight =1030
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4320
                    Top =2340
                    Width =1256
                    Height =576
                    TabIndex =3
                    Name ="cmd_BukaFormReferensi"
                    Caption ="Referensi"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Open Form"

                    LayoutCachedLeft =4320
                    LayoutCachedTop =2340
                    LayoutCachedWidth =5576
                    LayoutCachedHeight =2916
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    Left =4320
                    Top =4980
                    Width =1256
                    Height =300
                    FontSize =10
                    ForeColor =10319446
                    Name ="lbl_LinkLingkarRadius"
                    Caption ="Lingkar Busur"
                    OnClick ="[Event Procedure]"
                    LayoutCachedLeft =4320
                    LayoutCachedTop =4980
                    LayoutCachedWidth =5576
                    LayoutCachedHeight =5280
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    Left =4320
                    Top =4620
                    Width =1256
                    Height =300
                    FontSize =10
                    ForeColor =10319446
                    Name ="lbl_SkemaLining"
                    Caption ="Skema Lining"
                    OnClick ="[Event Procedure]"
                    LayoutCachedLeft =4320
                    LayoutCachedTop =4620
                    LayoutCachedWidth =5576
                    LayoutCachedHeight =4920
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
    End
End
CodeBehindForm
' See "F_MenuUtama.cls"
