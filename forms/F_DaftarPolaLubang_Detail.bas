Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    FilterOn = NotDefault
    OrderByOn = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =21960
    DatasheetFontHeight =11
    ItemSuffix =30
    Left =10230
    Top =170
    Right =24130
    Bottom =7630
    Filter ="[JumlahLubang] = 6"
    OrderBy ="T_ListPolaLubang.JumlahLubang"
    OnApplyFilter ="[Event Procedure]"
    RecSrcDt = Begin
        0x3d3ba4c90b6fe640
    End
    RecordSource ="SELECT T_ListPolaLubang.JumlahLubang, T_ListPolaLubang.PolaLubang, T_ListPolaLub"
        "ang.LokasiImage, GetHPImagePath ([PolaLubang]) AS ImagePath FROM T_ListPolaLuban"
        "g ORDER BY T_ListPolaLubang.JumlahLubang; "
    Caption ="Produk Dengan Pola Lubang"
    OnOpen ="[Event Procedure]"
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
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
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
        Begin ListBox
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
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin FormHeader
            Height =1380
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Rectangle
                    BackStyle =1
                    OverlapFlags =93
                    Left =5160
                    Width =3120
                    Height =1320
                    BackColor =8421504
                    Name ="Box19"
                    LayoutCachedLeft =5160
                    LayoutCachedWidth =8280
                    LayoutCachedHeight =1320
                    BackThemeColorIndex =-1
                End
                Begin ComboBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =6480
                    Top =600
                    Width =1020
                    Height =300
                    TabIndex =2
                    Name ="lsb_JmlLubangFilt"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT JumlahLubang, PolaLubang, LokasiImage FROM T_ListPolaLubang ORDE"
                        "R BY JumlahLubang; "
                    AfterUpdate ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"
                    OnLostFocus ="[Event Procedure]"

                    LayoutCachedLeft =6480
                    LayoutCachedTop =600
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =900
                    Begin
                        Begin Label
                            BorderWidth =2
                            OverlapFlags =215
                            Left =5400
                            Top =600
                            Width =850
                            Height =300
                            Name ="lbl_JumlahLubangFilter"
                            Caption ="Qty Lbng"
                            LayoutCachedLeft =5400
                            LayoutCachedTop =600
                            LayoutCachedWidth =6250
                            LayoutCachedHeight =900
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =360
                    Top =900
                    Width =660
                    Height =300
                    Name ="JumlahLubang_Label"
                    Caption ="Jumlah"
                    OnClick ="[Event Procedure]"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =360
                    LayoutCachedTop =900
                    LayoutCachedWidth =1020
                    LayoutCachedHeight =1200
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =1140
                    Top =900
                    Width =1620
                    Height =300
                    Name ="PolaLubang_Label"
                    Caption ="PolaLubang"
                    OnClick ="[Event Procedure]"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =1140
                    LayoutCachedTop =900
                    LayoutCachedWidth =2760
                    LayoutCachedHeight =1200
                End
                Begin Label
                    OverlapFlags =85
                    Left =60
                    Top =60
                    Width =3420
                    Height =540
                    FontSize =20
                    Name ="Label6"
                    Caption ="Daftar Pola Lubang"
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =600
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3720
                    Top =480
                    Width =1350
                    Name ="Tambah"
                    Caption ="Tambah "
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =3720
                    LayoutCachedTop =480
                    LayoutCachedWidth =5070
                    LayoutCachedHeight =840
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =6600
                    Top =60
                    Width =776
                    TabIndex =1
                    Name ="cmd_ClearFilter"
                    Caption ="Hapus"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add Record"

                    LayoutCachedLeft =6600
                    LayoutCachedTop =60
                    LayoutCachedWidth =7376
                    LayoutCachedHeight =420
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3720
                    Top =60
                    Width =1350
                    TabIndex =3
                    Name ="Command18"
                    Caption ="Edit "

                    LayoutCachedLeft =3720
                    LayoutCachedTop =60
                    LayoutCachedWidth =5070
                    LayoutCachedHeight =420
                    Overlaps =1
                End
                Begin ComboBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =6480
                    Top =960
                    Width =1020
                    Height =300
                    TabIndex =4
                    Name ="lsb_PolaLubangFiltr"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT JumlahLubang, PolaLubang, LokasiImage FROM T_ListPolaLubang ORDE"
                        "R BY JumlahLubang; "

                    LayoutCachedLeft =6480
                    LayoutCachedTop =960
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =1260
                    Begin
                        Begin Label
                            BorderWidth =2
                            OverlapFlags =215
                            Left =5400
                            Top =960
                            Width =850
                            Height =300
                            Name ="lbl_PolaLubangFilt"
                            Caption ="Pola"
                            LayoutCachedLeft =5400
                            LayoutCachedTop =960
                            LayoutCachedWidth =6250
                            LayoutCachedHeight =1260
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                        End
                    End
                End
                Begin Label
                    FontUnderline = NotDefault
                    OldBorderStyle =1
                    BorderWidth =2
                    OverlapFlags =215
                    TextAlign =2
                    Left =5400
                    Top =60
                    Width =1080
                    Height =360
                    FontSize =14
                    FontWeight =700
                    Name ="Label29"
                    Caption ="Filter"
                    LayoutCachedLeft =5400
                    LayoutCachedTop =60
                    LayoutCachedWidth =6480
                    LayoutCachedHeight =420
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =7200
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =360
                    Top =60
                    Width =660
                    Height =310
                    ColumnWidth =3000
                    ColumnOrder =0
                    Name ="JumlahLubang"
                    ControlSource ="JumlahLubang"

                    LayoutCachedLeft =360
                    LayoutCachedTop =60
                    LayoutCachedWidth =1020
                    LayoutCachedHeight =370
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1140
                    Top =60
                    Width =1620
                    Height =310
                    ColumnWidth =3000
                    ColumnOrder =1
                    TabIndex =1
                    Name ="PolaLubang"
                    ControlSource ="PolaLubang"

                    LayoutCachedLeft =1140
                    LayoutCachedTop =60
                    LayoutCachedWidth =2760
                    LayoutCachedHeight =370
                End
                Begin Image
                    PictureType =1
                    Left =2880
                    Top =60
                    Width =4621
                    Height =2821
                    Name ="PolaLubangImage"
                    ControlSource ="ImagePath"

                    LayoutCachedLeft =2880
                    LayoutCachedTop =60
                    LayoutCachedWidth =7501
                    LayoutCachedHeight =2881
                    TabIndex =3
                End
                Begin Subform
                    OverlapFlags =85
                    Left =120
                    Top =2940
                    Width =14280
                    Height =3840
                    TabIndex =2
                    Name ="F_PolaLubangView"
                    SourceObject ="Form.sfr_PolaLubangOverview"
                    LinkChildFields ="PolaLubang"
                    LinkMasterFields ="PolaLubang"

                    LayoutCachedLeft =120
                    LayoutCachedTop =2940
                    LayoutCachedWidth =14400
                    LayoutCachedHeight =6780
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
' See "F_DaftarPolaLubang_Detail.cls"
