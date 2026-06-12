Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    ViewsAllowed =1
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7560
    DatasheetFontHeight =11
    ItemSuffix =26
    Left =7880
    Top =420
    Right =15440
    Bottom =10430
    Filter ="[JumlahLubang] = 16"
    OrderBy ="T_ListPolaLubang.JumlahLubang"
    RecSrcDt = Begin
        0x3d3ba4c90b6fe640
    End
    RecordSource ="SELECT T_ListPolaLubang.JumlahLubang, T_ListPolaLubang.PolaLubang, T_ListPolaLub"
        "ang.LokasiImage, GetHPImagePath ([PolaLubang]) AS ImagePath FROM T_ListPolaLuban"
        "g ORDER BY T_ListPolaLubang.JumlahLubang; "
    Caption ="Qry_ListPolaLubang_byHoleQty"
    OnOpen ="[Event Procedure]"
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
        Begin ComboBox
            AddColon = NotDefault
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
            Height =1200
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
                    Top =60
                    Width =2280
                    Height =780
                    BackColor =8421504
                    Name ="Box19"
                    LayoutCachedLeft =5160
                    LayoutCachedTop =60
                    LayoutCachedWidth =7440
                    LayoutCachedHeight =840
                    BackThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =270
                    Top =840
                    Width =390
                    Height =300
                    Name ="JumlahLubang_Label"
                    Caption ="Qty"
                    OnClick ="[Event Procedure]"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =270
                    LayoutCachedTop =840
                    LayoutCachedWidth =660
                    LayoutCachedHeight =1140
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =720
                    Top =840
                    Width =1620
                    Height =300
                    Name ="PolaLubang_Label"
                    Caption ="PolaLubang"
                    OnClick ="[Event Procedure]"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =720
                    LayoutCachedTop =840
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =1140
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
                    Caption ="Tambah Pola "
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =3720
                    LayoutCachedTop =480
                    LayoutCachedWidth =5070
                    LayoutCachedHeight =840
                    Overlaps =1
                End
                Begin ComboBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =5220
                    Top =360
                    Width =1320
                    Height =360
                    TabIndex =2
                    Name ="JmlLubangFilt"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT JumlahLubang, PolaLubang, LokasiImage FROM T_ListPolaLubang ORDE"
                        "R BY JumlahLubang; "
                    AfterUpdate ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"
                    OnLostFocus ="[Event Procedure]"

                    LayoutCachedLeft =5220
                    LayoutCachedTop =360
                    LayoutCachedWidth =6540
                    LayoutCachedHeight =720
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =2
                            OverlapFlags =223
                            Left =5160
                            Top =60
                            Width =2280
                            Height =780
                            Name ="JumlahLubangFilter"
                            Caption ="    Filter Jumlah Lubang"
                            LayoutCachedLeft =5160
                            LayoutCachedTop =60
                            LayoutCachedWidth =7440
                            LayoutCachedHeight =840
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =6600
                    Top =360
                    Width =776
                    TabIndex =1
                    Name ="ClearFilter"
                    Caption ="Hapus"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add Record"

                    LayoutCachedLeft =6600
                    LayoutCachedTop =360
                    LayoutCachedWidth =7376
                    LayoutCachedHeight =720
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3720
                    Top =60
                    Width =1350
                    TabIndex =3
                    Name ="Command18"
                    Caption ="Edit Pola "

                    LayoutCachedLeft =3720
                    LayoutCachedTop =60
                    LayoutCachedWidth =5070
                    LayoutCachedHeight =420
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =2940
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =600
                    Height =310
                    ColumnWidth =3000
                    ColumnOrder =0
                    Name ="JumlahLubang"
                    ControlSource ="JumlahLubang"

                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =660
                    LayoutCachedHeight =370
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    EnterKeyBehavior = NotDefault
                    IsHyperlink = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =720
                    Top =60
                    Width =1620
                    Height =310
                    ColumnWidth =3000
                    ColumnOrder =1
                    TabIndex =1
                    Name ="PolaLubang"
                    ControlSource ="PolaLubang"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =720
                    LayoutCachedTop =60
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =370
                    ForeThemeColorIndex =10
                    ForeTint =100.0
                End
                Begin Image
                    PictureType =1
                    Left =2400
                    Top =60
                    Width =4621
                    Height =2821
                    Name ="PolaLubangImage"
                    ControlSource ="ImagePath"

                    LayoutCachedLeft =2400
                    LayoutCachedTop =60
                    LayoutCachedWidth =7021
                    LayoutCachedHeight =2881
                    TabIndex =2
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
' See "F_DaftarPolaLubang.cls"
