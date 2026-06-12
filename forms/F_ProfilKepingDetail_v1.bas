Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    DividingLines = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10860
    DatasheetFontHeight =11
    ItemSuffix =53
    Left =170
    Top =4200
    Right =11290
    Bottom =10500
    RecSrcDt = Begin
        0x21e3b371c85be640
    End
    RecordSource ="Qry_ProfilKepingDetail_Primary"
    OnOpen ="[Event Procedure]"
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
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin Tab
            FontSize =11
            FontName ="Calibri Light"
            ThemeFontIndex =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =3
            BackThemeColorIndex =1
            BackShade =85.0
            BorderLineStyle =0
            BorderThemeColorIndex =2
            BorderTint =60.0
            HoverThemeColorIndex =1
            PressedThemeColorIndex =1
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
            ForeThemeColorIndex =0
            ForeTint =75.0
        End
        Begin Page
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin PageHeader
            DisplayWhen =1
            Height =0
            BackColor =5676533
            Name ="PageHeaderSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
        End
        Begin Section
            CanGrow = NotDefault
            Height =6300
            BackColor =5676533
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =93
                    Top =360
                    Width =4500
                    Height =780
                    Name ="Box49"
                    LayoutCachedTop =360
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =1140
                    BackThemeColorIndex =4
                    BackTint =60.0
                    BorderThemeColorIndex =8
                    BorderTint =40.0
                    BorderShade =100.0
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =95
                    Left =4500
                    Top =60
                    Width =4320
                    Height =1980
                    Name ="Box18"
                    LayoutCachedLeft =4500
                    LayoutCachedTop =60
                    LayoutCachedWidth =8820
                    LayoutCachedHeight =2040
                    BackThemeColorIndex =4
                    BackTint =60.0
                    BorderThemeColorIndex =8
                    BorderTint =40.0
                    BorderShade =100.0
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =1560
                    Top =60
                    Width =2820
                    Height =300
                    ColumnOrder =3
                    ForeColor =0
                    Name ="KepingID"
                    ControlSource ="KepingID"

                    LayoutCachedLeft =1560
                    LayoutCachedTop =60
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =360
                    ForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =87
                            Left =60
                            Top =60
                            Width =1190
                            Height =300
                            ForeColor =0
                            Name ="ProfilKepingLbl"
                            Caption ="Profil Keping"
                            LayoutCachedLeft =60
                            LayoutCachedTop =60
                            LayoutCachedWidth =1250
                            LayoutCachedHeight =360
                            ForeTint =100.0
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =1440
                    Left =1560
                    Top =420
                    Width =2820
                    Height =300
                    ColumnWidth =1425
                    ColumnOrder =4
                    TabIndex =1
                    BoundColumn =1
                    Name ="BlankoKepingID"
                    ControlSource ="BlankoKepingID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT T_BlankoKeping.BlankoKepingAutoID, T_BlankoKeping.BlankoKepingID, T_Blank"
                        "oKeping.DeskripsiBlankoKeping FROM T_BlankoKeping; "
                    ColumnWidths ="0;1440;0"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =1560
                    LayoutCachedTop =420
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =720
                    ForeThemeColorIndex =0
                    ForeShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =60
                            Top =600
                            Width =1480
                            Height =300
                            ForeColor =0
                            Name ="BlankoKepingLbl"
                            Caption ="BlankoKepingID"
                            LayoutCachedLeft =60
                            LayoutCachedTop =600
                            LayoutCachedWidth =1540
                            LayoutCachedHeight =900
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =1560
                    Top =1140
                    Width =2820
                    Height =900
                    ColumnWidth =4300
                    ColumnOrder =5
                    TabIndex =2
                    ForeColor =0
                    Name ="DeskripsiKeping"
                    ControlSource ="DeskripsiProfilKeping"

                    LayoutCachedLeft =1560
                    LayoutCachedTop =1140
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =2040
                    ForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =87
                            Left =60
                            Top =1140
                            Width =1490
                            Height =300
                            ForeColor =0
                            Name ="Label4"
                            Caption ="DeskripsiProfilKeping"
                            LayoutCachedLeft =60
                            LayoutCachedTop =1140
                            LayoutCachedWidth =1550
                            LayoutCachedHeight =1440
                            ForeTint =100.0
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =3240
                    Top =2190
                    ColumnOrder =8
                    TabIndex =3
                    Name ="LubangFlag"
                    ControlSource ="LubangFlag"

                    LayoutCachedLeft =3240
                    LayoutCachedTop =2190
                    LayoutCachedWidth =3500
                    LayoutCachedHeight =2430
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3500
                            Top =2160
                            Width =2140
                            Height =300
                            ForeColor =0
                            Name ="Label7"
                            Caption ="Keping Berlobang Rivet"
                            LayoutCachedLeft =3500
                            LayoutCachedTop =2160
                            LayoutCachedWidth =5640
                            LayoutCachedHeight =2460
                            ForeTint =100.0
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =2550
                    Left =4200
                    Top =4740
                    Height =300
                    ColumnOrder =10
                    TabIndex =4
                    Name ="PolaLubangID"
                    ControlSource ="T_ProfilKeping.PolaLubangID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [T_UkuranPolaLubang].PolaLubangID, [T_UkuranPolaLubang].PolaLubang, T_Uku"
                        "ranPolaLubang.[DetailPolaLubang] FROM T_UkuranPolaLubang ORDER BY [T_UkuranPolaL"
                        "ubang].PolaLubang; "
                    ColumnWidths ="0;650;1900"
                    AllowValueListEdits =0

                    LayoutCachedLeft =4200
                    LayoutCachedTop =4740
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =5040
                    ForeThemeColorIndex =0
                    ForeShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2700
                            Top =4740
                            Width =1410
                            Height =300
                            ForeColor =0
                            Name ="Label9"
                            Caption ="T_ProfilKeping.PolaLubangID"
                            LayoutCachedLeft =2700
                            LayoutCachedTop =4740
                            LayoutCachedWidth =4110
                            LayoutCachedHeight =5040
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =1
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5940
                    Top =540
                    Width =600
                    Height =300
                    ColumnWidth =1020
                    ColumnOrder =12
                    TabIndex =5
                    BackColor =14396046
                    ForeColor =0
                    Name ="Panjang"
                    ControlSource ="Panjang"
                    Format ="General Number"
                    StatusBarText ="(mm)"

                    LayoutCachedLeft =5940
                    LayoutCachedTop =540
                    LayoutCachedWidth =6540
                    LayoutCachedHeight =840
                    BackThemeColorIndex =-1
                    ForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =4800
                            Top =540
                            Width =1080
                            Height =300
                            ForeColor =0
                            Name ="Label13"
                            Caption ="Panjang:"
                            LayoutCachedLeft =4800
                            LayoutCachedTop =540
                            LayoutCachedWidth =5880
                            LayoutCachedHeight =840
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5940
                    Top =900
                    Width =600
                    Height =300
                    ColumnWidth =1020
                    ColumnOrder =13
                    TabIndex =6
                    BackColor =14396046
                    ForeColor =0
                    Name ="Lebar"
                    ControlSource ="Lebar"
                    StatusBarText ="(mm)"

                    LayoutCachedLeft =5940
                    LayoutCachedTop =900
                    LayoutCachedWidth =6540
                    LayoutCachedHeight =1200
                    BackThemeColorIndex =-1
                    ForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =4800
                            Top =900
                            Width =1080
                            Height =300
                            ForeColor =0
                            Name ="Label14"
                            Caption ="Lebar:"
                            LayoutCachedLeft =4800
                            LayoutCachedTop =900
                            LayoutCachedWidth =5880
                            LayoutCachedHeight =1200
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5940
                    Top =1260
                    Width =600
                    Height =300
                    ColumnWidth =1370
                    ColumnOrder =14
                    TabIndex =7
                    BackColor =14396046
                    ForeColor =0
                    Name ="Tebal"
                    ControlSource ="Tebal"
                    StatusBarText ="(mm). Tebal (untuk keping yang tebal-tipis, ketebalan yang tertinggi- yang lebih"
                        " tepis di Tebal2)"

                    LayoutCachedLeft =5940
                    LayoutCachedTop =1260
                    LayoutCachedWidth =6540
                    LayoutCachedHeight =1560
                    BackThemeColorIndex =-1
                    ForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =4800
                            Top =1260
                            Width =1080
                            Height =300
                            ForeColor =0
                            Name ="Label15"
                            Caption ="Tebal:"
                            LayoutCachedLeft =4800
                            LayoutCachedTop =1260
                            LayoutCachedWidth =5880
                            LayoutCachedHeight =1560
                            ForeTint =100.0
                        End
                    End
                End
                Begin CheckBox
                    Locked = NotDefault
                    OverlapFlags =215
                    Left =6780
                    Top =540
                    ColumnOrder =15
                    TabIndex =8
                    Name ="Tebal/TipisFlag"
                    ControlSource ="Tebal_TipisFlag"
                    StatusBarText ="Untuk ketebalan yang tidak sama antara satu ujung dengan yang lain"
                    EventProcPrefix ="Tebal_TipisFlag"

                    LayoutCachedLeft =6780
                    LayoutCachedTop =540
                    LayoutCachedWidth =7040
                    LayoutCachedHeight =780
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =7140
                            Top =540
                            Width =1080
                            Height =300
                            ForeColor =0
                            Name ="Label16"
                            Caption ="Tebal Tipis"
                            LayoutCachedLeft =7140
                            LayoutCachedTop =540
                            LayoutCachedWidth =8220
                            LayoutCachedHeight =840
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5940
                    Top =1620
                    Width =600
                    Height =300
                    ColumnWidth =1180
                    ColumnOrder =16
                    TabIndex =9
                    BackColor =14396046
                    ForeColor =0
                    Name ="Tebal2"
                    ControlSource ="DiameterDrum"
                    StatusBarText ="(mm) Apabila ada yang bagian tebal tipis.  Ketebalan tipis untuk yang tipis"

                    LayoutCachedLeft =5940
                    LayoutCachedTop =1620
                    LayoutCachedWidth =6540
                    LayoutCachedHeight =1920
                    BackThemeColorIndex =-1
                    ForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =4800
                            Top =1620
                            Width =1080
                            Height =300
                            ForeColor =0
                            Name ="Label17"
                            Caption ="Dia. Drum:"
                            LayoutCachedLeft =4800
                            LayoutCachedTop =1620
                            LayoutCachedWidth =5880
                            LayoutCachedHeight =1920
                            ForeTint =100.0
                        End
                    End
                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =255
                    Left =2640
                    Top =2160
                    Width =8040
                    Height =3000
                    BorderColor =12349952
                    Name ="Box19"
                    LayoutCachedLeft =2640
                    LayoutCachedTop =2160
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =5160
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin Image
                    PictureType =1
                    Left =5880
                    Top =2220
                    Width =4741
                    Height =2821
                    Name ="PolaLubangImage"
                    ControlSource ="LokasiImage"

                    LayoutCachedLeft =5880
                    LayoutCachedTop =2220
                    LayoutCachedWidth =10621
                    LayoutCachedHeight =5041
                    TabIndex =26
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4200
                    Top =2940
                    Height =300
                    ColumnOrder =0
                    TabIndex =10
                    Name ="PolaLubang"
                    ControlSource ="PolaLubang"

                    LayoutCachedLeft =4200
                    LayoutCachedTop =2940
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =3240
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2700
                            Top =2940
                            Width =1410
                            Height =300
                            Name ="Label25"
                            Caption ="PolaLubang"
                            LayoutCachedLeft =2700
                            LayoutCachedTop =2940
                            LayoutCachedWidth =4110
                            LayoutCachedHeight =3240
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7920
                    Top =900
                    Width =600
                    Height =300
                    ColumnWidth =1180
                    TabIndex =11
                    ForeColor =0
                    Name ="Text30"
                    ControlSource ="Tebal2"
                    StatusBarText ="(mm) Apabila ada yang bagian tebal tipis.  Ketebalan tipis untuk yang tipis"

                    LayoutCachedLeft =7920
                    LayoutCachedTop =900
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =1200
                    BackThemeColorIndex =4
                    BackTint =60.0
                    ForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =6780
                            Top =900
                            Width =1080
                            Height =300
                            ForeColor =0
                            Name ="Label31"
                            Caption ="Tebal2"
                            LayoutCachedLeft =6780
                            LayoutCachedTop =900
                            LayoutCachedWidth =7860
                            LayoutCachedHeight =1200
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7920
                    Top =1260
                    Width =600
                    Height =300
                    TabIndex =12
                    ForeColor =0
                    Name ="DiameterBrakeShoe"
                    ControlSource ="DiameterBrakeShoe"

                    LayoutCachedLeft =7920
                    LayoutCachedTop =1260
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =1560
                    BackThemeColorIndex =4
                    BackTint =60.0
                    ForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =6780
                            Top =1260
                            Width =1080
                            Height =300
                            ForeColor =0
                            Name ="Label32"
                            Caption ="Dia. Shoe:"
                            LayoutCachedLeft =6780
                            LayoutCachedTop =1260
                            LayoutCachedWidth =7860
                            LayoutCachedHeight =1560
                            ForeTint =100.0
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    Left =5640
                    Top =120
                    Width =2280
                    Height =300
                    FontSize =12
                    FontWeight =700
                    ForeColor =0
                    Name ="Label33"
                    Caption =" Ukuran Blanko (mm)"
                    LayoutCachedLeft =5640
                    LayoutCachedTop =120
                    LayoutCachedWidth =7920
                    LayoutCachedHeight =420
                    BackThemeColorIndex =4
                    BackTint =60.0
                    ForeTint =100.0
                End
                Begin TextBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =4200
                    Top =3660
                    Height =300
                    TabIndex =13
                    Name ="KodeRivet"
                    ControlSource ="KodeRivet"

                    LayoutCachedLeft =4200
                    LayoutCachedTop =3660
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =3960
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2700
                            Top =3660
                            Width =1410
                            Height =300
                            Name ="Label34"
                            Caption ="KodeRivet"
                            LayoutCachedLeft =2700
                            LayoutCachedTop =3660
                            LayoutCachedWidth =4110
                            LayoutCachedHeight =3960
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =4200
                    Top =4020
                    Height =300
                    ColumnWidth =3105
                    TabIndex =14
                    Name ="Ukuran(mm)"
                    ControlSource ="Ukuran"
                    EventProcPrefix ="Ukuran_mm_"

                    LayoutCachedLeft =4200
                    LayoutCachedTop =4020
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =4320
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2700
                            Top =4020
                            Width =1410
                            Height =300
                            Name ="Label35"
                            Caption ="Ukuran"
                            LayoutCachedLeft =2700
                            LayoutCachedTop =4020
                            LayoutCachedWidth =4110
                            LayoutCachedHeight =4320
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =4200
                    Top =4380
                    Height =300
                    TabIndex =15
                    Name ="Text36"
                    ControlSource ="PolaLubang"

                    LayoutCachedLeft =4200
                    LayoutCachedTop =4380
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =4680
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2700
                            Top =4380
                            Width =1410
                            Height =300
                            Name ="Label37"
                            Caption ="PolaLubang"
                            LayoutCachedLeft =2700
                            LayoutCachedTop =4380
                            LayoutCachedWidth =4110
                            LayoutCachedHeight =4680
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =4200
                    Top =3300
                    Height =300
                    TabIndex =16
                    Name ="RivetID"
                    ControlSource ="RivetID"

                    LayoutCachedLeft =4200
                    LayoutCachedTop =3300
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =3600
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2700
                            Top =3300
                            Width =1410
                            Height =300
                            Name ="Label38"
                            Caption ="RivetID"
                            LayoutCachedLeft =2700
                            LayoutCachedTop =3300
                            LayoutCachedWidth =4110
                            LayoutCachedHeight =3600
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =480
                    Top =2370
                    TabIndex =17
                    Name ="Indikator/Lis"
                    ControlSource ="Indikator-Lis"
                    StatusBarText ="wear indikator untuk lis"
                    EventProcPrefix ="Indikator_Lis"

                    LayoutCachedLeft =480
                    LayoutCachedTop =2370
                    LayoutCachedWidth =740
                    LayoutCachedHeight =2610
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =710
                            Top =2340
                            Width =1200
                            Height =300
                            Name ="Label39"
                            Caption ="Indikator-Lis"
                            LayoutCachedLeft =710
                            LayoutCachedTop =2340
                            LayoutCachedWidth =1910
                            LayoutCachedHeight =2640
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =480
                    Top =2730
                    TabIndex =18
                    Name ="Coak"
                    ControlSource ="Coak"
                    StatusBarText ="Pencos (setengah lingkaran untuk wear indikator)"

                    LayoutCachedLeft =480
                    LayoutCachedTop =2730
                    LayoutCachedWidth =740
                    LayoutCachedHeight =2970
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =710
                            Top =2700
                            Width =500
                            Height =300
                            Name ="Label40"
                            Caption ="Coak"
                            LayoutCachedLeft =710
                            LayoutCachedTop =2700
                            LayoutCachedWidth =1210
                            LayoutCachedHeight =3000
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =480
                    Top =3090
                    TabIndex =19
                    Name ="Koak"
                    ControlSource ="Koak"
                    StatusBarText ="koak seperti k-492/k-492 koak"

                    LayoutCachedLeft =480
                    LayoutCachedTop =3090
                    LayoutCachedWidth =740
                    LayoutCachedHeight =3330
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =710
                            Top =3060
                            Width =490
                            Height =300
                            Name ="Label41"
                            Caption ="Koak"
                            LayoutCachedLeft =710
                            LayoutCachedTop =3060
                            LayoutCachedWidth =1200
                            LayoutCachedHeight =3360
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1560
                    Top =5280
                    Width =4320
                    Height =960
                    TabIndex =20
                    Name ="Catatan"
                    ControlSource ="Catatan"

                    LayoutCachedLeft =1560
                    LayoutCachedTop =5280
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =6240
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =5280
                            Width =760
                            Height =300
                            Name ="Label42"
                            Caption ="Catatan"
                            LayoutCachedLeft =60
                            LayoutCachedTop =5280
                            LayoutCachedWidth =820
                            LayoutCachedHeight =5580
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4200
                    Top =2520
                    Width =420
                    Height =300
                    TabIndex =21
                    Name ="JumlahLubang"
                    ControlSource ="JumlahLubang"

                    LayoutCachedLeft =4200
                    LayoutCachedTop =2520
                    LayoutCachedWidth =4620
                    LayoutCachedHeight =2820
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2700
                            Top =2520
                            Width =1410
                            Height =300
                            Name ="Label43"
                            Caption ="Jumlah Lubang"
                            LayoutCachedLeft =2700
                            LayoutCachedTop =2520
                            LayoutCachedWidth =4110
                            LayoutCachedHeight =2820
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =1440
                    Left =1560
                    Top =780
                    Width =2820
                    Height =300
                    TabIndex =22
                    BoundColumn =2
                    Name ="DeskripsiBlanko"
                    ControlSource ="DeskripsiBlankoKeping"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT T_BlankoKeping.BlankoKepingAutoID, T_BlankoKeping.BlankoKepingID, T_Blank"
                        "oKeping.DeskripsiBlankoKeping FROM T_BlankoKeping; "
                    ColumnWidths ="0;0;1440"
                    AllowValueListEdits =0

                    LayoutCachedLeft =1560
                    LayoutCachedTop =780
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =1080
                    ForeThemeColorIndex =0
                    ForeShade =100.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9060
                    Top =60
                    Width =1620
                    Height =420
                    TabIndex =23
                    Name ="EditProfilBtn"
                    Caption ="Edit Profil"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =9060
                    LayoutCachedTop =60
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =480
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9060
                    Top =600
                    Width =1620
                    Height =420
                    TabIndex =24
                    Name ="SaveProfilBtn"
                    Caption ="Save Profil"

                    LayoutCachedLeft =9060
                    LayoutCachedTop =600
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =1020
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9060
                    Top =1140
                    Width =1620
                    Height =420
                    TabIndex =25
                    Name ="NewProfilBtn"
                    Caption ="Tambah Profil"

                    LayoutCachedLeft =9060
                    LayoutCachedTop =1140
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =1560
                End
            End
        End
        Begin PageFooter
            DisplayWhen =1
            Height =0
            BackColor =5676533
            Name ="PageFooterSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
        End
    End
End
CodeBehindForm
' See "F_ProfilKepingDetail_v1.cls"
