Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10740
    DatasheetFontHeight =11
    ItemSuffix =44
    Left =1370
    Top =6640
    Right =12390
    Bottom =16420
    RecSrcDt = Begin
        0x4bbbd2c7bd43e640
    End
    RecordSource ="Qry_ProfilKepingDetail_Primary"
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
        Begin PageHeader
            DisplayWhen =1
            Height =420
            BackColor =5676533
            Name ="PageHeaderSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontFamily =0
                    IMESentenceMode =3
                    Left =1560
                    Top =60
                    Width =2820
                    Height =300
                    FontSize =12
                    ForeColor =0
                    Name ="KepingAutoID"
                    ControlSource ="KepingID"
                    FontName ="Helvetica"

                    LayoutCachedLeft =1560
                    LayoutCachedTop =60
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =360
                    ThemeFontIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =9780
            BackColor =5676533
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Left =4500
                    Top =60
                    Width =6180
                    Height =1980
                    Name ="Box18"
                    LayoutCachedLeft =4500
                    LayoutCachedTop =60
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =2040
                    BackThemeColorIndex =4
                    BackTint =60.0
                    BorderThemeColorIndex =8
                    BorderTint =40.0
                    BorderShade =100.0
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1560
                    Top =420
                    Width =2820
                    Height =300
                    ColumnOrder =3
                    ForeColor =0
                    Name ="KepingID"
                    ControlSource ="KepingID"

                    LayoutCachedLeft =1560
                    LayoutCachedTop =420
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =720
                    ForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =420
                            Width =860
                            Height =300
                            ForeColor =0
                            Name ="Label2"
                            Caption ="KepingID"
                            LayoutCachedLeft =60
                            LayoutCachedTop =420
                            LayoutCachedWidth =920
                            LayoutCachedHeight =720
                            ForeTint =100.0
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =1560
                    Top =780
                    Width =2820
                    Height =300
                    ColumnWidth =1425
                    ColumnOrder =4
                    TabIndex =1
                    Name ="BlankoKepingID"
                    ControlSource ="BlankoKepingID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [T_BlankoKeping].[BlankoKepingID], [T_BlankoKeping].[DeskripsiBlankoKepin"
                        "g] FROM T_BlankoKeping; "
                    ColumnWidths ="0;1440"
                    AllowValueListEdits =0

                    LayoutCachedLeft =1560
                    LayoutCachedTop =780
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =1080
                    ForeThemeColorIndex =0
                    ForeShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =780
                            Width =1480
                            Height =300
                            ForeColor =0
                            Name ="Label3"
                            Caption ="BlankoKepingID"
                            LayoutCachedLeft =60
                            LayoutCachedTop =780
                            LayoutCachedWidth =1540
                            LayoutCachedHeight =1080
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
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
                            OverlapFlags =85
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
                    Top =4020
                    Width =1620
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
                    LayoutCachedTop =4020
                    LayoutCachedWidth =5820
                    LayoutCachedHeight =4320
                    ForeThemeColorIndex =0
                    ForeShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2700
                            Top =4020
                            Width =1300
                            Height =300
                            ForeColor =0
                            Name ="Label9"
                            Caption ="T_ProfilKeping.PolaLubangID"
                            LayoutCachedLeft =2700
                            LayoutCachedTop =4020
                            LayoutCachedWidth =4000
                            LayoutCachedHeight =4320
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =1
                    OldBorderStyle =0
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =5940
                    Top =540
                    Width =1020
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
                    LayoutCachedWidth =6960
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
                    IMESentenceMode =3
                    Left =5940
                    Top =900
                    Width =1020
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
                    LayoutCachedWidth =6960
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
                    IMESentenceMode =3
                    Left =5940
                    Top =1260
                    Width =1020
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
                    LayoutCachedWidth =6960
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
                    Left =7140
                    Top =540
                    ColumnOrder =15
                    TabIndex =8
                    Name ="Tebal/TipisFlag"
                    ControlSource ="Tebal_TipisFlag"
                    StatusBarText ="Untuk ketebalan yang tidak sama antara satu ujung dengan yang lain"
                    EventProcPrefix ="Tebal_TipisFlag"

                    LayoutCachedLeft =7140
                    LayoutCachedTop =540
                    LayoutCachedWidth =7400
                    LayoutCachedHeight =780
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =7560
                            Top =540
                            Width =2210
                            Height =300
                            ForeColor =0
                            Name ="Label16"
                            Caption ="Profil Keping Tebal Tipis"
                            LayoutCachedLeft =7560
                            LayoutCachedTop =540
                            LayoutCachedWidth =9770
                            LayoutCachedHeight =840
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =5940
                    Top =1620
                    Width =1020
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
                    LayoutCachedWidth =6960
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
                    Height =3240
                    BorderColor =12349952
                    Name ="Box19"
                    LayoutCachedLeft =2640
                    LayoutCachedTop =2160
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =5400
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8520
                    Top =6420
                    Width =1740
                    TabIndex =10
                    Name ="Opn_PolaLubang_Detail"
                    Caption ="Detail Pola Lubang"
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="F_PolaLubangDetail"
                            Argument ="0"
                            Argument =""
                            Argument ="=\"[PolaLubangID]=\"  &[PolaLubangID]"
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"Opn_PolaLubang_Detail\" xmlns=\"http://schemas.microsoft.com"
                                "/office/accessservices/2009/11/application\"><Statements><Action Name=\"OpenForm"
                                "\"><Argument Name=\"FormName\">F_"
                        End
                        Begin
                            Comment ="_AXL:PolaLubangDetail</Argument><Argument Name=\"WhereCondition\">=\"[PolaLubang"
                                "ID]=\"  &amp;[PolaLubangID]</Argument></Action></Statements></UserInterfaceMacro"
                                ">"
                        End
                    End

                    LayoutCachedLeft =8520
                    LayoutCachedTop =6420
                    LayoutCachedWidth =10260
                    LayoutCachedHeight =6780
                    Overlaps =1
                End
                Begin Image
                    PictureType =1
                    Left =5880
                    Top =2880
                    Width =4741
                    Height =2161
                    Name ="PolaLubangImage"
                    ControlSource ="LokasiImage"

                    LayoutCachedLeft =5880
                    LayoutCachedTop =2880
                    LayoutCachedWidth =10621
                    LayoutCachedHeight =5041
                    TabIndex =24
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9180
                    Top =5040
                    Height =300
                    ColumnOrder =0
                    TabIndex =11
                    Name ="PolaLubang"
                    ControlSource ="PolaLubang"

                    LayoutCachedLeft =9180
                    LayoutCachedTop =5040
                    LayoutCachedWidth =10620
                    LayoutCachedHeight =5340
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =7680
                            Top =5040
                            Width =1100
                            Height =300
                            Name ="Label25"
                            Caption ="PolaLubang"
                            LayoutCachedLeft =7680
                            LayoutCachedTop =5040
                            LayoutCachedWidth =8780
                            LayoutCachedHeight =5340
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8520
                    Top =6840
                    Width =2010
                    TabIndex =12
                    Name ="Tambah Pola Lubang"
                    Caption ="Tambah Pola Lubang"
                    OnClick ="[Event Procedure]"
                    EventProcPrefix ="Tambah_Pola_Lubang"

                    LayoutCachedLeft =8520
                    LayoutCachedTop =6840
                    LayoutCachedWidth =10530
                    LayoutCachedHeight =7200
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =8280
                    Top =900
                    Width =1020
                    Height =300
                    ColumnWidth =1180
                    TabIndex =13
                    ForeColor =0
                    Name ="Text30"
                    ControlSource ="Tebal2"
                    StatusBarText ="(mm) Apabila ada yang bagian tebal tipis.  Ketebalan tipis untuk yang tipis"

                    LayoutCachedLeft =8280
                    LayoutCachedTop =900
                    LayoutCachedWidth =9300
                    LayoutCachedHeight =1200
                    BackThemeColorIndex =4
                    BackTint =60.0
                    ForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =7140
                            Top =900
                            Width =1080
                            Height =300
                            ForeColor =0
                            Name ="Label31"
                            Caption ="Tebal2"
                            LayoutCachedLeft =7140
                            LayoutCachedTop =900
                            LayoutCachedWidth =8220
                            LayoutCachedHeight =1200
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =8280
                    Top =1260
                    Width =1020
                    Height =300
                    TabIndex =14
                    ForeColor =0
                    Name ="DiameterBrakeShoe"
                    ControlSource ="DiameterBrakeShoe"

                    LayoutCachedLeft =8280
                    LayoutCachedTop =1260
                    LayoutCachedWidth =9300
                    LayoutCachedHeight =1560
                    BackThemeColorIndex =4
                    BackTint =60.0
                    ForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =7140
                            Top =1260
                            Width =1080
                            Height =300
                            ForeColor =0
                            Name ="Label32"
                            Caption ="Dia. Shoe:"
                            LayoutCachedLeft =7140
                            LayoutCachedTop =1260
                            LayoutCachedWidth =8220
                            LayoutCachedHeight =1560
                            ForeTint =100.0
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    Left =6600
                    Top =120
                    Width =2280
                    Height =300
                    FontSize =12
                    ForeColor =0
                    Name ="Label33"
                    Caption =" Ukuran Dalam mm:"
                    LayoutCachedLeft =6600
                    LayoutCachedTop =120
                    LayoutCachedWidth =8880
                    LayoutCachedHeight =420
                    BackThemeColorIndex =4
                    BackTint =60.0
                    ForeTint =100.0
                End
                Begin TextBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =4200
                    Top =2940
                    Height =300
                    TabIndex =15
                    Name ="KodeRivet"
                    ControlSource ="KodeRivet"

                    LayoutCachedLeft =4200
                    LayoutCachedTop =2940
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =3240
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2700
                            Top =2940
                            Width =960
                            Height =300
                            Name ="Label34"
                            Caption ="KodeRivet"
                            LayoutCachedLeft =2700
                            LayoutCachedTop =2940
                            LayoutCachedWidth =3660
                            LayoutCachedHeight =3240
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =4200
                    Top =3300
                    Height =300
                    ColumnWidth =3105
                    TabIndex =16
                    Name ="Ukuran(mm)"
                    ControlSource ="Ukuran"
                    EventProcPrefix ="Ukuran_mm_"

                    LayoutCachedLeft =4200
                    LayoutCachedTop =3300
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =3600
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2700
                            Top =3300
                            Width =1220
                            Height =300
                            Name ="Label35"
                            Caption ="Ukuran"
                            LayoutCachedLeft =2700
                            LayoutCachedTop =3300
                            LayoutCachedWidth =3920
                            LayoutCachedHeight =3600
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =4200
                    Top =3660
                    Height =300
                    TabIndex =17
                    Name ="Text36"
                    ControlSource ="PolaLubang"

                    LayoutCachedLeft =4200
                    LayoutCachedTop =3660
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =3960
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2700
                            Top =3660
                            Width =1100
                            Height =300
                            Name ="Label37"
                            Caption ="PolaLubang"
                            LayoutCachedLeft =2700
                            LayoutCachedTop =3660
                            LayoutCachedWidth =3800
                            LayoutCachedHeight =3960
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =4200
                    Top =2580
                    Height =300
                    TabIndex =18
                    Name ="RivetID"
                    ControlSource ="RivetID"

                    LayoutCachedLeft =4200
                    LayoutCachedTop =2580
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =2880
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2700
                            Top =2580
                            Width =700
                            Height =300
                            Name ="Label38"
                            Caption ="RivetID"
                            LayoutCachedLeft =2700
                            LayoutCachedTop =2580
                            LayoutCachedWidth =3400
                            LayoutCachedHeight =2880
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =480
                    Top =2370
                    TabIndex =19
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
                    TabIndex =20
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
                    TabIndex =21
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
                    Top =5520
                    Width =4320
                    Height =960
                    TabIndex =22
                    Name ="Catatan"
                    ControlSource ="Catatan"

                    LayoutCachedLeft =1560
                    LayoutCachedTop =5520
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =6480
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =5520
                            Width =760
                            Height =300
                            Name ="Label42"
                            Caption ="Catatan"
                            LayoutCachedLeft =60
                            LayoutCachedTop =5520
                            LayoutCachedWidth =820
                            LayoutCachedHeight =5820
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =9180
                    Top =2520
                    Height =300
                    TabIndex =23
                    Name ="JumlahLubang"
                    ControlSource ="JumlahLubang"

                    LayoutCachedLeft =9180
                    LayoutCachedTop =2520
                    LayoutCachedWidth =10620
                    LayoutCachedHeight =2820
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =7680
                            Top =2520
                            Width =1360
                            Height =300
                            Name ="Label43"
                            Caption ="JumlahLubang"
                            LayoutCachedLeft =7680
                            LayoutCachedTop =2520
                            LayoutCachedWidth =9040
                            LayoutCachedHeight =2820
                        End
                    End
                End
            End
        End
        Begin PageFooter
            DisplayWhen =1
            Height =360
            BackColor =5676533
            Name ="PageFooterSection"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
        End
    End
End
CodeBehindForm
' See "F_ProfilKepingDetail_bak_241103.cls"
