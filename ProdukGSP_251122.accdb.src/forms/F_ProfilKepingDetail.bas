Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10980
    DatasheetFontHeight =11
    ItemSuffix =77
    Left =6950
    Top =980
    Right =13820
    Bottom =7460
    Filter =" [KepingID] = 'KI-M0002'"
    RecSrcDt = Begin
        0x241713ec0e73e640
    End
    OnDirty ="[Event Procedure]"
    RecordSource ="Qry_ProfilKepingDetail_Primary"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
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
        Begin FormHeader
            Height =960
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =6780
                    Top =540
                    Width =1320
                    Name ="cmdNewProfilKeping"
                    Caption ="Baru "
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =6780
                    LayoutCachedTop =540
                    LayoutCachedWidth =8100
                    LayoutCachedHeight =900
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9540
                    Top =540
                    Width =1320
                    TabIndex =1
                    Name ="cmdSaveProfil"
                    Caption ="Simpan "
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =9540
                    LayoutCachedTop =540
                    LayoutCachedWidth =10860
                    LayoutCachedHeight =900
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8160
                    Top =540
                    Width =1320
                    TabIndex =2
                    Name ="cmdEditProfil"
                    Caption ="Edit "
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =8160
                    LayoutCachedTop =540
                    LayoutCachedWidth =9480
                    LayoutCachedHeight =900
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1620
                    Top =120
                    Height =360
                    FontSize =14
                    FontWeight =700
                    TabIndex =3
                    ForeColor =10040879
                    Name ="txtFormHeaderTitle"
                    ControlSource ="=[txtKepingID]"

                    LayoutCachedLeft =1620
                    LayoutCachedTop =120
                    LayoutCachedWidth =3060
                    LayoutCachedHeight =480
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =120
                            Width =1500
                            Height =360
                            FontSize =14
                            FontWeight =700
                            ForeColor =10040879
                            Name ="lblFormHeaderLabel"
                            Caption ="Profil Keping"
                            LayoutCachedLeft =60
                            LayoutCachedTop =120
                            LayoutCachedWidth =1560
                            LayoutCachedHeight =480
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6120
                    Top =120
                    Width =3360
                    Height =360
                    TabIndex =4
                    Name ="Combo82"
                    RowSourceType ="Table/Query"

                    LayoutCachedLeft =6120
                    LayoutCachedTop =120
                    LayoutCachedWidth =9480
                    LayoutCachedHeight =480
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =5520
                            Top =120
                            Width =540
                            Height =360
                            Name ="Label83"
                            Caption ="Cari"
                            LayoutCachedLeft =5520
                            LayoutCachedTop =120
                            LayoutCachedWidth =6060
                            LayoutCachedHeight =480
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9540
                    Top =120
                    Width =1320
                    TabIndex =5
                    Name ="cmdKembaliKeDaftarBaseProduk"
                    Caption ="Daftar PK"

                    LayoutCachedLeft =9540
                    LayoutCachedTop =120
                    LayoutCachedWidth =10860
                    LayoutCachedHeight =480
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =5400
                    Top =540
                    Width =1320
                    TabIndex =6
                    Name ="cmdCancel"
                    Caption ="Cancel"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =5400
                    LayoutCachedTop =540
                    LayoutCachedWidth =6720
                    LayoutCachedHeight =900
                    Overlaps =1
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =85
                    TextAlign =2
                    Top =540
                    Width =5340
                    Height =360
                    FontSize =16
                    FontWeight =700
                    BackColor =5676533
                    BorderColor =5676533
                    ForeColor =255
                    Name ="lblWarning"
                    Caption =" WARNING"
                    LayoutCachedTop =540
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =900
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =9360
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
                    Width =6420
                    Height =1980
                    Name ="Box18"
                    LayoutCachedLeft =4500
                    LayoutCachedTop =60
                    LayoutCachedWidth =10920
                    LayoutCachedHeight =2040
                    BackThemeColorIndex =4
                    BackTint =60.0
                    BorderThemeColorIndex =4
                    BorderTint =60.0
                    BorderShade =100.0
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =95
                    Top =360
                    Width =4500
                    Height =780
                    Name ="Box49"
                    LayoutCachedTop =360
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =1140
                    BackThemeColorIndex =4
                    BackTint =60.0
                    BorderThemeColorIndex =4
                    BorderTint =60.0
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
                    Name ="txtKepingID"
                    ControlSource ="KepingID"

                    LayoutCachedLeft =1560
                    LayoutCachedTop =60
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =360
                    ForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =95
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
                    Name ="cboBlankoKepingID"
                    ControlSource ="BlankoKepingAutoID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT T_BlankoKeping.BlankoKepingAutoID, T_BlankoKeping.BlankoKepingID, T_Blank"
                        "oKeping.DeskripsiBlankoKeping FROM T_BlankoKeping; "
                    ColumnWidths ="0;1440;0"
                    OnGotFocus ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =1560
                    LayoutCachedTop =420
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =720
                    ForeThemeColorIndex =0
                    ForeShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =1
                            Left =60
                            Top =360
                            Width =1480
                            Height =780
                            TopMargin =252
                            ForeColor =0
                            Name ="BlankoKepingLbl"
                            Caption ="BlankoKepingID"
                            LayoutCachedLeft =60
                            LayoutCachedTop =360
                            LayoutCachedWidth =1540
                            LayoutCachedHeight =1140
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
                    Name ="txtDeskripsiKeping"
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
                            Width =1190
                            Height =570
                            ForeColor =0
                            Name ="lblDeskripsiProfilKeping"
                            Caption ="Deskripsi\015\012Profil Keping"
                            LayoutCachedLeft =60
                            LayoutCachedTop =1140
                            LayoutCachedWidth =1250
                            LayoutCachedHeight =1710
                            ForeTint =100.0
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =180
                    Top =4350
                    ColumnOrder =8
                    TabIndex =3
                    Name ="LubangFlag"
                    ControlSource ="LubangFlag"

                    LayoutCachedLeft =180
                    LayoutCachedTop =4350
                    LayoutCachedWidth =440
                    LayoutCachedHeight =4590
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =440
                            Top =4320
                            Width =2140
                            Height =300
                            ForeColor =0
                            Name ="Label7"
                            Caption ="Keping Berlobang Rivet"
                            LayoutCachedLeft =440
                            LayoutCachedTop =4320
                            LayoutCachedWidth =2580
                            LayoutCachedHeight =4620
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
                    ControlSource ="TebalTipisFlag"
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
                            Caption ="Tebal/Tipis"
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
                    OverlapFlags =95
                    Left =120
                    Top =4620
                    Width =10800
                    Height =3960
                    BorderColor =12349952
                    Name ="Box19"
                    LayoutCachedLeft =120
                    LayoutCachedTop =4620
                    LayoutCachedWidth =10920
                    LayoutCachedHeight =8580
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin Image
                    PictureType =1
                    Left =3360
                    Top =4740
                    Width =4741
                    Height =2821
                    Name ="PolaLubangImage"
                    ControlSource ="HPImageCode"

                    LayoutCachedLeft =3360
                    LayoutCachedTop =4740
                    LayoutCachedWidth =8101
                    LayoutCachedHeight =7561
                    TabIndex =30
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8340
                    Top =900
                    Width =600
                    Height =300
                    ColumnWidth =1180
                    TabIndex =10
                    ForeColor =0
                    Name ="Text30"
                    ControlSource ="Panjang2"
                    StatusBarText ="(mm) Apabila ada yang bagian tebal tipis.  Ketebalan tipis untuk yang tipis"

                    LayoutCachedLeft =8340
                    LayoutCachedTop =900
                    LayoutCachedWidth =8940
                    LayoutCachedHeight =1200
                    BackThemeColorIndex =4
                    BackTint =60.0
                    ForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =6780
                            Top =900
                            Width =1500
                            Height =300
                            ForeColor =0
                            Name ="Label31"
                            Caption ="Panjang Bawah:"
                            LayoutCachedLeft =6780
                            LayoutCachedTop =900
                            LayoutCachedWidth =8280
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
                    Left =8340
                    Top =1620
                    Width =600
                    Height =300
                    TabIndex =11
                    ForeColor =0
                    Name ="DiameterBrakeShoe"
                    ControlSource ="DiameterBrakeShoe"

                    LayoutCachedLeft =8340
                    LayoutCachedTop =1620
                    LayoutCachedWidth =8940
                    LayoutCachedHeight =1920
                    BackThemeColorIndex =4
                    BackTint =60.0
                    ForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =6780
                            Top =1620
                            Width =1500
                            Height =300
                            ForeColor =0
                            Name ="Label32"
                            Caption ="Dia. Sepatu:"
                            LayoutCachedLeft =6780
                            LayoutCachedTop =1620
                            LayoutCachedWidth =8280
                            LayoutCachedHeight =1920
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
                Begin CheckBox
                    OverlapFlags =85
                    Left =5820
                    Top =2250
                    TabIndex =14
                    Name ="chkIndikator"
                    ControlSource ="Indikator-Lis"
                    StatusBarText ="wear indikator untuk lis"

                    LayoutCachedLeft =5820
                    LayoutCachedTop =2250
                    LayoutCachedWidth =6080
                    LayoutCachedHeight =2490
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =6050
                            Top =2160
                            Width =850
                            Height =540
                            ForeColor =0
                            Name ="lblIndikator"
                            Caption ="Indikator (Lis)"
                            LayoutCachedLeft =6050
                            LayoutCachedTop =2160
                            LayoutCachedWidth =6900
                            LayoutCachedHeight =2700
                            ForeTint =100.0
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =5820
                    Top =2790
                    TabIndex =15
                    Name ="chkCoak"
                    ControlSource ="Coak"
                    StatusBarText ="Pencos (setengah lingkaran untuk wear indikator)"

                    LayoutCachedLeft =5820
                    LayoutCachedTop =2790
                    LayoutCachedWidth =6080
                    LayoutCachedHeight =3030
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =6050
                            Top =2760
                            Width =500
                            Height =300
                            ForeColor =0
                            Name ="lblCoak"
                            Caption ="Coak"
                            LayoutCachedLeft =6050
                            LayoutCachedTop =2760
                            LayoutCachedWidth =6550
                            LayoutCachedHeight =3060
                            ForeTint =100.0
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =5820
                    Top =3150
                    TabIndex =16
                    Name ="chkKoak"
                    ControlSource ="Koak"
                    StatusBarText ="koak seperti k-492/k-492 koak"

                    LayoutCachedLeft =5820
                    LayoutCachedTop =3150
                    LayoutCachedWidth =6080
                    LayoutCachedHeight =3390
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =6050
                            Top =3120
                            Width =490
                            Height =300
                            ForeColor =0
                            Name ="lblKoak"
                            Caption ="Koak"
                            LayoutCachedLeft =6050
                            LayoutCachedTop =3120
                            LayoutCachedWidth =6540
                            LayoutCachedHeight =3420
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1560
                    Top =2160
                    Width =4140
                    Height =2040
                    TabIndex =17
                    ForeColor =0
                    Name ="Catatan"
                    ControlSource ="Catatan"

                    LayoutCachedLeft =1560
                    LayoutCachedTop =2160
                    LayoutCachedWidth =5700
                    LayoutCachedHeight =4200
                    ForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =2160
                            Width =1380
                            Height =720
                            ForeColor =0
                            Name ="Label42"
                            Caption ="Catatan Profil Keping:"
                            LayoutCachedLeft =60
                            LayoutCachedTop =2160
                            LayoutCachedWidth =1440
                            LayoutCachedHeight =2880
                            BackThemeColorIndex =-1
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
                    Top =780
                    Width =2820
                    Height =300
                    TabIndex =19
                    Name ="cboDeskripsiBlanko"
                    ControlSource ="BlankoKepingAutoID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT T_BlankoKeping.BlankoKepingAutoID, T_BlankoKeping.BlankoKepingID, T_Blank"
                        "oKeping.DeskripsiBlankoKeping FROM T_BlankoKeping; "
                    ColumnWidths ="0;0;2880"
                    AllowValueListEdits =0

                    LayoutCachedLeft =1560
                    LayoutCachedTop =780
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =1080
                    ForeThemeColorIndex =0
                    ForeShade =100.0
                End
                Begin Image
                    OldBorderStyle =1
                    Left =7200
                    Top =2520
                    Width =3720
                    Height =1740
                    BorderColor =0
                    Name ="imgGambarProfil"
                    ControlSource ="PKImageCode"

                    LayoutCachedLeft =7200
                    LayoutCachedTop =2520
                    LayoutCachedWidth =10920
                    LayoutCachedHeight =4260
                    TabIndex =31
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8340
                    Top =1260
                    Width =600
                    Height =300
                    TabIndex =20
                    ForeColor =0
                    Name ="Text54"
                    ControlSource ="Tebal2"
                    StatusBarText ="(mm) Apabila ada yang bagian tebal tipis.  Ketebalan tipis untuk yang tipis"

                    LayoutCachedLeft =8340
                    LayoutCachedTop =1260
                    LayoutCachedWidth =8940
                    LayoutCachedHeight =1560
                    BackThemeColorIndex =4
                    BackTint =60.0
                    ForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =6780
                            Top =1260
                            Width =1500
                            Height =300
                            ForeColor =0
                            Name ="Label55"
                            Caption ="Tebal2"
                            LayoutCachedLeft =6780
                            LayoutCachedTop =1260
                            LayoutCachedWidth =8280
                            LayoutCachedHeight =1560
                            ForeTint =100.0
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =9540
                    Top =120
                    Width =1320
                    TabIndex =21
                    Name ="cmnOpenBlankoKeping"
                    Caption ="Detail Blanko"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =9540
                    LayoutCachedTop =120
                    LayoutCachedWidth =10860
                    LayoutCachedHeight =480
                    Overlaps =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =215
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1680
                    Top =5100
                    Width =780
                    Height =300
                    TabIndex =22
                    Name ="T_ProfilKeping.PolaLubangID"
                    ControlSource ="T_ProfilKeping.PolaLubangID"
                    EventProcPrefix ="T_ProfilKeping_PolaLubangID"

                    LayoutCachedLeft =1680
                    LayoutCachedTop =5100
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =5400
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =9360
                    Top =5160
                    Width =1320
                    TabIndex =24
                    Name ="Command65"
                    Caption ="Detail Pola"
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="F_DaftarBlankoKeping"
                            Argument ="0"
                            Argument =""
                            Argument ="=\"[BlankoKepingAutoID]=\"  &[cboBlankoKepingID]"
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"Command65\" xmlns=\"http://schemas.microsoft.com/office/acce"
                                "ssservices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argument"
                                " Name=\"FormName\">F_DaftarBlanko"
                        End
                        Begin
                            Comment ="_AXL:Keping</Argument><Argument Name=\"WhereCondition\">=\"[BlankoKepingAutoID]="
                                "\"  &amp;[cboBlankoKepingID]</Argument></Action></Statements></UserInterfaceMacr"
                                "o>"
                        End
                    End

                    LayoutCachedLeft =9360
                    LayoutCachedTop =5160
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =5520
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9600
                    Top =2100
                    Width =1320
                    TabIndex =25
                    Name ="cmdKelolaFoto"
                    Caption ="Kelola Foto"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =9600
                    LayoutCachedTop =2100
                    LayoutCachedWidth =10920
                    LayoutCachedHeight =2460
                End
                Begin ComboBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =1680
                    Top =6180
                    Height =300
                    TabIndex =12
                    Name ="cboRivetID"
                    ControlSource ="RivetID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT T_UkuranRivet.RivetID, T_UkuranRivet.KodeBarang, T_UkuranRivet.Ukuran FRO"
                        "M T_UkuranRivet; "
                    ColumnWidths ="0;720;1440"
                    OnGotFocus ="[Event Procedure]"

                    LayoutCachedLeft =1680
                    LayoutCachedTop =6180
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =6480
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    ForeShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =120
                            Top =6180
                            Width =1410
                            Height =300
                            ForeColor =0
                            Name ="Label38"
                            Caption ="RivetID"
                            LayoutCachedLeft =120
                            LayoutCachedTop =6180
                            LayoutCachedWidth =1530
                            LayoutCachedHeight =6480
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1680
                    Top =5820
                    Height =300
                    TabIndex =23
                    ForeColor =0
                    Name ="cboCatatanPolaLubangID"
                    ControlSource ="DetailPolaLubang"

                    LayoutCachedLeft =1680
                    LayoutCachedTop =5820
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =6120
                    ForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =120
                            Top =5820
                            Width =1410
                            Height =300
                            ForeColor =0
                            Name ="Label69"
                            Caption ="Catatan P. Lbg"
                            LayoutCachedLeft =120
                            LayoutCachedTop =5820
                            LayoutCachedWidth =1530
                            LayoutCachedHeight =6120
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1680
                    Top =5460
                    Width =420
                    Height =300
                    TabIndex =18
                    ForeColor =0
                    Name ="JumlahLubang"
                    ControlSource ="JumlahLubang"

                    LayoutCachedLeft =1680
                    LayoutCachedTop =5460
                    LayoutCachedWidth =2100
                    LayoutCachedHeight =5760
                    ForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =120
                            Top =5460
                            Width =1410
                            Height =300
                            ForeColor =0
                            Name ="Label43"
                            Caption ="Jumlah Lubang"
                            LayoutCachedLeft =120
                            LayoutCachedTop =5460
                            LayoutCachedWidth =1530
                            LayoutCachedHeight =5760
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =215
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1680
                    Top =6540
                    Height =300
                    ColumnWidth =3105
                    TabIndex =13
                    Name ="Ukuran(mm)"
                    ControlSource ="Ukuran"
                    EventProcPrefix ="Ukuran_mm_"

                    LayoutCachedLeft =1680
                    LayoutCachedTop =6540
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =6840
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =120
                            Top =6540
                            Width =1410
                            Height =300
                            ForeColor =0
                            Name ="lbl_txtUkuranRivet"
                            Caption ="Ukuran Rivet"
                            LayoutCachedLeft =120
                            LayoutCachedTop =6540
                            LayoutCachedWidth =1530
                            LayoutCachedHeight =6840
                            ForeTint =100.0
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =2550
                    Left =1680
                    Top =4740
                    Height =300
                    ColumnOrder =10
                    TabIndex =4
                    Name ="cboPolaLubangID"
                    ControlSource ="PolaLubangID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT T_UkuranPolaLubang.PolaLubangID, T_UkuranPolaLubang.PolaLubang, T_UkuranP"
                        "olaLubang.DetailPolaLubang FROM T_UkuranPolaLubang ORDER BY T_UkuranPolaLubang.P"
                        "olaLubang; "
                    ColumnWidths ="0;650;1900"
                    OnGotFocus ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =1680
                    LayoutCachedTop =4740
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =5040
                    ForeThemeColorIndex =0
                    ForeShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =120
                            Top =4740
                            Width =1410
                            Height =660
                            TopMargin =187
                            ForeColor =0
                            Name ="lblPolaLubangID"
                            Caption ="PolaLubangID"
                            LayoutCachedLeft =120
                            LayoutCachedTop =4740
                            LayoutCachedWidth =1530
                            LayoutCachedHeight =5400
                            ForeTint =100.0
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1680
                    Top =6900
                    Height =300
                    TabIndex =26
                    Name ="cboUkuranLubangID"
                    ControlSource ="UkuranLubangID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT T_UkuranLubang.UkuranLubangID, T_UkuranLubang.UkuranLubang FROM T_UkuranL"
                        "ubang; "
                    ColumnWidths ="0;2880"
                    OnGotFocus ="[Event Procedure]"

                    LayoutCachedLeft =1680
                    LayoutCachedTop =6900
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =7200
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    ForeShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =120
                            Top =6900
                            Width =1410
                            Height =300
                            ForeColor =0
                            Name ="lbl_cboUkuranLubangID"
                            Caption ="Ukuran Lubang"
                            LayoutCachedLeft =120
                            LayoutCachedTop =6900
                            LayoutCachedWidth =1530
                            LayoutCachedHeight =7200
                            ForeTint =100.0
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =9360
                    Top =4680
                    TabIndex =27
                    Name ="cmd_HapusPolaLubang"
                    Caption ="Hapus Info Lubang"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =9360
                    LayoutCachedTop =4680
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =5040
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =9540
                    Top =600
                    Width =1320
                    TabIndex =28
                    Name ="cmdRefreshcbobox"
                    Caption ="cmd"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =9540
                    LayoutCachedTop =600
                    LayoutCachedWidth =10860
                    LayoutCachedHeight =960
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7560
                    Top =2160
                    Width =1080
                    Height =300
                    TabIndex =29
                    Name ="Command76"
                    Caption ="Debug"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =7560
                    LayoutCachedTop =2160
                    LayoutCachedWidth =8640
                    LayoutCachedHeight =2460
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =5676533
            Name ="FormFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
        End
    End
End
CodeBehindForm
' See "F_ProfilKepingDetail.cls"
