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
    DatasheetFontHeight =11
    ItemSuffix =79
    Left =9220
    Top =890
    Right =16420
    Bottom =7480
    Filter ="BlankoKepingAutoID = 11"
    RecSrcDt = Begin
        0x819543abc901e640
    End
    OnDirty ="[Event Procedure]"
    RecordSource ="T_BlankoKeping"
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =1020
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =300
                    Top =60
                    Width =2160
                    Height =460
                    FontSize =18
                    Name ="Auto_Header0"
                    Caption ="Blanko Keping"
                    FontName ="Calibri Light"
                    GroupTable =2
                    LayoutCachedLeft =300
                    LayoutCachedTop =60
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =520
                    LayoutGroup =1
                    ThemeFontIndex =0
                    BorderThemeColorIndex =2
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5700
                    Top =120
                    Width =1320
                    TabIndex =1
                    Name ="cmdKembaliKeDaftarBlanko"
                    Caption ="Daftar BK"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =5700
                    LayoutCachedTop =120
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =480
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4320
                    Top =600
                    Width =1320
                    TabIndex =4
                    Name ="cmdEditBlanko"
                    Caption ="Edit"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =4320
                    LayoutCachedTop =600
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =960
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2940
                    Top =600
                    Width =1320
                    TabIndex =3
                    Name ="cmdTambahBlanko"
                    Caption ="Baru"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =2940
                    LayoutCachedTop =600
                    LayoutCachedWidth =4260
                    LayoutCachedHeight =960
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5700
                    Top =600
                    Width =1320
                    TabIndex =5
                    Name ="cmdSimpanBlanko"
                    Caption ="Simpan"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =5700
                    LayoutCachedTop =600
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =960
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =3360
                    Top =120
                    Width =2280
                    Height =360
                    TabIndex =2
                    Name ="cboCariProduk"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT T_BlankoKeping.BlankoKepingAutoID, T_BlankoKeping.BlankoKepingID, T_Blank"
                        "oKeping.DeskripsiBlankoKeping FROM T_BlankoKeping; "
                    ColumnWidths ="0;864;288"
                    AfterUpdate ="[Event Procedure]"
                    TopMargin =43

                    LayoutCachedLeft =3360
                    LayoutCachedTop =120
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =480
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2880
                            Top =120
                            Width =420
                            Height =360
                            TopMargin =43
                            Name ="lbl_cboCariProduk"
                            Caption ="Cari"
                            LayoutCachedLeft =2880
                            LayoutCachedTop =120
                            LayoutCachedWidth =3300
                            LayoutCachedHeight =480
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =300
                    Top =600
                    Width =1140
                    Height =360
                    ColumnWidth =2400
                    ForeColor =8210719
                    Name ="txtBlankoKepingAutoID"
                    ControlSource ="BlankoKepingAutoID"

                    LayoutCachedLeft =300
                    LayoutCachedTop =600
                    LayoutCachedWidth =1440
                    LayoutCachedHeight =960
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =1560
                    Top =600
                    Width =1320
                    TabIndex =6
                    Name ="cmdCancel"
                    Caption ="Cancel"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =1560
                    LayoutCachedTop =600
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =960
                End
            End
        End
        Begin Section
            Height =5580
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1920
                    Top =120
                    Width =960
                    Height =300
                    ColumnWidth =590
                    Name ="txtBlankoKepingID"
                    ControlSource ="BlankoKepingID"

                    LayoutCachedLeft =1920
                    LayoutCachedTop =120
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =420
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =120
                            Width =1620
                            Height =300
                            Name ="Label46"
                            Caption ="BlankoKepingID"
                            LayoutCachedLeft =120
                            LayoutCachedTop =120
                            LayoutCachedWidth =1740
                            LayoutCachedHeight =420
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1920
                    Top =480
                    Width =5100
                    Height =300
                    ColumnWidth =3400
                    TabIndex =1
                    Name ="txtDeskripsiBlankoKeping"
                    ControlSource ="DeskripsiBlankoKeping"
                    StatusBarText ="Deskripsi Dasar Blanko (Sebelum lobang/finishing).  Lebih untuk mengetahui cetak"
                        " mould"

                    LayoutCachedLeft =1920
                    LayoutCachedTop =480
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =780
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =480
                            Width =1620
                            Height =300
                            Name ="Label47"
                            Caption ="Deskripsi"
                            LayoutCachedLeft =120
                            LayoutCachedTop =480
                            LayoutCachedWidth =1740
                            LayoutCachedHeight =780
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1920
                    Top =840
                    Width =960
                    Height =300
                    ColumnWidth =1020
                    TabIndex =2
                    Name ="txtPanjangA"
                    ControlSource ="Panjang"
                    Format ="General Number"
                    StatusBarText ="(mm)"

                    LayoutCachedLeft =1920
                    LayoutCachedTop =840
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =1140
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =840
                            Width =1620
                            Height =300
                            Name ="Label48"
                            Caption ="Panjang  (mm)"
                            LayoutCachedLeft =120
                            LayoutCachedTop =840
                            LayoutCachedWidth =1740
                            LayoutCachedHeight =1140
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1920
                    Top =1200
                    Width =960
                    Height =300
                    ColumnWidth =1020
                    TabIndex =3
                    Name ="txtLebar"
                    ControlSource ="Lebar"
                    StatusBarText ="(mm)"

                    LayoutCachedLeft =1920
                    LayoutCachedTop =1200
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =1500
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1200
                            Width =1620
                            Height =300
                            Name ="Label49"
                            Caption ="Lebar (mm)"
                            LayoutCachedLeft =120
                            LayoutCachedTop =1200
                            LayoutCachedWidth =1740
                            LayoutCachedHeight =1500
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =3540
                    Top =1230
                    TabIndex =5
                    Name ="chkTebalTipisFlag"
                    ControlSource ="Tebal_TipisFlag"
                    StatusBarText ="Untuk ketebalan yang tidak sama antara satu ujung dengan yang lain"
                    DefaultValue ="=[chkTebalTipisFlag]=True"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =3540
                    LayoutCachedTop =1230
                    LayoutCachedWidth =3800
                    LayoutCachedHeight =1470
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =3770
                            Top =1200
                            Width =1950
                            Height =300
                            Name ="Label50"
                            Caption ="Tebal/Tipis"
                            LayoutCachedLeft =3770
                            LayoutCachedTop =1200
                            LayoutCachedWidth =5720
                            LayoutCachedHeight =1500
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1920
                    Top =1560
                    Width =960
                    Height =300
                    ColumnWidth =2000
                    TabIndex =4
                    Name ="txtTebal1"
                    ControlSource ="Tebal"
                    StatusBarText ="(mm). Tebal (untuk keping yang tebal-tipis, ketebalan yang tertinggi- yang lebih"
                        " tepis di Tebal2)"

                    LayoutCachedLeft =1920
                    LayoutCachedTop =1560
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =1860
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1560
                            Width =1620
                            Height =300
                            Name ="Label51"
                            Caption ="Tebal (mm)"
                            LayoutCachedLeft =120
                            LayoutCachedTop =1560
                            LayoutCachedWidth =1740
                            LayoutCachedHeight =1860
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5580
                    Top =1560
                    Height =300
                    ColumnWidth =800
                    TabIndex =6
                    Name ="txtTebal2"
                    ControlSource ="Tebal2"
                    StatusBarText ="(mm) Apabila ada yang bagian tebal tipis.  Ketebalan tipis untuk yang tipis"

                    LayoutCachedLeft =5580
                    LayoutCachedTop =1560
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =1860
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3540
                            Top =1560
                            Width =1620
                            Height =300
                            Name ="Label52"
                            Caption ="Tebal2 (mm)"
                            LayoutCachedLeft =3540
                            LayoutCachedTop =1560
                            LayoutCachedWidth =5160
                            LayoutCachedHeight =1860
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1920
                    Top =1920
                    Width =960
                    Height =300
                    TabIndex =7
                    Name ="txtDiaDrum"
                    ControlSource ="DiameterDrum"
                    Format ="General Number"

                    LayoutCachedLeft =1920
                    LayoutCachedTop =1920
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =2220
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1920
                            Width =1620
                            Height =300
                            Name ="Label53"
                            Caption ="Ø Drum (mm)"
                            LayoutCachedLeft =120
                            LayoutCachedTop =1920
                            LayoutCachedWidth =1740
                            LayoutCachedHeight =2220
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1920
                    Top =2280
                    Width =5100
                    Height =1380
                    TabIndex =10
                    Name ="txtCatatanBlankoKeping"
                    ControlSource ="CatatanBlankoKeping"

                    LayoutCachedLeft =1920
                    LayoutCachedTop =2280
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =3660
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =2280
                            Width =1620
                            Height =300
                            Name ="Label54"
                            Caption ="Catatan"
                            LayoutCachedLeft =120
                            LayoutCachedTop =2280
                            LayoutCachedWidth =1740
                            LayoutCachedHeight =2580
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5580
                    Top =840
                    Height =300
                    TabIndex =8
                    Name ="txtPanjangB"
                    ControlSource ="Panjang2"
                    StatusBarText ="(mm) Untuk Panjang Bagian Bawah"

                    LayoutCachedLeft =5580
                    LayoutCachedTop =840
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =1140
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3540
                            Top =840
                            Width =1990
                            Height =300
                            Name ="Label56"
                            Caption ="Panjang Bawah (mm)"
                            LayoutCachedLeft =3540
                            LayoutCachedTop =840
                            LayoutCachedWidth =5530
                            LayoutCachedHeight =1140
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5580
                    Top =1920
                    Height =300
                    TabIndex =9
                    Name ="txtDiaBrakeShoe"
                    ControlSource ="DiameterBrakeShoe"

                    LayoutCachedLeft =5580
                    LayoutCachedTop =1920
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =2220
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3540
                            Top =1920
                            Width =1870
                            Height =300
                            Name ="Label57"
                            Caption ="Ø Sepatu (mm)"
                            LayoutCachedLeft =3540
                            LayoutCachedTop =1920
                            LayoutCachedWidth =5410
                            LayoutCachedHeight =2220
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =3480
                    Width =3540
                    Height =420
                    FontSize =16
                    FontWeight =700
                    BackColor =5676533
                    BorderColor =5676533
                    ForeColor =0
                    Name ="lblWarning"
                    Caption =" WARNING"
                    LayoutCachedLeft =3480
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =420
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
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
' See "F_BlankoKeping.cls"
