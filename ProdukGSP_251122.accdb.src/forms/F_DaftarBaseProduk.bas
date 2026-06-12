Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11520
    DatasheetFontHeight =11
    ItemSuffix =33
    Left =1050
    Top =70
    Right =12850
    Bottom =7320
    OrderBy ="BaseProdukID"
    RecSrcDt = Begin
        0x1a2389ff707de640
    End
    RecordSource ="Q_BaseProdukDetail"
    Caption ="Daftar Base Produk"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    OnKeyDown ="[Event Procedure]"
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
            Height =1800
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Rectangle
                    BackStyle =1
                    OverlapFlags =93
                    Left =7320
                    Top =60
                    Width =4080
                    Height =1260
                    BackColor =8421504
                    Name ="Box19"
                    LayoutCachedLeft =7320
                    LayoutCachedTop =60
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =1320
                    BackThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =180
                    Top =1380
                    Width =530
                    Height =300
                    Name ="Kode_Label"
                    Caption =" BPID"
                    OnClick ="[Event Procedure]"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =180
                    LayoutCachedTop =1380
                    LayoutCachedWidth =710
                    LayoutCachedHeight =1680
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =2160
                    Top =1380
                    Width =3060
                    Height =300
                    Name ="Deskripsi_Label"
                    Caption ="Deskripsi"
                    OnClick ="[Event Procedure]"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =2160
                    LayoutCachedTop =1380
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =1680
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =5280
                    Top =1380
                    Width =540
                    Height =300
                    Name ="AxlePos_Label"
                    Caption ="F/R/T"
                    OnClick ="[Event Procedure]"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =5280
                    LayoutCachedTop =1380
                    LayoutCachedWidth =5820
                    LayoutCachedHeight =1680
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =5880
                    Top =1380
                    Width =5460
                    Height =300
                    Name ="Catatan_Label"
                    Caption ="Catatan"
                    OnClick ="[Event Procedure]"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =5880
                    LayoutCachedTop =1380
                    LayoutCachedWidth =11340
                    LayoutCachedHeight =1680
                End
                Begin Label
                    OverlapFlags =85
                    Left =60
                    Top =60
                    Width =4376
                    Height =480
                    FontSize =20
                    Name ="Label16"
                    Caption ="Daftar Produk Dasar"
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =4436
                    LayoutCachedHeight =540
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =8400
                    Top =960
                    Width =1800
                    Height =310
                    ColumnWidth =3000
                    Name ="StatusBaseProdukFilt"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="0;1440"
                    StatusBarText ="(Lookup) List Status dari Sample, Prototype, Produksi"
                    OnGotFocus ="[Event Procedure]"
                    OnLostFocus ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =8400
                    LayoutCachedTop =960
                    LayoutCachedWidth =10200
                    LayoutCachedHeight =1270
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =1
                            Left =7380
                            Top =960
                            Width =960
                            Height =300
                            Name ="StatusBaseProduk_Lbl"
                            Caption ="Status"
                            Tag ="DetachedLabel"
                            GridlineStyleBottom =1
                            LayoutCachedLeft =7380
                            LayoutCachedTop =960
                            LayoutCachedWidth =8340
                            LayoutCachedHeight =1260
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =8400
                    Top =600
                    Width =1800
                    Height =310
                    ColumnWidth =3000
                    TabIndex =1
                    Name ="KatProdukFilt"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="0;1441"
                    StatusBarText ="(Lookup) dari T_KategoriProduk"
                    AfterUpdate ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"
                    OnLostFocus ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =8400
                    LayoutCachedTop =600
                    LayoutCachedWidth =10200
                    LayoutCachedHeight =910
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =1
                            Left =7380
                            Top =600
                            Width =960
                            Height =300
                            Name ="KatProdukFilt_Lbl"
                            Caption ="Kategori"
                            Tag ="DetachedLabel"
                            GridlineStyleBottom =1
                            LayoutCachedLeft =7380
                            LayoutCachedTop =600
                            LayoutCachedWidth =8340
                            LayoutCachedHeight =900
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5340
                    Top =480
                    Width =1860
                    TabIndex =2
                    Name ="TambahProdukDasar"
                    Caption ="Tambah Produk"
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =10
                        Begin
                            Action ="OnError"
                            Argument ="0"
                        End
                        Begin
                            Action ="GoToRecord"
                            Argument ="-1"
                            Argument =""
                            Argument ="5"
                        End
                        Begin
                            Condition ="[MacroError]<>0"
                            Action ="MsgBox"
                            Argument ="=[MacroError].[Description]"
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"TambahProdukDasar\" xmlns=\"http://schemas.microsoft.com/off"
                                "ice/accessservices/2009/11/application\"><Statements><Action Name=\"OnError\"/><"
                                "Action Name=\"GoToRecord\"><Argum"
                        End
                        Begin
                            Comment ="_AXL:ent Name=\"Record\">New</Argument></Action><ConditionalBlock><If><Condition"
                                ">[MacroError]&lt;&gt;0</Condition><Statements><Action Name=\"MessageBox\"><Argum"
                                "ent Name=\"Message\">=[MacroError].[Description]</Argument></Action></Statements"
                                "></If></Conditional"
                        End
                        Begin
                            Comment ="_AXL:Block></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =5340
                    LayoutCachedTop =480
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =840
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    Left =7440
                    Top =120
                    Width =1260
                    Height =360
                    FontSize =14
                    FontWeight =700
                    Name ="FILTER_lbl"
                    Caption ="Filter"
                    LayoutCachedLeft =7440
                    LayoutCachedTop =120
                    LayoutCachedWidth =8700
                    LayoutCachedHeight =480
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =9120
                    Top =120
                    Width =1860
                    TabIndex =3
                    Name ="HapusSemuaFilt"
                    Caption ="Hapus Semua"

                    LayoutCachedLeft =9120
                    LayoutCachedTop =120
                    LayoutCachedWidth =10980
                    LayoutCachedHeight =480
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =10260
                    Top =960
                    Width =1080
                    Height =300
                    TabIndex =4
                    Name ="HapusStatusProdFiltBtn"
                    Caption ="Hapus"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =10260
                    LayoutCachedTop =960
                    LayoutCachedWidth =11340
                    LayoutCachedHeight =1260
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =10260
                    Top =600
                    Width =1080
                    Height =300
                    TabIndex =5
                    Name ="HapusKatFiltBtn"
                    Caption ="Hapus"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =10260
                    LayoutCachedTop =600
                    LayoutCachedWidth =11340
                    LayoutCachedHeight =900
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =720
                    Top =1380
                    Width =1380
                    Height =300
                    Name ="Produk_Label"
                    Caption ="Produk"
                    OnClick ="[Event Procedure]"
                    LayoutCachedLeft =720
                    LayoutCachedTop =1380
                    LayoutCachedWidth =2100
                    LayoutCachedHeight =1680
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5340
                    Top =60
                    Width =1860
                    TabIndex =6
                    Name ="ExportExcelBtn"
                    Caption ="Export ke Excel"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =5340
                    LayoutCachedTop =60
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =420
                End
            End
        End
        Begin Section
            Height =420
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    FontUnderline = NotDefault
                    IsHyperlink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =180
                    Top =60
                    Width =480
                    Height =310
                    ColumnWidth =2240
                    Name ="SetID"
                    ControlSource ="BaseProdukID"
                    StatusBarText ="Copy dari SetAutoID"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =180
                    LayoutCachedTop =60
                    LayoutCachedWidth =660
                    LayoutCachedHeight =370
                    ForeThemeColorIndex =10
                    ForeTint =100.0
                    HyperlinkTarget ="form.F_BaseProdukDaftarUtama"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2160
                    Top =60
                    Width =3060
                    Height =310
                    ColumnWidth =3000
                    TabIndex =2
                    Name ="Deskripsi"
                    ControlSource ="Deskripsi"

                    LayoutCachedLeft =2160
                    LayoutCachedTop =60
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =370
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5880
                    Top =60
                    Width =5460
                    Height =300
                    ColumnWidth =3000
                    TabIndex =3
                    Name ="Catatan"
                    ControlSource ="Catatan"
                    StatusBarText ="Ini untuk memuat Catatan Panjang tentang Produk terkait"

                    LayoutCachedLeft =5880
                    LayoutCachedTop =60
                    LayoutCachedWidth =11340
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =720
                    Top =60
                    Width =1380
                    Height =300
                    TabIndex =4
                    Name ="Text27"
                    ControlSource ="BaseProdukName"
                    StatusBarText ="Copy dari SetAutoID - Now from SetID change to BaseProductName"

                    LayoutCachedLeft =720
                    LayoutCachedTop =60
                    LayoutCachedWidth =2100
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5280
                    Top =60
                    Width =540
                    Height =310
                    TabIndex =1
                    Name ="AxlePosID"
                    ControlSource ="AxlePos"
                    StatusBarText ="(Lookup) Jenis Produk: F/R,Handbrake, Trailer, Lempengan dari T_AxlePos"

                    LayoutCachedLeft =5280
                    LayoutCachedTop =60
                    LayoutCachedWidth =5820
                    LayoutCachedHeight =370
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
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
' See "F_DaftarBaseProduk.cls"
