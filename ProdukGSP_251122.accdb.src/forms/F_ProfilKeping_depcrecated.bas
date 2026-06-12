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
    Width =10080
    DatasheetFontHeight =11
    ItemSuffix =78
    Left =-3630
    Right =6710
    Bottom =11880
    RecSrcDt = Begin
        0x8133f7cb1877e640
    End
    RecordSource ="SELECT T_ProfilKeping.*, GetPKImagePath ([T_ProfilKeping.KepingAutoID]) AS PKIma"
        "geCode FROM T_ProfilKeping; "
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Edge
            OldBorderStyle =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =900
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =300
                    Top =60
                    Width =2580
                    Height =600
                    FontSize =22
                    ForeColor =0
                    Name ="Auto_Header0"
                    Caption ="Profil Keping"
                    FontName ="Calibri Light"
                    GroupTable =1
                    LayoutCachedLeft =300
                    LayoutCachedTop =60
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =660
                    LayoutGroup =1
                    ThemeFontIndex =0
                    BorderThemeColorIndex =2
                    BorderTint =100.0
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5520
                    Top =480
                    Width =1320
                    Name ="TambahProfilKpnBtn"
                    Caption ="Baru"

                    LayoutCachedLeft =5520
                    LayoutCachedTop =480
                    LayoutCachedWidth =6840
                    LayoutCachedHeight =840
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8280
                    Top =480
                    Width =1320
                    TabIndex =2
                    Name ="cmdEditProfilKpngBtn"
                    Caption ="Edit"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =8280
                    LayoutCachedTop =480
                    LayoutCachedWidth =9600
                    LayoutCachedHeight =840
                    PictureCaptionArrangement =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8280
                    Top =60
                    Width =1320
                    TabIndex =1
                    Name ="cmdKembaliKeDaftarBaseProduk"
                    Caption ="Daftar PK"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =8280
                    LayoutCachedTop =60
                    LayoutCachedWidth =9600
                    LayoutCachedHeight =420
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =4140
                    Top =480
                    Width =1320
                    TabIndex =3
                    Name ="cmdCancel"
                    Caption ="Cancel"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =4140
                    LayoutCachedTop =480
                    LayoutCachedWidth =5460
                    LayoutCachedHeight =840
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6900
                    Top =480
                    Width =1320
                    TabIndex =4
                    Name ="cmdSaveProfil"
                    Caption ="Simpan "
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =6900
                    LayoutCachedTop =480
                    LayoutCachedWidth =8220
                    LayoutCachedHeight =840
                End
            End
        End
        Begin PageHeader
            DisplayWhen =1
            Height =0
            Name ="PageHeaderSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin Section
            CanGrow = NotDefault
            Height =20640
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9300
                    Top =480
                    Width =660
                    Height =300
                    ColumnWidth =2170
                    Name ="txtKepingAutoID"
                    ControlSource ="KepingAutoID"

                    LayoutCachedLeft =9300
                    LayoutCachedTop =480
                    LayoutCachedWidth =9960
                    LayoutCachedHeight =780
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8520
                            Top =480
                            Width =720
                            Height =300
                            Name ="Label2"
                            Caption ="PKAID"
                            LayoutCachedLeft =8520
                            LayoutCachedTop =480
                            LayoutCachedWidth =9240
                            LayoutCachedHeight =780
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1560
                    Top =60
                    Height =300
                    TabIndex =1
                    ForeColor =0
                    Name ="txtKepingID"
                    ControlSource ="KepingID"

                    LayoutCachedLeft =1560
                    LayoutCachedTop =60
                    LayoutCachedWidth =3000
                    LayoutCachedHeight =360
                    ForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =60
                            Width =1440
                            Height =300
                            ForeColor =0
                            Name ="lbltxtKepingID"
                            Caption ="PKID"
                            LayoutCachedLeft =60
                            LayoutCachedTop =60
                            LayoutCachedWidth =1500
                            LayoutCachedHeight =360
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =1560
                    Top =420
                    Width =4500
                    Height =660
                    ColumnWidth =4300
                    TabIndex =3
                    ForeColor =0
                    Name ="txtDeskripsiProfilKeping"
                    ControlSource ="DeskripsiProfilKeping"

                    LayoutCachedLeft =1560
                    LayoutCachedTop =420
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =1080
                    ForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =420
                            Width =1440
                            Height =300
                            ForeColor =0
                            Name ="Label7"
                            Caption ="Deskripsi PK"
                            LayoutCachedLeft =60
                            LayoutCachedTop =420
                            LayoutCachedWidth =1500
                            LayoutCachedHeight =720
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3660
                    Top =6000
                    Width =660
                    Height =300
                    TabIndex =7
                    ForeColor =0
                    Name ="txtPolaLubangID"
                    ControlSource ="PolaLubangID"

                    LayoutCachedLeft =3660
                    LayoutCachedTop =6000
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =6300
                    ForeTint =100.0
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =6780
                    Top =6030
                    TabIndex =4
                    Name ="chkLis"
                    ControlSource ="Indikator-Lis"
                    StatusBarText ="wear indikator untuk lis"

                    LayoutCachedLeft =6780
                    LayoutCachedTop =6030
                    LayoutCachedWidth =7040
                    LayoutCachedHeight =6270
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =7010
                            Top =6000
                            Width =1170
                            Height =300
                            ForeColor =0
                            Name ="lblchklis"
                            Caption ="Indikator-Lis"
                            LayoutCachedLeft =7010
                            LayoutCachedTop =6000
                            LayoutCachedWidth =8180
                            LayoutCachedHeight =6300
                            ForeTint =100.0
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =6780
                    Top =6390
                    TabIndex =5
                    Name ="chkCoak"
                    ControlSource ="Coak"
                    StatusBarText ="Pencos (setengah lingkaran untuk wear indikator)"

                    LayoutCachedLeft =6780
                    LayoutCachedTop =6390
                    LayoutCachedWidth =7040
                    LayoutCachedHeight =6630
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =7010
                            Top =6360
                            Width =1170
                            Height =300
                            ForeColor =0
                            Name ="lblChkCoak"
                            Caption ="Coak"
                            LayoutCachedLeft =7010
                            LayoutCachedTop =6360
                            LayoutCachedWidth =8180
                            LayoutCachedHeight =6660
                            ForeTint =100.0
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =6780
                    Top =6750
                    TabIndex =6
                    Name ="chkKoak"
                    ControlSource ="Koak"
                    StatusBarText ="koak seperti k-492/k-492 koak"

                    LayoutCachedLeft =6780
                    LayoutCachedTop =6750
                    LayoutCachedWidth =7040
                    LayoutCachedHeight =6990
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =7010
                            Top =6720
                            Width =1170
                            Height =300
                            ForeColor =0
                            Name ="lblchkKoak"
                            Caption ="Koak"
                            LayoutCachedLeft =7010
                            LayoutCachedTop =6720
                            LayoutCachedWidth =8180
                            LayoutCachedHeight =7020
                            ForeTint =100.0
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =87
                    OldBorderStyle =5
                    BorderLineStyle =4
                    Left =600
                    Top =1860
                    Width =9480
                    Height =4032
                    TabIndex =8
                    Name ="sfrm_BlankoKeping"
                    SourceObject ="Form.sfrm_BlankoKeping"
                    LinkChildFields ="BlankoKepingAutoID"
                    LinkMasterFields ="BlankoKepingAutoID"

                    LayoutCachedLeft =600
                    LayoutCachedTop =1860
                    LayoutCachedWidth =10080
                    LayoutCachedHeight =5892
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =120
                    Top =6120
                    TabIndex =9
                    Name ="chkKepingRvt"
                    ControlSource ="LubangFlag"
                    StatusBarText ="'gunakan rivetID apabila tanpa lubang"

                    LayoutCachedLeft =120
                    LayoutCachedTop =6120
                    LayoutCachedWidth =380
                    LayoutCachedHeight =6360
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =350
                            Top =6000
                            Width =910
                            Height =660
                            ForeColor =0
                            Name ="lblchkKepingRivet"
                            Caption ="Keping dgn Rivet"
                            LayoutCachedLeft =350
                            LayoutCachedTop =6000
                            LayoutCachedWidth =1260
                            LayoutCachedHeight =6660
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
                    Top =1140
                    Width =2820
                    Height =300
                    TabIndex =10
                    Name ="cboBlankoKepingID"
                    ControlSource ="BlankoKepingAutoID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT T_BlankoKeping.BlankoKepingAutoID, T_BlankoKeping.BlankoKepingID, T_Blank"
                        "oKeping.DeskripsiBlankoKeping FROM T_BlankoKeping; "
                    ColumnWidths ="0;1440;0"
                    AllowValueListEdits =0

                    LayoutCachedLeft =1560
                    LayoutCachedTop =1140
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =1440
                    ForeThemeColorIndex =0
                    ForeShade =100.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =95
                            TextAlign =1
                            Top =1080
                            Width =10080
                            Height =780
                            TopMargin =252
                            BackColor =15788753
                            ForeColor =0
                            Name ="BlankoKepingLbl"
                            Caption ="BlankoKepingID"
                            LayoutCachedTop =1080
                            LayoutCachedWidth =10080
                            LayoutCachedHeight =1860
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
                    Top =1500
                    Width =2820
                    Height =300
                    TabIndex =11
                    Name ="cboDeskripsiBlanko"
                    ControlSource ="BlankoKepingAutoID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT T_BlankoKeping.BlankoKepingAutoID, T_BlankoKeping.BlankoKepingID, T_Blank"
                        "oKeping.DeskripsiBlankoKeping FROM T_BlankoKeping; "
                    ColumnWidths ="0;0;2880"
                    AllowValueListEdits =0

                    LayoutCachedLeft =1560
                    LayoutCachedTop =1500
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =1800
                    ForeThemeColorIndex =0
                    ForeShade =100.0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =2550
                    Left =2820
                    Top =6000
                    Width =780
                    Height =300
                    TabIndex =12
                    Name ="cboPolaLubangID"
                    ControlSource ="PolaLubangID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT T_UkuranPolaLubang.PolaLubangID, T_UkuranPolaLubang.PolaLubang, T_UkuranP"
                        "olaLubang.DetailPolaLubang FROM T_UkuranPolaLubang ORDER BY T_UkuranPolaLubang.P"
                        "olaLubang; "
                    ColumnWidths ="0;649;1900"
                    AllowValueListEdits =0

                    LayoutCachedLeft =2820
                    LayoutCachedTop =6000
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =6300
                    ForeThemeColorIndex =0
                    ForeShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1620
                            Top =6000
                            Width =1140
                            Height =300
                            ForeColor =0
                            Name ="lblPolaLubangID"
                            Caption ="Pola Lubang"
                            LayoutCachedLeft =1620
                            LayoutCachedTop =6000
                            LayoutCachedWidth =2760
                            LayoutCachedHeight =6300
                            ForeTint =100.0
                        End
                    End
                End
                Begin Image
                    BackStyle =1
                    OldBorderStyle =1
                    Left =5340
                    Top =7080
                    Width =4590
                    Height =3090
                    BorderColor =0
                    Name ="imgGambarProfil"
                    ControlSource ="PKImageCode"

                    LayoutCachedLeft =5340
                    LayoutCachedTop =7080
                    LayoutCachedWidth =9930
                    LayoutCachedHeight =10170
                    TabIndex =19
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5340
                    Top =6660
                    Width =1320
                    TabIndex =13
                    Name ="cmdManageImage"
                    Caption ="Kelola"
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="F_ImageProfilKeping"
                            Argument ="0"
                            Argument =""
                            Argument ="\"[KepingAutoID]=\" & [txtKepingAutoID]"
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"cmdManageImage\" xmlns=\"http://schemas.microsoft.com/office"
                                "/accessservices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Arg"
                                "ument Name=\"FormName\">F_ImagePr"
                        End
                        Begin
                            Comment ="_AXL:ofilKeping</Argument><Argument Name=\"WhereCondition\">\"[KepingAutoID]=\" "
                                "&amp; [txtKepingAutoID]</Argument></Action></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =5340
                    LayoutCachedTop =6660
                    LayoutCachedWidth =6660
                    LayoutCachedHeight =7020
                End
                Begin Subform
                    OverlapFlags =85
                    Left =150
                    Top =7080
                    Width =4590
                    Height =3090
                    TabIndex =14
                    Name ="sfrm_PolaLubangImage"
                    SourceObject ="Form.sfrm_PolaLubangImage"
                    LinkChildFields ="PolaLubangID"
                    LinkMasterFields ="PolaLubangID"

                    LayoutCachedLeft =150
                    LayoutCachedTop =7080
                    LayoutCachedWidth =4740
                    LayoutCachedHeight =10170
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2820
                    Top =6360
                    Width =1500
                    Height =300
                    TabIndex =15
                    Name ="cboUkuranLubangID"
                    ControlSource ="UkuranLubangID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT T_UkuranLubang.UkuranLubangID, T_UkuranLubang.UkuranLubang FROM T_UkuranL"
                        "ubang; "
                    ColumnWidths ="0;1440"

                    LayoutCachedLeft =2820
                    LayoutCachedTop =6360
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =6660
                    ForeThemeColorIndex =0
                    ForeShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1620
                            Top =6360
                            Width =1140
                            Height =300
                            ForeColor =0
                            Name ="Label75"
                            Caption ="Ukr Lubang"
                            LayoutCachedLeft =1620
                            LayoutCachedTop =6360
                            LayoutCachedWidth =2760
                            LayoutCachedHeight =6660
                            ForeTint =100.0
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =2820
                    Top =6720
                    Width =1500
                    Height =300
                    TabIndex =2
                    Name ="cboRivetID"
                    ControlSource ="RivetID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT T_UkuranRivet.RivetID, T_UkuranRivet.KodeRivet, T_UkuranRivet.Ukuran, T_U"
                        "kuranRivet.KodeBarang FROM T_UkuranRivet; "
                    ColumnWidths ="0;1440;1440;1440"

                    LayoutCachedLeft =2820
                    LayoutCachedTop =6720
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =7020
                    ForeThemeColorIndex =0
                    ForeShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1620
                            Top =6720
                            Width =1140
                            Height =300
                            ForeColor =0
                            Name ="Label11"
                            Caption ="RivetID"
                            LayoutCachedLeft =1620
                            LayoutCachedTop =6720
                            LayoutCachedWidth =2760
                            LayoutCachedHeight =7020
                            ForeTint =100.0
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =119
                    TextAlign =2
                    Left =4440
                    Width =5280
                    Height =430
                    FontSize =16
                    FontWeight =700
                    BackColor =5676533
                    BorderColor =5676533
                    ForeColor =255
                    Name ="lblWarning"
                    Caption =" MODE DATA EDIT"
                    LayoutCachedLeft =4440
                    LayoutCachedWidth =9720
                    LayoutCachedHeight =430
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    OverlapFlags =119
                    Left =180
                    Top =6660
                    Width =1320
                    TabIndex =16
                    Name ="cmdHapusLbg"
                    Caption ="Hapus Pencarian"

                    LayoutCachedLeft =180
                    LayoutCachedTop =6660
                    LayoutCachedWidth =1500
                    LayoutCachedHeight =7020
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8610
                    Top =6660
                    Width =1320
                    TabIndex =17
                    Name ="cmdLihatImg"
                    Caption ="Foto"
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="F_ImageProfilKeping"
                            Argument ="0"
                            Argument =""
                            Argument ="\"[KepingAutoID]=\" & [txtKepingAutoID]"
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"cmdLihatImg\" xmlns=\"http://schemas.microsoft.com/office/ac"
                                "cessservices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argume"
                                "nt Name=\"FormName\">F_ImageProfi"
                        End
                        Begin
                            Comment ="_AXL:lKeping</Argument><Argument Name=\"WhereCondition\">\"[KepingAutoID]=\" &am"
                                "p; [txtKepingAutoID]</Argument></Action></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =8610
                    LayoutCachedTop =6660
                    LayoutCachedWidth =9930
                    LayoutCachedHeight =7020
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =8280
                    Top =1140
                    Width =1320
                    TabIndex =18
                    Name ="cmdDaftarBlankoKeping"
                    Caption ="Daftar BK"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =8280
                    LayoutCachedTop =1140
                    LayoutCachedWidth =9600
                    LayoutCachedHeight =1500
                    Overlaps =1
                End
            End
        End
        Begin PageFooter
            DisplayWhen =1
            Height =0
            Name ="PageFooterSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
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
' See "F_ProfilKeping_depcrecated.cls"
