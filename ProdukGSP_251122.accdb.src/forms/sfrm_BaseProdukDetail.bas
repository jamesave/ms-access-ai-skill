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
    Width =11820
    DatasheetFontHeight =11
    ItemSuffix =68
    Left =200
    Top =100
    Right =15400
    Bottom =11260
    RecSrcDt = Begin
        0x32fdea221342e640
    End
    RecordSource ="Q_BaseProdukDetail"
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =545
            Name ="FormHeader"
            AutoHeight =1
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
                    Width =2120
                    Height =460
                    FontSize =18
                    Name ="Auto_Header0"
                    Caption ="Produk Dasar"
                    FontName ="Calibri Light"
                    GroupTable =2
                    LayoutCachedLeft =300
                    LayoutCachedTop =60
                    LayoutCachedWidth =2420
                    LayoutCachedHeight =520
                    LayoutGroup =1
                    ThemeFontIndex =0
                    BorderThemeColorIndex =2
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =5760
                    Top =240
                    Height =300
                    Name ="StatusProduk"
                    ControlSource ="StatusProduk"

                    LayoutCachedLeft =5760
                    LayoutCachedTop =240
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =540
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4260
                            Top =240
                            Width =1260
                            Height =300
                            Name ="Label54"
                            Caption ="StatusProduk"
                            LayoutCachedLeft =4260
                            LayoutCachedTop =240
                            LayoutCachedWidth =5520
                            LayoutCachedHeight =540
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8640
                    Top =240
                    Height =300
                    ColumnWidth =2775
                    TabIndex =1
                    Name ="KategoriProduk"
                    ControlSource ="KategoriProduk"

                    LayoutCachedLeft =8640
                    LayoutCachedTop =240
                    LayoutCachedWidth =10080
                    LayoutCachedHeight =540
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =7140
                            Top =240
                            Width =1450
                            Height =300
                            Name ="Label55"
                            Caption ="KategoriProduk"
                            LayoutCachedLeft =7140
                            LayoutCachedTop =240
                            LayoutCachedWidth =8590
                            LayoutCachedHeight =540
                        End
                    End
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =16400
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1930
                    Top =420
                    Width =7480
                    Height =300
                    ColumnWidth =2970
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    Name ="Deskripsi"
                    ControlSource ="Deskripsi"
                    BottomPadding =150

                    LayoutCachedLeft =1930
                    LayoutCachedTop =420
                    LayoutCachedWidth =9410
                    LayoutCachedHeight =720
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =130
                            Top =420
                            Width =1670
                            Height =300
                            LeftMargin =44
                            TopMargin =22
                            RightMargin =44
                            BottomMargin =22
                            Name ="Label3"
                            Caption ="Deskripsi"
                            BottomPadding =150
                            LayoutCachedLeft =130
                            LayoutCachedTop =420
                            LayoutCachedWidth =1800
                            LayoutCachedHeight =720
                            RowStart =1
                            RowEnd =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1930
                    Top =60
                    Height =300
                    TabIndex =1
                    Name ="BaseProdukName"
                    ControlSource ="BaseProdukName"
                    StatusBarText ="Copy dari SetAutoID - Now from SetID change to BaseProductName"

                    LayoutCachedLeft =1930
                    LayoutCachedTop =60
                    LayoutCachedWidth =3370
                    LayoutCachedHeight =360
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =60
                            Width =1670
                            Height =300
                            Name ="Label52"
                            Caption ="Kode Produk:"
                            LayoutCachedLeft =120
                            LayoutCachedTop =60
                            LayoutCachedWidth =1790
                            LayoutCachedHeight =360
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8940
                    Top =60
                    Width =480
                    Height =300
                    TabIndex =2
                    Name ="BaseProdukID"
                    ControlSource ="BaseProdukID"

                    LayoutCachedLeft =8940
                    LayoutCachedTop =60
                    LayoutCachedWidth =9420
                    LayoutCachedHeight =360
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7920
                            Top =60
                            Width =960
                            Height =300
                            Name ="Label53"
                            Caption ="Produk ID"
                            LayoutCachedLeft =7920
                            LayoutCachedTop =60
                            LayoutCachedWidth =8880
                            LayoutCachedHeight =360
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    OldBorderStyle =0
                    Left =120
                    Top =780
                    Width =9240
                    Height =600
                    TabIndex =3
                    Name ="sfrm_qcal_ProdukNmPcsPerSetCampurKeping"
                    SourceObject ="Form.sfrm_qcal_ProdukNmPcsPerSetCampurKeping"
                    LinkChildFields ="BaseProdukID"
                    LinkMasterFields ="BaseProdukID"

                    LayoutCachedLeft =120
                    LayoutCachedTop =780
                    LayoutCachedWidth =9360
                    LayoutCachedHeight =1380
                    ShowPageHeaderAndPageFooter =0
                End
                Begin Tab
                    OverlapFlags =85
                    Left =120
                    Top =1500
                    Width =11700
                    Height =14900
                    TabIndex =4
                    Name ="TabCtl62"

                    LayoutCachedLeft =120
                    LayoutCachedTop =1500
                    LayoutCachedWidth =11820
                    LayoutCachedHeight =16400
                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =170
                            Top =1910
                            Width =11600
                            Height =14440
                            Name ="Item Qty"
                            EventProcPrefix ="Item_Qty"
                            LayoutCachedLeft =170
                            LayoutCachedTop =1910
                            LayoutCachedWidth =11770
                            LayoutCachedHeight =16350
                            Begin
                                Begin Subform
                                    OverlapFlags =215
                                    Left =300
                                    Top =2150
                                    Width =6690
                                    Height =5180
                                    Name ="sfrm_PcsSetBOM"
                                    SourceObject ="Form.sfrm_PcsSetBOM"
                                    LinkChildFields ="BaseProdukID"
                                    LinkMasterFields ="BaseProdukID"

                                    LayoutCachedLeft =300
                                    LayoutCachedTop =2150
                                    LayoutCachedWidth =6990
                                    LayoutCachedHeight =7330
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =170
                            Top =1910
                            Width =11600
                            Height =14440
                            Name ="Page64"
                            LayoutCachedLeft =170
                            LayoutCachedTop =1910
                            LayoutCachedWidth =11770
                            LayoutCachedHeight =16350
                        End
                    End
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
' See "sfrm_BaseProdukDetail.cls"
