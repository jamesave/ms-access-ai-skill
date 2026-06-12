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
    Width =8884
    DatasheetFontHeight =11
    ItemSuffix =34
    Left =200
    Top =1590
    Right =10870
    Bottom =5480
    RecSrcDt = Begin
        0xba07fbd9fd42e640
    End
    RecordSource ="Qry_ProdukCust"
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
            Height =660
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin EmptyCell
                    Left =300
                    Top =60
                    Width =0
                    Height =470
                    Name ="Auto_LogoEmptyCell"
                    GroupTable =1
                    LayoutCachedLeft =300
                    LayoutCachedTop =60
                    LayoutCachedWidth =300
                    LayoutCachedHeight =530
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =360
                    Top =60
                    Width =3500
                    Height =470
                    FontSize =18
                    Name ="Auto_Header0"
                    Caption ="Produk Customer Detail"
                    FontName ="Calibri Light"
                    GroupTable =1
                    LayoutCachedLeft =360
                    LayoutCachedTop =60
                    LayoutCachedWidth =3860
                    LayoutCachedHeight =530
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ThemeFontIndex =0
                    BorderThemeColorIndex =2
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =1
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =2820
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =180
                    Height =300
                    Name ="ProdukID"
                    ControlSource ="ProdukID"

                    LayoutCachedLeft =1500
                    LayoutCachedTop =180
                    LayoutCachedWidth =2940
                    LayoutCachedHeight =480
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =180
                            Width =900
                            Height =300
                            Name ="Label0"
                            Caption ="ProdukID"
                            LayoutCachedTop =180
                            LayoutCachedWidth =900
                            LayoutCachedHeight =480
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =540
                    Height =300
                    TabIndex =1
                    Name ="KodeBarang"
                    ControlSource ="KodeBarang"

                    LayoutCachedLeft =1500
                    LayoutCachedTop =540
                    LayoutCachedWidth =2940
                    LayoutCachedHeight =840
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =540
                            Width =1150
                            Height =300
                            Name ="Label1"
                            Caption ="KodeBarang"
                            LayoutCachedTop =540
                            LayoutCachedWidth =1150
                            LayoutCachedHeight =840
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =1500
                    Top =900
                    Height =300
                    TabIndex =2
                    Name ="KodeBarangCust"
                    ControlSource ="KodeBarangCust"
                    StatusBarText ="Deskripsi Customer"

                    LayoutCachedLeft =1500
                    LayoutCachedTop =900
                    LayoutCachedWidth =2940
                    LayoutCachedHeight =1200
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Top =900
                            Width =1550
                            Height =300
                            Name ="Label2"
                            Caption ="KodeBarangCust"
                            LayoutCachedTop =900
                            LayoutCachedWidth =1550
                            LayoutCachedHeight =1200
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =1260
                    Height =300
                    ColumnWidth =5150
                    TabIndex =3
                    Name ="DeskripsiCust"
                    ControlSource ="DeskripsiCust"
                    StatusBarText ="Untuk Barang MI, SCB, GSP dsb"

                    LayoutCachedLeft =1500
                    LayoutCachedTop =1260
                    LayoutCachedWidth =2940
                    LayoutCachedHeight =1560
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =1260
                            Width =1280
                            Height =300
                            Name ="Label3"
                            Caption ="DeskripsiCust"
                            LayoutCachedTop =1260
                            LayoutCachedWidth =1280
                            LayoutCachedHeight =1560
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5220
                    Top =1260
                    Height =300
                    TabIndex =4
                    Name ="SetAutoID"
                    ControlSource ="SetAutoID"

                    LayoutCachedLeft =5220
                    LayoutCachedTop =1260
                    LayoutCachedWidth =6660
                    LayoutCachedHeight =1560
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3720
                            Top =1260
                            Width =970
                            Height =300
                            Name ="Label5"
                            Caption ="SetAutoID"
                            LayoutCachedLeft =3720
                            LayoutCachedTop =1260
                            LayoutCachedWidth =4690
                            LayoutCachedHeight =1560
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5220
                    Top =180
                    Height =300
                    TabIndex =5
                    Name ="BaseProdukID"
                    ControlSource ="BaseProdukID"

                    LayoutCachedLeft =5220
                    LayoutCachedTop =180
                    LayoutCachedWidth =6660
                    LayoutCachedHeight =480
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3720
                            Top =180
                            Width =1330
                            Height =300
                            Name ="Label6"
                            Caption ="BaseProdukID"
                            LayoutCachedLeft =3720
                            LayoutCachedTop =180
                            LayoutCachedWidth =5050
                            LayoutCachedHeight =480
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =360
            Name ="FormFooter"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
CodeBehindForm
' See "F_ProdukCustomerDetail.cls"
