Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9360
    DatasheetFontHeight =11
    ItemSuffix =14
    Left =5870
    Top =6570
    Right =15080
    Bottom =10330
    RecSrcDt = Begin
        0xdfe4418e0d77e640
    End
    RecordSource ="T_BlankoKeping"
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
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin Section
            Height =3840
            BackColor =15788753
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8040
                    Top =60
                    Width =1200
                    Height =300
                    ColumnWidth =2400
                    ForeColor =0
                    Name ="BlankoKepingAutoID"
                    ControlSource ="BlankoKepingAutoID"

                    LayoutCachedLeft =8040
                    LayoutCachedTop =60
                    LayoutCachedWidth =9240
                    LayoutCachedHeight =360
                    ForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6960
                            Top =60
                            Width =1020
                            Height =300
                            ForeColor =0
                            Name ="Label0"
                            Caption ="BKAID"
                            LayoutCachedLeft =6960
                            LayoutCachedTop =60
                            LayoutCachedWidth =7980
                            LayoutCachedHeight =360
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1020
                    Top =60
                    Height =300
                    TabIndex =1
                    Name ="BlankoKepingID"
                    ControlSource ="BlankoKepingID"

                    LayoutCachedLeft =1020
                    LayoutCachedTop =60
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =360
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =60
                            Width =900
                            Height =300
                            ForeColor =0
                            Name ="Label1"
                            Caption ="Blanko"
                            LayoutCachedLeft =60
                            LayoutCachedTop =60
                            LayoutCachedWidth =960
                            LayoutCachedHeight =360
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1020
                    Top =420
                    Width =5280
                    Height =660
                    ColumnWidth =3400
                    TabIndex =2
                    Name ="DeskripsiBlankoKeping"
                    ControlSource ="DeskripsiBlankoKeping"
                    StatusBarText ="Deskripsi Dasar Blanko (Sebelum lobang/finishing).  Lebih untuk mengetahui cetak"
                        " mould"

                    LayoutCachedLeft =1020
                    LayoutCachedTop =420
                    LayoutCachedWidth =6300
                    LayoutCachedHeight =1080
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =420
                            Width =900
                            Height =300
                            ForeColor =0
                            Name ="Label2"
                            Caption ="Deskripsi"
                            LayoutCachedLeft =60
                            LayoutCachedTop =420
                            LayoutCachedWidth =960
                            LayoutCachedHeight =720
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8040
                    Top =420
                    Width =1200
                    Height =300
                    TabIndex =3
                    ForeColor =0
                    Name ="JenisBlankoKepingID"
                    ControlSource ="JenisBlankoKepingID"
                    StatusBarText ="1 = Brake Lining; 2= Lempengan/SegiEmpat; 3= Kopling/Donut; 9=Lainnya"

                    LayoutCachedLeft =8040
                    LayoutCachedTop =420
                    LayoutCachedWidth =9240
                    LayoutCachedHeight =720
                    ForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6960
                            Top =420
                            Width =1020
                            Height =300
                            ForeColor =0
                            Name ="Label3"
                            Caption ="Jenis BKID"
                            LayoutCachedLeft =6960
                            LayoutCachedTop =420
                            LayoutCachedWidth =7980
                            LayoutCachedHeight =720
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1020
                    Top =1140
                    Width =720
                    Height =300
                    ColumnWidth =1020
                    TabIndex =4
                    Name ="Panjang"
                    ControlSource ="Panjang"
                    Format ="General Number"
                    StatusBarText ="(mm)"

                    LayoutCachedLeft =1020
                    LayoutCachedTop =1140
                    LayoutCachedWidth =1740
                    LayoutCachedHeight =1440
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1140
                            Width =900
                            Height =300
                            ForeColor =0
                            Name ="Label4"
                            Caption ="Pjg."
                            LayoutCachedLeft =60
                            LayoutCachedTop =1140
                            LayoutCachedWidth =960
                            LayoutCachedHeight =1440
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3240
                    Top =1140
                    Width =720
                    Height =300
                    TabIndex =5
                    ForeColor =0
                    Name ="Panjang2"
                    ControlSource ="Panjang2"
                    StatusBarText ="(mm) Untuk Panjang Bagian Bawah"

                    LayoutCachedLeft =3240
                    LayoutCachedTop =1140
                    LayoutCachedWidth =3960
                    LayoutCachedHeight =1440
                    ForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2340
                            Top =1140
                            Width =840
                            Height =300
                            ForeColor =0
                            Name ="Label5"
                            Caption ="Pjg. bwh"
                            LayoutCachedLeft =2340
                            LayoutCachedTop =1140
                            LayoutCachedWidth =3180
                            LayoutCachedHeight =1440
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1020
                    Top =1500
                    Width =720
                    Height =300
                    ColumnWidth =1020
                    TabIndex =6
                    Name ="Lebar"
                    ControlSource ="Lebar"
                    StatusBarText ="(mm)"

                    LayoutCachedLeft =1020
                    LayoutCachedTop =1500
                    LayoutCachedWidth =1740
                    LayoutCachedHeight =1800
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1500
                            Width =900
                            Height =300
                            ForeColor =0
                            Name ="Label6"
                            Caption ="Lebar"
                            LayoutCachedLeft =60
                            LayoutCachedTop =1500
                            LayoutCachedWidth =960
                            LayoutCachedHeight =1800
                            ForeTint =100.0
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =2280
                    Top =1860
                    TabIndex =7
                    Name ="Tebal_TipisFlag"
                    ControlSource ="Tebal_TipisFlag"
                    StatusBarText ="Untuk ketebalan yang tidak sama antara satu ujung dengan yang lain"

                    LayoutCachedLeft =2280
                    LayoutCachedTop =1860
                    LayoutCachedWidth =2540
                    LayoutCachedHeight =2100
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2520
                            Top =1860
                            Width =1080
                            Height =300
                            ForeColor =0
                            Name ="Label7"
                            Caption ="Tebal/Tipis"
                            LayoutCachedLeft =2520
                            LayoutCachedTop =1860
                            LayoutCachedWidth =3600
                            LayoutCachedHeight =2160
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1020
                    Top =1860
                    Width =720
                    Height =300
                    ColumnWidth =1370
                    TabIndex =8
                    Name ="Tebal"
                    ControlSource ="Tebal"
                    StatusBarText ="(mm). Tebal (untuk keping yang tebal-tipis, ketebalan yang tertinggi- yang lebih"
                        " tepis di Tebal2)"

                    LayoutCachedLeft =1020
                    LayoutCachedTop =1860
                    LayoutCachedWidth =1740
                    LayoutCachedHeight =2160
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1860
                            Width =900
                            Height =300
                            ForeColor =0
                            Name ="Label8"
                            Caption ="Tebal"
                            LayoutCachedLeft =60
                            LayoutCachedTop =1860
                            LayoutCachedWidth =960
                            LayoutCachedHeight =2160
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5580
                    Top =1860
                    Width =720
                    Height =300
                    ColumnWidth =1180
                    TabIndex =9
                    ForeColor =0
                    Name ="Tebal2"
                    ControlSource ="Tebal2"
                    StatusBarText ="(mm) Apabila ada yang bagian tebal tipis.  Ketebalan tipis untuk yang tipis"

                    LayoutCachedLeft =5580
                    LayoutCachedTop =1860
                    LayoutCachedWidth =6300
                    LayoutCachedHeight =2160
                    ForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4500
                            Top =1860
                            Width =1020
                            Height =300
                            ForeColor =0
                            Name ="Label9"
                            Caption ="Tebal2"
                            LayoutCachedLeft =4500
                            LayoutCachedTop =1860
                            LayoutCachedWidth =5520
                            LayoutCachedHeight =2160
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5580
                    Top =1140
                    Width =720
                    Height =300
                    ColumnWidth =1680
                    TabIndex =10
                    ForeColor =0
                    Name ="DiameterDrum"
                    ControlSource ="DiameterDrum"
                    Format ="General Number"

                    LayoutCachedLeft =5580
                    LayoutCachedTop =1140
                    LayoutCachedWidth =6300
                    LayoutCachedHeight =1440
                    ForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4500
                            Top =1140
                            Width =1020
                            Height =300
                            ForeColor =0
                            Name ="Label10"
                            Caption ="Dia. Drum"
                            LayoutCachedLeft =4500
                            LayoutCachedTop =1140
                            LayoutCachedWidth =5520
                            LayoutCachedHeight =1440
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5580
                    Top =1500
                    Width =720
                    Height =300
                    TabIndex =11
                    ForeColor =0
                    Name ="DiameterBrakeShoe"
                    ControlSource ="DiameterBrakeShoe"

                    LayoutCachedLeft =5580
                    LayoutCachedTop =1500
                    LayoutCachedWidth =6300
                    LayoutCachedHeight =1800
                    ForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4500
                            Top =1500
                            Width =1020
                            Height =300
                            ForeColor =0
                            Name ="Label11"
                            Caption ="Dia. Sepatu"
                            LayoutCachedLeft =4500
                            LayoutCachedTop =1500
                            LayoutCachedWidth =5520
                            LayoutCachedHeight =1800
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1020
                    Top =2220
                    Width =5280
                    Height =1620
                    TabIndex =12
                    Name ="CatatanBlankoKeping"
                    ControlSource ="CatatanBlankoKeping"

                    LayoutCachedLeft =1020
                    LayoutCachedTop =2220
                    LayoutCachedWidth =6300
                    LayoutCachedHeight =3840
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =2220
                            Width =900
                            Height =300
                            ForeColor =0
                            Name ="Label12"
                            Caption ="Catatan"
                            LayoutCachedLeft =60
                            LayoutCachedTop =2220
                            LayoutCachedWidth =960
                            LayoutCachedHeight =2520
                            ForeTint =100.0
                        End
                    End
                End
            End
        End
    End
End
