Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =9216
    DatasheetFontHeight =11
    ItemSuffix =30
    Left =3340
    Top =1140
    Right =16100
    Bottom =10800
    Filter ="[PolaLubangID] =42"
    RecSrcDt = Begin
        0xd0bc7f4d2c6fe640
    End
    RecordSource ="F_PolaLubangDetail"
    Caption ="Pola Lubang"
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
            Width =1701
            Height =1701
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
            Width =1701
            LabelX =-1701
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
            Height =1008
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    OverlapFlags =85
                    Width =3070
                    Height =520
                    FontSize =20
                    Name ="Label26"
                    Caption ="Detail Pola Lubang"
                    LayoutCachedWidth =3070
                    LayoutCachedHeight =520
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4980
                    Top =560
                    Width =1320
                    Name ="cmdNewProfilKeping"
                    Caption ="Baru "

                    LayoutCachedLeft =4980
                    LayoutCachedTop =560
                    LayoutCachedWidth =6300
                    LayoutCachedHeight =920
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7740
                    Top =560
                    Width =1320
                    TabIndex =1
                    Name ="cmdSaveProfil"
                    Caption ="Simpan "

                    LayoutCachedLeft =7740
                    LayoutCachedTop =560
                    LayoutCachedWidth =9060
                    LayoutCachedHeight =920
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6360
                    Top =560
                    Width =1320
                    TabIndex =2
                    Name ="cmdEditProfil"
                    Caption ="Edit "

                    LayoutCachedLeft =6360
                    LayoutCachedTop =560
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =920
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =3600
                    Top =560
                    Width =1320
                    TabIndex =3
                    Name ="cmdCancel"
                    Caption ="Cancel"

                    LayoutCachedLeft =3600
                    LayoutCachedTop =560
                    LayoutCachedWidth =4920
                    LayoutCachedHeight =920
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =85
                    TextAlign =2
                    Top =560
                    Width =3460
                    Height =360
                    FontSize =16
                    FontWeight =700
                    BackColor =5676533
                    BorderColor =5676533
                    ForeColor =255
                    Name ="lblWarning"
                    Caption =" WARNING"
                    LayoutCachedTop =560
                    LayoutCachedWidth =3460
                    LayoutCachedHeight =920
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            Height =5328
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8636
                    Top =144
                    Width =441
                    Height =310
                    ColumnWidth =1701
                    ForeColor =8210719
                    Name ="PolaLubangID"
                    ControlSource ="PolaLubangID"

                    LayoutCachedLeft =8636
                    LayoutCachedTop =144
                    LayoutCachedWidth =9077
                    LayoutCachedHeight =454
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7200
                            Top =140
                            Width =1360
                            Height =310
                            Name ="PolaLubangID_Label"
                            Caption ="PolaLubangID"
                            LayoutCachedLeft =7200
                            LayoutCachedTop =140
                            LayoutCachedWidth =8560
                            LayoutCachedHeight =450
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1570
                    Top =150
                    Width =1130
                    Height =310
                    ColumnWidth =2240
                    TabIndex =1
                    Name ="PolaLubang"
                    ControlSource ="PolaLubang"

                    LayoutCachedLeft =1570
                    LayoutCachedTop =150
                    LayoutCachedWidth =2700
                    LayoutCachedHeight =460
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =150
                            Top =150
                            Width =1360
                            Height =310
                            Name ="PolaLubang_Label"
                            Caption ="PolaLubang"
                            LayoutCachedLeft =150
                            LayoutCachedTop =150
                            LayoutCachedWidth =1510
                            LayoutCachedHeight =460
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1570
                    Top =549
                    Width =7500
                    Height =580
                    ColumnWidth =3000
                    TabIndex =2
                    Name ="Catatan"
                    ControlSource ="DetailPolaLubang"
                    StatusBarText ="Informasi Tentang Jarak Lobang yang mungkin berbeda dengan gambar"

                    LayoutCachedLeft =1570
                    LayoutCachedTop =549
                    LayoutCachedWidth =9070
                    LayoutCachedHeight =1129
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =150
                            Top =549
                            Width =1360
                            Height =310
                            Name ="Catatan_Label"
                            Caption ="Catatan"
                            LayoutCachedLeft =150
                            LayoutCachedTop =549
                            LayoutCachedWidth =1510
                            LayoutCachedHeight =859
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1570
                    Top =1233
                    Width =1130
                    Height =310
                    ColumnWidth =2240
                    TabIndex =3
                    Name ="L"
                    ControlSource ="L"

                    LayoutCachedLeft =1570
                    LayoutCachedTop =1233
                    LayoutCachedWidth =2700
                    LayoutCachedHeight =1543
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =720
                            Top =1233
                            Width =790
                            Height =310
                            Name ="L_Label"
                            Caption ="L"
                            LayoutCachedLeft =720
                            LayoutCachedTop =1233
                            LayoutCachedWidth =1510
                            LayoutCachedHeight =1543
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1570
                    Top =1632
                    Width =1130
                    Height =310
                    ColumnWidth =2240
                    TabIndex =4
                    Name ="X"
                    ControlSource ="X"

                    LayoutCachedLeft =1570
                    LayoutCachedTop =1632
                    LayoutCachedWidth =2700
                    LayoutCachedHeight =1942
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =720
                            Top =1632
                            Width =790
                            Height =310
                            Name ="X_Label"
                            Caption ="X"
                            LayoutCachedLeft =720
                            LayoutCachedTop =1632
                            LayoutCachedWidth =1510
                            LayoutCachedHeight =1942
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1570
                    Top =2031
                    Width =1130
                    Height =310
                    ColumnWidth =2240
                    TabIndex =5
                    Name ="A"
                    ControlSource ="A"

                    LayoutCachedLeft =1570
                    LayoutCachedTop =2031
                    LayoutCachedWidth =2700
                    LayoutCachedHeight =2341
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =720
                            Top =2031
                            Width =790
                            Height =310
                            Name ="A_Label"
                            Caption ="A"
                            LayoutCachedLeft =720
                            LayoutCachedTop =2031
                            LayoutCachedWidth =1510
                            LayoutCachedHeight =2341
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1570
                    Top =2430
                    Width =1130
                    Height =310
                    ColumnWidth =2240
                    TabIndex =6
                    Name ="B"
                    ControlSource ="B"

                    LayoutCachedLeft =1570
                    LayoutCachedTop =2430
                    LayoutCachedWidth =2700
                    LayoutCachedHeight =2740
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =720
                            Top =2430
                            Width =790
                            Height =310
                            Name ="B_Label"
                            Caption ="B"
                            LayoutCachedLeft =720
                            LayoutCachedTop =2430
                            LayoutCachedWidth =1510
                            LayoutCachedHeight =2740
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1570
                    Top =2829
                    Width =1130
                    Height =310
                    ColumnWidth =2240
                    TabIndex =7
                    Name ="C"
                    ControlSource ="C"

                    LayoutCachedLeft =1570
                    LayoutCachedTop =2829
                    LayoutCachedWidth =2700
                    LayoutCachedHeight =3139
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =720
                            Top =2829
                            Width =790
                            Height =310
                            Name ="C_Label"
                            Caption ="C"
                            LayoutCachedLeft =720
                            LayoutCachedTop =2829
                            LayoutCachedWidth =1510
                            LayoutCachedHeight =3139
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1570
                    Top =3228
                    Width =1130
                    Height =310
                    ColumnWidth =2240
                    TabIndex =8
                    Name ="D"
                    ControlSource ="D"

                    LayoutCachedLeft =1570
                    LayoutCachedTop =3228
                    LayoutCachedWidth =2700
                    LayoutCachedHeight =3538
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =720
                            Top =3228
                            Width =790
                            Height =310
                            Name ="D_Label"
                            Caption ="D"
                            LayoutCachedLeft =720
                            LayoutCachedTop =3228
                            LayoutCachedWidth =1510
                            LayoutCachedHeight =3538
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1570
                    Top =3627
                    Width =1130
                    Height =310
                    ColumnWidth =2240
                    TabIndex =9
                    Name ="E"
                    ControlSource ="E"

                    LayoutCachedLeft =1570
                    LayoutCachedTop =3627
                    LayoutCachedWidth =2700
                    LayoutCachedHeight =3937
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =720
                            Top =3627
                            Width =790
                            Height =310
                            Name ="E_Label"
                            Caption ="E"
                            LayoutCachedLeft =720
                            LayoutCachedTop =3627
                            LayoutCachedWidth =1510
                            LayoutCachedHeight =3937
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1570
                    Top =4026
                    Width =1130
                    Height =310
                    ColumnWidth =2240
                    TabIndex =10
                    Name ="F"
                    ControlSource ="F"

                    LayoutCachedLeft =1570
                    LayoutCachedTop =4026
                    LayoutCachedWidth =2700
                    LayoutCachedHeight =4336
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =720
                            Top =4026
                            Width =790
                            Height =310
                            Name ="F_Label"
                            Caption ="F"
                            LayoutCachedLeft =720
                            LayoutCachedTop =4026
                            LayoutCachedWidth =1510
                            LayoutCachedHeight =4336
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1570
                    Top =4425
                    Width =1130
                    Height =310
                    ColumnWidth =2240
                    TabIndex =11
                    Name ="G"
                    ControlSource ="G"

                    LayoutCachedLeft =1570
                    LayoutCachedTop =4425
                    LayoutCachedWidth =2700
                    LayoutCachedHeight =4735
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =720
                            Top =4425
                            Width =790
                            Height =310
                            Name ="G_Label"
                            Caption ="G"
                            LayoutCachedLeft =720
                            LayoutCachedTop =4425
                            LayoutCachedWidth =1510
                            LayoutCachedHeight =4735
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1570
                    Top =4824
                    Width =1130
                    Height =310
                    ColumnWidth =2240
                    TabIndex =12
                    Name ="H"
                    ControlSource ="H"

                    LayoutCachedLeft =1570
                    LayoutCachedTop =4824
                    LayoutCachedWidth =2700
                    LayoutCachedHeight =5134
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =720
                            Top =4824
                            Width =790
                            Height =310
                            Name ="H_Label"
                            Caption ="H"
                            LayoutCachedLeft =720
                            LayoutCachedTop =4824
                            LayoutCachedWidth =1510
                            LayoutCachedHeight =5134
                        End
                    End
                End
                Begin Image
                    PictureType =1
                    Left =2879
                    Top =1225
                    Width =6191
                    Height =3381
                    Name ="PolaLubangImage"
                    ControlSource ="ImageCode"

                    LayoutCachedLeft =2879
                    LayoutCachedTop =1225
                    LayoutCachedWidth =9070
                    LayoutCachedHeight =4606
                    TabIndex =14
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4180
                    Top =4747
                    Width =4891
                    Height =430
                    ColumnWidth =7430
                    TabIndex =13
                    Name ="LokasiImage"
                    ControlSource ="LokasiImage"

                    LayoutCachedLeft =4180
                    LayoutCachedTop =4747
                    LayoutCachedWidth =9071
                    LayoutCachedHeight =5177
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2880
                            Top =4752
                            Width =1170
                            Height =430
                            Name ="Label29"
                            Caption ="LokasiImage"
                            LayoutCachedLeft =2880
                            LayoutCachedTop =4752
                            LayoutCachedWidth =4050
                            LayoutCachedHeight =5182
                        End
                    End
                End
            End
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
' See "F_PolaLubangDetail.cls"
