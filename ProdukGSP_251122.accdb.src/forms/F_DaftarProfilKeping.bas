Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    ViewsAllowed =1
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =15480
    DatasheetFontHeight =11
    ItemSuffix =37
    Top =60
    Right =14160
    Bottom =7580
    RecSrcDt = Begin
        0x7e45b8ae8b80e640
    End
    RecordSource ="SELECT T_ProfilKeping.BlankoKepingAutoID, T_ProfilKeping.DeskripsiProfilKeping, "
        "T_ProfilKeping.PolaLubangID, T_ProfilKeping.RivetID, T_ProfilKeping.[Indikator-L"
        "is], T_ProfilKeping.Coak, T_ProfilKeping.Koak, T_ProfilKeping.Catatan, T_ProfilK"
        "eping.ImageCode, T_ProfilKeping.NoteGambar, T_ProfilKeping.LubangFlag, T_ProfilK"
        "eping.TebalTipisFlag, T_UkuranLubang.UkuranLubang FROM T_UkuranLubang RIGHT JOIN"
        " T_ProfilKeping ON T_UkuranLubang.UkuranLubangID = T_ProfilKeping.UkuranLubangID"
        "; "
    Caption ="T_DataKeping"
    DatasheetFontName ="Aptos"
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
            TextFontFamily =0
            FontSize =11
            FontName ="Aptos"
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
            TextFontFamily =0
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Aptos"
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
        Begin FormHeader
            Height =1110
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =118
                    Top =720
                    Width =1350
                    Height =315
                    Name ="KepingID_Label"
                    Caption ="KepingID"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =118
                    LayoutCachedTop =720
                    LayoutCachedWidth =1468
                    LayoutCachedHeight =1035
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =1
                    Left =1468
                    Top =720
                    Width =2760
                    Height =315
                    Name ="NamaBarang_Label"
                    Caption ="NamaBarang"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =1468
                    LayoutCachedTop =720
                    LayoutCachedWidth =4228
                    LayoutCachedHeight =1035
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =4320
                    Top =720
                    Width =420
                    Height =315
                    Name ="LubangFlag_Label"
                    Caption ="LubangFlag"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =4320
                    LayoutCachedTop =720
                    LayoutCachedWidth =4740
                    LayoutCachedHeight =1035
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =3
                    Left =5180
                    Top =720
                    Width =810
                    Height =315
                    Name ="PolaLubangID_Label"
                    Caption ="PolaLubangID"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =5180
                    LayoutCachedTop =720
                    LayoutCachedWidth =5990
                    LayoutCachedHeight =1035
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =3
                    Left =5990
                    Top =720
                    Width =600
                    Height =315
                    Name ="JumlahLubang_Label"
                    Caption ="JumlahLubang"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =5990
                    LayoutCachedTop =720
                    LayoutCachedWidth =6590
                    LayoutCachedHeight =1035
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =1
                    Left =6590
                    Top =720
                    Width =1350
                    Height =315
                    Name ="UkuranRivetID_Label"
                    Caption ="UkuranRivetID"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =6590
                    LayoutCachedTop =720
                    LayoutCachedWidth =7940
                    LayoutCachedHeight =1035
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =7800
                    Top =720
                    Width =540
                    Height =315
                    Name ="TebalTipisFlag_Label"
                    Caption ="TebalTipisFlag"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =7800
                    LayoutCachedTop =720
                    LayoutCachedWidth =8340
                    LayoutCachedHeight =1035
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =8400
                    Top =720
                    Width =300
                    Height =315
                    Name ="Marker_Label"
                    Caption ="Marker"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =8400
                    LayoutCachedTop =720
                    LayoutCachedWidth =8700
                    LayoutCachedHeight =1035
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =1
                    Left =8700
                    Top =720
                    Width =240
                    Height =315
                    Name ="Coak_Label"
                    Caption ="Coak"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =8700
                    LayoutCachedTop =720
                    LayoutCachedWidth =8940
                    LayoutCachedHeight =1035
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =1
                    Left =8940
                    Top =720
                    Width =4260
                    Height =315
                    Name ="Catatan_Label"
                    Caption ="Catatan"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =8940
                    LayoutCachedTop =720
                    LayoutCachedWidth =13200
                    LayoutCachedHeight =1035
                End
                Begin Label
                    OverlapFlags =85
                    Left =60
                    Top =60
                    Width =3900
                    Height =540
                    FontSize =20
                    Name ="Label36"
                    Caption ="Katalog Profil Keping"
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =3960
                    LayoutCachedHeight =600
                End
            End
        End
        Begin Section
            Height =435
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =118
                    Top =60
                    Width =1350
                    Height =330
                    ColumnWidth =3000
                    Name ="KepingID"
                    ControlSource ="KepingID"

                    LayoutCachedLeft =118
                    LayoutCachedTop =60
                    LayoutCachedWidth =1468
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =1468
                    Top =60
                    Width =2760
                    Height =330
                    ColumnWidth =3000
                    TabIndex =1
                    Name ="NamaBarang"
                    ControlSource ="DeskripsiProfilKeping"

                    LayoutCachedLeft =1468
                    LayoutCachedTop =60
                    LayoutCachedWidth =4228
                    LayoutCachedHeight =390
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =4320
                    Top =60
                    Height =330
                    TabIndex =2
                    Name ="LubangFlag"
                    ControlSource ="LubangFlag"

                    LayoutCachedLeft =4320
                    LayoutCachedTop =60
                    LayoutCachedWidth =4580
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =5180
                    Top =60
                    Width =810
                    Height =330
                    ColumnWidth =1530
                    TabIndex =3
                    Name ="PolaLubangID"
                    ControlSource ="PolaLubangID"

                    LayoutCachedLeft =5180
                    LayoutCachedTop =60
                    LayoutCachedWidth =5990
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =5990
                    Top =60
                    Width =600
                    Height =330
                    ColumnWidth =1530
                    TabIndex =4
                    Name ="JumlahLubang"
                    ControlSource ="JumlahLubang_deprecated"

                    LayoutCachedLeft =5990
                    LayoutCachedTop =60
                    LayoutCachedWidth =6590
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =6590
                    Top =60
                    Width =1350
                    Height =330
                    ColumnWidth =3000
                    TabIndex =5
                    Name ="UkuranRivetID"
                    ControlSource ="UkuranRivetID_deprecated"

                    LayoutCachedLeft =6590
                    LayoutCachedTop =60
                    LayoutCachedWidth =7940
                    LayoutCachedHeight =390
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =7940
                    Top =60
                    Height =330
                    TabIndex =6
                    Name ="TebalTipisFlag"
                    ControlSource ="TebalTipisFlag"
                    StatusBarText ="Apabila Kepingan ini ada Tebal Tipis"

                    LayoutCachedLeft =7940
                    LayoutCachedTop =60
                    LayoutCachedWidth =8200
                    LayoutCachedHeight =390
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =8440
                    Top =60
                    Height =330
                    TabIndex =7
                    Name ="Marker"
                    ControlSource ="Indikator-Lis"

                    LayoutCachedLeft =8440
                    LayoutCachedTop =60
                    LayoutCachedWidth =8700
                    LayoutCachedHeight =390
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =8700
                    Top =60
                    Width =240
                    Height =330
                    TabIndex =8
                    Name ="Coak"
                    ControlSource ="Coak"

                    LayoutCachedLeft =8700
                    LayoutCachedTop =60
                    LayoutCachedWidth =8940
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8940
                    Top =60
                    Width =4260
                    Height =330
                    ColumnWidth =3000
                    TabIndex =9
                    Name ="Catatan"
                    ControlSource ="Catatan"

                    LayoutCachedLeft =8940
                    LayoutCachedTop =60
                    LayoutCachedWidth =13200
                    LayoutCachedHeight =390
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
' See "F_DaftarProfilKeping.cls"
