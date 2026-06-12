Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =22360
    DatasheetFontHeight =11
    ItemSuffix =6
    Left =1950
    Top =2740
    RecSrcDt = Begin
        0x24238ed7cb1fe640
    End
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    ShowPageMargins =0
    FitToPage =1
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
        Begin Subform
            BorderLineStyle =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =545
            Name ="ReportHeader"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin EmptyCell
                    Left =300
                    Top =60
                    Width =690
                    Height =460
                    Name ="Auto_LogoEmptyCell"
                    GroupTable =3
                    GridlineColor =10921638
                    LayoutCachedLeft =300
                    LayoutCachedTop =60
                    LayoutCachedWidth =990
                    LayoutCachedHeight =520
                    RowEnd =1
                    LayoutGroup =2
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =9720
                    Top =60
                    Width =1150
                    Height =170
                    Name ="Auto_DateEmptyCell"
                    GroupTable =3
                    GridlineColor =10921638
                    LayoutCachedLeft =9720
                    LayoutCachedTop =60
                    LayoutCachedWidth =10870
                    LayoutCachedHeight =230
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =9720
                    Top =290
                    Width =1150
                    Height =230
                    Name ="Auto_TimeEmptyCell"
                    GroupTable =3
                    GridlineColor =10921638
                    LayoutCachedLeft =9720
                    LayoutCachedTop =290
                    LayoutCachedWidth =10870
                    LayoutCachedHeight =520
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    GroupTable =3
                End
                Begin Label
                    TextAlign =1
                    Left =1050
                    Top =60
                    Width =8611
                    Height =460
                    FontSize =18
                    Name ="Auto_Header0"
                    Caption ="Daftar Blanko Keping "
                    FontName ="Calibri Light"
                    GroupTable =3
                    GridlineColor =10921638
                    HorizontalAnchor =2
                    LayoutCachedLeft =1050
                    LayoutCachedTop =60
                    LayoutCachedWidth =9661
                    LayoutCachedHeight =520
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    ThemeFontIndex =0
                    BorderThemeColorIndex =2
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =3
                End
            End
        End
        Begin PageHeader
            Height =0
            Name ="PageHeaderSection"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =2060
            Name ="Detail"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Subform
                    Locked = NotDefault
                    CanShrink = NotDefault
                    Left =30
                    Top =30
                    Width =15160
                    Height =2000
                    Name ="T_BlankoKeping"
                    SourceObject ="Table.T_BlankoKeping"
                    GroupTable =2
                    GridlineColor =10921638

                    LayoutCachedLeft =30
                    LayoutCachedTop =30
                    LayoutCachedWidth =15190
                    LayoutCachedHeight =2030
                    LayoutGroup =1
                    GroupTable =2
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
