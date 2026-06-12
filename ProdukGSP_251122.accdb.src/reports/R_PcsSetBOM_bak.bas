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
    Width =16140
    DatasheetFontHeight =11
    ItemSuffix =129
    Left =590
    Top =2840
    RecSrcDt = Begin
        0xbb477158b874e640
    End
    RecordSource ="SELECT T_PcsSetBOM.*, T_ProfilKeping.KepingAutoID, T_ProfilKeping.DeskripsiProfi"
        "lKeping FROM T_ProfilKeping INNER JOIN T_PcsSetBOM ON T_ProfilKeping.KepingID = "
        "T_PcsSetBOM.KepingID; "
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
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
            GridlineColor =-2147483609
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =60.0
        End
        Begin Rectangle
            BorderLineStyle =0
            GridlineColor =-2147483609
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
        End
        Begin Line
            BorderLineStyle =0
            GridlineColor =-2147483609
            BorderThemeColorIndex =0
        End
        Begin Image
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            GridlineColor =-2147483609
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            GridlineColor =-2147483609
            ForeThemeColorIndex =0
            ForeTint =75.0
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
        Begin OptionButton
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            GridlineColor =-2147483609
            BorderThemeColorIndex =1
            BorderShade =65.0
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            GridlineColor =-2147483609
            BorderThemeColorIndex =1
            BorderShade =65.0
        End
        Begin OptionGroup
            BorderLineStyle =0
            GridlineColor =-2147483609
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
        End
        Begin BoundObjectFrame
            AddColon = NotDefault
            SizeMode =3
            BorderLineStyle =0
            LabelX =-1800
            GridlineColor =-2147483609
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            GridlineColor =-2147483609
            ShowDatePicker =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
        End
        Begin ListBox
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            GridlineColor =-2147483609
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =0
            ForeTint =75.0
        End
        Begin ComboBox
            AddColon = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            GridlineColor =-2147483609
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
        End
        Begin Subform
            BorderLineStyle =0
            GridlineColor =-2147483609
        End
        Begin Chart
            OldBorderStyle =1
            GridlineColor =-2147483609
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
        End
        Begin UnboundObjectFrame
            OldBorderStyle =1
            GridlineColor =-2147483609
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
        End
        Begin CustomControl
            OldBorderStyle =1
            GridlineColor =-2147483609
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
        End
        Begin ToggleButton
            FontName ="Calibri"
            GridlineColor =-2147483609
            ForeThemeColorIndex =0
            ForeTint =75.0
            UseTheme =1
            Shape =2
            Bevel =1
            BackThemeColorIndex =4
            BackTint =60.0
            OldBorderStyle =0
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
            PressedForeThemeColorIndex =1
        End
        Begin Tab
            FontSize =11
            FontName ="Calibri Light"
            GridlineColor =-2147483609
            ThemeFontIndex =0
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
            GridlineColor =-2147483609
            BorderThemeColorIndex =1
            BorderShade =65.0
        End
        Begin Attachment
            BackStyle =0
            BorderLineStyle =0
            PictureSizeMode =3
            Width =3200
            Height =2560
            GridlineColor =-2147483609
            LabelX =-1800
            AddColon =0
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
        End
        Begin EmptyCell
            Height =240
            GridlineColor =-2147483609
        End
        Begin WebBrowser
            OldBorderStyle =1
            GridlineColor =-2147483609
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
        End
        Begin NavigationControl
            BorderWidth =1
            BorderLineStyle =0
            ForeColor =3484194
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            GridlineColor =-2147483609
        End
        Begin NavigationButton
            GridlineColor =-2147483609
            HoverThemeColorIndex =2
            HoverTint =20.0
            PressedThemeColorIndex =2
            PressedTint =60.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
            BackThemeColorIndex =1
            OldBorderStyle =0
            BorderLineStyle =0
            BorderThemeColorIndex =3
            BorderShade =90.0
            ThemeFontIndex =1
            FontName ="Calibri"
            FontWeight =400
            FontSize =11
            ForeThemeColorIndex =0
            ForeTint =75.0
        End
        Begin ModernChart
            OldBorderStyle =0
            BorderLineStyle =0
            GridlineColor =-2147483609
            BackThemeColorIndex =1
            ChartTitle ="Chart Title"
            ChartTitleFontName ="Calibri"
            ThemeFontIndex =1
            ChartSubtitle ="Subtitle"
        End
        Begin Edge
            OldBorderStyle =1
            GridlineColor =-2147483609
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
        End
        Begin FormHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =0
            Name ="ReportHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
        End
        Begin PageHeader
            Height =0
            Name ="PageHeaderSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =2100
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Subform
                    Left =420
                    Top =540
                    Width =15720
                    Height =1496
                    TabIndex =3
                    Name ="R_ProfilKepingDetail"
                    SourceObject ="Report.R_ProfilKepingDetail"
                    LinkChildFields ="KepingID"
                    LinkMasterFields ="KepingID"

                    LayoutCachedLeft =420
                    LayoutCachedTop =540
                    LayoutCachedWidth =16140
                    LayoutCachedHeight =2036
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =10930
                    Top =30
                    Width =830
                    Height =390
                    TabIndex =1
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    Name ="Qty"
                    ControlSource ="Qty"
                    GroupTable =1

                    LayoutCachedLeft =10930
                    LayoutCachedTop =30
                    LayoutCachedWidth =11760
                    LayoutCachedHeight =420
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    IMESentenceMode =3
                    Left =3020
                    Top =30
                    Width =7840
                    Height =390
                    TabIndex =2
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    Name ="DeskripsiProfilKeping"
                    ControlSource ="DeskripsiProfilKeping"
                    GroupTable =1

                    LayoutCachedLeft =3020
                    LayoutCachedTop =30
                    LayoutCachedWidth =10860
                    LayoutCachedHeight =420
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            TextAlign =1
                            Left =2140
                            Top =30
                            Width =820
                            Height =390
                            LeftMargin =44
                            TopMargin =22
                            RightMargin =44
                            BottomMargin =22
                            Name ="Label1"
                            Caption ="Pcs/Set"
                            LayoutCachedLeft =2140
                            LayoutCachedTop =30
                            LayoutCachedWidth =2960
                            LayoutCachedHeight =420
                            ColumnStart =2
                            ColumnEnd =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1510
                    Top =30
                    Height =390
                    Name ="BaseProdukID"
                    ControlSource ="T_PcsSetBOM.BaseProdukID"
                    GroupTable =1

                    LayoutCachedLeft =1510
                    LayoutCachedTop =30
                    LayoutCachedWidth =2950
                    LayoutCachedHeight =420
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            TextAlign =1
                            Left =50
                            Top =30
                            Width =1390
                            Height =390
                            Name ="Label80"
                            Caption ="T_PcsSetBOM.BaseProdukID"
                            GroupTable =1
                            LayoutCachedLeft =50
                            LayoutCachedTop =30
                            LayoutCachedWidth =1440
                            LayoutCachedHeight =420
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =1500
                    Top =1140
                    Height =300
                    TabIndex =4
                    Name ="KepingID"
                    ControlSource ="KepingID"

                    LayoutCachedLeft =1500
                    LayoutCachedTop =1140
                    LayoutCachedWidth =2940
                    LayoutCachedHeight =1440
                    Begin
                        Begin Label
                            Top =1140
                            Width =850
                            Height =300
                            Name ="Label128"
                            Caption ="KepingID"
                            LayoutCachedTop =1140
                            LayoutCachedWidth =850
                            LayoutCachedHeight =1440
                        End
                    End
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin FormFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =0
            Name ="ReportFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
