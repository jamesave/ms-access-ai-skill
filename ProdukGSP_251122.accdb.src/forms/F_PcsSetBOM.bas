Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =2
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6120
    DatasheetFontHeight =11
    ItemSuffix =30
    Left =20
    Top =1050
    Right =8630
    Bottom =6370
    RecSrcDt = Begin
        0x869e5818b048e640
    End
    RecordSource ="SELECT T_PcsSetBOM.*, T_BaseProduk.BaseProdukName FROM T_BaseProduk LEFT JOIN T_"
        "PcsSetBOM ON T_BaseProduk.BaseProdukID = T_PcsSetBOM.BaseProdukID; "
    Caption ="F_PcsSetBOM_sfr"
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
        Begin FormHeader
            Height =1320
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Rectangle
                    BackStyle =1
                    OverlapFlags =93
                    Top =420
                    Width =4080
                    Height =840
                    BackColor =8421504
                    Name ="Box19"
                    LayoutCachedTop =420
                    LayoutCachedWidth =4080
                    LayoutCachedHeight =1260
                    BackThemeColorIndex =-1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4140
                    Top =60
                    Width =1860
                    Name ="TambahProdukDasar"
                    Caption ="Tambah Keping"
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

                    LayoutCachedLeft =4140
                    LayoutCachedTop =60
                    LayoutCachedWidth =6000
                    LayoutCachedHeight =420
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4140
                    Top =480
                    Width =1860
                    TabIndex =1
                    Name ="EditProdukDasarBtn"
                    Caption ="Edit Set"

                    LayoutCachedLeft =4140
                    LayoutCachedTop =480
                    LayoutCachedWidth =6000
                    LayoutCachedHeight =840
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =3000
                    Top =420
                    Width =1080
                    Height =300
                    TabIndex =2
                    Name ="HapusKatFiltBtn"
                    Caption ="Hapus"

                    LayoutCachedLeft =3000
                    LayoutCachedTop =420
                    LayoutCachedWidth =4080
                    LayoutCachedHeight =720
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =1920
                    Top =900
                    Height =300
                    TabIndex =3
                    Name ="BaseProdukName"
                    ControlSource ="BaseProdukName"
                    StatusBarText ="BaseProdukName"

                    LayoutCachedLeft =1920
                    LayoutCachedTop =900
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =1200
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =120
                            Top =900
                            Width =1670
                            Height =300
                            Name ="Label26"
                            Caption ="Produk Dasar"
                            LayoutCachedLeft =120
                            LayoutCachedTop =900
                            LayoutCachedWidth =1790
                            LayoutCachedHeight =1200
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4140
                    Top =900
                    Width =1860
                    TabIndex =4
                    Name ="Command28"
                    Caption ="Simpan Set"

                    LayoutCachedLeft =4140
                    LayoutCachedTop =900
                    LayoutCachedWidth =6000
                    LayoutCachedHeight =1260
                End
                Begin Label
                    OverlapFlags =85
                    Left =60
                    Width =850
                    Height =300
                    Name ="Label29"
                    Caption ="PCS / Set"
                    LayoutCachedLeft =60
                    LayoutCachedWidth =910
                    LayoutCachedHeight =300
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =3120
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Subform
                    OverlapFlags =85
                    Width =4590
                    Name ="F_PcsSetBOM_sfr"
                    SourceObject ="Form.F_PcsSetBOM_sfr_old"
                    LinkChildFields ="SetBOMID"
                    LinkMasterFields ="SetBOMID"

                    LayoutCachedWidth =4590
                    LayoutCachedHeight =1440
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
' See "F_PcsSetBOM.cls"
