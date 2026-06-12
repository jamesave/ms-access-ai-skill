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
    ItemSuffix =21
    Left =3000
    Top =1140
    Right =13340
    Bottom =10010
    Filter ="[KepingAutoID]=54"
    RecSrcDt = Begin
        0x2a59ec923a67e640
    End
    OnDirty ="[Event Procedure]"
    RecordSource ="SELECT T_ProfilKeping.KepingAutoID, T_ProfilKeping.KepingID, T_ProfilKeping.Desk"
        "ripsiProfilKeping, T_ProfilKeping.ImageCode, T_ProfilKeping.NoteGambar, GetPKIma"
        "gePath([KepingAutoID]) AS ImagePath FROM T_ProfilKeping; "
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
        Begin BoundObjectFrame
            AddColon = NotDefault
            SizeMode =3
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =-1800
            BackThemeColorIndex =1
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
        Begin FormHeader
            Height =960
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =240
                    Top =60
                    Width =4080
                    Height =470
                    FontSize =18
                    Name ="Auto_Header0"
                    Caption ="Kelola Gambar Profil Keping"
                    FontName ="Calibri Light"
                    GroupTable =1
                    LayoutCachedLeft =240
                    LayoutCachedTop =60
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =530
                    LayoutGroup =1
                    ThemeFontIndex =0
                    BorderThemeColorIndex =2
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8640
                    Top =540
                    Width =1320
                    TabIndex =1
                    Name ="cmdEditFotoProfil"
                    Caption ="Edit Gambar"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =8640
                    LayoutCachedTop =540
                    LayoutCachedWidth =9960
                    LayoutCachedHeight =900
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7260
                    Top =540
                    Width =1320
                    TabIndex =2
                    Name ="cmdSimpanFotoProfil"
                    Caption ="Simpan"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =7260
                    LayoutCachedTop =540
                    LayoutCachedWidth =8580
                    LayoutCachedHeight =900
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =5880
                    Top =540
                    Width =1320
                    Name ="cmdCancel"
                    Caption ="Cancel"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =5880
                    LayoutCachedTop =540
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =900
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =240
                    Top =600
                    Width =840
                    Height =300
                    ColumnWidth =2170
                    TabIndex =3
                    ForeColor =8210719
                    Name ="KepingAutoID"
                    ControlSource ="KepingAutoID"

                    LayoutCachedLeft =240
                    LayoutCachedTop =600
                    LayoutCachedWidth =1080
                    LayoutCachedHeight =900
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7200
                    Top =60
                    Width =1380
                    Height =360
                    TabIndex =4
                    Name ="Combo82"
                    RowSourceType ="Table/Query"

                    LayoutCachedLeft =7200
                    LayoutCachedTop =60
                    LayoutCachedWidth =8580
                    LayoutCachedHeight =420
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =6600
                            Top =60
                            Width =540
                            Height =360
                            Name ="Label83"
                            Caption ="Cari"
                            LayoutCachedLeft =6600
                            LayoutCachedTop =60
                            LayoutCachedWidth =7140
                            LayoutCachedHeight =420
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8640
                    Top =60
                    Width =1320
                    TabIndex =5
                    Name ="cmdKembaliKeDaftarBaseProduk"
                    Caption ="Daftar Produk"

                    LayoutCachedLeft =8640
                    LayoutCachedTop =60
                    LayoutCachedWidth =9960
                    LayoutCachedHeight =420
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4500
                    Top =540
                    Width =1320
                    TabIndex =6
                    Name ="cmdDeleteImg"
                    Caption ="Hapus"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =4500
                    LayoutCachedTop =540
                    LayoutCachedWidth =5820
                    LayoutCachedHeight =900
                    Overlaps =1
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =7560
            Name ="Detail"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =120
                    Width =1860
                    Height =300
                    Name ="txtKepingID"
                    ControlSource ="KepingID"

                    LayoutCachedLeft =1500
                    LayoutCachedTop =120
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =420
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =120
                            Width =860
                            Height =300
                            Name ="lblKepingID"
                            Caption ="KepingID"
                            LayoutCachedLeft =120
                            LayoutCachedTop =120
                            LayoutCachedWidth =980
                            LayoutCachedHeight =420
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =540
                    Width =3660
                    Height =600
                    ColumnWidth =4300
                    TabIndex =1
                    Name ="txtDeskripsiProfilKeping"
                    ControlSource ="DeskripsiProfilKeping"

                    LayoutCachedLeft =1500
                    LayoutCachedTop =540
                    LayoutCachedWidth =5160
                    LayoutCachedHeight =1140
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =540
                            Width =1320
                            Height =600
                            Name ="lblDeskripsiProfilKeping"
                            Caption ="DeskripsiProfilKeping"
                            LayoutCachedLeft =120
                            LayoutCachedTop =540
                            LayoutCachedWidth =1440
                            LayoutCachedHeight =1140
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6420
                    Top =3180
                    Width =3540
                    Height =300
                    TabIndex =2
                    Name ="txtImageCode"
                    ControlSource ="ImageCode"

                    LayoutCachedLeft =6420
                    LayoutCachedTop =3180
                    LayoutCachedWidth =9960
                    LayoutCachedHeight =3480
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5280
                            Top =3180
                            Width =1080
                            Height =300
                            Name ="lblImageCode"
                            Caption ="ImageCode"
                            LayoutCachedLeft =5280
                            LayoutCachedTop =3180
                            LayoutCachedWidth =6360
                            LayoutCachedHeight =3480
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =1200
                    Width =3660
                    Height =1920
                    TabIndex =3
                    Name ="txtNoteGambar"
                    ControlSource ="NoteGambar"

                    LayoutCachedLeft =1500
                    LayoutCachedTop =1200
                    LayoutCachedWidth =5160
                    LayoutCachedHeight =3120
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1200
                            Width =1230
                            Height =300
                            Name ="Label9"
                            Caption ="NoteGambar"
                            LayoutCachedLeft =120
                            LayoutCachedTop =1200
                            LayoutCachedWidth =1350
                            LayoutCachedHeight =1500
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =3540
                    Top =60
                    Width =6420
                    Height =420
                    FontSize =16
                    FontWeight =700
                    BackColor =5676533
                    BorderColor =5676533
                    ForeColor =255
                    Name ="lblWarning"
                    Caption =" WARNING"
                    LayoutCachedLeft =3540
                    LayoutCachedTop =60
                    LayoutCachedWidth =9960
                    LayoutCachedHeight =480
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =120
                    Top =3180
                    Width =5040
                    Height =300
                    FontWeight =700
                    Name ="Label17"
                    Caption ="Catatan: Form ini hanya untuk ubah/ lihat gambar"
                    LayoutCachedLeft =120
                    LayoutCachedTop =3180
                    LayoutCachedWidth =5160
                    LayoutCachedHeight =3480
                End
                Begin Image
                    PictureType =1
                    Left =5280
                    Top =540
                    Width =4680
                    Height =2580
                    Name ="ImgViewer"
                    ControlSource ="ImagePath"

                    LayoutCachedLeft =5280
                    LayoutCachedTop =540
                    LayoutCachedWidth =9960
                    LayoutCachedHeight =3120
                    TabIndex =4
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
' See "F_ImagePolaLubang.cls"
