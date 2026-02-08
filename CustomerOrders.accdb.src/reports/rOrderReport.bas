Version =20
VersionRequired =20
Begin Report
    RecordSource ="qOrderReport"
    Caption ="Order Report"
    Width =10000
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
        End
        Begin TextBox
            FELineBreak = NotDefault
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AsianLineBreak =1
        End
        Begin Line
            BorderLineStyle =0
            Width =1701
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="OrderID"
        End
        Begin FormHeader
            Height =600
            Name ="ReportHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =200
                    Top =100
                    Width =4000
                    Height =400
                    Name ="lblReportTitle"
                    Caption ="Order Report"
                    FontSize =16
                    FontWeight =700

                    LayoutCachedLeft =200
                    LayoutCachedTop =100
                    LayoutCachedWidth =4200
                    LayoutCachedHeight =500
                End
            End
        End
        Begin PageHeader
            Height =0
            Name ="PageHeader"
        End
        Begin BreakHeader
            Height =2400
            Name ="GroupHeader0"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =200
                    Top =100
                    Width =1500
                    Height =315
                    Name ="lblOrderNumber"
                    Caption ="Order Number:"
                    FontWeight =700

                    LayoutCachedLeft =200
                    LayoutCachedTop =100
                    LayoutCachedWidth =1700
                    LayoutCachedHeight =415
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1800
                    Top =100
                    Width =2500
                    Height =315
                    Name ="txtOrderNumber"
                    ControlSource ="OrderNumber"

                    LayoutCachedLeft =1800
                    LayoutCachedTop =100
                    LayoutCachedWidth =4300
                    LayoutCachedHeight =415
                End
                Begin Label
                    OverlapFlags =85
                    Left =5000
                    Top =100
                    Width =1200
                    Height =315
                    Name ="lblOrderDate"
                    Caption ="Order Date:"
                    FontWeight =700

                    LayoutCachedLeft =5000
                    LayoutCachedTop =100
                    LayoutCachedWidth =6200
                    LayoutCachedHeight =415
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =6300
                    Top =100
                    Width =2500
                    Height =315
                    TabIndex =1
                    Name ="txtOrderDate"
                    ControlSource ="OrderDate"
                    Format ="Short Date"

                    LayoutCachedLeft =6300
                    LayoutCachedTop =100
                    LayoutCachedWidth =8800
                    LayoutCachedHeight =415
                End
                Begin Label
                    OverlapFlags =85
                    Left =200
                    Top =500
                    Width =1500
                    Height =315
                    Name ="lblCustomer"
                    Caption ="Customer:"
                    FontWeight =700

                    LayoutCachedLeft =200
                    LayoutCachedTop =500
                    LayoutCachedWidth =1700
                    LayoutCachedHeight =815
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1800
                    Top =500
                    Width =4000
                    Height =315
                    TabIndex =2
                    Name ="txtCustomerName"
                    ControlSource ="=[FirstName] & \" \" & [LastName]"

                    LayoutCachedLeft =1800
                    LayoutCachedTop =500
                    LayoutCachedWidth =5800
                    LayoutCachedHeight =815
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1800
                    Top =900
                    Width =6000
                    Height =315
                    TabIndex =3
                    Name ="txtAddress"
                    ControlSource ="Address"

                    LayoutCachedLeft =1800
                    LayoutCachedTop =900
                    LayoutCachedWidth =7800
                    LayoutCachedHeight =1215
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1800
                    Top =1300
                    Width =6000
                    Height =315
                    TabIndex =4
                    Name ="txtCityStatePost"
                    ControlSource ="=[City] & \", \" & [State] & \" \" & [PostCode]"

                    LayoutCachedLeft =1800
                    LayoutCachedTop =1300
                    LayoutCachedWidth =7800
                    LayoutCachedHeight =1615
                End
                Begin Label
                    OverlapFlags =85
                    Left =200
                    Top =1800
                    Width =4000
                    Height =315
                    Name ="lblColItemDescription"
                    Caption ="Item Description"
                    FontWeight =700

                    LayoutCachedLeft =200
                    LayoutCachedTop =1800
                    LayoutCachedWidth =4200
                    LayoutCachedHeight =2115
                End
                Begin Label
                    OverlapFlags =85
                    Left =4400
                    Top =1800
                    Width =2000
                    Height =315
                    Name ="lblColUnitPrice"
                    Caption ="Unit Price"
                    FontWeight =700

                    LayoutCachedLeft =4400
                    LayoutCachedTop =1800
                    LayoutCachedWidth =6400
                    LayoutCachedHeight =2115
                End
                Begin Label
                    OverlapFlags =85
                    Left =6500
                    Top =1800
                    Width =1200
                    Height =315
                    Name ="lblColQuantity"
                    Caption ="Qty"
                    FontWeight =700

                    LayoutCachedLeft =6500
                    LayoutCachedTop =1800
                    LayoutCachedWidth =7700
                    LayoutCachedHeight =2115
                End
                Begin Label
                    OverlapFlags =85
                    Left =7800
                    Top =1800
                    Width =2000
                    Height =315
                    Name ="lblColSubTotal"
                    Caption ="Sub Total"
                    FontWeight =700

                    LayoutCachedLeft =7800
                    LayoutCachedTop =1800
                    LayoutCachedWidth =9800
                    LayoutCachedHeight =2115
                End
                Begin Line
                    OverlapFlags =85
                    Left =200
                    Top =2200
                    Width =9600
                    Height =0
                    Name ="lineHeaderSep"
                    BorderWidth =2

                    LayoutCachedLeft =200
                    LayoutCachedTop =2200
                    LayoutCachedWidth =9800
                    LayoutCachedHeight =2200
                End
            End
        End
        Begin Section
            Height =360
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    Left =200
                    Top =0
                    Width =4000
                    Height =315
                    Name ="txtItemDescription"
                    ControlSource ="ItemDescription"

                    LayoutCachedLeft =200
                    LayoutCachedTop =0
                    LayoutCachedWidth =4200
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =4400
                    Top =0
                    Width =2000
                    Height =315
                    TabIndex =1
                    Name ="txtUnitPriceNet"
                    ControlSource ="UnitPriceNet"
                    Format ="Currency"

                    LayoutCachedLeft =4400
                    LayoutCachedTop =0
                    LayoutCachedWidth =6400
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =6500
                    Top =0
                    Width =1200
                    Height =315
                    TabIndex =2
                    Name ="txtQuantity"
                    ControlSource ="Quantity"

                    LayoutCachedLeft =6500
                    LayoutCachedTop =0
                    LayoutCachedWidth =7700
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =7800
                    Top =0
                    Width =2000
                    Height =315
                    TabIndex =3
                    Name ="txtSubTotalNet"
                    ControlSource ="SubTotalNet"
                    Format ="Currency"

                    LayoutCachedLeft =7800
                    LayoutCachedTop =0
                    LayoutCachedWidth =9800
                    LayoutCachedHeight =315
                End
            End
        End
        Begin BreakFooter
            Height =1200
            Name ="GroupFooter0"
            Begin
                Begin Line
                    OverlapFlags =85
                    Left =200
                    Top =50
                    Width =9600
                    Height =0
                    Name ="lineFooterSep"
                    BorderWidth =2

                    LayoutCachedLeft =200
                    LayoutCachedTop =50
                    LayoutCachedWidth =9800
                    LayoutCachedHeight =50
                End
                Begin Label
                    OverlapFlags =85
                    Left =5500
                    Top =100
                    Width =2200
                    Height =315
                    Name ="lblTotalNet"
                    Caption ="Total Net:"
                    FontWeight =700

                    LayoutCachedLeft =5500
                    LayoutCachedTop =100
                    LayoutCachedWidth =7700
                    LayoutCachedHeight =415
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =7800
                    Top =100
                    Width =2000
                    Height =315
                    Name ="txtTotalNet"
                    ControlSource ="=Sum([SubTotalNet])"
                    Format ="Currency"

                    LayoutCachedLeft =7800
                    LayoutCachedTop =100
                    LayoutCachedWidth =9800
                    LayoutCachedHeight =415
                End
                Begin Label
                    OverlapFlags =85
                    Left =5500
                    Top =450
                    Width =2200
                    Height =315
                    Name ="lblGST"
                    Caption ="GST (10%):"
                    FontWeight =700

                    LayoutCachedLeft =5500
                    LayoutCachedTop =450
                    LayoutCachedWidth =7700
                    LayoutCachedHeight =765
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =7800
                    Top =450
                    Width =2000
                    Height =315
                    TabIndex =1
                    Name ="txtGST"
                    ControlSource ="=Sum([SubTotalNet]) * 0.1"
                    Format ="Currency"

                    LayoutCachedLeft =7800
                    LayoutCachedTop =450
                    LayoutCachedWidth =9800
                    LayoutCachedHeight =765
                End
                Begin Label
                    OverlapFlags =85
                    Left =5500
                    Top =800
                    Width =2200
                    Height =315
                    Name ="lblGrandTotal"
                    Caption ="Grand Total Inc. GST:"
                    FontWeight =700

                    LayoutCachedLeft =5500
                    LayoutCachedTop =800
                    LayoutCachedWidth =7700
                    LayoutCachedHeight =1115
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =7800
                    Top =800
                    Width =2000
                    Height =315
                    TabIndex =2
                    Name ="txtGrandTotal"
                    ControlSource ="=Sum([SubTotalNet]) * 1.1"
                    Format ="Currency"
                    FontWeight =700

                    LayoutCachedLeft =7800
                    LayoutCachedTop =800
                    LayoutCachedWidth =9800
                    LayoutCachedHeight =1115
                End
            End
        End
        Begin PageFooter
            Height =400
            Name ="PageFooter"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    Left =7000
                    Top =50
                    Width =2800
                    Height =315
                    Name ="txtPageNumber"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"

                    LayoutCachedLeft =7000
                    LayoutCachedTop =50
                    LayoutCachedWidth =9800
                    LayoutCachedHeight =365
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="ReportFooter"
        End
    End
End
