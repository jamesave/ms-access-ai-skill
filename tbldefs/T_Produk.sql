CREATE TABLE [T_Produk] (
  [ProdukID] AUTOINCREMENT,
  [KodeBarang] VARCHAR (20) CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [KodeBarangCust] VARCHAR (255),
  [DeskripsiCust] VARCHAR (255),
  [CustomerID] LONG CONSTRAINT [T_ListCustomerT_Produk] REFERENCES [T_ListCustomer] ([CustomerID]),
  [SetAutoID] LONG,
  [BaseProdukID] LONG CONSTRAINT [T_BaseProdukT_Produk] REFERENCES [T_BaseProduk] ([BaseProdukID])
)
