CREATE TABLE [T_BaseProduk] (
  [BaseProdukID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [BaseProdukName] VARCHAR (20),
  [Deskripsi] VARCHAR (50),
  [Catatan] LONGTEXT,
  [AxlePosID] LONG CONSTRAINT [T_ListAxlePosT_BaseProduk] REFERENCES [T_ListAxlePos] ([AxlePosID]),
  [KategoriProdukID] LONG CONSTRAINT [T_KategoriProdukT_BaseProduk] REFERENCES [T_KategoriProduk] ([KategoriProdukID]),
  [StatusProdukID] LONG CONSTRAINT [T_ListStatusBaseProdukT_BaseProduk] REFERENCES [T_ListStatusBaseProduk] ([StatusProdukID])
)
