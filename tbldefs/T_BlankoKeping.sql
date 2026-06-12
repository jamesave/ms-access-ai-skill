CREATE TABLE [T_BlankoKeping] (
  [BlankoKepingAutoID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [BlankoKepingID] VARCHAR (255),
  [DeskripsiBlankoKeping] VARCHAR (255),
  [JenisBlankoKepingID] SINGLE,
  [Panjang] DOUBLE,
  [Panjang2] LONG,
  [Lebar] LONG,
  [Tebal_TipisFlag] BIT,
  [Tebal] SINGLE,
  [Tebal2] SINGLE,
  [DiameterDrum] LONG,
  [DiameterBrakeShoe] LONG,
  [CatatanBlankoKeping] VARCHAR (255)
)
