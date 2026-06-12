SELECT
  T_Produk.ProdukID,
  T_Produk.KodeBarang,
  T_Produk.KodeBarangCust,
  T_Produk.DeskripsiCust,
  T_Produk.CustomerID,
  T_Produk.SetAutoID,
  T_Produk.BaseProdukID,
  T_ListCustomer.NamaCustomer
FROM
  T_ListCustomer
  RIGHT JOIN T_Produk ON T_ListCustomer.CustomerID = T_Produk.CustomerID;
