DROP DATABASE IF EXISTS db_penjualan;
CREATE DATABASE db_penjualan;
 USE db_penjualan;

 CREATE TABLE kategori (
    id_kategori INT AUTO_INCREMENT PRIMARY KEY,
    nama_kategori VARCHAR (50) NOT NULL
 );

 INSERT INTO kategori (nama_kategori)
 VALUES
 ('Mukena');

 CREATE TABLE produk (
    id_produk INT AUTO_INCREMENT PRIMARY KEY,
    id_kategori INT NOT NULL,
    kode_produk VARCHAR(20) NOT NULL,
    nama_produk VARCHAR(100) NOT NULL,
    harga_jual decimal(12,2) NOT NULL,
    status ENUM('Aktif','Non Aktif') DEFAULT 'aktif',

FOREIGN KEY (id_kategori)
REFERENCES kategori(id_kategori)
 );

INSERT INTO produk (nama_produk)
VALUES
('mukena abaya'),
('mukena two piece'),
('mukena ponco'),
('mukena travel');

CREATE TABLE warna(
    id_warna INT AUTO_INCREMENT PRIMARY KEY,
    nama_warna VARCHAR(70) NOT NULL UNIQUE
);

INSERT INTO warna (nama_warna)
VALUES
('Black'),
('white'),
('Taro'),
('Burgundy'),
('Navy'),
('Baby Pink'),
('Blue Electric'),
('Baby Gold'),
('Lavender'),
('Nude Pink'),
('Mustard'),
('Terracota');

CREATE TABLE varian_produk(
    id_varian INT AUTO_INCREMENT PRIMARY KEY,
    id_produk INT NOT NULL,
    id_warna INT NOT NULL,

    FOREIGN KEY (id_produk)
    REFERENCES produk(id_produk),

    FOREIGN KEY (id_warna)
    REFERENCES warna(id_warna)
);

CREATE TABLE barang_masuk(
    id_barang_masuk INT AUTO_INCREMENT PRIMARY KEY,
    no_faktur VARCHAR(30) NOT NULL,
    nama_penjahit VARCHAR(100),
    tanggal DATE NOT NULL,
    keterangan TEXT
);

CREATE TABLE detail_barang_masuk (
    id_detail INT AUTO_INCREMENT PRIMARY KEY,
    id_barang_masuk INT NOT NULL,
    id_varian INT NOT NULL,
    jumlah INT NOT NULL,
    harga_beli DECIMAL(12,2) NOT NULL,

    FOREIGN KEY (id_barang_masuk)
    REFERENCES barang_masuk(id_barang_masuk),

    FOREIGN KEY (id_varian)
    REFERENCES varian_produk(id_varian)
);

CREATE TABLE barang_keluar (
    id_barang_keluar INT AUTO_INCREMENT PRIMARY KEY,
    id_varian INT NOT NULL,
    jumlah INT NOT NULL,
    alasan ENUM('Penjualan','Rusak','Hilang','Retur') DEFAULT 'Penjualan',
    tanggal DATE NOT NULL,
    keterangan TEXT,

    FOREIGN KEY (id_varian)
    REFERENCEs varian_produk(id_varian)
);

CREATE TABLE stok_barang(
    id_stok INT AUTO_INCREMENT PRIMARY KEY,
    id_varian INT NOT NULL,
    stok INT DEFAULT 0,
    stok_minimum INT DEFAULT 5,
    lokasi_rak VARCHAR(50),

    FOREIGN KEY (id_varian)
    REFERENCEs varian_produk(id_varian)
);

CREATE TABLE rak(
    id_rak INT NOT NULL,
    kode_rak INT NOT NULL,
    nama_rak INT NOT NULL,

);