# excel_module_NumberWords
NumberWords adalah modul atau formula excel yang digunakan untuk merubah nominal (angka) menjadi terbilang (kata-kata)
  # cara import modul
  1. buka file excel yang diinginkan, atau bisa juga membuka file excel baru
  2. klik menu Developer di menu bar pada file excel yang sedang dibuka. Jika menu Developer belum aktif, maka aktifkan terlebih     dahulu dengan cara :
	    ## aktifasi menu Developer
	    a. di menu File tab, klik Options > Customize Ribbon.
	    b. pada Customize the Ribbon dan di bawah Main Tabs, pilih Developer check dan centang pada box tersebut, maka menu                 Developer akan tampil.
	
  3. pilih menu Visual Basic
  4. maka jendela baru akan terbuka yaitu jendela Visual Basic Application (VBA)
  5. pada sidebar sebelah kiri, pilih ThisWorkbook
  6. kemudian klik kanan pada ThisWorkbook dan pilih Import File
  7. cari file Number_to_Words.bas dan klik open
  8. maka akan ada module baru yang bernama Number_to_Words yang dapat di cek di side bar sebelah kiri di dalam folder Modules
  9. modul pun siap digunakan dan jangan lupa untuk menyimpan file excel tersebut

  # cara penggunaan modul
  1. tutup jendela VBA dan kembali ke halaman utama
  2. buka sembarang Sheet
  3. masukkan sembarang angka pada suatu cell, misal 750000 pada cell A3
  4. letakkan kursor di luar cell A3, misal A4
  5. pada cell A4 tersebut kita masukkan formula dengan cara =NumberWords(A3)
  6. maka cell A4 akan menampilkan tulisan "tujuh ratus lima puluh ribu"
  7. untuk membuat huruf menjadi huruf besar tiap awal kata maka tambahkan formula proper sehingga penulisan formula akan menjadi     =proper(NumberWords(A3))
  8. maka hasilnya akan menjadi "Tujuh Ratus Lima Puluh Ribu"
  9. untuk menambah kata "rupiah" di akhir kalimat, maka formula akan menjadi =proper(NumberWords(A3)&" rupiah") dan akan             menghasilkan keluaran "Tujuh Ratus Lima Puluh Ribu Rupiah"
