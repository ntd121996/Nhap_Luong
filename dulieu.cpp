#include "dulieu.h"

DuLieu::DuLieu()
{

}
DuLieu::DuLieu( QString MaNhanVien,
        QString TenNhanVien,
        QString MaSanPham,
        QString Ngay,
        QString So,
        QString Loai )
{
    this->MaNhanVien = MaNhanVien;
    this->TenNhanVien = TenNhanVien;
    this->MaSanPham = MaSanPham;
    this->Ngay = Ngay;
    this->So = So;
    this->Loai = Loai;
}
