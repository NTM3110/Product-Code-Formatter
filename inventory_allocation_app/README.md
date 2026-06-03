# Phan bo kho hang hoa

Ung dung desktop/web doc Excel de phan bo so luong ban ra tu kho hang hoa truoc khi chuyen phan con thieu sang kho thanh pham.

## Du lieu vao

- `Hoa don mua vao`: file ket qua tu ung dung tao ma VT. Ma da tach gia nhu `ABC.001`, `ABC.002` duoc xem la cac lo ton kho cua ma goc `ABC`.
- `Hoa don ban ra`: file can bo sung ket qua phan bo.
- `Ton dau ky`: file tuy chon gom `Ma VT`, `Ten hang`, `So luong dau ky`, `Don gia von (tuy chon)`.

Neu mot ma VT khong co trong file `Ton dau ky`, ung dung tu dong dat `SL dau ky = 0`. Quy tac nay cung duoc hien ro cho cac ma chi co o hoa don ban ra.
Dong co `Ma VT` trong hoac bang `0` duoc xem la chua xu ly va khong duoc nhap vao kho hang hoa.

Nguoi dung map cot bang ky tu Excel trong giao dien. Mac dinh cho hoa don la:

| Gia tri | Cot |
| --- | --- |
| Ma VT | M |
| Ten hang | N |
| So luong | P |
| Don gia | Q |
| Header / du lieu bat dau | 2 / 3 |

Sau khi upload file, moi dropdown mapping hien theo dang `ky tu cot - ten cot tai dong header`, vi du `M - Ma VT`, de nguoi dung doi chieu dung cot truoc khi tinh phan bo.

## Quy tac phan bo

1. Ton dau ky va dong mua vao tao thanh cac lo trong `kho hang hoa`.
2. Ma `ABC.001`, `ABC.002` cung thuoc ma goc `ABC`.
3. Moi dong ban ra lay toi da tu cac lo cung ma goc.
4. Lo co `don gia ban - don gia von` nho nhat duoc lay truoc, sau do den lo lai cao hon.
5. Lo khong co don gia von van duoc dung truoc kho thanh pham, nhung duoc xep sau cac lo co the tinh lai.
6. Phan so luong khong du trong kho hang hoa duoc ghi vao `SL lay tu kho thanh pham`.

## Khoang lai/lo chap nhan

Nguoi dung co the nhap gioi han `lo toi da` va `lai toi da` tren giao dien. Ty le cua tung lo duoc tinh:

```text
(don gia ban - don gia von) / don gia ban * 100
```

Vi du: `lo toi da = 10`, `lai toi da = 25` chi nhan cac lo trong khoang `-10%` den `+25%`. Lo ngoai khoang khong duoc lay tu kho hang hoa; neu khong con lo hop le, so luong ban se chuyen sang kho thanh pham. Neu bat gioi han ma lo khong co don gia von hoac dong ban khong co don gia ban, lo do cung khong du dieu kien kiem tra.

## File ket qua

Sheet hoa don ban ra goc duoc giu nguyen va them hai cot so luong ket qua:

- `SL lay tu kho hang hoa`
- `SL lay tu kho thanh pham`

De xem tung dong duoc xu ly the nao, file ban ra con duoc them cac cot giai thich:

- `Ton kho truoc khi ban`
- `Chi tiet lay tu kho hang hoa`
- `Ton kho sau khi ban`
- `Lo khong dat khoang lai/lo`

Chi tiet neu co bao gom ma VT lo mua vao, so hoa don mua vao, so luong, gia von va ty le lai/lo cua lo voi dong ban ra.

File cung co cac sheet doi chieu:

- `TongHopKho`
- `PhanBoKho`
- `TonKhoHangHoa`
- `MaChiBanRaKhongTon`: danh sach ma VT co trong hoa don ban ra nhung khong co ma goc tuong ung trong ton dau ky hoac hoa don mua vao.

Ung dung moi chay tren cong `5082`, rieng voi ung dung tao ma VT truoc do.

## Chay va dong goi

Chay tu source:

```bat
run_app.bat
```

Tao file desktop:

```bat
build_exe_auto.bat
```

File chay duoc dat tai `deploy\InventoryAllocator.exe`.
