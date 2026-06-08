# PyCompare Studio

Desktop tool Python de compare va dong bo file/folder, chay duoc tren Windows va Linux.

## Tinh nang

- Folder Compare: so sanh thu muc, o dia cuc bo, hoac file nen `.zip`, `.tar`, `.tar.gz`, `.tgz`, `.gz`; `.7z` neu cai `py7zr`.
- Folder Sync: copy chieu trai sang phai, phai sang trai, hoac mirror mot ben sang ben kia.
- Text Compare: load, so sanh, highlight khac biet, chinh sua truc tiep, luu file, word wrap, bo qua whitespace/dong trong/comment.
- Table Compare: so sanh tung o cho `.csv` va `.xlsx`.
- Picture Compare: side-by-side, overlay, va highlight pixel khac nhau.
- Hex Compare: so sanh file nhi phan theo tung dong byte.
- Settings: doi mau highlight khac biet va bat/tat khoi dong cung Windows.

## Chay app

Python 3.10+:

```bash
python compare_tool.py
```

Neu muon day du tinh nang `.xlsx`, anh va `.7z`:

```bash
pip install -r requirements.txt
python compare_tool.py
```

## Build executable

Dung script build:

```bash
python build_executable.py --clean
```

Ket qua nam trong thu muc `dist/`.

Tren Windows, lenh tren tao `dist/PyCompareStudio.exe`.

Tren Linux, chay cung script do tren may Linux de tao binary Linux:

```bash
python3 -m pip install -r requirements.txt
python3 build_executable.py --clean --target linux
```

Neu muon tao dang folder thay vi mot file don:

```bash
python build_executable.py --onedir --clean
```

Luu y: PyInstaller build theo he dieu hanh hien tai. Muon co file cho Windows va Linux thi can build mot lan tren Windows va mot lan tren Linux.

## Menu chuot phai

Cai menu chuot phai khi chay bang source Python:

```bash
python context_menu.py install
```

Go menu chuot phai:

```bash
python context_menu.py uninstall
```

Sau khi build `.exe`, cai menu tro thang vao file executable:

```bash
python context_menu.py install --app-path dist\PyCompareStudio.exe
```

### Cach dung tren Windows

Chon tung file don le:

1. Click chuot phai File A.
2. Chon `PyCompareStudio > Select left file to compare`.
3. Di chuyen den File B.
4. Click chuot phai File B.
5. Chon `PyCompareStudio > Compare to <ten File A>`.

Chon nhanh 2 file cung luc:

1. Quet chon dung 2 file hoac folder.
2. Click chuot phai.
3. Chon `PyCompareStudio > Compare`.

Tren Windows 11, neu menu compact chua hien submenu nay, chon `Show more options`.

Neu Windows Explorer khong truyen du 2 file cho menu top-level o mot so may, dung fallback on dinh:

1. Quet chon dung 2 file hoac folder.
2. Click chuot phai.
3. Chon `Send to > PyCompare Studio Compare`.

### Cach dung tren Linux

Nautilus:

```bash
python3 context_menu.py install
```

Sau do click chuot phai file/folder va vao `Scripts > PyCompare Studio`.

Dolphin:

```bash
python3 context_menu.py install
```

Sau do click chuot phai file/folder va vao `Actions`.

## Ghi chu

- Cac tinh nang loi Folder/Text/CSV/Hex dung thu vien chuan Python.
- Picture Compare can `pillow`.
- Table Compare voi `.xlsx` can `openpyxl`.
- File nen `.7z` can `py7zr`; `.zip` va `.tar` khong can cai them.
- Folder Sync chi ap dung voi thu muc that tren may, khong ghi nguoc vao file nen.
