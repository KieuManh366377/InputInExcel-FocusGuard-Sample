# InputInExcel-FocusGuard-Sample

Delphi COM Add-in mẫu cho Excel (VCL Form modeless, RAD Studio 13.1) minh
họa cách xử lý triệt để vấn đề Excel "cướp" keyboard focus qua lại giữa
Add-in Form và Sheet - một vấn đề kinh điển không có lời giải trọn vẹn
trong các tài liệu phổ biến.

## Vấn đề gốc

Khi mở một VCL Form modeless (không modal) từ trong COM Add-in của Excel,
gõ dữ liệu trên Form có thể bị "rò" xuống Sheet đang active, hoặc ngược
lại, thao tác trên Sheet làm Form mất focus / không nhận phím đúng cách.
Nguyên nhân là cơ chế dispatch message và quản lý focus của Excel can
thiệp sâu vào WM_SETFOCUS / WM_KILLFOCUS theo cách không được tài liệu
hóa đầy đủ.

Tham khảo thảo luận chi tiết về hiện tượng này:
https://stackoverflow.com/questions/56058580/excel-stealing-keyboard-focus-from-vcl-form-in-addin

Các hướng xử lý phổ biến (subclass ở tầng VCL, `Application.OnMessage`,
`CM_UIACTIVATE`, hook `WH_CBT`...) thường chỉ vá được một phần, vẫn bị
Excel giành lại focus khi user thao tác nhanh qua lại giữa Form và Sheet.

## Giải pháp trong project này

`OfficeFocusGuard.dll` (viết bằng C++ Builder, dùng chung được cho nhiều
ngôn ngữ: Delphi, Visual C++, C#...) subclass trực tiếp WndProc thật sự
của cả Form và các control liên quan, chặn đúng message trước khi nó lan
tới message loop của Excel. Kết quả: gõ trên Form không rò xuống Sheet,
gõ trên Sheet không bị Form giành phím, kể cả khi Form đang ở chế độ
luôn-hiện-trên-cùng (`fsStayOnTop`).

Cơ chế này dùng được cho cả hai kiểu UI Add-in phổ biến:
- Modeless Form top-level (`OFG_Install` trên handle của Form).
- Custom TaskPane dock vào cửa sổ Excel (`OFG_InstallControlGuard` trên
  handle của control bên trong TaskPane).

**Lưu ý quan trọng:** repo này CHỈ chia sẻ file nhị phân đã biên dịch
`OfficeFocusGuard.dll` cùng với unit binding Delphi (`OfficeFocusGuard.pas`,
chỉ khai báo `external`, không chứa logic). Mã nguồn C++ Builder của
`OfficeFocusGuard.dll` KHÔNG được công khai trong repo này.

## Cấu trúc project

```
InputInExcel_Clean/
  InputInExcel.dpr          Project chinh (COM Add-in library)
  InputInExcel_TLB.pas      Unit sinh tu Type Library (khong sua tay)
  uDelphiAddIn.pas          IDTExtensibility2, quan ly vong doi Form theo
                             cac su kien cua Excel (startup/shutdown)
  OfficeFocusGuard.pas      Binding external toi OfficeFocusGuard.dll
  Form/
    Form.pas                VCL Form mau (singleton, tu giai phong an toan)
    Form.dfm
```

## Điểm thiết kế đáng chú ý

- **Vòng đời Form gắn với vòng đời Excel:** `Form.FreeForm` được gọi
  trong `OnBeginShutdown` và `OnDisconnection` của Add-in, TRƯỚC khi
  Excel thật sự phá hủy handle/COM object. Nhờ vậy `OFG_Uninstall` luôn
  chạy đúng thời điểm, tránh Access Violation lúc thoát Excel.
- **Form1 là singleton:** gọi `ShowForm` nhiều lần không tạo instance
  trùng, chỉ đưa form đã mở lên trước.
- **Không tự import ngược chính DLL:** `ShowForm` / `FreeForm` được gọi
  trực tiếp qua tham chiếu unit `Form`, không dùng `external` trỏ về
  chính file DLL đang biên dịch.

## Yêu cầu build

- RAD Studio 13.1 (Delphi), Excel COM Add-in project.
- Đã import Type Library `Excel2010` (Microsoft Excel Object Library) và
  `AddInDesignerObjects_TLB`.
- `OfficeFocusGuard.dll` (đúng bitness với Add-in: 32-bit hoặc 64-bit)
  phải nằm cùng thư mục với `InputInExcel.dll` sau khi build, hoặc nằm
  trong PATH hệ thống.

## Ghi chú bản quyền / chia sẻ

- Code Delphi trong repo: xem file `LICENSE` (MIT).
- `OfficeFocusGuard.dll`: chỉ phân phối bản biên dịch sẵn, không kèm mã
  nguồn.
