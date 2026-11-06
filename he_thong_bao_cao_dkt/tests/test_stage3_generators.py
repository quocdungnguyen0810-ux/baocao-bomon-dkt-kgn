# -*- coding: utf-8 -*-
"""
Unit Test & Kiểm toán Giai đoạn 3 (Audit Stage 3):
Kiểm toán Bộ Sinh Báo cáo Excel (.xlsx Mẫu 01, 02, 03) và Bản cứng Word (.docx).
Đọc ngược lại các tệp đã xuất bằng openpyxl và python-docx để đối chiếu từng ô công thức.
"""

from pathlib import Path
import tempfile
import unittest
import docx
import openpyxl

from he_thong_bao_cao_dkt.database_manager import QuanLyCSDLBộMôn
from he_thong_bao_cao_dkt.docx_generator import BoSinhBaoCaoDocx
from he_thong_bao_cao_dkt.excel_generator import BoSinhBaoCaoExcel


class TestStage3Generators(unittest.TestCase):
    """Kiểm toán tính nguyên vẹn công thức Excel và định dạng DOCX của Mẫu 01, 02, 03."""

    def setUp(self) -> None:
        self.tmp_dir = tempfile.TemporaryDirectory()
        self.out_dir = Path(self.tmp_dir.name)
        self.db = QuanLyCSDLBộMôn(self.out_dir / "CSDL_Test.xlsx")
        self.db.khoi_tao_csdl_mau_neu_chua_co(ghi_de=True)
        self.excel_gen = BoSinhBaoCaoExcel()
        self.docx_gen = BoSinhBaoCaoDocx()

    def tearDown(self) -> None:
        self.tmp_dir.cleanup()

    def test_01_xuat_va_kiem_toan_mau_01_excel(self) -> None:
        """Kiểm toán từng ô dữ liệu và công thức của tệp BCThang_DKT_NguyenHaiHa_11-2026.xlsx."""
        bc_ha = self.db.doc_bao_cao_ca_nhan("1051", thang=11, nam=2026)
        file_01 = self.excel_gen.xuat_mau_01_ca_nhan(bc_ha, self.out_dir)
        self.assertTrue(file_01.exists())
        self.assertEqual(file_01.name, "BCThang_DKT_NguyenHaiHa_11-2026.xlsx")

        wb = openpyxl.load_workbook(file_01, data_only=False)
        self.assertIn("BC tháng", wb.sheetnames)
        self.assertIn("BC tuần", wb.sheetnames)
        self.assertIn("Danh mục", wb.sheetnames)

        ws = wb["BC tháng"]
        self.assertEqual(ws["C5"].value, 11)
        self.assertEqual(ws["E5"].value, 2026)
        self.assertEqual(ws["G5"].value, '=IF(OR(C5="",E5=""),"",DATE(E5,C5+1,0))')
        self.assertEqual(ws["C7"].value, "Nguyễn Hải Hà")
        self.assertEqual(ws["C9"].value, "Trần Thế Truyền")

        # Kiểm tra dòng 13 đã được thay bằng dữ liệu thật (không còn dòng ví dụ) và giữ nguyên công thức K13, L13, M13
        self.assertIn("Giảng dạy các học phần Cơ học đất", ws["B13"].value)
        self.assertEqual(ws["I13"].value, 1.0)
        self.assertEqual(
            ws["K13"].value,
            '=IF(B13="","",IF(OR(AND(F13<>"",$G$5<>"",F13<=$G$5),G13<>""),1,0))',
        )
        self.assertEqual(
            ws["L13"].value,
            '=IF(B13="","",IF(AND(G13<>"",OR(F13="",G13<=F13)),1,0))',
        )
        self.assertTrue(str(ws["M13"].value).startswith("=IF("))

        # Kiểm tra công thức vùng tổng hợp E30:E38, H42:H44, J50, F51, E52
        self.assertEqual(ws["E30"].value, "=COUNTA(B13:B27)")
        self.assertEqual(ws["E36"].value, "=SUM(K13:K27)")
        self.assertEqual(ws["E37"].value, "=SUM(L13:L27)")
        self.assertEqual(ws["E38"].value, '=IFERROR(E37/E36,"")')
        self.assertEqual(ws["E42"].value, 36.0)
        self.assertEqual(ws["H42"].value, '=IFERROR(F42/G42,"")')
        self.assertEqual(ws["F49"].value, 29.0)
        self.assertEqual(ws["F50"].value, 70.0)
        self.assertEqual(ws["J50"].value, '=IFERROR(ROUND(70*E38,0),"")')
        self.assertEqual(ws["F51"].value, '=IF(COUNT(F49:F50)=0,"",SUM(F49:F50))')
        self.assertTrue(str(ws["E52"].value).startswith("=IF("))
        wb.close()

    def test_02_xuat_va_kiem_toan_mau_02_va_mau_03_don_vi(self) -> None:
        """Kiểm toán tổng hợp Mẫu 02 (02b & 02c) và Mẫu 03 (03a, 03b, 03c) cho toàn bộ 22 thành viên Bộ môn."""
        ds_bc = self.db.doc_toan_bo_bao_cao_bo_mon(thang=11, nam=2026)
        self.assertEqual(len(ds_bc), 22)

        file_02 = self.excel_gen.xuat_mau_02_don_vi(ds_bc, self.out_dir, thang=11, nam=2026)
        self.assertTrue(file_02.exists())
        wb2 = openpyxl.load_workbook(file_02, data_only=False)
        ws_02b = wb2["BC tháng ĐV"]
        ws_02c = wb2["Theo dõi cá nhân"]

        # Kiểm tra Mẫu 02b
        self.assertEqual(ws_02b["E44"].value, '=IF(E40="","",E40-E41-E42-E43)')
        self.assertEqual(ws_02b["E49"].value, 22)

        # Kiểm tra Mẫu 02c có đủ 22 thành viên từ dòng 7 đến dòng 28
        ten_gv_02c = [ws_02c.cell(r, 2).value for r in range(7, 29)]
        self.assertEqual(len(ten_gv_02c), 22)
        self.assertIn("Nguyễn Hải Hà", ten_gv_02c)
        self.assertIn("Trần Thế Truyền", ten_gv_02c)
        self.assertIn("Nguyễn Thị Thanh Giang", ten_gv_02c)
        self.assertNotIn("Nguyễn Văn A", ten_gv_02c)
        self.assertNotIn("Nguyễn Thị Định", ten_gv_02c)
        self.assertEqual(ws_02c["G49"].value, "=COUNTA(B7:B46)")
        wb2.close()

        # Kiểm tra Mẫu 03
        file_03 = self.excel_gen.xuat_mau_03_so_giao_viec(ds_bc, self.out_dir, thang=11, nam=2026)
        self.assertTrue(file_03.exists())
        wb3 = openpyxl.load_workbook(file_03, data_only=False)
        ws_03b = wb3["Sổ giao việc"]
        ws_03c = wb3["Tổng hợp theo người"]
        self.assertEqual(ws_03b["C8"].value, '=IF(B8="","","GV"&TEXT(ROW()-7,"000"))')
        ten_gv_03c = [ws_03c.cell(r, 2).value for r in range(9, 30)]
        self.assertIn("Nguyễn Hải Hà", ten_gv_03c)
        wb3.close()

    def test_03_xuat_va_kiem_toan_docx_ban_cung(self) -> None:
        """Kiểm toán tệp bản cứng Word (.docx) cá nhân và đơn vị."""
        bc_ha = self.db.doc_bao_cao_ca_nhan("1051", thang=11, nam=2026)
        file_docx = self.docx_gen.xuat_docx_ca_nhan_01b(bc_ha, self.out_dir)
        self.assertTrue(file_docx.exists())
        doc = docx.Document(file_docx)
        full_text = "\n".join(p.text for p in doc.paragraphs)
        self.assertIn("Nguyễn Hải Hà", full_text)
        self.assertIn("Trần Thế Truyền", full_text)
        self.assertIn("HOẠT ĐỘNG CHUYÊN MÔN", full_text)
        self.assertGreaterEqual(len(doc.tables), 4)

        # Kiểm tra riêng đối với Nguyễn Thị Thanh Giang (0537): Đơn vị = Trung tâm KHCN GTVT, Trưởng đơn vị = Lương Xuân Chiểu
        bc_giang = self.db.doc_bao_cao_ca_nhan("0537", thang=11, nam=2026)
        f_giang_xlsx = self.excel_gen.xuat_mau_01_ca_nhan(bc_giang, self.out_dir, cap_nhat_ca_bc_tuan=True)
        wb_g = openpyxl.load_workbook(f_giang_xlsx, data_only=False)
        self.assertEqual(wb_g["BC tháng"]["A2"].value, "ĐƠN VỊ: TRUNG TÂM KHCN GTVT")
        self.assertEqual(wb_g["BC tháng"]["C9"].value, "Lương Xuân Chiểu")
        self.assertEqual(wb_g["BC tháng"]["B57"].value, "Bảo trì thiết bị máy móc định kỳ")
        self.assertEqual(wb_g["BC tháng"]["B58"].value, "NCKH và xâm nhập thực tiễn")
        self.assertIsNone(wb_g["BC tháng"]["B59"].value)
        self.assertEqual(wb_g["BC tuần"]["A2"].value, "ĐƠN VỊ: TRUNG TÂM KHCN GTVT")
        self.assertEqual(wb_g["BC tuần"]["C9"].value, "Lương Xuân Chiểu")
        self.assertEqual(wb_g["BC tuần"]["B39"].value, "Bảo trì thiết bị máy móc định kỳ")
        self.assertEqual(wb_g["BC tuần"]["B40"].value, "NCKH và xâm nhập thực tiễn")
        wb_g.close()

        f_giang_docx = self.docx_gen.xuat_docx_ca_nhan_01b(bc_giang, self.out_dir)
        doc_g = docx.Document(f_giang_docx)
        text_g = "\n".join(p.text for p in doc_g.paragraphs)
        tbl_text_g = "\n".join(c.text for t in doc_g.tables for r in t.rows for c in r.cells)
        self.assertIn("Trung tâm KHCN GTVT", text_g)
        self.assertIn("Lương Xuân Chiểu", text_g)
        self.assertIn("TRUNG TÂM KHCN GTVT", tbl_text_g)
        self.assertIn("Lương Xuân Chiểu", tbl_text_g)


if __name__ == "__main__":
    unittest.main()
