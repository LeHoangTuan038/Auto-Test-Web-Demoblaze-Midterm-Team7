from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR
from pptx.util import Inches, Pt


ROOT = Path(__file__).resolve().parents[1]
NAVY = RGBColor(24, 45, 77)
BLUE = RGBColor(39, 111, 191)
TEAL = RGBColor(0, 143, 145)
PALE = RGBColor(237, 244, 250)
WHITE = RGBColor(255, 255, 255)
INK = RGBColor(38, 50, 56)
MUTED = RGBColor(95, 109, 120)


TEST_CASES = [
    (
        "TC01", "Smoke", "Mở Sign up, đăng ký username mới, sau đó đăng nhập.",
        "Username duy nhất; password Test-pass-123.",
        "Đăng ký thành công; đăng nhập và thấy lời chào chứa username.",
        "test_sign_up_and_login_successfully",
    ),
    (
        "TC02", "Smoke", "Mở trang chủ và kiểm tra danh sách sản phẩm.",
        "Demoblaze home page.",
        "Danh sách có sản phẩm và mỗi tên sản phẩm không rỗng.",
        "test_product_catalog_is_displayed",
    ),
    (
        "TC03", "Smoke", "Mở sản phẩm đầu tiên, thêm vào giỏ và mở Cart.",
        "Sản phẩm đầu tiên trong catalog.",
        "Giỏ hàng hiển thị đúng một sản phẩm.",
        "test_user_can_add_a_product_to_cart",
    ),
    (
        "TC04", "Smoke", "Thêm sản phẩm; mở Cart, Place Order, điền thông tin và Purchase.",
        "Test Buyer; Vietnam; Hanoi; thẻ 4111111111111111; tháng 10; năm 2026.",
        "Hiển thị thông báo Thank you for your purchase!",
        "test_user_can_complete_purchase",
    ),
    (
        "TC05", "Smoke", "Đăng ký và đăng nhập, sau đó chọn Log out.",
        "Username duy nhất; password Test-pass-123.",
        "Đăng xuất; liên kết Log in và Sign up hiện lại.",
        "test_user_can_log_out",
    ),
    (
        "TC06", "Regression", "Đăng ký user mới rồi đăng nhập với password sai.",
        "Username đã đăng ký; password Wrong-pass-456.",
        "Hiện alert Wrong password.",
        "test_login_rejects_wrong_password",
    ),
    (
        "TC07", "Regression", "Mở Log in, để trống username, nhập password và submit.",
        "Username rỗng; password Test-pass-123.",
        "Hiện alert yêu cầu Username và Password.",
        "test_login_requires_username",
    ),
    (
        "TC08", "Regression", "Mở Log in, nhập username, để trống password và submit.",
        "Username mới; password rỗng.",
        "Hiện alert yêu cầu Username và Password.",
        "test_login_requires_password",
    ),
    (
        "TC09", "Regression", "Đăng ký một username; thử đăng ký lại username đó.",
        "Cùng một username duy nhất cho cả hai lần.",
        "Lần đầu thành công; lần hai báo user đã tồn tại.",
        "test_sign_up_rejects_duplicate_username",
    ),
    (
        "TC10", "Regression", "Tại trang chủ chọn danh mục Phones.",
        "Danh mục Phones.",
        "Danh sách được lọc; sản phẩm laptop VAIO không còn hiển thị.",
        "test_category_filter_shows_phones",
    ),
    (
        "TC11", "Regression", "Ghi tên sản phẩm đầu trang; chọn Next.",
        "Trang đầu catalog.",
        "Danh sách chuyển trang và tên sản phẩm đầu tiên thay đổi.",
        "test_user_can_navigate_to_next_catalog_page",
    ),
    (
        "TC12", "Regression", "Thêm một sản phẩm; vào Cart và xóa sản phẩm cuối.",
        "Một sản phẩm.",
        "Giỏ không còn dòng sản phẩm; không có badge số lượng.",
        "test_removing_last_product_leaves_cart_empty",
    ),
    (
        "TC13", "Regression", "Thêm hàng, mở Place Order, để trống Name rồi Purchase.",
        "Name rỗng; Credit card có giá trị.",
        "Hiện alert yêu cầu Name và Creditcard.",
        "test_purchase_requires_name",
    ),
    (
        "TC14", "Regression", "Thêm hàng, mở Place Order, để trống Credit card rồi Purchase.",
        "Name có giá trị; Credit card rỗng.",
        "Hiện alert yêu cầu Name và Creditcard.",
        "test_purchase_requires_credit_card",
    ),
    (
        "TC15", "Regression", "Mở Cart khi chưa thêm sản phẩm; kiểm tra trạng thái và Place Order.",
        "Giỏ hàng rỗng.",
        "Không có sản phẩm, tổng rỗng/0 và mở được form đặt hàng; không gửi đơn.",
        "test_empty_cart_has_no_products_and_order_form_can_be_opened",
    ),
    (
        "TC16", "Regression", "Thêm sản phẩm; mở Cart; cộng giá từng dòng và so với Total.",
        "Một sản phẩm có giá hiển thị.",
        "Total bằng tổng giá của các sản phẩm trong giỏ.",
        "test_cart_total_matches_sum_of_item_prices",
    ),
]


def create_workbook():
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Test cases"
    headers = [
        "ID",
        "Loại (Smoke/Regression)",
        "Các bước",
        "Dữ liệu",
        "Kết quả mong đợi",
        "Tên hàm test tương ứng",
        "Kết quả chạy",
    ]
    sheet.append(headers)
    for case in TEST_CASES:
        sheet.append([*case, "PASS"])

    header_fill = PatternFill("solid", fgColor="182D4D")
    smoke_fill = PatternFill("solid", fgColor="E5F4EE")
    regression_fill = PatternFill("solid", fgColor="EDF4FA")
    for cell in sheet[1]:
        cell.fill = header_fill
        cell.font = Font(name="Aptos", color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for row in sheet.iter_rows(min_row=2):
        row_fill = smoke_fill if row[1].value == "Smoke" else regression_fill
        for cell in row:
            cell.fill = row_fill
            cell.font = Font(name="Aptos", color="263238", size=10)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        row[0].font = Font(name="Aptos", color="182D4D", bold=True)
        row[6].font = Font(name="Aptos", color="18794E", bold=True)
        sheet.row_dimensions[row[0].row].height = 54

    widths = [10, 24, 48, 43, 48, 54, 16]
    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[chr(64 + index)].width = width
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    sheet.row_dimensions[1].height = 34
    table = Table(displayName="DemoblazeTestCases", ref=sheet.dimensions)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    sheet.add_table(table)
    sheet.sheet_view.showGridLines = False
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.fitToWidth = 1
    sheet.sheet_properties.pageSetUpPr.fitToPage = True

    notes = workbook.create_sheet("Ghi chú")
    notes.append(["Hạng mục", "Ghi chú"])
    notes.append(
        [
            "Website",
            "Demoblaze (https://demoblaze.com/); bộ test dùng Selenium WebDriver + Python + pytest.",
        ]
    )
    notes.append(
        [
            "Kết quả",
            f"PASS là kết quả lần chạy tự động xác nhận ngày {date.today().isoformat()}; chạy lại có thể cho kết quả khác.",
        ]
    )
    notes.append(
        [
            "Thay đổi kịch bản",
            "Demoblaze không có tài khoản khóa, sắp xếp, ZIP code, badge giỏ hoặc thuế. Các case được thay bằng đăng ký trùng, lọc/phân trang, kiểm tra Credit card, giỏ rỗng và tổng giá sản phẩm.",
        ]
    )
    notes.append(
        [
            "Checkout giỏ trống",
            "Chỉ kiểm tra trạng thái rỗng và khả năng mở form. Không bấm Purchase vì ứng dụng có thể báo thành công dù giỏ trống.",
        ]
    )
    notes.column_dimensions["A"].width = 27
    notes.column_dimensions["B"].width = 110
    for cell in notes[1]:
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
    for row in notes.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        notes.row_dimensions[row[0].row].height = 44
    notes.sheet_view.showGridLines = False
    workbook.save(ROOT / "Test-Cases-Demoblaze.xlsx")


def add_text(slide, x, y, w, h, text, *, size=20, color=INK, bold=False):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    frame = box.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.margin_left = Inches(0.06)
    frame.margin_right = Inches(0.06)
    frame.margin_top = Inches(0.04)
    frame.margin_bottom = Inches(0.02)
    for index, line in enumerate(text.split("\n")):
        paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        paragraph.text = line
        paragraph.font.name = "Aptos"
        paragraph.font.size = Pt(size)
        paragraph.font.bold = bold
        paragraph.font.color.rgb = color
        paragraph.space_after = Pt(9)
    return box


def add_bullets(slide, lines, *, x=0.85, y=1.55, w=11.5, size=22, gap=0.64):
    for index, line in enumerate(lines):
        add_text(slide, x, y + index * gap, 0.35, 0.35, "•", size=size + 2, color=TEAL, bold=True)
        add_text(slide, x + 0.42, y + index * gap, w - 0.42, gap, line, size=size)


def add_footer(slide, number):
    add_text(slide, 0.55, 7.12, 9.8, 0.2, "KIỂM THỬ TỰ ĐỘNG • DEMOBLAZE • SELENIUM + PYTHON", size=9, color=MUTED)
    add_text(slide, 12.15, 7.08, 0.6, 0.25, f"{number:02}", size=10, color=MUTED)


def add_title(slide, title, subtitle=None):
    add_text(slide, 0.65, 0.38, 12.05, 0.55, title, size=28, color=NAVY, bold=True)
    if subtitle:
        add_text(slide, 0.68, 1.02, 11.8, 0.35, subtitle, size=13, color=MUTED)


def add_content_slide(prs, number, title, bullets, subtitle=None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = WHITE
    add_title(slide, title, subtitle)
    add_bullets(slide, bullets)
    add_footer(slide, number)
    return slide


def create_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    cover = prs.slides.add_slide(prs.slide_layouts[6])
    cover.background.fill.solid()
    cover.background.fill.fore_color.rgb = NAVY
    add_text(cover, 0.9, 1.15, 11.4, 0.45, "ĐỀ TÀI GIỮA KỲ", size=18, color=RGBColor(117, 220, 214), bold=True)
    add_text(cover, 0.9, 2.0, 11.5, 1.25, "KIỂM THỬ TỰ ĐỘNG\nDEMOBLAZE", size=36, color=WHITE, bold=True)
    add_text(cover, 0.94, 4.0, 10.8, 0.8, "Selenium WebDriver • Python • pytest", size=24, color=RGBColor(218, 230, 242))
    add_text(cover, 0.94, 5.55, 11.0, 0.5, "Thành viên: [Tên thành viên 1] • [Tên thành viên 2] • [Tên thành viên 3]", size=16, color=WHITE)
    add_text(cover, 0.94, 6.14, 10.5, 0.35, "Lớp: [Tên lớp]    Giảng viên: [Tên giảng viên]", size=13, color=RGBColor(185, 202, 219))
    add_footer(cover, 1)

    add_content_slide(prs, 2, "Mục tiêu và phạm vi", [
        "Tự động hóa các luồng mua hàng trên website Demoblaze.",
        "Dùng Selenium WebDriver điều khiển Chrome bằng Python; pytest tổ chức và chạy test.",
        "Bao phủ đăng ký/đăng nhập, catalog, giỏ hàng và checkout.",
        "Phân loại 5 Smoke Test và 11 Regression Test; lưu kết quả HTML và bảng Excel.",
    ])
    add_content_slide(prs, 3, "Automation Test là gì?", [
        "Dùng chương trình thực hiện thao tác và so sánh kết quả thực tế với kết quả mong đợi.",
        "Có thể chạy lặp lại cùng một bộ dữ liệu, giảm thao tác thủ công và lỗi nhập liệu.",
        "Phù hợp cho luồng ổn định, kiểm thử hồi quy và kiểm tra sau mỗi thay đổi.",
        "Không thay thế hoàn toàn kiểm thử khám phá hoặc đánh giá trải nghiệm người dùng.",
    ])
    add_content_slide(prs, 4, "Khi nào nên tự động hóa?", [
        "Luồng nghiệp vụ lặp lại nhiều lần và có kết quả kiểm tra rõ ràng.",
        "Các chức năng quan trọng cần kiểm tra ở mỗi lần cập nhật.",
        "Có môi trường và dữ liệu kiểm thử ổn định; kết quả có thể xác minh.",
        "Cân nhắc chi phí bảo trì locator, thay đổi giao diện và phụ thuộc mạng.",
    ])
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_title(slide, "So sánh công cụ và lựa chọn")
    headers = ["Công cụ", "Điểm mạnh", "Lưu ý"]
    rows = [
        ["Selenium WebDriver", "Hệ sinh thái lớn; Python/Java; phù hợp kiểm thử trình duyệt.", "Chờ phần tử và quản lý driver cần thiết kế rõ ràng."],
        ["Playwright", "Tự chờ tốt; API hiện đại; có Codegen.", "Khác công cụ Selenium được yêu cầu trong đề."],
        ["Postman", "Thuận tiện kiểm thử API.", "Không tự kiểm thử giao diện người dùng như yêu cầu luồng web."],
    ]
    table = slide.shapes.add_table(4, 3, Inches(0.7), Inches(1.55), Inches(11.95), Inches(3.8)).table
    table.columns[0].width = Inches(2.15)
    table.columns[1].width = Inches(4.7)
    table.columns[2].width = Inches(5.1)
    for col, header in enumerate(headers):
        cell = table.cell(0, col)
        cell.text = header
    for row_index, row in enumerate(rows, start=1):
        for col_index, value in enumerate(row):
            table.cell(row_index, col_index).text = value
    for row_index in range(4):
        for col_index in range(3):
            cell = table.cell(row_index, col_index)
            cell.margin_left = Inches(0.1)
            cell.margin_right = Inches(0.1)
            cell.margin_top = Inches(0.08)
            cell.margin_bottom = Inches(0.05)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            cell.fill.fore_color.rgb = NAVY if row_index == 0 else (PALE if row_index % 2 else WHITE)
            for paragraph in cell.text_frame.paragraphs:
                paragraph.font.name = "Aptos"
                paragraph.font.size = Pt(14 if row_index else 16)
                paragraph.font.bold = row_index == 0 or col_index == 0
                paragraph.font.color.rgb = WHITE if row_index == 0 else INK
    add_text(slide, 0.9, 5.75, 11.5, 0.65, "Lựa chọn: Selenium WebDriver + Python theo yêu cầu đề tài; pytest để nhóm test, marker và tạo báo cáo.", size=18, color=NAVY, bold=True)
    add_footer(slide, 5)

    add_content_slide(prs, 6, "Website được kiểm thử: Demoblaze", [
        "URL: https://demoblaze.com/",
        "Website demo bán điện thoại, laptop và màn hình; có đăng ký/đăng nhập.",
        "Luồng chính: chọn sản phẩm → thêm vào Cart → Place Order → Purchase.",
        "Trang dùng modal và JavaScript alert; test cần chờ trạng thái thay đổi trước khi thao tác.",
    ])
    add_content_slide(prs, 7, "Cấu trúc project", [
        "conftest.py: cấu hình Chrome, fixture browser và tùy chọn --headed.",
        "tests/helpers.py: thao tác dùng chung, wait, đăng nhập, giỏ hàng và checkout.",
        "tests/test_login.py: đăng ký, đăng nhập, đăng xuất và trường hợp lỗi.",
        "tests/test_cart.py và tests/test_checkout.py: danh mục, giỏ hàng, đặt hàng.",
        "pytest.ini, requirements.txt, README.md; kèm test case Excel và slide trình bày.",
    ])
    add_content_slide(prs, 8, "Luồng thực thi của Selenium", [
        "pytest khởi tạo fixture → Selenium WebDriver mở Chrome.",
        "WebDriver tìm phần tử theo ID/CSS/XPath và thực hiện click/nhập liệu.",
        "WebDriverWait đợi điều kiện cụ thể thay vì chờ cố định.",
        "Assertion so sánh URL, alert, dòng sản phẩm hoặc thông báo xác nhận.",
        "Fixture đóng trình duyệt sau mỗi test để giảm ảnh hưởng dữ liệu giữa các case.",
    ])
    add_content_slide(prs, 9, "Smoke Test và Regression Test", [
        "Smoke (5 case): kiểm tra các luồng thiết yếu có thể sử dụng.",
        "Đăng ký/đăng nhập, hiển thị catalog, thêm sản phẩm, checkout thành công, đăng xuất.",
        "Regression (11 case): kiểm tra lỗi nhập liệu, lọc/phân trang và biên của giỏ/checkout.",
        "Chạy nhanh Smoke bằng pytest -m smoke; chạy đầy đủ bằng pytest.",
    ])
    add_content_slide(prs, 10, "Đăng nhập, catalog và giỏ hàng", [
        "Tài khoản kiểm thử được tạo với username ngẫu nhiên để tránh trùng dữ liệu.",
        "Kiểm tra sai mật khẩu, thiếu username/password và đăng ký username đã tồn tại.",
        "Catalog: xác nhận sản phẩm hiển thị, lọc Phones và chuyển sang trang kế tiếp.",
        "Giỏ: thêm sản phẩm, xác nhận dòng hàng, xóa sản phẩm cuối và kiểm tra giỏ rỗng.",
    ])
    add_content_slide(prs, 11, "Checkout và xác minh tổng tiền", [
        "Luồng thành công: Place Order → nhập thông tin → Purchase → thông báo cảm ơn.",
        "Kiểm tra validation khi thiếu Name hoặc Credit card.",
        "Giỏ rỗng: xác nhận không có dòng hàng và có thể mở form; không gửi đơn.",
        "Đối chiếu Total trong Cart với tổng giá từng sản phẩm.",
    ])
    add_content_slide(prs, 12, "Chạy test và demo", [
        "Cài: py -m venv .venv; kích hoạt môi trường; pip install -r requirements.txt.",
        "Chạy toàn bộ: pytest; trình diễn trình duyệt: pytest --headed.",
        "Chạy theo nhóm: pytest -m smoke hoặc pytest -m regression.",
        "Mỗi lần chạy pytest tạo results/report.html tĩnh; báo cáo tương tác tùy chọn: pytest --html=results/pytest-html-report.html --self-contained-html.",
        "Selenium Manager cần mạng lần đầu để chuẩn bị driver tương thích Chrome.",
    ])
    add_content_slide(prs, 13, "Kết quả chạy", [
        "Bộ test gồm 16 case: 5 Smoke và 11 Regression.",
        "Kết quả xác minh gần nhất: 16/16 PASS.",
        "Báo cáo HTML được tạo trong thư mục results/ sau khi chạy test.",
        "Excel liệt kê ID, bước, dữ liệu, kết quả mong đợi, hàm test và kết quả chạy.",
    ])
    add_content_slide(prs, 14, "Khó khăn, giới hạn và bài học", [
        "Ứng dụng dùng alert JavaScript và dữ liệu tải bất đồng bộ: cần explicit wait.",
        "Demoblaze không có locked account, sắp xếp, ZIP code, badge hoặc thuế.",
        "Đã thay bằng chức năng tương đương có thật: user trùng, lọc/phân trang, kiểm tra card và tổng giá.",
        "Website demo công cộng có thể thay đổi hoặc chậm; test phụ thuộc kết nối mạng.",
        "Bài học: locator rõ ràng, test độc lập, dữ liệu riêng và báo cáo kết quả có thể tái lập.",
    ])
    add_content_slide(prs, 15, "Kết luận và phân công", [
        "Selenium + Python + pytest tự động hóa được các luồng cốt lõi trên Demoblaze.",
        "Người 1: cấu trúc, cài đặt, fixture và demo; Người 2: đăng nhập và test case Excel.",
        "Người 3: giỏ hàng, checkout và nội dung trình bày.",
        "Cập nhật tên thành viên, lớp và giảng viên trước khi nộp.",
        "Cảm ơn thầy/cô và các bạn!",
    ])

    prs.save(ROOT / "Midterm-Automation-Test.pptx")


if __name__ == "__main__":
    create_workbook()
    create_presentation()
    print(f"Created deliverables in {ROOT}")
